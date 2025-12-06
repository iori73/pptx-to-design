import JSZip from 'jszip';
import {
  PPTXSlide,
  PPTXTheme,
  ConversionContext,
  emuToPixels,
} from '../types/pptx';
import { parseSlide } from './slideParser';
import {
  parseXML,
  ParsedElement,
  getElementByLocalName,
  getChildByLocalName,
  getChildrenByLocalName,
  getAttribute,
  getAttributeAsNumber,
  parseColor,
} from './xmlUtils';

export interface PPTXParseResult {
  slides: PPTXSlide[];
  slideWidth: number;
  slideHeight: number;
  theme?: PPTXTheme;
}

/**
 * Parse a .pptx file
 */
export async function parsePPTX(fileData: ArrayBuffer): Promise<PPTXParseResult> {
  const zip = await JSZip.loadAsync(fileData);

  // Parse presentation.xml to get slide size and slide list
  const presentationXml = await zip.file('ppt/presentation.xml')?.async('string');
  if (!presentationXml) {
    throw new Error('Invalid PPTX file: presentation.xml not found');
  }

  const { slideWidth, slideHeight, slideRIds } = parsePresentationXml(presentationXml);

  // Parse presentation relationships
  const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels')?.async('string');
  if (!presRelsXml) {
    throw new Error('Invalid PPTX file: presentation relationships not found');
  }

  const presRelationships = parseRelationships(presRelsXml);

  // Parse theme
  const theme = await parseTheme(zip, presRelationships);

  // Load media files
  const mediaFiles = await loadMediaFiles(zip);

  // Parse each slide
  const slides: PPTXSlide[] = [];

  for (let i = 0; i < slideRIds.length; i++) {
    const rId = slideRIds[i];
    const slidePath = presRelationships.get(rId);
    if (!slidePath) continue;

    // Normalize path (remove leading ../ if present)
    const normalizedPath = slidePath.startsWith('../')
      ? slidePath.substring(3)
      : slidePath.startsWith('/')
      ? slidePath.substring(1)
      : `ppt/${slidePath}`;

    const slideXml = await zip.file(normalizedPath)?.async('string');
    if (!slideXml) continue;

    // Load slide relationships
    const slideRelsPath = normalizedPath.replace(
      /slides\/slide(\d+)\.xml/,
      'slides/_rels/slide$1.xml.rels'
    );
    const slideRelsXml = await zip.file(slideRelsPath)?.async('string');
    const slideRelationships = slideRelsXml
      ? parseRelationships(slideRelsXml)
      : new Map<string, string>();

    // Create conversion context
    const context: ConversionContext = {
      slideWidth,
      slideHeight,
      theme,
      mediaFiles,
      relationships: slideRelationships,
    };

    // Parse the slide
    const slide = parseSlide(slideXml, i + 1, context);
    slides.push(slide);
  }

  return {
    slides,
    slideWidth,
    slideHeight,
    theme,
  };
}

/**
 * Parse presentation.xml to get slide size and slide list
 */
function parsePresentationXml(xml: string): {
  slideWidth: number;
  slideHeight: number;
  slideRIds: string[];
} {
  const doc = parseXML(xml);

  // Get slide size
  const sldSz = getElementByLocalName(doc, 'sldSz');
  let slideWidth = 960; // Default 10 inches at 96 DPI
  let slideHeight = 540; // Default 7.5 inches at 96 DPI

  if (sldSz) {
    const cx = getAttributeAsNumber(sldSz, 'cx', 9144000);
    const cy = getAttributeAsNumber(sldSz, 'cy', 6858000);
    slideWidth = emuToPixels(cx);
    slideHeight = emuToPixels(cy);
  }

  // Get slide list
  const slideRIds: string[] = [];
  const sldIdLst = getElementByLocalName(doc, 'sldIdLst');
  if (sldIdLst) {
    const sldIds = getChildrenByLocalName(sldIdLst, 'sldId');
    for (const sldId of sldIds) {
      const rId = getAttribute(sldId, 'id', 'r');
      if (rId) {
        slideRIds.push(rId);
      }
    }
  }

  return { slideWidth, slideHeight, slideRIds };
}

/**
 * Parse relationship file
 */
function parseRelationships(xml: string): Map<string, string> {
  const relationships = new Map<string, string>();
  const doc = parseXML(xml);

  // Get all Relationship elements
  const relElements = getChildrenByLocalName(doc, 'Relationship');
  
  // Also check inside Relationships container
  const relationshipsContainer = getChildByLocalName(doc, 'Relationships');
  if (relationshipsContainer) {
    const innerRels = getChildrenByLocalName(relationshipsContainer, 'Relationship');
    relElements.push(...innerRels);
  }
  
  for (const rel of relElements) {
    const id = getAttribute(rel, 'Id');
    const target = getAttribute(rel, 'Target');
    if (id && target) {
      relationships.set(id, target);
    }
  }

  return relationships;
}

/**
 * Parse theme file
 */
async function parseTheme(
  zip: JSZip,
  presRelationships: Map<string, string>
): Promise<PPTXTheme | undefined> {
  // Find theme file
  let themePath: string | undefined;
  for (const [, target] of presRelationships) {
    if (target.includes('theme')) {
      themePath = target.startsWith('../')
        ? target.substring(3)
        : target.startsWith('/')
        ? target.substring(1)
        : `ppt/${target}`;
      break;
    }
  }

  if (!themePath) return undefined;

  const themeXml = await zip.file(themePath)?.async('string');
  if (!themeXml) return undefined;

  const doc = parseXML(themeXml);

  // Parse color scheme
  const colors = new Map<string, string>();
  const clrScheme = getElementByLocalName(doc, 'clrScheme');

  if (clrScheme) {
    const colorNames = [
      'dk1', 'lt1', 'dk2', 'lt2',
      'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
      'hlink', 'folHlink',
    ];

    for (const name of colorNames) {
      const colorEl = getChildByLocalName(clrScheme, name);
      if (colorEl) {
        const color = parseColor(colorEl);
        if (color) {
          colors.set(name, color);
        }
      }
    }
  }

  // Parse font scheme
  const fontScheme = getElementByLocalName(doc, 'fontScheme');
  let majorFont: string | undefined;
  let minorFont: string | undefined;

  if (fontScheme) {
    const majorFontEl = getElementByLocalName(fontScheme, 'majorFont');
    const minorFontEl = getElementByLocalName(fontScheme, 'minorFont');

    if (majorFontEl) {
      const latin = getChildByLocalName(majorFontEl, 'latin');
      if (latin) {
        majorFont = getAttribute(latin, 'typeface') || undefined;
      }
    }

    if (minorFontEl) {
      const latin = getChildByLocalName(minorFontEl, 'latin');
      if (latin) {
        minorFont = getAttribute(latin, 'typeface') || undefined;
      }
    }
  }

  return {
    name: 'Theme',
    colors,
    fonts: { majorFont, minorFont },
  };
}

/**
 * Load media files from the archive
 */
async function loadMediaFiles(zip: JSZip): Promise<Map<string, Uint8Array>> {
  const mediaFiles = new Map<string, Uint8Array>();

  // Get all files in ppt/media/
  const mediaFolder = zip.folder('ppt/media');
  if (mediaFolder) {
    const files = Object.keys(zip.files).filter((path) =>
      path.startsWith('ppt/media/')
    );

    for (const filePath of files) {
      const file = zip.file(filePath);
      if (file && !file.dir) {
        const data = await file.async('uint8array');
        // Store with relative path from slides folder
        const relativePath = `../media/${filePath.split('/').pop()}`;
        mediaFiles.set(relativePath, data);
      }
    }
  }

  return mediaFiles;
}
