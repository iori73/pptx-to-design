import {
  PPTXSlide,
  PPTXElement,
  PPTXTextElement,
  PPTXShapeElement,
  PPTXImageElement,
  PPTXTableElement,
  PPTXGroupElement,
  PPTXParagraph,
  PPTXTextRun,
  PPTXTableRow,
  PPTXTableCell,
  PPTXShapeType,
  PPTXFill,
  PPTXStroke,
  emuToPixels,
  ConversionContext,
} from '../types/pptx';
import {
  parseXML,
  ParsedElement,
  getElementsByLocalName,
  getElementByLocalName,
  getChildByLocalName,
  getChildrenByLocalName,
  getChildren,
  getAttribute,
  getAttributeAsNumber,
  getTextContent,
  parseColor,
  parseLineProperties,
  parseFillProperties,
} from './xmlUtils';

/**
 * Parse a slide XML content
 */
export function parseSlide(
  slideXml: string,
  slideIndex: number,
  context: ConversionContext
): PPTXSlide {
  const doc = parseXML(slideXml);
  const elements: PPTXElement[] = [];

  // Find the shape tree (spTree)
  const spTree = getElementByLocalName(doc, 'spTree');
  if (spTree) {
    parseShapeTree(spTree, elements, context);
  }

  return {
    index: slideIndex,
    width: context.slideWidth,
    height: context.slideHeight,
    elements,
  };
}

/**
 * Parse shape tree and extract all elements
 */
function parseShapeTree(
  spTree: ParsedElement,
  elements: PPTXElement[],
  context: ConversionContext
): void {
  const children = getChildren(spTree);
  for (const { name, element } of children) {
    const parsed = parseElement(name, element, context);
    if (parsed) {
      elements.push(parsed);
    }
  }
}

/**
 * Parse a single element
 */
function parseElement(
  elementName: string,
  element: ParsedElement,
  context: ConversionContext
): PPTXElement | null {
  // Get local name (remove namespace prefix)
  const localName = elementName.includes(':') ? elementName.split(':')[1] : elementName;

  switch (localName) {
    case 'sp': // Shape
      return parseShape(element, context);
    case 'pic': // Picture
      return parsePicture(element, context);
    case 'graphicFrame': // Table, Chart, etc.
      return parseGraphicFrame(element, context);
    case 'grpSp': // Group
      return parseGroup(element, context);
    case 'cxnSp': // Connector
      return parseConnector(element, context);
    default:
      return null;
  }
}

/**
 * Parse shape element
 */
function parseShape(
  shapeElement: ParsedElement,
  context: ConversionContext
): PPTXTextElement | PPTXShapeElement | null {
  // Get shape properties
  const spPr = getChildByLocalName(shapeElement, 'spPr');
  if (!spPr) return null;

  // Get transform (position and size)
  const xfrm = getChildByLocalName(spPr, 'xfrm');
  if (!xfrm) return null;

  const position = parseTransform(xfrm);
  if (!position) return null;

  // Get non-visual properties for name
  const nvSpPr = getChildByLocalName(shapeElement, 'nvSpPr');
  const cNvPr = nvSpPr ? getChildByLocalName(nvSpPr, 'cNvPr') : null;
  const name = cNvPr ? getAttribute(cNvPr, 'name') || undefined : undefined;

  // Check if it's a text box or shape
  const txBody = getChildByLocalName(shapeElement, 'txBody');
  const prstGeom = getChildByLocalName(spPr, 'prstGeom');
  const prst = prstGeom ? getAttribute(prstGeom, 'prst') : null;

  // Parse fill and stroke
  const fill = parseFillFromSpPr(spPr, context);
  const stroke = parseStrokeFromSpPr(spPr, context);

  // If it has text content, check if it's primarily a text element
  if (txBody) {
    const paragraphs = parseTxBody(txBody, context);
    const hasVisibleShape = fill?.type !== 'none' || stroke !== undefined;
    const isTextBox = prst === 'rect' && !hasVisibleShape;

    if (isTextBox || prst === null) {
      // Text element
      const textElement: PPTXTextElement = {
        type: 'text',
        ...position,
        name,
        paragraphs,
        backgroundColor: fill?.type === 'solid' ? fill.color : undefined,
      };
      return textElement;
    } else {
      // Shape with text
      const shapeType = mapPresetGeometry(prst || 'rect');
      const shapeResult: PPTXShapeElement = {
        type: 'shape',
        ...position,
        name,
        shapeType,
        fill,
        stroke,
        text: paragraphs,
        cornerRadius: parseCornerRadius(prstGeom),
      };
      return shapeResult;
    }
  }

  // Pure shape without text
  const shapeType = mapPresetGeometry(prst || 'rect');
  return {
    type: 'shape',
    ...position,
    name,
    shapeType,
    fill,
    stroke,
    cornerRadius: parseCornerRadius(prstGeom),
  };
}

/**
 * Parse transform element to get position and size
 */
function parseTransform(
  xfrm: ParsedElement
): { x: number; y: number; width: number; height: number; rotation?: number } | null {
  const off = getChildByLocalName(xfrm, 'off');
  const ext = getChildByLocalName(xfrm, 'ext');

  if (!off || !ext) return null;

  const x = emuToPixels(getAttributeAsNumber(off, 'x'));
  const y = emuToPixels(getAttributeAsNumber(off, 'y'));
  const width = emuToPixels(getAttributeAsNumber(ext, 'cx'));
  const height = emuToPixels(getAttributeAsNumber(ext, 'cy'));

  // Get rotation (in 60000ths of a degree)
  const rot = getAttributeAsNumber(xfrm, 'rot', 0);
  const rotation = rot !== 0 ? rot / 60000 : undefined;

  return { x, y, width, height, rotation };
}

/**
 * Parse fill from shape properties
 */
function parseFillFromSpPr(spPr: ParsedElement, context: ConversionContext): PPTXFill | undefined {
  const fillResult = parseFillProperties(spPr, context.theme?.colors);
  if (!fillResult) return undefined;

  return {
    type: fillResult.type,
    color: fillResult.color,
    opacity: fillResult.opacity,
  };
}

/**
 * Parse stroke from shape properties
 */
function parseStrokeFromSpPr(spPr: ParsedElement, context: ConversionContext): PPTXStroke | undefined {
  const ln = getChildByLocalName(spPr, 'ln');
  const lineProps = parseLineProperties(ln, context.theme?.colors);
  if (!lineProps) return undefined;

  return {
    color: lineProps.color,
    width: lineProps.width,
    opacity: 1,
    dashPattern: 'solid',
  };
}

/**
 * Parse corner radius from preset geometry
 */
function parseCornerRadius(prstGeom: ParsedElement | null): number | undefined {
  if (!prstGeom) return undefined;

  const avLst = getChildByLocalName(prstGeom, 'avLst');
  if (!avLst) return undefined;

  const gd = getChildByLocalName(avLst, 'gd');
  if (!gd) return undefined;

  const name = getAttribute(gd, 'name');
  const fmla = getAttribute(gd, 'fmla');

  if (name === 'adj' && fmla) {
    const match = fmla.match(/val\s+(\d+)/);
    if (match) {
      // Convert from 100000ths to percentage then to pixels
      const value = parseInt(match[1], 10);
      return (value / 100000) * 50; // Rough approximation
    }
  }

  return undefined;
}

/**
 * Parse text body
 */
function parseTxBody(txBody: ParsedElement, context: ConversionContext): PPTXParagraph[] {
  const paragraphs: PPTXParagraph[] = [];
  const pElements = getChildrenByLocalName(txBody, 'p');

  for (const pElement of pElements) {
    const paragraph = parseParagraph(pElement, context);
    if (paragraph.runs.length > 0 || paragraphs.length === 0) {
      paragraphs.push(paragraph);
    }
  }

  return paragraphs;
}

/**
 * Parse a paragraph
 */
function parseParagraph(pElement: ParsedElement, context: ConversionContext): PPTXParagraph {
  const runs: PPTXTextRun[] = [];

  // Get paragraph properties
  const pPr = getChildByLocalName(pElement, 'pPr');
  let alignment: 'left' | 'center' | 'right' | 'justify' = 'left';
  let bulletType: 'none' | 'bullet' | 'number' = 'none';
  let indentLevel = 0;

  if (pPr) {
    const algn = getAttribute(pPr, 'algn');
    if (algn) {
      alignment = mapAlignment(algn);
    }

    indentLevel = getAttributeAsNumber(pPr, 'lvl', 0);

    // Check for bullets
    const buNone = getChildByLocalName(pPr, 'buNone');
    const buChar = getChildByLocalName(pPr, 'buChar');
    const buAutoNum = getChildByLocalName(pPr, 'buAutoNum');

    if (buChar) {
      bulletType = 'bullet';
    } else if (buAutoNum) {
      bulletType = 'number';
    } else if (!buNone) {
      // Check default bullet
      const defRPr = getChildByLocalName(pPr, 'defRPr');
      if (defRPr) {
        // Has default run properties, might have bullet
      }
    }
  }

  // Parse text runs using getChildren to iterate through children
  const children = getChildren(pElement);
  for (const { name, element } of children) {
    const localName = name.includes(':') ? name.split(':')[1] : name;
    
    if (localName === 'r') {
      const run = parseTextRun(element, context);
      if (run) {
        runs.push(run);
      }
    } else if (localName === 'br') {
      // Line break - add newline to previous run or create empty run
      if (runs.length > 0) {
        runs[runs.length - 1].text += '\n';
      }
    } else if (localName === 'fld') {
      // Field (like page number) - parse as text run
      const run = parseTextRun(element, context);
      if (run) {
        runs.push(run);
      }
    }
  }

  return {
    runs,
    alignment,
    bulletType,
    indentLevel,
  };
}

/**
 * Parse a text run - IMPROVED VERSION
 */
function parseTextRun(rElement: ParsedElement, context: ConversionContext): PPTXTextRun | null {
  // Get text content - try multiple approaches
  let text = '';
  
  // Approach 1: Standard <a:t> element
  const tElement = getChildByLocalName(rElement, 't');
  if (tElement) {
    text = getTextContent(tElement);
  }
  
  // Approach 2: Direct #text content on the element (some parsers store it here)
  if (!text && rElement['#text'] !== undefined) {
    text = String(rElement['#text']);
  }
  
  // Approach 3: Check for t as direct property (fast-xml-parser might do this)
  if (!text && typeof rElement['t'] === 'string') {
    text = rElement['t'];
  }
  if (!text && typeof rElement['t'] === 'number') {
    text = String(rElement['t']);
  }
  
  // If still no text and the entire element is just a string or number
  if (!text && typeof rElement === 'string') {
    text = rElement;
  }
  if (!text && typeof rElement === 'number') {
    text = String(rElement);
  }
  
  // Return null only if truly no text (preserve spaces)
  if (text === '') return null;

  // Get run properties
  const rPr = getChildByLocalName(rElement, 'rPr');

  let fontFamily: string | undefined;
  let fontSize: number | undefined;
  let fontWeight: 'normal' | 'bold' = 'normal';
  let fontStyle: 'normal' | 'italic' = 'normal';
  let textDecoration: 'none' | 'underline' | 'strikethrough' = 'none';
  let color: string | undefined;

  if (rPr) {
    // Font size (in hundredths of a point)
    const sz = getAttributeAsNumber(rPr, 'sz', 0);
    if (sz > 0) {
      fontSize = sz / 100;
    }

    // Bold
    const b = getAttribute(rPr, 'b');
    if (b === '1' || b === 'true') {
      fontWeight = 'bold';
    }

    // Italic
    const i = getAttribute(rPr, 'i');
    if (i === '1' || i === 'true') {
      fontStyle = 'italic';
    }

    // Underline
    const u = getAttribute(rPr, 'u');
    if (u && u !== 'none') {
      textDecoration = 'underline';
    }

    // Strikethrough
    const strike = getAttribute(rPr, 'strike');
    if (strike && strike !== 'noStrike') {
      textDecoration = 'strikethrough';
    }

    // Font family - check multiple sources
    const latin = getChildByLocalName(rPr, 'latin');
    if (latin) {
      fontFamily = getAttribute(latin, 'typeface') || undefined;
    }
    // Also check for ea (East Asian) and cs (Complex Script) fonts
    if (!fontFamily) {
      const ea = getChildByLocalName(rPr, 'ea');
      if (ea) {
        fontFamily = getAttribute(ea, 'typeface') || undefined;
      }
    }

    // Color - check solidFill
    const solidFill = getChildByLocalName(rPr, 'solidFill');
    if (solidFill) {
      color = parseColor(solidFill, context.theme?.colors);
    }
    
    // If no explicit color, check for inherited color from theme
    if (!color) {
      // Default text color is typically black (dk1/tx1)
      // We'll let the converter handle the default
    }
  }

  return {
    text,
    fontFamily,
    fontSize,
    fontWeight,
    fontStyle,
    textDecoration,
    color,
  };
}

/**
 * Map PowerPoint alignment to standard values
 */
function mapAlignment(algn: string): 'left' | 'center' | 'right' | 'justify' {
  switch (algn) {
    case 'l':
      return 'left';
    case 'ctr':
      return 'center';
    case 'r':
      return 'right';
    case 'just':
      return 'justify';
    default:
      return 'left';
  }
}

/**
 * Parse picture element
 */
function parsePicture(
  picElement: ParsedElement,
  context: ConversionContext
): PPTXImageElement | null {
  // Get shape properties
  const spPr = getChildByLocalName(picElement, 'spPr');
  if (!spPr) return null;

  // Get transform
  const xfrm = getChildByLocalName(spPr, 'xfrm');
  if (!xfrm) return null;

  const position = parseTransform(xfrm);
  if (!position) return null;

  // Get non-visual properties
  const nvPicPr = getChildByLocalName(picElement, 'nvPicPr');
  const cNvPr = nvPicPr ? getChildByLocalName(nvPicPr, 'cNvPr') : null;
  const name = cNvPr ? getAttribute(cNvPr, 'name') || undefined : undefined;

  // Get image reference
  const blipFill = getChildByLocalName(picElement, 'blipFill');
  if (!blipFill) return null;

  const blip = getChildByLocalName(blipFill, 'blip');
  if (!blip) return null;

  // Get relationship ID
  const rEmbed = getAttribute(blip, 'embed', 'r');
  if (!rEmbed) return null;

  // Get image data from context
  const imagePath = context.relationships.get(rEmbed);
  if (!imagePath) return null;

  const imageData = context.mediaFiles.get(imagePath);
  if (!imageData) return null;

  // Determine MIME type from extension
  const ext = imagePath.split('.').pop()?.toLowerCase() || 'png';
  const mimeType = getMimeType(ext);

  return {
    type: 'image',
    ...position,
    name,
    imageData,
    mimeType,
  };
}

/**
 * Get MIME type from extension
 */
function getMimeType(ext: string): string {
  const mimeTypes: Record<string, string> = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    svg: 'image/svg+xml',
    webp: 'image/webp',
    emf: 'image/emf',
    wmf: 'image/wmf',
  };
  return mimeTypes[ext] || 'image/png';
}

/**
 * Parse graphic frame (table, chart, etc.)
 */
function parseGraphicFrame(
  frameElement: ParsedElement,
  context: ConversionContext
): PPTXTableElement | PPTXElement | null {
  // Get transform
  const xfrm = getElementByLocalName(frameElement, 'xfrm');
  if (!xfrm) return null;

  const position = parseTransform(xfrm);
  if (!position) return null;

  // Check for table
  const tbl = getElementByLocalName(frameElement, 'tbl');
  if (tbl) {
    return parseTable(tbl, position, context);
  }

  // Check for chart
  const chart = getElementByLocalName(frameElement, 'chart');
  if (chart) {
    // Charts are complex, return as placeholder for now
    return {
      type: 'shape',
      ...position,
      shapeType: 'rect',
      fill: { type: 'solid', color: '#E0E0E0' },
      text: [{ runs: [{ text: '[Chart]' }], alignment: 'center' }],
    } as PPTXShapeElement;
  }

  return null;
}

/**
 * Parse table element
 */
function parseTable(
  tblElement: ParsedElement,
  position: { x: number; y: number; width: number; height: number },
  context: ConversionContext
): PPTXTableElement {
  const rows: PPTXTableRow[] = [];
  const columnWidths: number[] = [];

  // Parse grid columns
  const tblGrid = getChildByLocalName(tblElement, 'tblGrid');
  if (tblGrid) {
    const gridCols = getChildrenByLocalName(tblGrid, 'gridCol');
    for (const gridCol of gridCols) {
      const w = getAttributeAsNumber(gridCol, 'w', 0);
      columnWidths.push(emuToPixels(w));
    }
  }

  // Parse rows
  const trElements = getChildrenByLocalName(tblElement, 'tr');
  for (const tr of trElements) {
    const row = parseTableRow(tr, context);
    rows.push(row);
  }

  return {
    type: 'table',
    ...position,
    rows,
    columnWidths,
  };
}

/**
 * Parse table row
 */
function parseTableRow(trElement: ParsedElement, context: ConversionContext): PPTXTableRow {
  const cells: PPTXTableCell[] = [];
  const height = emuToPixels(getAttributeAsNumber(trElement, 'h', 0));

  const tcElements = getChildrenByLocalName(trElement, 'tc');
  for (const tc of tcElements) {
    const cell = parseTableCell(tc, context);
    cells.push(cell);
  }

  return { height, cells };
}

/**
 * Parse table cell
 */
function parseTableCell(tcElement: ParsedElement, context: ConversionContext): PPTXTableCell {
  const content: PPTXParagraph[] = [];

  // Parse cell text content
  const txBody = getChildByLocalName(tcElement, 'txBody');
  if (txBody) {
    const paragraphs = parseTxBody(txBody, context);
    content.push(...paragraphs);
  }

  // Parse cell properties
  const tcPr = getChildByLocalName(tcElement, 'tcPr');
  let backgroundColor: string | undefined;
  let verticalAlign: 'top' | 'middle' | 'bottom' = 'top';

  if (tcPr) {
    // Background color
    const solidFill = getChildByLocalName(tcPr, 'solidFill');
    backgroundColor = parseColor(solidFill, context.theme?.colors);

    // Vertical alignment
    const anchor = getAttribute(tcPr, 'anchor');
    if (anchor === 'ctr') {
      verticalAlign = 'middle';
    } else if (anchor === 'b') {
      verticalAlign = 'bottom';
    }
  }

  // Get span info
  const rowSpan = getAttributeAsNumber(tcElement, 'rowSpan', 1);
  const gridSpan = getAttributeAsNumber(tcElement, 'gridSpan', 1);

  return {
    content,
    rowSpan: rowSpan > 1 ? rowSpan : undefined,
    colSpan: gridSpan > 1 ? gridSpan : undefined,
    backgroundColor,
    verticalAlign,
  };
}

/**
 * Parse group element
 */
function parseGroup(
  grpElement: ParsedElement,
  context: ConversionContext
): PPTXGroupElement | null {
  // Get group shape properties
  const grpSpPr = getChildByLocalName(grpElement, 'grpSpPr');
  if (!grpSpPr) return null;

  // Get transform
  const xfrm = getChildByLocalName(grpSpPr, 'xfrm');
  if (!xfrm) return null;

  const position = parseTransform(xfrm);
  if (!position) return null;

  // Get non-visual properties
  const nvGrpSpPr = getChildByLocalName(grpElement, 'nvGrpSpPr');
  const cNvPr = nvGrpSpPr ? getChildByLocalName(nvGrpSpPr, 'cNvPr') : null;
  const name = cNvPr ? getAttribute(cNvPr, 'name') || undefined : undefined;

  // Parse children
  const children: PPTXElement[] = [];
  const childElements = getChildren(grpElement);
  
  for (const { name: childName, element } of childElements) {
    const localName = childName.includes(':') ? childName.split(':')[1] : childName;
    // Skip property elements
    if (localName === 'nvGrpSpPr' || localName === 'grpSpPr') {
      continue;
    }
    const parsed = parseElement(childName, element, context);
    if (parsed) {
      children.push(parsed);
    }
  }

  return {
    type: 'group',
    ...position,
    name,
    children,
  };
}

/**
 * Parse connector shape
 */
function parseConnector(
  cxnElement: ParsedElement,
  context: ConversionContext
): PPTXShapeElement | null {
  const spPr = getChildByLocalName(cxnElement, 'spPr');
  if (!spPr) return null;

  const xfrm = getChildByLocalName(spPr, 'xfrm');
  if (!xfrm) return null;

  const position = parseTransform(xfrm);
  if (!position) return null;

  const stroke = parseStrokeFromSpPr(spPr, context);

  return {
    type: 'shape',
    ...position,
    shapeType: 'line',
    fill: { type: 'none' },
    stroke: stroke || { color: '#000000', width: 1 },
  };
}

/**
 * Map PowerPoint preset geometry to shape type
 */
function mapPresetGeometry(prst: string): PPTXShapeType {
  const mapping: Record<string, PPTXShapeType> = {
    rect: 'rect',
    roundRect: 'roundRect',
    ellipse: 'ellipse',
    triangle: 'triangle',
    rtTriangle: 'triangle',
    diamond: 'diamond',
    pentagon: 'pentagon',
    hexagon: 'hexagon',
    star4: 'star',
    star5: 'star',
    star6: 'star',
    star8: 'star',
    star10: 'star',
    star12: 'star',
    star16: 'star',
    star24: 'star',
    star32: 'star',
    rightArrow: 'arrow',
    leftArrow: 'arrow',
    upArrow: 'arrow',
    downArrow: 'arrow',
    line: 'line',
    straightConnector1: 'connector',
    bentConnector2: 'connector',
    bentConnector3: 'connector',
    bentConnector4: 'connector',
    bentConnector5: 'connector',
    curvedConnector2: 'connector',
    curvedConnector3: 'connector',
    curvedConnector4: 'connector',
    curvedConnector5: 'connector',
  };

  return mapping[prst] || 'rect';
}
