// Polyfill for setImmediate (required by JSZip in browser environments)
declare global {
  interface Window {
    setImmediate: (callback: (...args: any[]) => void, ...args: any[]) => number;
    clearImmediate: (handle: number) => void;
  }
}

if (typeof setImmediate === 'undefined') {
  (globalThis as any).setImmediate = (fn: (...args: any[]) => void, ...args: any[]) => {
    return setTimeout(() => fn(...args), 0) as unknown as number;
  };
  (globalThis as any).clearImmediate = (id: number) => {
    clearTimeout(id);
  };
}

import { parsePPTX, PPTXParseResult } from './parser/pptxParser';
import { convertSlideElements } from './converters/groupConverter';
import { PPTXSlide } from './types/pptx';

// Show UI
figma.showUI(__html__, {
  width: 440,
  height: 580,
  themeColors: true,
});

// Handle messages from UI
figma.ui.onmessage = async (msg) => {
  if (msg.type === 'import-pptx') {
    await handleImport(msg.data, msg.options);
  }
};

interface ImportOptions {
  separateFrames: boolean;
  preserveNames: boolean;
}

/**
 * Handle the PPTX import
 */
async function handleImport(data: number[], options: ImportOptions): Promise<void> {
  try {
    // Convert array back to ArrayBuffer
    const arrayBuffer = new Uint8Array(data).buffer;

    // Update progress
    sendProgress(15, 'Parsing PowerPoint file...');

    // Parse the PPTX file
    let parseResult: PPTXParseResult;
    try {
      parseResult = await parsePPTX(arrayBuffer);
    } catch (error) {
      throw new Error(`Failed to parse PPTX file: ${error}`);
    }

    if (parseResult.slides.length === 0) {
      throw new Error('No slides found in the PowerPoint file');
    }

    // Debug: Log parse results
    console.log('[DEBUG] Parse result:', JSON.stringify({
      slideCount: parseResult.slides.length,
      slideWidth: parseResult.slideWidth,
      slideHeight: parseResult.slideHeight,
      themeColors: parseResult.theme?.colors ? Object.fromEntries(parseResult.theme.colors) : null,
      slides: parseResult.slides.map((s, i) => ({
        index: i,
        elementCount: s.elements.length,
        elements: s.elements.map(e => ({
          type: e.type,
          name: e.name,
          x: e.x,
          y: e.y,
          width: e.width,
          height: e.height,
          // Include text info for debugging
          ...(e.type === 'text' ? { paragraphs: (e as any).paragraphs } : {}),
          ...(e.type === 'shape' && (e as any).text ? { text: (e as any).text } : {}),
        }))
      }))
    }, null, 2));

    sendProgress(30, `Found ${parseResult.slides.length} slide(s). Creating Figma elements...`);

    // Create slides in Figma
    const createdFrames: FrameNode[] = [];
    const totalSlides = parseResult.slides.length;

    for (let i = 0; i < totalSlides; i++) {
      const slide = parseResult.slides[i];
      const progressPercent = 30 + Math.floor((i / totalSlides) * 60);
      sendProgress(progressPercent, `Converting slide ${i + 1}/${totalSlides}...`);

      const frame = await createSlideFrame(
        slide,
        parseResult.slideWidth,
        parseResult.slideHeight,
        options,
        i
      );

      createdFrames.push(frame);
    }

    // Position frames in a row
    if (options.separateFrames && createdFrames.length > 1) {
      let offsetX = 0;
      for (const frame of createdFrames) {
        frame.x = offsetX;
        frame.y = 0;
        offsetX += frame.width + 50; // 50px gap between slides
      }
    }

    // Select the created frames
    figma.currentPage.selection = createdFrames;
    figma.viewport.scrollAndZoomIntoView(createdFrames);

    sendProgress(100, 'Import complete!');
    
    // Send success message
    figma.ui.postMessage({
      type: 'success',
      message: `Successfully imported ${totalSlides} slide(s) with ${countElements(parseResult.slides)} elements`,
    });
  } catch (error) {
    console.error('Import error:', error);
    figma.ui.postMessage({
      type: 'error',
      message: `Import failed: ${error instanceof Error ? error.message : 'Unknown error'}`,
    });
  }
}

/**
 * Create a Figma frame for a slide
 */
async function createSlideFrame(
  slide: PPTXSlide,
  width: number,
  height: number,
  options: ImportOptions,
  slideIndex: number
): Promise<FrameNode> {
  // Create the main slide frame
  const frame = figma.createFrame();
  frame.name = options.preserveNames ? `Slide ${slide.index}` : `Slide ${slideIndex + 1}`;
  frame.resize(width, height);

  // Set background
  if (slide.background) {
    if (slide.background.type === 'solid' && slide.background.color) {
      frame.fills = [{
        type: 'SOLID',
        color: hexToRgb(slide.background.color),
      }];
    } else if (slide.background.type === 'gradient' && slide.background.gradientStops) {
      const stops: ColorStop[] = slide.background.gradientStops.map((stop) => ({
        color: { ...hexToRgb(stop.color), a: 1 },
        position: stop.position,
      }));
      frame.fills = [{
        type: 'GRADIENT_LINEAR',
        gradientTransform: [[1, 0, 0], [0, 1, 0]],
        gradientStops: stops,
      }];
    }
  } else {
    // Default white background
    frame.fills = [{
      type: 'SOLID',
      color: { r: 1, g: 1, b: 1 },
    }];
  }

  // Convert all elements
  await convertSlideElements(
    slide.elements,
    frame,
    (current, total, message) => {
      // Optional: More granular progress updates
    }
  );

  return frame;
}

/**
 * Send progress update to UI
 */
function sendProgress(percent: number, text: string): void {
  figma.ui.postMessage({
    type: 'progress',
    percent,
    text,
  });
}

/**
 * Count total elements across all slides
 */
function countElements(slides: PPTXSlide[]): number {
  let count = 0;
  for (const slide of slides) {
    count += countSlideElements(slide.elements);
  }
  return count;
}

/**
 * Count elements recursively
 */
function countSlideElements(elements: any[]): number {
  let count = elements.length;
  for (const element of elements) {
    if (element.type === 'group' && element.children) {
      count += countSlideElements(element.children);
    }
  }
  return count;
}

/**
 * Convert hex color to RGB
 */
function hexToRgb(hex: string): RGB {
  const cleanHex = hex.replace('#', '');
  const r = parseInt(cleanHex.substring(0, 2), 16) / 255;
  const g = parseInt(cleanHex.substring(2, 4), 16) / 255;
  const b = parseInt(cleanHex.substring(4, 6), 16) / 255;
  return { r, g, b };
}

