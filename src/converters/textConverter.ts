import { PPTXTextElement, PPTXParagraph, PPTXTextRun } from '../types/pptx';

// Default text color (black)
const DEFAULT_TEXT_COLOR = '#000000';
// Default font size in points
const DEFAULT_FONT_SIZE = 12;

/**
 * Convert PPTX text element to Figma text node
 */
export async function convertTextElement(
  element: PPTXTextElement,
  parent: FrameNode | GroupNode
): Promise<TextNode | FrameNode> {
  // If there's a background color, create a frame with text inside
  if (element.backgroundColor) {
    const frame = figma.createFrame();
    frame.name = element.name || 'Text Box';
    frame.x = element.x;
    frame.y = element.y;
    frame.resize(element.width, element.height);
    frame.fills = [{ type: 'SOLID', color: hexToRgb(element.backgroundColor) }];
    frame.clipsContent = false;

    const textNode = await createTextNode(element);
    textNode.x = 8; // Padding
    textNode.y = 8;
    frame.appendChild(textNode);

    parent.appendChild(frame);
    return frame;
  }

  // Simple text node
  const textNode = await createTextNode(element);
  textNode.x = element.x;
  textNode.y = element.y;
  
  parent.appendChild(textNode);
  return textNode;
}

/**
 * Create a text node from PPTX text element
 */
async function createTextNode(element: PPTXTextElement): Promise<TextNode> {
  const textNode = figma.createText();
  textNode.name = element.name || 'Text';

  // Build full text content and collect style ranges
  const { fullText, styleRanges } = buildTextContent(element.paragraphs);

  if (!fullText.trim()) {
    // Empty text - set minimal content
    await figma.loadFontAsync({ family: 'Inter', style: 'Regular' });
    textNode.characters = ' ';
    textNode.resize(element.width, element.height);
    return textNode;
  }

  // Load all required fonts first
  const fontsToLoad = new Set<string>();
  for (const range of styleRanges) {
    const fontFamily = range.fontFamily || 'Inter';
    const fontStyle = getFontStyle(range.fontWeight, range.fontStyle);
    fontsToLoad.add(`${fontFamily}|${fontStyle}`);
  }
  // Always ensure Inter Regular is loaded as fallback
  fontsToLoad.add('Inter|Regular');

  for (const fontKey of fontsToLoad) {
    const [family, style] = fontKey.split('|');
    try {
      await figma.loadFontAsync({ family, style });
    } catch {
      // Fallback to Inter if font not available
      try {
        await figma.loadFontAsync({ family: 'Inter', style });
      } catch {
        await figma.loadFontAsync({ family: 'Inter', style: 'Regular' });
      }
    }
  }

  // Set the text content
  textNode.characters = fullText;

  // Apply styles to each range
  for (const range of styleRanges) {
    const { start, end, fontFamily, fontSize, fontWeight, fontStyle, color, textDecoration } = range;

    if (start >= end) continue;

    // Font family and style
    const family = fontFamily || 'Inter';
    const style = getFontStyle(fontWeight, fontStyle);
    
    try {
      textNode.setRangeFontName(start, end, { family, style });
    } catch {
      try {
        textNode.setRangeFontName(start, end, { family: 'Inter', style });
      } catch {
        textNode.setRangeFontName(start, end, { family: 'Inter', style: 'Regular' });
      }
    }

    // Font size - use default if not specified
    const finalFontSize = fontSize || DEFAULT_FONT_SIZE;
    textNode.setRangeFontSize(start, end, finalFontSize);

    // Color - use default if not specified
    const finalColor = color || DEFAULT_TEXT_COLOR;
    textNode.setRangeFills(start, end, [{ type: 'SOLID', color: hexToRgb(finalColor) }]);

    // Text decoration
    if (textDecoration === 'underline') {
      textNode.setRangeTextDecoration(start, end, 'UNDERLINE');
    } else if (textDecoration === 'strikethrough') {
      textNode.setRangeTextDecoration(start, end, 'STRIKETHROUGH');
    }
  }

  // Set text box size
  textNode.resize(element.width, element.height);
  textNode.textAutoResize = 'HEIGHT';

  // Set paragraph alignment (use first paragraph's alignment as default)
  if (element.paragraphs.length > 0 && element.paragraphs[0].alignment) {
    textNode.textAlignHorizontal = mapAlignment(element.paragraphs[0].alignment);
  }

  return textNode;
}

interface StyleRange {
  start: number;
  end: number;
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: 'normal' | 'bold';
  fontStyle?: 'normal' | 'italic';
  color?: string;
  textDecoration?: 'none' | 'underline' | 'strikethrough';
}

/**
 * Build full text content and style ranges from paragraphs
 */
function buildTextContent(paragraphs: PPTXParagraph[]): {
  fullText: string;
  styleRanges: StyleRange[];
} {
  let fullText = '';
  const styleRanges: StyleRange[] = [];
  let currentPosition = 0;

  for (let pIndex = 0; pIndex < paragraphs.length; pIndex++) {
    const paragraph = paragraphs[pIndex];

    // Add bullet if needed
    if (paragraph.bulletType === 'bullet') {
      const indent = '  '.repeat(paragraph.indentLevel || 0);
      const bulletText = `${indent}• `;
      fullText += bulletText;
      
      // Add style range for bullet (use first run's style if available)
      const bulletStyle = paragraph.runs[0] || {};
      styleRanges.push({
        start: currentPosition,
        end: currentPosition + bulletText.length,
        fontFamily: bulletStyle.fontFamily,
        fontSize: bulletStyle.fontSize,
        fontWeight: bulletStyle.fontWeight,
        fontStyle: bulletStyle.fontStyle,
        color: bulletStyle.color,
      });
      
      currentPosition += bulletText.length;
    } else if (paragraph.bulletType === 'number') {
      const indent = '  '.repeat(paragraph.indentLevel || 0);
      const bulletText = `${indent}${pIndex + 1}. `;
      fullText += bulletText;
      
      // Add style range for number (use first run's style if available)
      const bulletStyle = paragraph.runs[0] || {};
      styleRanges.push({
        start: currentPosition,
        end: currentPosition + bulletText.length,
        fontFamily: bulletStyle.fontFamily,
        fontSize: bulletStyle.fontSize,
        fontWeight: bulletStyle.fontWeight,
        fontStyle: bulletStyle.fontStyle,
        color: bulletStyle.color,
      });
      
      currentPosition += bulletText.length;
    }

    // Process runs
    for (const run of paragraph.runs) {
      const start = currentPosition;
      const text = run.text;
      fullText += text;
      currentPosition += text.length;

      styleRanges.push({
        start,
        end: currentPosition,
        fontFamily: run.fontFamily,
        fontSize: run.fontSize,
        fontWeight: run.fontWeight,
        fontStyle: run.fontStyle,
        color: run.color,
        textDecoration: run.textDecoration,
      });
    }

    // Add newline between paragraphs (except for last)
    if (pIndex < paragraphs.length - 1) {
      fullText += '\n';
      currentPosition += 1;
    }
  }

  return { fullText, styleRanges };
}

/**
 * Get Figma font style string
 */
function getFontStyle(
  weight?: 'normal' | 'bold',
  style?: 'normal' | 'italic'
): string {
  const isBold = weight === 'bold';
  const isItalic = style === 'italic';

  if (isBold && isItalic) return 'Bold Italic';
  if (isBold) return 'Bold';
  if (isItalic) return 'Italic';
  return 'Regular';
}

/**
 * Map alignment to Figma format
 */
function mapAlignment(alignment: 'left' | 'center' | 'right' | 'justify'): 'LEFT' | 'CENTER' | 'RIGHT' | 'JUSTIFIED' {
  switch (alignment) {
    case 'left': return 'LEFT';
    case 'center': return 'CENTER';
    case 'right': return 'RIGHT';
    case 'justify': return 'JUSTIFIED';
  }
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

/**
 * Convert paragraphs to plain text (for shapes with text)
 */
export function paragraphsToPlainText(paragraphs: PPTXParagraph[]): string {
  return paragraphs
    .map((p) => p.runs.map((r) => r.text).join(''))
    .join('\n');
}

/**
 * Get first text style from paragraphs (for shapes with text)
 */
export function getFirstTextStyle(paragraphs: PPTXParagraph[]): {
  fontSize?: number;
  color?: string;
  fontWeight?: 'normal' | 'bold';
  alignment?: 'left' | 'center' | 'right' | 'justify';
} {
  if (paragraphs.length === 0) return {};
  
  const firstParagraph = paragraphs[0];
  const firstRun = firstParagraph.runs[0];
  
  return {
    fontSize: firstRun?.fontSize,
    color: firstRun?.color || DEFAULT_TEXT_COLOR,
    fontWeight: firstRun?.fontWeight,
    alignment: firstParagraph.alignment,
  };
}
