import { PPTXTableElement, PPTXTableRow, PPTXTableCell, PPTXParagraph } from '../types/pptx';
import { paragraphsToPlainText, getFirstTextStyle } from './textConverter';

// Default text color (black)
const DEFAULT_TEXT_COLOR = '#000000';
// Default font size in points
const DEFAULT_FONT_SIZE = 10;

/**
 * Convert PPTX table element to Figma frame structure
 */
export async function convertTableElement(
  element: PPTXTableElement,
  parent: FrameNode | GroupNode
): Promise<FrameNode> {
  // Create main table frame
  const tableFrame = figma.createFrame();
  tableFrame.name = element.name || 'Table';
  tableFrame.x = element.x;
  tableFrame.y = element.y;
  tableFrame.resize(element.width, element.height);
  tableFrame.fills = [];
  tableFrame.clipsContent = false;

  // Calculate cell positions
  const columnWidths = element.columnWidths.length > 0
    ? element.columnWidths
    : calculateEvenColumnWidths(element.width, element.rows[0]?.cells.length || 1);

  let currentY = 0;

  for (let rowIndex = 0; rowIndex < element.rows.length; rowIndex++) {
    const row = element.rows[rowIndex];
    const rowHeight = row.height || element.height / element.rows.length;
    let currentX = 0;

    for (let colIndex = 0; colIndex < row.cells.length; colIndex++) {
      const cell = row.cells[colIndex];
      const cellWidth = getCellWidth(columnWidths, colIndex, cell.colSpan || 1);
      const cellHeight = getCellHeight(element.rows, rowIndex, rowHeight, cell.rowSpan || 1);

      // Create cell frame
      const cellFrame = await createCellFrame(cell, cellWidth, cellHeight);
      cellFrame.x = currentX;
      cellFrame.y = currentY;

      tableFrame.appendChild(cellFrame);

      currentX += cellWidth;
    }

    currentY += rowHeight;
  }

  // Apply rotation if present
  if (element.rotation) {
    tableFrame.rotation = -element.rotation;
  }

  parent.appendChild(tableFrame);
  return tableFrame;
}

/**
 * Create a table cell frame - IMPROVED VERSION
 */
async function createCellFrame(
  cell: PPTXTableCell,
  width: number,
  height: number
): Promise<FrameNode> {
  const cellFrame = figma.createFrame();
  cellFrame.name = 'Cell';
  cellFrame.resize(width, height);

  // Apply background color
  if (cell.backgroundColor) {
    cellFrame.fills = [{
      type: 'SOLID',
      color: hexToRgb(cell.backgroundColor),
    }];
  } else {
    cellFrame.fills = [{
      type: 'SOLID',
      color: { r: 1, g: 1, b: 1 },
    }];
  }

  // Apply borders
  const strokeWeight = 1;
  cellFrame.strokes = [{
    type: 'SOLID',
    color: { r: 0.8, g: 0.8, b: 0.8 },
  }];
  cellFrame.strokeWeight = strokeWeight;

  // Add cell content
  if (cell.content && cell.content.length > 0) {
    const text = paragraphsToPlainText(cell.content);
    if (text.trim()) {
      const textNode = figma.createText();
      textNode.name = 'Cell Text';

      // Get text style from cell content
      const textStyle = getFirstTextStyle(cell.content);

      // Load font
      const fontStyle = textStyle.fontWeight === 'bold' ? 'Bold' : 'Regular';
      try {
        await figma.loadFontAsync({ family: 'Inter', style: fontStyle });
        if (textStyle.fontWeight === 'bold') {
          textNode.fontName = { family: 'Inter', style: 'Bold' };
        }
      } catch {
        await figma.loadFontAsync({ family: 'Inter', style: 'Regular' });
      }
      
      textNode.characters = text;

      // Apply font size
      textNode.fontSize = textStyle.fontSize || DEFAULT_FONT_SIZE;
      
      // Apply text color - use default if not specified
      const textColor = textStyle.color || DEFAULT_TEXT_COLOR;
      textNode.fills = [{ type: 'SOLID', color: hexToRgb(textColor) }];

      // Set text alignment
      if (textStyle.alignment) {
        textNode.textAlignHorizontal = mapAlignment(textStyle.alignment);
      } else if (cell.content[0]?.alignment) {
        textNode.textAlignHorizontal = mapAlignment(cell.content[0].alignment);
      }

      // Vertical alignment
      textNode.textAlignVertical = mapVerticalAlign(cell.verticalAlign);

      // Set size and position with padding
      const padding = cell.padding || { top: 4, right: 4, bottom: 4, left: 4 };
      
      // Ensure content dimensions are non-negative
      const minContentSize = 1;
      const contentWidth = Math.max(minContentSize, width - padding.left - padding.right);
      const contentHeight = Math.max(minContentSize, height - padding.top - padding.bottom);

      textNode.resize(contentWidth, contentHeight);
      textNode.textAutoResize = 'HEIGHT';
      textNode.x = padding.left;
      textNode.y = padding.top;

      cellFrame.appendChild(textNode);
    }
  }

  return cellFrame;
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
 * Calculate even column widths
 */
function calculateEvenColumnWidths(tableWidth: number, columnCount: number): number[] {
  const width = tableWidth / columnCount;
  return Array(columnCount).fill(width);
}

/**
 * Get cell width considering column span
 */
function getCellWidth(columnWidths: number[], startCol: number, colSpan: number): number {
  let width = 0;
  for (let i = startCol; i < startCol + colSpan && i < columnWidths.length; i++) {
    width += columnWidths[i];
  }
  return width;
}

/**
 * Get cell height considering row span
 */
function getCellHeight(
  rows: PPTXTableRow[],
  startRow: number,
  defaultHeight: number,
  rowSpan: number
): number {
  let height = 0;
  for (let i = startRow; i < startRow + rowSpan && i < rows.length; i++) {
    height += rows[i].height || defaultHeight;
  }
  return height;
}

/**
 * Map vertical alignment
 */
function mapVerticalAlign(align?: 'top' | 'middle' | 'bottom'): 'TOP' | 'CENTER' | 'BOTTOM' {
  switch (align) {
    case 'middle': return 'CENTER';
    case 'bottom': return 'BOTTOM';
    default: return 'TOP';
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
