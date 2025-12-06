// PowerPoint element types

export interface PPTXSlide {
  index: number;
  width: number;
  height: number;
  elements: PPTXElement[];
  background?: PPTXBackground;
}

export interface PPTXBackground {
  type: 'solid' | 'gradient' | 'image';
  color?: string;
  gradientStops?: Array<{ color: string; position: number }>;
  imageData?: Uint8Array;
}

export type PPTXElement = 
  | PPTXTextElement 
  | PPTXShapeElement 
  | PPTXImageElement 
  | PPTXTableElement 
  | PPTXChartElement 
  | PPTXGroupElement;

export interface PPTXBaseElement {
  type: string;
  x: number;
  y: number;
  width: number;
  height: number;
  rotation?: number;
  name?: string;
}

export interface PPTXTextElement extends PPTXBaseElement {
  type: 'text';
  paragraphs: PPTXParagraph[];
  backgroundColor?: string;
  borderColor?: string;
  borderWidth?: number;
}

export interface PPTXParagraph {
  runs: PPTXTextRun[];
  alignment?: 'left' | 'center' | 'right' | 'justify';
  lineSpacing?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  bulletType?: 'none' | 'bullet' | 'number';
  bulletChar?: string;
  indentLevel?: number;
}

export interface PPTXTextRun {
  text: string;
  fontFamily?: string;
  fontSize?: number;
  fontWeight?: 'normal' | 'bold';
  fontStyle?: 'normal' | 'italic';
  textDecoration?: 'none' | 'underline' | 'strikethrough';
  color?: string;
  highlight?: string;
}

export interface PPTXShapeElement extends PPTXBaseElement {
  type: 'shape';
  shapeType: PPTXShapeType;
  fill?: PPTXFill;
  stroke?: PPTXStroke;
  text?: PPTXParagraph[];
  cornerRadius?: number;
  points?: Array<{ x: number; y: number }>;
}

export type PPTXShapeType = 
  | 'rect' 
  | 'roundRect' 
  | 'ellipse' 
  | 'triangle' 
  | 'diamond'
  | 'pentagon'
  | 'hexagon'
  | 'star'
  | 'arrow'
  | 'line'
  | 'connector'
  | 'custom';

export interface PPTXFill {
  type: 'solid' | 'gradient' | 'none';
  color?: string;
  opacity?: number;
  gradientStops?: Array<{ color: string; position: number; opacity?: number }>;
  gradientAngle?: number;
}

export interface PPTXStroke {
  color: string;
  width: number;
  opacity?: number;
  dashPattern?: 'solid' | 'dash' | 'dot' | 'dashDot';
}

export interface PPTXImageElement extends PPTXBaseElement {
  type: 'image';
  imageData: Uint8Array;
  mimeType: string;
  cropRect?: { x: number; y: number; width: number; height: number };
}

export interface PPTXTableElement extends PPTXBaseElement {
  type: 'table';
  rows: PPTXTableRow[];
  columnWidths: number[];
}

export interface PPTXTableRow {
  height: number;
  cells: PPTXTableCell[];
}

export interface PPTXTableCell {
  content: PPTXParagraph[];
  rowSpan?: number;
  colSpan?: number;
  backgroundColor?: string;
  borderTop?: PPTXStroke;
  borderRight?: PPTXStroke;
  borderBottom?: PPTXStroke;
  borderLeft?: PPTXStroke;
  padding?: { top: number; right: number; bottom: number; left: number };
  verticalAlign?: 'top' | 'middle' | 'bottom';
}

export interface PPTXChartElement extends PPTXBaseElement {
  type: 'chart';
  chartType: 'bar' | 'line' | 'pie' | 'area' | 'scatter' | 'other';
  title?: string;
  // Chart data is complex, we'll render as image for now
  imageData?: Uint8Array;
}

export interface PPTXGroupElement extends PPTXBaseElement {
  type: 'group';
  children: PPTXElement[];
}

// Conversion context
export interface ConversionContext {
  slideWidth: number;
  slideHeight: number;
  theme?: PPTXTheme;
  mediaFiles: Map<string, Uint8Array>;
  relationships: Map<string, string>;
}

export interface PPTXTheme {
  name: string;
  colors: Map<string, string>;
  fonts: {
    majorFont?: string;
    minorFont?: string;
  };
}

// EMU to pixel conversion constants
export const EMU_PER_INCH = 914400;
export const PIXELS_PER_INCH = 72;
export const EMU_TO_PIXEL = PIXELS_PER_INCH / EMU_PER_INCH;

// Helper function to convert EMU to pixels
export function emuToPixels(emu: number): number {
  return emu * EMU_TO_PIXEL;
}

// Helper function to convert points to pixels
export function pointsToPixels(points: number): number {
  return points * (PIXELS_PER_INCH / 72);
}

// Helper function to convert hundredths of a point to pixels
export function hundredthsToPixels(hundredths: number): number {
  return (hundredths / 100) * (PIXELS_PER_INCH / 72);
}

