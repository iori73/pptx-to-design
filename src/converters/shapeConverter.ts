import { PPTXShapeElement, PPTXShapeType, PPTXFill, PPTXStroke, PPTXParagraph } from '../types/pptx';
import { paragraphsToPlainText, getFirstTextStyle } from './textConverter';

// Default text color (black)
const DEFAULT_TEXT_COLOR = '#000000';
// Default font size in points
const DEFAULT_FONT_SIZE = 12;

/**
 * Convert PPTX shape element to Figma node
 */
export async function convertShapeElement(
  element: PPTXShapeElement,
  parent: FrameNode | GroupNode
): Promise<SceneNode> {
  let shapeNode: SceneNode;

  switch (element.shapeType) {
    case 'rect':
    case 'roundRect':
      shapeNode = createRectangle(element);
      break;
    case 'ellipse':
      shapeNode = createEllipse(element);
      break;
    case 'triangle':
      shapeNode = createTriangle(element);
      break;
    case 'diamond':
      shapeNode = createDiamond(element);
      break;
    case 'pentagon':
      shapeNode = createPentagon(element);
      break;
    case 'hexagon':
      shapeNode = createHexagon(element);
      break;
    case 'star':
      shapeNode = createStar(element);
      break;
    case 'arrow':
      shapeNode = createArrow(element);
      break;
    case 'line':
    case 'connector':
      shapeNode = createLine(element);
      break;
    default:
      shapeNode = createRectangle(element);
  }

  // Apply position
  shapeNode.x = element.x;
  shapeNode.y = element.y;

  // Apply rotation if present
  if (element.rotation) {
    shapeNode.rotation = -element.rotation; // Figma uses opposite direction
  }

  // If shape has text, create a group with shape and text
  if (element.text && element.text.length > 0) {
    const text = paragraphsToPlainText(element.text);
    if (text.trim()) {
      const group = await createShapeWithText(shapeNode, element, text);
      parent.appendChild(group);
      return group;
    }
  }

  parent.appendChild(shapeNode);
  return shapeNode;
}

/**
 * Create a rectangle shape
 */
function createRectangle(element: PPTXShapeElement): RectangleNode {
  const rect = figma.createRectangle();
  rect.name = element.name || 'Rectangle';
  rect.resize(element.width, element.height);

  // Apply corner radius
  if (element.shapeType === 'roundRect' && element.cornerRadius) {
    rect.cornerRadius = element.cornerRadius;
  }

  // Apply fill
  applyFill(rect, element.fill);

  // Apply stroke
  applyStroke(rect, element.stroke);

  return rect;
}

/**
 * Create an ellipse shape
 */
function createEllipse(element: PPTXShapeElement): EllipseNode {
  const ellipse = figma.createEllipse();
  ellipse.name = element.name || 'Ellipse';
  ellipse.resize(element.width, element.height);

  applyFill(ellipse, element.fill);
  applyStroke(ellipse, element.stroke);

  return ellipse;
}

/**
 * Create a triangle shape using polygon
 */
function createTriangle(element: PPTXShapeElement): PolygonNode {
  const polygon = figma.createPolygon();
  polygon.name = element.name || 'Triangle';
  polygon.pointCount = 3;
  polygon.resize(element.width, element.height);

  applyFill(polygon, element.fill);
  applyStroke(polygon, element.stroke);

  return polygon;
}

/**
 * Create a diamond shape
 */
function createDiamond(element: PPTXShapeElement): PolygonNode {
  const polygon = figma.createPolygon();
  polygon.name = element.name || 'Diamond';
  polygon.pointCount = 4;
  polygon.resize(element.width, element.height);
  polygon.rotation = 45;

  applyFill(polygon, element.fill);
  applyStroke(polygon, element.stroke);

  return polygon;
}

/**
 * Create a pentagon shape
 */
function createPentagon(element: PPTXShapeElement): PolygonNode {
  const polygon = figma.createPolygon();
  polygon.name = element.name || 'Pentagon';
  polygon.pointCount = 5;
  polygon.resize(element.width, element.height);

  applyFill(polygon, element.fill);
  applyStroke(polygon, element.stroke);

  return polygon;
}

/**
 * Create a hexagon shape
 */
function createHexagon(element: PPTXShapeElement): PolygonNode {
  const polygon = figma.createPolygon();
  polygon.name = element.name || 'Hexagon';
  polygon.pointCount = 6;
  polygon.resize(element.width, element.height);

  applyFill(polygon, element.fill);
  applyStroke(polygon, element.stroke);

  return polygon;
}

/**
 * Create a star shape
 */
function createStar(element: PPTXShapeElement): StarNode {
  const star = figma.createStar();
  star.name = element.name || 'Star';
  star.pointCount = 5;
  star.innerRadius = 0.4;
  star.resize(element.width, element.height);

  applyFill(star, element.fill);
  applyStroke(star, element.stroke);

  return star;
}

/**
 * Create an arrow shape using vector
 */
function createArrow(element: PPTXShapeElement): FrameNode {
  // Create a frame to contain the arrow
  const frame = figma.createFrame();
  frame.name = element.name || 'Arrow';
  frame.resize(element.width, element.height);
  frame.fills = [];
  frame.clipsContent = false;

  // Create arrow body
  const body = figma.createRectangle();
  body.resize(element.width * 0.7, element.height * 0.4);
  body.x = 0;
  body.y = element.height * 0.3;

  // Create arrow head
  const head = figma.createPolygon();
  head.pointCount = 3;
  head.resize(element.width * 0.4, element.height);
  head.x = element.width * 0.6;
  head.y = 0;
  head.rotation = 90;

  applyFill(body, element.fill);
  applyFill(head, element.fill);
  applyStroke(body, element.stroke);
  applyStroke(head, element.stroke);

  frame.appendChild(body);
  frame.appendChild(head);

  return frame;
}

/**
 * Create a line shape
 */
function createLine(element: PPTXShapeElement): LineNode {
  const line = figma.createLine();
  line.name = element.name || 'Line';
  line.resize(element.width, 0);

  // Apply stroke
  if (element.stroke) {
    line.strokes = [{
      type: 'SOLID',
      color: hexToRgb(element.stroke.color),
      opacity: element.stroke.opacity ?? 1,
    }];
    line.strokeWeight = element.stroke.width;

    // Apply dash pattern
    if (element.stroke.dashPattern === 'dash') {
      line.dashPattern = [10, 5];
    } else if (element.stroke.dashPattern === 'dot') {
      line.dashPattern = [2, 2];
    } else if (element.stroke.dashPattern === 'dashDot') {
      line.dashPattern = [10, 5, 2, 5];
    }
  } else {
    line.strokes = [{ type: 'SOLID', color: { r: 0, g: 0, b: 0 } }];
    line.strokeWeight = 1;
  }

  return line;
}

/**
 * Create a shape with text inside - IMPROVED VERSION
 */
async function createShapeWithText(
  shapeNode: SceneNode,
  element: PPTXShapeElement,
  text: string
): Promise<GroupNode> {
  // Create text node
  const textNode = figma.createText();
  textNode.name = 'Text';

  // Get text style from paragraphs
  const textStyle = element.text ? getFirstTextStyle(element.text) : {};

  // Load font
  const fontStyle = textStyle.fontWeight === 'bold' ? 'Bold' : 'Regular';
  try {
    await figma.loadFontAsync({ family: 'Inter', style: fontStyle });
  } catch {
    await figma.loadFontAsync({ family: 'Inter', style: 'Regular' });
  }
  
  textNode.characters = text;
  
  // Apply font size
  if (textStyle.fontSize) {
    textNode.fontSize = textStyle.fontSize;
  } else {
    textNode.fontSize = DEFAULT_FONT_SIZE;
  }
  
  // Apply text color - use default if not specified
  const textColor = textStyle.color || DEFAULT_TEXT_COLOR;
  textNode.fills = [{ type: 'SOLID', color: hexToRgb(textColor) }];
  
  // Apply alignment
  if (textStyle.alignment) {
    textNode.textAlignHorizontal = mapAlignment(textStyle.alignment);
  } else {
    textNode.textAlignHorizontal = 'CENTER';
  }
  textNode.textAlignVertical = 'CENTER';
  
  // Set text box size
  // Calculate safe size for text box to avoid negative dimensions
  const minSize = 1; // Minimum size to avoid errors
  const paddingX = Math.min(8, element.width / 4);
  const paddingY = Math.min(8, element.height / 4);
  
  const textBoxWidth = Math.max(minSize, element.width - (paddingX * 2));
  const textBoxHeight = Math.max(minSize, element.height - (paddingY * 2));
  
  textNode.resize(textBoxWidth, textBoxHeight);
  textNode.textAutoResize = 'HEIGHT';

  // Position text centered in shape
  textNode.x = element.x + (element.width - textNode.width) / 2;
  textNode.y = element.y + (element.height - textNode.height) / 2;

  // Group shape and text
  const group = figma.group([shapeNode, textNode], figma.currentPage);
  group.name = element.name || 'Shape with Text';

  return group;
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
 * Apply fill to a shape
 */
function applyFill(
  node: GeometryMixin & MinimalFillsMixin,
  fill?: PPTXFill
): void {
  if (!fill || fill.type === 'none') {
    node.fills = [];
    return;
  }

  if (fill.type === 'solid' && fill.color) {
    node.fills = [{
      type: 'SOLID',
      color: hexToRgb(fill.color),
      opacity: fill.opacity ?? 1,
    }];
  } else if (fill.type === 'gradient' && fill.gradientStops) {
    // Create gradient fill
    const stops: ColorStop[] = fill.gradientStops.map((stop) => ({
      color: { ...hexToRgb(stop.color), a: stop.opacity ?? 1 },
      position: stop.position,
    }));

    node.fills = [{
      type: 'GRADIENT_LINEAR',
      gradientTransform: [
        [Math.cos((fill.gradientAngle || 0) * Math.PI / 180), -Math.sin((fill.gradientAngle || 0) * Math.PI / 180), 0],
        [Math.sin((fill.gradientAngle || 0) * Math.PI / 180), Math.cos((fill.gradientAngle || 0) * Math.PI / 180), 0],
      ],
      gradientStops: stops,
    }];
  } else if (fill.color) {
    // Fallback to solid color
    node.fills = [{
      type: 'SOLID',
      color: hexToRgb(fill.color),
      opacity: fill.opacity ?? 1,
    }];
  }
}

/**
 * Apply stroke to a shape
 */
function applyStroke(
  node: GeometryMixin & MinimalStrokesMixin,
  stroke?: PPTXStroke
): void {
  if (!stroke) {
    node.strokes = [];
    return;
  }

  node.strokes = [{
    type: 'SOLID',
    color: hexToRgb(stroke.color),
    opacity: stroke.opacity ?? 1,
  }];
  node.strokeWeight = stroke.width;

  // Apply dash pattern
  if (stroke.dashPattern === 'dash') {
    node.dashPattern = [10, 5];
  } else if (stroke.dashPattern === 'dot') {
    node.dashPattern = [2, 2];
  } else if (stroke.dashPattern === 'dashDot') {
    node.dashPattern = [10, 5, 2, 5];
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
