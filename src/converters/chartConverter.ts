import { PPTXChartElement } from '../types/pptx';

/**
 * Convert PPTX chart element to Figma
 * Charts are complex, so we create a placeholder with chart info
 */
export async function convertChartElement(
  element: PPTXChartElement,
  parent: FrameNode | GroupNode
): Promise<FrameNode> {
  // If we have image data, use it
  if (element.imageData) {
    return await createChartFromImage(element, parent);
  }

  // Otherwise create a placeholder
  return createChartPlaceholder(element, parent);
}

/**
 * Create chart from image data
 */
async function createChartFromImage(
  element: PPTXChartElement,
  parent: FrameNode | GroupNode
): Promise<FrameNode> {
  const frame = figma.createFrame();
  frame.name = element.name || `Chart (${element.chartType})`;
  frame.x = element.x;
  frame.y = element.y;
  frame.resize(element.width, element.height);

  try {
    const image = figma.createImage(element.imageData!);
    frame.fills = [{
      type: 'IMAGE',
      imageHash: image.hash,
      scaleMode: 'FILL',
    }];
  } catch (error) {
    console.error('Error creating chart image:', error);
    // Fallback to placeholder
    return createChartPlaceholder(element, parent);
  }

  if (element.rotation) {
    frame.rotation = -element.rotation;
  }

  parent.appendChild(frame);
  return frame;
}

/**
 * Create a chart placeholder
 */
function createChartPlaceholder(
  element: PPTXChartElement,
  parent: FrameNode | GroupNode
): FrameNode {
  const frame = figma.createFrame();
  frame.name = element.name || `Chart (${element.chartType})`;
  frame.x = element.x;
  frame.y = element.y;
  frame.resize(element.width, element.height);
  frame.fills = [{
    type: 'SOLID',
    color: { r: 0.97, g: 0.97, b: 0.97 },
  }];
  frame.strokes = [{
    type: 'SOLID',
    color: { r: 0.85, g: 0.85, b: 0.85 },
  }];
  frame.strokeWeight = 1;
  frame.cornerRadius = 8;

  // Create chart icon based on type
  const icon = createChartIcon(element.chartType, element.width, element.height);
  frame.appendChild(icon);

  // Add chart type label
  createChartLabel(frame, element);

  if (element.rotation) {
    frame.rotation = -element.rotation;
  }

  parent.appendChild(frame);
  return frame;
}

/**
 * Create a visual icon representing the chart type
 */
function createChartIcon(
  chartType: string,
  containerWidth: number,
  containerHeight: number
): FrameNode {
  const iconContainer = figma.createFrame();
  iconContainer.name = 'Chart Icon';
  iconContainer.fills = [];
  
  const iconSize = Math.min(containerWidth, containerHeight) * 0.4;
  iconContainer.resize(iconSize, iconSize);
  iconContainer.x = (containerWidth - iconSize) / 2;
  iconContainer.y = (containerHeight - iconSize) / 2 - 20;

  switch (chartType) {
    case 'bar':
      createBarChartIcon(iconContainer, iconSize);
      break;
    case 'line':
      createLineChartIcon(iconContainer, iconSize);
      break;
    case 'pie':
      createPieChartIcon(iconContainer, iconSize);
      break;
    default:
      createGenericChartIcon(iconContainer, iconSize);
  }

  return iconContainer;
}

/**
 * Create bar chart icon
 */
function createBarChartIcon(container: FrameNode, size: number): void {
  const colors = [
    { r: 0.3, g: 0.5, b: 0.9 },
    { r: 0.4, g: 0.7, b: 0.4 },
    { r: 0.9, g: 0.6, b: 0.3 },
  ];

  const barWidth = size / 5;
  const heights = [0.6, 0.9, 0.45];
  const gap = (size - barWidth * 3) / 4;

  heights.forEach((heightRatio, index) => {
    const bar = figma.createRectangle();
    bar.resize(barWidth, size * heightRatio);
    bar.x = gap + (barWidth + gap) * index;
    bar.y = size - size * heightRatio;
    bar.fills = [{ type: 'SOLID', color: colors[index] }];
    bar.cornerRadius = 2;
    container.appendChild(bar);
  });
}

/**
 * Create line chart icon
 */
function createLineChartIcon(container: FrameNode, size: number): void {
  // Create axes
  const xAxis = figma.createLine();
  xAxis.resize(size * 0.8, 0);
  xAxis.x = size * 0.1;
  xAxis.y = size * 0.85;
  xAxis.strokes = [{ type: 'SOLID', color: { r: 0.6, g: 0.6, b: 0.6 } }];
  xAxis.strokeWeight = 2;
  container.appendChild(xAxis);

  const yAxis = figma.createLine();
  yAxis.resize(0, size * 0.7);
  yAxis.x = size * 0.1;
  yAxis.y = size * 0.15;
  yAxis.strokes = [{ type: 'SOLID', color: { r: 0.6, g: 0.6, b: 0.6 } }];
  yAxis.strokeWeight = 2;
  yAxis.rotation = -90;
  container.appendChild(yAxis);

  // Create line data points
  const points = [
    { x: 0.15, y: 0.7 },
    { x: 0.35, y: 0.4 },
    { x: 0.55, y: 0.55 },
    { x: 0.75, y: 0.25 },
  ];

  // Create dots at each point
  points.forEach((point) => {
    const dot = figma.createEllipse();
    dot.resize(8, 8);
    dot.x = size * point.x - 4;
    dot.y = size * point.y - 4;
    dot.fills = [{ type: 'SOLID', color: { r: 0.3, g: 0.5, b: 0.9 } }];
    container.appendChild(dot);
  });
}

/**
 * Create pie chart icon
 */
function createPieChartIcon(container: FrameNode, size: number): void {
  const radius = size * 0.4;
  const centerX = size / 2;
  const centerY = size / 2;

  // Create a simple pie representation using ellipse
  const pie = figma.createEllipse();
  pie.resize(radius * 2, radius * 2);
  pie.x = centerX - radius;
  pie.y = centerY - radius;
  pie.fills = [{ type: 'SOLID', color: { r: 0.3, g: 0.5, b: 0.9 } }];
  container.appendChild(pie);

  // Add a slice indicator
  const slice = figma.createEllipse();
  slice.resize(radius * 2, radius * 2);
  slice.x = centerX - radius;
  slice.y = centerY - radius;
  slice.arcData = {
    startingAngle: 0,
    endingAngle: Math.PI * 0.7,
    innerRadius: 0,
  };
  slice.fills = [{ type: 'SOLID', color: { r: 0.4, g: 0.7, b: 0.4 } }];
  container.appendChild(slice);

  // Add another slice
  const slice2 = figma.createEllipse();
  slice2.resize(radius * 2, radius * 2);
  slice2.x = centerX - radius;
  slice2.y = centerY - radius;
  slice2.arcData = {
    startingAngle: Math.PI * 0.7,
    endingAngle: Math.PI * 1.2,
    innerRadius: 0,
  };
  slice2.fills = [{ type: 'SOLID', color: { r: 0.9, g: 0.6, b: 0.3 } }];
  container.appendChild(slice2);
}

/**
 * Create generic chart icon
 */
function createGenericChartIcon(container: FrameNode, size: number): void {
  // Just create a simple representation
  createBarChartIcon(container, size);
}

/**
 * Create chart label
 */
async function createChartLabel(frame: FrameNode, element: PPTXChartElement): Promise<void> {
  const textNode = figma.createText();
  textNode.name = 'Label';

  await figma.loadFontAsync({ family: 'Inter', style: 'Medium' });
  
  const label = element.title || `${capitalizeFirst(element.chartType)} Chart`;
  textNode.characters = label;
  textNode.fontSize = 12;
  textNode.fontName = { family: 'Inter', style: 'Medium' };
  textNode.fills = [{ type: 'SOLID', color: { r: 0.5, g: 0.5, b: 0.5 } }];
  textNode.textAlignHorizontal = 'CENTER';
  
  textNode.x = (element.width - textNode.width) / 2;
  textNode.y = element.height - 30;
  
  frame.appendChild(textNode);
}

/**
 * Capitalize first letter
 */
function capitalizeFirst(str: string): string {
  return str.charAt(0).toUpperCase() + str.slice(1);
}

