import { PPTXGroupElement, PPTXElement } from '../types/pptx';
import { convertTextElement } from './textConverter';
import { convertShapeElement } from './shapeConverter';
import { convertImageElement } from './imageConverter';
import { convertTableElement } from './tableConverter';
import { convertChartElement } from './chartConverter';

/**
 * Convert PPTX group element to Figma group
 */
export async function convertGroupElement(
  element: PPTXGroupElement,
  parent: FrameNode | GroupNode,
  progressCallback?: (message: string) => void
): Promise<GroupNode | FrameNode> {
  // Create a temporary frame to hold children
  const tempFrame = figma.createFrame();
  tempFrame.name = element.name || 'Group';
  tempFrame.x = element.x;
  tempFrame.y = element.y;
  tempFrame.resize(element.width, element.height);
  tempFrame.fills = [];
  tempFrame.clipsContent = false;

  // Convert all children
  const children: SceneNode[] = [];
  for (const child of element.children) {
    const childNode = await convertElement(child, tempFrame, progressCallback);
    if (childNode) {
      children.push(childNode);
    }
  }

  // If we have children, create a group
  if (children.length > 0) {
    // Move children to parent first
    for (const child of children) {
      parent.appendChild(child);
    }
    
    // Create group from children
    const group = figma.group(children, parent);
    group.name = element.name || 'Group';
    
    // Apply position
    group.x = element.x;
    group.y = element.y;
    
    // Apply rotation if present
    if (element.rotation) {
      group.rotation = -element.rotation;
    }
    
    // Remove the temporary frame
    tempFrame.remove();
    
    return group;
  }

  // No children, return empty frame
  parent.appendChild(tempFrame);
  return tempFrame;
}

/**
 * Convert a single element (used for recursive group conversion)
 */
export async function convertElement(
  element: PPTXElement,
  parent: FrameNode | GroupNode,
  progressCallback?: (message: string) => void
): Promise<SceneNode | null> {
  try {
    switch (element.type) {
      case 'text':
        return await convertTextElement(element, parent);
      case 'shape':
        return await convertShapeElement(element, parent);
      case 'image':
        return await convertImageElement(element, parent);
      case 'table':
        return await convertTableElement(element, parent);
      case 'chart':
        return await convertChartElement(element, parent);
      case 'group':
        return await convertGroupElement(element, parent, progressCallback);
      default:
        console.warn('Unknown element type:', (element as any).type);
        return null;
    }
  } catch (error) {
    console.error('Error converting element:', error);
    return null;
  }
}

/**
 * Convert all elements on a slide
 */
export async function convertSlideElements(
  elements: PPTXElement[],
  parent: FrameNode,
  progressCallback?: (current: number, total: number, message: string) => void
): Promise<void> {
  const total = elements.length;
  
  for (let i = 0; i < elements.length; i++) {
    const element = elements[i];
    
    if (progressCallback) {
      progressCallback(i + 1, total, `Converting element ${i + 1}/${total}`);
    }
    
    await convertElement(element, parent);
  }
}

