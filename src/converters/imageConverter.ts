import { PPTXImageElement } from '../types/pptx';

/**
 * Convert PPTX image element to Figma image node
 */
export async function convertImageElement(
  element: PPTXImageElement,
  parent: FrameNode | GroupNode
): Promise<RectangleNode> {
  // Create a rectangle to hold the image
  const imageNode = figma.createRectangle();
  imageNode.name = element.name || 'Image';
  imageNode.x = element.x;
  imageNode.y = element.y;
  imageNode.resize(element.width, element.height);

  // Apply rotation if present
  if (element.rotation) {
    imageNode.rotation = -element.rotation;
  }

  try {
    // Create image from data
    const imageHash = await createImageFromData(element.imageData, element.mimeType);
    
    if (imageHash) {
      imageNode.fills = [{
        type: 'IMAGE',
        imageHash,
        scaleMode: 'FILL',
      }];
    } else {
      // Fallback to placeholder
      imageNode.fills = [{
        type: 'SOLID',
        color: { r: 0.9, g: 0.9, b: 0.9 },
      }];
    }
  } catch (error) {
    console.error('Error creating image:', error);
    // Fallback to placeholder
    imageNode.fills = [{
      type: 'SOLID',
      color: { r: 0.9, g: 0.9, b: 0.9 },
    }];
  }

  parent.appendChild(imageNode);
  return imageNode;
}

/**
 * Create a Figma image from raw data
 */
async function createImageFromData(
  data: Uint8Array,
  mimeType: string
): Promise<string | null> {
  try {
    // Handle different image formats
    if (mimeType === 'image/emf' || mimeType === 'image/wmf') {
      // EMF/WMF formats are not directly supported by Figma
      // Return null to use placeholder
      console.warn('EMF/WMF images are not supported, using placeholder');
      return null;
    }

    // Create image from bytes
    const image = figma.createImage(data);
    return image.hash;
  } catch (error) {
    console.error('Error creating image from data:', error);
    return null;
  }
}

/**
 * Create a placeholder image node with text
 */
export function createImagePlaceholder(
  element: PPTXImageElement,
  parent: FrameNode | GroupNode,
  message: string = 'Image'
): FrameNode {
  const frame = figma.createFrame();
  frame.name = element.name || 'Image Placeholder';
  frame.x = element.x;
  frame.y = element.y;
  frame.resize(element.width, element.height);
  frame.fills = [{
    type: 'SOLID',
    color: { r: 0.95, g: 0.95, b: 0.95 },
  }];
  frame.strokes = [{
    type: 'SOLID',
    color: { r: 0.8, g: 0.8, b: 0.8 },
  }];
  frame.strokeWeight = 1;

  // Add placeholder icon/text
  const iconFrame = figma.createFrame();
  iconFrame.name = 'Icon';
  iconFrame.resize(48, 48);
  iconFrame.x = (element.width - 48) / 2;
  iconFrame.y = (element.height - 48) / 2 - 16;
  iconFrame.fills = [];

  // Create simple image icon using rectangles
  const iconBg = figma.createRectangle();
  iconBg.resize(48, 40);
  iconBg.y = 4;
  iconBg.fills = [{
    type: 'SOLID',
    color: { r: 0.7, g: 0.7, b: 0.7 },
  }];
  iconBg.cornerRadius = 4;

  const iconCircle = figma.createEllipse();
  iconCircle.resize(12, 12);
  iconCircle.x = 8;
  iconCircle.y = 12;
  iconCircle.fills = [{
    type: 'SOLID',
    color: { r: 1, g: 0.85, b: 0.3 },
  }];

  iconFrame.appendChild(iconBg);
  iconFrame.appendChild(iconCircle);

  frame.appendChild(iconFrame);

  parent.appendChild(frame);
  return frame;
}

