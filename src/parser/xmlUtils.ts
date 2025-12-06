// XML utility functions for parsing PPTX content
import { XMLParser } from 'fast-xml-parser';

// XML namespaces used in PPTX files
export const NAMESPACES = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
};

// Create XML parser instance with options
const xmlParser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  textNodeName: '#text',
  preserveOrder: false,
  removeNSPrefix: true,
  parseAttributeValue: false,
  trimValues: false,
  isArray: (tagName) => {
    // These elements commonly appear multiple times
    const arrayTags = ['p', 'r', 'sp', 'pic', 'grpSp', 'cxnSp', 'graphicFrame', 
                       'sldId', 'Relationship', 'tr', 'tc', 'gridCol', 'gs'];
    return arrayTags.includes(tagName);
  },
});

/**
 * Parsed XML element type
 */
export interface ParsedElement {
  [key: string]: any;
  '@_attributes'?: Record<string, string>;
  '#text'?: string | number;
}

/**
 * Parse XML string to object
 */
export function parseXML(xmlString: string): ParsedElement {
  try {
    const result = xmlParser.parse(xmlString);
    return result;
  } catch (error) {
    throw new Error(`XML parsing error: ${error}`);
  }
}

/**
 * Get the root element name from a parsed XML object
 */
export function getRootElement(doc: ParsedElement): { name: string; element: ParsedElement } | null {
  const keys = Object.keys(doc).filter(k => !k.startsWith('@_') && !k.startsWith('#') && !k.startsWith('?'));
  if (keys.length === 0) return null;
  const name = keys[0];
  return { name, element: doc[name] };
}

/**
 * Normalize element to always be an array
 */
function toArray<T>(value: T | T[] | undefined): T[] {
  if (value === undefined || value === null) return [];
  return Array.isArray(value) ? value : [value];
}

/**
 * Get elements by local name (ignoring namespace)
 * Works with both prefixed and unprefixed names
 */
export function getElementsByLocalName(parent: ParsedElement, localName: string): ParsedElement[] {
  const results: ParsedElement[] = [];
  
  function search(obj: any): void {
    if (!obj || typeof obj !== 'object') return;
    
    if (Array.isArray(obj)) {
      for (const item of obj) {
        search(item);
      }
      return;
    }
    
    for (const key of Object.keys(obj)) {
      if (key.startsWith('@_') || key.startsWith('#') || key.startsWith('?')) continue;
      
      // Check if this key matches (with or without namespace prefix)
      const keyLocalName = key.includes(':') ? key.split(':')[1] : key;
      if (keyLocalName === localName) {
        const value = obj[key];
        if (Array.isArray(value)) {
          results.push(...value);
        } else if (value !== null && value !== undefined) {
          results.push(value);
        }
      }
      
      // Recurse into children
      if (typeof obj[key] === 'object') {
        search(obj[key]);
      }
    }
  }
  
  search(parent);
  return results;
}

/**
 * Get first element by local name
 */
export function getElementByLocalName(parent: ParsedElement, localName: string): ParsedElement | null {
  const elements = getElementsByLocalName(parent, localName);
  return elements.length > 0 ? elements[0] : null;
}

/**
 * Get direct children by local name (not recursive)
 */
export function getChildrenByLocalName(parent: ParsedElement, localName: string): ParsedElement[] {
  if (!parent || typeof parent !== 'object') return [];
  
  const results: ParsedElement[] = [];
  
  for (const key of Object.keys(parent)) {
    if (key.startsWith('@_') || key.startsWith('#') || key.startsWith('?')) continue;
    
    const keyLocalName = key.includes(':') ? key.split(':')[1] : key;
    if (keyLocalName === localName) {
      const value = parent[key];
      if (Array.isArray(value)) {
        results.push(...value);
      } else if (value !== null && value !== undefined) {
        results.push(value);
      }
    }
  }
  
  return results;
}

/**
 * Get first direct child by local name
 */
export function getChildByLocalName(parent: ParsedElement, localName: string): ParsedElement | null {
  const children = getChildrenByLocalName(parent, localName);
  return children.length > 0 ? children[0] : null;
}

/**
 * Get all direct children elements (excluding attributes and text)
 */
export function getChildren(parent: ParsedElement): { name: string; element: ParsedElement }[] {
  if (!parent || typeof parent !== 'object') return [];
  
  const results: { name: string; element: ParsedElement }[] = [];
  
  for (const key of Object.keys(parent)) {
    if (key.startsWith('@_') || key.startsWith('#') || key.startsWith('?')) continue;
    
    const value = parent[key];
    if (Array.isArray(value)) {
      for (const item of value) {
        if (item && typeof item === 'object') {
          results.push({ name: key, element: item });
        }
      }
    } else if (value !== null && value !== undefined && typeof value === 'object') {
      results.push({ name: key, element: value });
    }
  }
  
  return results;
}

/**
 * Get attribute value
 */
export function getAttribute(element: ParsedElement, name: string, nsPrefix?: string): string | null {
  if (!element || typeof element !== 'object') return null;
  
  // Try with namespace prefix first
  if (nsPrefix) {
    const prefixedKey = `@_${nsPrefix}:${name}`;
    if (element[prefixedKey] !== undefined) {
      return String(element[prefixedKey]);
    }
  }
  
  // Try without prefix
  const key = `@_${name}`;
  if (element[key] !== undefined) {
    return String(element[key]);
  }
  
  return null;
}

/**
 * Get attribute as number
 */
export function getAttributeAsNumber(element: ParsedElement, name: string, defaultValue: number = 0): number {
  const value = getAttribute(element, name);
  if (value === null) return defaultValue;
  
  const parsed = parseInt(value, 10);
  return isNaN(parsed) ? defaultValue : parsed;
}

/**
 * Get text content from element - IMPROVED VERSION
 * Handles multiple formats that fast-xml-parser might return
 */
export function getTextContent(element: ParsedElement): string {
  if (!element) return '';
  
  // Direct string value
  if (typeof element === 'string') return element;
  
  // Direct number value
  if (typeof element === 'number') return String(element);
  
  // Check for #text property (standard fast-xml-parser text node)
  if (element['#text'] !== undefined) {
    return String(element['#text']);
  }
  
  // Check if element itself is an empty object (might be self-closing tag)
  if (typeof element === 'object') {
    const keys = Object.keys(element);
    
    // If no keys, return empty
    if (keys.length === 0) return '';
    
    // Check if there's a direct text value stored under any non-attribute key
    for (const key of keys) {
      if (key.startsWith('@_') || key.startsWith('?')) continue;
      const value = element[key];
      if (typeof value === 'string') {
        return value;
      }
      if (typeof value === 'number') {
        return String(value);
      }
    }
  }
  
  return '';
}

/**
 * Theme color aliases mapping
 * Maps PowerPoint scheme color names to theme color keys
 */
const schemeColorAliases: Record<string, string> = {
  'tx1': 'dk1',     // Text 1 -> Dark 1
  'tx2': 'dk2',     // Text 2 -> Dark 2
  'bg1': 'lt1',     // Background 1 -> Light 1
  'bg2': 'lt2',     // Background 2 -> Light 2
  'phClr': 'accent1', // Placeholder color -> Accent 1
};

/**
 * Apply color modifiers (tint, shade, lumMod, lumOff)
 */
function applyColorModifiers(baseColor: string, colorElement: ParsedElement): string {
  if (!baseColor || !colorElement) return baseColor;
  
  let color = baseColor;
  
  // Get modifiers
  const tint = getChildByLocalName(colorElement, 'tint');
  const shade = getChildByLocalName(colorElement, 'shade');
  const lumMod = getChildByLocalName(colorElement, 'lumMod');
  const lumOff = getChildByLocalName(colorElement, 'lumOff');
  const alpha = getChildByLocalName(colorElement, 'alpha');
  
  // Parse base color to RGB
  let r = parseInt(color.slice(1, 3), 16);
  let g = parseInt(color.slice(3, 5), 16);
  let b = parseInt(color.slice(5, 7), 16);
  
  // Apply tint (lighten toward white)
  if (tint) {
    const val = getAttributeAsNumber(tint, 'val', 100000) / 100000;
    r = Math.round(r + (255 - r) * (1 - val));
    g = Math.round(g + (255 - g) * (1 - val));
    b = Math.round(b + (255 - b) * (1 - val));
  }
  
  // Apply shade (darken toward black)
  if (shade) {
    const val = getAttributeAsNumber(shade, 'val', 100000) / 100000;
    r = Math.round(r * val);
    g = Math.round(g * val);
    b = Math.round(b * val);
  }
  
  // Apply luminance modification
  if (lumMod || lumOff) {
    // Convert to HSL
    const hsl = rgbToHsl(r, g, b);
    
    if (lumMod) {
      const val = getAttributeAsNumber(lumMod, 'val', 100000) / 100000;
      hsl.l = hsl.l * val;
    }
    
    if (lumOff) {
      const val = getAttributeAsNumber(lumOff, 'val', 0) / 100000;
      hsl.l = Math.min(1, Math.max(0, hsl.l + val));
    }
    
    // Convert back to RGB
    const rgb = hslToRgb(hsl.h, hsl.s, hsl.l);
    r = rgb.r;
    g = rgb.g;
    b = rgb.b;
  }
  
  // Clamp values
  r = Math.min(255, Math.max(0, r));
  g = Math.min(255, Math.max(0, g));
  b = Math.min(255, Math.max(0, b));
  
  return `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`.toUpperCase();
}

/**
 * Convert RGB to HSL
 */
function rgbToHsl(r: number, g: number, b: number): { h: number; s: number; l: number } {
  r /= 255;
  g /= 255;
  b /= 255;
  
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  let s = 0;
  const l = (max + min) / 2;
  
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    
    switch (max) {
      case r:
        h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
        break;
      case g:
        h = ((b - r) / d + 2) / 6;
        break;
      case b:
        h = ((r - g) / d + 4) / 6;
        break;
    }
  }
  
  return { h, s, l };
}

/**
 * Convert HSL to RGB
 */
function hslToRgb(h: number, s: number, l: number): { r: number; g: number; b: number } {
  let r: number, g: number, b: number;
  
  if (s === 0) {
    r = g = b = l;
  } else {
    const hue2rgb = (p: number, q: number, t: number): number => {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      if (t < 1/6) return p + (q - p) * 6 * t;
      if (t < 1/2) return q;
      if (t < 2/3) return p + (q - p) * (2/3 - t) * 6;
      return p;
    };
    
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = hue2rgb(p, q, h + 1/3);
    g = hue2rgb(p, q, h);
    b = hue2rgb(p, q, h - 1/3);
  }
  
  return {
    r: Math.round(r * 255),
    g: Math.round(g * 255),
    b: Math.round(b * 255),
  };
}

/**
 * Convert PowerPoint color formats to hex - IMPROVED VERSION
 */
export function parseColor(colorElement: ParsedElement | null, themeColors?: Map<string, string>): string | undefined {
  if (!colorElement) return undefined;
  
  // Check for srgbClr (RGB color) - most specific, check first
  const srgbClr = getChildByLocalName(colorElement, 'srgbClr');
  if (srgbClr) {
    const val = getAttribute(srgbClr, 'val');
    if (val) {
      const baseColor = `#${val}`;
      return applyColorModifiers(baseColor, srgbClr);
    }
  }
  
  // Check for schemeClr (theme color)
  const schemeClr = getChildByLocalName(colorElement, 'schemeClr');
  if (schemeClr) {
    const val = getAttribute(schemeClr, 'val');
    if (val) {
      // Apply alias mapping
      const mappedVal = schemeColorAliases[val] || val;
      
      if (themeColors && themeColors.has(mappedVal)) {
        const baseColor = themeColors.get(mappedVal)!;
        return applyColorModifiers(baseColor, schemeClr);
      }
      
      // Also try the original value
      if (themeColors && themeColors.has(val)) {
        const baseColor = themeColors.get(val)!;
        return applyColorModifiers(baseColor, schemeClr);
      }
      
      // Fallback for common scheme colors without theme
      const defaultSchemeColors: Record<string, string> = {
        'dk1': '#000000',
        'lt1': '#FFFFFF',
        'dk2': '#1F497D',
        'lt2': '#EEECE1',
        'accent1': '#4F81BD',
        'accent2': '#C0504D',
        'accent3': '#9BBB59',
        'accent4': '#8064A2',
        'accent5': '#4BACC6',
        'accent6': '#F79646',
        'hlink': '#0000FF',
        'folHlink': '#800080',
        'tx1': '#000000',
        'tx2': '#1F497D',
        'bg1': '#FFFFFF',
        'bg2': '#EEECE1',
      };
      
      if (defaultSchemeColors[mappedVal]) {
        return applyColorModifiers(defaultSchemeColors[mappedVal], schemeClr);
      }
      if (defaultSchemeColors[val]) {
        return applyColorModifiers(defaultSchemeColors[val], schemeClr);
      }
    }
  }
  
  // Check for prstClr (preset color)
  const prstClr = getChildByLocalName(colorElement, 'prstClr');
  if (prstClr) {
    const val = getAttribute(prstClr, 'val');
    if (val) {
      const baseColor = presetColorToHex(val);
      return applyColorModifiers(baseColor, prstClr);
    }
  }
  
  // Check for sysClr (system color)
  const sysClr = getChildByLocalName(colorElement, 'sysClr');
  if (sysClr) {
    const lastClr = getAttribute(sysClr, 'lastClr');
    if (lastClr) {
      const baseColor = `#${lastClr}`;
      return applyColorModifiers(baseColor, sysClr);
    }
    // Try val attribute
    const val = getAttribute(sysClr, 'val');
    if (val) {
      // System color names
      const systemColors: Record<string, string> = {
        'windowText': '#000000',
        'window': '#FFFFFF',
        'highlightText': '#FFFFFF',
        'highlight': '#0078D7',
        'grayText': '#6D6D6D',
        'btnText': '#000000',
        'btnFace': '#F0F0F0',
        'btnHighlight': '#FFFFFF',
        'btnShadow': '#A0A0A0',
        '3dLight': '#E3E3E3',
        '3dDkShadow': '#696969',
      };
      if (systemColors[val]) {
        return applyColorModifiers(systemColors[val], sysClr);
      }
    }
  }
  
  return undefined;
}

/**
 * Convert preset color names to hex values
 */
function presetColorToHex(colorName: string): string {
  const presetColors: Record<string, string> = {
    aliceBlue: '#F0F8FF',
    antiqueWhite: '#FAEBD7',
    aqua: '#00FFFF',
    aquamarine: '#7FFFD4',
    azure: '#F0FFFF',
    beige: '#F5F5DC',
    bisque: '#FFE4C4',
    black: '#000000',
    blanchedAlmond: '#FFEBCD',
    blue: '#0000FF',
    blueViolet: '#8A2BE2',
    brown: '#A52A2A',
    burlyWood: '#DEB887',
    cadetBlue: '#5F9EA0',
    chartreuse: '#7FFF00',
    chocolate: '#D2691E',
    coral: '#FF7F50',
    cornflowerBlue: '#6495ED',
    cornsilk: '#FFF8DC',
    crimson: '#DC143C',
    cyan: '#00FFFF',
    darkBlue: '#00008B',
    darkCyan: '#008B8B',
    darkGoldenrod: '#B8860B',
    darkGray: '#A9A9A9',
    darkGreen: '#006400',
    darkKhaki: '#BDB76B',
    darkMagenta: '#8B008B',
    darkOliveGreen: '#556B2F',
    darkOrange: '#FF8C00',
    darkOrchid: '#9932CC',
    darkRed: '#8B0000',
    darkSalmon: '#E9967A',
    darkSeaGreen: '#8FBC8F',
    darkSlateBlue: '#483D8B',
    darkSlateGray: '#2F4F4F',
    darkTurquoise: '#00CED1',
    darkViolet: '#9400D3',
    deepPink: '#FF1493',
    deepSkyBlue: '#00BFFF',
    dimGray: '#696969',
    dkBlue: '#00008B',
    dkCyan: '#008B8B',
    dkGoldenrod: '#B8860B',
    dkGray: '#A9A9A9',
    dkGreen: '#006400',
    dkKhaki: '#BDB76B',
    dkMagenta: '#8B008B',
    dkOliveGreen: '#556B2F',
    dkOrange: '#FF8C00',
    dkOrchid: '#9932CC',
    dkRed: '#8B0000',
    dkSalmon: '#E9967A',
    dkSeaGreen: '#8FBC8F',
    dkSlateBlue: '#483D8B',
    dkSlateGray: '#2F4F4F',
    dkTurquoise: '#00CED1',
    dkViolet: '#9400D3',
    dodgerBlue: '#1E90FF',
    firebrick: '#B22222',
    floralWhite: '#FFFAF0',
    forestGreen: '#228B22',
    fuchsia: '#FF00FF',
    gainsboro: '#DCDCDC',
    ghostWhite: '#F8F8FF',
    gold: '#FFD700',
    goldenrod: '#DAA520',
    gray: '#808080',
    green: '#008000',
    greenYellow: '#ADFF2F',
    honeydew: '#F0FFF0',
    hotPink: '#FF69B4',
    indianRed: '#CD5C5C',
    indigo: '#4B0082',
    ivory: '#FFFFF0',
    khaki: '#F0E68C',
    lavender: '#E6E6FA',
    lavenderBlush: '#FFF0F5',
    lawnGreen: '#7CFC00',
    lemonChiffon: '#FFFACD',
    lightBlue: '#ADD8E6',
    lightCoral: '#F08080',
    lightCyan: '#E0FFFF',
    lightGoldenrodYellow: '#FAFAD2',
    lightGray: '#D3D3D3',
    lightGreen: '#90EE90',
    lightPink: '#FFB6C1',
    lightSalmon: '#FFA07A',
    lightSeaGreen: '#20B2AA',
    lightSkyBlue: '#87CEFA',
    lightSlateGray: '#778899',
    lightSteelBlue: '#B0C4DE',
    lightYellow: '#FFFFE0',
    lime: '#00FF00',
    limeGreen: '#32CD32',
    linen: '#FAF0E6',
    ltBlue: '#ADD8E6',
    ltCoral: '#F08080',
    ltCyan: '#E0FFFF',
    ltGoldenrodYellow: '#FAFAD2',
    ltGray: '#D3D3D3',
    ltGreen: '#90EE90',
    ltPink: '#FFB6C1',
    ltSalmon: '#FFA07A',
    ltSeaGreen: '#20B2AA',
    ltSkyBlue: '#87CEFA',
    ltSlateGray: '#778899',
    ltSteelBlue: '#B0C4DE',
    ltYellow: '#FFFFE0',
    magenta: '#FF00FF',
    maroon: '#800000',
    medAquamarine: '#66CDAA',
    medBlue: '#0000CD',
    mediumAquamarine: '#66CDAA',
    mediumBlue: '#0000CD',
    mediumOrchid: '#BA55D3',
    mediumPurple: '#9370DB',
    mediumSeaGreen: '#3CB371',
    mediumSlateBlue: '#7B68EE',
    mediumSpringGreen: '#00FA9A',
    mediumTurquoise: '#48D1CC',
    mediumVioletRed: '#C71585',
    medOrchid: '#BA55D3',
    medPurple: '#9370DB',
    medSeaGreen: '#3CB371',
    medSlateBlue: '#7B68EE',
    medSpringGreen: '#00FA9A',
    medTurquoise: '#48D1CC',
    medVioletRed: '#C71585',
    midnightBlue: '#191970',
    mintCream: '#F5FFFA',
    mistyRose: '#FFE4E1',
    moccasin: '#FFE4B5',
    navajoWhite: '#FFDEAD',
    navy: '#000080',
    oldLace: '#FDF5E6',
    olive: '#808000',
    oliveDrab: '#6B8E23',
    orange: '#FFA500',
    orangeRed: '#FF4500',
    orchid: '#DA70D6',
    paleGoldenrod: '#EEE8AA',
    paleGreen: '#98FB98',
    paleTurquoise: '#AFEEEE',
    paleVioletRed: '#DB7093',
    papayaWhip: '#FFEFD5',
    peachPuff: '#FFDAB9',
    peru: '#CD853F',
    pink: '#FFC0CB',
    plum: '#DDA0DD',
    powderBlue: '#B0E0E6',
    purple: '#800080',
    red: '#FF0000',
    rosyBrown: '#BC8F8F',
    royalBlue: '#4169E1',
    saddleBrown: '#8B4513',
    salmon: '#FA8072',
    sandyBrown: '#F4A460',
    seaGreen: '#2E8B57',
    seaShell: '#FFF5EE',
    sienna: '#A0522D',
    silver: '#C0C0C0',
    skyBlue: '#87CEEB',
    slateBlue: '#6A5ACD',
    slateGray: '#708090',
    snow: '#FFFAFA',
    springGreen: '#00FF7F',
    steelBlue: '#4682B4',
    tan: '#D2B48C',
    teal: '#008080',
    thistle: '#D8BFD8',
    tomato: '#FF6347',
    turquoise: '#40E0D0',
    violet: '#EE82EE',
    wheat: '#F5DEB3',
    white: '#FFFFFF',
    whiteSmoke: '#F5F5F5',
    yellow: '#FFFF00',
    yellowGreen: '#9ACD32',
  };
  
  return presetColors[colorName] || '#000000';
}

/**
 * Parse line/border properties
 */
export function parseLineProperties(lnElement: ParsedElement | null, themeColors?: Map<string, string>): { color: string; width: number } | undefined {
  if (!lnElement) return undefined;
  
  const width = getAttributeAsNumber(lnElement, 'w', 12700); // Default 1pt in EMU
  const widthPx = width / 12700; // Convert from EMU to points/pixels
  
  // Get color from solidFill
  const solidFill = getChildByLocalName(lnElement, 'solidFill');
  const color = parseColor(solidFill, themeColors) || '#000000';
  
  return { color, width: widthPx };
}

/**
 * Parse fill properties
 */
export function parseFillProperties(element: ParsedElement, themeColors?: Map<string, string>): { type: 'solid' | 'gradient' | 'none'; color?: string; opacity?: number } | undefined {
  // Check for noFill
  if (getChildByLocalName(element, 'noFill')) {
    return { type: 'none' };
  }
  
  // Check for solidFill
  const solidFill = getChildByLocalName(element, 'solidFill');
  if (solidFill) {
    const color = parseColor(solidFill, themeColors);
    return { type: 'solid', color, opacity: 1 };
  }
  
  // Check for gradFill
  const gradFill = getChildByLocalName(element, 'gradFill');
  if (gradFill) {
    // For now, return first color in gradient
    const gsLst = getChildByLocalName(gradFill, 'gsLst');
    if (gsLst) {
      const gsElements = getChildrenByLocalName(gsLst, 'gs');
      if (gsElements.length > 0) {
        const color = parseColor(gsElements[0], themeColors);
        return { type: 'gradient', color, opacity: 1 };
      }
    }
  }
  
  return undefined;
}
