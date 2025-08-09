/**
 * @fileoverview Excel styles.xml generation
 * Handles the generation of Excel styling XML including borders, fonts, fills, and cell formats
 * This file creates the styles.xml file that defines all visual styling for the workbook
 * 
 * @author Maifee Ul Asad <maifeeulasad@gmail.com>
 * @license MIT
 */

import { IBorder, BorderStyle, ICellStyle } from "..";

/**
 * Generates XML representation of border styling for a single border configuration
 * @param {IBorder} border - Border configuration object
 * @returns {string} XML string representing the border styling
 * @internal
 */
const generateBorderXml = (border: IBorder): string => {
  /**
   * Generates color XML attribute for borders
   * @param {string} color - Optional hex color string
   * @returns {string} Color XML or empty string
   */
  const getColorXml = (color?: string) => 
    color ? `<color rgb="${color.replace('#', 'FF')}" />` : '';
  
  /**
   * Generates XML for individual border sides (unused helper function)
   * @deprecated - Replaced by individual side XML generation
   */
  const getBorderSideXml = (side: BorderStyle | undefined, color?: string) => {
    if (!side || side === BorderStyle.none) {
      return '<left />';
    }
    return `<left style="${side}">${getColorXml(color)}</left>`;
  };

  // Generate XML for each border side
  const leftXml = !border.left || border.left === BorderStyle.none 
    ? '<left />' 
    : `<left style="${border.left}">${getColorXml(border.color)}</left>`;
    
  const rightXml = !border.right || border.right === BorderStyle.none 
    ? '<right />' 
    : `<right style="${border.right}">${getColorXml(border.color)}</right>`;
    
  const topXml = !border.top || border.top === BorderStyle.none 
    ? '<top />' 
    : `<top style="${border.top}">${getColorXml(border.color)}</top>`;
    
  const bottomXml = !border.bottom || border.bottom === BorderStyle.none 
    ? '<bottom />' 
    : `<bottom style="${border.bottom}">${getColorXml(border.color)}</bottom>`;

  return `
    <border>
      ${leftXml}
      ${rightXml}
      ${topXml}
      ${bottomXml}
      <diagonal />
    </border>`;
};

/**
 * Generates XML representation of font styling for foreground color
 * @param {string} color - Optional hex color string for font color
 * @returns {string} XML string representing the font styling
 * @internal
 */
const generateFontXml = (color?: string): string => {
  const colorXml = color 
    ? `<color rgb="${color.replace('#', 'FF')}" />`
    : '<color theme="1" />';
    
  return `
    <font>
      <sz val="11" />
      ${colorXml}
      <name val="Calibri" />
      <family val="2" />
      <scheme val="minor" />
    </font>`;
};

/**
 * Generates XML representation of fill styling for background color
 * @param {string} color - Optional hex color string for background color
 * @returns {string} XML string representing the fill styling
 * @internal
 */
const generateFillXml = (color?: string): string => {
  if (!color) {
    return `
    <fill>
      <patternFill patternType="none" />
    </fill>`;
  }
  
  return `
    <fill>
      <patternFill patternType="solid">
        <fgColor rgb="${color.replace('#', 'FF')}" />
        <bgColor indexed="64" />
      </patternFill>
    </fill>`;
};

/**
 * Generates the complete styles.xml content for an Excel workbook
 * Creates a comprehensive styling definition including fonts, fills, borders, and cell formats
 * @param {Map<string, ICellStyle>} styleMap - Map of unique complete styles used in the workbook
 * @param {boolean} hasDateCells - Whether the workbook contains date cells requiring date formatting
 * @returns {string} Complete XML content for styles.xml file
 * @internal
 */
const generateStyleXml = (styleMap: Map<string, ICellStyle>, hasDateCells: boolean = false) => {
  const styles = Array.from(styleMap.values());
  const styleCount = styles.length;
  
  // Extract unique borders, fonts, and fills from all styles
  const uniqueBorders = new Map<string, IBorder>();
  const uniqueFonts = new Map<string, string>();
  const uniqueFills = new Map<string, string>();
  
  // Add default styles
  uniqueBorders.set("none", {});
  uniqueFonts.set("default", "");
  uniqueFills.set("none", "");
  uniqueFills.set("gray125", ""); // Excel default
  
  // Collect unique styles
  styles.forEach(style => {
    if (style.border) {
      uniqueBorders.set(JSON.stringify(style.border), style.border);
    }
    if (style.foregroundColor) {
      uniqueFonts.set(style.foregroundColor, style.foregroundColor);
    }
    if (style.backgroundColor) {
      uniqueFills.set(style.backgroundColor, style.backgroundColor);
    }
  });
  
  // Generate XML sections
  const borderArray = Array.from(uniqueBorders.values());
  const fontArray = Array.from(uniqueFonts.values());
  const fillArray = Array.from(uniqueFills.values());
  
  const bordersXml = borderArray.map(border => generateBorderXml(border)).join('');
  const fontsXml = fontArray.map(color => generateFontXml(color || undefined)).join('');
  const fillsXml = fillArray.map(color => generateFillXml(color || undefined)).join('');
  
  // Generate number formats if date cells are present
  const numFmtsXml = hasDateCells 
    ? `<numFmts count="1">
        <numFmt numFmtId="164" formatCode="mm/dd/yyyy" />
    </numFmts>`
    : '';
  
  // Generate cell format definitions
  let cellXfsXml = '';
  let cellXfsCount = styleCount;
  
  if (hasDateCells) {
    // Generate formats for regular cells
    cellXfsXml += styles.map((style, index) => {
      const borderIndex = Array.from(uniqueBorders.keys()).indexOf(JSON.stringify(style.border || {}));
      const fontIndex = Array.from(uniqueFonts.keys()).indexOf(style.foregroundColor || "default");
      const fillIndex = Array.from(uniqueFills.keys()).indexOf(style.backgroundColor || "none");
      return `<xf numFmtId="0" fontId="${fontIndex}" fillId="${fillIndex}" borderId="${borderIndex}" xfId="0" />`;
    }).join('\n        ');
    
    // Generate formats for date cells
    cellXfsXml += '\n        ';
    cellXfsXml += styles.map((style, index) => {
      const borderIndex = Array.from(uniqueBorders.keys()).indexOf(JSON.stringify(style.border || {}));
      const fontIndex = Array.from(uniqueFonts.keys()).indexOf(style.foregroundColor || "default");
      const fillIndex = Array.from(uniqueFills.keys()).indexOf(style.backgroundColor || "none");
      return `<xf numFmtId="164" fontId="${fontIndex}" fillId="${fillIndex}" borderId="${borderIndex}" xfId="0" />`;
    }).join('\n        ');
    
    cellXfsCount = styleCount * 2;
  } else {
    cellXfsXml = styles.map((style, index) => {
      const borderIndex = Array.from(uniqueBorders.keys()).indexOf(JSON.stringify(style.border || {}));
      const fontIndex = Array.from(uniqueFonts.keys()).indexOf(style.foregroundColor || "default");
      const fillIndex = Array.from(uniqueFills.keys()).indexOf(style.backgroundColor || "none");
      return `<xf numFmtId="0" fontId="${fontIndex}" fillId="${fillIndex}" borderId="${borderIndex}" xfId="0" />`;
    }).join('\n        ');
  }

  /**
   * Complete XML template for Excel styles.xml file
   * Includes standard fonts, fills, borders, and cell format definitions
   * Follows OpenXML specification for Excel styling
   */
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
    ${numFmtsXml}
    <fonts count="${fontArray.length}" x14ac:knownFonts="1">
        ${fontsXml}
    </fonts>
    <fills count="${fillArray.length}">
        ${fillsXml}
    </fills>
    <borders count="${borderArray.length}">
        ${bordersXml}
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
    </cellStyleXfs>
    <cellXfs count="${cellXfsCount}">
        ${cellXfsXml}
    </cellXfs>
    <cellStyles count="1">
        <cellStyle name="Normal" xfId="0" builtinId="0" />
    </cellStyles>
    <dxfs count="0" />
    <tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16" />
    <extLst>
        <ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">
            <slicerStyles defaultSlicerStyle="SlicerStyleLight1" />
        </ext>
        <ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
            <timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1" />
        </ext>
    </extLst>
</styleSheet>
`;
};

/**
 * Exports the main styling XML generation function
 * @name generateStyleXml
 * @function
 * @description Generates complete styles.xml content for Excel workbook with border styling support
 * @see {@link generateStyleXml} - Main style generation function
 */
export { generateStyleXml };
