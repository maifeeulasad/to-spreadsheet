/**
 * @fileoverview Excel package relationships XML generation
 * Handles the generation of .rels file which defines top-level relationships in the Excel package
 * This file maps the main relationships between the package and its primary components
 * 
 * @author Maifee Ul Asad <maifeeulasad@gmail.com>
 * @license MIT
 */

/**
 * Generates the main .rels file for an Excel package
 * Creates the top-level relationships that define how Excel locates the main workbook and properties
 * This file is required by OpenXML specification and must be in the _rels folder
 * @returns {string} Complete XML content for .rels file
 * @internal
 */
const generateRels = () =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml" />
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml" />
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" />
</Relationships>
`;

/**
 * Exports the main package relationships XML generation function
 * @name generateRels
 * @function
 * @description Generates .rels file defining top-level relationships in the Excel package
 * @see {@link generateRels} - Main package relationships generation function
 */
export { generateRels };
