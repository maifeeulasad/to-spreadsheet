/**
 * @fileoverview Excel content types XML generation
 * Handles the generation of [Content_Types].xml which defines MIME types for all parts of the Excel package
 * This file is required by the OpenXML specification to identify content types of various XML parts
 * 
 * @author Maifee Ul Asad <maifeeulasad@gmail.com>
 * @license MIT
 */

import { IWorkbook } from "./index"
import { indexToVbIndex } from './util'

/**
 * Generates the [Content_Types].xml file for an Excel workbook
 * This file defines the content types (MIME types) for all parts of the Excel package
 * Required by OpenXML specification for proper Excel file recognition
 * @param {IWorkbook} workbook - The workbook data containing sheets information
 * @returns {string} Complete XML content for [Content_Types].xml file
 * @internal
 */
const generateContentTypesXml = (workbook: IWorkbook) =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  ${workbook.sheets.map((_, index) => `
  <Override PartName="/xl/worksheets/sheet${indexToVbIndex(index)}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  `).join('\n')}
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`

/**
 * Exports the content types XML generation function
 * @name generateContentTypesXml
 * @function
 * @description Generates [Content_Types].xml defining MIME types for Excel package parts
 * @see {@link generateContentTypesXml} - Main content types generation function
 */
export { generateContentTypesXml };
