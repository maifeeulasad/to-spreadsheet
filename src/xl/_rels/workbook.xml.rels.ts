/**
 * @fileoverview Excel workbook relationships XML generation
 * Handles the generation of workbook.xml.rels which defines relationships between workbook and its parts
 * This file maps the relationships between the main workbook and its associated files (styles, strings, sheets)
 * 
 * @author Maifee Ul Asad <maifeeulasad@gmail.com>
 * @license MIT
 */

import { IWorkbook } from "../..";
import { indexToVbIndex, indexToVbRelationIndex } from '../../util'

/**
 * Generates the workbook.xml.rels file for an Excel workbook
 * Defines the relationships between the main workbook.xml and all its associated parts
 * Each relationship has an ID, type, and target that Excel uses to locate related files
 * @param {IWorkbook} workbook - The workbook data containing sheets information for relationship mapping
 * @returns {string} Complete XML content for workbook.xml.rels file
 * @internal
 */
const generateWorkBookXmlRels = (workbook: IWorkbook) =>
	`<?xml version="1.0" encoding="UTF-8"?>
	<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
		<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
		<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
		${workbook.sheets.map((_, index) => `
		<Relationship Id="rId${indexToVbRelationIndex(index)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${indexToVbIndex(index)}.xml"/>
		`).join('\n')}
	</Relationships>
`;

/**
 * Exports the workbook relationships XML generation function
 * @name generateWorkBookXmlRels
 * @function
 * @description Generates workbook.xml.rels defining relationships between workbook and its associated parts
 * @see {@link generateWorkBookXmlRels} - Main workbook relationships generation function
 */
export { generateWorkBookXmlRels };
