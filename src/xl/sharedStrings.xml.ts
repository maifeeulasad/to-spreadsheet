/**
 * @fileoverview Excel shared strings XML generation
 * Handles the generation of sharedStrings.xml which contains all unique text strings used in the workbook
 * Excel uses shared strings to optimize file size by storing unique text values once and referencing them
 * 
 * @author Maifee Ul Asad <maifeeulasad@gmail.com>
 * @license MIT
 */

import { IWorkbook } from ".."

/**
 * Generates the sharedStrings.xml file for an Excel workbook
 * Creates a centralized repository of all unique text strings used across all worksheets
 * Excel references these strings by index to reduce file size and improve performance
 * @param {IWorkbook} workbook - The workbook data containing all strings from cells
 * @returns {string} Complete XML content for sharedStrings.xml file
 * @internal
 */
const generateSharedStrings = (workbook: IWorkbook) => {

	// console.log(workbook.strings)

	/**
	 * Calculate the number of unique strings for the uniqueCount attribute
	 * Excel requires both total count and unique count in the XML
	 * @type {number}
	 */
	const uniqueCount = workbook.strings.reduce((acc:string[], cur) => acc.includes(cur) ? acc : [...acc, cur], []).length;


	/**
	 * Complete XML template for Excel sharedStrings.xml file
	 * Each string is wrapped in <si><t> tags as per OpenXML specification
	 * The count and uniqueCount attributes are required for Excel compatibility
	 */
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${workbook.strings.length}" uniqueCount="${uniqueCount}">
		${workbook.strings.map(string => `
		<si><t>${string}</t></si>
		`).join('\n')}
	</sst>`
}

/**
 * Exports the shared strings XML generation function
 * @name generateSharedStrings
 * @function
 * @description Generates sharedStrings.xml containing all unique text strings from the workbook
 * @see {@link generateSharedStrings} - Main shared strings generation function
 */
export { generateSharedStrings }