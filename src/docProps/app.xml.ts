/**
 * @fileoverview Excel application properties XML generation
 * Handles the generation of app.xml which contains application-specific properties and metadata
 * This file includes information about worksheets, document security, and application version
 * 
 * @author Maifee Ul Asad <maifeeulasad@gmail.com>
 * @license MIT
 */

import { IWorkbook } from '../index'

/**
 * Generates the app.xml file for an Excel workbook
 * Creates application properties including worksheet information, security settings, and version data
 * This file is part of the OpenXML document properties and provides metadata about the application
 * @param {IWorkbook} workbook - The workbook data containing sheets and application metadata
 * @returns {string} Complete XML content for app.xml file
 * @internal
 */
const generateAppXml = (workbook: IWorkbook) =>
    `
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
		        xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
		<Template></Template>
		<TotalTime>0</TotalTime>
		<Application>
			Microsoft Excel
		</Application>
		<DocSecurity>
			0
		</DocSecurity>
		<ScaleCrop>
			false
		</ScaleCrop>
		<HeadingPairs>
			<vt:vector size="${workbook.sheets.length}" baseType="variant">
				<vt:variant>
					<vt:lpstr>Worksheets</vt:lpstr>
				</vt:variant>
				<vt:variant>
					<vt:i4>${workbook.sheets.length}</vt:i4>
				</vt:variant>
			</vt:vector>
		</HeadingPairs>
		<TitlesOfParts>
			<vt:vector size="${workbook.sheets.length}" baseType="lpstr">
				${workbook.sheets.map(sheet => `
				<vt:lpstr>${sheet.title}</vt:lpstr>
				`)}
			</vt:vector>
		</TitlesOfParts>
		<LinksUpToDate>
			false
		</LinksUpToDate>
		<SharedDoc>
			false
		</SharedDoc>
		<HyperlinksChanged>
			false
		</HyperlinksChanged>
		<AppVersion>
			16.0300
		</AppVersion>
	</Properties>
`;

/**
 * Exports the application properties XML generation function
 * @name generateAppXml
 * @function
 * @description Generates app.xml containing application-specific properties and worksheet metadata
 * @see {@link generateAppXml} - Main application properties generation function
 */
export { generateAppXml };
