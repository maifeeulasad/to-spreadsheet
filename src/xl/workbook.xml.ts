/**
 * @fileoverview Excel workbook.xml generation
 * Handles the generation of the main workbook.xml file which defines workbook structure and sheet references
 * This is the central file that ties together all worksheets and defines the workbook properties
 * 
 * @author Maifee Ul Asad <maifeeulasad@gmail.com>
 * @license MIT
 */

import { IWorkbook } from "..";
import { indexToVbIndex, indexToVbRelationIndex } from '../util'

/**
 * Generates the main workbook.xml file for an Excel workbook
 * Creates the workbook structure definition including sheet references, properties, and metadata
 * Contains workbook-level settings and references to all worksheets
 * @param {IWorkbook} workbook - The workbook data containing sheets and metadata
 * @returns {string} Complete XML content for workbook.xml file
 * @internal
 */
const generateWorkBookXml = (workbook: IWorkbook) => 
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">
    <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="25427" />
		<workbookProtection/>
		<sheets>
			${workbook.sheets.map((sheet, index) => `
			<sheet name="${sheet.title}" sheetId="${indexToVbIndex(index)}" r:id="rId${indexToVbRelationIndex(index)}"/>
			`).join('\n')}
		</sheets>
    <workbookPr defaultThemeVersion="166925" />
    <AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
        <Choice Requires="x15">
            <absPath url="C:\\projects\\to-spreadsheet\\files\\" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac" />
        </Choice>
    </AlternateContent>
    <revisionPtr revIDLastSave="0" documentId="13_ncr:1_{0266953E-4E47-4F87-8C93-A72373D927CC}" xr6:coauthVersionLast="47" xr6:coauthVersionMax="47" xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}" />
    <bookViews>
        <workbookView xWindow="825" yWindow="-120" windowWidth="28095" windowHeight="16440" xr2:uid="{BABDBE04-4BCD-4C42-B58D-30BEF5B5054D}" />
    </bookViews>
    <calcPr calcId="181029" />
    <extLst>
        <ext uri="{140A7094-0E35-4892-8432-C4D2E57EDEB5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">
            <workbookPr chartTrackingRefBase="1" />
        </ext>
        <ext uri="{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}" xmlns:xcalcf="http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures">
            <calcFeatures>
                <feature name="microsoft.com:RD" />
                <feature name="microsoft.com:FV" />
                <feature name="microsoft.com:LET_WF" />
                <feature name="microsoft.com:LAMBDA_WF" />
            </calcFeatures>
        </ext>
    </extLst>
</workbook>
`;

/**
 * Exports the main workbook XML generation function
 * @name generateWorkBookXml
 * @function
 * @description Generates complete workbook.xml content defining workbook structure and sheet references
 * @see {@link generateWorkBookXml} - Main workbook generation function
 */
export { generateWorkBookXml };
