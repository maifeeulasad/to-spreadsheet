import { IWorkbook } from "../..";
import { indexToVbIndex, indexToVbRelationIndex } from '../../util'

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

export { generateWorkBookXmlRels };
