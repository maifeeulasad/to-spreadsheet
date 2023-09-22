import { IWorkbook } from '../index'

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

export { generateAppXml };
