import { generateSheet1Xml, ICellEntry } from "./xl/worksheets/sheet1.xml";

let data = [
  [
    { readOnly: true, value: "" },
    { value: "A", readOnly: true },
    { value: "B", readOnly: true },
    { value: "C", readOnly: true },
    { value: "D", readOnly: true },
    { value: "D", readOnly: true },
  ],
  [
    { readOnly: true, value: 1 },
    { value: 1 },
    { value: 3 },
    { value: 3 },
    { value: 3 },
  ],
  [
    { readOnly: true, value: 2 },
    { value: 2 },
    { value: 4 },
    { value: 4 },
    { value: 4 },
  ],
  [
    { readOnly: true, value: 3 },
    { value: 1 },
    { value: 3 },
    { value: 3 },
    { value: 3 },
  ],
  [
    { readOnly: true, value: 4 },
    { value: 2 },
    { value: 4 },
    { value: 4 },
    { value: 4 },
  ],
];

const sheet = (
  grid: ICellEntry[][]
) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac xr xr2 xr3" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" xmlns:xr3="http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" xr:uid="{6A0CD924-A203-41BB-8331-42E46B7D5F20}">
	<dimension ref="B3:I7" />
	<sheetViews>
		<sheetView tabSelected="1" workbookViewId="0">
			<selection activeCell="B3" sqref="B3:D4" />
		</sheetView>
	</sheetViews>
	<sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25" />
    ${generateSheet1Xml(grid)}
    <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />
</worksheet>
`;

console.log(sheet(data));
//console.log(index(30))

// todo: time with moment js ->  2022-08-11T16:18:46Z
const createDocPropCore = (username: string = "maifeeulasad/to-spreadsheet") =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <creator>
        ${username}
    </creator>
    <lastModifiedBy>
        ${username}
    </lastModifiedBy>
    <created xsi:type="dcterms:W3CDTF">
        ${Date.now()}
    </created>
    <modified xsi:type="dcterms:W3CDTF">
        ${Date.now()}
    </modified>
</coreProperties>
`;

const createDocPropApp = () =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
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
        <vector size="2" baseType="variant">
            <variant>
                <lpstr>
                    Worksheets
                </lpstr>
            </variant>
            <variant>
                <i4>
                    1
                </i4>
            </variant>
        </vector>
    </HeadingPairs>
    <TitlesOfParts>
        <vector size="1" baseType="lpstr">
            <lpstr>
                Sheet1
            </lpstr>
        </vector>
    </TitlesOfParts>
    <Company>
    </Company>
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

const createRel = () =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml" />
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml" />
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" />
</Relationships>
`;

const contentTypeXml = () =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
    <Default Extension="xml" ContentType="application/xml" />
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
    <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml" />
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml" />
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" />
</Types>
`;

const workBook = () =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15 xr xr6 xr10 xr2" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" xmlns:xr6="http://schemas.microsoft.com/office/spreadsheetml/2016/revision6" xmlns:xr10="http://schemas.microsoft.com/office/spreadsheetml/2016/revision10" xmlns:xr2="http://schemas.microsoft.com/office/spreadsheetml/2015/revision2">
    <fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="25427" />
    <workbookPr defaultThemeVersion="166925" />
    <AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
        <Choice Requires="x15">
            <absPath url="D:\\projects\    o-spreadsheet\\files\\" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac" />
        </Choice>
    </AlternateContent>
    <revisionPtr revIDLastSave="0" documentId="13_ncr:1_{0266953E-4E47-4F87-8C93-A72373D927CC}" xr6:coauthVersionLast="47" xr6:coauthVersionMax="47" xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}" />
    <bookViews>
        <workbookView xWindow="825" yWindow="-120" windowWidth="28095" windowHeight="16440" xr2:uid="{BABDBE04-4BCD-4C42-B58D-30BEF5B5054D}" />
    </bookViews>
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1" />
    </sheets>
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

const styles = () =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2 xr" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main" xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision">
    <fonts count="1" x14ac:knownFonts="1">
        <font>
            <sz val="11" />
            <color theme="1" />
            <name val="Calibri" />
            <family val="2" />
            <scheme val="minor" />
        </font>
    </fonts>
    <fills count="2">
        <fill>
            <patternFill patternType="none" />
        </fill>
        <fill>
            <patternFill patternType="gray125" />
        </fill>
    </fills>
    <borders count="1">
        <border>
            <left />
            <right />
            <top />
            <bottom />
            <diagonal />
        </border>
    </borders>
    <cellStyleXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
    </cellStyleXfs>
    <cellXfs count="1">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
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

const theme = () =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
    <themeElements>
        <clrScheme name="Office">
            <dk1>
                <sysClr val="windowText" lastClr="000000" />
            </dk1>
            <lt1>
                <sysClr val="window" lastClr="FFFFFF" />
            </lt1>
            <dk2>
                <srgbClr val="44546A" />
            </dk2>
            <lt2>
                <srgbClr val="E7E6E6" />
            </lt2>
            <accent1>
                <srgbClr val="4472C4" />
            </accent1>
            <accent2>
                <srgbClr val="ED7D31" />
            </accent2>
            <accent3>
                <srgbClr val="A5A5A5" />
            </accent3>
            <accent4>
                <srgbClr val="FFC000" />
            </accent4>
            <accent5>
                <srgbClr val="5B9BD5" />
            </accent5>
            <accent6>
                <srgbClr val="70AD47" />
            </accent6>
            <hlink>
                <srgbClr val="0563C1" />
            </hlink>
            <folHlink>
                <srgbClr val="954F72" />
            </folHlink>
        </clrScheme>
        <fontScheme name="Office">
            <majorFont>
                <latin typeface="Calibri Light" panose="020F0302020204030204" />
                <ea typeface="" />
                <cs typeface="" />
                <font script="Jpan" typeface="游ゴシック Light" />
                <font script="Hang" typeface="맑은 고딕" />
                <font script="Hans" typeface="等线 Light" />
                <font script="Hant" typeface="新細明體" />
                <font script="Arab" typeface="Times New Roman" />
                <font script="Hebr" typeface="Times New Roman" />
                <font script="Thai" typeface="Tahoma" />
                <font script="Ethi" typeface="Nyala" />
                <font script="Beng" typeface="Vrinda" />
                <font script="Gujr" typeface="Shruti" />
                <font script="Khmr" typeface="MoolBoran" />
                <font script="Knda" typeface="Tunga" />
                <font script="Guru" typeface="Raavi" />
                <font script="Cans" typeface="Euphemia" />
                <font script="Cher" typeface="Plantagenet Cherokee" />
                <font script="Yiii" typeface="Microsoft Yi Baiti" />
                <font script="Tibt" typeface="Microsoft Himalaya" />
                <font script="Thaa" typeface="MV Boli" />
                <font script="Deva" typeface="Mangal" />
                <font script="Telu" typeface="Gautami" />
                <font script="Taml" typeface="Latha" />
                <font script="Syrc" typeface="Estrangelo Edessa" />
                <font script="Orya" typeface="Kalinga" />
                <font script="Mlym" typeface="Kartika" />
                <font script="Laoo" typeface="DokChampa" />
                <font script="Sinh" typeface="Iskoola Pota" />
                <font script="Mong" typeface="Mongolian Baiti" />
                <font script="Viet" typeface="Times New Roman" />
                <font script="Uigh" typeface="Microsoft Uighur" />
                <font script="Geor" typeface="Sylfaen" />
                <font script="Armn" typeface="Arial" />
                <font script="Bugi" typeface="Leelawadee UI" />
                <font script="Bopo" typeface="Microsoft JhengHei" />
                <font script="Java" typeface="Javanese Text" />
                <font script="Lisu" typeface="Segoe UI" />
                <font script="Mymr" typeface="Myanmar Text" />
                <font script="Nkoo" typeface="Ebrima" />
                <font script="Olck" typeface="Nirmala UI" />
                <font script="Osma" typeface="Ebrima" />
                <font script="Phag" typeface="Phagspa" />
                <font script="Syrn" typeface="Estrangelo Edessa" />
                <font script="Syrj" typeface="Estrangelo Edessa" />
                <font script="Syre" typeface="Estrangelo Edessa" />
                <font script="Sora" typeface="Nirmala UI" />
                <font script="Tale" typeface="Microsoft Tai Le" />
                <font script="Talu" typeface="Microsoft New Tai Lue" />
                <font script="Tfng" typeface="Ebrima" />
            </majorFont>
            <minorFont>
                <latin typeface="Calibri" panose="020F0502020204030204" />
                <ea typeface="" />
                <cs typeface="" />
                <font script="Jpan" typeface="游ゴシック" />
                <font script="Hang" typeface="맑은 고딕" />
                <font script="Hans" typeface="等线" />
                <font script="Hant" typeface="新細明體" />
                <font script="Arab" typeface="Arial" />
                <font script="Hebr" typeface="Arial" />
                <font script="Thai" typeface="Tahoma" />
                <font script="Ethi" typeface="Nyala" />
                <font script="Beng" typeface="Vrinda" />
                <font script="Gujr" typeface="Shruti" />
                <font script="Khmr" typeface="DaunPenh" />
                <font script="Knda" typeface="Tunga" />
                <font script="Guru" typeface="Raavi" />
                <font script="Cans" typeface="Euphemia" />
                <font script="Cher" typeface="Plantagenet Cherokee" />
                <font script="Yiii" typeface="Microsoft Yi Baiti" />
                <font script="Tibt" typeface="Microsoft Himalaya" />
                <font script="Thaa" typeface="MV Boli" />
                <font script="Deva" typeface="Mangal" />
                <font script="Telu" typeface="Gautami" />
                <font script="Taml" typeface="Latha" />
                <font script="Syrc" typeface="Estrangelo Edessa" />
                <font script="Orya" typeface="Kalinga" />
                <font script="Mlym" typeface="Kartika" />
                <font script="Laoo" typeface="DokChampa" />
                <font script="Sinh" typeface="Iskoola Pota" />
                <font script="Mong" typeface="Mongolian Baiti" />
                <font script="Viet" typeface="Arial" />
                <font script="Uigh" typeface="Microsoft Uighur" />
                <font script="Geor" typeface="Sylfaen" />
                <font script="Armn" typeface="Arial" />
                <font script="Bugi" typeface="Leelawadee UI" />
                <font script="Bopo" typeface="Microsoft JhengHei" />
                <font script="Java" typeface="Javanese Text" />
                <font script="Lisu" typeface="Segoe UI" />
                <font script="Mymr" typeface="Myanmar Text" />
                <font script="Nkoo" typeface="Ebrima" />
                <font script="Olck" typeface="Nirmala UI" />
                <font script="Osma" typeface="Ebrima" />
                <font script="Phag" typeface="Phagspa" />
                <font script="Syrn" typeface="Estrangelo Edessa" />
                <font script="Syrj" typeface="Estrangelo Edessa" />
                <font script="Syre" typeface="Estrangelo Edessa" />
                <font script="Sora" typeface="Nirmala UI" />
                <font script="Tale" typeface="Microsoft Tai Le" />
                <font script="Talu" typeface="Microsoft New Tai Lue" />
                <font script="Tfng" typeface="Ebrima" />
            </minorFont>
        </fontScheme>
        <fmtScheme name="Office">
            <fillStyleLst>
                <solidFill>
                    <schemeClr val="phClr" />
                </solidFill>
                <gradFill rotWithShape="1">
                    <gsLst>
                        <gs pos="0">
                            <schemeClr val="phClr">
                                <lumMod val="110000" />
                                <satMod val="105000" />
                                <tint val="67000" />
                            </schemeClr>
                        </gs>
                        <gs pos="50000">
                            <schemeClr val="phClr">
                                <lumMod val="105000" />
                                <satMod val="103000" />
                                <tint val="73000" />
                            </schemeClr>
                        </gs>
                        <gs pos="100000">
                            <schemeClr val="phClr">
                                <lumMod val="105000" />
                                <satMod val="109000" />
                                <tint val="81000" />
                            </schemeClr>
                        </gs>
                    </gsLst>
                    <lin ang="5400000" scaled="0" />
                </gradFill>
                <gradFill rotWithShape="1">
                    <gsLst>
                        <gs pos="0">
                            <schemeClr val="phClr">
                                <satMod val="103000" />
                                <lumMod val="102000" />
                                <tint val="94000" />
                            </schemeClr>
                        </gs>
                        <gs pos="50000">
                            <schemeClr val="phClr">
                                <satMod val="110000" />
                                <lumMod val="100000" />
                                <shade val="100000" />
                            </schemeClr>
                        </gs>
                        <gs pos="100000">
                            <schemeClr val="phClr">
                                <lumMod val="99000" />
                                <satMod val="120000" />
                                <shade val="78000" />
                            </schemeClr>
                        </gs>
                    </gsLst>
                    <lin ang="5400000" scaled="0" />
                </gradFill>
            </fillStyleLst>
            <lnStyleLst>
                <ln w="6350" cap="flat" cmpd="sng" algn="ctr">
                    <solidFill>
                        <schemeClr val="phClr" />
                    </solidFill>
                    <prstDash val="solid" />
                    <miter lim="800000" />
                </ln>
                <ln w="12700" cap="flat" cmpd="sng" algn="ctr">
                    <solidFill>
                        <schemeClr val="phClr" />
                    </solidFill>
                    <prstDash val="solid" />
                    <miter lim="800000" />
                </ln>
                <ln w="19050" cap="flat" cmpd="sng" algn="ctr">
                    <solidFill>
                        <schemeClr val="phClr" />
                    </solidFill>
                    <prstDash val="solid" />
                    <miter lim="800000" />
                </ln>
            </lnStyleLst>
            <effectStyleLst>
                <effectStyle>
                    <effectLst />
                </effectStyle>
                <effectStyle>
                    <effectLst />
                </effectStyle>
                <effectStyle>
                    <effectLst>
                        <outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
                            <srgbClr val="000000">
                                <alpha val="63000" />
                            </srgbClr>
                        </outerShdw>
                    </effectLst>
                </effectStyle>
            </effectStyleLst>
            <bgFillStyleLst>
                <solidFill>
                    <schemeClr val="phClr" />
                </solidFill>
                <solidFill>
                    <schemeClr val="phClr">
                        <tint val="95000" />
                        <satMod val="170000" />
                    </schemeClr>
                </solidFill>
                <gradFill rotWithShape="1">
                    <gsLst>
                        <gs pos="0">
                            <schemeClr val="phClr">
                                <tint val="93000" />
                                <satMod val="150000" />
                                <shade val="98000" />
                                <lumMod val="102000" />
                            </schemeClr>
                        </gs>
                        <gs pos="50000">
                            <schemeClr val="phClr">
                                <tint val="98000" />
                                <satMod val="130000" />
                                <shade val="90000" />
                                <lumMod val="103000" />
                            </schemeClr>
                        </gs>
                        <gs pos="100000">
                            <schemeClr val="phClr">
                                <shade val="63000" />
                                <satMod val="120000" />
                            </schemeClr>
                        </gs>
                    </gsLst>
                    <lin ang="5400000" scaled="0" />
                </gradFill>
            </bgFillStyleLst>
        </fmtScheme>
    </themeElements>
    <objectDefaults />
    <extraClrSchemeLst />
    <extLst>
        <ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}">
            <themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}" />
        </ext>
    </extLst>
</theme>
`;

const workbookXmlRel = () =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" />
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />
</Relationships>
`;
