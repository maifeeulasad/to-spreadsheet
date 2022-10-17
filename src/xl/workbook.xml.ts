const generateWorkBookXml = () =>
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
`

export {generateWorkBookXml}