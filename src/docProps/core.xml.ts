// todo: time with moment js ->  2022-08-11T16:18:46Z
const generateCoreXml = (username: string = "maifeeulasad/to-spreadsheet") =>
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

export {generateCoreXml}