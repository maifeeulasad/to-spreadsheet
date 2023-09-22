import { format } from "date-fns";

interface IGenerateCoreXml {
    username?: string
    description?: string
    language?: string
    version?: string
    subject?: string
    title?: string
}

const generateCoreXml = ({ username = "to-spreadsheet", description = "", language = "en-US", version = "1", subject = "", title = "" }: IGenerateCoreXml) => {
    const date =
        format(Date.now(), "yyyy-MM-dd") +
        "T" +
        format(Date.now(), "HH:mm:ss") +
        "Z";

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <cp:coreProperties
          xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" 
          xmlns:dc="http://purl.org/dc/elements/1.1/" 
          xmlns:dcterms="http://purl.org/dc/terms/" 
          xmlns:dcmitype="http://purl.org/dc/dcmitype/" 
          xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
      <dcterms:created xsi:type="dcterms:W3CDTF">${date}</dcterms:created>
      <dc:creator>
        ${username}
    </dc:creator>
      <dc:description>
      ${description}</dc:description>
      <dc:language>${language}</dc:language>
      <cp:lastModifiedBy>${date}</cp:lastModifiedBy>
      <dcterms:modified xsi:type="dcterms:W3CDTF">${date}</dcterms:modified>
      <cp:revision>${version}</cp:revision>
      <dc:subject>${subject}</dc:subject>
      <dc:title>${title}</dc:title>
  </cp:coreProperties>
`;
};

export { generateCoreXml };
