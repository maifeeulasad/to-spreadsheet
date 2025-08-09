/**
 * @fileoverview Excel core properties XML generation
 * Handles the generation of core.xml which contains core document properties and metadata
 * This file includes document creation info, author details, language, version, and other core metadata
 * 
 * @author Maifee Ul Asad <maifeeulasad@gmail.com>
 * @license MIT
 */

/**
 * Interface for core XML generation parameters
 * Defines optional metadata that can be included in the core properties
 * @interface IGenerateCoreXml
 */
interface IGenerateCoreXml {
  /** Username/creator of the document */
  username?: string
  /** Document description */
  description?: string
  /** Document language code (e.g., 'en-US') */
  language?: string
  /** Document version/revision number */
  version?: string
  /** Document subject */
  subject?: string
  /** Document title */
  title?: string
}

/**
 * Formats a JavaScript Date object to ISO 8601 format required by Excel
 * @param {Date} date - The date to format
 * @returns {string} ISO 8601 formatted date string (YYYY-MM-DDTHH:mm:ssZ)
 * @internal
 */
const formatDate = (date: Date) => {
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const day = String(date.getUTCDate()).padStart(2, '0');
  const hours = String(date.getUTCHours()).padStart(2, '0');
  const minutes = String(date.getUTCMinutes()).padStart(2, '0');
  const seconds = String(date.getUTCSeconds()).padStart(2, '0');

  return `${year}-${month}-${day}T${hours}:${minutes}:${seconds}Z`;
};

/**
 * Generates the core.xml file for an Excel workbook
 * Creates core document properties including creation date, author, language, and other metadata
 * This file is part of the OpenXML core properties specification
 * @param {IGenerateCoreXml} options - Configuration object with optional metadata properties
 * @param {string} [options.username="to-spreadsheet"] - Document creator username
 * @param {string} [options.description=""] - Document description
 * @param {string} [options.language="en-US"] - Document language
 * @param {string} [options.version="1"] - Document version/revision
 * @param {string} [options.subject=""] - Document subject
 * @param {string} [options.title=""] - Document title
 * @returns {string} Complete XML content for core.xml file
 * @internal
 */
const generateCoreXml = ({ username = "to-spreadsheet", description = "", language = "en-US", version = "1", subject = "", title = "" }: IGenerateCoreXml) => {
  const date = formatDate(new Date());

  /**
   * Complete XML template for Excel core.xml file
   * Includes creation date, creator, language, modification info, and other core metadata
   * Follows OpenXML specification for core document properties
   */
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

/**
 * Exports the core properties XML generation function
 * @name generateCoreXml
 * @function
 * @description Generates core.xml containing core document properties and metadata
 * @see {@link generateCoreXml} - Main core properties generation function
 */
export { generateCoreXml };
