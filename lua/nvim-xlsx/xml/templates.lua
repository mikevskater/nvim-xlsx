--- XML templates and namespace constants for xlsx
--- @module xlsx.xml.templates

local M = {}

-- Namespace URIs used in xlsx files
M.NS = {
  -- Main namespaces
  SPREADSHEET = "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types",
  PACKAGE_RELS = "http://schemas.openxmlformats.org/package/2006/relationships",

  -- Relationship types
  REL_OFFICE_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
  REL_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
  REL_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
  REL_SHARED_STRINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
  REL_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",

  -- Document properties
  CORE_PROPS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
  DC = "http://purl.org/dc/elements/1.1/",
  DCTERMS = "http://purl.org/dc/terms/",
  DCMITYPE = "http://purl.org/dc/dcmitype/",
  XSI = "http://www.w3.org/2001/XMLSchema-instance",

  -- Extended properties
  EXT_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
  VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes",

  -- Relationship types for properties
  REL_CORE_PROPS = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
  REL_EXT_PROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
  REL_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
}

-- Content types for xlsx parts
M.CONTENT_TYPES = {
  WORKBOOK = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
  WORKSHEET = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
  STYLES = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
  SHARED_STRINGS = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
  THEME = "application/vnd.openxmlformats-officedocument.theme+xml",
  CORE_PROPS = "application/vnd.openxmlformats-package.core-properties+xml",
  EXT_PROPS = "application/vnd.openxmlformats-officedocument.extended-properties+xml",
  RELS = "application/vnd.openxmlformats-package.relationships+xml",
  TABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
}

-- Default extension mappings
M.DEFAULT_EXTENSIONS = {
  rels = "application/vnd.openxmlformats-package.relationships+xml",
  xml = "application/xml",
}

return M
