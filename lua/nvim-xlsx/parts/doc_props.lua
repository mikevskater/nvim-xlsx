--- Document properties generation for xlsx
--- Generates docProps/core.xml and docProps/app.xml
--- @module nvim-xlsx.parts.doc_props

local xml = require("nvim-xlsx.xml.writer")
local templates = require("nvim-xlsx.xml.templates")

local M = {}

--- Generate docProps/core.xml (Dublin Core metadata)
--- @param props table Document properties
--- @return string XML content
function M.generate_core(props)
  props = props or {}

  local b = xml.builder()
  b:declaration()
  b:open("cp:coreProperties", {
    ["xmlns:cp"] = templates.NS.CORE_PROPS,
    ["xmlns:dc"] = templates.NS.DC,
    ["xmlns:dcterms"] = templates.NS.DCTERMS,
    ["xmlns:dcmitype"] = templates.NS.DCMITYPE,
    ["xmlns:xsi"] = templates.NS.XSI,
  })

  -- Creator
  if props.creator then
    b:elem("dc:creator", props.creator)
  end

  -- Last modified by
  if props.last_modified_by then
    b:elem("cp:lastModifiedBy", props.last_modified_by)
  elseif props.creator then
    b:elem("cp:lastModifiedBy", props.creator)
  end

  -- Created date
  if props.created then
    b:raw('<dcterms:created xsi:type="dcterms:W3CDTF">' .. props.created .. '</dcterms:created>')
  end

  -- Modified date
  if props.modified then
    b:raw('<dcterms:modified xsi:type="dcterms:W3CDTF">' .. props.modified .. '</dcterms:modified>')
  end

  -- Title
  if props.title then
    b:elem("dc:title", props.title)
  end

  -- Subject
  if props.subject then
    b:elem("dc:subject", props.subject)
  end

  -- Description
  if props.description then
    b:elem("dc:description", props.description)
  end

  -- Keywords
  if props.keywords then
    b:elem("cp:keywords", props.keywords)
  end

  -- Category
  if props.category then
    b:elem("cp:category", props.category)
  end

  b:close("cp:coreProperties")
  return b:to_string()
end

--- Generate docProps/app.xml (Application properties)
--- @param props table Document properties
--- @param workbook table Workbook for sheet information
--- @return string XML content
function M.generate_app(props, workbook)
  props = props or {}

  local b = xml.builder()
  b:declaration()
  b:open("Properties", {
    xmlns = templates.NS.EXT_PROPS,
    ["xmlns:vt"] = templates.NS.VT,
  })

  -- Application name
  b:elem("Application", props.application or "nvim-xlsx")

  -- Document security (0 = none)
  b:elem("DocSecurity", "0")

  -- Scale crop
  b:elem("ScaleCrop", "false")

  -- Heading pairs (Sheet names header)
  if workbook and #workbook.sheets > 0 then
    b:open("HeadingPairs")
    b:open("vt:vector", { size = "2", baseType = "variant" })
    b:raw("<vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>")
    b:raw("<vt:variant><vt:i4>" .. #workbook.sheets .. "</vt:i4></vt:variant>")
    b:close("vt:vector")
    b:close("HeadingPairs")

    -- Title of parts (sheet names)
    b:open("TitlesOfParts")
    b:open("vt:vector", { size = tostring(#workbook.sheets), baseType = "lpstr" })
    for _, sheet in ipairs(workbook.sheets) do
      b:elem("vt:lpstr", sheet.name)
    end
    b:close("vt:vector")
    b:close("TitlesOfParts")
  end

  -- Company
  if props.company then
    b:elem("Company", props.company)
  end

  -- Links up to date
  b:elem("LinksUpToDate", "false")

  -- Shared doc
  b:elem("SharedDoc", "false")

  -- Hyperlinks changed
  b:elem("HyperlinksChanged", "false")

  -- App version
  b:elem("AppVersion", props.app_version or "1.0")

  b:close("Properties")
  return b:to_string()
end

return M
