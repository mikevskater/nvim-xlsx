--- Workbook XML parsing for xlsx
--- @module nvim-xlsx.parts.workbook_part
---
--- Handles parsing of xl/workbook.xml and xl/_rels/workbook.xml.rels

local parser = require("nvim-xlsx.xml.parser")

local M = {}

---@class SheetInfo
---@field name string Sheet name
---@field sheet_id integer Sheet ID
---@field r_id string Relationship ID (e.g., "rId1")
---@field state? string Sheet state ("visible", "hidden", "veryHidden")
---@field target? string Target file path from relationships

---@class WorkbookInfo
---@field sheets SheetInfo[] Array of sheet information
---@field active_sheet integer? Index of active sheet (0-indexed)
---@field relationships table<string, string> Map of rId to target path

--- Parse relationships from .rels XML content
--- @param xml_content string The XML content of workbook.xml.rels
--- @return table<string, string> Map of rId to target path
function M.parse_relationships(xml_content)
  local rels = {}

  -- Find all <Relationship> elements
  local rel_elements = parser.find_all(xml_content, "Relationship")

  for _, rel in ipairs(rel_elements) do
    local id = rel.attrs.Id
    local target = rel.attrs.Target
    if id and target then
      rels[id] = target
    end
  end

  return rels
end

--- Parse workbook XML content
--- @param xml_content string The XML content of workbook.xml
--- @param rels_content? string The XML content of workbook.xml.rels (optional)
--- @return WorkbookInfo
function M.parse(xml_content, rels_content)
  local info = {
    sheets = {},
    active_sheet = nil,
    relationships = {},
  }

  -- Parse relationships if provided
  if rels_content then
    info.relationships = M.parse_relationships(rels_content)
  end

  -- Parse workbookView for active sheet
  local workbook_view = xml_content:match("<workbookView[^>]*activeTab=\"(%d+)\"")
  if workbook_view then
    info.active_sheet = tonumber(workbook_view)
  end

  -- Find all <sheet> elements
  local sheet_elements = parser.find_all(xml_content, "sheet")

  for _, sheet in ipairs(sheet_elements) do
    local sheet_info = {
      name = sheet.attrs.name or ("Sheet" .. #info.sheets + 1),
      sheet_id = tonumber(sheet.attrs.sheetId) or (#info.sheets + 1),
      r_id = sheet.attrs["r:id"],
      state = sheet.attrs.state,  -- nil = visible, "hidden", "veryHidden"
    }

    -- Look up target from relationships
    if sheet_info.r_id and info.relationships[sheet_info.r_id] then
      sheet_info.target = info.relationships[sheet_info.r_id]
    end

    table.insert(info.sheets, sheet_info)
  end

  return info
end

--- Parse root relationships (_rels/.rels)
--- @param xml_content string The XML content of .rels
--- @return table<string, table> Map of relationship type to info
function M.parse_root_rels(xml_content)
  local rels = {}

  local rel_elements = parser.find_all(xml_content, "Relationship")

  for _, rel in ipairs(rel_elements) do
    local rel_type = rel.attrs.Type
    local target = rel.attrs.Target
    local id = rel.attrs.Id

    if rel_type then
      -- Extract the relationship type name from the full URI
      local type_name = rel_type:match("([^/]+)$")
      rels[type_name] = {
        id = id,
        type = rel_type,
        target = target,
      }
    end
  end

  return rels
end

--- Parse content types ([Content_Types].xml)
--- @param xml_content string The XML content
--- @return table Content types info
function M.parse_content_types(xml_content)
  local types = {
    defaults = {},     -- Extension -> ContentType
    overrides = {},    -- PartName -> ContentType
  }

  -- Parse Default elements
  local default_elements = parser.find_all(xml_content, "Default")
  for _, elem in ipairs(default_elements) do
    local ext = elem.attrs.Extension
    local content_type = elem.attrs.ContentType
    if ext and content_type then
      types.defaults[ext] = content_type
    end
  end

  -- Parse Override elements
  local override_elements = parser.find_all(xml_content, "Override")
  for _, elem in ipairs(override_elements) do
    local part_name = elem.attrs.PartName
    local content_type = elem.attrs.ContentType
    if part_name and content_type then
      types.overrides[part_name] = content_type
    end
  end

  return types
end

return M
