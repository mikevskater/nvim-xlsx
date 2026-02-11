--- Workbook representation for xlsx
--- @module nvim-xlsx.workbook

local Worksheet = require("nvim-xlsx.worksheet")
local xml = require("nvim-xlsx.xml.writer")
local templates = require("nvim-xlsx.xml.templates")
local zip = require("nvim-xlsx.zip")
local doc_props = require("nvim-xlsx.parts.doc_props")
local Style = require("nvim-xlsx.style")
local column_utils = require("nvim-xlsx.utils.column")
local validation = require("nvim-xlsx.utils.validation")

local M = {}

---@class Workbook
---@field sheets Worksheet[] List of worksheets
---@field sheet_map table<string, Worksheet> Map of sheet names to worksheets
---@field properties table Document properties
---@field active_sheet integer Index of active sheet
---@field styles StyleRegistry Style registry for this workbook
---@field _table_id_counter integer Global table ID counter
---@field defined_names table[] List of defined names (named ranges)
local Workbook = {}
Workbook.__index = Workbook

--- Create a new workbook
--- @return Workbook
function M.new()
  local self = setmetatable({}, Workbook)
  self.sheets = {}
  self.sheet_map = {}
  self.properties = {
    creator = "nvim-xlsx",
    created = os.date("!%Y-%m-%dT%H:%M:%SZ"),
    modified = os.date("!%Y-%m-%dT%H:%M:%SZ"),
  }
  self.active_sheet = 1
  self.styles = Style.new_registry()
  self._table_id_counter = 0
  self.defined_names = {}
  return self
end

--- Add a new worksheet
--- @param name? string Sheet name (defaults to "Sheet1", "Sheet2", etc.)
--- @return Worksheet? worksheet
--- @return string? error_message
function Workbook:add_sheet(name)
  -- Generate default name if not provided
  if not name then
    local num = #self.sheets + 1
    name = "Sheet" .. num
    -- Ensure unique
    while self.sheet_map[name] do
      num = num + 1
      name = "Sheet" .. num
    end
  end

  -- Check for duplicate name
  if self.sheet_map[name] then
    return nil, "Sheet with name '" .. name .. "' already exists"
  end

  local index = #self.sheets + 1
  local sheet, err = Worksheet.new(name, index, self)
  if not sheet then
    return nil, err
  end

  table.insert(self.sheets, sheet)
  self.sheet_map[name] = sheet

  return sheet
end

--- Get a worksheet by name or index
--- @param name_or_index string|integer Sheet name or index
--- @return Worksheet?
function Workbook:get_sheet(name_or_index)
  if type(name_or_index) == "number" then
    return self.sheets[name_or_index]
  else
    return self.sheet_map[name_or_index]
  end
end

--- Set the active (selected) sheet
--- @param name_or_index string|integer Sheet name or index
--- @return boolean success
function Workbook:set_active_sheet(name_or_index)
  if type(name_or_index) == "number" then
    if name_or_index >= 1 and name_or_index <= #self.sheets then
      self.active_sheet = name_or_index
      return true
    end
  else
    for i, sheet in ipairs(self.sheets) do
      if sheet.name == name_or_index then
        self.active_sheet = i
        return true
      end
    end
  end
  return false
end

--- Set document properties
--- @param props table Properties to set (creator, title, subject, etc.)
function Workbook:set_properties(props)
  for k, v in pairs(props) do
    self.properties[k] = v
  end
  -- Update modified timestamp
  self.properties.modified = os.date("!%Y-%m-%dT%H:%M:%SZ")
end

--- Create a style and return its index
--- @param def table Style definition
--- @return integer Style index
function Workbook:create_style(def)
  return self.styles:create_style(def)
end

--- Get the next globally unique table ID
--- @return integer
function Workbook:_next_table_id()
  self._table_id_counter = self._table_id_counter + 1
  return self._table_id_counter
end

--- Add a raw defined name to the workbook
--- @param name string The defined name
--- @param ref string The reference string (pre-formatted, e.g., "Sales!$A$1:$C$4" or "42")
--- @param options? table Optional: { local_sheet_id?: integer, hidden?: boolean, comment?: string }
--- @return Workbook self For chaining
function Workbook:add_defined_name(name, ref, options)
  options = options or {}
  validation.check(validation.validate_defined_name(name), "add_defined_name")
  table.insert(self.defined_names, {
    name = name,
    ref = ref,
    local_sheet_id = options.local_sheet_id,
    hidden = options.hidden or false,
    comment = options.comment,
  })
  return self
end

--- Add a named range (convenience method that builds the ref string)
--- @param name string The range name
--- @param sheet_name string The sheet name
--- @param r1 integer Start row
--- @param c1 integer Start column
--- @param r2 integer End row
--- @param c2 integer End column
--- @param options? table Optional: { local_sheet_id?: integer, hidden?: boolean, comment?: string }
--- @return Workbook self For chaining
function Workbook:add_named_range(name, sheet_name, r1, c1, r2, c2, options)
  -- Quote sheet name if it contains spaces
  local quoted_name
  if sheet_name:find(" ") then
    quoted_name = "'" .. sheet_name .. "'"
  else
    quoted_name = sheet_name
  end
  local abs_range = column_utils.make_abs_range(r1, c1, r2, c2)
  local ref = quoted_name .. "!" .. abs_range
  return self:add_defined_name(name, ref, options)
end

--- Generate table XML for a single ExcelTable
--- @param tbl ExcelTable
--- @return string
function Workbook:_generate_table_xml(tbl)
  local b = xml.builder()
  b:declaration()
  b:open("table", {
    xmlns = templates.NS.SPREADSHEET,
    id = tbl.id,
    name = tbl.name,
    displayName = tbl.name,
    ref = tbl.ref,
    totalsRowShown = "0",
  })

  -- Auto-filter
  if tbl.auto_filter then
    b:empty("autoFilter", { ref = tbl.ref })
  end

  -- Table columns
  local col_parts = {}
  for _, col in ipairs(tbl.columns) do
    table.insert(col_parts, xml.empty_element("tableColumn", {
      id = col.id,
      name = col.name,
    }))
  end
  b:elem_raw("tableColumns", table.concat(col_parts), { count = #tbl.columns })

  -- Table style info
  b:empty("tableStyleInfo", {
    name = tbl.style_name,
    showFirstColumn = tbl.show_first_col and "1" or "0",
    showLastColumn = tbl.show_last_col and "1" or "0",
    showRowStripes = tbl.show_row_stripes and "1" or "0",
    showColumnStripes = tbl.show_col_stripes and "1" or "0",
  })

  b:close("table")
  return b:to_string()
end

--- Generate [Content_Types].xml
--- @return string
function Workbook:_generate_content_types()
  local b = xml.builder()
  b:declaration()
  b:open("Types", { xmlns = templates.NS.CONTENT_TYPES })

  -- Default extensions
  b:empty("Default", { Extension = "rels", ContentType = templates.DEFAULT_EXTENSIONS.rels })
  b:empty("Default", { Extension = "xml", ContentType = templates.DEFAULT_EXTENSIONS.xml })

  -- Document properties
  b:empty("Override", {
    PartName = "/docProps/core.xml",
    ContentType = templates.CONTENT_TYPES.CORE_PROPS,
  })
  b:empty("Override", {
    PartName = "/docProps/app.xml",
    ContentType = templates.CONTENT_TYPES.EXT_PROPS,
  })

  -- Override for workbook
  b:empty("Override", {
    PartName = "/xl/workbook.xml",
    ContentType = templates.CONTENT_TYPES.WORKBOOK,
  })

  -- Override for each worksheet
  for i = 1, #self.sheets do
    b:empty("Override", {
      PartName = "/xl/worksheets/sheet" .. i .. ".xml",
      ContentType = templates.CONTENT_TYPES.WORKSHEET,
    })
  end

  -- Override for each table
  for _, sheet in ipairs(self.sheets) do
    for _, tbl in ipairs(sheet.tables) do
      b:empty("Override", {
        PartName = "/xl/tables/table" .. tbl.id .. ".xml",
        ContentType = templates.CONTENT_TYPES.TABLE,
      })
    end
  end

  -- Override for styles
  b:empty("Override", {
    PartName = "/xl/styles.xml",
    ContentType = templates.CONTENT_TYPES.STYLES,
  })

  b:close("Types")
  return b:to_string()
end

--- Generate _rels/.rels
--- @return string
function Workbook:_generate_root_rels()
  local b = xml.builder()
  b:declaration()
  b:open("Relationships", { xmlns = templates.NS.PACKAGE_RELS })

  b:empty("Relationship", {
    Id = "rId1",
    Type = templates.NS.REL_OFFICE_DOC,
    Target = "xl/workbook.xml",
  })

  b:empty("Relationship", {
    Id = "rId2",
    Type = templates.NS.REL_CORE_PROPS,
    Target = "docProps/core.xml",
  })

  b:empty("Relationship", {
    Id = "rId3",
    Type = templates.NS.REL_EXT_PROPS,
    Target = "docProps/app.xml",
  })

  b:close("Relationships")
  return b:to_string()
end

--- Generate xl/_rels/workbook.xml.rels
--- @return string
function Workbook:_generate_workbook_rels()
  local b = xml.builder()
  b:declaration()
  b:open("Relationships", { xmlns = templates.NS.PACKAGE_RELS })

  local rid = 1

  -- Worksheet relationships
  for i = 1, #self.sheets do
    b:empty("Relationship", {
      Id = "rId" .. rid,
      Type = templates.NS.REL_WORKSHEET,
      Target = "worksheets/sheet" .. i .. ".xml",
    })
    rid = rid + 1
  end

  -- Styles relationship
  b:empty("Relationship", {
    Id = "rId" .. rid,
    Type = templates.NS.REL_STYLES,
    Target = "styles.xml",
  })

  b:close("Relationships")
  return b:to_string()
end

--- Generate xl/workbook.xml
--- @return string
function Workbook:_generate_workbook_xml()
  local b = xml.builder()
  b:declaration()
  b:open("workbook", {
    xmlns = templates.NS.SPREADSHEET,
    ["xmlns:r"] = templates.NS.RELATIONSHIPS,
  })

  -- Workbook views with active tab
  local active_tab = (self.active_sheet or 1) - 1  -- 0-indexed
  b:raw('<bookViews><workbookView activeTab="' .. active_tab .. '"/></bookViews>')

  -- Sheets
  b:open("sheets")
  for i, sheet in ipairs(self.sheets) do
    b:empty("sheet", {
      name = sheet.name,
      sheetId = i,
      ["r:id"] = "rId" .. i,
    })
  end
  b:close("sheets")

  -- Defined names (named ranges)
  if #self.defined_names > 0 then
    b:open("definedNames")
    for _, dn in ipairs(self.defined_names) do
      local attrs = { name = dn.name }
      if dn.local_sheet_id then
        attrs.localSheetId = dn.local_sheet_id
      end
      if dn.hidden then
        attrs.hidden = "1"
      end
      if dn.comment then
        attrs.comment = dn.comment
      end
      b:elem("definedName", dn.ref, attrs)
    end
    b:close("definedNames")
  end

  b:close("workbook")
  return b:to_string()
end

--- Generate xl/styles.xml
--- @return string
function Workbook:_generate_styles_xml()
  return self.styles:to_xml()
end

--- Generate xl/worksheets/_rels/sheetN.xml.rels for a worksheet
--- @param sheet Worksheet The worksheet
--- @return string? XML content, or nil if no relationships needed
function Workbook:_generate_worksheet_rels(sheet)
  local table_rels = sheet:get_table_relationships()
  local hyperlink_rels = sheet:get_hyperlink_relationships()
  if #table_rels == 0 and #hyperlink_rels == 0 then
    return nil
  end

  local b = xml.builder()
  b:declaration()
  b:open("Relationships", { xmlns = templates.NS.PACKAGE_RELS })

  -- Table relationships first (rId1, rId2, ...)
  for _, rel in ipairs(table_rels) do
    b:empty("Relationship", {
      Id = rel.id,
      Type = rel.type,
      Target = rel.target,
    })
  end

  -- Hyperlink relationships (rId1001+, no collision)
  for _, rel in ipairs(hyperlink_rels) do
    b:empty("Relationship", {
      Id = rel.id,
      Type = rel.type,
      Target = rel.target,
      TargetMode = rel.targetMode,
    })
  end

  b:close("Relationships")
  return b:to_string()
end

--- Save workbook to file
--- @param filepath string Output file path
--- @return boolean success
--- @return string? error_message
function Workbook:save(filepath)
  -- Ensure we have at least one sheet
  if #self.sheets == 0 then
    local _, err = self:add_sheet()
    if err then
      return false, err
    end
  end

  -- Update modified timestamp
  self.properties.modified = os.date("!%Y-%m-%dT%H:%M:%SZ")

  -- Create temporary directory
  local temp_dir = zip.create_temp_dir()

  local function cleanup_and_fail(msg)
    zip.cleanup_temp_dir(temp_dir)
    return false, msg
  end

  -- Write [Content_Types].xml
  local ok, err = zip.write_file(
    temp_dir .. "/[Content_Types].xml",
    self:_generate_content_types()
  )
  if not ok then return cleanup_and_fail(err) end

  -- Write _rels/.rels
  ok, err = zip.write_file(
    temp_dir .. "/_rels/.rels",
    self:_generate_root_rels()
  )
  if not ok then return cleanup_and_fail(err) end

  -- Write docProps/core.xml
  ok, err = zip.write_file(
    temp_dir .. "/docProps/core.xml",
    doc_props.generate_core(self.properties)
  )
  if not ok then return cleanup_and_fail(err) end

  -- Write docProps/app.xml
  ok, err = zip.write_file(
    temp_dir .. "/docProps/app.xml",
    doc_props.generate_app(self.properties, self)
  )
  if not ok then return cleanup_and_fail(err) end

  -- Write xl/workbook.xml
  ok, err = zip.write_file(
    temp_dir .. "/xl/workbook.xml",
    self:_generate_workbook_xml()
  )
  if not ok then return cleanup_and_fail(err) end

  -- Write xl/_rels/workbook.xml.rels
  ok, err = zip.write_file(
    temp_dir .. "/xl/_rels/workbook.xml.rels",
    self:_generate_workbook_rels()
  )
  if not ok then return cleanup_and_fail(err) end

  -- Write xl/styles.xml
  ok, err = zip.write_file(
    temp_dir .. "/xl/styles.xml",
    self:_generate_styles_xml()
  )
  if not ok then return cleanup_and_fail(err) end

  -- Write each worksheet and its relationships
  for i, sheet in ipairs(self.sheets) do
    -- Pass whether this sheet is active
    local is_active = (i == self.active_sheet)
    ok, err = zip.write_file(
      temp_dir .. "/xl/worksheets/sheet" .. i .. ".xml",
      sheet:to_xml(is_active)
    )
    if not ok then return cleanup_and_fail(err) end

    -- Write worksheet relationships if needed (for tables and external hyperlinks)
    local sheet_rels = self:_generate_worksheet_rels(sheet)
    if sheet_rels then
      ok, err = zip.write_file(
        temp_dir .. "/xl/worksheets/_rels/sheet" .. i .. ".xml.rels",
        sheet_rels
      )
      if not ok then return cleanup_and_fail(err) end
    end

    -- Write table XML files
    for _, tbl in ipairs(sheet.tables) do
      ok, err = zip.write_file(
        temp_dir .. "/xl/tables/table" .. tbl.id .. ".xml",
        self:_generate_table_xml(tbl)
      )
      if not ok then return cleanup_and_fail(err) end
    end
  end

  -- Create ZIP file
  ok, err = zip.zip_directory(temp_dir, filepath)
  if not ok then return cleanup_and_fail(err) end

  -- Cleanup
  zip.cleanup_temp_dir(temp_dir)

  return true
end

M.Workbook = Workbook

return M
