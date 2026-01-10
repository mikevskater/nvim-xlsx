--- Worksheet representation for xlsx
--- @module xlsx.worksheet

local Cell = require("xlsx.cell")
local column_utils = require("xlsx.utils.column")
local xml = require("xlsx.xml.writer")
local templates = require("xlsx.xml.templates")

local M = {}

---@class Worksheet
---@field name string Sheet name
---@field index integer Sheet index (1-indexed)
---@field rows table<integer, table<integer, Cell>> Sparse row/column storage
---@field min_row integer Minimum row with data
---@field max_row integer Maximum row with data
---@field min_col integer Minimum column with data
---@field max_col integer Maximum column with data
---@field column_widths table<integer, number> Custom column widths
---@field row_heights table<integer, number> Custom row heights
---@field merged_cells table[] List of merged cell ranges
---@field workbook Workbook Parent workbook reference
local Worksheet = {}
Worksheet.__index = Worksheet

--- Validate sheet name
--- @param name string Sheet name to validate
--- @return boolean valid
--- @return string? error_message
local function validate_sheet_name(name)
  if not name or name == "" then
    return false, "Sheet name cannot be empty"
  end
  if #name > 31 then
    return false, "Sheet name cannot exceed 31 characters"
  end
  -- Check for forbidden characters
  if name:match("[\\/%*%?%[%]:]") then
    return false, "Sheet name cannot contain: \\ / * ? [ ] :"
  end
  -- Cannot start or end with apostrophe
  if name:sub(1, 1) == "'" or name:sub(-1) == "'" then
    return false, "Sheet name cannot start or end with apostrophe"
  end
  return true
end

--- Create a new worksheet
--- @param name string Sheet name
--- @param index integer Sheet index
--- @param workbook table Parent workbook
--- @return Worksheet? worksheet
--- @return string? error_message
function M.new(name, index, workbook)
  local valid, err = validate_sheet_name(name)
  if not valid then
    return nil, err
  end

  local self = setmetatable({}, Worksheet)
  self.name = name
  self.index = index
  self.workbook = workbook
  self.rows = {}
  self.min_row = nil
  self.max_row = nil
  self.min_col = nil
  self.max_col = nil
  self.column_widths = {}
  self.row_heights = {}
  self.merged_cells = {}

  return self
end

--- Update dimension tracking
--- @param row integer
--- @param col integer
function Worksheet:_update_dimensions(row, col)
  if not self.min_row or row < self.min_row then
    self.min_row = row
  end
  if not self.max_row or row > self.max_row then
    self.max_row = row
  end
  if not self.min_col or col < self.min_col then
    self.min_col = col
  end
  if not self.max_col or col > self.max_col then
    self.max_col = col
  end
end

--- Set a cell value
--- @param row integer Row number (1-indexed)
--- @param col integer Column number (1-indexed)
--- @param value any Cell value
--- @return Cell
function Worksheet:set_cell(row, col, value)
  -- Validate bounds
  if row < 1 or row > 1048576 then
    error("Row out of range (1-1048576): " .. tostring(row))
  end
  if col < 1 or col > 16384 then
    error("Column out of range (1-16384): " .. tostring(col))
  end

  -- Create row if needed
  if not self.rows[row] then
    self.rows[row] = {}
  end

  -- Create or update cell
  local cell = Cell.new(row, col, value)
  self.rows[row][col] = cell

  -- Update dimensions
  self:_update_dimensions(row, col)

  return cell
end

--- Get a cell (or nil if not set)
--- @param row integer Row number
--- @param col integer Column number
--- @return Cell?
function Worksheet:get_cell(row, col)
  if self.rows[row] then
    return self.rows[row][col]
  end
  return nil
end

--- Set a cell using A1 notation
--- @param ref string Cell reference (e.g., "A1")
--- @param value any Cell value
--- @return Cell
function Worksheet:set(ref, value)
  local parsed = column_utils.parse_ref(ref)
  return self:set_cell(parsed.row, parsed.col, value)
end

--- Get a cell using A1 notation
--- @param ref string Cell reference (e.g., "A1")
--- @return Cell?
function Worksheet:get(ref)
  local parsed = column_utils.parse_ref(ref)
  return self:get_cell(parsed.row, parsed.col)
end

--- Set column width
--- @param col integer Column number
--- @param width number Width in characters
function Worksheet:set_column_width(col, width)
  self.column_widths[col] = width
end

--- Set row height
--- @param row integer Row number
--- @param height number Height in points
function Worksheet:set_row_height(row, height)
  self.row_heights[row] = height
end

--- Set style for a cell
--- @param row integer Row number
--- @param col integer Column number
--- @param style_index integer Style index from workbook:create_style()
function Worksheet:set_cell_style(row, col, style_index)
  local cell = self:get_cell(row, col)
  if cell then
    cell.style_index = style_index
  else
    -- Create empty cell with style
    local Cell = require("xlsx.cell")
    if not self.rows[row] then
      self.rows[row] = {}
    end
    cell = Cell.new(row, col, nil)
    cell.style_index = style_index
    self.rows[row][col] = cell
    self:_update_dimensions(row, col)
  end
end

--- Set style for a range of cells
--- @param r1 integer Start row
--- @param c1 integer Start column
--- @param r2 integer End row
--- @param c2 integer End column
--- @param style_index integer Style index from workbook:create_style()
function Worksheet:set_range_style(r1, c1, r2, c2, style_index)
  for row = r1, r2 do
    for col = c1, c2 do
      self:set_cell_style(row, col, style_index)
    end
  end
end

--- Check if a range overlaps with any existing merged cell range
--- @param r1 integer Start row
--- @param c1 integer Start column
--- @param r2 integer End row
--- @param c2 integer End column
--- @return boolean overlaps
--- @return string? conflicting_range
local function check_merge_overlap(merged_cells, r1, c1, r2, c2)
  for _, merge_ref in ipairs(merged_cells) do
    local range = column_utils.parse_range(merge_ref)
    local mr1, mc1 = range.start.row, range.start.col
    local mr2, mc2 = range.finish.row, range.finish.col

    -- Check if ranges overlap
    local overlap = not (r2 < mr1 or r1 > mr2 or c2 < mc1 or c1 > mc2)
    if overlap then
      return true, merge_ref
    end
  end
  return false
end

--- Merge cells in a range
--- The value of the top-left cell will be displayed across the merged area
--- @param r1 integer Start row
--- @param c1 integer Start column
--- @param r2 integer End row
--- @param c2 integer End column
--- @return boolean success
--- @return string? error_message
function Worksheet:merge_cells(r1, c1, r2, c2)
  -- Validate bounds
  if r1 < 1 or r1 > 1048576 or r2 < 1 or r2 > 1048576 then
    return false, "Row out of range (1-1048576)"
  end
  if c1 < 1 or c1 > 16384 or c2 < 1 or c2 > 16384 then
    return false, "Column out of range (1-16384)"
  end

  -- Normalize coordinates (ensure r1 <= r2, c1 <= c2)
  if r1 > r2 then r1, r2 = r2, r1 end
  if c1 > c2 then c1, c2 = c2, c1 end

  -- Check for single cell (no merge needed)
  if r1 == r2 and c1 == c2 then
    return false, "Cannot merge a single cell"
  end

  -- Check for overlaps with existing merges
  local overlaps, conflict = check_merge_overlap(self.merged_cells, r1, c1, r2, c2)
  if overlaps then
    return false, "Range overlaps with existing merge: " .. conflict
  end

  -- Add the merge reference
  local merge_ref = column_utils.make_range(r1, c1, r2, c2)
  table.insert(self.merged_cells, merge_ref)

  -- Update dimensions to include the merged range
  self:_update_dimensions(r1, c1)
  self:_update_dimensions(r2, c2)

  return true
end

--- Merge cells using A1:B2 notation
--- @param range string Range reference (e.g., "A1:D5")
--- @return boolean success
--- @return string? error_message
function Worksheet:merge_range(range)
  local parsed = column_utils.parse_range(range)
  return self:merge_cells(
    parsed.start.row, parsed.start.col,
    parsed.finish.row, parsed.finish.col
  )
end

--- Set cell value with optional style
--- @param row integer Row number
--- @param col integer Column number
--- @param value any Cell value
--- @param style_index? integer Optional style index
--- @return Cell
function Worksheet:set_cell_value(row, col, value, style_index)
  local cell = self:set_cell(row, col, value)
  if style_index then
    cell.style_index = style_index
  end
  return cell
end

--- Set a formula in a cell
--- @param row integer Row number
--- @param col integer Column number
--- @param formula string Formula string (with or without leading =)
--- @param style_index? integer Optional style index
--- @return Cell
function Worksheet:set_formula(row, col, formula, style_index)
  -- Ensure formula starts with = for consistency
  if formula:sub(1, 1) ~= "=" then
    formula = "=" .. formula
  end
  return self:set_cell_value(row, col, formula, style_index)
end

--- Set a date value in a cell (as Excel serial number)
--- @param row integer Row number
--- @param col integer Column number
--- @param date_value number|table Excel serial number, or date table {year, month, day, hour?, min?, sec?}
--- @param style_index? integer Optional style index (recommend using a date number format)
--- @return Cell
function Worksheet:set_date(row, col, date_value, style_index)
  local serial
  if type(date_value) == "table" then
    local date_utils = require("xlsx.utils.date")
    serial = date_utils.to_serial(date_value)
  else
    serial = date_value
  end
  return self:set_cell_value(row, col, serial, style_index)
end

--- Set a boolean value in a cell
--- @param row integer Row number
--- @param col integer Column number
--- @param value boolean Boolean value
--- @param style_index? integer Optional style index
--- @return Cell
function Worksheet:set_boolean(row, col, value, style_index)
  -- The cell module handles boolean detection automatically
  return self:set_cell_value(row, col, value, style_index)
end

--- Get the dimension string (e.g., "A1:C10")
--- @return string
function Worksheet:get_dimension()
  if not self.min_row then
    return "A1"
  end
  return column_utils.make_range(self.min_row, self.min_col, self.max_row, self.max_col)
end

--- Generate the sheetData XML content
--- @return string
function Worksheet:_generate_sheet_data()
  if not self.min_row then
    return ""
  end

  local parts = {}

  -- Get sorted row numbers
  local row_nums = {}
  for row_num in pairs(self.rows) do
    table.insert(row_nums, row_num)
  end
  table.sort(row_nums)

  -- Generate each row
  for _, row_num in ipairs(row_nums) do
    local row_data = self.rows[row_num]
    local cell_parts = {}

    -- Get sorted column numbers for this row
    local col_nums = {}
    for col_num in pairs(row_data) do
      table.insert(col_nums, col_num)
    end
    table.sort(col_nums)

    -- Generate cells
    for _, col_num in ipairs(col_nums) do
      local cell = row_data[col_num]
      if cell and cell:has_content() then
        table.insert(cell_parts, cell:to_xml())
      end
    end

    if #cell_parts > 0 then
      local row_attrs = { r = row_num }
      -- Add custom height if set
      if self.row_heights[row_num] then
        row_attrs.ht = self.row_heights[row_num]
        row_attrs.customHeight = "1"
      end
      table.insert(parts, xml.element_raw("row", table.concat(cell_parts), row_attrs))
    end
  end

  return table.concat(parts)
end

--- Generate column definitions XML
--- @return string
function Worksheet:_generate_cols()
  if not next(self.column_widths) then
    return ""
  end

  local parts = {}
  local cols = {}
  for col in pairs(self.column_widths) do
    table.insert(cols, col)
  end
  table.sort(cols)

  for _, col in ipairs(cols) do
    local width = self.column_widths[col]
    table.insert(parts, xml.empty_element("col", {
      min = col,
      max = col,
      width = width,
      customWidth = "1",
    }))
  end

  return xml.element_raw("cols", table.concat(parts))
end

--- Generate the complete worksheet XML
--- @param is_active? boolean Whether this sheet is the active/selected sheet
--- @return string
function Worksheet:to_xml(is_active)
  local b = xml.builder()

  b:declaration()
  b:open("worksheet", {
    xmlns = templates.NS.SPREADSHEET,
    ["xmlns:r"] = templates.NS.RELATIONSHIPS,
  })

  -- Dimension
  b:empty("dimension", { ref = self:get_dimension() })

  -- Sheet views with tabSelected based on active status
  local tab_selected = is_active and "1" or "0"
  b:raw('<sheetViews><sheetView tabSelected="' .. tab_selected .. '" workbookViewId="0"/></sheetViews>')

  -- Sheet format defaults
  b:empty("sheetFormatPr", { defaultRowHeight = "15" })

  -- Column widths
  local cols = self:_generate_cols()
  if cols ~= "" then
    b:raw(cols)
  end

  -- Sheet data
  local sheet_data = self:_generate_sheet_data()
  b:elem_raw("sheetData", sheet_data)

  -- Merged cells (if any)
  if #self.merged_cells > 0 then
    local merge_parts = {}
    for _, merge in ipairs(self.merged_cells) do
      table.insert(merge_parts, xml.empty_element("mergeCell", { ref = merge }))
    end
    b:elem_raw("mergeCells", table.concat(merge_parts), { count = #self.merged_cells })
  end

  b:close("worksheet")

  return b:to_string()
end

M.Worksheet = Worksheet

return M
