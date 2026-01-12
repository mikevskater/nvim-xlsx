--- Worksheet core module - base class and cell operations
--- @module nvim-xlsx.worksheet.core

local Cell = require("nvim-xlsx.cell")
local column_utils = require("nvim-xlsx.utils.column")
local date_utils = require("nvim-xlsx.utils.date")
local validation = require("nvim-xlsx.utils.validation")

local LIMITS = validation.LIMITS

local M = {}

---@class WorksheetFreezePane
---@field rows integer Number of rows to freeze at top (0 for none)
---@field cols integer Number of columns to freeze at left (0 for none)

---@class WorksheetAutoFilter
---@field ref string Range reference (e.g., "A1:D10")

---@class WorksheetDataValidation
---@field ref string Cell or range reference
---@field type string Validation type: "list", "whole", "decimal", "date", "time", "textLength", "custom"
---@field operator? string Operator: "between", "notBetween", "equal", "notEqual", "greaterThan", "lessThan", "greaterThanOrEqual", "lessThanOrEqual"
---@field formula1? string First formula/value
---@field formula2? string Second formula/value (for between)
---@field allowBlank? boolean Allow blank cells
---@field showDropDown? boolean Show dropdown for list validation (default true)
---@field showInputMessage? boolean Show input message
---@field showErrorMessage? boolean Show error message
---@field promptTitle? string Input message title
---@field prompt? string Input message text
---@field errorTitle? string Error message title
---@field error? string Error message text
---@field errorStyle? string Error style: "stop", "warning", "information"

---@class WorksheetHyperlink
---@field ref string Cell reference
---@field target string URL or cell reference
---@field location? string Internal location (sheet + cell)
---@field tooltip? string Tooltip text
---@field is_external boolean Whether it's an external link

---@class WorksheetPrintSettings
---@field orientation? string "portrait" or "landscape"
---@field paperSize? integer Paper size code (1=Letter, 9=A4)
---@field fitToWidth? integer Fit to n pages wide
---@field fitToHeight? integer Fit to n pages tall
---@field scale? integer Scale percentage (10-400)
---@field margins? table {top, bottom, left, right, header, footer} in inches
---@field printArea? string Range to print (e.g., "A1:G20")
---@field printTitleRows? string Rows to repeat at top (e.g., "1:1")
---@field printTitleCols? string Columns to repeat at left (e.g., "A:A")
---@field gridLines? boolean Print grid lines
---@field headings? boolean Print row/column headings

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
---@field freeze_pane? WorksheetFreezePane Frozen pane configuration
---@field auto_filter? WorksheetAutoFilter Auto-filter configuration
---@field data_validations WorksheetDataValidation[] List of data validations
---@field hyperlinks WorksheetHyperlink[] List of hyperlinks
---@field print_settings? WorksheetPrintSettings Print settings
local Worksheet = {}
Worksheet.__index = Worksheet

--- Create a new worksheet
--- @param name string Sheet name
--- @param index integer Sheet index
--- @param workbook table Parent workbook
--- @return Worksheet? worksheet
--- @return string? error_message
function M.new(name, index, workbook)
  local result = validation.validate_sheet_name(name)
  if not result.valid then
    return nil, result.error
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
  self.freeze_pane = nil
  self.auto_filter = nil
  self.data_validations = {}
  self.hyperlinks = {}
  self.print_settings = nil

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
  validation.check(validation.validate_row(row))
  validation.check(validation.validate_col(col))

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
  local ok, err = validation.check_soft(validation.validate_range(r1, c1, r2, c2))
  if not ok then
    return false, err
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

-- Export the class and constructor
M.Worksheet = Worksheet
M.LIMITS = LIMITS

return M
