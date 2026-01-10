--- Worksheet representation for xlsx
--- @module xlsx.worksheet

local Cell = require("xlsx.cell")
local column_utils = require("xlsx.utils.column")
local xml = require("xlsx.xml.writer")
local templates = require("xlsx.xml.templates")

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

-- ============================================
-- Phase 6: Advanced Features
-- ============================================

--- Freeze panes at the specified row and/or column
--- @param rows integer Number of rows to freeze at top (0 for none)
--- @param cols integer Number of columns to freeze at left (0 for none)
--- @return Worksheet self For chaining
function Worksheet:freeze_panes(rows, cols)
  rows = rows or 0
  cols = cols or 0

  if rows < 0 or rows > 1048576 then
    error("Freeze rows out of range (0-1048576): " .. tostring(rows))
  end
  if cols < 0 or cols > 16384 then
    error("Freeze cols out of range (0-16384): " .. tostring(cols))
  end

  if rows == 0 and cols == 0 then
    self.freeze_pane = nil
  else
    self.freeze_pane = { rows = rows, cols = cols }
  end

  return self
end

--- Freeze the first N rows
--- @param rows integer Number of rows to freeze
--- @return Worksheet self For chaining
function Worksheet:freeze_rows(rows)
  return self:freeze_panes(rows, self.freeze_pane and self.freeze_pane.cols or 0)
end

--- Freeze the first N columns
--- @param cols integer Number of columns to freeze
--- @return Worksheet self For chaining
function Worksheet:freeze_cols(cols)
  return self:freeze_panes(self.freeze_pane and self.freeze_pane.rows or 0, cols)
end

--- Set auto-filter for a range
--- @param r1 integer Start row (or nil to clear)
--- @param c1? integer Start column
--- @param r2? integer End row
--- @param c2? integer End column
--- @return Worksheet self For chaining
function Worksheet:set_auto_filter(r1, c1, r2, c2)
  if r1 == nil then
    self.auto_filter = nil
    return self
  end

  -- Validate bounds
  if r1 < 1 or r1 > 1048576 or r2 < 1 or r2 > 1048576 then
    error("Row out of range (1-1048576)")
  end
  if c1 < 1 or c1 > 16384 or c2 < 1 or c2 > 16384 then
    error("Column out of range (1-16384)")
  end

  -- Normalize coordinates
  if r1 > r2 then r1, r2 = r2, r1 end
  if c1 > c2 then c1, c2 = c2, c1 end

  self.auto_filter = {
    ref = column_utils.make_range(r1, c1, r2, c2)
  }

  return self
end

--- Set auto-filter using A1:B2 notation
--- @param range string Range reference (e.g., "A1:D10")
--- @return Worksheet self For chaining
function Worksheet:set_auto_filter_range(range)
  if not range then
    self.auto_filter = nil
    return self
  end

  local parsed = column_utils.parse_range(range)
  return self:set_auto_filter(
    parsed.start.row, parsed.start.col,
    parsed.finish.row, parsed.finish.col
  )
end

--- Add data validation to a cell or range
--- @param ref string Cell or range reference (e.g., "A1" or "A1:A10")
--- @param validation table Validation options
--- @return Worksheet self For chaining
function Worksheet:add_data_validation(ref, validation)
  if not ref then
    error("Data validation requires a cell reference")
  end
  if not validation or not validation.type then
    error("Data validation requires a type")
  end

  local dv = {
    ref = ref,
    type = validation.type,
    operator = validation.operator,
    formula1 = validation.formula1,
    formula2 = validation.formula2,
    allowBlank = validation.allowBlank ~= false,  -- default true
    showDropDown = validation.showDropDown ~= false,  -- default true
    showInputMessage = validation.showInputMessage or false,
    showErrorMessage = validation.showErrorMessage ~= false,  -- default true
    promptTitle = validation.promptTitle,
    prompt = validation.prompt,
    errorTitle = validation.errorTitle,
    error = validation.error,
    errorStyle = validation.errorStyle or "stop",
  }

  table.insert(self.data_validations, dv)
  return self
end

--- Add a dropdown list validation to a cell or range
--- @param ref string Cell or range reference
--- @param items string[]|string List of items or a formula reference
--- @param options? table Optional: { allowBlank?, prompt?, promptTitle?, error?, errorTitle? }
--- @return Worksheet self For chaining
function Worksheet:add_dropdown(ref, items, options)
  options = options or {}

  local formula1
  if type(items) == "table" then
    -- Create comma-separated list in quotes
    formula1 = '"' .. table.concat(items, ",") .. '"'
  else
    -- It's a formula reference (e.g., "Sheet2!$A$1:$A$10")
    formula1 = items
  end

  return self:add_data_validation(ref, {
    type = "list",
    formula1 = formula1,
    allowBlank = options.allowBlank,
    showDropDown = true,
    showInputMessage = options.prompt ~= nil,
    prompt = options.prompt,
    promptTitle = options.promptTitle,
    showErrorMessage = options.error ~= nil or true,
    error = options.error,
    errorTitle = options.errorTitle,
    errorStyle = options.errorStyle,
  })
end

--- Add a number range validation
--- @param ref string Cell or range reference
--- @param min number|string Minimum value or formula
--- @param max number|string Maximum value or formula
--- @param options? table Optional: { allowBlank?, allowDecimal?, prompt?, error? }
--- @return Worksheet self For chaining
function Worksheet:add_number_validation(ref, min, max, options)
  options = options or {}

  return self:add_data_validation(ref, {
    type = options.allowDecimal and "decimal" or "whole",
    operator = "between",
    formula1 = tostring(min),
    formula2 = tostring(max),
    allowBlank = options.allowBlank,
    showInputMessage = options.prompt ~= nil,
    prompt = options.prompt,
    promptTitle = options.promptTitle,
    showErrorMessage = true,
    error = options.error or string.format("Value must be between %s and %s", min, max),
    errorTitle = options.errorTitle or "Invalid Input",
    errorStyle = options.errorStyle,
  })
end

--- Add a hyperlink to a cell
--- @param row integer Row number
--- @param col integer Column number
--- @param target string URL or internal reference
--- @param options? table Optional: { tooltip?, display_text? }
--- @return Worksheet self For chaining
function Worksheet:add_hyperlink(row, col, target, options)
  options = options or {}

  local ref = column_utils.make_ref(row, col)
  -- Check if it's an external link (URL, mailto, or file)
  local is_external = (target:match("^https?://") ~= nil)
    or (target:match("^mailto:") ~= nil)
    or (target:match("^file://") ~= nil)

  local link = {
    ref = ref,
    target = target,
    tooltip = options.tooltip,
    is_external = is_external,
  }

  -- For internal links, parse the location
  if not is_external then
    link.location = target
  end

  table.insert(self.hyperlinks, link)

  -- If display text is provided, set the cell value
  if options.display_text then
    self:set_cell(row, col, options.display_text)
  elseif not self:get_cell(row, col) then
    -- Default to showing the URL if no cell value exists
    self:set_cell(row, col, target)
  end

  return self
end

--- Set print settings for the worksheet
--- @param settings WorksheetPrintSettings Print settings
--- @return Worksheet self For chaining
function Worksheet:set_print_settings(settings)
  self.print_settings = settings
  return self
end

--- Set page margins (in inches)
--- @param top number Top margin
--- @param bottom number Bottom margin
--- @param left number Left margin
--- @param right number Right margin
--- @param header? number Header margin (default 0.3)
--- @param footer? number Footer margin (default 0.3)
--- @return Worksheet self For chaining
function Worksheet:set_margins(top, bottom, left, right, header, footer)
  if not self.print_settings then
    self.print_settings = {}
  end
  self.print_settings.margins = {
    top = top,
    bottom = bottom,
    left = left,
    right = right,
    header = header or 0.3,
    footer = footer or 0.3,
  }
  return self
end

--- Set page orientation
--- @param orientation string "portrait" or "landscape"
--- @return Worksheet self For chaining
function Worksheet:set_orientation(orientation)
  if orientation ~= "portrait" and orientation ~= "landscape" then
    error("Orientation must be 'portrait' or 'landscape'")
  end
  if not self.print_settings then
    self.print_settings = {}
  end
  self.print_settings.orientation = orientation
  return self
end

--- Set print area
--- @param range string Range to print (e.g., "A1:G20"), or nil to clear
--- @return Worksheet self For chaining
function Worksheet:set_print_area(range)
  if not self.print_settings then
    self.print_settings = {}
  end
  self.print_settings.printArea = range
  return self
end

--- Set print title rows (rows to repeat at top of each page)
--- @param rows string Row range (e.g., "1:2" for first two rows)
--- @return Worksheet self For chaining
function Worksheet:set_print_title_rows(rows)
  if not self.print_settings then
    self.print_settings = {}
  end
  self.print_settings.printTitleRows = rows
  return self
end

--- Set print title columns (columns to repeat at left of each page)
--- @param cols string Column range (e.g., "A:B" for first two columns)
--- @return Worksheet self For chaining
function Worksheet:set_print_title_cols(cols)
  if not self.print_settings then
    self.print_settings = {}
  end
  self.print_settings.printTitleCols = cols
  return self
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

--- Generate the sheetViews XML with freeze pane support
--- @param is_active boolean Whether this sheet is the active/selected sheet
--- @return string
function Worksheet:_generate_sheet_views(is_active)
  local tab_selected = is_active and "1" or "0"

  if not self.freeze_pane then
    -- Simple sheet view without freeze panes
    return '<sheetViews><sheetView tabSelected="' .. tab_selected .. '" workbookViewId="0"/></sheetViews>'
  end

  local fp = self.freeze_pane
  local parts = {}

  table.insert(parts, '<sheetViews>')
  table.insert(parts, '<sheetView tabSelected="' .. tab_selected .. '" workbookViewId="0">')

  -- Calculate the top-left cell of the unfrozen region
  local top_left_row = fp.rows + 1
  local top_left_col = fp.cols + 1
  local top_left_cell = column_utils.make_ref(top_left_row, top_left_col)

  -- Determine the active pane and pane state
  local active_pane
  local pane_state = "frozen"

  if fp.rows > 0 and fp.cols > 0 then
    -- Both rows and columns frozen
    active_pane = "bottomRight"
  elseif fp.rows > 0 then
    -- Only rows frozen
    active_pane = "bottomLeft"
  else
    -- Only columns frozen
    active_pane = "topRight"
  end

  -- Generate pane element
  local pane_attrs = {
    state = pane_state,
    topLeftCell = top_left_cell,
    activePane = active_pane,
  }

  if fp.cols > 0 then
    pane_attrs.xSplit = fp.cols
  end
  if fp.rows > 0 then
    pane_attrs.ySplit = fp.rows
  end

  table.insert(parts, xml.empty_element("pane", pane_attrs))

  -- Generate selection elements for the panes
  if fp.rows > 0 and fp.cols > 0 then
    -- Four panes: need selections for topRight, bottomLeft, bottomRight
    table.insert(parts, xml.empty_element("selection", { pane = "topRight", activeCell = column_utils.make_ref(1, top_left_col), sqref = column_utils.make_ref(1, top_left_col) }))
    table.insert(parts, xml.empty_element("selection", { pane = "bottomLeft", activeCell = column_utils.make_ref(top_left_row, 1), sqref = column_utils.make_ref(top_left_row, 1) }))
    table.insert(parts, xml.empty_element("selection", { pane = "bottomRight", activeCell = top_left_cell, sqref = top_left_cell }))
  elseif fp.rows > 0 then
    -- Two panes (top/bottom)
    table.insert(parts, xml.empty_element("selection", { pane = "bottomLeft", activeCell = top_left_cell, sqref = top_left_cell }))
  else
    -- Two panes (left/right)
    table.insert(parts, xml.empty_element("selection", { pane = "topRight", activeCell = top_left_cell, sqref = top_left_cell }))
  end

  table.insert(parts, '</sheetView>')
  table.insert(parts, '</sheetViews>')

  return table.concat(parts)
end

--- Generate auto-filter XML
--- @return string
function Worksheet:_generate_auto_filter()
  if not self.auto_filter then
    return ""
  end
  return xml.empty_element("autoFilter", { ref = self.auto_filter.ref })
end

--- Generate data validations XML
--- @return string
function Worksheet:_generate_data_validations()
  if #self.data_validations == 0 then
    return ""
  end

  local parts = {}

  for _, dv in ipairs(self.data_validations) do
    local attrs = {
      type = dv.type,
      sqref = dv.ref,
      allowBlank = dv.allowBlank and "1" or "0",
      showErrorMessage = dv.showErrorMessage and "1" or "0",
      showInputMessage = dv.showInputMessage and "1" or "0",
    }

    -- Note: Excel uses showDropDown="1" to HIDE the dropdown (counter-intuitive)
    -- We expose showDropDown=true to SHOW the dropdown, so we invert it
    if dv.type == "list" and dv.showDropDown == false then
      attrs.showDropDown = "1"  -- Hide dropdown
    end

    if dv.operator then
      attrs.operator = dv.operator
    end
    if dv.errorStyle and dv.errorStyle ~= "stop" then
      attrs.errorStyle = dv.errorStyle
    end
    if dv.errorTitle then
      attrs.errorTitle = dv.errorTitle
    end
    if dv.error then
      attrs.error = dv.error
    end
    if dv.promptTitle then
      attrs.promptTitle = dv.promptTitle
    end
    if dv.prompt then
      attrs.prompt = dv.prompt
    end

    -- Build inner content (formulas)
    local inner = {}
    if dv.formula1 then
      table.insert(inner, xml.element("formula1", dv.formula1))
    end
    if dv.formula2 then
      table.insert(inner, xml.element("formula2", dv.formula2))
    end

    if #inner > 0 then
      table.insert(parts, xml.element_raw("dataValidation", table.concat(inner), attrs))
    else
      table.insert(parts, xml.empty_element("dataValidation", attrs))
    end
  end

  return xml.element_raw("dataValidations", table.concat(parts), { count = #self.data_validations })
end

--- Generate hyperlinks XML
--- @param hyperlink_rels table Table to populate with relationship info for external links
--- @return string
function Worksheet:_generate_hyperlinks(hyperlink_rels)
  if #self.hyperlinks == 0 then
    return ""
  end

  local parts = {}
  local rel_id = 1

  for _, link in ipairs(self.hyperlinks) do
    local attrs = { ref = link.ref }

    if link.is_external then
      -- External link needs a relationship
      local rid = "rId" .. (1000 + rel_id)  -- Use high IDs to avoid conflicts
      attrs["r:id"] = rid
      if link.tooltip then
        attrs.tooltip = link.tooltip
      end
      table.insert(hyperlink_rels, {
        id = rid,
        target = link.target,
        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        targetMode = "External"
      })
      rel_id = rel_id + 1
    else
      -- Internal link uses location attribute
      attrs.location = link.location
      if link.tooltip then
        attrs.tooltip = link.tooltip
      end
    end

    table.insert(parts, xml.empty_element("hyperlink", attrs))
  end

  return xml.element_raw("hyperlinks", table.concat(parts))
end

--- Generate print settings XML (pageMargins and pageSetup)
--- @return string
function Worksheet:_generate_print_settings()
  if not self.print_settings then
    return ""
  end

  local parts = {}
  local ps = self.print_settings

  -- Page margins
  if ps.margins then
    local m = ps.margins
    table.insert(parts, xml.empty_element("pageMargins", {
      left = m.left or 0.7,
      right = m.right or 0.7,
      top = m.top or 0.75,
      bottom = m.bottom or 0.75,
      header = m.header or 0.3,
      footer = m.footer or 0.3,
    }))
  end

  -- Page setup
  local setup_attrs = {}
  local has_setup = false

  if ps.orientation then
    setup_attrs.orientation = ps.orientation
    has_setup = true
  end
  if ps.paperSize then
    setup_attrs.paperSize = ps.paperSize
    has_setup = true
  end
  if ps.scale then
    setup_attrs.scale = ps.scale
    has_setup = true
  end
  if ps.fitToWidth then
    setup_attrs.fitToWidth = ps.fitToWidth
    has_setup = true
  end
  if ps.fitToHeight then
    setup_attrs.fitToHeight = ps.fitToHeight
    has_setup = true
  end
  if ps.gridLines then
    setup_attrs.gridLines = "1"
    has_setup = true
  end
  if ps.headings then
    setup_attrs.headings = "1"
    has_setup = true
  end

  if has_setup then
    table.insert(parts, xml.empty_element("pageSetup", setup_attrs))
  end

  return table.concat(parts)
end

--- Get hyperlink relationships for this worksheet
--- @return table[] Array of relationship info for external hyperlinks
function Worksheet:get_hyperlink_relationships()
  local rels = {}
  self:_generate_hyperlinks(rels)
  return rels
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

  -- Sheet views (with freeze pane support)
  b:raw(self:_generate_sheet_views(is_active or false))

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

  -- Auto-filter (must come after sheetData)
  local auto_filter = self:_generate_auto_filter()
  if auto_filter ~= "" then
    b:raw(auto_filter)
  end

  -- Merged cells (if any)
  if #self.merged_cells > 0 then
    local merge_parts = {}
    for _, merge in ipairs(self.merged_cells) do
      table.insert(merge_parts, xml.empty_element("mergeCell", { ref = merge }))
    end
    b:elem_raw("mergeCells", table.concat(merge_parts), { count = #self.merged_cells })
  end

  -- Data validations
  local data_validations = self:_generate_data_validations()
  if data_validations ~= "" then
    b:raw(data_validations)
  end

  -- Hyperlinks
  local hyperlink_rels = {}
  local hyperlinks = self:_generate_hyperlinks(hyperlink_rels)
  if hyperlinks ~= "" then
    b:raw(hyperlinks)
  end

  -- Print settings (pageMargins and pageSetup)
  local print_settings = self:_generate_print_settings()
  if print_settings ~= "" then
    b:raw(print_settings)
  end

  b:close("worksheet")

  return b:to_string()
end

M.Worksheet = Worksheet

return M
