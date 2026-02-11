--- Worksheet features module - freeze panes, filters, validation, hyperlinks, print settings
--- @module nvim-xlsx.worksheet.features

local column_utils = require("nvim-xlsx.utils.column")
local validation = require("nvim-xlsx.utils.validation")

local LIMITS = validation.LIMITS

local M = {}

-- ============================================
-- Freeze Panes
-- ============================================

--- Freeze panes at the specified row and/or column
--- @param self Worksheet
--- @param rows integer Number of rows to freeze at top (0 for none)
--- @param cols integer Number of columns to freeze at left (0 for none)
--- @return Worksheet self For chaining
function M.freeze_panes(self, rows, cols)
  rows = rows or 0
  cols = cols or 0

  if rows < 0 or rows > LIMITS.MAX_ROWS then
    error("Freeze rows out of range (0-" .. LIMITS.MAX_ROWS .. "): " .. tostring(rows))
  end
  if cols < 0 or cols > LIMITS.MAX_COLS then
    error("Freeze cols out of range (0-" .. LIMITS.MAX_COLS .. "): " .. tostring(cols))
  end

  if rows == 0 and cols == 0 then
    self.freeze_pane = nil
  else
    self.freeze_pane = { rows = rows, cols = cols }
  end

  return self
end

--- Freeze the first N rows
--- @param self Worksheet
--- @param rows integer Number of rows to freeze
--- @return Worksheet self For chaining
function M.freeze_rows(self, rows)
  return M.freeze_panes(self, rows, self.freeze_pane and self.freeze_pane.cols or 0)
end

--- Freeze the first N columns
--- @param self Worksheet
--- @param cols integer Number of columns to freeze
--- @return Worksheet self For chaining
function M.freeze_cols(self, cols)
  return M.freeze_panes(self, self.freeze_pane and self.freeze_pane.rows or 0, cols)
end

-- ============================================
-- Auto-Filter
-- ============================================

--- Set auto-filter for a range
--- @param self Worksheet
--- @param r1 integer Start row (or nil to clear)
--- @param c1? integer Start column
--- @param r2? integer End row
--- @param c2? integer End column
--- @return Worksheet self For chaining
function M.set_auto_filter(self, r1, c1, r2, c2)
  if r1 == nil then
    self.auto_filter = nil
    return self
  end

  -- Validate bounds
  validation.check(validation.validate_range(r1, c1, r2, c2))

  -- Normalize coordinates
  if r1 > r2 then r1, r2 = r2, r1 end
  if c1 > c2 then c1, c2 = c2, c1 end

  self.auto_filter = {
    ref = column_utils.make_range(r1, c1, r2, c2)
  }

  return self
end

--- Set auto-filter using A1:B2 notation
--- @param self Worksheet
--- @param range string Range reference (e.g., "A1:D10")
--- @return Worksheet self For chaining
function M.set_auto_filter_range(self, range)
  if not range then
    self.auto_filter = nil
    return self
  end

  local parsed = column_utils.parse_range(range)
  return M.set_auto_filter(self,
    parsed.start.row, parsed.start.col,
    parsed.finish.row, parsed.finish.col
  )
end

-- ============================================
-- Data Validation
-- ============================================

--- Add data validation to a cell or range
--- @param self Worksheet
--- @param ref string Cell or range reference (e.g., "A1" or "A1:A10")
--- @param validation_opts table Validation options
--- @return Worksheet self For chaining
function M.add_data_validation(self, ref, validation_opts)
  if not ref then
    error("Data validation requires a cell reference")
  end
  if not validation_opts or not validation_opts.type then
    error("Data validation requires a type")
  end

  local dv = {
    ref = ref,
    type = validation_opts.type,
    operator = validation_opts.operator,
    formula1 = validation_opts.formula1,
    formula2 = validation_opts.formula2,
    allowBlank = validation_opts.allowBlank ~= false,  -- default true
    showDropDown = validation_opts.showDropDown ~= false,  -- default true
    showInputMessage = validation_opts.showInputMessage or false,
    showErrorMessage = validation_opts.showErrorMessage ~= false,  -- default true
    promptTitle = validation_opts.promptTitle,
    prompt = validation_opts.prompt,
    errorTitle = validation_opts.errorTitle,
    error = validation_opts.error,
    errorStyle = validation_opts.errorStyle or "stop",
  }

  table.insert(self.data_validations, dv)
  return self
end

--- Add a dropdown list validation to a cell or range
--- @param self Worksheet
--- @param ref string Cell or range reference
--- @param items string[]|string List of items or a formula reference
--- @param options? table Optional: { allowBlank?, prompt?, promptTitle?, error?, errorTitle? }
--- @return Worksheet self For chaining
function M.add_dropdown(self, ref, items, options)
  options = options or {}

  local formula1
  if type(items) == "table" then
    -- Create comma-separated list in quotes
    formula1 = '"' .. table.concat(items, ",") .. '"'
  else
    -- It's a formula reference (e.g., "Sheet2!$A$1:$A$10")
    formula1 = items
  end

  return M.add_data_validation(self, ref, {
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
--- @param self Worksheet
--- @param ref string Cell or range reference
--- @param min number|string Minimum value or formula
--- @param max number|string Maximum value or formula
--- @param options? table Optional: { allowBlank?, allowDecimal?, prompt?, error? }
--- @return Worksheet self For chaining
function M.add_number_validation(self, ref, min, max, options)
  options = options or {}

  return M.add_data_validation(self, ref, {
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

-- ============================================
-- Hyperlinks
-- ============================================

--- Add a hyperlink to a cell
--- @param self Worksheet
--- @param row integer Row number
--- @param col integer Column number
--- @param target string URL or internal reference
--- @param options? table Optional: { tooltip?, display_text? }
--- @return Worksheet self For chaining
function M.add_hyperlink(self, row, col, target, options)
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

  -- Store display text for the hyperlink element
  if options.display_text then
    link.display = options.display_text
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

-- ============================================
-- Print Settings
-- ============================================

--- Set print settings for the worksheet
--- @param self Worksheet
--- @param settings WorksheetPrintSettings Print settings
--- @return Worksheet self For chaining
function M.set_print_settings(self, settings)
  self.print_settings = settings
  return self
end

--- Ensure print_settings table exists
--- @param self Worksheet
--- @return WorksheetPrintSettings
function M._ensure_print_settings(self)
  if not self.print_settings then
    self.print_settings = {}
  end
  return self.print_settings
end

--- Set page margins (in inches)
--- @param self Worksheet
--- @param top number Top margin
--- @param bottom number Bottom margin
--- @param left number Left margin
--- @param right number Right margin
--- @param header? number Header margin (default 0.3)
--- @param footer? number Footer margin (default 0.3)
--- @return Worksheet self For chaining
function M.set_margins(self, top, bottom, left, right, header, footer)
  M._ensure_print_settings(self).margins = {
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
--- @param self Worksheet
--- @param orientation string "portrait" or "landscape"
--- @return Worksheet self For chaining
function M.set_orientation(self, orientation)
  if orientation ~= "portrait" and orientation ~= "landscape" then
    error("Orientation must be 'portrait' or 'landscape'")
  end
  M._ensure_print_settings(self).orientation = orientation
  return self
end

--- Set print area
--- @param self Worksheet
--- @param range string Range to print (e.g., "A1:G20"), or nil to clear
--- @return Worksheet self For chaining
function M.set_print_area(self, range)
  M._ensure_print_settings(self).printArea = range
  return self
end

--- Set print title rows (rows to repeat at top of each page)
--- @param self Worksheet
--- @param rows string Row range (e.g., "1:2" for first two rows)
--- @return Worksheet self For chaining
function M.set_print_title_rows(self, rows)
  M._ensure_print_settings(self).printTitleRows = rows
  return self
end

--- Set print title columns (columns to repeat at left of each page)
--- @param self Worksheet
--- @param cols string Column range (e.g., "A:B" for first two columns)
--- @return Worksheet self For chaining
function M.set_print_title_cols(self, cols)
  M._ensure_print_settings(self).printTitleCols = cols
  return self
end

-- ============================================
-- Auto-fit Column Widths
-- ============================================

--- Auto-fit column widths based on cell content
--- Measures the display width of all cell values and sets column widths accordingly.
--- @param self Worksheet
--- @param opts? table Options: { min_width?: integer, max_width?: integer, padding?: integer, columns?: (integer|string)[] }
--- @return Worksheet self For chaining
function M.auto_fit_columns(self, opts)
  opts = opts or {}
  local min_width = opts.min_width or 8
  local max_width = opts.max_width or 50
  local padding = opts.padding or 2

  if not self.min_col or not self.max_col then
    return self
  end

  -- Build set of target columns (resolve string names via header row mapping)
  local target_cols
  if opts.columns then
    target_cols = {}
    for _, col in ipairs(opts.columns) do
      if type(col) == "string" then
        local resolved = self.column_names and self.column_names[col]
        if resolved then
          target_cols[resolved] = true
        end
      else
        target_cols[col] = true
      end
    end
  end

  -- Measure max content width per column
  local col_widths = {}
  if target_cols then
    for col in pairs(target_cols) do
      col_widths[col] = 0
    end
  else
    for col = self.min_col, self.max_col do
      col_widths[col] = 0
    end
  end

  for _, row_data in pairs(self.rows) do
    for col, cell in pairs(row_data) do
      if col_widths[col] ~= nil and cell and cell.value ~= nil then
        local len = #tostring(cell.value)
        if len > col_widths[col] then
          col_widths[col] = len
        end
      end
    end
  end

  -- Apply constraints and store
  for col, width in pairs(col_widths) do
    width = math.max(min_width, math.min(max_width, width + padding))
    self.column_widths[col] = width
  end

  return self
end

-- ============================================
-- Excel Tables
-- ============================================

--- Add an Excel table (structured ListObject) to the worksheet
--- @param self Worksheet
--- @param r1 integer Start row (header row)
--- @param c1 integer Start column
--- @param r2 integer End row (must be > r1 for at least one data row)
--- @param c2 integer End column
--- @param options? table Optional: { name?: string, auto_filter?: boolean, style_name?: string, show_first_col?: boolean, show_last_col?: boolean, show_row_stripes?: boolean, show_col_stripes?: boolean }
--- @return Worksheet self For chaining
function M.add_table(self, r1, c1, r2, c2, options)
  options = options or {}

  -- Validate bounds
  validation.check(validation.validate_range(r1, c1, r2, c2))

  -- Normalize coordinates
  if r1 > r2 then r1, r2 = r2, r1 end
  if c1 > c2 then c1, c2 = c2, c1 end

  -- Must have at least a header row and one data row
  if r1 == r2 then
    error("Table must have at least a header row and one data row (r1 == r2)")
  end

  -- Get globally unique table ID
  local id = self.workbook:_next_table_id()

  -- Determine table name
  local name = options.name or ("Table" .. id)
  validation.check(validation.validate_table_name(name), "add_table")

  -- Derive column names from header row cells
  local columns = {}
  for col = c1, c2 do
    local col_id = col - c1 + 1
    local col_name
    local cell = self:get_cell(r1, col)
    if cell and cell.value ~= nil and tostring(cell.value) ~= "" then
      col_name = tostring(cell.value)
    else
      col_name = "Column" .. col_id
    end
    table.insert(columns, { id = col_id, name = col_name })
  end

  -- Build the ExcelTable struct
  local ref = column_utils.make_range(r1, c1, r2, c2)
  ---@type ExcelTable
  local tbl = {
    id = id,
    name = name,
    ref = ref,
    header_row = true,
    r1 = r1,
    c1 = c1,
    r2 = r2,
    c2 = c2,
    columns = columns,
    auto_filter = options.auto_filter ~= false, -- default true
    style_name = options.style_name or "TableStyleMedium2",
    show_first_col = options.show_first_col or false,
    show_last_col = options.show_last_col or false,
    show_row_stripes = options.show_row_stripes ~= false, -- default true
    show_col_stripes = options.show_col_stripes or false,
  }

  table.insert(self.tables, tbl)
  return self
end

return M
