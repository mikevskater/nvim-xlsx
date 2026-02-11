--- Validation utilities for xlsx
--- @module xlsx.utils.validation

local M = {}

-- Excel limits
M.LIMITS = {
  MAX_ROWS = 1048576,
  MAX_COLS = 16384,
  MAX_CELL_TEXT = 32767,
  MAX_FORMULA_LENGTH = 8192,
  MAX_SHEET_NAME = 31,
  MAX_URL_LENGTH = 2083,  -- IE limit, commonly used
}

-- Forbidden characters in sheet names
M.SHEET_NAME_FORBIDDEN = { "\\", "/", "*", "?", ":", "[", "]" }

---@class ValidationResult
---@field valid boolean Whether validation passed
---@field error? string Error message if validation failed

--- Validate a row number
--- @param row any Value to validate
--- @return ValidationResult
function M.validate_row(row)
  if type(row) ~= "number" then
    return { valid = false, error = "Row must be a number, got " .. type(row) }
  end
  if row ~= math.floor(row) then
    return { valid = false, error = "Row must be an integer, got " .. tostring(row) }
  end
  if row < 1 then
    return { valid = false, error = "Row must be >= 1, got " .. tostring(row) }
  end
  if row > M.LIMITS.MAX_ROWS then
    return { valid = false, error = string.format("Row %d exceeds Excel limit of %d", row, M.LIMITS.MAX_ROWS) }
  end
  return { valid = true }
end

--- Validate a column number
--- @param col any Value to validate
--- @return ValidationResult
function M.validate_col(col)
  if type(col) ~= "number" then
    return { valid = false, error = "Column must be a number, got " .. type(col) }
  end
  if col ~= math.floor(col) then
    return { valid = false, error = "Column must be an integer, got " .. tostring(col) }
  end
  if col < 1 then
    return { valid = false, error = "Column must be >= 1, got " .. tostring(col) }
  end
  if col > M.LIMITS.MAX_COLS then
    return { valid = false, error = string.format("Column %d exceeds Excel limit of %d", col, M.LIMITS.MAX_COLS) }
  end
  return { valid = true }
end

--- Validate a cell reference (row and column)
--- @param row any Row number
--- @param col any Column number
--- @return ValidationResult
function M.validate_cell_ref(row, col)
  local row_result = M.validate_row(row)
  if not row_result.valid then
    return row_result
  end
  local col_result = M.validate_col(col)
  if not col_result.valid then
    return col_result
  end
  return { valid = true }
end

--- Validate a range
--- @param r1 any Start row
--- @param c1 any Start column
--- @param r2 any End row
--- @param c2 any End column
--- @return ValidationResult
function M.validate_range(r1, c1, r2, c2)
  local start_result = M.validate_cell_ref(r1, c1)
  if not start_result.valid then
    return { valid = false, error = "Invalid range start: " .. start_result.error }
  end
  local end_result = M.validate_cell_ref(r2, c2)
  if not end_result.valid then
    return { valid = false, error = "Invalid range end: " .. end_result.error }
  end
  return { valid = true }
end

--- Validate a sheet name
--- @param name any Value to validate
--- @return ValidationResult
function M.validate_sheet_name(name)
  if type(name) ~= "string" then
    return { valid = false, error = "Sheet name must be a string, got " .. type(name) }
  end
  if name == "" then
    return { valid = false, error = "Sheet name cannot be empty" }
  end
  if #name > M.LIMITS.MAX_SHEET_NAME then
    return { valid = false, error = string.format("Sheet name '%s' exceeds %d character limit", name, M.LIMITS.MAX_SHEET_NAME) }
  end
  -- Check for forbidden characters
  for _, char in ipairs(M.SHEET_NAME_FORBIDDEN) do
    if name:find(char, 1, true) then
      return { valid = false, error = string.format("Sheet name cannot contain '%s'", char) }
    end
  end
  -- Cannot start or end with apostrophe
  if name:sub(1, 1) == "'" then
    return { valid = false, error = "Sheet name cannot start with apostrophe" }
  end
  if name:sub(-1) == "'" then
    return { valid = false, error = "Sheet name cannot end with apostrophe" }
  end
  -- Cannot be 'History' (reserved by Excel)
  if name:lower() == "history" then
    return { valid = false, error = "Sheet name 'History' is reserved by Excel" }
  end
  return { valid = true }
end

--- Validate cell text length
--- @param text any Value to validate
--- @return ValidationResult
function M.validate_cell_text(text)
  if text == nil then
    return { valid = true }
  end
  local str = tostring(text)
  if #str > M.LIMITS.MAX_CELL_TEXT then
    return { valid = false, error = string.format("Cell text length %d exceeds Excel limit of %d", #str, M.LIMITS.MAX_CELL_TEXT) }
  end
  return { valid = true }
end

--- Validate formula length
--- @param formula any Value to validate
--- @return ValidationResult
function M.validate_formula(formula)
  if type(formula) ~= "string" then
    return { valid = false, error = "Formula must be a string, got " .. type(formula) }
  end
  -- Remove leading = if present for length check
  local f = formula:match("^=(.*)$") or formula
  if #f > M.LIMITS.MAX_FORMULA_LENGTH then
    return { valid = false, error = string.format("Formula length %d exceeds Excel limit of %d", #f, M.LIMITS.MAX_FORMULA_LENGTH) }
  end
  return { valid = true }
end

--- Validate a URL for hyperlinks
--- @param url any Value to validate
--- @return ValidationResult
function M.validate_url(url)
  if type(url) ~= "string" then
    return { valid = false, error = "URL must be a string, got " .. type(url) }
  end
  if url == "" then
    return { valid = false, error = "URL cannot be empty" }
  end
  if #url > M.LIMITS.MAX_URL_LENGTH then
    return { valid = false, error = string.format("URL length %d exceeds limit of %d", #url, M.LIMITS.MAX_URL_LENGTH) }
  end
  return { valid = true }
end

--- Validate a number format code
--- @param format any Value to validate
--- @return ValidationResult
function M.validate_number_format(format)
  if type(format) ~= "string" then
    return { valid = false, error = "Number format must be a string, got " .. type(format) }
  end
  -- Basic check - format shouldn't be empty
  if format == "" then
    return { valid = false, error = "Number format cannot be empty" }
  end
  return { valid = true }
end

--- Validate a color value (hex format)
--- @param color any Value to validate
--- @return ValidationResult
function M.validate_color(color)
  if type(color) ~= "string" then
    return { valid = false, error = "Color must be a string, got " .. type(color) }
  end
  -- Check for valid hex color formats: #RGB, #RRGGBB, #AARRGGBB, or RRGGBB, AARRGGBB
  local clean = color:gsub("^#", "")
  if not clean:match("^%x+$") then
    return { valid = false, error = "Color must be a valid hex color (e.g., #FF0000 or FF0000)" }
  end
  local len = #clean
  if len ~= 3 and len ~= 6 and len ~= 8 then
    return { valid = false, error = "Color must be 3, 6, or 8 hex digits" }
  end
  return { valid = true }
end

--- Validate column width
--- @param width any Value to validate
--- @return ValidationResult
function M.validate_column_width(width)
  if type(width) ~= "number" then
    return { valid = false, error = "Column width must be a number, got " .. type(width) }
  end
  if width < 0 then
    return { valid = false, error = "Column width cannot be negative" }
  end
  if width > 255 then
    return { valid = false, error = "Column width cannot exceed 255" }
  end
  return { valid = true }
end

--- Validate row height
--- @param height any Value to validate
--- @return ValidationResult
function M.validate_row_height(height)
  if type(height) ~= "number" then
    return { valid = false, error = "Row height must be a number, got " .. type(height) }
  end
  if height < 0 then
    return { valid = false, error = "Row height cannot be negative" }
  end
  if height > 409 then
    return { valid = false, error = "Row height cannot exceed 409 points" }
  end
  return { valid = true }
end

--- Validate an Excel table name
--- @param name any Value to validate
--- @return ValidationResult
function M.validate_table_name(name)
  if type(name) ~= "string" then
    return { valid = false, error = "Table name must be a string, got " .. type(name) }
  end
  if name == "" then
    return { valid = false, error = "Table name cannot be empty" }
  end
  if #name > 255 then
    return { valid = false, error = string.format("Table name '%s' exceeds 255 character limit", name) }
  end
  -- Must start with letter, underscore, or backslash
  if not name:match("^[A-Za-z_\\]") then
    return { valid = false, error = "Table name must start with a letter, underscore, or backslash" }
  end
  -- Cannot contain spaces
  if name:find(" ") then
    return { valid = false, error = "Table name cannot contain spaces" }
  end
  -- Cannot look like a cell reference (e.g., A1, XFD1048576)
  if name:match("^[A-Za-z]+%d+$") then
    local letters = name:match("^([A-Za-z]+)")
    if #letters <= 3 then
      return { valid = false, error = "Table name cannot look like a cell reference: " .. name }
    end
  end
  return { valid = true }
end

--- Validate a defined name (named range name)
--- @param name any Value to validate
--- @return ValidationResult
function M.validate_defined_name(name)
  if type(name) ~= "string" then
    return { valid = false, error = "Defined name must be a string, got " .. type(name) }
  end
  if name == "" then
    return { valid = false, error = "Defined name cannot be empty" }
  end
  if #name > 255 then
    return { valid = false, error = string.format("Defined name '%s' exceeds 255 character limit", name) }
  end
  -- Must start with letter, underscore, or backslash
  if not name:match("^[A-Za-z_\\]") then
    return { valid = false, error = "Defined name must start with a letter, underscore, or backslash" }
  end
  -- Cannot contain spaces
  if name:find(" ") then
    return { valid = false, error = "Defined name cannot contain spaces" }
  end
  return { valid = true }
end

--- Helper to check validation and raise error if invalid
--- @param result ValidationResult
--- @param context? string Optional context for error message
function M.check(result, context)
  if not result.valid then
    local msg = result.error
    if context then
      msg = context .. ": " .. msg
    end
    error(msg, 3)
  end
end

--- Helper to check validation and return nil, error if invalid
--- @param result ValidationResult
--- @param context? string Optional context for error message
--- @return boolean valid
--- @return string? error
function M.check_soft(result, context)
  if not result.valid then
    local msg = result.error
    if context then
      msg = context .. ": " .. msg
    end
    return false, msg
  end
  return true
end

return M
