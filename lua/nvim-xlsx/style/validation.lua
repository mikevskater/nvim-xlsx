--- Style validation for xlsx
--- Validates style definitions before applying
--- @module nvim-xlsx.style.validation

local constants = require("nvim-xlsx.style.constants")
local color = require("nvim-xlsx.utils.color")

local M = {}

-- Build lookup tables for valid values
local VALID_HALIGN = {}
for k in pairs(constants.HALIGN) do
  VALID_HALIGN[k] = true
end

local VALID_VALIGN = {}
for k in pairs(constants.VALIGN) do
  VALID_VALIGN[k] = true
end

local VALID_BORDER_STYLES = { none = true }
for k in pairs(constants.BORDER_STYLES) do
  if k ~= "none" then
    VALID_BORDER_STYLES[k] = true
  end
end

local VALID_UNDERLINE = { none = true }
for k in pairs(constants.UNDERLINE) do
  if k ~= "none" then
    VALID_UNDERLINE[k] = true
  end
end

local VALID_BUILTIN_FORMATS = {}
for k in pairs(constants.BUILTIN_FORMATS) do
  VALID_BUILTIN_FORMATS[k] = true
end

-- Fill patterns supported by Excel
local VALID_PATTERNS = {
  none = true,
  solid = true,
  gray125 = true,
  gray0625 = true,
  darkGray = true,
  mediumGray = true,
  lightGray = true,
  darkHorizontal = true,
  darkVertical = true,
  darkDown = true,
  darkUp = true,
  darkGrid = true,
  darkTrellis = true,
  lightHorizontal = true,
  lightVertical = true,
  lightDown = true,
  lightUp = true,
  lightGrid = true,
  lightTrellis = true,
}

--- Format a list of valid options for error messages
--- @param tbl table Lookup table
--- @return string Comma-separated list
local function format_options(tbl)
  local opts = {}
  for k in pairs(tbl) do
    table.insert(opts, '"' .. k .. '"')
  end
  table.sort(opts)
  return table.concat(opts, ", ")
end

--- Validate a color value
--- @param value any The color value to validate
--- @param field string Field name for error message
--- @return string? error Error message if invalid
local function validate_color(value, field)
  if value == nil then
    return nil
  end
  if type(value) ~= "string" then
    return string.format("%s must be a string, got %s", field, type(value))
  end
  -- Check if it converts successfully
  local argb = color.to_argb(value)
  if not argb then
    return string.format(
      "Invalid %s value '%s'. Expected hex color (#RRGGBB, #AARRGGBB, RRGGBB) or named color (%s)",
      field, value, "black, white, red, green, blue, yellow, gray, etc."
    )
  end
  return nil
end

--- Validate a boolean value
--- @param value any The value to validate
--- @param field string Field name for error message
--- @return string? error Error message if invalid
local function validate_boolean(value, field)
  if value == nil then
    return nil
  end
  if type(value) ~= "boolean" then
    return string.format("%s must be a boolean, got %s", field, type(value))
  end
  return nil
end

--- Validate a positive number
--- @param value any The value to validate
--- @param field string Field name for error message
--- @param min? number Minimum value (inclusive)
--- @param max? number Maximum value (inclusive)
--- @return string? error Error message if invalid
local function validate_number(value, field, min, max)
  if value == nil then
    return nil
  end
  if type(value) ~= "number" then
    return string.format("%s must be a number, got %s", field, type(value))
  end
  if min and value < min then
    return string.format("%s must be >= %s, got %s", field, min, value)
  end
  if max and value > max then
    return string.format("%s must be <= %s, got %s", field, max, value)
  end
  return nil
end

--- Validate a non-negative integer
--- @param value any The value to validate
--- @param field string Field name for error message
--- @return string? error Error message if invalid
local function validate_non_negative_integer(value, field)
  if value == nil then
    return nil
  end
  if type(value) ~= "number" then
    return string.format("%s must be a number, got %s", field, type(value))
  end
  if value < 0 then
    return string.format("%s must be >= 0, got %s", field, value)
  end
  if value ~= math.floor(value) then
    return string.format("%s must be an integer, got %s", field, value)
  end
  return nil
end

--- Validate an enum value against a lookup table
--- @param value any The value to validate
--- @param field string Field name for error message
--- @param valid_values table Lookup table of valid values
--- @return string? error Error message if invalid
local function validate_enum(value, field, valid_values)
  if value == nil then
    return nil
  end
  if type(value) ~= "string" then
    return string.format("%s must be a string, got %s", field, type(value))
  end
  if not valid_values[value] then
    return string.format(
      "Invalid %s value '%s'. Valid options: %s",
      field, value, format_options(valid_values)
    )
  end
  return nil
end

--- Validate a single border edge definition
--- @param value any The edge value (string, table, or boolean)
--- @param edge string Edge name (left, right, top, bottom)
--- @return string? error Error message if invalid
local function validate_border_edge(value, edge)
  if value == nil then
    return nil
  end

  local field = "border_" .. edge

  if type(value) == "boolean" then
    return nil  -- boolean is always valid
  end

  if type(value) == "string" then
    return validate_enum(value, field, VALID_BORDER_STYLES)
  end

  if type(value) == "table" then
    local errors = {}
    if value.style ~= nil then
      local err = validate_enum(value.style, field .. ".style", VALID_BORDER_STYLES)
      if err then table.insert(errors, err) end
    end
    if value.color ~= nil then
      local err = validate_color(value.color, field .. ".color")
      if err then table.insert(errors, err) end
    end
    if #errors > 0 then
      return table.concat(errors, "; ")
    end
    return nil
  end

  return string.format("%s must be a string, boolean, or table, got %s", field, type(value))
end

--- Validate a complete style definition
--- Returns all validation errors found
--- @param def table Style definition to validate
--- @return boolean valid Whether the definition is valid
--- @return string[]? errors Array of error messages (nil if valid)
function M.validate_style(def)
  if def == nil then
    return true, nil
  end

  if type(def) ~= "table" then
    return false, { "Style definition must be a table, got " .. type(def) }
  end

  local errors = {}
  local err

  -- Font properties
  if def.font_name ~= nil and type(def.font_name) ~= "string" then
    table.insert(errors, "font_name must be a string, got " .. type(def.font_name))
  end
  if def.font ~= nil and type(def.font) ~= "string" then
    table.insert(errors, "font must be a string, got " .. type(def.font))
  end

  err = validate_number(def.font_size, "font_size", 1, 409)
  if err then table.insert(errors, err) end

  err = validate_boolean(def.bold, "bold")
  if err then table.insert(errors, err) end

  err = validate_boolean(def.italic, "italic")
  if err then table.insert(errors, err) end

  err = validate_boolean(def.strike, "strike")
  if err then table.insert(errors, err) end

  err = validate_enum(def.underline, "underline", VALID_UNDERLINE)
  if err then table.insert(errors, err) end

  err = validate_color(def.font_color, "font_color")
  if err then table.insert(errors, err) end

  -- Fill properties
  err = validate_color(def.bg_color, "bg_color")
  if err then table.insert(errors, err) end

  err = validate_color(def.fill_color, "fill_color")
  if err then table.insert(errors, err) end

  err = validate_enum(def.pattern, "pattern", VALID_PATTERNS)
  if err then table.insert(errors, err) end

  -- Border properties
  err = validate_boolean(def.border, "border")
  if err then table.insert(errors, err) end

  err = validate_enum(def.border_style, "border_style", VALID_BORDER_STYLES)
  if err then table.insert(errors, err) end

  err = validate_color(def.border_color, "border_color")
  if err then table.insert(errors, err) end

  -- Individual border edges
  for _, edge in ipairs({ "left", "right", "top", "bottom" }) do
    err = validate_border_edge(def["border_" .. edge], edge)
    if err then table.insert(errors, err) end
  end

  -- Alignment properties
  err = validate_enum(def.halign, "halign", VALID_HALIGN)
  if err then table.insert(errors, err) end

  err = validate_enum(def.align, "align", VALID_HALIGN)
  if err then table.insert(errors, err) end

  err = validate_enum(def.valign, "valign", VALID_VALIGN)
  if err then table.insert(errors, err) end

  err = validate_boolean(def.wrap_text, "wrap_text")
  if err then table.insert(errors, err) end

  err = validate_boolean(def.wrap, "wrap")
  if err then table.insert(errors, err) end

  -- Rotation: -90 to 90, or 255 for vertical text
  if def.rotation ~= nil then
    if type(def.rotation) ~= "number" then
      table.insert(errors, "rotation must be a number, got " .. type(def.rotation))
    elseif def.rotation ~= 255 and (def.rotation < -90 or def.rotation > 90) then
      table.insert(errors, "rotation must be between -90 and 90, or 255 for vertical text, got " .. def.rotation)
    end
  end

  err = validate_non_negative_integer(def.indent, "indent")
  if err then table.insert(errors, err) end

  -- Number format
  if def.num_format ~= nil then
    if type(def.num_format) ~= "string" and type(def.num_format) ~= "number" then
      table.insert(errors, "num_format must be a string or number, got " .. type(def.num_format))
    end
  end
  if def.number_format ~= nil then
    if type(def.number_format) ~= "string" and type(def.number_format) ~= "number" then
      table.insert(errors, "number_format must be a string or number, got " .. type(def.number_format))
    end
  end

  if #errors > 0 then
    return false, errors
  end

  return true, nil
end

--- Get list of all valid horizontal alignment values
--- @return string[]
function M.get_valid_halign()
  local result = {}
  for k in pairs(VALID_HALIGN) do
    table.insert(result, k)
  end
  table.sort(result)
  return result
end

--- Get list of all valid vertical alignment values
--- @return string[]
function M.get_valid_valign()
  local result = {}
  for k in pairs(VALID_VALIGN) do
    table.insert(result, k)
  end
  table.sort(result)
  return result
end

--- Get list of all valid border styles
--- @return string[]
function M.get_valid_border_styles()
  local result = {}
  for k in pairs(VALID_BORDER_STYLES) do
    table.insert(result, k)
  end
  table.sort(result)
  return result
end

--- Get list of all valid underline styles
--- @return string[]
function M.get_valid_underline()
  local result = {}
  for k in pairs(VALID_UNDERLINE) do
    table.insert(result, k)
  end
  table.sort(result)
  return result
end

--- Get list of all valid pattern types
--- @return string[]
function M.get_valid_patterns()
  local result = {}
  for k in pairs(VALID_PATTERNS) do
    table.insert(result, k)
  end
  table.sort(result)
  return result
end

--- Get list of all built-in number format names
--- @return string[]
function M.get_valid_number_formats()
  local result = {}
  for k in pairs(VALID_BUILTIN_FORMATS) do
    table.insert(result, k)
  end
  table.sort(result)
  return result
end

--- Get list of all named colors
--- @return string[]
function M.get_valid_colors()
  local result = {}
  for k in pairs(color.COLORS) do
    table.insert(result, k)
  end
  table.sort(result)
  return result
end

return M
