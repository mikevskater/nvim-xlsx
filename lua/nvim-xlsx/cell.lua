--- Cell representation for xlsx
--- @module nvim-xlsx.cell

local column_utils = require("nvim-xlsx.utils.column")
local xml = require("nvim-xlsx.xml.writer")

local M = {}

--- Cell value types
M.TYPE = {
  NUMBER = "n",
  STRING = "s",        -- shared string reference
  INLINE_STRING = "inlineStr",
  BOOLEAN = "b",
  ERROR = "e",
  FORMULA = "str",     -- formula result type
}

---@class Cell
---@field row integer Row number (1-indexed)
---@field col integer Column number (1-indexed)
---@field value any The cell value
---@field value_type string Type of value (from M.TYPE)
---@field style_index? integer Index into styles
---@field formula? string Formula string (without leading =)
local Cell = {}
Cell.__index = Cell

--- Create a new cell
--- @param row integer Row number
--- @param col integer Column number
--- @param value? any Cell value
--- @return Cell
function M.new(row, col, value)
  local self = setmetatable({}, Cell)
  self.row = row
  self.col = col
  self.style_index = nil
  self.formula = nil
  self:set_value(value)
  return self
end

--- Detect the type of a value and set appropriately
--- @param value any The value to set
function Cell:set_value(value)
  if value == nil then
    self.value = nil
    self.value_type = nil
  elseif type(value) == "number" then
    self.value = value
    self.value_type = M.TYPE.NUMBER
  elseif type(value) == "boolean" then
    self.value = value and 1 or 0
    self.value_type = M.TYPE.BOOLEAN
  elseif type(value) == "string" then
    -- Check if it's a formula (starts with =)
    if value:sub(1, 1) == "=" then
      self.formula = value:sub(2)  -- Store without the =
      self.value = nil
      self.value_type = nil
    else
      self.value = value
      self.value_type = M.TYPE.INLINE_STRING
    end
  else
    -- Convert other types to string
    self.value = tostring(value)
    self.value_type = M.TYPE.INLINE_STRING
  end
end

--- Get the cell reference (e.g., "A1")
--- @return string
function Cell:get_ref()
  return column_utils.make_ref(self.row, self.col)
end

--- Check if cell has any content (value, formula, or style)
--- @return boolean
function Cell:has_content()
  return self.value ~= nil or self.formula ~= nil or (self.style_index and self.style_index > 0)
end

--- Generate XML for this cell
--- @return string XML representation
function Cell:to_xml()
  if not self:has_content() then
    return ""
  end

  -- Handle styled empty cells
  if self.value == nil and self.formula == nil then
    local ref = self:get_ref()
    return xml.empty_element("c", { r = ref, s = self.style_index })
  end

  local ref = self:get_ref()
  local attrs = { r = ref }

  -- Add style index if present
  if self.style_index and self.style_index > 0 then
    attrs.s = self.style_index
  end

  -- Handle formula cells
  if self.formula then
    local content = xml.element("f", self.formula)
    -- If there's a cached value, include it
    if self.value ~= nil then
      content = content .. xml.element("v", self.value)
    end
    return xml.element_raw("c", content, attrs)
  end

  -- Add type attribute for non-number types
  if self.value_type and self.value_type ~= M.TYPE.NUMBER then
    attrs.t = self.value_type
  end

  -- Generate value content based on type
  local content
  if self.value_type == M.TYPE.INLINE_STRING then
    -- Inline strings use <is><t>value</t></is>
    content = xml.element_raw("is", xml.element("t", self.value))
  elseif self.value ~= nil then
    -- Numbers, booleans, shared string refs use <v>
    content = xml.element("v", self.value)
  else
    content = ""
  end

  return xml.element_raw("c", content, attrs)
end

M.Cell = Cell

return M
