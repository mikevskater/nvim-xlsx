--- Style registry for xlsx
--- Manages fonts, fills, borders, number formats, and cell styles
--- @module nvim-xlsx.style.registry

local color_utils = require("nvim-xlsx.utils.color")
local constants = require("nvim-xlsx.style.constants")

local M = {}

---@class StyleRegistry
---@field fonts table[] Font definitions
---@field fills table[] Fill definitions
---@field borders table[] Border definitions
---@field numFmts table[] Custom number format definitions
---@field cellXfs table[] Cell format definitions
---@field font_map table<string, integer> Font dedup map
---@field fill_map table<string, integer> Fill dedup map
---@field border_map table<string, integer> Border dedup map
---@field numFmt_map table<string, integer> NumFmt dedup map
---@field xf_map table<string, integer> CellXf dedup map
---@field next_numFmt_id integer Next custom numFmt ID (starts at 164)
local StyleRegistry = {}
StyleRegistry.__index = StyleRegistry

--- Create a new style registry
--- @return StyleRegistry
function M.new_registry()
  local self = setmetatable({}, StyleRegistry)

  -- Initialize with required defaults
  self.fonts = {}
  self.fills = {}
  self.borders = {}
  self.numFmts = {}
  self.cellXfs = {}

  self.font_map = {}
  self.fill_map = {}
  self.border_map = {}
  self.numFmt_map = {}
  self.xf_map = {}

  self.next_numFmt_id = 164  -- Custom formats start at 164

  -- Add required defaults
  self:_add_default_font()
  self:_add_default_fills()
  self:_add_default_border()
  self:_add_default_xf()

  return self
end

--- Add default font (index 0)
function StyleRegistry:_add_default_font()
  local font = {
    name = "Calibri",
    size = 11,
  }
  table.insert(self.fonts, font)
  self.font_map[self:_font_key(font)] = 0
end

--- Add required default fills (index 0 = none, index 1 = gray125)
function StyleRegistry:_add_default_fills()
  -- Fill 0: none
  local fill_none = { pattern = "none" }
  table.insert(self.fills, fill_none)
  self.fill_map[self:_fill_key(fill_none)] = 0

  -- Fill 1: gray125 (required by Excel)
  local fill_gray = { pattern = "gray125" }
  table.insert(self.fills, fill_gray)
  self.fill_map[self:_fill_key(fill_gray)] = 1
end

--- Add default border (index 0 = no borders)
function StyleRegistry:_add_default_border()
  local border = {}
  table.insert(self.borders, border)
  self.border_map[self:_border_key(border)] = 0
end

--- Add default cell format (index 0)
function StyleRegistry:_add_default_xf()
  local xf = {
    fontId = 0,
    fillId = 0,
    borderId = 0,
    numFmtId = 0,
  }
  table.insert(self.cellXfs, xf)
  self.xf_map[self:_xf_key(xf)] = 0
end

--- Generate a unique key for font deduplication
--- @param font table Font definition
--- @return string
function StyleRegistry:_font_key(font)
  return string.format("%s|%s|%s|%s|%s|%s|%s",
    font.name or "",
    font.size or "",
    font.bold and "B" or "",
    font.italic and "I" or "",
    font.underline or "",
    font.strike and "S" or "",
    font.color or ""
  )
end

--- Generate a unique key for fill deduplication
--- @param fill table Fill definition
--- @return string
function StyleRegistry:_fill_key(fill)
  return string.format("%s|%s|%s",
    fill.pattern or "",
    fill.fgColor or "",
    fill.bgColor or ""
  )
end

--- Generate a unique key for border deduplication
--- @param border table Border definition
--- @return string
function StyleRegistry:_border_key(border)
  local parts = {}
  for _, edge in ipairs({"left", "right", "top", "bottom", "diagonal"}) do
    local e = border[edge] or {}
    table.insert(parts, string.format("%s:%s", e.style or "", e.color or ""))
  end
  return table.concat(parts, "|")
end

--- Generate a unique key for cell format deduplication
--- @param xf table Cell format definition
--- @return string
function StyleRegistry:_xf_key(xf)
  return string.format("%d|%d|%d|%d|%s|%s|%s|%s|%d",
    xf.fontId or 0,
    xf.fillId or 0,
    xf.borderId or 0,
    xf.numFmtId or 0,
    xf.halign or "",
    xf.valign or "",
    xf.wrapText and "W" or "",
    xf.rotation or "",
    xf.indent or 0
  )
end

--- Register or get existing font index
--- @param font table Font definition
--- @return integer Font index
function StyleRegistry:register_font(font)
  local key = self:_font_key(font)
  if self.font_map[key] then
    return self.font_map[key]
  end

  local idx = #self.fonts
  table.insert(self.fonts, font)
  self.font_map[key] = idx
  return idx
end

--- Register or get existing fill index
--- @param fill table Fill definition
--- @return integer Fill index
function StyleRegistry:register_fill(fill)
  local key = self:_fill_key(fill)
  if self.fill_map[key] then
    return self.fill_map[key]
  end

  local idx = #self.fills
  table.insert(self.fills, fill)
  self.fill_map[key] = idx
  return idx
end

--- Register or get existing border index
--- @param border table Border definition
--- @return integer Border index
function StyleRegistry:register_border(border)
  local key = self:_border_key(border)
  if self.border_map[key] then
    return self.border_map[key]
  end

  local idx = #self.borders
  table.insert(self.borders, border)
  self.border_map[key] = idx
  return idx
end

--- Register or get existing number format ID
--- @param format string|integer Format code or built-in ID
--- @return integer NumFmt ID
function StyleRegistry:register_numfmt(format)
  -- If it's already an integer, assume it's a built-in
  if type(format) == "number" then
    return format
  end

  -- Check if it's a built-in name
  if constants.BUILTIN_FORMATS[format] then
    return constants.BUILTIN_FORMATS[format]
  end

  -- Check if we already have this custom format
  if self.numFmt_map[format] then
    return self.numFmt_map[format]
  end

  -- Register new custom format
  local id = self.next_numFmt_id
  self.next_numFmt_id = self.next_numFmt_id + 1

  table.insert(self.numFmts, { id = id, code = format })
  self.numFmt_map[format] = id

  return id
end

--- Create a style and return its index
--- @param def table Style definition
--- @return integer Style index (xf index)
function StyleRegistry:create_style(def)
  def = def or {}

  -- Build font if any font properties specified
  local fontId = 0
  if def.font or def.bold or def.italic or def.underline or def.strike or def.font_color or def.font_size or def.font_name then
    local font = {
      name = def.font_name or def.font or "Calibri",
      size = def.font_size or 11,
      bold = def.bold,
      italic = def.italic,
      underline = def.underline,
      strike = def.strike,
      color = def.font_color and color_utils.to_argb(def.font_color),
    }
    fontId = self:register_font(font)
  end

  -- Build fill if background specified
  local fillId = 0
  if def.bg_color or def.fill_color or def.pattern then
    local fill = {
      pattern = def.pattern or "solid",
      fgColor = color_utils.to_argb(def.bg_color or def.fill_color),
    }
    fillId = self:register_fill(fill)
  end

  -- Build border if any border properties specified
  local borderId = 0
  if def.border or def.border_style or def.border_color or
     def.border_left or def.border_right or def.border_top or def.border_bottom then

    local default_style = def.border_style or (def.border and "thin") or nil
    local default_color = def.border_color and color_utils.to_argb(def.border_color)

    local border = {}

    -- Handle individual edges or use default
    for _, edge in ipairs({"left", "right", "top", "bottom"}) do
      local edge_def = def["border_" .. edge]
      if edge_def then
        if type(edge_def) == "string" then
          border[edge] = { style = edge_def, color = default_color }
        elseif type(edge_def) == "table" then
          border[edge] = {
            style = edge_def.style or default_style,
            color = edge_def.color and color_utils.to_argb(edge_def.color) or default_color,
          }
        elseif edge_def == true then
          border[edge] = { style = default_style, color = default_color }
        end
      elseif def.border then
        border[edge] = { style = default_style, color = default_color }
      end
    end

    borderId = self:register_border(border)
  end

  -- Number format
  local numFmtId = 0
  if def.num_format or def.number_format then
    numFmtId = self:register_numfmt(def.num_format or def.number_format)
  end

  -- Build cell format
  local xf = {
    fontId = fontId,
    fillId = fillId,
    borderId = borderId,
    numFmtId = numFmtId,
    halign = def.halign or def.align,
    valign = def.valign,
    wrapText = def.wrap_text or def.wrap,
    rotation = def.rotation,
    indent = def.indent,
  }

  -- Check for existing
  local key = self:_xf_key(xf)
  if self.xf_map[key] then
    return self.xf_map[key]
  end

  -- Register new
  local idx = #self.cellXfs
  table.insert(self.cellXfs, xf)
  self.xf_map[key] = idx

  return idx
end

-- Export
M.StyleRegistry = StyleRegistry

return M
