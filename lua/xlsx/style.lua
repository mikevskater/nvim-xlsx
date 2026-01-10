--- Style management for xlsx
--- Handles fonts, fills, borders, number formats, and cell styles
--- @module xlsx.style

local xml = require("xlsx.xml.writer")
local color_utils = require("xlsx.utils.color")

local M = {}

-- Built-in number format IDs (0-163 are reserved)
M.BUILTIN_FORMATS = {
  general = 0,
  number = 1,           -- 0
  number_d2 = 2,        -- 0.00
  number_thousands = 3, -- #,##0
  number_thousands_d2 = 4, -- #,##0.00
  percent = 9,          -- 0%
  percent_d2 = 10,      -- 0.00%
  scientific = 11,      -- 0.00E+00
  fraction = 12,        -- # ?/?
  fraction_d2 = 13,     -- # ??/??
  date = 14,            -- m/d/yyyy (locale dependent)
  date_d_mon_yy = 15,   -- d-mmm-yy
  date_d_mon = 16,      -- d-mmm
  date_mon_yy = 17,     -- mmm-yy
  time_12h = 18,        -- h:mm AM/PM
  time_12h_ss = 19,     -- h:mm:ss AM/PM
  time_24h = 20,        -- h:mm
  time_24h_ss = 21,     -- h:mm:ss
  datetime = 22,        -- m/d/yyyy h:mm
  accounting = 37,      -- #,##0_);(#,##0)
  accounting_red = 38,  -- #,##0_);[Red](#,##0)
  accounting_d2 = 39,   -- #,##0.00_);(#,##0.00)
  accounting_d2_red = 40, -- #,##0.00_);[Red](#,##0.00)
  text = 49,            -- @
}

-- Border styles
M.BORDER_STYLES = {
  none = nil,
  thin = "thin",
  medium = "medium",
  thick = "thick",
  dashed = "dashed",
  dotted = "dotted",
  double = "double",
  hair = "hair",
  mediumDashed = "mediumDashed",
  dashDot = "dashDot",
  mediumDashDot = "mediumDashDot",
  dashDotDot = "dashDotDot",
  mediumDashDotDot = "mediumDashDotDot",
  slantDashDot = "slantDashDot",
}

-- Horizontal alignment
M.HALIGN = {
  left = "left",
  center = "center",
  right = "right",
  fill = "fill",
  justify = "justify",
  centerContinuous = "centerContinuous",
  distributed = "distributed",
}

-- Vertical alignment
M.VALIGN = {
  top = "top",
  center = "center",
  bottom = "bottom",
  justify = "justify",
  distributed = "distributed",
}

-- Underline styles
M.UNDERLINE = {
  none = nil,
  single = "single",
  double = "double",
  singleAccounting = "singleAccounting",
  doubleAccounting = "doubleAccounting",
}

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
  if M.BUILTIN_FORMATS[format] then
    return M.BUILTIN_FORMATS[format]
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

--- Generate XML for a single font
--- @param font table Font definition
--- @return string XML
function StyleRegistry:_font_to_xml(font)
  local parts = {}

  if font.bold then
    table.insert(parts, "<b/>")
  end
  if font.italic then
    table.insert(parts, "<i/>")
  end
  if font.strike then
    table.insert(parts, "<strike/>")
  end
  if font.underline then
    local u = font.underline == true and "single" or font.underline
    table.insert(parts, xml.empty_element("u", { val = u }))
  end

  table.insert(parts, xml.empty_element("sz", { val = font.size or 11 }))

  if font.color then
    table.insert(parts, xml.empty_element("color", { rgb = font.color }))
  end

  table.insert(parts, xml.empty_element("name", { val = font.name or "Calibri" }))

  return xml.element_raw("font", table.concat(parts))
end

--- Generate XML for a single fill
--- @param fill table Fill definition
--- @return string XML
function StyleRegistry:_fill_to_xml(fill)
  local parts = {}

  local pattern = fill.pattern or "none"
  local attrs = { patternType = pattern }

  if fill.fgColor and pattern ~= "none" then
    table.insert(parts, xml.empty_element("fgColor", { rgb = fill.fgColor }))
  end
  if fill.bgColor then
    table.insert(parts, xml.empty_element("bgColor", { rgb = fill.bgColor }))
  end

  local patternFill
  if #parts > 0 then
    patternFill = xml.element_raw("patternFill", table.concat(parts), attrs)
  else
    patternFill = xml.empty_element("patternFill", attrs)
  end

  return xml.element_raw("fill", patternFill)
end

--- Generate XML for a single border edge
--- @param edge string Edge name
--- @param def table Edge definition
--- @return string XML
function StyleRegistry:_border_edge_to_xml(edge, def)
  if not def or not def.style then
    return "<" .. edge .. "/>"
  end

  local parts = {}
  if def.color then
    table.insert(parts, xml.empty_element("color", { rgb = def.color }))
  end

  if #parts > 0 then
    return xml.element_raw(edge, table.concat(parts), { style = def.style })
  else
    return xml.empty_element(edge, { style = def.style })
  end
end

--- Generate XML for a single border
--- @param border table Border definition
--- @return string XML
function StyleRegistry:_border_to_xml(border)
  local parts = {}

  table.insert(parts, self:_border_edge_to_xml("left", border.left))
  table.insert(parts, self:_border_edge_to_xml("right", border.right))
  table.insert(parts, self:_border_edge_to_xml("top", border.top))
  table.insert(parts, self:_border_edge_to_xml("bottom", border.bottom))
  table.insert(parts, self:_border_edge_to_xml("diagonal", border.diagonal))

  return xml.element_raw("border", table.concat(parts))
end

--- Generate XML for a cell format (xf)
--- @param xf table Cell format definition
--- @param for_style boolean Whether this is for cellStyleXfs (vs cellXfs)
--- @return string XML
function StyleRegistry:_xf_to_xml(xf, for_style)
  local attrs = {
    numFmtId = xf.numFmtId or 0,
    fontId = xf.fontId or 0,
    fillId = xf.fillId or 0,
    borderId = xf.borderId or 0,
  }

  if not for_style then
    attrs.xfId = 0

    if xf.fontId and xf.fontId > 0 then
      attrs.applyFont = "1"
    end
    if xf.fillId and xf.fillId > 0 then
      attrs.applyFill = "1"
    end
    if xf.borderId and xf.borderId > 0 then
      attrs.applyBorder = "1"
    end
    if xf.numFmtId and xf.numFmtId > 0 then
      attrs.applyNumberFormat = "1"
    end
  end

  -- Alignment
  local has_alignment = xf.halign or xf.valign or xf.wrapText or xf.rotation or xf.indent
  if has_alignment then
    attrs.applyAlignment = "1"
    local align_attrs = {}
    if xf.halign then align_attrs.horizontal = xf.halign end
    if xf.valign then align_attrs.vertical = xf.valign end
    if xf.wrapText then align_attrs.wrapText = "1" end
    if xf.rotation then align_attrs.textRotation = xf.rotation end
    if xf.indent then align_attrs.indent = xf.indent end

    return xml.element_raw("xf", xml.empty_element("alignment", align_attrs), attrs)
  end

  return xml.empty_element("xf", attrs)
end

--- Generate the complete styles.xml content
--- @return string XML
function StyleRegistry:to_xml()
  local b = xml.builder()

  b:declaration()
  b:open("styleSheet", { xmlns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main" })

  -- Number formats (only custom ones)
  if #self.numFmts > 0 then
    b:open("numFmts", { count = #self.numFmts })
    for _, fmt in ipairs(self.numFmts) do
      b:empty("numFmt", { numFmtId = fmt.id, formatCode = fmt.code })
    end
    b:close("numFmts")
  end

  -- Fonts
  b:open("fonts", { count = #self.fonts })
  for _, font in ipairs(self.fonts) do
    b:raw(self:_font_to_xml(font))
  end
  b:close("fonts")

  -- Fills
  b:open("fills", { count = #self.fills })
  for _, fill in ipairs(self.fills) do
    b:raw(self:_fill_to_xml(fill))
  end
  b:close("fills")

  -- Borders
  b:open("borders", { count = #self.borders })
  for _, border in ipairs(self.borders) do
    b:raw(self:_border_to_xml(border))
  end
  b:close("borders")

  -- Cell style formats (just one default)
  b:raw('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')

  -- Cell formats
  b:open("cellXfs", { count = #self.cellXfs })
  for _, xf in ipairs(self.cellXfs) do
    b:raw(self:_xf_to_xml(xf, false))
  end
  b:close("cellXfs")

  -- Cell styles
  b:raw('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>')

  b:close("styleSheet")

  return b:to_string()
end

M.StyleRegistry = StyleRegistry

return M
