--- Styles XML parsing for xlsx
--- @module nvim-xlsx.parts.styles_part
---
--- Handles parsing of xl/styles.xml
--- For reading files, we preserve the original styles and map them
--- to our internal style format for modification.

local parser = require("nvim-xlsx.xml.parser")

local M = {}

---@class NumberFormatInfo
---@field id integer Format ID
---@field code string Format code

---@class FontInfo
---@field name string? Font name
---@field size number? Font size
---@field bold boolean
---@field italic boolean
---@field underline string?
---@field strike boolean
---@field color string? ARGB color

---@class FillInfo
---@field pattern_type string?
---@field fg_color string? Foreground ARGB color
---@field bg_color string? Background ARGB color

---@class BorderEdge
---@field style string?
---@field color string?

---@class BorderInfo
---@field left BorderEdge?
---@field right BorderEdge?
---@field top BorderEdge?
---@field bottom BorderEdge?
---@field diagonal BorderEdge?

---@class CellXf
---@field font_id integer?
---@field fill_id integer?
---@field border_id integer?
---@field num_fmt_id integer?
---@field apply_font boolean
---@field apply_fill boolean
---@field apply_border boolean
---@field apply_number_format boolean
---@field apply_alignment boolean
---@field halign string?
---@field valign string?
---@field wrap_text boolean
---@field rotation integer?
---@field indent integer?

---@class StylesData
---@field num_formats NumberFormatInfo[]
---@field fonts FontInfo[]
---@field fills FillInfo[]
---@field borders BorderInfo[]
---@field cell_xfs CellXf[]
---@field raw_xml string Original XML for preservation

--- Parse color from element
--- @param elem_text string Element text containing color
--- @return string? ARGB color
local function parse_color(elem_text)
  if not elem_text then return nil end

  -- Check for rgb attribute
  local rgb = elem_text:match('rgb="([^"]+)"')
  if rgb then
    return rgb
  end

  -- Check for theme + tint (simplified - just return theme indicator)
  local theme = elem_text:match('theme="([^"]+)"')
  if theme then
    return "theme:" .. theme
  end

  -- Check for indexed color
  local indexed = elem_text:match('indexed="([^"]+)"')
  if indexed then
    return "indexed:" .. indexed
  end

  return nil
end

--- Parse a border edge element
--- @param edge_text string? Edge element text
--- @return BorderEdge?
local function parse_border_edge(edge_text)
  if not edge_text or edge_text == "" then
    return nil
  end

  local style = edge_text:match('style="([^"]+)"')
  if not style then
    return nil
  end

  local color = parse_color(edge_text)

  return {
    style = style,
    color = color,
  }
end

--- Parse number formats section
--- @param xml_content string
--- @return NumberFormatInfo[]
local function parse_num_formats(xml_content)
  local formats = {}

  local num_fmts_section = xml_content:match("<numFmts[^>]*>(.-)</numFmts>")
  if num_fmts_section then
    local fmt_elements = parser.find_all(num_fmts_section, "numFmt")
    for _, fmt in ipairs(fmt_elements) do
      table.insert(formats, {
        id = tonumber(fmt.attrs.numFmtId),
        code = fmt.attrs.formatCode,
      })
    end
  end

  return formats
end

--- Parse fonts section
--- @param xml_content string
--- @return FontInfo[]
local function parse_fonts(xml_content)
  local fonts = {}

  local fonts_section = xml_content:match("<fonts[^>]*>(.-)</fonts>")
  if fonts_section then
    local font_elements = parser.find_all(fonts_section, "font")
    for _, font in ipairs(font_elements) do
      local text = font.text or ""

      local font_info = {
        name = text:match('<name val="([^"]+)"'),
        size = tonumber(text:match('<sz val="([^"]+)"')),
        bold = text:find("<b") ~= nil and not text:find('<b val="0"'),
        italic = text:find("<i") ~= nil and not text:find('<i val="0"'),
        underline = text:match('<u val="([^"]+)"') or (text:find("<u") and "single"),
        strike = text:find("<strike") ~= nil,
        color = nil,
      }

      -- Parse color
      local color_elem = text:match("<color([^>]+)>") or text:match("<color([^/]+)/>")
      if color_elem then
        font_info.color = parse_color(color_elem)
      end

      table.insert(fonts, font_info)
    end
  end

  return fonts
end

--- Parse fills section
--- @param xml_content string
--- @return FillInfo[]
local function parse_fills(xml_content)
  local fills = {}

  local fills_section = xml_content:match("<fills[^>]*>(.-)</fills>")
  if fills_section then
    local fill_elements = parser.find_all(fills_section, "fill")
    for _, fill in ipairs(fill_elements) do
      local text = fill.text or ""

      local fill_info = {
        pattern_type = text:match('<patternFill patternType="([^"]+)"'),
        fg_color = nil,
        bg_color = nil,
      }

      -- Parse foreground color
      local fg_elem = text:match("<fgColor([^>]+)>") or text:match("<fgColor([^/]+)/>")
      if fg_elem then
        fill_info.fg_color = parse_color(fg_elem)
      end

      -- Parse background color
      local bg_elem = text:match("<bgColor([^>]+)>") or text:match("<bgColor([^/]+)/>")
      if bg_elem then
        fill_info.bg_color = parse_color(bg_elem)
      end

      table.insert(fills, fill_info)
    end
  end

  return fills
end

--- Parse borders section
--- @param xml_content string
--- @return BorderInfo[]
local function parse_borders(xml_content)
  local borders = {}

  local borders_section = xml_content:match("<borders[^>]*>(.-)</borders>")
  if borders_section then
    local border_elements = parser.find_all(borders_section, "border")
    for _, border in ipairs(border_elements) do
      local text = border.text or ""

      local border_info = {
        left = parse_border_edge(text:match("<left([^>]*>.-</left>)") or text:match("<left([^/]*)/>")),
        right = parse_border_edge(text:match("<right([^>]*>.-</right>)") or text:match("<right([^/]*)/>")),
        top = parse_border_edge(text:match("<top([^>]*>.-</top>)") or text:match("<top([^/]*)/>")),
        bottom = parse_border_edge(text:match("<bottom([^>]*>.-</bottom>)") or text:match("<bottom([^/]*)/>")),
        diagonal = parse_border_edge(text:match("<diagonal([^>]*>.-</diagonal>)") or text:match("<diagonal([^/]*)/>")),
      }

      table.insert(borders, border_info)
    end
  end

  return borders
end

--- Parse cellXfs section (cell formats)
--- @param xml_content string
--- @return CellXf[]
local function parse_cell_xfs(xml_content)
  local xfs = {}

  local cell_xfs_section = xml_content:match("<cellXfs[^>]*>(.-)</cellXfs>")
  if cell_xfs_section then
    local xf_elements = parser.find_all(cell_xfs_section, "xf")
    for _, xf in ipairs(xf_elements) do
      local text = xf.text or ""

      local xf_info = {
        font_id = tonumber(xf.attrs.fontId),
        fill_id = tonumber(xf.attrs.fillId),
        border_id = tonumber(xf.attrs.borderId),
        num_fmt_id = tonumber(xf.attrs.numFmtId),
        apply_font = xf.attrs.applyFont == "1",
        apply_fill = xf.attrs.applyFill == "1",
        apply_border = xf.attrs.applyBorder == "1",
        apply_number_format = xf.attrs.applyNumberFormat == "1",
        apply_alignment = xf.attrs.applyAlignment == "1",
        halign = nil,
        valign = nil,
        wrap_text = false,
        rotation = nil,
        indent = nil,
      }

      -- Parse alignment
      local align = text:match("<alignment([^/]*)/?>")
      if align then
        xf_info.halign = align:match('horizontal="([^"]+)"')
        xf_info.valign = align:match('vertical="([^"]+)"')
        xf_info.wrap_text = align:find('wrapText="1"') ~= nil
        xf_info.rotation = tonumber(align:match('textRotation="([^"]+)"'))
        xf_info.indent = tonumber(align:match('indent="([^"]+)"'))
      end

      table.insert(xfs, xf_info)
    end
  end

  return xfs
end

--- Parse styles.xml content
--- @param xml_content string The XML content of styles.xml
--- @return StylesData
function M.parse(xml_content)
  local data = {
    num_formats = parse_num_formats(xml_content),
    fonts = parse_fonts(xml_content),
    fills = parse_fills(xml_content),
    borders = parse_borders(xml_content),
    cell_xfs = parse_cell_xfs(xml_content),
    raw_xml = xml_content,  -- Preserve for write-back
  }

  return data
end

--- Get number format code by ID
--- @param data StylesData
--- @param num_fmt_id integer
--- @return string?
function M.get_number_format(data, num_fmt_id)
  -- Check custom formats
  for _, fmt in ipairs(data.num_formats) do
    if fmt.id == num_fmt_id then
      return fmt.code
    end
  end

  -- Return nil for built-in formats (0-163)
  -- The caller should handle built-in format codes
  return nil
end

--- Get font info by index
--- @param data StylesData
--- @param font_id integer 0-indexed
--- @return FontInfo?
function M.get_font(data, font_id)
  return data.fonts[font_id + 1]
end

--- Get fill info by index
--- @param data StylesData
--- @param fill_id integer 0-indexed
--- @return FillInfo?
function M.get_fill(data, fill_id)
  return data.fills[fill_id + 1]
end

--- Get border info by index
--- @param data StylesData
--- @param border_id integer 0-indexed
--- @return BorderInfo?
function M.get_border(data, border_id)
  return data.borders[border_id + 1]
end

--- Get cell format info by index
--- @param data StylesData
--- @param xf_id integer 0-indexed
--- @return CellXf?
function M.get_cell_xf(data, xf_id)
  return data.cell_xfs[xf_id + 1]
end

return M
