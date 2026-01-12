--- Style XML generation for xlsx
--- Generates styles.xml content
--- @module nvim-xlsx.style.xml

local xml = require("nvim-xlsx.xml.writer")

local M = {}

--- Generate XML for a single font
--- @param self StyleRegistry
--- @param font table Font definition
--- @return string XML
function M._font_to_xml(self, font)
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
--- @param self StyleRegistry
--- @param fill table Fill definition
--- @return string XML
function M._fill_to_xml(self, fill)
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
--- @param self StyleRegistry
--- @param edge string Edge name
--- @param def table Edge definition
--- @return string XML
function M._border_edge_to_xml(self, edge, def)
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
--- @param self StyleRegistry
--- @param border table Border definition
--- @return string XML
function M._border_to_xml(self, border)
  local parts = {}

  table.insert(parts, M._border_edge_to_xml(self, "left", border.left))
  table.insert(parts, M._border_edge_to_xml(self, "right", border.right))
  table.insert(parts, M._border_edge_to_xml(self, "top", border.top))
  table.insert(parts, M._border_edge_to_xml(self, "bottom", border.bottom))
  table.insert(parts, M._border_edge_to_xml(self, "diagonal", border.diagonal))

  return xml.element_raw("border", table.concat(parts))
end

--- Generate XML for a cell format (xf)
--- @param self StyleRegistry
--- @param xf table Cell format definition
--- @param for_style boolean Whether this is for cellStyleXfs (vs cellXfs)
--- @return string XML
function M._xf_to_xml(self, xf, for_style)
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
--- @param self StyleRegistry
--- @return string XML
function M.to_xml(self)
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
    b:raw(M._font_to_xml(self, font))
  end
  b:close("fonts")

  -- Fills
  b:open("fills", { count = #self.fills })
  for _, fill in ipairs(self.fills) do
    b:raw(M._fill_to_xml(self, fill))
  end
  b:close("fills")

  -- Borders
  b:open("borders", { count = #self.borders })
  for _, border in ipairs(self.borders) do
    b:raw(M._border_to_xml(self, border))
  end
  b:close("borders")

  -- Cell style formats (just one default)
  b:raw('<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')

  -- Cell formats
  b:open("cellXfs", { count = #self.cellXfs })
  for _, xf in ipairs(self.cellXfs) do
    b:raw(M._xf_to_xml(self, xf, false))
  end
  b:close("cellXfs")

  -- Cell styles
  b:raw('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>')

  b:close("styleSheet")

  return b:to_string()
end

return M
