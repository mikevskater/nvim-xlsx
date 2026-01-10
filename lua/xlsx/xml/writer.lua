--- XML Writer for generating well-formed XML
--- @module xlsx.xml.writer

local M = {}

--- Escape special XML characters in text content
--- @param str string The string to escape
--- @return string Escaped string
function M.escape(str)
  if str == nil then
    return ""
  end
  str = tostring(str)
  -- Order matters: & must be first
  str = str:gsub("&", "&amp;")
  str = str:gsub("<", "&lt;")
  str = str:gsub(">", "&gt;")
  str = str:gsub('"', "&quot;")
  str = str:gsub("'", "&apos;")
  return str
end

--- Escape attribute value (same as text but always quoted)
--- @param str string The attribute value
--- @return string Escaped and quoted attribute value
function M.escape_attr(str)
  return '"' .. M.escape(str) .. '"'
end

--- Format attributes table as XML attribute string
--- @param attrs? table<string, string|number|boolean> Attribute key-value pairs
--- @return string Formatted attributes string (with leading space if non-empty)
function M.format_attrs(attrs)
  if not attrs then
    return ""
  end

  local parts = {}
  -- Sort keys for consistent output
  local keys = {}
  for k in pairs(attrs) do
    table.insert(keys, k)
  end
  table.sort(keys)

  for _, k in ipairs(keys) do
    local v = attrs[k]
    if v ~= nil then
      table.insert(parts, k .. "=" .. M.escape_attr(tostring(v)))
    end
  end

  if #parts == 0 then
    return ""
  end
  return " " .. table.concat(parts, " ")
end

--- Create a self-closing XML element
--- @param name string Element tag name
--- @param attrs? table<string, string|number|boolean> Attributes
--- @return string XML element string
function M.empty_element(name, attrs)
  return "<" .. name .. M.format_attrs(attrs) .. "/>"
end

--- Create an XML element with text content
--- @param name string Element tag name
--- @param content string|number|nil Text content
--- @param attrs? table<string, string|number|boolean> Attributes
--- @return string XML element string
function M.element(name, content, attrs)
  if content == nil or content == "" then
    return M.empty_element(name, attrs)
  end
  return "<" .. name .. M.format_attrs(attrs) .. ">" .. M.escape(tostring(content)) .. "</" .. name .. ">"
end

--- Create an XML element with raw (unescaped) content
--- @param name string Element tag name
--- @param content string Raw XML content (already escaped/formatted)
--- @param attrs? table<string, string|number|boolean> Attributes
--- @return string XML element string
function M.element_raw(name, content, attrs)
  if content == nil or content == "" then
    return M.empty_element(name, attrs)
  end
  return "<" .. name .. M.format_attrs(attrs) .. ">" .. content .. "</" .. name .. ">"
end

--- Create opening tag
--- @param name string Element tag name
--- @param attrs? table<string, string|number|boolean> Attributes
--- @return string Opening tag string
function M.open_tag(name, attrs)
  return "<" .. name .. M.format_attrs(attrs) .. ">"
end

--- Create closing tag
--- @param name string Element tag name
--- @return string Closing tag string
function M.close_tag(name)
  return "</" .. name .. ">"
end

--- XML declaration
--- @param version? string XML version (default "1.0")
--- @param encoding? string Encoding (default "UTF-8")
--- @param standalone? string Standalone attribute
--- @return string XML declaration
function M.declaration(version, encoding, standalone)
  version = version or "1.0"
  encoding = encoding or "UTF-8"
  local decl = '<?xml version="' .. version .. '" encoding="' .. encoding .. '"'
  if standalone then
    decl = decl .. ' standalone="' .. standalone .. '"'
  end
  return decl .. "?>"
end

---@class XmlBuilder
---@field private _parts string[]
local XmlBuilder = {}
XmlBuilder.__index = XmlBuilder

--- Create a new XML builder for efficient string building
--- @return XmlBuilder
function M.builder()
  local self = setmetatable({}, XmlBuilder)
  self._parts = {}
  return self
end

--- Add XML declaration
--- @param version? string
--- @param encoding? string
--- @param standalone? string
--- @return XmlBuilder self for chaining
function XmlBuilder:declaration(version, encoding, standalone)
  table.insert(self._parts, M.declaration(version, encoding, standalone))
  return self
end

--- Add raw string (no escaping)
--- @param str string
--- @return XmlBuilder self for chaining
function XmlBuilder:raw(str)
  table.insert(self._parts, str)
  return self
end

--- Add newline
--- @return XmlBuilder self for chaining
function XmlBuilder:nl()
  table.insert(self._parts, "\n")
  return self
end

--- Add empty element
--- @param name string
--- @param attrs? table
--- @return XmlBuilder self for chaining
function XmlBuilder:empty(name, attrs)
  table.insert(self._parts, M.empty_element(name, attrs))
  return self
end

--- Add element with text content
--- @param name string
--- @param content string|number|nil
--- @param attrs? table
--- @return XmlBuilder self for chaining
function XmlBuilder:elem(name, content, attrs)
  table.insert(self._parts, M.element(name, content, attrs))
  return self
end

--- Add element with raw content
--- @param name string
--- @param content string
--- @param attrs? table
--- @return XmlBuilder self for chaining
function XmlBuilder:elem_raw(name, content, attrs)
  table.insert(self._parts, M.element_raw(name, content, attrs))
  return self
end

--- Add opening tag
--- @param name string
--- @param attrs? table
--- @return XmlBuilder self for chaining
function XmlBuilder:open(name, attrs)
  table.insert(self._parts, M.open_tag(name, attrs))
  return self
end

--- Add closing tag
--- @param name string
--- @return XmlBuilder self for chaining
function XmlBuilder:close(name)
  table.insert(self._parts, M.close_tag(name))
  return self
end

--- Get the built XML string
--- @return string
function XmlBuilder:to_string()
  return table.concat(self._parts)
end

return M
