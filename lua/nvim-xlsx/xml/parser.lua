--- XML Parser for xlsx files
--- @module xlsx.xml.parser
---
--- A focused XML parser for xlsx-specific patterns.
--- Not a full XML parser - optimized for the subset of XML used in xlsx files.

local M = {}

--- Unescape XML entities
--- @param str string
--- @return string
function M.unescape(str)
  if not str then return "" end
  str = str:gsub("&lt;", "<")
  str = str:gsub("&gt;", ">")
  str = str:gsub("&amp;", "&")
  str = str:gsub("&apos;", "'")
  str = str:gsub("&quot;", '"')
  -- Handle numeric entities
  str = str:gsub("&#(%d+);", function(n)
    return string.char(tonumber(n))
  end)
  str = str:gsub("&#x(%x+);", function(n)
    return string.char(tonumber(n, 16))
  end)
  return str
end

--- Parse attributes from an attribute string
--- @param attr_str string The attribute string (e.g., 'name="value" id="123"')
--- @return table<string, string> Attribute key-value pairs
function M.parse_attributes(attr_str)
  local attrs = {}
  if not attr_str or attr_str == "" then
    return attrs
  end

  -- Match attribute="value" or attribute='value'
  for name, quote, value in attr_str:gmatch('([%w_:%-]+)%s*=%s*(["\'])(.-)%2') do
    attrs[name] = M.unescape(value)
  end

  return attrs
end

--- Parse a single XML element (non-recursive, for simple elements)
--- @param xml string XML string
--- @param tag string Tag name to find
--- @return table? Element with attrs and text, or nil if not found
function M.parse_element(xml, tag)
  -- Try self-closing first
  local pattern_self = "<" .. tag .. "([^>]*)/>"
  local attrs_str = xml:match(pattern_self)
  if attrs_str then
    return {
      tag = tag,
      attrs = M.parse_attributes(attrs_str),
      text = nil,
      children = {},
    }
  end

  -- Try element with content
  local pattern_open = "<" .. tag .. "([^>]*)>"
  local pattern_close = "</" .. tag .. ">"

  local start_pos = xml:find("<" .. tag .. "[%s>]")
  if not start_pos then
    return nil
  end

  local _, open_end, attrs_str_match = xml:find("<" .. tag .. "([^>]*)>", start_pos)
  if not open_end then
    return nil
  end

  local close_start = xml:find(pattern_close, open_end + 1)
  if not close_start then
    return nil
  end

  local content = xml:sub(open_end + 1, close_start - 1)

  return {
    tag = tag,
    attrs = M.parse_attributes(attrs_str_match or ""),
    text = content,
    children = {},
  }
end

--- Find all elements with a given tag name
--- @param xml string XML string
--- @param tag string Tag name to find
--- @return table[] Array of elements
function M.find_all(xml, tag)
  local elements = {}

  -- Pattern for self-closing elements
  for attrs_str in xml:gmatch("<" .. tag .. "([^>]*)/>" ) do
    table.insert(elements, {
      tag = tag,
      attrs = M.parse_attributes(attrs_str),
      text = nil,
    })
  end

  -- Pattern for elements with content
  -- This is more complex - need to handle nested same-tags properly
  local pos = 1
  while true do
    local start_pos = xml:find("<" .. tag .. "[%s>]", pos)
    if not start_pos then break end

    local _, open_end, attrs_str = xml:find("<" .. tag .. "([^>]*)>", start_pos)
    if not open_end then
      pos = start_pos + 1
      break
    end

    -- Check if it's self-closing (already handled above)
    if xml:sub(open_end - 1, open_end - 1) == "/" then
      pos = open_end + 1
    else
      -- Find matching close tag (simplified - doesn't handle nested same-name tags)
      local close_start = xml:find("</" .. tag .. ">", open_end + 1)
      if close_start then
        local content = xml:sub(open_end + 1, close_start - 1)
        table.insert(elements, {
          tag = tag,
          attrs = M.parse_attributes(attrs_str or ""),
          text = content,
        })
        pos = close_start + #tag + 3
      else
        pos = open_end + 1
      end
    end
  end

  return elements
end

--- Extract text content from an element (strips child tags)
--- @param element table Element with text field
--- @return string Plain text content
function M.get_text(element)
  if not element or not element.text then
    return ""
  end
  -- Remove all XML tags to get plain text
  local text = element.text:gsub("<[^>]+>", "")
  return M.unescape(text)
end

--- Parse a simple XML document into a tree structure
--- @param xml string XML string
--- @return table Root element
function M.parse(xml)
  -- Remove XML declaration
  xml = xml:gsub("^%s*<%?xml[^?]*%?>%s*", "")

  -- Find root element
  local root_tag = xml:match("^%s*<([%w_:%-]+)")
  if not root_tag then
    return { tag = "root", attrs = {}, children = {}, text = xml }
  end

  return M.parse_element_recursive(xml, 1)
end

--- Recursively parse an element and its children
--- @param xml string XML string
--- @param pos integer Starting position
--- @return table Element tree
--- @return integer End position
function M.parse_element_recursive(xml, pos)
  -- Skip whitespace
  local _, new_pos = xml:find("^%s*", pos)
  pos = new_pos + 1

  -- Check for element start
  local tag_start, tag_end, tag, attrs_str = xml:find("^<([%w_:%-]+)([^>]*)>", pos)
  if not tag_start then
    -- No element found, return text content
    local text_end = xml:find("<", pos) or #xml + 1
    local text = xml:sub(pos, text_end - 1)
    return { tag = "#text", text = M.unescape(text) }, text_end
  end

  local element = {
    tag = tag,
    attrs = M.parse_attributes(attrs_str or ""),
    children = {},
    text = nil,
  }

  -- Check for self-closing
  if attrs_str and attrs_str:match("/%s*$") then
    element.attrs = M.parse_attributes(attrs_str:gsub("/%s*$", ""))
    return element, tag_end + 1
  end

  -- Parse children
  pos = tag_end + 1
  local close_tag = "</" .. tag .. ">"

  while pos <= #xml do
    -- Skip whitespace
    local _, ws_end = xml:find("^%s*", pos)
    pos = ws_end + 1

    -- Check for close tag
    if xml:sub(pos, pos + #close_tag - 1) == close_tag then
      pos = pos + #close_tag
      break
    end

    -- Check for child element
    if xml:sub(pos, pos) == "<" then
      if xml:sub(pos + 1, pos + 1) == "/" then
        -- Unexpected close tag
        break
      end
      local child, end_pos = M.parse_element_recursive(xml, pos)
      table.insert(element.children, child)
      pos = end_pos
    else
      -- Text content
      local text_end = xml:find("<", pos) or #xml + 1
      local text = xml:sub(pos, text_end - 1):gsub("^%s+", ""):gsub("%s+$", "")
      if text ~= "" then
        element.text = M.unescape(text)
      end
      pos = text_end
    end
  end

  return element, pos
end

--- Find child element by tag name
--- @param element table Parent element
--- @param tag string Tag name to find
--- @return table? Child element or nil
function M.find_child(element, tag)
  if not element or not element.children then
    return nil
  end
  for _, child in ipairs(element.children) do
    if child.tag == tag then
      return child
    end
  end
  return nil
end

--- Find all children with a given tag name
--- @param element table Parent element
--- @param tag string Tag name to find
--- @return table[] Array of matching children
function M.find_children(element, tag)
  local result = {}
  if not element or not element.children then
    return result
  end
  for _, child in ipairs(element.children) do
    if child.tag == tag then
      table.insert(result, child)
    end
  end
  return result
end

--- Get attribute value with optional default
--- @param element table Element
--- @param attr string Attribute name
--- @param default? any Default value if not found
--- @return any Attribute value or default
function M.get_attr(element, attr, default)
  if not element or not element.attrs then
    return default
  end
  local value = element.attrs[attr]
  if value == nil then
    return default
  end
  return value
end

--- Get numeric attribute value
--- @param element table Element
--- @param attr string Attribute name
--- @param default? number Default value if not found
--- @return number? Numeric value or default
function M.get_attr_number(element, attr, default)
  local value = M.get_attr(element, attr)
  if value == nil then
    return default
  end
  return tonumber(value) or default
end

--- Get boolean attribute value (handles "0", "1", "true", "false")
--- @param element table Element
--- @param attr string Attribute name
--- @param default? boolean Default value if not found
--- @return boolean? Boolean value or default
function M.get_attr_bool(element, attr, default)
  local value = M.get_attr(element, attr)
  if value == nil then
    return default
  end
  if value == "1" or value == "true" then
    return true
  elseif value == "0" or value == "false" then
    return false
  end
  return default
end

return M
