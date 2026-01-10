--- Shared Strings parsing and generation for xlsx
--- @module xlsx.parts.shared_strings
---
--- Handles xl/sharedStrings.xml which contains deduplicated string values
--- referenced by index from cells.

local parser = require("xlsx.xml.parser")
local xml_writer = require("xlsx.xml.writer")
local templates = require("xlsx.xml.templates")

local M = {}

---@class SharedStrings
---@field strings string[] Array of unique strings (1-indexed)
---@field string_map table<string, integer> Map from string to index (0-indexed for Excel compatibility)
---@field count integer Total string references count
---@field unique_count integer Number of unique strings
local SharedStrings = {}
SharedStrings.__index = SharedStrings

--- Create a new SharedStrings instance
--- @return SharedStrings
function M.new()
  local self = setmetatable({}, SharedStrings)
  self.strings = {}
  self.string_map = {}
  self.count = 0
  self.unique_count = 0
  return self
end

--- Add a string and return its index
--- @param str string The string to add
--- @return integer index The 0-indexed position of the string
function SharedStrings:add(str)
  self.count = self.count + 1

  -- Check if string already exists
  local existing = self.string_map[str]
  if existing ~= nil then
    return existing
  end

  -- Add new string
  local index = self.unique_count
  self.strings[index + 1] = str  -- 1-indexed Lua array
  self.string_map[str] = index    -- 0-indexed for Excel
  self.unique_count = self.unique_count + 1

  return index
end

--- Get a string by its 0-indexed position
--- @param index integer 0-indexed position
--- @return string? The string at that index, or nil
function SharedStrings:get(index)
  return self.strings[index + 1]  -- Convert to 1-indexed Lua
end

--- Check if a string exists
--- @param str string The string to check
--- @return boolean
function SharedStrings:has(str)
  return self.string_map[str] ~= nil
end

--- Get the index of a string (or nil if not found)
--- @param str string The string to find
--- @return integer? 0-indexed position or nil
function SharedStrings:index_of(str)
  return self.string_map[str]
end

--- Parse shared strings from XML content
--- @param xml_content string The XML content of sharedStrings.xml
--- @return SharedStrings
function M.parse(xml_content)
  local ss = M.new()

  -- Find all <si> elements (string items)
  local si_elements = parser.find_all(xml_content, "si")

  for _, si in ipairs(si_elements) do
    local text = ""

    -- Check for simple text: <si><t>text</t></si>
    local t_match = si.text and si.text:match("<t[^>]*>([^<]*)</t>")
    if t_match then
      text = parser.unescape(t_match)
    else
      -- Check for rich text: <si><r><t>text</t></r>...</si>
      -- Concatenate all <t> elements within <r> elements
      if si.text then
        for t_content in si.text:gmatch("<t[^>]*>([^<]*)</t>") do
          text = text .. parser.unescape(t_content)
        end
      end
    end

    -- Add to the strings array (maintain order for index lookup)
    table.insert(ss.strings, text)
    ss.string_map[text] = #ss.strings - 1  -- 0-indexed
  end

  ss.unique_count = #ss.strings

  return ss
end

--- Generate XML content for shared strings
--- @return string XML content
function SharedStrings:to_xml()
  if self.unique_count == 0 then
    return nil  -- No shared strings needed
  end

  local b = xml_writer.builder()
  b:declaration()
  b:open("sst", {
    xmlns = templates.NS.SPREADSHEET,
    count = self.count,
    uniqueCount = self.unique_count,
  })

  for _, str in ipairs(self.strings) do
    local escaped = xml_writer.escape(str)
    b:raw("<si><t>" .. escaped .. "</t></si>")
  end

  b:close("sst")
  return b:to_string()
end

M.SharedStrings = SharedStrings

return M
