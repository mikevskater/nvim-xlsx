--- Worksheet XML parsing for xlsx
--- @module xlsx.parts.worksheet_part
---
--- Handles parsing of xl/worksheets/sheetN.xml

local parser = require("xlsx.xml.parser")
local column_utils = require("xlsx.utils.column")

local M = {}

---@class CellData
---@field row integer Row number (1-indexed)
---@field col integer Column number (1-indexed)
---@field ref string Cell reference (e.g., "A1")
---@field value any Cell value
---@field value_type string? Cell type ("n", "s", "b", "str", "inlineStr", "e")
---@field formula string? Formula string (without =)
---@field style_index integer? Style index

---@class RowData
---@field row_num integer Row number
---@field height number? Custom height
---@field hidden boolean? Is row hidden
---@field cells table<integer, CellData> Cells by column number

---@class ColumnData
---@field min integer Start column
---@field max integer End column
---@field width number Column width
---@field hidden boolean? Is column hidden

---@class WorksheetData
---@field dimension string? Sheet dimension (e.g., "A1:D10")
---@field rows table<integer, RowData> Rows by row number
---@field columns ColumnData[] Column definitions
---@field merged_cells string[] Merged cell ranges
---@field cells table<integer, table<integer, CellData>> Quick cell access [row][col]

--- Parse a cell reference to row and column
--- @param ref string Cell reference (e.g., "A1")
--- @return integer row, integer col
local function parse_cell_ref(ref)
  local parsed = column_utils.parse_ref(ref)
  return parsed.row, parsed.col
end

--- Parse cell value based on type
--- @param cell_xml table Parsed cell element
--- @param shared_strings table? Shared strings array
--- @return any value, string? value_type, string? formula
local function parse_cell_value(cell_xml, shared_strings)
  local cell_type = cell_xml.attrs.t
  local value = nil
  local formula = nil

  -- Extract value text
  local v_text = cell_xml.text and cell_xml.text:match("<v>([^<]*)</v>")

  -- Extract formula
  local f_text = cell_xml.text and cell_xml.text:match("<f>([^<]*)</f>")
  if f_text then
    formula = parser.unescape(f_text)
  end

  -- Parse value based on type
  if cell_type == "s" then
    -- Shared string reference
    local index = tonumber(v_text)
    if index and shared_strings then
      value = shared_strings:get(index)
    else
      value = v_text
    end
  elseif cell_type == "b" then
    -- Boolean
    value = v_text == "1"
  elseif cell_type == "e" then
    -- Error
    value = v_text  -- Keep error code as string
  elseif cell_type == "str" then
    -- Formula result string
    value = v_text and parser.unescape(v_text)
  elseif cell_type == "inlineStr" then
    -- Inline string: <is><t>text</t></is>
    local inline = cell_xml.text and cell_xml.text:match("<is>.-<t>([^<]*)</t>.-</is>")
    if inline then
      value = parser.unescape(inline)
    else
      -- Try simpler pattern
      inline = cell_xml.text and cell_xml.text:match("<t>([^<]*)</t>")
      value = inline and parser.unescape(inline)
    end
  else
    -- Number (default) or no type
    if v_text then
      value = tonumber(v_text)
      if not value then
        -- Fallback to string if not a valid number
        value = parser.unescape(v_text)
      end
    end
  end

  return value, cell_type, formula
end

--- Parse worksheet XML content
--- @param xml_content string The XML content of the worksheet
--- @param shared_strings table? SharedStrings instance for resolving string refs
--- @return WorksheetData
function M.parse(xml_content, shared_strings)
  local data = {
    dimension = nil,
    rows = {},
    columns = {},
    merged_cells = {},
    cells = {},
  }

  -- Parse dimension
  local dim_match = xml_content:match('<dimension ref="([^"]+)"')
  if dim_match then
    data.dimension = dim_match
  end

  -- Parse column definitions <col>
  local cols_section = xml_content:match("<cols>(.-)</cols>")
  if cols_section then
    local col_elements = parser.find_all(cols_section, "col")
    for _, col in ipairs(col_elements) do
      local col_data = {
        min = tonumber(col.attrs.min) or 1,
        max = tonumber(col.attrs.max) or 1,
        width = tonumber(col.attrs.width),
        hidden = col.attrs.hidden == "1",
      }
      table.insert(data.columns, col_data)
    end
  end

  -- Parse merged cells
  local merge_section = xml_content:match("<mergeCells[^>]*>(.-)</mergeCells>")
  if merge_section then
    for ref in merge_section:gmatch('<mergeCell ref="([^"]+)"') do
      table.insert(data.merged_cells, ref)
    end
  end

  -- Parse sheet data (rows and cells)
  local sheet_data = xml_content:match("<sheetData>(.-)</sheetData>")
  if sheet_data then
    -- Parse each row
    local row_elements = parser.find_all(sheet_data, "row")
    for _, row_elem in ipairs(row_elements) do
      local row_num = tonumber(row_elem.attrs.r)
      if row_num then
        local row_data = {
          row_num = row_num,
          height = tonumber(row_elem.attrs.ht),
          hidden = row_elem.attrs.hidden == "1",
          cells = {},
        }

        -- Parse cells in this row
        if row_elem.text then
          local cell_elements = parser.find_all(row_elem.text, "c")
          for _, cell_elem in ipairs(cell_elements) do
            local ref = cell_elem.attrs.r
            if ref then
              local r, c = parse_cell_ref(ref)
              local value, value_type, formula = parse_cell_value(cell_elem, shared_strings)

              local cell_data = {
                row = r,
                col = c,
                ref = ref,
                value = value,
                value_type = value_type,
                formula = formula,
                style_index = tonumber(cell_elem.attrs.s),
              }

              row_data.cells[c] = cell_data

              -- Also add to quick access table
              if not data.cells[r] then
                data.cells[r] = {}
              end
              data.cells[r][c] = cell_data
            end
          end
        end

        data.rows[row_num] = row_data
      end
    end
  end

  return data
end

--- Get cell data by row and column
--- @param data WorksheetData Parsed worksheet data
--- @param row integer Row number (1-indexed)
--- @param col integer Column number (1-indexed)
--- @return CellData?
function M.get_cell(data, row, col)
  if data.cells[row] then
    return data.cells[row][col]
  end
  return nil
end

--- Get cell value by row and column
--- @param data WorksheetData Parsed worksheet data
--- @param row integer Row number (1-indexed)
--- @param col integer Column number (1-indexed)
--- @return any? value
function M.get_value(data, row, col)
  local cell = M.get_cell(data, row, col)
  return cell and cell.value
end

--- Get all non-empty cells as a flat list
--- @param data WorksheetData Parsed worksheet data
--- @return CellData[]
function M.get_all_cells(data)
  local cells = {}
  for _, row in pairs(data.cells) do
    for _, cell in pairs(row) do
      table.insert(cells, cell)
    end
  end
  return cells
end

--- Get dimension bounds
--- @param data WorksheetData Parsed worksheet data
--- @return integer min_row, integer min_col, integer max_row, integer max_col
function M.get_bounds(data)
  local min_row, min_col = math.huge, math.huge
  local max_row, max_col = 0, 0

  for row_num, row in pairs(data.cells) do
    for col_num, _ in pairs(row) do
      min_row = math.min(min_row, row_num)
      max_row = math.max(max_row, row_num)
      min_col = math.min(min_col, col_num)
      max_col = math.max(max_col, col_num)
    end
  end

  if min_row == math.huge then
    return 1, 1, 1, 1  -- Empty sheet
  end

  return min_row, min_col, max_row, max_col
end

return M
