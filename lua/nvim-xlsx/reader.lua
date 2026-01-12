--- XLSX file reader
--- @module nvim-xlsx.reader
---
--- Coordinates reading and parsing of xlsx files.

local zip = require("nvim-xlsx.zip")
local shared_strings_mod = require("nvim-xlsx.parts.shared_strings")
local workbook_part = require("nvim-xlsx.parts.workbook_part")
local worksheet_part = require("nvim-xlsx.parts.worksheet_part")
local styles_part = require("nvim-xlsx.parts.styles_part")

local M = {}

---@class ReadWorkbook
---@field filepath string Original file path
---@field temp_dir string? Temporary directory (if still open)
---@field workbook_info table Parsed workbook info
---@field shared_strings table SharedStrings instance
---@field styles table? Parsed styles
---@field worksheets table<string, table> Parsed worksheet data by name
---@field _worksheet_cache table<string, table> Worksheet instances cache

--- Read and parse an xlsx file
--- @param filepath string Path to the xlsx file
--- @return table? workbook Parsed workbook data, or nil on error
--- @return string? error Error message if failed
function M.read(filepath)
  -- Check file exists
  if vim.fn.filereadable(filepath) == 0 then
    return nil, "File not found: " .. filepath
  end

  -- Create temp directory for extraction
  local temp_dir = zip.create_temp_dir()

  -- Extract xlsx
  local ok, err = zip.unzip_file(filepath, temp_dir)
  if not ok then
    zip.cleanup_temp_dir(temp_dir)
    return nil, "Failed to extract xlsx: " .. (err or "unknown error")
  end

  -- Normalize path separators for cross-platform compatibility
  local sep = package.config:sub(1, 1)  -- "/" on Unix, "\" on Windows
  local function path_join(...)
    return table.concat({ ... }, sep)
  end

  local workbook = {
    filepath = filepath,
    temp_dir = temp_dir,
    workbook_info = nil,
    shared_strings = shared_strings_mod.new(),
    styles = nil,
    worksheets = {},
    _worksheet_cache = {},
  }

  -- Parse shared strings (optional file)
  local ss_path = path_join(temp_dir, "xl", "sharedStrings.xml")
  local ss_content = zip.read_file(ss_path)
  if ss_content then
    workbook.shared_strings = shared_strings_mod.parse(ss_content)
  end

  -- Parse styles
  local styles_path = path_join(temp_dir, "xl", "styles.xml")
  local styles_content = zip.read_file(styles_path)
  if styles_content then
    workbook.styles = styles_part.parse(styles_content)
  end

  -- Parse workbook relationships
  local wb_rels_path = path_join(temp_dir, "xl", "_rels", "workbook.xml.rels")
  local wb_rels_content = zip.read_file(wb_rels_path)

  -- Parse workbook
  local wb_path = path_join(temp_dir, "xl", "workbook.xml")
  local wb_content, wb_err = zip.read_file(wb_path)
  if not wb_content then
    zip.cleanup_temp_dir(temp_dir)
    return nil, "Failed to read workbook.xml: " .. (wb_err or "unknown error")
  end

  workbook.workbook_info = workbook_part.parse(wb_content, wb_rels_content)

  -- Parse each worksheet
  for _, sheet_info in ipairs(workbook.workbook_info.sheets) do
    local sheet_target = sheet_info.target
    if sheet_target then
      -- Target is relative to xl/ directory
      -- Replace forward slashes in target with platform separator
      local normalized_target = sheet_target:gsub("/", sep)
      local sheet_path = path_join(temp_dir, "xl", normalized_target)
      local sheet_content, read_err = zip.read_file(sheet_path)
      if sheet_content then
        local sheet_data = worksheet_part.parse(sheet_content, workbook.shared_strings)
        sheet_data.name = sheet_info.name
        sheet_data.sheet_id = sheet_info.sheet_id
        workbook.worksheets[sheet_info.name] = sheet_data
      end
    end
  end

  -- Clean up temp directory
  zip.cleanup_temp_dir(temp_dir)
  workbook.temp_dir = nil

  return workbook
end

--- Get list of sheet names
--- @param workbook table Parsed workbook
--- @return string[] Sheet names in order
function M.get_sheet_names(workbook)
  local names = {}
  for _, sheet_info in ipairs(workbook.workbook_info.sheets) do
    table.insert(names, sheet_info.name)
  end
  return names
end

--- Get worksheet data by name
--- @param workbook table Parsed workbook
--- @param name string Sheet name
--- @return table? Worksheet data or nil
function M.get_sheet(workbook, name)
  return workbook.worksheets[name]
end

--- Get worksheet data by index (1-based)
--- @param workbook table Parsed workbook
--- @param index integer Sheet index
--- @return table? Worksheet data or nil
function M.get_sheet_by_index(workbook, index)
  local sheet_info = workbook.workbook_info.sheets[index]
  if sheet_info then
    return workbook.worksheets[sheet_info.name]
  end
  return nil
end

--- Get cell value from a worksheet
--- @param sheet_data table Parsed worksheet data
--- @param row integer Row number (1-indexed)
--- @param col integer Column number (1-indexed)
--- @return any? Cell value
function M.get_cell(sheet_data, row, col)
  return worksheet_part.get_value(sheet_data, row, col)
end

--- Get cell data (full info) from a worksheet
--- @param sheet_data table Parsed worksheet data
--- @param row integer Row number (1-indexed)
--- @param col integer Column number (1-indexed)
--- @return table? Cell data with value, formula, style, etc.
function M.get_cell_data(sheet_data, row, col)
  return worksheet_part.get_cell(sheet_data, row, col)
end

--- Get a range of values as a 2D array
--- @param sheet_data table Parsed worksheet data
--- @param r1 integer Start row
--- @param c1 integer Start column
--- @param r2 integer End row
--- @param c2 integer End column
--- @return any[][] 2D array of values
function M.get_range(sheet_data, r1, c1, r2, c2)
  local result = {}
  for row = r1, r2 do
    local row_data = {}
    for col = c1, c2 do
      table.insert(row_data, worksheet_part.get_value(sheet_data, row, col))
    end
    table.insert(result, row_data)
  end
  return result
end

--- Get all data as a 2D array
--- @param sheet_data table Parsed worksheet data
--- @return any[][] 2D array of values
--- @return integer min_row, integer min_col, integer max_row, integer max_col Bounds
function M.get_all_data(sheet_data)
  local min_row, min_col, max_row, max_col = worksheet_part.get_bounds(sheet_data)
  local data = M.get_range(sheet_data, min_row, min_col, max_row, max_col)
  return data, min_row, min_col, max_row, max_col
end

--- Get merged cell ranges
--- @param sheet_data table Parsed worksheet data
--- @return string[] Array of merge range references (e.g., "A1:B2")
function M.get_merged_cells(sheet_data)
  return sheet_data.merged_cells or {}
end

--- Get column widths
--- @param sheet_data table Parsed worksheet data
--- @return table<integer, number> Map of column number to width
function M.get_column_widths(sheet_data)
  local widths = {}
  for _, col in ipairs(sheet_data.columns or {}) do
    if col.width then
      for c = col.min, col.max do
        widths[c] = col.width
      end
    end
  end
  return widths
end

--- Get row heights
--- @param sheet_data table Parsed worksheet data
--- @return table<integer, number> Map of row number to height
function M.get_row_heights(sheet_data)
  local heights = {}
  for row_num, row in pairs(sheet_data.rows or {}) do
    if row.height then
      heights[row_num] = row.height
    end
  end
  return heights
end

return M
