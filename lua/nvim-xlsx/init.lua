--- nvim-xlsx: Pure Lua xlsx library for Neovim
--- @module nvim-xlsx

local Workbook = require("nvim-xlsx.workbook")
local reader = require("nvim-xlsx.reader")

local M = {}

M._VERSION = "0.7.0"

--- Create a new empty workbook
--- @return Workbook
function M.new_workbook()
  return Workbook.new()
end

--- Convenience function to export a 2D table to xlsx
--- @param data table[] Array of rows, each row is an array of values
--- @param filepath string Output file path
--- @param options? table Options: { sheet_name?: string, headers?: string[] }
--- @return boolean success
--- @return string? error_message
function M.export_table(data, filepath, options)
  options = options or {}

  local wb = M.new_workbook()
  local sheet, err = wb:add_sheet(options.sheet_name or "Sheet1")
  if not sheet then
    return false, err
  end

  local start_row = 1

  -- Write headers if provided
  if options.headers then
    for col, header in ipairs(options.headers) do
      sheet:set_cell(1, col, header)
    end
    start_row = 2
  end

  -- Write data
  for row_idx, row_data in ipairs(data) do
    for col_idx, value in ipairs(row_data) do
      sheet:set_cell(start_row + row_idx - 1, col_idx, value)
    end
  end

  return wb:save(filepath)
end

--- Open an existing xlsx file for reading
--- @param filepath string Path to the xlsx file
--- @return table? workbook Parsed workbook data, or nil on error
--- @return string? error Error message if failed
function M.open(filepath)
  return reader.read(filepath)
end

--- Get sheet names from an opened workbook
--- @param workbook table Opened workbook from xlsx.open()
--- @return string[] Sheet names
function M.get_sheet_names(workbook)
  return reader.get_sheet_names(workbook)
end

--- Get a worksheet by name from an opened workbook
--- @param workbook table Opened workbook from xlsx.open()
--- @param name string Sheet name
--- @return table? Worksheet data or nil
function M.get_sheet(workbook, name)
  return reader.get_sheet(workbook, name)
end

--- Get a worksheet by index from an opened workbook
--- @param workbook table Opened workbook from xlsx.open()
--- @param index integer Sheet index (1-based)
--- @return table? Worksheet data or nil
function M.get_sheet_by_index(workbook, index)
  return reader.get_sheet_by_index(workbook, index)
end

--- Get a cell value from a worksheet
--- @param sheet table Worksheet from get_sheet()
--- @param row integer Row number (1-indexed)
--- @param col integer Column number (1-indexed)
--- @return any? Cell value
function M.get_cell(sheet, row, col)
  return reader.get_cell(sheet, row, col)
end

--- Get a range of values as a 2D array
--- @param sheet table Worksheet from get_sheet()
--- @param r1 integer Start row
--- @param c1 integer Start column
--- @param r2 integer End row
--- @param c2 integer End column
--- @return any[][] 2D array of values
function M.get_range(sheet, r1, c1, r2, c2)
  return reader.get_range(sheet, r1, c1, r2, c2)
end

--- Import all data from an xlsx file as a table
--- @param filepath string Path to the xlsx file
--- @param options? table Options: { sheet_name?: string, sheet_index?: integer }
--- @return any[][]? data 2D array of values, or nil on error
--- @return string? error Error message if failed
function M.import_table(filepath, options)
  options = options or {}

  local wb, err = M.open(filepath)
  if not wb then
    return nil, err
  end

  local sheet
  if options.sheet_name then
    sheet = M.get_sheet(wb, options.sheet_name)
  elseif options.sheet_index then
    sheet = M.get_sheet_by_index(wb, options.sheet_index)
  else
    -- Default to first sheet
    sheet = M.get_sheet_by_index(wb, 1)
  end

  if not sheet then
    return nil, "Sheet not found"
  end

  local data = reader.get_all_data(sheet)
  return data
end

--- Get information about an xlsx file without fully loading it
--- @param filepath string Path to the xlsx file
--- @return table? info File info, or nil on error
--- @return string? error Error message if failed
function M.info(filepath)
  local wb, err = M.open(filepath)
  if not wb then
    return nil, err
  end

  local sheets = M.get_sheet_names(wb)
  local sheet_info = {}

  for i, name in ipairs(sheets) do
    local sheet = M.get_sheet_by_index(wb, i)
    local dim = sheet and sheet.dimension or "A1"
    table.insert(sheet_info, {
      index = i,
      name = name,
      dimension = dim,
    })
  end

  return {
    sheets = sheet_info,
    sheet_count = #sheets,
    properties = wb.properties or {},
  }
end

--- Create xlsx from a CSV string
--- @param csv_string string CSV content
--- @param filepath string Output xlsx path
--- @param options? table Options: { delimiter?: string, has_headers?: boolean, sheet_name?: string }
--- @return boolean success
--- @return string? error_message
function M.from_csv(csv_string, filepath, options)
  options = options or {}
  local delimiter = options.delimiter or ","

  local data = {}
  local headers = nil

  -- Simple CSV parsing (handles basic cases)
  for line in csv_string:gmatch("[^\r\n]+") do
    local row = {}
    -- Handle quoted fields
    local pos = 1
    while pos <= #line do
      local c = line:sub(pos, pos)
      if c == '"' then
        -- Quoted field
        local end_quote = line:find('"', pos + 1)
        while end_quote and line:sub(end_quote + 1, end_quote + 1) == '"' do
          -- Escaped quote
          end_quote = line:find('"', end_quote + 2)
        end
        if end_quote then
          local value = line:sub(pos + 1, end_quote - 1):gsub('""', '"')
          table.insert(row, value)
          pos = end_quote + 2  -- Skip closing quote and delimiter
        else
          -- Malformed, take rest of line
          table.insert(row, line:sub(pos + 1))
          break
        end
      else
        -- Unquoted field
        local next_delim = line:find(delimiter, pos, true)
        if next_delim then
          local value = line:sub(pos, next_delim - 1)
          -- Try to convert to number
          local num = tonumber(value)
          table.insert(row, num or value)
          pos = next_delim + 1
        else
          local value = line:sub(pos)
          local num = tonumber(value)
          table.insert(row, num or value)
          break
        end
      end
    end

    if options.has_headers and not headers then
      headers = row
    else
      table.insert(data, row)
    end
  end

  return M.export_table(data, filepath, {
    sheet_name = options.sheet_name or "Sheet1",
    headers = headers,
  })
end

--- Create xlsx from a CSV file
--- @param csv_path string Path to CSV file
--- @param xlsx_path string Output xlsx path
--- @param options? table Options: { delimiter?: string, has_headers?: boolean, sheet_name?: string }
--- @return boolean success
--- @return string? error_message
function M.from_csv_file(csv_path, xlsx_path, options)
  local file, err = io.open(csv_path, "r")
  if not file then
    return false, "Cannot open CSV file: " .. tostring(err)
  end

  local content = file:read("*a")
  file:close()

  return M.from_csv(content, xlsx_path, options)
end

--- Export xlsx to CSV string
--- @param filepath string Path to xlsx file
--- @param options? table Options: { sheet_name?: string, sheet_index?: integer, delimiter?: string }
--- @return string? csv CSV content, or nil on error
--- @return string? error Error message if failed
function M.to_csv(filepath, options)
  options = options or {}
  local delimiter = options.delimiter or ","

  local data, err = M.import_table(filepath, options)
  if not data then
    return nil, err
  end

  local lines = {}
  for _, row in ipairs(data) do
    local cells = {}
    for _, value in ipairs(row) do
      local str = tostring(value or "")
      -- Quote if contains delimiter, quote, or newline
      if str:find(delimiter, 1, true) or str:find('"') or str:find("\n") then
        str = '"' .. str:gsub('"', '""') .. '"'
      end
      table.insert(cells, str)
    end
    table.insert(lines, table.concat(cells, delimiter))
  end

  return table.concat(lines, "\n")
end

--- Export submodules for advanced usage
M.Workbook = Workbook.Workbook
M.Worksheet = require("nvim-xlsx.worksheet").Worksheet
M.Cell = require("nvim-xlsx.cell").Cell
M.Style = require("nvim-xlsx.style")
M.xml = require("nvim-xlsx.xml")
M.utils = require("nvim-xlsx.utils")
M.zip = require("nvim-xlsx.zip")
M.reader = reader

-- Style constants for convenience
M.BORDER_STYLES = M.Style.BORDER_STYLES
M.HALIGN = M.Style.HALIGN
M.VALIGN = M.Style.VALIGN
M.UNDERLINE = M.Style.UNDERLINE
M.BUILTIN_FORMATS = M.Style.BUILTIN_FORMATS

-- Date utilities for convenience
M.date = require("nvim-xlsx.utils.date")

-- Validation utilities for convenience
M.validation = require("nvim-xlsx.utils.validation")

-- Excel limits for reference
M.LIMITS = M.validation.LIMITS

return M
