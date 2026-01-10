--- nvim-xlsx: Pure Lua xlsx library for Neovim
--- @module xlsx

local Workbook = require("xlsx.workbook")
local reader = require("xlsx.reader")

local M = {}

M._VERSION = "0.5.0"

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

--- Export submodules for advanced usage
M.Workbook = Workbook.Workbook
M.Worksheet = require("xlsx.worksheet").Worksheet
M.Cell = require("xlsx.cell").Cell
M.Style = require("xlsx.style")
M.xml = require("xlsx.xml")
M.utils = require("xlsx.utils")
M.zip = require("xlsx.zip")
M.reader = reader

-- Style constants for convenience
M.BORDER_STYLES = M.Style.BORDER_STYLES
M.HALIGN = M.Style.HALIGN
M.VALIGN = M.Style.VALIGN
M.UNDERLINE = M.Style.UNDERLINE
M.BUILTIN_FORMATS = M.Style.BUILTIN_FORMATS

-- Date utilities for convenience
M.date = require("xlsx.utils.date")

return M
