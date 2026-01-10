--- nvim-xlsx: Pure Lua xlsx library for Neovim
--- @module xlsx

local Workbook = require("xlsx.workbook")

local M = {}

M._VERSION = "0.2.0"

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

--- Export submodules for advanced usage
M.Workbook = Workbook.Workbook
M.Worksheet = require("xlsx.worksheet").Worksheet
M.Cell = require("xlsx.cell").Cell
M.Style = require("xlsx.style")
M.xml = require("xlsx.xml")
M.utils = require("xlsx.utils")
M.zip = require("xlsx.zip")

-- Style constants for convenience
M.BORDER_STYLES = M.Style.BORDER_STYLES
M.HALIGN = M.Style.HALIGN
M.VALIGN = M.Style.VALIGN
M.UNDERLINE = M.Style.UNDERLINE
M.BUILTIN_FORMATS = M.Style.BUILTIN_FORMATS

return M
