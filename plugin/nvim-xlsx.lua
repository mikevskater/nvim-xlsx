--- nvim-xlsx plugin initialization
--- Sets up user commands for xlsx file operations

if vim.g.loaded_nvim_xlsx then
  return
end
vim.g.loaded_nvim_xlsx = true

-- Lazy load the module
local function get_xlsx()
  return require("nvim-xlsx")
end

-- Command: Export current buffer or range to xlsx
vim.api.nvim_create_user_command("XlsxExport", function(opts)
  local xlsx = get_xlsx()
  local filepath = opts.args

  if filepath == "" then
    vim.notify("Usage: :XlsxExport <filepath>", vim.log.levels.ERROR)
    return
  end

  -- Get lines from buffer (range or entire buffer)
  local lines
  if opts.range == 2 then
    lines = vim.api.nvim_buf_get_lines(0, opts.line1 - 1, opts.line2, false)
  else
    lines = vim.api.nvim_buf_get_lines(0, 0, -1, false)
  end

  -- Parse lines as tab or comma separated values
  local data = {}
  for _, line in ipairs(lines) do
    local row = {}
    -- Try tab first, then comma
    local sep = line:find("\t") and "\t" or ","
    for value in (line .. sep):gmatch("([^" .. sep .. "]*)" .. sep) do
      -- Try to convert to number
      local num = tonumber(value)
      table.insert(row, num or value)
    end
    if #row > 0 then
      table.insert(data, row)
    end
  end

  if #data == 0 then
    vim.notify("No data to export", vim.log.levels.WARN)
    return
  end

  local ok, err = xlsx.export_table(data, filepath)
  if ok then
    vim.notify("Exported to " .. filepath, vim.log.levels.INFO)
  else
    vim.notify("Export failed: " .. (err or "unknown error"), vim.log.levels.ERROR)
  end
end, {
  nargs = 1,
  range = true,
  complete = "file",
  desc = "Export buffer/range to xlsx file",
})

-- Command: Import xlsx to new buffer
vim.api.nvim_create_user_command("XlsxImport", function(opts)
  local xlsx = get_xlsx()
  local filepath = opts.args

  if filepath == "" then
    vim.notify("Usage: :XlsxImport <filepath>", vim.log.levels.ERROR)
    return
  end

  local data, err = xlsx.import_table(filepath)
  if not data then
    vim.notify("Import failed: " .. (err or "unknown error"), vim.log.levels.ERROR)
    return
  end

  -- Create new buffer with the data
  vim.cmd("enew")
  local lines = {}
  for _, row in ipairs(data) do
    local cells = {}
    for _, value in ipairs(row) do
      table.insert(cells, tostring(value or ""))
    end
    table.insert(lines, table.concat(cells, "\t"))
  end
  vim.api.nvim_buf_set_lines(0, 0, -1, false, lines)
  vim.notify("Imported " .. #data .. " rows from " .. filepath, vim.log.levels.INFO)
end, {
  nargs = 1,
  complete = "file",
  desc = "Import xlsx file to new buffer",
})

-- Command: Get info about xlsx file
vim.api.nvim_create_user_command("XlsxInfo", function(opts)
  local xlsx = get_xlsx()
  local filepath = opts.args

  if filepath == "" then
    vim.notify("Usage: :XlsxInfo <filepath>", vim.log.levels.ERROR)
    return
  end

  local info, err = xlsx.info(filepath)
  if not info then
    vim.notify("Failed to read file: " .. (err or "unknown error"), vim.log.levels.ERROR)
    return
  end

  local output = { "Sheets: " .. info.sheet_count }
  for _, sheet in ipairs(info.sheets) do
    table.insert(output, string.format("  %d. %s (%s)", sheet.index, sheet.name, sheet.dimension))
  end
  vim.notify(table.concat(output, "\n"), vim.log.levels.INFO)
end, {
  nargs = 1,
  complete = "file",
  desc = "Show xlsx file info",
})

-- Command: Convert CSV to xlsx
vim.api.nvim_create_user_command("XlsxFromCsv", function(opts)
  local xlsx = get_xlsx()
  local args = vim.split(opts.args, " ")

  if #args < 2 then
    vim.notify("Usage: :XlsxFromCsv <csv_file> <xlsx_file>", vim.log.levels.ERROR)
    return
  end

  local csv_path = args[1]
  local xlsx_path = args[2]

  local ok, err = xlsx.from_csv_file(csv_path, xlsx_path)
  if ok then
    vim.notify("Converted " .. csv_path .. " to " .. xlsx_path, vim.log.levels.INFO)
  else
    vim.notify("Conversion failed: " .. (err or "unknown error"), vim.log.levels.ERROR)
  end
end, {
  nargs = "+",
  complete = "file",
  desc = "Convert CSV file to xlsx",
})

-- Command: Convert xlsx to CSV
vim.api.nvim_create_user_command("XlsxToCsv", function(opts)
  local xlsx = get_xlsx()
  local args = vim.split(opts.args, " ")

  if #args < 1 then
    vim.notify("Usage: :XlsxToCsv <xlsx_file> [output_file]", vim.log.levels.ERROR)
    return
  end

  local xlsx_path = args[1]
  local output_path = args[2]

  local csv, err = xlsx.to_csv(xlsx_path)
  if not csv then
    vim.notify("Conversion failed: " .. (err or "unknown error"), vim.log.levels.ERROR)
    return
  end

  if output_path then
    local file = io.open(output_path, "w")
    if file then
      file:write(csv)
      file:close()
      vim.notify("Wrote CSV to " .. output_path, vim.log.levels.INFO)
    else
      vim.notify("Failed to write to " .. output_path, vim.log.levels.ERROR)
    end
  else
    -- Open in new buffer
    vim.cmd("enew")
    local lines = vim.split(csv, "\n")
    vim.api.nvim_buf_set_lines(0, 0, -1, false, lines)
    vim.bo.filetype = "csv"
    vim.notify("Converted xlsx to CSV in new buffer", vim.log.levels.INFO)
  end
end, {
  nargs = "+",
  complete = "file",
  desc = "Convert xlsx to CSV",
})
