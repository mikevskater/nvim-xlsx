--- Performance benchmark suite for nvim-xlsx
--- Run with: nvim --headless -l tests/benchmark.lua

-- Setup package path
local script_path = debug.getinfo(1, "S").source:sub(2)
script_path = script_path:gsub("\\", "/")
local base_dir = script_path:match("(.+)/tests/")
if not base_dir then
  base_dir = vim.fn.fnamemodify(script_path, ":p:h:h")
  base_dir = base_dir:gsub("\\", "/")
end
package.path = base_dir .. "/lua/?.lua;" .. base_dir .. "/lua/?/init.lua;" .. package.path

local xlsx = require("xlsx")

print("=== nvim-xlsx Performance Benchmark ===\n")

-- Utility function to measure time
local function measure(name, fn)
  collectgarbage("collect")
  local start_mem = collectgarbage("count")
  local start_time = os.clock()

  local result = fn()

  local end_time = os.clock()
  local end_mem = collectgarbage("count")

  local elapsed = (end_time - start_time) * 1000
  local mem_used = end_mem - start_mem

  print(string.format("  %-40s %8.2f ms  %8.1f KB", name, elapsed, mem_used))

  return result, elapsed, mem_used
end

-- ============================================
-- Benchmark 1: Create workbook with many cells
-- ============================================
print("Benchmark 1: Cell Writing Performance")
print(string.rep("-", 70))

local function benchmark_write(rows, cols)
  local wb = xlsx.new_workbook()
  local sheet = wb:add_sheet("Data")

  for r = 1, rows do
    for c = 1, cols do
      sheet:set_cell(r, c, r * c)
    end
  end

  return wb
end

measure("100 rows x 10 cols (1,000 cells)", function()
  return benchmark_write(100, 10)
end)

measure("1,000 rows x 10 cols (10,000 cells)", function()
  return benchmark_write(1000, 10)
end)

measure("1,000 rows x 50 cols (50,000 cells)", function()
  return benchmark_write(1000, 50)
end)

measure("5,000 rows x 20 cols (100,000 cells)", function()
  return benchmark_write(5000, 20)
end)

print("")

-- ============================================
-- Benchmark 2: Save workbook to file
-- ============================================
print("Benchmark 2: File Save Performance")
print(string.rep("-", 70))

local output_dir = base_dir .. "/tests/fixtures"

local small_wb = benchmark_write(100, 10)
local medium_wb = benchmark_write(1000, 10)
local large_wb = benchmark_write(1000, 50)

measure("Save 1,000 cells", function()
  return small_wb:save(output_dir .. "/bench_small.xlsx")
end)

measure("Save 10,000 cells", function()
  return medium_wb:save(output_dir .. "/bench_medium.xlsx")
end)

measure("Save 50,000 cells", function()
  return large_wb:save(output_dir .. "/bench_large.xlsx")
end)

print("")

-- ============================================
-- Benchmark 3: Read workbook from file
-- ============================================
print("Benchmark 3: File Read Performance")
print(string.rep("-", 70))

local function read_all_cells(wb)
  local total = 0
  local names = xlsx.get_sheet_names(wb)
  for _, name in ipairs(names) do
    local sheet = xlsx.get_sheet(wb, name)
    if sheet and sheet.dimension then
      local dim = sheet.dimension
      -- Count cells in dimension
      local parsed = xlsx.utils.column.parse_range(dim)
      for r = parsed.start.row, parsed.finish.row do
        for c = parsed.start.col, parsed.finish.col do
          local val = xlsx.get_cell(sheet, r, c)
          if val then total = total + 1 end
        end
      end
    end
  end
  return total
end

measure("Read 1,000 cells", function()
  local wb = xlsx.open(output_dir .. "/bench_small.xlsx")
  return read_all_cells(wb)
end)

measure("Read 10,000 cells", function()
  local wb = xlsx.open(output_dir .. "/bench_medium.xlsx")
  return read_all_cells(wb)
end)

measure("Read 50,000 cells", function()
  local wb = xlsx.open(output_dir .. "/bench_large.xlsx")
  return read_all_cells(wb)
end)

print("")

-- ============================================
-- Benchmark 4: String handling
-- ============================================
print("Benchmark 4: String Cell Performance")
print(string.rep("-", 70))

local function benchmark_strings(rows, avg_length)
  local wb = xlsx.new_workbook()
  local sheet = wb:add_sheet("Strings")

  for r = 1, rows do
    local text = string.rep("x", avg_length) .. tostring(r)
    sheet:set_cell(r, 1, text)
  end

  return wb
end

measure("1,000 short strings (10 chars)", function()
  return benchmark_strings(1000, 10)
end)

measure("1,000 medium strings (100 chars)", function()
  return benchmark_strings(1000, 100)
end)

measure("1,000 long strings (1000 chars)", function()
  return benchmark_strings(1000, 1000)
end)

print("")

-- ============================================
-- Benchmark 5: Styled cells
-- ============================================
print("Benchmark 5: Styled Cell Performance")
print(string.rep("-", 70))

local function benchmark_styles(rows)
  local wb = xlsx.new_workbook()
  local sheet = wb:add_sheet("Styled")

  -- Create some styles
  local bold = wb:create_style({ font = { bold = true } })
  local red = wb:create_style({ font = { color = "#FF0000" } })
  local bg = wb:create_style({ fill = { color = "#FFFF00" } })

  for r = 1, rows do
    sheet:set_cell(r, 1, "Bold")
    sheet:set_cell_style(r, 1, bold)
    sheet:set_cell(r, 2, "Red")
    sheet:set_cell_style(r, 2, red)
    sheet:set_cell(r, 3, "Yellow BG")
    sheet:set_cell_style(r, 3, bg)
  end

  return wb
end

measure("100 styled rows", function()
  return benchmark_styles(100)
end)

measure("1,000 styled rows", function()
  return benchmark_styles(1000)
end)

measure("5,000 styled rows", function()
  return benchmark_styles(5000)
end)

print("")

-- ============================================
-- Benchmark 6: XML generation
-- ============================================
print("Benchmark 6: XML Generation Performance")
print(string.rep("-", 70))

local xml_wb = benchmark_write(1000, 20)
local xml_sheet = xml_wb:get_sheet(1)

measure("Generate XML for 20,000 cells", function()
  return xml_sheet:to_xml(true)
end)

local styled_wb = benchmark_styles(1000)

measure("Generate styled XML for 3,000 cells", function()
  local sheet = styled_wb:get_sheet(1)
  return sheet:to_xml(true)
end)

print("")

-- ============================================
-- Benchmark 7: Formulas
-- ============================================
print("Benchmark 7: Formula Cell Performance")
print(string.rep("-", 70))

local function benchmark_formulas(rows)
  local wb = xlsx.new_workbook()
  local sheet = wb:add_sheet("Formulas")

  for r = 1, rows do
    sheet:set_cell(r, 1, r)
    sheet:set_cell(r, 2, r * 2)
    sheet:set_formula(r, 3, "=A" .. r .. "+B" .. r)
  end

  return wb
end

measure("1,000 formula cells", function()
  return benchmark_formulas(1000)
end)

measure("5,000 formula cells", function()
  return benchmark_formulas(5000)
end)

print("")

-- ============================================
-- Summary
-- ============================================
print("=== Benchmark Complete ===")
print("\nNote: Times may vary based on system load and disk speed.")
print("Memory values are approximate due to Lua GC behavior.")

-- Cleanup benchmark files
os.remove(output_dir .. "/bench_small.xlsx")
os.remove(output_dir .. "/bench_medium.xlsx")
os.remove(output_dir .. "/bench_large.xlsx")
