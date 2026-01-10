--- Tests for reading xlsx files
--- Run with: nvim --headless -l tests/test_reader.lua

dofile(arg[0]:match("(.*/)").. "test_helper.lua")
local h = require("tests.test_helper")
local xlsx = h.require_xlsx()

h.suite("Reader Tests")

-- First, create a test file to read
local function create_test_file()
  local wb = xlsx.new_workbook()

  local sheet1 = wb:add_sheet("Data")
  sheet1:set_cell(1, 1, "Name")
  sheet1:set_cell(1, 2, "Value")
  sheet1:set_cell(2, 1, "Item A")
  sheet1:set_cell(2, 2, 100)
  sheet1:set_cell(3, 1, "Item B")
  sheet1:set_cell(3, 2, 200)
  sheet1:set_formula(4, 2, "SUM(B2:B3)")
  sheet1:set_boolean(5, 1, true)
  sheet1:set_boolean(5, 2, false)
  sheet1:merge_cells(6, 1, 6, 2)
  sheet1:set_cell(6, 1, "Merged")
  sheet1:set_column_width(1, 15)

  local sheet2 = wb:add_sheet("Styled")
  local style = wb:create_style({ bold = true, bg_color = "#FFFF00" })
  sheet2:set_cell(1, 1, "Styled Cell")
  sheet2:set_cell_style(1, 1, style)

  local path = h.fixtures_dir .. "/test_reader_source.xlsx"
  wb:save(path)
  return path
end

local test_file = create_test_file()

-- ============================================
h.section("Test 1: Open file")
-- ============================================

local wb, err = xlsx.open(test_file)
h.test("file opens successfully", wb ~= nil)
h.test("no error returned", err == nil)

-- ============================================
h.section("Test 2: Sheet enumeration")
-- ============================================

local sheet_names = xlsx.get_sheet_names(wb)
h.test("get_sheet_names works", sheet_names ~= nil)
h.test("has 2 sheets", #sheet_names == 2)
h.test("first sheet is Data", sheet_names[1] == "Data")
h.test("second sheet is Styled", sheet_names[2] == "Styled")

-- ============================================
h.section("Test 3: Get sheet by name")
-- ============================================

local data_sheet = xlsx.get_sheet(wb, "Data")
h.test("get_sheet by name works", data_sheet ~= nil)
h.test("correct sheet returned", data_sheet.name == "Data")

-- ============================================
h.section("Test 4: Get sheet by index")
-- ============================================

local first_sheet = xlsx.get_sheet_by_index(wb, 1)
h.test("get_sheet_by_index works", first_sheet ~= nil)
h.test("correct sheet returned", first_sheet.name == "Data")

-- ============================================
h.section("Test 5: Read cell values")
-- ============================================

local cell_a1 = xlsx.get_cell(data_sheet, 1, 1)
h.test("cell A1 value is Name", cell_a1 == "Name")

local cell_b2 = xlsx.get_cell(data_sheet, 2, 2)
h.test("cell B2 value is 100", cell_b2 == 100)

-- ============================================
h.section("Test 6: Read range of values")
-- ============================================

local range = xlsx.get_range(data_sheet, 1, 1, 3, 2)
h.test("range has 3 rows", #range == 3)
h.test("range row 1 has 2 cols", #range[1] == 2)
h.test("range[1][1] is Name", range[1][1] == "Name")
h.test("range[2][2] is 100", range[2][2] == 100)

-- ============================================
h.section("Test 7: Read using reader module directly")
-- ============================================

local reader_range = xlsx.reader.get_range(data_sheet, 1, 1, 3, 2)
h.test("reader.get_range works", reader_range ~= nil and #reader_range == 3)

-- ============================================
h.section("Test 8: Read merged cells")
-- ============================================

local merged = xlsx.reader.get_merged_cells(data_sheet)
h.test("merged_cells returned", merged ~= nil)
h.test("has merged regions", #merged > 0)
h.test("merge reference correct", merged[1] == "A6:B6")

-- ============================================
h.section("Test 9: Read column widths")
-- ============================================

local widths = xlsx.reader.get_column_widths(data_sheet)
h.test("column_widths returned", widths ~= nil)
h.test("column 1 has width", widths[1] ~= nil)
h.test("column width approximately correct", math.abs(widths[1] - 15) < 1)

-- ============================================
h.section("Test 10: Round-trip test")
-- ============================================

-- Save the read workbook using new workbook (recreate data)
local wb_new = xlsx.new_workbook()
local sheet_new = wb_new:add_sheet("Data")

-- Copy data from range
for r = 1, 3 do
  for c = 1, 2 do
    local val = xlsx.get_cell(data_sheet, r, c)
    sheet_new:set_cell(r, c, val)
  end
end

local roundtrip_path = h.fixtures_dir .. "/test_reader_roundtrip.xlsx"
local save_ok = wb_new:save(roundtrip_path)
h.test("roundtrip save succeeds", save_ok)

-- Read it again
local wb2, err2 = xlsx.open(roundtrip_path)
h.test("roundtrip file opens", wb2 ~= nil)
local rt_sheet = xlsx.get_sheet(wb2, "Data")
h.test("roundtrip sheet exists", rt_sheet ~= nil)
local rt_cell = xlsx.get_cell(rt_sheet, 2, 2)
h.test("roundtrip value preserved", rt_cell == 100)

-- ============================================
h.section("Test 11: Non-existent file")
-- ============================================

local bad_wb, bad_err = xlsx.open("/nonexistent/file.xlsx")
h.test("non-existent file returns nil", bad_wb == nil)
h.test("error message returned", bad_err ~= nil)

-- ============================================
h.section("Test 12: Import table convenience")
-- ============================================

local imported, imp_err = xlsx.import_table(test_file)
h.test("import_table works", imported ~= nil)
h.test("no import error", imp_err == nil)
h.test("imported has data", #imported > 0)

h.summary("Reader Tests")
