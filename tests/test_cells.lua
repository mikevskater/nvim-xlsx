--- Tests for cell operations, values, formulas, dates, booleans
--- Run with: nvim --headless -l tests/test_cells.lua

dofile(arg[0]:match("(.*/)").. "test_helper.lua")
local h = require("tests.test_helper")
local xlsx = h.require_xlsx()

h.suite("Cell Tests")

local wb = xlsx.new_workbook()
local sheet = wb:add_sheet("Cells")

-- ============================================
h.section("Test 1: Basic cell operations")
-- ============================================

local cell = sheet:set_cell(1, 1, "Hello")
h.test("set_cell returns cell", cell ~= nil)
h.test("cell has correct row", cell.row == 1)
h.test("cell has correct col", cell.col == 1)
h.test("cell has correct value", cell.value == "Hello")

local retrieved = sheet:get_cell(1, 1)
h.test("get_cell returns same cell", retrieved == cell)

local missing = sheet:get_cell(99, 99)
h.test("get_cell for empty cell returns nil", missing == nil)

-- ============================================
h.section("Test 2: A1 notation")
-- ============================================

local cell_a1 = sheet:set("B2", "World")
h.test("set with A1 notation works", cell_a1 ~= nil)
h.test("A1 cell has correct row", cell_a1.row == 2)
h.test("A1 cell has correct col", cell_a1.col == 2)

local get_a1 = sheet:get("B2")
h.test("get with A1 notation works", get_a1 == cell_a1)

-- ============================================
h.section("Test 3: Cell value types")
-- ============================================

-- String
local str_cell = sheet:set_cell(3, 1, "Text value")
h.test("string value stored", str_cell.value == "Text value")
h.test("string type is inlineStr", str_cell.value_type == "inlineStr")

-- Number
local num_cell = sheet:set_cell(3, 2, 42.5)
h.test("number value stored", num_cell.value == 42.5)
h.test("number type is n", num_cell.value_type == "n")

-- Integer
local int_cell = sheet:set_cell(3, 3, 100)
h.test("integer value stored", int_cell.value == 100)

-- ============================================
h.section("Test 4: Boolean values")
-- ============================================

local bool_true = sheet:set_boolean(4, 1, true)
h.test("true boolean stored as 1", bool_true.value == 1)
h.test("boolean type is b", bool_true.value_type == "b")

local bool_false = sheet:set_boolean(4, 2, false)
h.test("false boolean stored as 0", bool_false.value == 0)

-- ============================================
h.section("Test 5: Formula values")
-- ============================================

local formula1 = sheet:set_formula(5, 1, "SUM(A1:A10)")
h.test("formula stored without =", formula1.formula == "SUM(A1:A10)")

local formula2 = sheet:set_formula(5, 2, "=AVERAGE(B1:B10)")
h.test("formula with = has = stripped", formula2.formula == "AVERAGE(B1:B10)")

-- Formula via set_cell
local formula3 = sheet:set_cell(5, 3, "=COUNT(C1:C10)")
h.test("formula via set_cell detected", formula3.formula == "COUNT(C1:C10)")

-- ============================================
h.section("Test 6: Date values")
-- ============================================

-- Date as table
local date_table = sheet:set_date(6, 1, { year = 2024, month = 6, day = 15 })
h.test("date from table is number", type(date_table.value) == "number")
h.test("date serial is correct", math.floor(date_table.value) == 45458)

-- Date as serial number
local date_serial = sheet:set_date(6, 2, 45000)
h.test("date from serial stored", date_serial.value == 45000)

-- Date utilities
local date_utils = xlsx.date
local serial = date_utils.to_serial({ year = 2024, month = 1, day = 1 })
h.test("date to_serial works", serial == 45292)

local parsed = date_utils.from_serial(45292)
h.test("date from_serial year", parsed.year == 2024)
h.test("date from_serial month", parsed.month == 1)
h.test("date from_serial day", parsed.day == 1)

-- ============================================
h.section("Test 7: set_cell_value with style")
-- ============================================

local style_idx = wb:create_style({ bold = true })
local styled_cell = sheet:set_cell_value(7, 1, "Styled", style_idx)
h.test("set_cell_value returns cell", styled_cell ~= nil)
h.test("style index applied", styled_cell.style_index == style_idx)

-- ============================================
h.section("Test 8: Column widths and row heights")
-- ============================================

sheet:set_column_width(1, 20)
h.test("column width set", sheet.column_widths[1] == 20)

sheet:set_row_height(1, 30)
h.test("row height set", sheet.row_heights[1] == 30)

-- ============================================
h.section("Test 9: Dimension tracking")
-- ============================================

local dim_sheet = wb:add_sheet("Dimensions")
h.test("empty sheet dimension is A1", dim_sheet:get_dimension() == "A1")

dim_sheet:set_cell(5, 3, "Data")
h.test("dimension updates for single cell", dim_sheet:get_dimension() == "C5:C5")

dim_sheet:set_cell(1, 1, "Start")
dim_sheet:set_cell(10, 10, "End")
h.test("dimension spans full range", dim_sheet:get_dimension() == "A1:J10")

-- ============================================
h.section("Test 10: Cell bounds validation")
-- ============================================

local function throws(fn)
  local ok = pcall(fn)
  return not ok
end

h.test("row 0 throws error", throws(function() sheet:set_cell(0, 1, "x") end))
h.test("col 0 throws error", throws(function() sheet:set_cell(1, 0, "x") end))
h.test("negative row throws error", throws(function() sheet:set_cell(-1, 1, "x") end))
h.test("row > MAX_ROWS throws", throws(function() sheet:set_cell(1048577, 1, "x") end))
h.test("col > MAX_COLS throws", throws(function() sheet:set_cell(1, 16385, "x") end))

-- ============================================
h.section("Test 11: Cell XML generation")
-- ============================================

local xml_cell = sheet:set_cell(11, 1, "XML Test")
local xml = xml_cell:to_xml()
h.test("cell XML contains reference", xml:match('r="A11"') ~= nil)
h.test("cell XML contains value", xml:match("XML Test") ~= nil)

local formula_cell = sheet:set_formula(11, 2, "A1+B1")
local formula_xml = formula_cell:to_xml()
h.test("formula XML contains <f>", formula_xml:match("<f>") ~= nil)

-- ============================================
h.section("Test 12: Save and verify")
-- ============================================

local output_path = h.fixtures_dir .. "/test_cells.xlsx"
local ok, err = wb:save(output_path)
h.test("save succeeds", ok == true)

h.summary("Cell Tests")
