--- Tests for print settings
--- Run with: nvim --headless -l tests/test_print.lua

dofile(arg[0]:match("(.*/)").. "test_helper.lua")
local h = require("tests.test_helper")
local xlsx = h.require_xlsx()

h.suite("Print Settings Tests")

local wb = xlsx.new_workbook()

-- ============================================
h.section("Test 1: Set orientation")
-- ============================================

local sheet1 = wb:add_sheet("Orientation")
local ok = sheet1:set_orientation("landscape")
h.test("set_orientation succeeds", ok ~= nil)
h.test("print_settings created", sheet1.print_settings ~= nil)
h.test("orientation is landscape", sheet1.print_settings.orientation == "landscape")

sheet1:set_orientation("portrait")
h.test("orientation changed to portrait", sheet1.print_settings.orientation == "portrait")

-- ============================================
h.section("Test 2: Set margins")
-- ============================================

local sheet2 = wb:add_sheet("Margins")
-- set_margins(top, bottom, left, right, header, footer)
local ok2 = sheet2:set_margins(1.0, 1.0, 0.75, 0.75, 0.5, 0.5)
h.test("set_margins succeeds", ok2 ~= nil)
h.test("print_settings exists", sheet2.print_settings ~= nil)
h.test("margins exists", sheet2.print_settings.margins ~= nil)
h.test("top margin set", sheet2.print_settings.margins.top == 1.0)
h.test("bottom margin set", sheet2.print_settings.margins.bottom == 1.0)
h.test("left margin set", sheet2.print_settings.margins.left == 0.75)
h.test("header margin set", sheet2.print_settings.margins.header == 0.5)

-- ============================================
h.section("Test 3: Set print area")
-- ============================================

local sheet3 = wb:add_sheet("PrintArea")
local ok3 = sheet3:set_print_area("A1:G50")
h.test("set_print_area succeeds", ok3 ~= nil)
h.test("printArea set", sheet3.print_settings.printArea == "A1:G50")

-- ============================================
h.section("Test 4: Set print title rows")
-- ============================================

local sheet4 = wb:add_sheet("TitleRows")
-- set_print_title_rows takes a string like "1:2"
local ok4 = sheet4:set_print_title_rows("1:2")
h.test("set_print_title_rows succeeds", ok4 ~= nil)
h.test("printTitleRows set", sheet4.print_settings.printTitleRows == "1:2")

-- ============================================
h.section("Test 5: Set print title columns")
-- ============================================

local sheet5 = wb:add_sheet("TitleCols")
local ok5 = sheet5:set_print_title_cols("A:B")
h.test("set_print_title_cols succeeds", ok5 ~= nil)
h.test("printTitleCols set", sheet5.print_settings.printTitleCols == "A:B")

-- ============================================
h.section("Test 6: Comprehensive print settings")
-- ============================================

local sheet6 = wb:add_sheet("Comprehensive")
local ok6 = sheet6:set_print_settings({
  orientation = "landscape",
  paperSize = 9, -- A4
  scale = 85,
  fitToWidth = 1,
  fitToHeight = 0,
  horizontalCentered = true,
  verticalCentered = false,
  gridLines = true,
  headings = true,
})
h.test("comprehensive print settings succeeds", ok6 ~= nil)
h.test("orientation set", sheet6.print_settings.orientation == "landscape")
h.test("paperSize set", sheet6.print_settings.paperSize == 9)
h.test("scale set", sheet6.print_settings.scale == 85)

-- ============================================
h.section("Test 7: Print settings chaining")
-- ============================================

local sheet7 = wb:add_sheet("Chaining")
local result = sheet7:set_orientation("portrait"):set_print_area("A1:D20")
h.test("set_orientation returns worksheet", result == sheet7)
result:set_cell(1, 1, "Test")
h.test("chained set_cell works", sheet7:get_cell(1, 1).value == "Test")
h.test("both settings applied",
  sheet7.print_settings.orientation == "portrait" and
  sheet7.print_settings.printArea == "A1:D20")

-- ============================================
h.section("Test 8: Invalid orientation")
-- ============================================

local function throws(fn)
  local ok = pcall(fn)
  return not ok
end

local sheet8 = wb:add_sheet("InvalidOrientation")
h.test("invalid orientation throws", throws(function()
  sheet8:set_orientation("diagonal")
end))

-- ============================================
h.section("Test 9: No print settings (default)")
-- ============================================

local sheet9 = wb:add_sheet("NoSettings")
sheet9:set_cell(1, 1, "Normal")

-- Default sheet might or might not have print settings
h.test("default sheet renders", sheet9:to_xml() ~= nil)

-- ============================================
h.section("Test 10: XML generation")
-- ============================================

local sheet10 = wb:add_sheet("XMLTest")
sheet10:set_orientation("landscape")
sheet10:set_margins(1.0, 1.0, 0.75, 0.75)
local xml = sheet10:to_xml()

h.test("XML contains pageSetup", xml:match("<pageSetup") ~= nil)
h.test("XML contains orientation", xml:match('orientation="landscape"') ~= nil)
h.test("XML contains pageMargins", xml:match("<pageMargins") ~= nil)

-- ============================================
h.section("Test 11: Clear print area")
-- ============================================

local sheet11 = wb:add_sheet("ClearArea")
sheet11:set_print_area("A1:Z100")
h.test("print area initially set", sheet11.print_settings.printArea == "A1:Z100")

sheet11:set_print_area(nil)
h.test("print area cleared", sheet11.print_settings.printArea == nil)

-- ============================================
h.section("Test 12: Save and verify")
-- ============================================

local output_path = h.fixtures_dir .. "/test_print.xlsx"
local save_ok = wb:save(output_path)
h.test("save succeeds", save_ok == true)

-- Read back
local wb2 = xlsx.open(output_path)
h.test("file opens", wb2 ~= nil)

local read_sheet = xlsx.get_sheet(wb2, "Orientation")
h.test("print sheet exists", read_sheet ~= nil)

h.summary("Print Settings Tests")
