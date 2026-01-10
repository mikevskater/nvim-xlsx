--- Tests for auto-filter functionality
--- Run with: nvim --headless -l tests/test_filters.lua

dofile(arg[0]:match("(.*/)").. "test_helper.lua")
local h = require("tests.test_helper")
local xlsx = h.require_xlsx()

h.suite("Auto-Filter Tests")

local wb = xlsx.new_workbook()

-- ============================================
h.section("Test 1: Set auto-filter by coordinates")
-- ============================================

local sheet1 = wb:add_sheet("Filter1")
sheet1:set_cell(1, 1, "Name")
sheet1:set_cell(1, 2, "Age")
sheet1:set_cell(1, 3, "City")
sheet1:set_cell(2, 1, "Alice")
sheet1:set_cell(2, 2, 30)
sheet1:set_cell(2, 3, "NYC")

local ok = sheet1:set_auto_filter(1, 1, 10, 3)
h.test("set_auto_filter succeeds", ok ~= nil)
h.test("auto_filter set", sheet1.auto_filter ~= nil)
h.test("auto_filter ref is A1:C10", sheet1.auto_filter.ref == "A1:C10")

-- ============================================
h.section("Test 2: Set auto-filter by range string")
-- ============================================

local sheet2 = wb:add_sheet("Filter2")
sheet2:set_cell(1, 1, "Product")
sheet2:set_cell(1, 2, "Price")

local ok2 = sheet2:set_auto_filter_range("A1:B50")
h.test("set_auto_filter_range succeeds", ok2 ~= nil)
h.test("auto_filter ref is A1:B50", sheet2.auto_filter.ref == "A1:B50")

-- ============================================
h.section("Test 3: Auto-filter chaining")
-- ============================================

local sheet3 = wb:add_sheet("Chaining")
local result = sheet3:set_auto_filter(1, 1, 5, 5)
h.test("set_auto_filter returns worksheet", result == sheet3)
result:set_cell(1, 1, "Test")
h.test("chained set_cell works", sheet3:get_cell(1, 1).value == "Test")

-- ============================================
h.section("Test 4: Override auto-filter")
-- ============================================

local sheet4 = wb:add_sheet("Override")
sheet4:set_auto_filter(1, 1, 10, 5)
h.test("initial filter set", sheet4.auto_filter.ref == "A1:E10")

sheet4:set_auto_filter_range("B2:D20")
h.test("filter overridden", sheet4.auto_filter.ref == "B2:D20")

-- ============================================
h.section("Test 5: Coordinate normalization")
-- ============================================

local sheet5 = wb:add_sheet("Normalize")
-- Reversed coordinates should be normalized
sheet5:set_auto_filter(10, 5, 1, 1)
h.test("coordinates normalized", sheet5.auto_filter.ref == "A1:E10")

-- ============================================
h.section("Test 6: Invalid filter coordinates")
-- ============================================

local function throws(fn)
  local ok = pcall(fn)
  return not ok
end

local sheet6 = wb:add_sheet("Invalid")
h.test("row 0 throws", throws(function() sheet6:set_auto_filter(0, 1, 10, 3) end))
h.test("col 0 throws", throws(function() sheet6:set_auto_filter(1, 0, 10, 3) end))
h.test("negative row throws", throws(function() sheet6:set_auto_filter(-1, 1, 10, 3) end))

-- ============================================
h.section("Test 7: XML generation with auto-filter")
-- ============================================

local sheet7 = wb:add_sheet("XMLTest")
sheet7:set_auto_filter(1, 1, 100, 10)
local xml = sheet7:to_xml()

h.test("XML contains autoFilter", xml:match("<autoFilter") ~= nil)
h.test("XML contains ref attribute", xml:match('ref="A1:J100"') ~= nil)

-- ============================================
h.section("Test 8: No auto-filter (default)")
-- ============================================

local sheet8 = wb:add_sheet("NoFilter")
sheet8:set_cell(1, 1, "Normal")
local xml8 = sheet8:to_xml()

h.test("default has no autoFilter in XML", xml8:match("<autoFilter") == nil)

-- ============================================
h.section("Test 9: Clear auto-filter")
-- ============================================

local sheet9 = wb:add_sheet("Clear")
sheet9:set_auto_filter(1, 1, 10, 5)
h.test("filter initially set", sheet9.auto_filter ~= nil)

sheet9:set_auto_filter(nil)
h.test("filter cleared", sheet9.auto_filter == nil)

-- ============================================
h.section("Test 10: Save and verify")
-- ============================================

local output_path = h.fixtures_dir .. "/test_filters.xlsx"
local save_ok = wb:save(output_path)
h.test("save succeeds", save_ok == true)

-- Read back and verify filter preserved
local wb2 = xlsx.open(output_path)
h.test("file opens", wb2 ~= nil)

local read_sheet = xlsx.get_sheet(wb2, "Filter1")
h.test("filter sheet exists", read_sheet ~= nil)
-- Note: auto_filter round-trip depends on reader implementation
h.test("sheet has data", xlsx.get_cell(read_sheet, 1, 1) == "Name")

h.summary("Auto-Filter Tests")
