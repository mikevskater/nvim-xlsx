--- Tests for freeze panes functionality
--- Run with: nvim --headless -l tests/test_freeze_panes.lua

dofile(arg[0]:match("(.*/)").. "test_helper.lua")
local h = require("tests.test_helper")
local xlsx = h.require_xlsx()

h.suite("Freeze Panes Tests")

local wb = xlsx.new_workbook()

-- ============================================
h.section("Test 1: Freeze rows only")
-- ============================================

local sheet1 = wb:add_sheet("FreezeRows")
sheet1:set_cell(1, 1, "Header 1")
sheet1:set_cell(1, 2, "Header 2")
sheet1:set_cell(2, 1, "Data 1")
sheet1:set_cell(2, 2, "Data 2")

local ok = sheet1:freeze_rows(1)
h.test("freeze_rows succeeds", ok ~= nil)
h.test("freeze_pane set", sheet1.freeze_pane ~= nil)
h.test("freeze_pane rows is 1", sheet1.freeze_pane.rows == 1)
h.test("freeze_pane cols is 0", sheet1.freeze_pane.cols == 0)

-- ============================================
h.section("Test 2: Freeze columns only")
-- ============================================

local sheet2 = wb:add_sheet("FreezeCols")
sheet2:set_cell(1, 1, "Row Label")
sheet2:set_cell(1, 2, "Value 1")
sheet2:set_cell(1, 3, "Value 2")

local ok2 = sheet2:freeze_cols(1)
h.test("freeze_cols succeeds", ok2 ~= nil)
h.test("freeze_pane cols is 1", sheet2.freeze_pane.cols == 1)
h.test("freeze_pane rows is 0", sheet2.freeze_pane.rows == 0)

-- ============================================
h.section("Test 3: Freeze both rows and columns")
-- ============================================

local sheet3 = wb:add_sheet("FreezeBoth")
sheet3:set_cell(1, 1, "Corner")
sheet3:set_cell(1, 2, "Header 1")
sheet3:set_cell(2, 1, "Row 1")
sheet3:set_cell(2, 2, "Data")

local ok3 = sheet3:freeze_panes(2, 2)
h.test("freeze_panes succeeds", ok3 ~= nil)
h.test("freeze_pane rows is 2", sheet3.freeze_pane.rows == 2)
h.test("freeze_pane cols is 2", sheet3.freeze_pane.cols == 2)

-- ============================================
h.section("Test 4: Freeze panes chaining")
-- ============================================

local sheet4 = wb:add_sheet("Chaining")
local result = sheet4:freeze_rows(3)
h.test("freeze_rows returns worksheet", result == sheet4)
result:set_cell(1, 1, "Test")
h.test("chained set_cell works", sheet4:get_cell(1, 1).value == "Test")

-- ============================================
h.section("Test 5: Override freeze panes")
-- ============================================

local sheet5 = wb:add_sheet("Override")
sheet5:freeze_rows(2)
h.test("initial freeze at rows 2", sheet5.freeze_pane.rows == 2)

sheet5:freeze_panes(5, 3)
h.test("overridden rows to 5", sheet5.freeze_pane.rows == 5)
h.test("overridden cols to 3", sheet5.freeze_pane.cols == 3)

-- ============================================
h.section("Test 6: Freeze at different positions")
-- ============================================

local positions = {
  { rows = 1, cols = 0, desc = "rows 1" },
  { rows = 5, cols = 0, desc = "rows 5" },
  { rows = 0, cols = 1, desc = "cols 1" },
  { rows = 0, cols = 3, desc = "cols 3" },
  { rows = 10, cols = 5, desc = "rows 10, cols 5" },
}

for i, pos in ipairs(positions) do
  local s = wb:add_sheet("Position" .. i)
  s:freeze_panes(pos.rows, pos.cols)
  h.test("freeze at " .. pos.desc, s.freeze_pane.rows == pos.rows and s.freeze_pane.cols == pos.cols)
end

-- ============================================
h.section("Test 7: Invalid freeze values")
-- ============================================

local function throws(fn)
  local ok = pcall(fn)
  return not ok
end

local sheet7 = wb:add_sheet("Invalid")
h.test("negative rows throws", throws(function() sheet7:freeze_panes(-1, 0) end))
h.test("negative cols throws", throws(function() sheet7:freeze_panes(0, -1) end))

-- ============================================
h.section("Test 8: XML generation with freeze")
-- ============================================

local sheet8 = wb:add_sheet("XMLTest")
sheet8:freeze_panes(2, 1)
local xml = sheet8:to_xml()

h.test("XML contains sheetViews", xml:match("<sheetViews") ~= nil)
h.test("XML contains pane element", xml:match("<pane") ~= nil)
h.test("XML contains ySplit", xml:match('ySplit="2"') ~= nil)
h.test("XML contains xSplit", xml:match('xSplit="1"') ~= nil)

-- ============================================
h.section("Test 9: No freeze pane (default)")
-- ============================================

local sheet9 = wb:add_sheet("NoFreeze")
sheet9:set_cell(1, 1, "Normal")
local xml9 = sheet9:to_xml()

-- Should not have pane element when no freeze is set
h.test("default has no pane in XML", xml9:match("<pane") == nil)

-- ============================================
h.section("Test 10: Clear freeze pane")
-- ============================================

local sheet10 = wb:add_sheet("ClearFreeze")
sheet10:freeze_panes(2, 2)
h.test("freeze pane initially set", sheet10.freeze_pane ~= nil)

sheet10:freeze_panes(0, 0)
h.test("freeze pane cleared", sheet10.freeze_pane == nil)

-- ============================================
h.section("Test 11: Save and verify")
-- ============================================

local output_path = h.fixtures_dir .. "/test_freeze_panes.xlsx"
local save_ok = wb:save(output_path)
h.test("save succeeds", save_ok == true)

h.summary("Freeze Panes Tests")
