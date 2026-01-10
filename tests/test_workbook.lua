--- Tests for workbook creation, sheets, properties, and saving
--- Run with: nvim --headless -l tests/test_workbook.lua

dofile(arg[0]:match("(.*/)").. "test_helper.lua")
local h = require("tests.test_helper")
local xlsx = h.require_xlsx()

h.suite("Workbook Tests")

-- ============================================
h.section("Test 1: Basic workbook creation")
-- ============================================

local wb = xlsx.new_workbook()
h.test("new_workbook returns workbook", wb ~= nil)
h.test("workbook has sheets array", type(wb.sheets) == "table")
h.test("workbook has sheet_map", type(wb.sheet_map) == "table")
h.test("workbook has styles", wb.styles ~= nil)

-- ============================================
h.section("Test 2: Adding sheets")
-- ============================================

local sheet1 = wb:add_sheet("Sales")
h.test("add_sheet returns worksheet", sheet1 ~= nil)
h.test("sheet has correct name", sheet1.name == "Sales")
h.test("sheet has correct index", sheet1.index == 1)
h.test("sheet is in sheets array", wb.sheets[1] == sheet1)
h.test("sheet is in sheet_map", wb.sheet_map["Sales"] == sheet1)

local sheet2 = wb:add_sheet("Expenses")
h.test("second sheet has correct index", sheet2.index == 2)
h.test("workbook has 2 sheets", #wb.sheets == 2)

-- ============================================
h.section("Test 3: Auto-generated sheet names")
-- ============================================

local wb2 = xlsx.new_workbook()
local auto1 = wb2:add_sheet()
h.test("auto sheet name is Sheet1", auto1.name == "Sheet1")
local auto2 = wb2:add_sheet()
h.test("auto sheet name is Sheet2", auto2.name == "Sheet2")

-- ============================================
h.section("Test 4: Duplicate sheet name prevention")
-- ============================================

local dup_sheet, dup_err = wb:add_sheet("Sales")
h.test("duplicate name returns nil", dup_sheet == nil)
h.test("duplicate name returns error", dup_err ~= nil)
h.test("error mentions existing name", dup_err:match("already exists") ~= nil)

-- ============================================
h.section("Test 5: Invalid sheet names")
-- ============================================

local invalid_names = {
  { name = "", desc = "empty string" },
  { name = "A/B", desc = "contains /" },
  { name = "A\\B", desc = "contains \\" },
  { name = "A*B", desc = "contains *" },
  { name = "A?B", desc = "contains ?" },
  { name = "A:B", desc = "contains :" },
  { name = "[Test]", desc = "contains []" },
  { name = "'Test", desc = "starts with '" },
  { name = "Test'", desc = "ends with '" },
  { name = "History", desc = "reserved name" },
  { name = string.rep("A", 32), desc = "too long (32 chars)" },
}

for _, case in ipairs(invalid_names) do
  local wb_test = xlsx.new_workbook()
  local s, err = wb_test:add_sheet(case.name)
  h.test("rejects " .. case.desc, s == nil and err ~= nil)
end

-- ============================================
h.section("Test 6: Get sheet by name or index")
-- ============================================

h.test("get_sheet by name works", wb:get_sheet("Sales") == sheet1)
h.test("get_sheet by index works", wb:get_sheet(2) == sheet2)
h.test("get_sheet invalid name returns nil", wb:get_sheet("NonExistent") == nil)
h.test("get_sheet invalid index returns nil", wb:get_sheet(99) == nil)

-- ============================================
h.section("Test 7: Active sheet selection")
-- ============================================

h.test("default active sheet is 1", wb.active_sheet == 1)
h.test("set_active_sheet by index", wb:set_active_sheet(2) == true)
h.test("active sheet changed to 2", wb.active_sheet == 2)
h.test("set_active_sheet by name", wb:set_active_sheet("Sales") == true)
h.test("active sheet changed back to 1", wb.active_sheet == 1)
h.test("set_active_sheet invalid returns false", wb:set_active_sheet("Nope") == false)

-- ============================================
h.section("Test 8: Document properties")
-- ============================================

wb:set_properties({
  creator = "Test User",
  title = "Test Workbook",
  subject = "Testing",
})
h.test("creator property set", wb.properties.creator == "Test User")
h.test("title property set", wb.properties.title == "Test Workbook")
h.test("subject property set", wb.properties.subject == "Testing")
h.test("modified timestamp updated", wb.properties.modified ~= nil)

-- ============================================
h.section("Test 9: Saving workbook")
-- ============================================

-- Add some data first
sheet1:set_cell(1, 1, "Test Data")
sheet1:set_cell(1, 2, 42)

local output_path = h.fixtures_dir .. "/test_workbook.xlsx"
local save_ok, save_err = wb:save(output_path)
h.test("save succeeds", save_ok == true)
h.test("no save error", save_err == nil)

-- Verify file exists
local f = io.open(output_path, "rb")
h.test("file was created", f ~= nil)
if f then
  local size = f:seek("end")
  h.test("file has content", size > 0)
  f:close()
end

-- ============================================
h.section("Test 10: Empty workbook auto-creates sheet")
-- ============================================

local wb_empty = xlsx.new_workbook()
local empty_path = h.fixtures_dir .. "/test_workbook_empty.xlsx"
local empty_ok = wb_empty:save(empty_path)
h.test("empty workbook save succeeds", empty_ok == true)
h.test("empty workbook now has 1 sheet", #wb_empty.sheets == 1)

h.summary("Workbook Tests")
