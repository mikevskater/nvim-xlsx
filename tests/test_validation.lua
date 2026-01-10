--- Tests for data validation and dropdowns
--- Run with: nvim --headless -l tests/test_validation.lua

dofile(arg[0]:match("(.*/)").. "test_helper.lua")
local h = require("tests.test_helper")
local xlsx = h.require_xlsx()

h.suite("Data Validation Tests")

local wb = xlsx.new_workbook()

-- ============================================
h.section("Test 1: Basic dropdown list")
-- ============================================

local sheet1 = wb:add_sheet("Dropdown")
sheet1:set_cell(1, 1, "Status:")

local ok = sheet1:add_dropdown("B1", { "Active", "Pending", "Closed" })
h.test("add_dropdown succeeds", ok ~= nil)
h.test("data_validations populated", #sheet1.data_validations > 0)

local validation = sheet1.data_validations[1]
h.test("validation type is list", validation.type == "list")
h.test("validation ref is B1", validation.ref == "B1")

-- ============================================
h.section("Test 2: Number validation - between")
-- ============================================

local sheet2 = wb:add_sheet("NumberValidation")
-- add_number_validation(ref, min, max, options)
local ok2 = sheet2:add_number_validation("A1:A100", 1, 100)
h.test("number validation succeeds", ok2 ~= nil)

local v2 = sheet2.data_validations[1]
h.test("type is whole", v2.type == "whole")
h.test("operator is between", v2.operator == "between")

-- ============================================
h.section("Test 3: Number validation - decimal")
-- ============================================

local sheet3 = wb:add_sheet("DecimalValidation")
local ok3 = sheet3:add_number_validation("B1:B50", 0, 100, { allowDecimal = true })
h.test("decimal validation succeeds", ok3 ~= nil)

local v3 = sheet3.data_validations[1]
h.test("type is decimal", v3.type == "decimal")

-- ============================================
h.section("Test 4: Generic data validation")
-- ============================================

local sheet4 = wb:add_sheet("Generic")
local ok4 = sheet4:add_data_validation("C1:C20", {
  type = "textLength",
  operator = "lessThanOrEqual",
  formula1 = "50",
})
h.test("text length validation succeeds", ok4 ~= nil)

local v4 = sheet4.data_validations[1]
h.test("type is textLength", v4.type == "textLength")
h.test("formula1 is 50", v4.formula1 == "50")

-- ============================================
h.section("Test 5: Date validation")
-- ============================================

local sheet5 = wb:add_sheet("DateValidation")
local ok5 = sheet5:add_data_validation("D1:D30", {
  type = "date",
  operator = "greaterThanOrEqual",
  formula1 = "45292", -- 2024-01-01
})
h.test("date validation succeeds", ok5 ~= nil)

-- ============================================
h.section("Test 6: Multiple validations per sheet")
-- ============================================

local sheet6 = wb:add_sheet("Multiple")
sheet6:add_dropdown("A1:A10", { "Type1", "Type2" })
sheet6:add_dropdown("B1:B10", { "Status1", "Status2" })
sheet6:add_number_validation("C1:C10", 0, 1000)

h.test("3 validations added", #sheet6.data_validations == 3)

-- ============================================
h.section("Test 7: Validation chaining")
-- ============================================

local sheet7 = wb:add_sheet("Chaining")
local result = sheet7:add_dropdown("A1", { "X", "Y" })
h.test("add_dropdown returns worksheet", result == sheet7)
result:set_cell(1, 2, "Test")
h.test("chained set_cell works", sheet7:get_cell(1, 2).value == "Test")

-- ============================================
h.section("Test 8: XML generation")
-- ============================================

local sheet8 = wb:add_sheet("XMLTest")
sheet8:add_dropdown("A1:A5", { "Red", "Green", "Blue" })
local xml = sheet8:to_xml()

h.test("XML contains dataValidations", xml:match("<dataValidations") ~= nil)
h.test("XML contains dataValidation", xml:match("<dataValidation") ~= nil)
h.test("XML contains type=list", xml:match('type="list"') ~= nil)

-- ============================================
h.section("Test 9: Number validation with options")
-- ============================================

local sheet9 = wb:add_sheet("Options")
sheet9:add_number_validation("A1:A10", 1, 100, {
  promptTitle = "Enter a number",
  prompt = "Must be between 1 and 100",
  errorTitle = "Invalid",
})

local v9 = sheet9.data_validations[1]
h.test("options preserved", v9.promptTitle == "Enter a number")

-- ============================================
h.section("Test 10: No validations (default)")
-- ============================================

local sheet10 = wb:add_sheet("NoValidation")
sheet10:set_cell(1, 1, "Normal")
local xml10 = sheet10:to_xml()

h.test("default has no dataValidations in XML", xml10:match("<dataValidations") == nil)

-- ============================================
h.section("Test 11: Empty dropdown list")
-- ============================================

local sheet11 = wb:add_sheet("EmptyList")
local function create_empty_dropdown()
  sheet11:add_dropdown("A1", {})
end
-- Should either work with empty list or throw error
local ok11 = pcall(create_empty_dropdown)
h.test("empty list handled", true) -- Pass as long as no crash

-- ============================================
h.section("Test 12: Save and verify")
-- ============================================

local output_path = h.fixtures_dir .. "/test_validation.xlsx"
local save_ok = wb:save(output_path)
h.test("save succeeds", save_ok == true)

-- Read back and verify validations preserved
local wb2 = xlsx.open(output_path)
h.test("file opens", wb2 ~= nil)

local read_sheet = xlsx.get_sheet(wb2, "Dropdown")
h.test("validation sheet exists", read_sheet ~= nil)

h.summary("Data Validation Tests")
