--- Tests for Named Ranges (definedNames)
--- Run with: nvim --headless -l tests/test_named_ranges.lua

dofile(arg[0]:match("(.*/)").. "test_helper.lua")
local h = require("tests.test_helper")
local xlsx = h.require_xlsx()

h.suite("Named Range Tests")

-- ============================================
h.section("Test 1: Workbook-scoped defined name")
-- ============================================

local wb1 = xlsx.new_workbook()
local sheet1 = wb1:add_sheet("Data")
sheet1:set_cell(1, 1, "Value")
sheet1:set_cell(2, 1, 42)

local result = wb1:add_defined_name("MyConstant", "42")
h.test("add_defined_name returns workbook for chaining", result == wb1)
h.test("defined_names populated", #wb1.defined_names == 1)

local dn = wb1.defined_names[1]
h.test("name stored", dn.name == "MyConstant")
h.test("ref stored", dn.ref == "42")
h.test("no local_sheet_id", dn.local_sheet_id == nil)
h.test("hidden default false", dn.hidden == false)
h.test("comment default nil", dn.comment == nil)

-- ============================================
h.section("Test 2: Sheet-scoped defined name")
-- ============================================

wb1:add_defined_name("LocalName", "Data!$A$1:$A$2", { local_sheet_id = 0 })

local dn2 = wb1.defined_names[2]
h.test("local_sheet_id set", dn2.local_sheet_id == 0)

-- ============================================
h.section("Test 3: Convenience add_named_range")
-- ============================================

local wb3 = xlsx.new_workbook()
wb3:add_sheet("Sales")

local result3 = wb3:add_named_range("SalesRange", "Sales", 1, 1, 4, 3)
h.test("add_named_range returns workbook for chaining", result3 == wb3)
h.test("defined_names populated", #wb3.defined_names == 1)

local dn3 = wb3.defined_names[1]
h.test("name correct", dn3.name == "SalesRange")
h.test("ref has absolute range", dn3.ref == "Sales!$A$1:$C$4")

-- ============================================
h.section("Test 4: Sheet name quoting for spaces")
-- ============================================

local wb4 = xlsx.new_workbook()
wb4:add_sheet("My Sheet")
wb4:add_named_range("QuotedRange", "My Sheet", 1, 1, 5, 2)

local dn4 = wb4.defined_names[1]
h.test("sheet name quoted", dn4.ref == "'My Sheet'!$A$1:$B$5")

-- ============================================
h.section("Test 5: Sheet name without spaces not quoted")
-- ============================================

local wb5 = xlsx.new_workbook()
wb5:add_sheet("Sheet1")
wb5:add_named_range("SimpleRange", "Sheet1", 2, 1, 10, 3)

local dn5 = wb5.defined_names[1]
h.test("sheet name not quoted", dn5.ref == "Sheet1!$A$2:$C$10")

-- ============================================
h.section("Test 6: Hidden flag")
-- ============================================

local wb6 = xlsx.new_workbook()
wb6:add_defined_name("HiddenName", "Sheet1!$A$1", { hidden = true })

local dn6 = wb6.defined_names[1]
h.test("hidden flag set", dn6.hidden == true)

-- ============================================
h.section("Test 7: Comment attribute")
-- ============================================

local wb7 = xlsx.new_workbook()
wb7:add_defined_name("CommentName", "Sheet1!$A$1", { comment = "This is a note" })

local dn7 = wb7.defined_names[1]
h.test("comment set", dn7.comment == "This is a note")

-- ============================================
h.section("Test 8: Invalid name validation")
-- ============================================

local wb8 = xlsx.new_workbook()

local ok8a = pcall(function()
  wb8:add_defined_name("123bad", "ref")
end)
h.test("invalid name starting with number", ok8a == false)

local ok8b = pcall(function()
  wb8:add_defined_name("has space", "ref")
end)
h.test("invalid name with space", ok8b == false)

local ok8c = pcall(function()
  wb8:add_defined_name("", "ref")
end)
h.test("empty name rejected", ok8c == false)

-- ============================================
h.section("Test 9: Method chaining")
-- ============================================

local wb9 = xlsx.new_workbook()
wb9:add_sheet("S1")

local chain = wb9:add_defined_name("First", "S1!$A$1")
  :add_defined_name("Second", "S1!$B$1")
  :add_named_range("Third", "S1", 1, 1, 10, 5)

h.test("chaining produces 3 names", #wb9.defined_names == 3)
h.test("chain returns workbook", chain == wb9)

-- ============================================
h.section("Test 10: XML - definedNames in workbook XML")
-- ============================================

local wb10 = xlsx.new_workbook()
wb10:add_sheet("Sheet1")
wb10:add_defined_name("TestName", "Sheet1!$A$1:$C$10")
wb10:add_defined_name("LocalName", "Sheet1!$A$1", { local_sheet_id = 0 })
wb10:add_defined_name("HiddenName", "42", { hidden = true })
wb10:add_defined_name("CommentedName", "Sheet1!$B$2", { comment = "A comment" })

local wb_xml = wb10:_generate_workbook_xml()
h.test("XML contains definedNames", wb_xml:find("<definedNames>") ~= nil)
h.test("XML contains definedName", wb_xml:find("<definedName ") ~= nil)
h.test("XML contains TestName", wb_xml:find('name="TestName"') ~= nil)
h.test("XML contains localSheetId", wb_xml:find('localSheetId="0"') ~= nil)
h.test("XML contains hidden attribute", wb_xml:find('hidden="1"') ~= nil)
h.test("XML contains comment attribute", wb_xml:find('comment="A comment"') ~= nil)
h.test("XML contains ref value", wb_xml:find(">Sheet1!%$A%$1:%$C%$10<") ~= nil)
h.test("XML closes definedNames", wb_xml:find("</definedNames>") ~= nil)

-- ============================================
h.section("Test 11: No defined names (default)")
-- ============================================

local wb11 = xlsx.new_workbook()
wb11:add_sheet("Empty")
local wb_xml11 = wb11:_generate_workbook_xml()
h.test("no definedNames in default XML", wb_xml11:find("<definedNames") == nil)

-- ============================================
h.section("Test 12: Round-trip save with named ranges")
-- ============================================

local wb12 = xlsx.new_workbook()
local sheet12 = wb12:add_sheet("Sales")
sheet12:set_cell(1, 1, "Product")
sheet12:set_cell(1, 2, "Amount")
sheet12:set_cell(2, 1, "Widget")
sheet12:set_cell(2, 2, 100)

wb12:add_named_range("SalesData", "Sales", 1, 1, 2, 2)
wb12:add_defined_name("Multiplier", "1.5")

local output_path = h.fixtures_dir .. "/test_named_ranges.xlsx"
local save_ok = wb12:save(output_path)
h.test("save with named ranges succeeds", save_ok == true)

-- ============================================
h.section("Test 13: Combined tables and named ranges")
-- ============================================

local wb13 = xlsx.new_workbook()
local sheet13 = wb13:add_sheet("Combined")
sheet13:set_cell(1, 1, "Name")
sheet13:set_cell(1, 2, "Score")
sheet13:set_cell(2, 1, "Alice")
sheet13:set_cell(2, 2, 95)

sheet13:add_table(1, 1, 2, 2, { name = "Scores" })
wb13:add_named_range("ScoreRange", "Combined", 1, 1, 2, 2)

local output_combo = h.fixtures_dir .. "/test_combined.xlsx"
local save_combo = wb13:save(output_combo)
h.test("save tables+named ranges succeeds", save_combo == true)

-- Verify XML has both
local wb_xml13 = wb13:_generate_workbook_xml()
h.test("workbook XML has definedNames", wb_xml13:find("<definedNames") ~= nil)
local sheet_xml13 = sheet13:to_xml()
h.test("sheet XML has tableParts", sheet_xml13:find("<tableParts") ~= nil)

h.summary("Named Range Tests")
