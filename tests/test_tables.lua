--- Tests for Excel Tables (structured ListObjects)
--- Run with: nvim --headless -l tests/test_tables.lua

dofile(arg[0]:match("(.*/)").. "test_helper.lua")
local h = require("tests.test_helper")
local xlsx = h.require_xlsx()

h.suite("Excel Table Tests")

-- ============================================
h.section("Test 1: Basic table creation")
-- ============================================

local wb1 = xlsx.new_workbook()
local sheet1 = wb1:add_sheet("Basic")
sheet1:set_cell(1, 1, "Name")
sheet1:set_cell(1, 2, "Amount")
sheet1:set_cell(1, 3, "Region")
sheet1:set_cell(2, 1, "Alice")
sheet1:set_cell(2, 2, 500)
sheet1:set_cell(2, 3, "East")
sheet1:set_cell(3, 1, "Bob")
sheet1:set_cell(3, 2, 300)
sheet1:set_cell(3, 3, "West")

local result = sheet1:add_table(1, 1, 3, 3)
h.test("add_table returns worksheet for chaining", result == sheet1)
h.test("tables array populated", #sheet1.tables == 1)

local tbl = sheet1.tables[1]
h.test("table has id", tbl.id == 1)
h.test("table has default name", tbl.name == "Table1")
h.test("table ref correct", tbl.ref == "A1:C3")
h.test("table has 3 columns", #tbl.columns == 3)
h.test("column 1 name from header", tbl.columns[1].name == "Name")
h.test("column 2 name from header", tbl.columns[2].name == "Amount")
h.test("column 3 name from header", tbl.columns[3].name == "Region")
h.test("auto_filter default true", tbl.auto_filter == true)
h.test("show_row_stripes default true", tbl.show_row_stripes == true)
h.test("show_first_col default false", tbl.show_first_col == false)
h.test("show_last_col default false", tbl.show_last_col == false)
h.test("show_col_stripes default false", tbl.show_col_stripes == false)
h.test("default style", tbl.style_name == "TableStyleMedium2")

-- ============================================
h.section("Test 2: Custom table name and style")
-- ============================================

local wb2 = xlsx.new_workbook()
local sheet2 = wb2:add_sheet("Custom")
sheet2:set_cell(1, 1, "ID")
sheet2:set_cell(1, 2, "Value")
sheet2:set_cell(2, 1, 1)
sheet2:set_cell(2, 2, "test")

sheet2:add_table(1, 1, 2, 2, {
  name = "MyCustomTable",
  style_name = "TableStyleMedium9",
})

local tbl2 = sheet2.tables[1]
h.test("custom name set", tbl2.name == "MyCustomTable")
h.test("custom style set", tbl2.style_name == "TableStyleMedium9")

-- ============================================
h.section("Test 3: Disabled auto-filter")
-- ============================================

local wb3 = xlsx.new_workbook()
local sheet3 = wb3:add_sheet("NoFilter")
sheet3:set_cell(1, 1, "A")
sheet3:set_cell(1, 2, "B")
sheet3:set_cell(2, 1, 1)
sheet3:set_cell(2, 2, 2)

sheet3:add_table(1, 1, 2, 2, { auto_filter = false })
h.test("auto_filter disabled", sheet3.tables[1].auto_filter == false)

-- ============================================
h.section("Test 4: Style options")
-- ============================================

local wb4 = xlsx.new_workbook()
local sheet4 = wb4:add_sheet("StyleOpts")
sheet4:set_cell(1, 1, "H1")
sheet4:set_cell(1, 2, "H2")
sheet4:set_cell(2, 1, "d1")
sheet4:set_cell(2, 2, "d2")

sheet4:add_table(1, 1, 2, 2, {
  show_first_col = true,
  show_last_col = true,
  show_row_stripes = false,
  show_col_stripes = true,
})

local tbl4 = sheet4.tables[1]
h.test("show_first_col true", tbl4.show_first_col == true)
h.test("show_last_col true", tbl4.show_last_col == true)
h.test("show_row_stripes false", tbl4.show_row_stripes == false)
h.test("show_col_stripes true", tbl4.show_col_stripes == true)

-- ============================================
h.section("Test 5: Column name fallback for empty headers")
-- ============================================

local wb5 = xlsx.new_workbook()
local sheet5 = wb5:add_sheet("Fallback")
sheet5:set_cell(1, 1, "Header1")
-- Column 2 has no header value
sheet5:set_cell(1, 3, "Header3")
sheet5:set_cell(2, 1, "a")
sheet5:set_cell(2, 2, "b")
sheet5:set_cell(2, 3, "c")

sheet5:add_table(1, 1, 2, 3)

local tbl5 = sheet5.tables[1]
h.test("col 1 from header", tbl5.columns[1].name == "Header1")
h.test("col 2 fallback name", tbl5.columns[2].name == "Column2")
h.test("col 3 from header", tbl5.columns[3].name == "Header3")

-- ============================================
h.section("Test 6: Coordinate normalization")
-- ============================================

local wb6 = xlsx.new_workbook()
local sheet6 = wb6:add_sheet("Normalize")
sheet6:set_cell(1, 1, "H")
sheet6:set_cell(2, 1, "D")

-- Pass reversed coordinates
sheet6:add_table(2, 1, 1, 1)
-- After normalization: r1=1, r2=2, c1=1, c2=1 — but that's single column
-- Actually r1=1 < r2=2 is fine, c1=1 == c2=1 is fine (single column table)
local tbl6 = sheet6.tables[1]
h.test("normalized ref", tbl6.ref == "A1:A2")

-- ============================================
h.section("Test 7: Single-row error")
-- ============================================

local wb7 = xlsx.new_workbook()
local sheet7 = wb7:add_sheet("Error")
sheet7:set_cell(1, 1, "H")

local ok7, err7 = pcall(function()
  sheet7:add_table(1, 1, 1, 1)
end)
h.test("single-row table errors", ok7 == false)
h.test("error message mentions r1 == r2", err7 and err7:find("r1 == r2") ~= nil)

-- ============================================
h.section("Test 8: Invalid table name error")
-- ============================================

local wb8 = xlsx.new_workbook()
local sheet8 = wb8:add_sheet("BadName")
sheet8:set_cell(1, 1, "H")
sheet8:set_cell(2, 1, "D")

local ok8 = pcall(function()
  sheet8:add_table(1, 1, 2, 1, { name = "123invalid" })
end)
h.test("invalid name errors", ok8 == false)

local ok8b = pcall(function()
  sheet8:add_table(1, 1, 2, 1, { name = "has space" })
end)
h.test("name with spaces errors", ok8b == false)

-- ============================================
h.section("Test 9: Method chaining")
-- ============================================

local wb9 = xlsx.new_workbook()
local sheet9 = wb9:add_sheet("Chain")
sheet9:set_cell(1, 1, "A")
sheet9:set_cell(1, 2, "B")
sheet9:set_cell(2, 1, 1)
sheet9:set_cell(2, 2, 2)
sheet9:set_cell(3, 1, 3)
sheet9:set_cell(3, 2, 4)

-- Chain add_table with set_cell
local chain_result = sheet9:add_table(1, 1, 3, 2):set_cell(4, 1, "after")
h.test("chaining works", chain_result ~= nil)
h.test("chained set_cell value", sheet9:get_cell(4, 1).value == "after")

-- ============================================
h.section("Test 10: Multiple tables per sheet")
-- ============================================

local wb10 = xlsx.new_workbook()
local sheet10 = wb10:add_sheet("Multi")
-- Table 1
sheet10:set_cell(1, 1, "A")
sheet10:set_cell(2, 1, "a")
-- Table 2
sheet10:set_cell(1, 3, "C")
sheet10:set_cell(2, 3, "c")

sheet10:add_table(1, 1, 2, 1, { name = "First" })
sheet10:add_table(1, 3, 2, 3, { name = "Second" })

h.test("two tables on sheet", #sheet10.tables == 2)
h.test("first table name", sheet10.tables[1].name == "First")
h.test("second table name", sheet10.tables[2].name == "Second")
h.test("unique IDs", sheet10.tables[1].id ~= sheet10.tables[2].id)

-- ============================================
h.section("Test 11: Globally unique IDs across sheets")
-- ============================================

local wb11 = xlsx.new_workbook()
local s11a = wb11:add_sheet("SheetA")
s11a:set_cell(1, 1, "H")
s11a:set_cell(2, 1, "D")
s11a:add_table(1, 1, 2, 1, { name = "TableA" })

local s11b = wb11:add_sheet("SheetB")
s11b:set_cell(1, 1, "H")
s11b:set_cell(2, 1, "D")
s11b:add_table(1, 1, 2, 1, { name = "TableB" })

h.test("IDs globally unique", s11a.tables[1].id ~= s11b.tables[1].id)
h.test("ID is 1", s11a.tables[1].id == 1)
h.test("ID is 2", s11b.tables[1].id == 2)

-- ============================================
h.section("Test 12: XML - tableParts in worksheet")
-- ============================================

local wb12 = xlsx.new_workbook()
local sheet12 = wb12:add_sheet("XMLTest")
sheet12:set_cell(1, 1, "Col1")
sheet12:set_cell(1, 2, "Col2")
sheet12:set_cell(2, 1, "a")
sheet12:set_cell(2, 2, "b")
sheet12:add_table(1, 1, 2, 2)

local xml12 = sheet12:to_xml()
h.test("XML contains tableParts", xml12:find("<tableParts") ~= nil)
h.test("XML contains tablePart with rId", xml12:find('r:id="rId1"') ~= nil)
h.test("tableParts count=1", xml12:find('count="1"') ~= nil)

-- ============================================
h.section("Test 13: Table XML structure")
-- ============================================

local table_xml = wb12:_generate_table_xml(sheet12.tables[1])
h.test("table XML has declaration", table_xml:find("<?xml") ~= nil)
h.test("table XML has table element", table_xml:find("<table ") ~= nil)
h.test("table XML has displayName", table_xml:find("displayName=") ~= nil)
h.test("table XML has autoFilter", table_xml:find("<autoFilter") ~= nil)
h.test("table XML has tableColumns", table_xml:find("<tableColumns") ~= nil)
h.test("table XML has tableColumn", table_xml:find("<tableColumn") ~= nil)
h.test("table XML has tableStyleInfo", table_xml:find("<tableStyleInfo") ~= nil)
h.test("table XML has ref", table_xml:find('ref="A1:B2"') ~= nil)

-- ============================================
h.section("Test 14: Table without auto-filter XML")
-- ============================================

local wb14 = xlsx.new_workbook()
local sheet14 = wb14:add_sheet("NoFilterXML")
sheet14:set_cell(1, 1, "H")
sheet14:set_cell(2, 1, "D")
sheet14:add_table(1, 1, 2, 1, { auto_filter = false })
local xml14 = wb14:_generate_table_xml(sheet14.tables[1])
h.test("no autoFilter in XML", xml14:find("<autoFilter") == nil)

-- ============================================
h.section("Test 15: Worksheet relationships")
-- ============================================

local rels = sheet12:get_table_relationships()
h.test("one table rel", #rels == 1)
h.test("rel id is rId1", rels[1].id == "rId1")
h.test("rel target correct", rels[1].target:find("tables/table") ~= nil)
h.test("rel type is table", rels[1].type:find("table") ~= nil)

-- ============================================
h.section("Test 16: Content types")
-- ============================================

local ct_xml = wb12:_generate_content_types()
h.test("content types has table override", ct_xml:find("table+xml", 1, true) ~= nil)
h.test("content types has table path", ct_xml:find("/xl/tables/table") ~= nil)

-- ============================================
h.section("Test 17: No tables (default)")
-- ============================================

local wb17 = xlsx.new_workbook()
local sheet17 = wb17:add_sheet("Empty")
sheet17:set_cell(1, 1, "data")
local xml17 = sheet17:to_xml()
h.test("no tableParts in default XML", xml17:find("<tableParts") == nil)

-- ============================================
h.section("Test 18: Round-trip save")
-- ============================================

local output_path = h.fixtures_dir .. "/test_tables.xlsx"
local save_ok = wb1:save(output_path)
h.test("save with tables succeeds", save_ok == true)

-- Also save multi-table workbook
local output_path2 = h.fixtures_dir .. "/test_tables_multi.xlsx"
local save_ok2 = wb10:save(output_path2)
h.test("save multi-table workbook succeeds", save_ok2 == true)

-- Save tables + hyperlinks together
local wb_combo = xlsx.new_workbook()
local sheet_combo = wb_combo:add_sheet("Combo")
sheet_combo:set_cell(1, 1, "Link")
sheet_combo:set_cell(1, 2, "Value")
sheet_combo:set_cell(2, 1, "click")
sheet_combo:set_cell(2, 2, 42)
sheet_combo:add_table(1, 1, 2, 2, { name = "ComboTable" })
sheet_combo:add_hyperlink(2, 1, "https://example.com")

local output_path3 = h.fixtures_dir .. "/test_tables_combo.xlsx"
local save_ok3 = wb_combo:save(output_path3)
h.test("save tables+hyperlinks succeeds", save_ok3 == true)

h.summary("Excel Table Tests")
