--- Tests for import/export functionality: export_table, import_table, to_csv
--- Run with: nvim --headless -l tests/test_import_export.lua

dofile(arg[0]:match("(.*/)").. "test_helper.lua")
local h = require("tests.test_helper")
local xlsx = h.require_xlsx()

h.suite("Import/Export Tests")

-- ============================================
h.section("Test 1: export_table basic")
-- ============================================

local data = {
  { "Alice", 30, "New York" },
  { "Bob", 25, "London" },
  { "Charlie", 35, "Tokyo" },
}

local export_path = h.fixtures_dir .. "/test_export_basic.xlsx"
local ok, err = xlsx.export_table(data, export_path)
h.test("export_table succeeds", ok == true)
h.test("no error returned", err == nil)

-- Verify file exists
local f = io.open(export_path, "rb")
h.test("file was created", f ~= nil)
if f then f:close() end

-- ============================================
h.section("Test 2: export_table with headers")
-- ============================================

local headers_path = h.fixtures_dir .. "/test_export_headers.xlsx"
local ok2 = xlsx.export_table(data, headers_path, {
  headers = { "Name", "Age", "City" }
})
h.test("export_table with headers succeeds", ok2 == true)

-- Read it back and verify headers
local wb2 = xlsx.open(headers_path)
local sheet2 = xlsx.get_sheet_by_index(wb2, 1)
local header_val = xlsx.get_cell(sheet2, 1, 1)
h.test("header row written", header_val == "Name")

local data_val = xlsx.get_cell(sheet2, 2, 1)
h.test("data row at row 2", data_val == "Alice")

-- ============================================
h.section("Test 3: export_table with sheet name")
-- ============================================

local sheet_path = h.fixtures_dir .. "/test_export_sheet.xlsx"
local ok3 = xlsx.export_table(data, sheet_path, {
  sheet_name = "MyData"
})
h.test("export_table with sheet_name succeeds", ok3 == true)

local wb3 = xlsx.open(sheet_path)
local names3 = xlsx.get_sheet_names(wb3)
h.test("custom sheet name used", names3[1] == "MyData")

-- ============================================
h.section("Test 4: import_table basic")
-- ============================================

local imported, imp_err = xlsx.import_table(export_path)
h.test("import_table succeeds", imported ~= nil)
h.test("no import error", imp_err == nil)
h.test("imported has 3 rows", #imported == 3)
h.test("imported[1][1] is Alice", imported[1][1] == "Alice")
h.test("imported[1][2] is 30", imported[1][2] == 30)
h.test("imported[3][3] is Tokyo", imported[3][3] == "Tokyo")

-- ============================================
h.section("Test 5: import_table with sheet_name")
-- ============================================

local imported_named, imp_err2 = xlsx.import_table(sheet_path, {
  sheet_name = "MyData"
})
h.test("import with sheet_name works", imported_named ~= nil)
h.test("imported has data", #imported_named == 3)

-- ============================================
h.section("Test 6: import_table with sheet_index")
-- ============================================

local imported_idx = xlsx.import_table(export_path, { sheet_index = 1 })
h.test("import with sheet_index works", imported_idx ~= nil)

-- ============================================
h.section("Test 7: import_table non-existent file")
-- ============================================

local bad_import, bad_err = xlsx.import_table("/nonexistent/file.xlsx")
h.test("non-existent file returns nil", bad_import == nil)
h.test("error returned", bad_err ~= nil)

-- ============================================
h.section("Test 8: Round-trip export/import")
-- ============================================

local roundtrip_data = {
  { 1, 2, 3 },
  { 4, 5, 6 },
  { 7, 8, 9 },
}

local rt_path = h.fixtures_dir .. "/test_roundtrip.xlsx"
xlsx.export_table(roundtrip_data, rt_path)
local rt_imported = xlsx.import_table(rt_path)

h.test("roundtrip preserves structure", #rt_imported == 3 and #rt_imported[1] == 3)
h.test("roundtrip preserves values", rt_imported[1][1] == 1 and rt_imported[3][3] == 9)

-- ============================================
h.section("Test 9: Export empty data")
-- ============================================

local empty_path = h.fixtures_dir .. "/test_export_empty.xlsx"
local empty_ok = xlsx.export_table({}, empty_path)
h.test("empty export succeeds", empty_ok == true)

-- ============================================
h.section("Test 10: to_csv function")
-- ============================================

local csv_path = h.fixtures_dir .. "/test_to_csv.csv"
local csv_result, csv_err = xlsx.to_csv(export_path)
h.test("to_csv returns string", type(csv_result) == "string")
h.test("no csv error", csv_err == nil)
h.test("csv contains data", csv_result:match("Alice") ~= nil)
h.test("csv has commas", csv_result:match(",") ~= nil)

-- ============================================
h.section("Test 11: Mixed data types")
-- ============================================

local mixed_data = {
  { "Text", 123, 45.67, true, false },
  { "More", 0, -10.5, false, true },
}

local mixed_path = h.fixtures_dir .. "/test_mixed.xlsx"
local mixed_ok = xlsx.export_table(mixed_data, mixed_path)
h.test("mixed types export succeeds", mixed_ok == true)

local mixed_import = xlsx.import_table(mixed_path)
h.test("mixed import succeeds", mixed_import ~= nil)
h.test("string preserved", mixed_import[1][1] == "Text")
h.test("integer preserved", mixed_import[1][2] == 123)
h.test("float preserved", mixed_import[1][3] == 45.67)

-- ============================================
h.section("Test 12: Large dataset")
-- ============================================

local large_data = {}
for i = 1, 100 do
  table.insert(large_data, { i, "Row " .. i, i * 10 })
end

local large_path = h.fixtures_dir .. "/test_large.xlsx"
local large_ok = xlsx.export_table(large_data, large_path)
h.test("large export succeeds", large_ok == true)

local large_import = xlsx.import_table(large_path)
h.test("large import succeeds", large_import ~= nil)
h.test("large import has 100 rows", #large_import == 100)
h.test("large import first row correct", large_import[1][1] == 1)
h.test("large import last row correct", large_import[100][1] == 100)

h.summary("Import/Export Tests")
