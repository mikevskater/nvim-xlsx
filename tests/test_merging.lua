--- Tests for merged cells
--- Run with: nvim --headless -l tests/test_merging.lua

dofile(arg[0]:match("(.*/)").. "test_helper.lua")
local h = require("tests.test_helper")
local xlsx = h.require_xlsx()

h.suite("Cell Merging Tests")

local wb = xlsx.new_workbook()
local sheet = wb:add_sheet("Merging")

-- ============================================
h.section("Test 1: Basic merge")
-- ============================================

local ok, err = sheet:merge_cells(1, 1, 1, 4)
h.test("merge_cells succeeds", ok == true)
h.test("no error returned", err == nil)
h.test("merged_cells has entry", #sheet.merged_cells == 1)
h.test("merge reference is correct", sheet.merged_cells[1] == "A1:D1")

-- ============================================
h.section("Test 2: Merge with A1 notation")
-- ============================================

local ok2, err2 = sheet:merge_range("A3:C3")
h.test("merge_range succeeds", ok2 == true)
h.test("merged_cells count increased", #sheet.merged_cells == 2)

-- ============================================
h.section("Test 3: Vertical merge")
-- ============================================

local ok3 = sheet:merge_cells(5, 1, 8, 1)
h.test("vertical merge succeeds", ok3 == true)

-- ============================================
h.section("Test 4: Block merge")
-- ============================================

local ok4 = sheet:merge_cells(10, 1, 12, 4)
h.test("block merge succeeds", ok4 == true)

-- ============================================
h.section("Test 5: Coordinate normalization")
-- ============================================

-- Reversed coordinates should still work
local ok5 = sheet:merge_cells(16, 4, 14, 1)
h.test("reversed coordinates normalized", ok5 == true)

-- ============================================
h.section("Test 6: Single cell merge prevention")
-- ============================================

local single_ok, single_err = sheet:merge_cells(20, 1, 20, 1)
h.test("single cell merge returns false", single_ok == false)
h.test("single cell merge has error", single_err ~= nil)
h.test("error mentions single cell", single_err:match("single cell") ~= nil)

-- ============================================
h.section("Test 7: Overlap prevention")
-- ============================================

-- Try to merge overlapping with first merge (A1:D1)
local overlap_ok, overlap_err = sheet:merge_cells(1, 2, 2, 3)
h.test("overlapping merge returns false", overlap_ok == false)
h.test("overlap error returned", overlap_err ~= nil)
h.test("error mentions overlap", overlap_err:match("overlap") ~= nil)

-- ============================================
h.section("Test 8: Adjacent merges allowed")
-- ============================================

-- Adjacent to A1:D1 (row 2) should work
local adjacent_ok = sheet:merge_cells(2, 1, 2, 4)
h.test("adjacent merge succeeds", adjacent_ok == true)

-- ============================================
h.section("Test 9: Merge with content")
-- ============================================

local content_sheet = wb:add_sheet("Content")
content_sheet:set_cell(1, 1, "Merged Title")
local content_ok = content_sheet:merge_cells(1, 1, 1, 5)
h.test("merge with existing content succeeds", content_ok == true)

-- ============================================
h.section("Test 10: Dimension tracking with merges")
-- ============================================

local dim_sheet = wb:add_sheet("Dimensions")
h.test("empty sheet dimension", dim_sheet:get_dimension() == "A1")

dim_sheet:merge_cells(1, 1, 5, 5)
h.test("merge updates dimensions", dim_sheet:get_dimension() == "A1:E5")

-- ============================================
h.section("Test 11: Invalid range validation")
-- ============================================

local function returns_error(r1, c1, r2, c2)
  local ok, err = sheet:merge_cells(r1, c1, r2, c2)
  return ok == false and err ~= nil
end

h.test("invalid row returns error", returns_error(0, 1, 1, 1))
h.test("invalid col returns error", returns_error(1, 0, 1, 1))

-- ============================================
h.section("Test 12: XML generation")
-- ============================================

local xml = sheet:to_xml()
h.test("XML contains mergeCells", xml:match("<mergeCells") ~= nil)
h.test("XML contains mergeCell", xml:match("<mergeCell") ~= nil)

-- ============================================
h.section("Test 13: Save and verify")
-- ============================================

local output_path = h.fixtures_dir .. "/test_merging.xlsx"
local save_ok = wb:save(output_path)
h.test("save succeeds", save_ok == true)

h.summary("Cell Merging Tests")
