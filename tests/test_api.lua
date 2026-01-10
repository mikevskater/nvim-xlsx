--- Tests for public API verification
--- Run with: nvim --headless -l tests/test_api.lua

dofile(arg[0]:match("(.*/)").. "test_helper.lua")
local h = require("tests.test_helper")
local xlsx = h.require_xlsx()

h.suite("Public API Tests")

-- ============================================
h.section("Test 1: Top-level exports")
-- ============================================

h.test("xlsx module exists", xlsx ~= nil)
h.test("new_workbook function", type(xlsx.new_workbook) == "function")
h.test("open function", type(xlsx.open) == "function")
h.test("reader module", xlsx.reader ~= nil)
h.test("date module", xlsx.date ~= nil)
h.test("utils module", xlsx.utils ~= nil)

-- Module-level convenience functions
h.test("export_table function", type(xlsx.export_table) == "function")
h.test("import_table function", type(xlsx.import_table) == "function")
h.test("to_csv function", type(xlsx.to_csv) == "function")

-- ============================================
h.section("Test 2: Style constants")
-- ============================================

h.test("BORDER_STYLES exported", xlsx.BORDER_STYLES ~= nil)
h.test("HALIGN exported", xlsx.HALIGN ~= nil)
h.test("VALIGN exported", xlsx.VALIGN ~= nil)
h.test("UNDERLINE exported", xlsx.UNDERLINE ~= nil)
h.test("BUILTIN_FORMATS exported", xlsx.BUILTIN_FORMATS ~= nil)

-- Verify constant values
h.test("BORDER_STYLES.thin", xlsx.BORDER_STYLES.thin == "thin")
h.test("BORDER_STYLES.medium", xlsx.BORDER_STYLES.medium == "medium")
h.test("BORDER_STYLES.thick", xlsx.BORDER_STYLES.thick == "thick")
h.test("BORDER_STYLES.double", xlsx.BORDER_STYLES.double == "double")
h.test("BORDER_STYLES.dashed", xlsx.BORDER_STYLES.dashed == "dashed")
h.test("BORDER_STYLES.dotted", xlsx.BORDER_STYLES.dotted == "dotted")

h.test("HALIGN.left", xlsx.HALIGN.left == "left")
h.test("HALIGN.center", xlsx.HALIGN.center == "center")
h.test("HALIGN.right", xlsx.HALIGN.right == "right")

h.test("VALIGN.top", xlsx.VALIGN.top == "top")
h.test("VALIGN.center", xlsx.VALIGN.center == "center")
h.test("VALIGN.bottom", xlsx.VALIGN.bottom == "bottom")

h.test("BUILTIN_FORMATS.percent", xlsx.BUILTIN_FORMATS.percent == 9)
h.test("BUILTIN_FORMATS.date", xlsx.BUILTIN_FORMATS.date == 14)

-- ============================================
h.section("Test 3: Workbook API")
-- ============================================

local wb = xlsx.new_workbook()
h.test("add_sheet method", type(wb.add_sheet) == "function")
h.test("get_sheet method", type(wb.get_sheet) == "function")
h.test("set_active_sheet method", type(wb.set_active_sheet) == "function")
h.test("set_properties method", type(wb.set_properties) == "function")
h.test("create_style method", type(wb.create_style) == "function")
h.test("save method", type(wb.save) == "function")
h.test("sheets array", type(wb.sheets) == "table")
h.test("sheet_map table", type(wb.sheet_map) == "table")
h.test("styles object", wb.styles ~= nil)

-- ============================================
h.section("Test 4: Worksheet API - Core")
-- ============================================

local sheet = wb:add_sheet("API Test")
h.test("set_cell method", type(sheet.set_cell) == "function")
h.test("get_cell method", type(sheet.get_cell) == "function")
h.test("set method (A1)", type(sheet.set) == "function")
h.test("get method (A1)", type(sheet.get) == "function")
h.test("set_cell_value method", type(sheet.set_cell_value) == "function")
h.test("set_cell_style method", type(sheet.set_cell_style) == "function")
h.test("set_range_style method", type(sheet.set_range_style) == "function")
h.test("set_formula method", type(sheet.set_formula) == "function")
h.test("set_date method", type(sheet.set_date) == "function")
h.test("set_boolean method", type(sheet.set_boolean) == "function")
h.test("set_column_width method", type(sheet.set_column_width) == "function")
h.test("set_row_height method", type(sheet.set_row_height) == "function")
h.test("get_dimension method", type(sheet.get_dimension) == "function")
h.test("to_xml method", type(sheet.to_xml) == "function")

-- ============================================
h.section("Test 5: Worksheet API - Merging")
-- ============================================

h.test("merge_cells method", type(sheet.merge_cells) == "function")
h.test("merge_range method", type(sheet.merge_range) == "function")

-- ============================================
h.section("Test 6: Worksheet API - Features")
-- ============================================

h.test("freeze_panes method", type(sheet.freeze_panes) == "function")
h.test("freeze_rows method", type(sheet.freeze_rows) == "function")
h.test("freeze_cols method", type(sheet.freeze_cols) == "function")
h.test("set_auto_filter method", type(sheet.set_auto_filter) == "function")
h.test("set_auto_filter_range method", type(sheet.set_auto_filter_range) == "function")
h.test("add_data_validation method", type(sheet.add_data_validation) == "function")
h.test("add_dropdown method", type(sheet.add_dropdown) == "function")
h.test("add_number_validation method", type(sheet.add_number_validation) == "function")
h.test("add_hyperlink method", type(sheet.add_hyperlink) == "function")

-- ============================================
h.section("Test 7: Worksheet API - Print")
-- ============================================

h.test("set_print_settings method", type(sheet.set_print_settings) == "function")
h.test("set_margins method", type(sheet.set_margins) == "function")
h.test("set_orientation method", type(sheet.set_orientation) == "function")
h.test("set_print_area method", type(sheet.set_print_area) == "function")
h.test("set_print_title_rows method", type(sheet.set_print_title_rows) == "function")
h.test("set_print_title_cols method", type(sheet.set_print_title_cols) == "function")

-- ============================================
h.section("Test 8: Date utilities API")
-- ============================================

h.test("date.to_serial function", type(xlsx.date.to_serial) == "function")
h.test("date.from_serial function", type(xlsx.date.from_serial) == "function")

-- Verify they work
local serial = xlsx.date.to_serial({ year = 2024, month = 1, day = 1 })
h.test("to_serial returns number", type(serial) == "number")

local date = xlsx.date.from_serial(45292)
h.test("from_serial returns table", type(date) == "table")
h.test("from_serial has year", date.year == 2024)

-- ============================================
h.section("Test 9: Reader API")
-- ============================================

h.test("reader.get_range function", type(xlsx.reader.get_range) == "function")
h.test("reader.get_cell function", type(xlsx.reader.get_cell) == "function")
h.test("reader.get_sheet function", type(xlsx.reader.get_sheet) == "function")
h.test("reader.get_sheet_names function", type(xlsx.reader.get_sheet_names) == "function")
h.test("reader.get_merged_cells function", type(xlsx.reader.get_merged_cells) == "function")
h.test("reader.get_column_widths function", type(xlsx.reader.get_column_widths) == "function")

-- ============================================
h.section("Test 10: Utils API")
-- ============================================

h.test("utils.column module", xlsx.utils.column ~= nil)
h.test("utils.column.to_letter function", type(xlsx.utils.column.to_letter) == "function")
h.test("utils.column.to_number function", type(xlsx.utils.column.to_number) == "function")
h.test("utils.column.parse_ref function", type(xlsx.utils.column.parse_ref) == "function")
h.test("utils.column.make_ref function", type(xlsx.utils.column.make_ref) == "function")

-- Verify they work
h.test("to_letter(1) = A", xlsx.utils.column.to_letter(1) == "A")
h.test("to_letter(26) = Z", xlsx.utils.column.to_letter(26) == "Z")
h.test("to_letter(27) = AA", xlsx.utils.column.to_letter(27) == "AA")
h.test("to_number(A) = 1", xlsx.utils.column.to_number("A") == 1)
h.test("to_number(AA) = 27", xlsx.utils.column.to_number("AA") == 27)

local parsed = xlsx.utils.column.parse_ref("B3")
h.test("parse_ref B3 row", parsed.row == 3)
h.test("parse_ref B3 col", parsed.col == 2)

local ref = xlsx.utils.column.make_ref(5, 3)
h.test("make_ref(5,3) = C5", ref == "C5")

-- ============================================
h.section("Test 11: Validation utilities")
-- ============================================

h.test("utils.validation module", xlsx.utils.validation ~= nil)
h.test("validation.LIMITS", xlsx.utils.validation.LIMITS ~= nil)
h.test("LIMITS.MAX_ROWS", xlsx.utils.validation.LIMITS.MAX_ROWS == 1048576)
h.test("LIMITS.MAX_COLS", xlsx.utils.validation.LIMITS.MAX_COLS == 16384)
h.test("LIMITS.MAX_SHEET_NAME", xlsx.utils.validation.LIMITS.MAX_SHEET_NAME == 31)

h.test("validate_row function", type(xlsx.utils.validation.validate_row) == "function")
h.test("validate_col function", type(xlsx.utils.validation.validate_col) == "function")
h.test("validate_cell_ref function", type(xlsx.utils.validation.validate_cell_ref) == "function")
h.test("validate_sheet_name function", type(xlsx.utils.validation.validate_sheet_name) == "function")

-- ============================================
h.section("Test 12: Cell API")
-- ============================================

local cell = sheet:set_cell(1, 1, "Test")
h.test("cell has row", cell.row == 1)
h.test("cell has col", cell.col == 1)
h.test("cell has value", cell.value == "Test")
h.test("cell has value_type", cell.value_type ~= nil)
h.test("cell to_xml method", type(cell.to_xml) == "function")

-- ============================================
h.section("Test 13: Method chaining verification")
-- ============================================

local chain_sheet = wb:add_sheet("Chaining")

-- These should all return the worksheet for chaining
local result1 = chain_sheet:freeze_panes(1, 1)
h.test("freeze_panes returns worksheet", result1 == chain_sheet)

local result2 = chain_sheet:set_auto_filter(1, 1, 10, 5)
h.test("set_auto_filter returns worksheet", result2 == chain_sheet)

local result3 = chain_sheet:add_dropdown("A1", { "X", "Y" })
h.test("add_dropdown returns worksheet", result3 == chain_sheet)

local result4 = chain_sheet:add_hyperlink(2, 1, "https://test.com")
h.test("add_hyperlink returns worksheet", result4 == chain_sheet)

local result5 = chain_sheet:set_orientation("landscape")
h.test("set_orientation returns worksheet", result5 == chain_sheet)

-- These return Cell, not worksheet
local cell_result = chain_sheet:set_cell(2, 2, "Cell")
h.test("set_cell returns cell", cell_result.value == "Cell")

local formula_result = chain_sheet:set_formula(3, 3, "A1+B1")
h.test("set_formula returns cell", formula_result.formula ~= nil)

h.summary("Public API Tests")
