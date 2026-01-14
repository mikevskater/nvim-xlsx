--- Tests for style validation
--- Run with: nvim --headless -l tests/test_style_validation.lua

dofile(arg[0]:match("(.*/)").. "test_helper.lua")
local h = require("tests.test_helper")
local xlsx = h.require_xlsx()

h.suite("Style Validation Tests")

-- ============================================
h.section("Test 1: validate_style exposed at top level")
-- ============================================

h.test("validate_style function exists", type(xlsx.validate_style) == "function")

-- ============================================
h.section("Test 2: Valid styles pass validation")
-- ============================================

local valid, errors

valid, errors = xlsx.validate_style({})
h.test("empty style is valid", valid == true and errors == nil)

valid, errors = xlsx.validate_style({ bold = true })
h.test("bold=true is valid", valid == true)

valid, errors = xlsx.validate_style({ halign = "center" })
h.test("halign=center is valid", valid == true)

valid, errors = xlsx.validate_style({ valign = "top" })
h.test("valign=top is valid", valid == true)

valid, errors = xlsx.validate_style({ font_color = "#FF0000" })
h.test("font_color hex is valid", valid == true)

valid, errors = xlsx.validate_style({ font_color = "red" })
h.test("font_color named is valid", valid == true)

valid, errors = xlsx.validate_style({ bg_color = "#4472C4" })
h.test("bg_color hex is valid", valid == true)

valid, errors = xlsx.validate_style({ border = true, border_style = "thin" })
h.test("border with style is valid", valid == true)

valid, errors = xlsx.validate_style({ underline = "single" })
h.test("underline=single is valid", valid == true)

valid, errors = xlsx.validate_style({ font_size = 12 })
h.test("font_size=12 is valid", valid == true)

valid, errors = xlsx.validate_style({ rotation = 45 })
h.test("rotation=45 is valid", valid == true)

valid, errors = xlsx.validate_style({ rotation = 255 })
h.test("rotation=255 (vertical) is valid", valid == true)

valid, errors = xlsx.validate_style({ indent = 2 })
h.test("indent=2 is valid", valid == true)

valid, errors = xlsx.validate_style({ wrap_text = true })
h.test("wrap_text=true is valid", valid == true)

valid, errors = xlsx.validate_style({ num_format = "percent" })
h.test("num_format builtin name is valid", valid == true)

valid, errors = xlsx.validate_style({ num_format = "$#,##0.00" })
h.test("num_format custom string is valid", valid == true)

valid, errors = xlsx.validate_style({ num_format = 14 })
h.test("num_format builtin ID is valid", valid == true)

-- ============================================
h.section("Test 3: Invalid halign values")
-- ============================================

valid, errors = xlsx.validate_style({ halign = "middle" })
h.test("halign=middle is invalid", valid == false)
h.test("error mentions 'middle'", errors and errors[1]:find("middle") ~= nil)
h.test("error shows valid options", errors and errors[1]:find("center") ~= nil)

valid, errors = xlsx.validate_style({ align = "invalid" })
h.test("align=invalid is invalid", valid == false)

-- ============================================
h.section("Test 4: Invalid valign values")
-- ============================================

valid, errors = xlsx.validate_style({ valign = "middle" })
h.test("valign=middle is invalid", valid == false)
h.test("valign error mentions valid options", errors and errors[1]:find("center") ~= nil)

-- ============================================
h.section("Test 5: Invalid color values")
-- ============================================

valid, errors = xlsx.validate_style({ font_color = "not-a-color" })
h.test("font_color=not-a-color is invalid", valid == false)
h.test("color error is descriptive", errors and errors[1]:find("Invalid font_color") ~= nil)

valid, errors = xlsx.validate_style({ bg_color = "invalid-hex" })
h.test("bg_color=invalid-hex is invalid", valid == false)

valid, errors = xlsx.validate_style({ border_color = 12345 })
h.test("border_color=number is invalid", valid == false)

-- ============================================
h.section("Test 6: Invalid border styles")
-- ============================================

valid, errors = xlsx.validate_style({ border_style = "super-thick" })
h.test("border_style=super-thick is invalid", valid == false)

valid, errors = xlsx.validate_style({ border_left = "invalid-style" })
h.test("border_left=invalid-style is invalid", valid == false)

-- ============================================
h.section("Test 7: Invalid underline values")
-- ============================================

valid, errors = xlsx.validate_style({ underline = "triple" })
h.test("underline=triple is invalid", valid == false)
h.test("underline error mentions valid options", errors and errors[1]:find("single") ~= nil)

-- ============================================
h.section("Test 8: Invalid numeric values")
-- ============================================

valid, errors = xlsx.validate_style({ font_size = "large" })
h.test("font_size=string is invalid", valid == false)

valid, errors = xlsx.validate_style({ font_size = 0 })
h.test("font_size=0 is invalid", valid == false)

valid, errors = xlsx.validate_style({ font_size = 500 })
h.test("font_size=500 is invalid", valid == false)

valid, errors = xlsx.validate_style({ rotation = 100 })
h.test("rotation=100 is invalid", valid == false)

valid, errors = xlsx.validate_style({ indent = -1 })
h.test("indent=-1 is invalid", valid == false)

valid, errors = xlsx.validate_style({ indent = 1.5 })
h.test("indent=1.5 (non-integer) is invalid", valid == false)

-- ============================================
h.section("Test 9: Invalid boolean values")
-- ============================================

valid, errors = xlsx.validate_style({ bold = "yes" })
h.test("bold=string is invalid", valid == false)

valid, errors = xlsx.validate_style({ italic = 1 })
h.test("italic=number is invalid", valid == false)

valid, errors = xlsx.validate_style({ wrap_text = "true" })
h.test("wrap_text=string is invalid", valid == false)

-- ============================================
h.section("Test 10: Multiple errors collected")
-- ============================================

valid, errors = xlsx.validate_style({
  halign = "invalid1",
  valign = "invalid2",
  font_color = "bad-color",
})
h.test("multiple invalid values fails", valid == false)
h.test("multiple errors collected", errors and #errors == 3)

-- ============================================
h.section("Test 11: create_style returns errors")
-- ============================================

local wb = xlsx.new_workbook()

local style_idx, err = wb:create_style({ halign = "middle" })
h.test("create_style returns nil on invalid", style_idx == nil)
h.test("create_style returns error message", err ~= nil and err:find("middle") ~= nil)

style_idx, err = wb:create_style({ bold = true })
h.test("create_style succeeds with valid style", style_idx ~= nil and style_idx > 0)
h.test("create_style returns no error when valid", err == nil)

-- ============================================
h.section("Test 12: Complex style validation")
-- ============================================

valid, errors = xlsx.validate_style({
  bold = true,
  italic = true,
  font_color = "#FFFFFF",
  bg_color = "#4472C4",
  halign = "center",
  valign = "center",
  border = true,
  border_style = "thin",
  num_format = "percent",
})
h.test("complex valid style passes", valid == true)

valid, errors = xlsx.validate_style({
  bold = true,
  halign = "center",
  font_color = "xyz-not-valid",  -- invalid (not hex, not named)
})
h.test("complex style with one invalid field fails", valid == false)
h.test("only the invalid field reported", errors and #errors == 1)

-- ============================================
h.section("Test 13: Edge case - nil input")
-- ============================================

valid, errors = xlsx.validate_style(nil)
h.test("nil style is valid", valid == true)

-- ============================================
h.section("Test 14: Edge case - non-table input")
-- ============================================

valid, errors = xlsx.validate_style("not a table")
h.test("string input is invalid", valid == false)

valid, errors = xlsx.validate_style(123)
h.test("number input is invalid", valid == false)

-- ============================================
h.section("Test 15: Border edge table validation")
-- ============================================

valid, errors = xlsx.validate_style({
  border_left = { style = "thin", color = "#FF0000" }
})
h.test("border edge as table with style/color is valid", valid == true)

valid, errors = xlsx.validate_style({
  border_left = { style = "invalid-style" }
})
h.test("border edge with invalid style fails", valid == false)

valid, errors = xlsx.validate_style({
  border_left = { color = "not-a-color" }
})
h.test("border edge with invalid color fails", valid == false)

-- ============================================
h.section("Test 16: Helper functions")
-- ============================================

local halign_list = xlsx.Style.get_valid_halign()
h.test("get_valid_halign returns table", type(halign_list) == "table")
h.test("get_valid_halign contains center", vim.tbl_contains(halign_list, "center"))

local colors_list = xlsx.Style.get_valid_colors()
h.test("get_valid_colors returns table", type(colors_list) == "table")
h.test("get_valid_colors contains red", vim.tbl_contains(colors_list, "red"))

h.summary("Style Validation Tests")
