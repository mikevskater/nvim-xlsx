--- Tests for styles: fonts, fills, borders, number formats, alignment
--- Run with: nvim --headless -l tests/test_styles.lua

dofile(arg[0]:match("(.*/)").. "test_helper.lua")
local h = require("tests.test_helper")
local xlsx = h.require_xlsx()

h.suite("Style Tests")

local wb = xlsx.new_workbook()

-- ============================================
h.section("Test 1: Style constants accessible")
-- ============================================

h.test("BORDER_STYLES exists", xlsx.BORDER_STYLES ~= nil)
h.test("BORDER_STYLES.thin", xlsx.BORDER_STYLES.thin == "thin")
h.test("BORDER_STYLES.medium", xlsx.BORDER_STYLES.medium == "medium")
h.test("BORDER_STYLES.thick", xlsx.BORDER_STYLES.thick == "thick")

h.test("HALIGN exists", xlsx.HALIGN ~= nil)
h.test("HALIGN.left", xlsx.HALIGN.left == "left")
h.test("HALIGN.center", xlsx.HALIGN.center == "center")
h.test("HALIGN.right", xlsx.HALIGN.right == "right")

h.test("VALIGN exists", xlsx.VALIGN ~= nil)
h.test("VALIGN.top", xlsx.VALIGN.top == "top")
h.test("VALIGN.center", xlsx.VALIGN.center == "center")
h.test("VALIGN.bottom", xlsx.VALIGN.bottom == "bottom")

h.test("UNDERLINE exists", xlsx.UNDERLINE ~= nil)
h.test("BUILTIN_FORMATS exists", xlsx.BUILTIN_FORMATS ~= nil)
h.test("BUILTIN_FORMATS.date", xlsx.BUILTIN_FORMATS.date == 14)

-- ============================================
h.section("Test 2: Font styles")
-- ============================================

local sheet = wb:add_sheet("Fonts")

local bold_style = wb:create_style({ bold = true })
h.test("bold style created", bold_style > 0)

local italic_style = wb:create_style({ italic = true })
h.test("italic style created", italic_style > 0)

local underline_style = wb:create_style({ underline = "single" })
h.test("underline style created", underline_style > 0)

local strike_style = wb:create_style({ strike = true })
h.test("strikethrough style created", strike_style > 0)

local font_color_style = wb:create_style({ font_color = "#FF0000" })
h.test("font color style created", font_color_style > 0)

local font_size_style = wb:create_style({ font_size = 16 })
h.test("font size style created", font_size_style > 0)

local font_name_style = wb:create_style({ font_name = "Arial" })
h.test("font name style created", font_name_style > 0)

-- Apply styles
sheet:set_cell(1, 1, "Bold")
sheet:set_cell_style(1, 1, bold_style)
sheet:set_cell(2, 1, "Italic")
sheet:set_cell_style(2, 1, italic_style)

-- ============================================
h.section("Test 3: Fill styles")
-- ============================================

local fill_sheet = wb:add_sheet("Fills")

local bg_style = wb:create_style({ bg_color = "#FFFF00" })
h.test("background color style created", bg_style > 0)

local fill_style = wb:create_style({ fill_color = "#00FF00" })
h.test("fill color style created", fill_style > 0)

fill_sheet:set_cell(1, 1, "Yellow BG")
fill_sheet:set_cell_style(1, 1, bg_style)

-- ============================================
h.section("Test 4: Border styles")
-- ============================================

local border_sheet = wb:add_sheet("Borders")

local all_border = wb:create_style({ border = true })
h.test("all borders style created", all_border > 0)

local thin_border = wb:create_style({ border = true, border_style = "thin" })
h.test("thin border style created", thin_border > 0)

local medium_border = wb:create_style({ border = true, border_style = "medium" })
h.test("medium border style created", medium_border > 0)

local colored_border = wb:create_style({
  border = true,
  border_style = "thin",
  border_color = "#0000FF"
})
h.test("colored border style created", colored_border > 0)

local partial_border = wb:create_style({
  border_top = "thin",
  border_bottom = "thick"
})
h.test("partial border style created", partial_border > 0)

border_sheet:set_cell(1, 1, "All borders")
border_sheet:set_cell_style(1, 1, all_border)

-- ============================================
h.section("Test 5: Number formats")
-- ============================================

local format_sheet = wb:add_sheet("Formats")

local percent_style = wb:create_style({ num_format = "percent" })
h.test("percent format style created", percent_style > 0)

local date_style = wb:create_style({ num_format = "date" })
h.test("date format style created", date_style > 0)

local currency_style = wb:create_style({ num_format = "$#,##0.00" })
h.test("custom currency format created", currency_style > 0)

local builtin_style = wb:create_style({ num_format = 14 })
h.test("builtin format by ID created", builtin_style > 0)

format_sheet:set_cell(1, 1, 0.75)
format_sheet:set_cell_style(1, 1, percent_style)
format_sheet:set_cell(2, 1, 45000)
format_sheet:set_cell_style(2, 1, date_style)

-- ============================================
h.section("Test 6: Alignment")
-- ============================================

local align_sheet = wb:add_sheet("Alignment")

local center_style = wb:create_style({ halign = "center" })
h.test("horizontal center style created", center_style > 0)

local vcenter_style = wb:create_style({ valign = "center" })
h.test("vertical center style created", vcenter_style > 0)

local wrap_style = wb:create_style({ wrap_text = true })
h.test("wrap text style created", wrap_style > 0)

local rotate_style = wb:create_style({ rotation = 45 })
h.test("rotation style created", rotate_style > 0)

local indent_style = wb:create_style({ indent = 2 })
h.test("indent style created", indent_style > 0)

align_sheet:set_cell(1, 1, "Centered")
align_sheet:set_cell_style(1, 1, center_style)

-- ============================================
h.section("Test 7: Combined styles")
-- ============================================

local combined_sheet = wb:add_sheet("Combined")

local header_style = wb:create_style({
  bold = true,
  bg_color = "#4472C4",
  font_color = "#FFFFFF",
  halign = "center",
  border = true,
  border_style = "thin"
})
h.test("combined header style created", header_style > 0)

combined_sheet:set_cell(1, 1, "Header")
combined_sheet:set_cell_style(1, 1, header_style)

-- ============================================
h.section("Test 8: Style deduplication")
-- ============================================

local style_a = wb:create_style({ bold = true })
local style_b = wb:create_style({ bold = true })
h.test("duplicate styles return same index", style_a == style_b)

local style_c = wb:create_style({ italic = true })
h.test("different styles return different index", style_a ~= style_c)

-- ============================================
h.section("Test 9: Range styling")
-- ============================================

local range_sheet = wb:add_sheet("Range")
local range_style = wb:create_style({ bg_color = "#E0E0E0" })

range_sheet:set_range_style(1, 1, 3, 3, range_style)
local corner_cell = range_sheet:get_cell(1, 1)
h.test("range style applied to corner", corner_cell ~= nil and corner_cell.style_index == range_style)

-- ============================================
h.section("Test 10: Save and verify")
-- ============================================

local output_path = h.fixtures_dir .. "/test_styles.xlsx"
local ok, err = wb:save(output_path)
h.test("save succeeds", ok == true)

h.summary("Style Tests")
