--- Tests for hyperlinks
--- Run with: nvim --headless -l tests/test_hyperlinks.lua

dofile(arg[0]:match("(.*/)").. "test_helper.lua")
local h = require("tests.test_helper")
local xlsx = h.require_xlsx()

h.suite("Hyperlink Tests")

local wb = xlsx.new_workbook()

-- ============================================
h.section("Test 1: Basic URL hyperlink")
-- ============================================

local sheet1 = wb:add_sheet("URLs")
sheet1:set_cell(1, 1, "Click here")

-- add_hyperlink(row, col, target, options)
local ok = sheet1:add_hyperlink(1, 1, "https://example.com")
h.test("add_hyperlink succeeds", ok ~= nil)
h.test("hyperlinks populated", #sheet1.hyperlinks > 0)

local link = sheet1.hyperlinks[1]
h.test("hyperlink ref is A1", link.ref == "A1")
h.test("hyperlink target is URL", link.target == "https://example.com")

-- ============================================
h.section("Test 2: Hyperlink with display text")
-- ============================================

local sheet2 = wb:add_sheet("Display")
local ok2 = sheet2:add_hyperlink(1, 1, "https://google.com", { display_text = "Google" })
h.test("hyperlink with display succeeds", ok2 ~= nil)

local link2 = sheet2.hyperlinks[1]
h.test("display text set", link2.display == "Google")

-- Cell should also have the display text
local cell2 = sheet2:get_cell(1, 1)
h.test("cell has display text", cell2 and cell2.value == "Google")

-- ============================================
h.section("Test 3: Hyperlink with tooltip")
-- ============================================

local sheet3 = wb:add_sheet("Tooltip")
local ok3 = sheet3:add_hyperlink(2, 2, "https://github.com", { tooltip = "Visit GitHub" })
h.test("hyperlink with tooltip succeeds", ok3 ~= nil)

local link3 = sheet3.hyperlinks[1]
h.test("tooltip set", link3.tooltip == "Visit GitHub")

-- ============================================
h.section("Test 4: Internal reference hyperlink")
-- ============================================

local sheet4 = wb:add_sheet("Internal")
local ok4 = sheet4:add_hyperlink(1, 1, "Sheet2!A1")
h.test("internal hyperlink succeeds", ok4 ~= nil)

local link4 = sheet4.hyperlinks[1]
h.test("location set for internal", link4.location == "Sheet2!A1")
h.test("not marked as external", link4.is_external == false)

-- ============================================
h.section("Test 5: Email hyperlink")
-- ============================================

local sheet5 = wb:add_sheet("Email")
local ok5 = sheet5:add_hyperlink(1, 1, "mailto:test@example.com")
h.test("email hyperlink succeeds", ok5 ~= nil)

local link5 = sheet5.hyperlinks[1]
h.test("email target set", link5.target:match("mailto:") ~= nil)
h.test("email is external", link5.is_external == true)

-- ============================================
h.section("Test 6: File hyperlink")
-- ============================================

local sheet6 = wb:add_sheet("File")
local ok6 = sheet6:add_hyperlink(1, 1, "file:///C:/Documents/report.pdf")
h.test("file hyperlink succeeds", ok6 ~= nil)

local link6 = sheet6.hyperlinks[1]
h.test("file is external", link6.is_external == true)

-- ============================================
h.section("Test 7: Multiple hyperlinks")
-- ============================================

local sheet7 = wb:add_sheet("Multiple")
sheet7:add_hyperlink(1, 1, "https://site1.com")
sheet7:add_hyperlink(2, 1, "https://site2.com")
sheet7:add_hyperlink(3, 1, "https://site3.com")

h.test("3 hyperlinks added", #sheet7.hyperlinks == 3)

-- ============================================
h.section("Test 8: Hyperlink chaining")
-- ============================================

local sheet8 = wb:add_sheet("Chaining")
local result = sheet8:add_hyperlink(1, 1, "https://test.com")
h.test("add_hyperlink returns worksheet", result == sheet8)
result:set_cell(1, 2, "Data")
h.test("chained set_cell works", sheet8:get_cell(1, 2).value == "Data")

-- ============================================
h.section("Test 9: Hyperlink with all options")
-- ============================================

local sheet9 = wb:add_sheet("AllOptions")
local ok9 = sheet9:add_hyperlink(3, 3, "https://full.example.com", {
  display_text = "Full Example",
  tooltip = "Click to visit full example",
})
h.test("full options hyperlink succeeds", ok9 ~= nil)

local link9 = sheet9.hyperlinks[1]
h.test("all options preserved",
  link9.display == "Full Example" and
  link9.tooltip == "Click to visit full example")

-- ============================================
h.section("Test 10: XML generation")
-- ============================================

local sheet10 = wb:add_sheet("XMLTest")
sheet10:add_hyperlink(1, 1, "https://xml.test.com", { display_text = "Test" })
local xml = sheet10:to_xml()

h.test("XML contains hyperlinks element", xml:match("<hyperlinks") ~= nil)
h.test("XML contains hyperlink element", xml:match("<hyperlink") ~= nil)

-- ============================================
h.section("Test 11: No hyperlinks (default)")
-- ============================================

local sheet11 = wb:add_sheet("NoLinks")
sheet11:set_cell(1, 1, "Normal")
local xml11 = sheet11:to_xml()

h.test("default has no hyperlinks in XML", xml11:match("<hyperlinks") == nil)

-- ============================================
h.section("Test 12: Save and verify")
-- ============================================

local output_path = h.fixtures_dir .. "/test_hyperlinks.xlsx"
local save_ok = wb:save(output_path)
h.test("save succeeds", save_ok == true)

-- Read back
local wb2 = xlsx.open(output_path)
h.test("file opens", wb2 ~= nil)

local read_sheet = xlsx.get_sheet(wb2, "URLs")
h.test("hyperlink sheet exists", read_sheet ~= nil)

h.summary("Hyperlink Tests")
