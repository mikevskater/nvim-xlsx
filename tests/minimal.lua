--- Minimal test for nvim-xlsx
--- Run with: nvim -l tests/minimal.lua

-- Add the lua directory to package path
local script_dir = debug.getinfo(1, "S").source:match("@(.*/)")
if script_dir then
  package.path = script_dir .. "../lua/?.lua;" .. script_dir .. "../lua/?/init.lua;" .. package.path
end

-- Also handle Windows paths
local script_path = debug.getinfo(1, "S").source:sub(2)
local base_dir = script_path:match("(.+)[/\\]tests[/\\]")
if base_dir then
  package.path = base_dir .. "/lua/?.lua;" .. base_dir .. "/lua/?/init.lua;" .. package.path
end

local function test(name, fn)
  local ok, err = pcall(fn)
  if ok then
    print("PASS: " .. name)
  else
    print("FAIL: " .. name)
    print("  " .. tostring(err))
  end
end

print("=== nvim-xlsx Minimal Tests ===\n")

-- Test column utilities
test("column.to_letter(1) == 'A'", function()
  local col = require("xlsx.utils.column")
  assert(col.to_letter(1) == "A", "Expected A, got " .. col.to_letter(1))
end)

test("column.to_letter(27) == 'AA'", function()
  local col = require("xlsx.utils.column")
  assert(col.to_letter(27) == "AA", "Expected AA, got " .. col.to_letter(27))
end)

test("column.to_number('A') == 1", function()
  local col = require("xlsx.utils.column")
  assert(col.to_number("A") == 1, "Expected 1, got " .. col.to_number("A"))
end)

test("column.to_number('AA') == 27", function()
  local col = require("xlsx.utils.column")
  assert(col.to_number("AA") == 27, "Expected 27, got " .. col.to_number("AA"))
end)

test("column.parse_ref('A1')", function()
  local col = require("xlsx.utils.column")
  local ref = col.parse_ref("A1")
  assert(ref.row == 1, "Expected row 1")
  assert(ref.col == 1, "Expected col 1")
end)

test("column.make_ref(1, 1) == 'A1'", function()
  local col = require("xlsx.utils.column")
  assert(col.make_ref(1, 1) == "A1", "Expected A1")
end)

-- Test XML writer
test("xml.escape handles special chars", function()
  local xml = require("xlsx.xml.writer")
  local escaped = xml.escape("<test>&\"'</test>")
  assert(escaped:find("&lt;"), "Should escape <")
  assert(escaped:find("&gt;"), "Should escape >")
  assert(escaped:find("&amp;"), "Should escape &")
end)

test("xml.element creates valid element", function()
  local xml = require("xlsx.xml.writer")
  local elem = xml.element("test", "content")
  assert(elem == "<test>content</test>", "Expected <test>content</test>, got " .. elem)
end)

test("xml.empty_element creates self-closing", function()
  local xml = require("xlsx.xml.writer")
  local elem = xml.empty_element("test", { a = "1" })
  assert(elem:find("<test"), "Should start with <test")
  assert(elem:find("/>"), "Should be self-closing")
end)

-- Test cell
test("Cell creation and XML generation", function()
  local Cell = require("xlsx.cell")
  local cell = Cell.new(1, 1, 42)
  local xml_str = cell:to_xml()
  assert(xml_str:find('r="A1"'), "Should have ref A1")
  assert(xml_str:find("<v>42</v>"), "Should have value 42")
end)

test("Cell with string value", function()
  local Cell = require("xlsx.cell")
  local cell = Cell.new(1, 1, "Hello")
  local xml_str = cell:to_xml()
  assert(xml_str:find('t="inlineStr"'), "Should have inline string type")
  assert(xml_str:find("<t>Hello</t>"), "Should have text content")
end)

-- Test worksheet
test("Worksheet creation", function()
  local WS = require("xlsx.worksheet")
  local sheet, err = WS.new("Test", 1, {})
  assert(sheet, "Should create worksheet: " .. tostring(err))
  assert(sheet.name == "Test", "Name should be Test")
end)

test("Worksheet set_cell", function()
  local WS = require("xlsx.worksheet")
  local sheet = WS.new("Test", 1, {})
  sheet:set_cell(1, 1, "A1")
  sheet:set_cell(2, 2, 123)
  assert(sheet.min_row == 1)
  assert(sheet.max_row == 2)
  assert(sheet.max_col == 2)
end)

test("Worksheet invalid name rejected", function()
  local WS = require("xlsx.worksheet")
  local sheet, err = WS.new("Test/Invalid", 1, {})
  assert(not sheet, "Should reject invalid name")
  assert(err:find("cannot contain"), "Should mention forbidden chars")
end)

-- Test workbook
test("Workbook creation", function()
  local WB = require("xlsx.workbook")
  local wb = WB.new()
  assert(wb, "Should create workbook")
  assert(#wb.sheets == 0, "Should start with no sheets")
end)

test("Workbook add_sheet", function()
  local WB = require("xlsx.workbook")
  local wb = WB.new()
  local sheet = wb:add_sheet("MySheet")
  assert(sheet, "Should add sheet")
  assert(sheet.name == "MySheet")
  assert(#wb.sheets == 1)
end)

print("\n=== Tests Complete ===")
