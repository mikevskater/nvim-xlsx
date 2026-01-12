--- Worksheet module - combines core, features, and xml submodules
--- @module nvim-xlsx.worksheet
---
--- This module provides backward compatibility by exposing the same API
--- as the original worksheet.lua while internally splitting the code
--- into smaller, focused modules.
---
--- ## Error Handling Patterns
---
--- The Worksheet API uses two error handling patterns:
---
--- ### 1. Throws error() - For programmer errors
--- These methods throw errors for invalid arguments or out-of-range values.
--- Use pcall() if you need to catch these errors.
---
--- Methods that throw:
---   - set_cell(), set(), set_cell_value(), set_formula(), set_date(), set_boolean()
---   - freeze_panes(), freeze_rows(), freeze_cols()
---   - set_auto_filter(), set_auto_filter_range()
---   - add_data_validation(), add_dropdown(), add_number_validation()
---   - set_orientation()
---
--- ### 2. Returns nil/false, error - For recoverable errors
--- These methods return error information that the caller can handle gracefully.
---
--- Methods that return errors:
---   - M.new(name, index, workbook) -> Worksheet?, error_message?
---   - merge_cells(r1, c1, r2, c2) -> boolean, error_message?
---   - merge_range(range) -> boolean, error_message?
---
--- ### 3. Returns nil - For lookups (not an error)
--- These methods return nil when no data is found, which is expected behavior.
---
---   - get_cell(row, col) -> Cell?
---   - get(ref) -> Cell?
---
--- ## Method Chaining
---
--- Feature methods support chaining by returning self:
---   sheet:freeze_rows(1):set_auto_filter(1, 1, 10, 5):set_orientation("landscape")
---
--- Cell methods return the Cell object instead (for further cell manipulation):
---   local cell = sheet:set_cell(1, 1, "Hello")
---   cell.style_index = my_style
---
--- Chainable methods (return Worksheet):
---   - freeze_panes(), freeze_rows(), freeze_cols()
---   - set_auto_filter(), set_auto_filter_range()
---   - add_data_validation(), add_dropdown(), add_number_validation()
---   - add_hyperlink()
---   - set_print_settings(), set_margins(), set_orientation()
---   - set_print_area(), set_print_title_rows(), set_print_title_cols()
---
--- Cell methods (return Cell):
---   - set_cell(), set(), set_cell_value(), set_formula(), set_date(), set_boolean()
---

local core = require("nvim-xlsx.worksheet.core")
local features = require("nvim-xlsx.worksheet.features")
local xml_gen = require("nvim-xlsx.worksheet.xml")

local M = {}

-- Get the Worksheet class from core
local Worksheet = core.Worksheet

-- ============================================
-- Inject feature methods into Worksheet class
-- ============================================

-- Freeze panes
function Worksheet:freeze_panes(rows, cols)
  return features.freeze_panes(self, rows, cols)
end

function Worksheet:freeze_rows(rows)
  return features.freeze_rows(self, rows)
end

function Worksheet:freeze_cols(cols)
  return features.freeze_cols(self, cols)
end

-- Auto-filter
function Worksheet:set_auto_filter(r1, c1, r2, c2)
  return features.set_auto_filter(self, r1, c1, r2, c2)
end

function Worksheet:set_auto_filter_range(range)
  return features.set_auto_filter_range(self, range)
end

-- Data validation
function Worksheet:add_data_validation(ref, validation_opts)
  return features.add_data_validation(self, ref, validation_opts)
end

function Worksheet:add_dropdown(ref, items, options)
  return features.add_dropdown(self, ref, items, options)
end

function Worksheet:add_number_validation(ref, min, max, options)
  return features.add_number_validation(self, ref, min, max, options)
end

-- Hyperlinks
function Worksheet:add_hyperlink(row, col, target, options)
  return features.add_hyperlink(self, row, col, target, options)
end

-- Print settings
function Worksheet:set_print_settings(settings)
  return features.set_print_settings(self, settings)
end

function Worksheet:_ensure_print_settings()
  return features._ensure_print_settings(self)
end

function Worksheet:set_margins(top, bottom, left, right, header, footer)
  return features.set_margins(self, top, bottom, left, right, header, footer)
end

function Worksheet:set_orientation(orientation)
  return features.set_orientation(self, orientation)
end

function Worksheet:set_print_area(range)
  return features.set_print_area(self, range)
end

function Worksheet:set_print_title_rows(rows)
  return features.set_print_title_rows(self, rows)
end

function Worksheet:set_print_title_cols(cols)
  return features.set_print_title_cols(self, cols)
end

-- ============================================
-- Inject XML generation methods into Worksheet class
-- ============================================

function Worksheet:_generate_sheet_data()
  return xml_gen._generate_sheet_data(self)
end

function Worksheet:_generate_cols()
  return xml_gen._generate_cols(self)
end

function Worksheet:_generate_sheet_views(is_active)
  return xml_gen._generate_sheet_views(self, is_active)
end

function Worksheet:_generate_auto_filter()
  return xml_gen._generate_auto_filter(self)
end

function Worksheet:_generate_data_validations()
  return xml_gen._generate_data_validations(self)
end

function Worksheet:_generate_hyperlinks(hyperlink_rels)
  return xml_gen._generate_hyperlinks(self, hyperlink_rels)
end

function Worksheet:_generate_print_settings()
  return xml_gen._generate_print_settings(self)
end

function Worksheet:get_hyperlink_relationships()
  return xml_gen.get_hyperlink_relationships(self)
end

function Worksheet:to_xml(is_active)
  return xml_gen.to_xml(self, is_active)
end

-- ============================================
-- Export module API (backward compatible)
-- ============================================

M.new = core.new
M.Worksheet = Worksheet

return M
