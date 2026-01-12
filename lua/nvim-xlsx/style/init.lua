--- Style module - combines constants, registry, and xml submodules
--- @module nvim-xlsx.style
---
--- This module provides backward compatibility by exposing the same API
--- as the original style.lua while internally splitting the code
--- into smaller, focused modules.

local constants = require("nvim-xlsx.style.constants")
local registry = require("nvim-xlsx.style.registry")
local xml_gen = require("nvim-xlsx.style.xml")

local M = {}

-- Get the StyleRegistry class from registry module
local StyleRegistry = registry.StyleRegistry

-- ============================================
-- Inject XML generation methods into StyleRegistry class
-- ============================================

function StyleRegistry:_font_to_xml(font)
  return xml_gen._font_to_xml(self, font)
end

function StyleRegistry:_fill_to_xml(fill)
  return xml_gen._fill_to_xml(self, fill)
end

function StyleRegistry:_border_edge_to_xml(edge, def)
  return xml_gen._border_edge_to_xml(self, edge, def)
end

function StyleRegistry:_border_to_xml(border)
  return xml_gen._border_to_xml(self, border)
end

function StyleRegistry:_xf_to_xml(xf, for_style)
  return xml_gen._xf_to_xml(self, xf, for_style)
end

function StyleRegistry:to_xml()
  return xml_gen.to_xml(self)
end

-- ============================================
-- Export module API (backward compatible)
-- ============================================

-- Re-export constants at top level
M.BUILTIN_FORMATS = constants.BUILTIN_FORMATS
M.BORDER_STYLES = constants.BORDER_STYLES
M.HALIGN = constants.HALIGN
M.VALIGN = constants.VALIGN
M.UNDERLINE = constants.UNDERLINE

-- Export constructor and class
M.new_registry = registry.new_registry
M.StyleRegistry = StyleRegistry

return M
