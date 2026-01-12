--- Color utilities for xlsx
--- Handles color format conversions for Excel
--- @module xlsx.utils.color

local M = {}

--- Named colors with their ARGB values
M.COLORS = {
  -- Basic colors
  black = "FF000000",
  white = "FFFFFFFF",
  red = "FFFF0000",
  green = "FF00FF00",
  blue = "FF0000FF",
  yellow = "FFFFFF00",
  cyan = "FF00FFFF",
  magenta = "FFFF00FF",

  -- Common colors
  orange = "FFFF6600",
  purple = "FF800080",
  pink = "FFFFC0CB",
  brown = "FF8B4513",
  gray = "FF808080",
  grey = "FF808080",
  lightgray = "FFD3D3D3",
  lightgrey = "FFD3D3D3",
  darkgray = "FFA9A9A9",
  darkgrey = "FFA9A9A9",

  -- Excel-like colors
  darkred = "FF8B0000",
  darkgreen = "FF006400",
  darkblue = "FF00008B",
  darkyellow = "FF9B870C",
  olive = "FF808000",
  teal = "FF008080",
  navy = "FF000080",
  maroon = "FF800000",

  -- Light colors (good for backgrounds)
  lightred = "FFFFCCCC",
  lightgreen = "FFCCFFCC",
  lightblue = "FFCCE5FF",
  lightyellow = "FFFFFFCC",
  lightorange = "FFFFD699",
  lightpurple = "FFE6CCFF",
}

--- Convert a color specification to Excel ARGB format
--- Accepts: "AARRGGBB", "RRGGBB", "#AARRGGBB", "#RRGGBB", named colors
--- @param color string|nil Color specification
--- @return string|nil ARGB color string (8 hex chars) or nil
function M.to_argb(color)
  if not color then
    return nil
  end

  -- Check named colors first
  local named = M.COLORS[color:lower()]
  if named then
    return named
  end

  -- Remove # prefix if present
  if color:sub(1, 1) == "#" then
    color = color:sub(2)
  end

  -- Validate hex characters
  if not color:match("^[0-9A-Fa-f]+$") then
    return nil
  end

  -- Handle different lengths
  local len = #color
  if len == 6 then
    -- RRGGBB -> FFRRGGBB (fully opaque)
    return "FF" .. color:upper()
  elseif len == 8 then
    -- AARRGGBB
    return color:upper()
  elseif len == 3 then
    -- RGB shorthand -> RRGGBB -> FFRRGGBB
    local r, g, b = color:sub(1,1), color:sub(2,2), color:sub(3,3)
    return "FF" .. r:upper() .. r:upper() .. g:upper() .. g:upper() .. b:upper() .. b:upper()
  end

  return nil
end

--- Parse ARGB into components
--- @param argb string ARGB color string
--- @return table {a: number, r: number, g: number, b: number} (0-255 each)
function M.parse_argb(argb)
  if not argb or #argb ~= 8 then
    return { a = 255, r = 0, g = 0, b = 0 }
  end

  return {
    a = tonumber(argb:sub(1, 2), 16) or 255,
    r = tonumber(argb:sub(3, 4), 16) or 0,
    g = tonumber(argb:sub(5, 6), 16) or 0,
    b = tonumber(argb:sub(7, 8), 16) or 0,
  }
end

--- Create ARGB from RGB components
--- @param r number Red (0-255)
--- @param g number Green (0-255)
--- @param b number Blue (0-255)
--- @param a? number Alpha (0-255, default 255)
--- @return string ARGB color string
function M.from_rgb(r, g, b, a)
  a = a or 255
  return string.format("%02X%02X%02X%02X", a, r, g, b)
end

--- Apply tint to a color (Excel uses tints for theme color variations)
--- @param argb string Base ARGB color
--- @param tint number Tint value (-1 to 1, negative = darker, positive = lighter)
--- @return string Modified ARGB color
function M.apply_tint(argb, tint)
  if not tint or tint == 0 then
    return argb
  end

  local c = M.parse_argb(argb)

  local function apply_tint_component(value, t)
    if t < 0 then
      -- Darken
      return math.floor(value * (1 + t))
    else
      -- Lighten
      return math.floor(value + (255 - value) * t)
    end
  end

  return M.from_rgb(
    apply_tint_component(c.r, tint),
    apply_tint_component(c.g, tint),
    apply_tint_component(c.b, tint),
    c.a
  )
end

return M
