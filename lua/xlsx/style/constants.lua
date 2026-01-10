--- Style constants for xlsx
--- Built-in formats, border styles, alignment options
--- @module xlsx.style.constants

local M = {}

-- Built-in number format IDs (0-163 are reserved)
M.BUILTIN_FORMATS = {
  general = 0,
  number = 1,           -- 0
  number_d2 = 2,        -- 0.00
  number_thousands = 3, -- #,##0
  number_thousands_d2 = 4, -- #,##0.00
  percent = 9,          -- 0%
  percent_d2 = 10,      -- 0.00%
  scientific = 11,      -- 0.00E+00
  fraction = 12,        -- # ?/?
  fraction_d2 = 13,     -- # ??/??
  date = 14,            -- m/d/yyyy (locale dependent)
  date_d_mon_yy = 15,   -- d-mmm-yy
  date_d_mon = 16,      -- d-mmm
  date_mon_yy = 17,     -- mmm-yy
  time_12h = 18,        -- h:mm AM/PM
  time_12h_ss = 19,     -- h:mm:ss AM/PM
  time_24h = 20,        -- h:mm
  time_24h_ss = 21,     -- h:mm:ss
  datetime = 22,        -- m/d/yyyy h:mm
  accounting = 37,      -- #,##0_);(#,##0)
  accounting_red = 38,  -- #,##0_);[Red](#,##0)
  accounting_d2 = 39,   -- #,##0.00_);(#,##0.00)
  accounting_d2_red = 40, -- #,##0.00_);[Red](#,##0.00)
  text = 49,            -- @
}

-- Border styles
M.BORDER_STYLES = {
  none = nil,
  thin = "thin",
  medium = "medium",
  thick = "thick",
  dashed = "dashed",
  dotted = "dotted",
  double = "double",
  hair = "hair",
  mediumDashed = "mediumDashed",
  dashDot = "dashDot",
  mediumDashDot = "mediumDashDot",
  dashDotDot = "dashDotDot",
  mediumDashDotDot = "mediumDashDotDot",
  slantDashDot = "slantDashDot",
}

-- Horizontal alignment
M.HALIGN = {
  left = "left",
  center = "center",
  right = "right",
  fill = "fill",
  justify = "justify",
  centerContinuous = "centerContinuous",
  distributed = "distributed",
}

-- Vertical alignment
M.VALIGN = {
  top = "top",
  center = "center",
  bottom = "bottom",
  justify = "justify",
  distributed = "distributed",
}

-- Underline styles
M.UNDERLINE = {
  none = nil,
  single = "single",
  double = "double",
  singleAccounting = "singleAccounting",
  doubleAccounting = "doubleAccounting",
}

return M
