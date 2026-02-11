--- Column utilities for Excel column/cell reference handling
--- @module nvim-xlsx.utils.column

local validation = require("nvim-xlsx.utils.validation")
local LIMITS = validation.LIMITS

local M = {}

--- Convert a column number to Excel letter notation
--- @param num integer Column number (1-indexed)
--- @return string Column letter(s) (A, B, ..., Z, AA, AB, ...)
function M.to_letter(num)
  if num < 1 then
    error("Column number must be >= 1, got: " .. tostring(num))
  end
  if num > LIMITS.MAX_COLS then
    error("Column number exceeds Excel maximum (" .. LIMITS.MAX_COLS .. "), got: " .. tostring(num))
  end

  local result = ""
  while num > 0 do
    local remainder = (num - 1) % 26
    result = string.char(65 + remainder) .. result
    num = math.floor((num - 1) / 26)
  end
  return result
end

--- Convert Excel column letter notation to number
--- @param str string Column letter(s) (A, B, ..., Z, AA, AB, ...)
--- @return integer Column number (1-indexed)
function M.to_number(str)
  if not str or str == "" then
    error("Column letter cannot be empty")
  end

  str = str:upper()
  local num = 0
  for i = 1, #str do
    local char = str:sub(i, i)
    if char < "A" or char > "Z" then
      error("Invalid column letter: " .. str)
    end
    num = num * 26 + (string.byte(char) - 64)
  end

  if num > LIMITS.MAX_COLS then
    error("Column exceeds Excel maximum (XFD/" .. LIMITS.MAX_COLS .. "): " .. str)
  end

  return num
end

--- Parse a cell reference into row and column
--- @param ref string Cell reference (e.g., "A1", "$A$1", "AA100")
--- @return table {row: integer, col: integer, abs_row: boolean, abs_col: boolean}
function M.parse_ref(ref)
  if not ref or ref == "" then
    error("Cell reference cannot be empty")
  end

  local abs_col = ref:sub(1, 1) == "$"
  if abs_col then
    ref = ref:sub(2)
  end

  local col_str, rest = ref:match("^([A-Za-z]+)(.*)$")
  if not col_str then
    error("Invalid cell reference: " .. ref)
  end

  local abs_row = rest:sub(1, 1) == "$"
  if abs_row then
    rest = rest:sub(2)
  end

  local row_str = rest:match("^(%d+)$")
  if not row_str then
    error("Invalid cell reference: " .. ref)
  end

  local row = tonumber(row_str)
  if row < 1 or row > LIMITS.MAX_ROWS then
    error("Row number out of range (1-" .. LIMITS.MAX_ROWS .. "): " .. row_str)
  end

  return {
    row = row,
    col = M.to_number(col_str),
    abs_row = abs_row,
    abs_col = abs_col,
  }
end

--- Create a cell reference from row and column numbers
--- @param row integer Row number (1-indexed)
--- @param col integer Column number (1-indexed)
--- @param abs_row? boolean Make row absolute ($)
--- @param abs_col? boolean Make column absolute ($)
--- @return string Cell reference (e.g., "A1", "$A$1")
function M.make_ref(row, col, abs_row, abs_col)
  if row < 1 or row > LIMITS.MAX_ROWS then
    error("Row number out of range (1-" .. LIMITS.MAX_ROWS .. "): " .. tostring(row))
  end

  local ref = ""
  if abs_col then
    ref = "$"
  end
  ref = ref .. M.to_letter(col)
  if abs_row then
    ref = ref .. "$"
  end
  ref = ref .. tostring(row)
  return ref
end

--- Parse a range reference into start and end cells
--- @param range string Range reference (e.g., "A1:B10", "$A$1:$B$10")
--- @return table {start: {row, col}, finish: {row, col}}
function M.parse_range(range)
  local start_ref, end_ref = range:match("^([^:]+):([^:]+)$")
  if not start_ref then
    error("Invalid range reference: " .. range)
  end

  return {
    start = M.parse_ref(start_ref),
    finish = M.parse_ref(end_ref),
  }
end

--- Create a range reference from coordinates
--- @param r1 integer Start row
--- @param c1 integer Start column
--- @param r2 integer End row
--- @param c2 integer End column
--- @return string Range reference (e.g., "A1:B10")
function M.make_range(r1, c1, r2, c2)
  return M.make_ref(r1, c1) .. ":" .. M.make_ref(r2, c2)
end

--- Create an absolute range reference (e.g., "$A$1:$C$10")
--- @param r1 integer Start row
--- @param c1 integer Start column
--- @param r2 integer End row
--- @param c2 integer End column
--- @return string Absolute range reference
function M.make_abs_range(r1, c1, r2, c2)
  return M.make_ref(r1, c1, true, true) .. ":" .. M.make_ref(r2, c2, true, true)
end

return M
