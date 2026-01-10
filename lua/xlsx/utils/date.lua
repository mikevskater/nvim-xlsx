--- Date utilities for xlsx
--- @module xlsx.utils.date
---
--- Excel date system:
--- - Day 1 = January 1, 1900
--- - Excel incorrectly treats 1900 as a leap year (Lotus 1-2-3 compatibility)
--- - Day 60 = Feb 29, 1900 (doesn't exist but Excel thinks it does)
--- - For dates >= March 1, 1900, serial numbers are off by 1 day

local M = {}

--- Excel epoch: January 1, 1900 (but see leap year bug notes)
--- We use December 30, 1899 as the reference because:
--- Excel day 1 = Jan 1, 1900, but due to the 1900 leap year bug,
--- serial numbers >= 60 are off by 1 day
local EXCEL_EPOCH_OFFSET = 25569  -- Days between Unix epoch (1970) and Excel epoch (1900) adjusted

--- Seconds in a day
local SECONDS_PER_DAY = 86400

--- Check if a year is a leap year (correct algorithm)
--- @param year integer
--- @return boolean
function M.is_leap_year(year)
  if year % 400 == 0 then
    return true
  elseif year % 100 == 0 then
    return false
  elseif year % 4 == 0 then
    return true
  else
    return false
  end
end

--- Days in each month (non-leap year)
local DAYS_IN_MONTH = { 31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31 }

--- Get days in a specific month
--- @param year integer
--- @param month integer (1-12)
--- @return integer
function M.days_in_month(year, month)
  if month == 2 and M.is_leap_year(year) then
    return 29
  end
  return DAYS_IN_MONTH[month]
end

--- Convert a Lua date table to Excel serial number
--- @param date table Date table with year, month, day, (optional: hour, min, sec)
--- @return number Excel serial number (with fractional time component if applicable)
function M.to_serial(date)
  local year = date.year
  local month = date.month
  local day = date.day
  local hour = date.hour or 0
  local min = date.min or 0
  local sec = date.sec or 0

  -- Validate
  if year < 1900 then
    error("Excel dates must be >= 1900")
  end

  -- Calculate days from Jan 1, 1900
  local serial = 0

  -- Add days for complete years from 1900 to year-1
  for y = 1900, year - 1 do
    serial = serial + (M.is_leap_year(y) and 366 or 365)
  end

  -- Add days for complete months in current year
  for m = 1, month - 1 do
    serial = serial + M.days_in_month(year, m)
  end

  -- Add days in current month
  serial = serial + day

  -- Excel 1900 leap year bug: Excel thinks Feb 29, 1900 exists
  -- For dates on or after March 1, 1900, we need to add 1
  -- because Excel's serial numbers are off by 1 starting from day 60
  if serial >= 60 then
    serial = serial + 1
  end

  -- Add time component as fractional day
  local time_fraction = (hour * 3600 + min * 60 + sec) / SECONDS_PER_DAY

  return serial + time_fraction
end

--- Convert an Excel serial number to a Lua date table
--- @param serial number Excel serial number
--- @return table Date table with year, month, day, hour, min, sec
function M.from_serial(serial)
  if serial < 1 then
    error("Excel serial must be >= 1")
  end

  -- Extract time component
  local days = math.floor(serial)
  local time_fraction = serial - days

  -- Handle Excel 1900 leap year bug
  -- Excel thinks day 60 = Feb 29, 1900 (doesn't exist)
  -- For days >= 61, subtract 1 to get the correct date
  local adjusted_days = days
  if days >= 61 then
    adjusted_days = days - 1
  elseif days == 60 then
    -- This is the fictitious Feb 29, 1900
    return {
      year = 1900,
      month = 2,
      day = 29,  -- Doesn't really exist, but that's what Excel thinks
      hour = 0,
      min = 0,
      sec = 0,
    }
  end

  -- Calculate date from adjusted days (day 1 = Jan 1, 1900)
  local year = 1900
  local remaining = adjusted_days

  -- Subtract days for complete years
  while true do
    local days_in_year = M.is_leap_year(year) and 366 or 365
    if remaining <= days_in_year then
      break
    end
    remaining = remaining - days_in_year
    year = year + 1
  end

  -- Subtract days for complete months
  local month = 1
  while month <= 12 do
    local days_in_m = M.days_in_month(year, month)
    if remaining <= days_in_m then
      break
    end
    remaining = remaining - days_in_m
    month = month + 1
  end

  local day = remaining

  -- Calculate time from fraction
  local total_seconds = math.floor(time_fraction * SECONDS_PER_DAY + 0.5)
  local hour = math.floor(total_seconds / 3600)
  local min = math.floor((total_seconds % 3600) / 60)
  local sec = total_seconds % 60

  return {
    year = year,
    month = month,
    day = day,
    hour = hour,
    min = min,
    sec = sec,
  }
end

--- Convert a Unix timestamp to Excel serial number
--- @param timestamp integer Unix timestamp (seconds since 1970-01-01)
--- @return number Excel serial number
function M.from_unix_timestamp(timestamp)
  -- Convert to Excel serial
  -- Unix epoch = Jan 1, 1970 = Excel day 25569 (accounting for leap year bug)
  local days = timestamp / SECONDS_PER_DAY
  return days + EXCEL_EPOCH_OFFSET
end

--- Convert an Excel serial number to Unix timestamp
--- @param serial number Excel serial number
--- @return integer Unix timestamp
function M.to_unix_timestamp(serial)
  local days = serial - EXCEL_EPOCH_OFFSET
  return math.floor(days * SECONDS_PER_DAY)
end

--- Parse an ISO 8601 date string to Excel serial number
--- Supports formats: YYYY-MM-DD, YYYY-MM-DDTHH:MM:SS, YYYY-MM-DD HH:MM:SS
--- @param str string ISO date string
--- @return number? serial Excel serial number, or nil if parse fails
--- @return string? error Error message if parse fails
function M.parse_iso(str)
  -- Try YYYY-MM-DD with optional time
  local year, month, day, hour, min, sec = str:match("^(%d%d%d%d)%-(%d%d)%-(%d%d)[T ]?(%d?%d?):?(%d?%d?):?(%d?%d?)$")

  if not year then
    -- Try without dashes: YYYYMMDD
    year, month, day = str:match("^(%d%d%d%d)(%d%d)(%d%d)$")
    hour, min, sec = "0", "0", "0"
  end

  if not year then
    return nil, "Invalid date format: " .. str
  end

  year = tonumber(year)
  month = tonumber(month)
  day = tonumber(day)
  hour = tonumber(hour) or 0
  min = tonumber(min) or 0
  sec = tonumber(sec) or 0

  -- Validate
  if month < 1 or month > 12 then
    return nil, "Invalid month: " .. month
  end
  if day < 1 or day > M.days_in_month(year, month) then
    return nil, "Invalid day: " .. day .. " for month " .. month
  end

  return M.to_serial({
    year = year,
    month = month,
    day = day,
    hour = hour,
    min = min,
    sec = sec,
  })
end

--- Format an Excel serial number as ISO date string
--- @param serial number Excel serial number
--- @param include_time? boolean Whether to include time (default: false)
--- @return string ISO date string
function M.format_iso(serial, include_time)
  local date = M.from_serial(serial)
  local iso = string.format("%04d-%02d-%02d", date.year, date.month, date.day)

  if include_time then
    iso = iso .. string.format("T%02d:%02d:%02d", date.hour, date.min, date.sec)
  end

  return iso
end

--- Get current date/time as Excel serial number
--- @return number Excel serial number
function M.now()
  local now = os.date("*t")
  return M.to_serial({
    year = now.year,
    month = now.month,
    day = now.day,
    hour = now.hour,
    min = now.min,
    sec = now.sec,
  })
end

--- Get today's date as Excel serial number (no time component)
--- @return number Excel serial number (integer)
function M.today()
  local now = os.date("*t")
  return M.to_serial({
    year = now.year,
    month = now.month,
    day = now.day,
  })
end

--- Create a date serial from components
--- @param year integer Year (>= 1900)
--- @param month integer Month (1-12)
--- @param day integer Day (1-31)
--- @param hour? integer Hour (0-23)
--- @param min? integer Minute (0-59)
--- @param sec? integer Second (0-59)
--- @return number Excel serial number
function M.date(year, month, day, hour, min, sec)
  return M.to_serial({
    year = year,
    month = month,
    day = day,
    hour = hour,
    min = min,
    sec = sec,
  })
end

--- Create a time-only value (fractional day)
--- @param hour integer Hour (0-23)
--- @param min integer Minute (0-59)
--- @param sec? integer Second (0-59)
--- @return number Fractional day (0-1)
function M.time(hour, min, sec)
  sec = sec or 0
  return (hour * 3600 + min * 60 + sec) / SECONDS_PER_DAY
end

return M
