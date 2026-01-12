--- Test helper utilities for nvim-xlsx tests
--- @module tests.test_helper

local M = {}

-- Setup package path
local function setup_path()
  local script_path = debug.getinfo(2, "S").source:sub(2)
  script_path = script_path:gsub("\\", "/")
  local base_dir = script_path:match("(.+)/tests/")
  if not base_dir then
    base_dir = vim.fn.fnamemodify(script_path, ":p:h:h")
    base_dir = base_dir:gsub("\\", "/")
  end
  package.path = base_dir .. "/lua/?.lua;" .. base_dir .. "/lua/?/init.lua;" .. package.path
  return base_dir
end

M.base_dir = setup_path()
M.fixtures_dir = M.base_dir .. "/tests/fixtures"

-- Test state
M.all_passed = true
M.pass_count = 0
M.fail_count = 0

--- Run a single test assertion
--- @param name string Test name
--- @param condition boolean Test condition
function M.test(name, condition)
  if condition then
    print("  [PASS] " .. name)
    M.pass_count = M.pass_count + 1
  else
    print("  [FAIL] " .. name)
    M.fail_count = M.fail_count + 1
    M.all_passed = false
  end
end

--- Print a section header
--- @param name string Section name
function M.section(name)
  print("\n" .. name .. "...")
end

--- Print test suite header
--- @param name string Suite name
function M.suite(name)
  print("=== " .. name .. " ===")
  M.all_passed = true
  M.pass_count = 0
  M.fail_count = 0
end

--- Print test suite summary and exit
--- @param name string Suite name
function M.summary(name)
  print("\n=== " .. name .. " Complete ===")
  print(string.format("\nResults: %d passed, %d failed", M.pass_count, M.fail_count))
  if M.all_passed then
    print("All tests PASSED!")
  else
    print("Some tests FAILED!")
    os.exit(1)
  end
end

--- Get the xlsx module
--- @return table xlsx module
function M.require_xlsx()
  return require("nvim-xlsx")
end

return M
