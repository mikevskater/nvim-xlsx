#!/usr/bin/env -S nvim --headless -l
--- Run all nvim-xlsx tests and provide a summary
--- Usage: nvim --headless -l tests/run_all.lua
---
--- Options (set as environment variables):
---   XLSX_TEST_VERBOSE=1    Show all test output
---   XLSX_TEST_STOP=1       Stop on first failure

-- Setup path
dofile(arg[0]:match("(.*/)").. "test_helper.lua")

local test_files = {
  "test_workbook",
  "test_cells",
  "test_styles",
  "test_merging",
  "test_reader",
  "test_import_export",
  "test_freeze_panes",
  "test_filters",
  "test_validation",
  "test_hyperlinks",
  "test_print",
  "test_api",
}

local verbose = os.getenv("XLSX_TEST_VERBOSE") == "1"
local stop_on_failure = os.getenv("XLSX_TEST_STOP") == "1"

local results = {}
local total_passed = 0
local total_failed = 0
local total_time = 0

print("╔══════════════════════════════════════════════════════════════╗")
print("║              nvim-xlsx Test Suite                            ║")
print("╠══════════════════════════════════════════════════════════════╣")
print(string.format("║  Running %d test files...                                    ║", #test_files))
print("╚══════════════════════════════════════════════════════════════╝")
print("")

local base_dir = arg[0]:match("(.*/)")

for _, test_name in ipairs(test_files) do
  local test_path = base_dir .. test_name .. ".lua"

  local start_time = os.clock()

  -- Capture output
  local output = {}
  local passed = 0
  local failed = 0

  -- Load and run the test file in a protected call
  local success, err = pcall(function()
    -- Override print to capture output
    local original_print = print
    print = function(...)
      local args = {...}
      local line = table.concat(vim.tbl_map(tostring, args), "\t")
      table.insert(output, line)

      -- Count results
      if line:match("%[PASS%]") then
        passed = passed + 1
      elseif line:match("%[FAIL%]") then
        failed = failed + 1
      end

      if verbose then
        original_print(...)
      end
    end

    -- Run the test
    dofile(test_path)

    -- Restore print
    print = original_print
  end)

  local elapsed = os.clock() - start_time

  local result = {
    name = test_name,
    success = success and failed == 0,
    passed = passed,
    failed = failed,
    time = elapsed,
    error = err,
    output = output,
  }

  table.insert(results, result)
  total_passed = total_passed + passed
  total_failed = total_failed + failed
  total_time = total_time + elapsed

  -- Print result line
  local status_icon = result.success and "✓" or "✗"
  local status_color = result.success and "" or ""
  print(string.format("  %s %-25s %3d passed, %2d failed  (%.2fs)",
    status_icon, test_name, passed, failed, elapsed))

  if not success then
    print(string.format("    ERROR: %s", tostring(err)))
  end

  if stop_on_failure and not result.success then
    print("\nStopping on first failure (XLSX_TEST_STOP=1)")
    break
  end
end

print("")
print("══════════════════════════════════════════════════════════════")
print(string.format("  Total: %d passed, %d failed in %.2fs", total_passed, total_failed, total_time))
print("══════════════════════════════════════════════════════════════")

-- Summary of failures
local failures = vim.tbl_filter(function(r) return not r.success end, results)
if #failures > 0 then
  print("")
  print("Failed tests:")
  for _, result in ipairs(failures) do
    print(string.format("  - %s (%d failures)", result.name, result.failed))
    if result.error then
      print(string.format("    Error: %s", result.error))
    end
    -- Show failed test lines
    for _, line in ipairs(result.output) do
      if line:match("%[FAIL%]") then
        print("    " .. line)
      end
    end
  end
  print("")
end

-- Exit with appropriate code
if total_failed > 0 then
  print("FAILED")
  vim.cmd("cquit 1")
else
  print("ALL TESTS PASSED")
  vim.cmd("quit")
end
