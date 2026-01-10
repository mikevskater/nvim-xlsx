--- Test runner for nvim-xlsx
--- Provides Neovim commands and functions for running tests
---
--- Usage from Neovim:
---   :lua require("tests.runner").run_all()
---   :lua require("tests.runner").run("test_workbook")
---   :lua require("tests.runner").run_current()
---
--- Or use the provided commands after loading:
---   :XlsxTestAll          - Run all tests
---   :XlsxTest <name>      - Run specific test (tab completion available)
---   :XlsxTestCurrent      - Run test file in current buffer
---   :XlsxTestList         - List available tests

local M = {}

-- Get the tests directory path
local function get_tests_dir()
  local info = debug.getinfo(1, "S")
  local script_path = info.source:sub(2) -- Remove the @ prefix
  -- Normalize path separators
  script_path = script_path:gsub("\\", "/")
  return script_path:match("(.*/)")
end

M.tests_dir = get_tests_dir()

-- List of all test files (without .lua extension)
M.test_files = {
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

--- Run a single test file
--- @param name string Test file name (with or without .lua extension)
--- @param opts? table Options: { silent?: boolean, callback?: function }
function M.run(name, opts)
  opts = opts or {}

  -- Remove .lua extension if present
  name = name:gsub("%.lua$", "")

  -- Validate test exists
  local found = false
  for _, test in ipairs(M.test_files) do
    if test == name then
      found = true
      break
    end
  end

  if not found then
    vim.notify("Test not found: " .. name, vim.log.levels.ERROR)
    vim.notify("Available tests: " .. table.concat(M.test_files, ", "), vim.log.levels.INFO)
    return false
  end

  local test_path = M.tests_dir .. name .. ".lua"

  if not opts.silent then
    vim.notify("Running test: " .. name, vim.log.levels.INFO)
  end

  -- Run test in a new terminal buffer
  local cmd = string.format('nvim --headless -l "%s"', test_path)

  -- Create a new split for output
  vim.cmd("botright new")
  vim.cmd("resize 15")

  local buf = vim.api.nvim_get_current_buf()
  vim.bo[buf].buftype = "nofile"
  vim.bo[buf].bufhidden = "wipe"
  vim.bo[buf].swapfile = false
  vim.api.nvim_buf_set_name(buf, "Test: " .. name)

  -- Run the test and capture output
  vim.fn.termopen(cmd, {
    on_exit = function(_, exit_code)
      if opts.callback then
        opts.callback(exit_code == 0, name)
      end
      if exit_code == 0 then
        vim.notify("Test passed: " .. name, vim.log.levels.INFO)
      else
        vim.notify("Test failed: " .. name, vim.log.levels.ERROR)
      end
    end,
  })

  -- Enter insert mode to see output scrolling
  vim.cmd("startinsert")

  return true
end

--- Run all tests
--- @param opts? table Options: { silent?: boolean, stop_on_failure?: boolean }
function M.run_all(opts)
  opts = opts or {}

  vim.notify("Running all " .. #M.test_files .. " test files...", vim.log.levels.INFO)

  -- Build command to run all tests sequentially
  local commands = {}
  for _, test in ipairs(M.test_files) do
    local test_path = M.tests_dir .. test .. ".lua"
    table.insert(commands, string.format('echo "\\n=== Running %s ===" && nvim --headless -l "%s"', test, test_path))
  end

  local cmd = table.concat(commands, " && ")

  -- Create output buffer
  vim.cmd("botright new")
  vim.cmd("resize 20")

  local buf = vim.api.nvim_get_current_buf()
  vim.bo[buf].buftype = "nofile"
  vim.bo[buf].bufhidden = "wipe"
  vim.bo[buf].swapfile = false
  vim.api.nvim_buf_set_name(buf, "All Tests")

  vim.fn.termopen(cmd, {
    on_exit = function(_, exit_code)
      if exit_code == 0 then
        vim.notify("All tests passed!", vim.log.levels.INFO)
      else
        vim.notify("Some tests failed!", vim.log.levels.ERROR)
      end
    end,
  })

  vim.cmd("startinsert")
end

--- Run the test file currently open in the buffer
function M.run_current()
  local bufname = vim.api.nvim_buf_get_name(0)
  local filename = vim.fn.fnamemodify(bufname, ":t:r") -- Get filename without extension

  -- Check if it's a test file
  if not filename:match("^test_") then
    vim.notify("Current file is not a test file (should start with 'test_')", vim.log.levels.WARN)
    return false
  end

  return M.run(filename)
end

--- List available tests
function M.list()
  vim.notify("Available test files:", vim.log.levels.INFO)
  for i, test in ipairs(M.test_files) do
    print(string.format("  %2d. %s", i, test))
  end
end

--- Get test completion list
function M.complete(arglead, cmdline, cursorpos)
  local matches = {}
  for _, test in ipairs(M.test_files) do
    if test:find(arglead, 1, true) == 1 then
      table.insert(matches, test)
    end
  end
  return matches
end

--- Quick summary run (minimal output)
--- @param opts? table Options
function M.run_quick(opts)
  opts = opts or {}

  local test_path = M.tests_dir .. "run_all.lua"
  local cmd = string.format('nvim --headless -l "%s"', test_path)

  vim.cmd("botright new")
  vim.cmd("resize 20")

  local buf = vim.api.nvim_get_current_buf()
  vim.bo[buf].buftype = "nofile"
  vim.bo[buf].bufhidden = "wipe"
  vim.bo[buf].swapfile = false
  vim.api.nvim_buf_set_name(buf, "Quick Test Run")

  vim.fn.termopen(cmd, {
    on_exit = function(_, exit_code)
      if exit_code == 0 then
        vim.notify("All tests passed!", vim.log.levels.INFO)
      else
        vim.notify("Some tests failed!", vim.log.levels.ERROR)
      end
    end,
  })

  vim.cmd("startinsert")
end

--- Setup commands (call this to register vim commands)
function M.setup()
  -- :XlsxTestAll - Run all tests
  vim.api.nvim_create_user_command("XlsxTestAll", function()
    M.run_all()
  end, { desc = "Run all nvim-xlsx tests" })

  -- :XlsxTest <name> - Run specific test with completion
  vim.api.nvim_create_user_command("XlsxTest", function(opts)
    M.run(opts.args)
  end, {
    nargs = 1,
    complete = function(arglead)
      return M.complete(arglead)
    end,
    desc = "Run specific nvim-xlsx test",
  })

  -- :XlsxTestCurrent - Run current test file
  vim.api.nvim_create_user_command("XlsxTestCurrent", function()
    M.run_current()
  end, { desc = "Run current nvim-xlsx test file" })

  -- :XlsxTestList - List available tests
  vim.api.nvim_create_user_command("XlsxTestList", function()
    M.list()
  end, { desc = "List available nvim-xlsx tests" })

  -- :XlsxTestQuick - Quick run all tests
  vim.api.nvim_create_user_command("XlsxTestQuick", function()
    M.run_quick()
  end, { desc = "Quick run all nvim-xlsx tests" })

  vim.notify("nvim-xlsx test commands registered", vim.log.levels.INFO)
end

return M
