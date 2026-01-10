--- nvim-xlsx test commands plugin
--- This file registers test runner commands when loaded
---
--- Commands:
---   :XlsxTestAll      - Run all tests in a split terminal
---   :XlsxTest <name>  - Run a specific test (with tab completion)
---   :XlsxTestCurrent  - Run the test file currently being edited
---   :XlsxTestList     - List all available tests
---   :XlsxTestQuick    - Quick run all tests with summary
---
--- Usage:
---   After opening Neovim in the nvim-xlsx directory, run:
---     :XlsxTestAll
---   to run all tests, or:
---     :XlsxTest test_workbook
---   to run a specific test.

-- Only load if we're in the nvim-xlsx directory
local cwd = vim.fn.getcwd()
if not cwd:match("nvim%-xlsx") then
  return
end

-- Check if tests directory exists
local tests_dir = cwd .. "/tests"
if vim.fn.isdirectory(tests_dir) == 0 then
  return
end

-- Lazy load the runner when first command is used
local runner_loaded = false
local function ensure_runner()
  if not runner_loaded then
    -- Add the plugin directory to package path
    package.path = cwd .. "/?.lua;" .. cwd .. "/?/init.lua;" .. package.path

    local ok, runner = pcall(require, "tests.runner")
    if ok then
      runner.setup()
      runner_loaded = true
      return runner
    else
      vim.notify("Failed to load test runner: " .. tostring(runner), vim.log.levels.ERROR)
      return nil
    end
  end
  return require("tests.runner")
end

-- Register commands
vim.api.nvim_create_user_command("XlsxTestAll", function()
  local runner = ensure_runner()
  if runner then runner.run_all() end
end, { desc = "Run all nvim-xlsx tests" })

vim.api.nvim_create_user_command("XlsxTest", function(opts)
  local runner = ensure_runner()
  if runner then runner.run(opts.args) end
end, {
  nargs = 1,
  complete = function(arglead)
    local runner = ensure_runner()
    if runner then
      return runner.complete(arglead)
    end
    return {}
  end,
  desc = "Run specific nvim-xlsx test",
})

vim.api.nvim_create_user_command("XlsxTestCurrent", function()
  local runner = ensure_runner()
  if runner then runner.run_current() end
end, { desc = "Run current nvim-xlsx test file" })

vim.api.nvim_create_user_command("XlsxTestList", function()
  local runner = ensure_runner()
  if runner then runner.list() end
end, { desc = "List available nvim-xlsx tests" })

vim.api.nvim_create_user_command("XlsxTestQuick", function()
  local runner = ensure_runner()
  if runner then runner.run_quick() end
end, { desc = "Quick run all nvim-xlsx tests" })
