# nvim-xlsx Tests

This directory contains the test suite for nvim-xlsx.

## Running Tests

### From Command Line

Run all tests:
```bash
nvim --headless -l tests/run_all.lua
```

Run a specific test:
```bash
nvim --headless -l tests/test_workbook.lua
```

### From Neovim

When editing files in the nvim-xlsx directory, you can use the following commands:

| Command | Description |
|---------|-------------|
| `:XlsxTestAll` | Run all tests in a split terminal |
| `:XlsxTest <name>` | Run a specific test (with tab completion) |
| `:XlsxTestCurrent` | Run the test file in the current buffer |
| `:XlsxTestList` | List all available tests |
| `:XlsxTestQuick` | Quick run with summary output |

Example:
```vim
:XlsxTest test_workbook
:XlsxTestAll
```

### Programmatic Usage

```lua
local runner = require("tests.runner")

-- Run all tests
runner.run_all()

-- Run specific test
runner.run("test_workbook")

-- Run current buffer's test
runner.run_current()

-- List available tests
runner.list()
```

## Test Files

| File | Description | Tests |
|------|-------------|-------|
| `test_workbook.lua` | Workbook creation, sheets, properties, saving | 47 |
| `test_cells.lua` | Cell operations, values, formulas, dates, booleans | 44 |
| `test_styles.lua` | Fonts, fills, borders, number formats, alignment | 43 |
| `test_merging.lua` | Merged cells | 24 |
| `test_reader.lua` | Reading xlsx files | 32 |
| `test_import_export.lua` | export_table, import_table, to_csv | 36 |
| `test_freeze_panes.lua` | Freeze rows/columns | 30 |
| `test_filters.lua` | Auto-filter | 22 |
| `test_validation.lua` | Data validation, dropdowns | 25 |
| `test_hyperlinks.lua` | Hyperlinks | 28 |
| `test_print.lua` | Print settings | 34 |
| `test_api.lua` | Public API verification | 113 |

**Total: 478 tests**

## Writing Tests

Tests use the shared `test_helper.lua` module:

```lua
dofile(arg[0]:match("(.*/)").. "test_helper.lua")
local h = require("tests.test_helper")
local xlsx = h.require_xlsx()

h.suite("My Tests")

h.section("Test 1: Basic test")
h.test("condition is true", some_value == expected_value)

h.summary("My Tests")
```

### Test Helper Functions

- `h.suite(name)` - Start a test suite
- `h.section(name)` - Start a test section
- `h.test(name, condition)` - Assert a condition
- `h.summary(name)` - Print summary and exit with appropriate code
- `h.require_xlsx()` - Load the xlsx module with correct paths
- `h.fixtures_dir` - Path to fixtures directory for test files

## Environment Variables

- `XLSX_TEST_VERBOSE=1` - Show all test output when running `run_all.lua`
- `XLSX_TEST_STOP=1` - Stop on first failure when running `run_all.lua`
