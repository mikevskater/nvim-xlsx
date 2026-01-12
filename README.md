<div align="center">

# nvim-xlsx

**Pure Lua library for creating, reading, and modifying Excel .xlsx files in Neovim**

[![Neovim](https://img.shields.io/badge/Neovim-0.9+-57A143?style=for-the-badge&logo=neovim&logoColor=white)](https://neovim.io/)
[![Lua](https://img.shields.io/badge/Lua-2C2D72?style=for-the-badge&logo=lua&logoColor=white)](https://www.lua.org/)
[![License](https://img.shields.io/badge/License-MIT-blue.svg?style=for-the-badge)](LICENSE)

[Features](#features) •
[Installation](#installation) •
[Quick Start](#quick-start) •
[API Reference](#api-reference) •
[Contributing](#contributing)

</div>

---

## Features

<table>
<tr>
<td width="50%" valign="top">

### Core
- Create and save Excel `.xlsx` files
- Multiple worksheet support
- Read existing `.xlsx` files
- Import/export from Lua tables and CSV

### Cell Operations
- Text, numbers, and booleans
- Formulas with cached values
- Date and time values (Excel serial numbers)
- A1 notation support (`sheet:set("A1", value)`)

### Styling
- Font properties (name, size, bold, italic, underline, color)
- Background fills with colors and patterns
- Borders (all styles including dashed, double, etc.)
- Alignment (horizontal, vertical, wrap, rotation, indent)
- Number formats (built-in and custom)

</td>
<td width="50%" valign="top">

### Advanced Features
- Cell merging with overlap detection
- Freeze panes (rows, columns, or both)
- Auto-filter for data ranges
- Data validation (dropdowns, number ranges, custom rules)
- Hyperlinks (URLs, email, internal references)

### Print Settings
- Page orientation (portrait/landscape)
- Margins (top, bottom, left, right, header, footer)
- Print area definition
- Repeat title rows and columns

### Limitations
- No chart support
- No image insertion
- No pivot tables
- No password-protected files

</td>
</tr>
</table>

---

## Requirements

<table>
<tr>
<td>

**Neovim 0.9+** (0.11+ recommended)

</td>
<td>

**zip/unzip commands** (for .xlsx file handling)

</td>
</tr>
</table>

<details>
<summary><b>System Requirements by OS</b></summary>

<br>

| OS | Requirement | Notes |
|:---|:------------|:------|
| **Windows** | PowerShell 3.0+ | Built-in on Windows 7 SP1+ |
| **Linux** | `zip` and `unzip` | `apt install zip unzip` or `dnf install zip unzip` |
| **macOS** | Xcode CLI Tools | `xcode-select --install` |

</details>

---

## Installation

### lazy.nvim <sub>(Recommended)</sub>
#### IMPORT METHOD:
##### --lazy.lua
```lua
  require("lazy").setup({
      {import = 'config.plugins.xlsx'},
    }
  )
```
##### --xlsx.lua
```lua
  return {
    "mikevskater/nvim-xlsx",
    opts = {},
    lazy = true,
  }
```

---

#### DIRECT METHOD:
##### --lazy.lua
```lua
  require("lazy").setup({
      {"mikevskater/nvim-xlsx"},
    }
  )
```

<details>
<summary><b>Other Package Managers</b></summary>

<br>

**packer.nvim**
```lua
use "mikevskater/nvim-xlsx"
```

**vim-plug**
```vim
Plug 'mikevskater/nvim-xlsx'
```

**mini.deps**
```lua
add("mikevskater/nvim-xlsx")
```

</details>

---

## Quick Start

<table>
<tr>
<td width="33%">

**Create a Spreadsheet**

```lua
local xlsx = require("nvim-xlsx")

local wb = xlsx.new_workbook()
local sheet = wb:add_sheet("Data")

sheet:set_cell(1, 1, "Name")
sheet:set_cell(1, 2, "Age")
sheet:set_cell(2, 1, "Alice")
sheet:set_cell(2, 2, 30)

wb:save("output.xlsx")
```

</td>
<td width="33%">

**Export a Lua Table**

```lua
local xlsx = require("nvim-xlsx")

local data = {
  {"Product", "Price", "Qty"},
  {"Widget", 9.99, 100},
  {"Gadget", 24.99, 50},
}

xlsx.export_table(data, "products.xlsx")
```

</td>
<td width="33%">

**Read an Existing File**

```lua
local xlsx = require("nvim-xlsx")

local data = xlsx.import_table("input.xlsx")

for _, row in ipairs(data) do
  print(row[1], row[2], row[3])
end
```

</td>
</tr>
</table>

---

## API Reference

### Main Module

```lua
local xlsx = require("nvim-xlsx")
```

| Function | Description |
|:---------|:------------|
| `xlsx.new_workbook()` | Create a new workbook |
| `xlsx.open(filepath)` | Open existing XLSX file |
| `xlsx.export_table(data, filepath, opts?)` | Export 2D table to XLSX |
| `xlsx.import_table(filepath, opts?)` | Import XLSX to 2D table |
| `xlsx.info(filepath)` | Get file info without fully loading |
| `xlsx.from_csv(csv_string, filepath, opts?)` | Create XLSX from CSV string |
| `xlsx.from_csv_file(csv_path, xlsx_path, opts?)` | Create XLSX from CSV file |
| `xlsx.to_csv(filepath, opts?)` | Export XLSX to CSV string |

---

### Workbook

```lua
local wb = xlsx.new_workbook()  -- Create new
local wb = xlsx.open("file.xlsx")  -- Open existing
```

| Method | Description |
|:-------|:------------|
| `wb:add_sheet(name?)` | Add a worksheet (returns sheet, error) |
| `wb:get_sheet(name_or_index)` | Get sheet by name or 1-based index |
| `wb:set_active_sheet(name_or_index)` | Set the active sheet |
| `wb:create_style(def)` | Create a style, returns style index |
| `wb:set_properties(props)` | Set document properties |
| `wb:save(filepath)` | Save workbook to file |

---

### Worksheet

```lua
local sheet = wb:add_sheet("Sheet1")  -- Create new
local sheet = wb:get_sheet("Sheet1")  -- Get by name
local sheet = wb:get_sheet(1)  -- Get by index (1-based)
```

<details open>
<summary><b>Cell Operations</b></summary>

| Method | Description |
|:-------|:------------|
| `sheet:set_cell(row, col, value)` | Set cell value |
| `sheet:get_cell(row, col)` | Get cell (or nil) |
| `sheet:set(ref, value)` | Set using A1 notation |
| `sheet:get(ref)` | Get using A1 notation |
| `sheet:set_cell_value(row, col, value, style?)` | Set cell with optional style |
| `sheet:set_formula(row, col, formula, style?)` | Set formula |
| `sheet:set_date(row, col, date, style?)` | Set date value |
| `sheet:set_boolean(row, col, value, style?)` | Set boolean |

</details>

<details>
<summary><b>Dimension and Layout</b></summary>

| Method | Description |
|:-------|:------------|
| `sheet:get_dimension()` | Get dimension string (e.g., "A1:C10") |
| `sheet:set_column_width(col, width)` | Set column width in characters |
| `sheet:set_row_height(row, height)` | Set row height in points |

</details>

<details>
<summary><b>Styling</b></summary>

| Method | Description |
|:-------|:------------|
| `sheet:set_cell_style(row, col, style_index)` | Apply style to cell |
| `sheet:set_range_style(r1, c1, r2, c2, style_index)` | Apply style to range |
| `sheet:merge_cells(r1, c1, r2, c2)` | Merge cell range |
| `sheet:merge_range(range)` | Merge using A1:B2 notation |

</details>

<details>
<summary><b>Features</b></summary>

| Method | Description |
|:-------|:------------|
| `sheet:freeze_panes(rows, cols)` | Freeze rows and/or columns |
| `sheet:freeze_rows(n)` | Freeze first N rows |
| `sheet:freeze_cols(n)` | Freeze first N columns |
| `sheet:set_auto_filter(r1, c1, r2, c2)` | Set auto-filter range |
| `sheet:set_auto_filter_range(range)` | Set auto-filter using A1:B2 notation |
| `sheet:add_data_validation(ref, opts)` | Add data validation |
| `sheet:add_dropdown(ref, items, opts?)` | Add dropdown validation |
| `sheet:add_number_validation(ref, min, max, opts?)` | Add number validation |
| `sheet:add_hyperlink(row, col, target, opts?)` | Add hyperlink |

</details>

<details>
<summary><b>Print Settings</b></summary>

| Method | Description |
|:-------|:------------|
| `sheet:set_print_settings(settings)` | Set print settings |
| `sheet:set_orientation(orientation)` | "portrait" or "landscape" |
| `sheet:set_margins(top, bottom, left, right, header?, footer?)` | Set page margins |
| `sheet:set_print_area(range)` | Set print area |
| `sheet:set_print_title_rows(rows)` | Rows to repeat (e.g., "1:2") |
| `sheet:set_print_title_cols(cols)` | Columns to repeat (e.g., "A:B") |

</details>

---

### Styling

Create styles with `wb:create_style(definition)`:

```lua
local style = wb:create_style({
  -- Font
  font_name = "Arial",
  font_size = 12,
  bold = true,
  italic = false,
  underline = "single",
  strike = false,
  font_color = "#FFFFFF",

  -- Fill
  bg_color = "#4472C4",
  pattern = "solid",

  -- Border (all edges)
  border = true,
  border_style = "thin",
  border_color = "#000000",

  -- Or individual borders
  border_left = "thin",
  border_right = { style = "medium", color = "#FF0000" },
  border_top = "thin",
  border_bottom = "double",

  -- Alignment
  halign = "center",
  valign = "center",
  wrap_text = true,
  indent = 1,
  rotation = 45,

  -- Number format
  num_format = "0.00",
})

sheet:set_cell_style(1, 1, style)
```

<table>
<tr>
<td width="50%" valign="top">

**Color Formats**

Colors can be specified as:
- Hex with hash: `"#RRGGBB"` or `"#AARRGGBB"`
- Hex without hash: `"RRGGBB"` or `"AARRGGBB"`
- Named: `"red"`, `"blue"`, `"green"`, `"yellow"`, `"white"`, `"black"`

</td>
<td width="50%" valign="top">

**Border Styles**

`thin` `medium` `thick` `dashed` `dotted` `double` `hair` `mediumDashed` `dashDot` `mediumDashDot` `dashDotDot` `slantDashDot`

</td>
</tr>
</table>

<details>
<summary><b>Built-in Number Formats</b></summary>

<br>

| Name | ID | Format |
|:-----|:--:|:-------|
| `general` | 0 | General |
| `number` | 1 | 0 |
| `number_d2` | 2 | 0.00 |
| `number_thousands` | 3 | #,##0 |
| `percent` | 9 | 0% |
| `percent_d2` | 10 | 0.00% |
| `date` | 14 | m/d/yyyy |
| `time_24h` | 20 | h:mm |
| `datetime` | 22 | m/d/yyyy h:mm |
| `text` | 49 | @ |

**Custom formats:** `"$#,##0.00"` (currency), `"yyyy-mm-dd"` (ISO date), `"0.00%"` (percentage)

</details>

---

### Reading Files

```lua
-- Get file info
local info = xlsx.info("file.xlsx")
print(info.sheet_count)

-- Import all data
local data = xlsx.import_table("file.xlsx", { sheet_index = 1 })

-- Open for detailed access
local wb = xlsx.open("file.xlsx")
local sheets = xlsx.get_sheet_names(wb)
local sheet = xlsx.get_sheet(wb, "Sheet1")
local value = xlsx.get_cell(sheet, 1, 1)
local range = xlsx.get_range(sheet, 1, 1, 10, 5)
```

> **Note:** See `:help nvim-xlsx-reading` for complete reading API documentation.

---

### Utilities

<details>
<summary><b>Date Utilities</b></summary>

```lua
local xlsx = require("nvim-xlsx")

xlsx.date.to_serial({year=2024, month=1, day=15})  -- Lua date to Excel serial
xlsx.date.from_serial(45306)                        -- Excel serial to Lua date
xlsx.date.now()                                     -- Current date/time serial
xlsx.date.today()                                   -- Today's date serial
xlsx.date.date(2024, 1, 15)                         -- Create from components
xlsx.date.date(2024, 1, 15, 14, 30, 0)              -- With time
xlsx.date.time(14, 30, 0)                           -- Time-only value
xlsx.date.parse_iso("2024-01-15")                   -- Parse ISO 8601 string
xlsx.date.format_iso(45306)                         -- Format as ISO string
xlsx.date.from_unix_timestamp(1705305600)           -- Unix to Excel
xlsx.date.to_unix_timestamp(45306)                  -- Excel to Unix
```

</details>

<details>
<summary><b>Column Utilities</b></summary>

```lua
local xlsx = require("nvim-xlsx")

xlsx.utils.to_letter(1)               -- "A"
xlsx.utils.to_letter(27)              -- "AA"
xlsx.utils.to_number("A")             -- 1
xlsx.utils.to_number("AA")            -- 27
xlsx.utils.parse_ref("A1")            -- {row=1, col=1}
xlsx.utils.parse_ref("$A$1")          -- {row=1, col=1, abs_row=true, abs_col=true}
xlsx.utils.make_ref(1, 1)             -- "A1"
xlsx.utils.make_ref(1, 1, true, true) -- "$A$1"
xlsx.utils.parse_range("A1:B10")      -- Parse range reference
xlsx.utils.make_range(1, 1, 10, 2)    -- "A1:B10"
```

</details>

<details>
<summary><b>Color Utilities</b></summary>

```lua
local color = require("nvim-xlsx.utils.color")

color.to_argb("#FF0000")              -- "FFFF0000"
color.to_argb("red")                  -- "FFFF0000"
color.parse_argb("FFFF0000")          -- {a=255, r=255, g=0, b=0}
color.from_rgb(255, 0, 0)             -- "FFFF0000"
color.from_rgb(255, 0, 0, 128)        -- "80FF0000" (semi-transparent)
color.apply_tint("FFFF0000", 0.5)     -- lighter
color.apply_tint("FFFF0000", -0.5)    -- darker
```

</details>

<details>
<summary><b>Constants</b></summary>

```lua
local xlsx = require("nvim-xlsx")

xlsx.LIMITS.MAX_ROWS           -- 1,048,576
xlsx.LIMITS.MAX_COLS           -- 16,384
xlsx.LIMITS.MAX_CELL_TEXT      -- 32,767
xlsx.LIMITS.MAX_FORMULA_LENGTH -- 8,192
xlsx.LIMITS.MAX_SHEET_NAME     -- 31
xlsx.LIMITS.MAX_URL_LENGTH     -- 2,083

xlsx.BORDER_STYLES             -- Border style names
xlsx.HALIGN                    -- Horizontal alignment options
xlsx.VALIGN                    -- Vertical alignment options
xlsx.UNDERLINE                 -- Underline style options
xlsx.BUILTIN_FORMATS           -- Built-in number format IDs
```

</details>

---

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Run the tests (`nvim --headless -l tests/run_all.lua`)
5. Commit your changes (`git commit -m 'Add amazing feature'`)
6. Push to the branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

<details>
<summary><b>Running Tests</b></summary>

```bash
# Run all tests
nvim --headless -l tests/run_all.lua

# Run specific test
nvim --headless -l tests/test_cells.lua

# Run tests in Neovim (with plugin loaded)
:XlsxTestAll
```

</details>

---

## Issues

Found a bug or have a feature request? Please open an issue on [GitHub Issues](https://github.com/mikevskater/nvim-xlsx/issues).

<table>
<tr>
<td>

**When opening an issue, please include:**

- **Bug or Feature** — Label your issue type
- **Description** — Clear description of the issue or feature
- **Steps to Reproduce** — Minimal steps to reproduce (for bugs)
- **Expected Behavior** — What you expected to happen

</td>
</tr>
</table>

---

<div align="center">

## License

MIT License — see [LICENSE](LICENSE) for details.

<br>

Made with Lua for Neovim

</div>
