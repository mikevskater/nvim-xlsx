--- Worksheet XML generation module
--- @module nvim-xlsx.worksheet.xml

local column_utils = require("nvim-xlsx.utils.column")
local xml = require("nvim-xlsx.xml.writer")
local templates = require("nvim-xlsx.xml.templates")

local M = {}

--- Generate the sheetData XML content
--- @param self Worksheet
--- @return string
function M._generate_sheet_data(self)
  if not self.min_row then
    return ""
  end

  local parts = {}

  -- Get sorted row numbers
  local row_nums = {}
  for row_num in pairs(self.rows) do
    table.insert(row_nums, row_num)
  end
  table.sort(row_nums)

  -- Generate each row
  for _, row_num in ipairs(row_nums) do
    local row_data = self.rows[row_num]
    local cell_parts = {}

    -- Get sorted column numbers for this row
    local col_nums = {}
    for col_num in pairs(row_data) do
      table.insert(col_nums, col_num)
    end
    table.sort(col_nums)

    -- Generate cells
    for _, col_num in ipairs(col_nums) do
      local cell = row_data[col_num]
      if cell and cell:has_content() then
        table.insert(cell_parts, cell:to_xml())
      end
    end

    if #cell_parts > 0 then
      local row_attrs = { r = row_num }
      -- Add custom height if set
      if self.row_heights[row_num] then
        row_attrs.ht = self.row_heights[row_num]
        row_attrs.customHeight = "1"
      end
      table.insert(parts, xml.element_raw("row", table.concat(cell_parts), row_attrs))
    end
  end

  return table.concat(parts)
end

--- Generate column definitions XML
--- @param self Worksheet
--- @return string
function M._generate_cols(self)
  if not next(self.column_widths) then
    return ""
  end

  local parts = {}
  local cols = {}
  for col in pairs(self.column_widths) do
    table.insert(cols, col)
  end
  table.sort(cols)

  for _, col in ipairs(cols) do
    local width = self.column_widths[col]
    table.insert(parts, xml.empty_element("col", {
      min = col,
      max = col,
      width = width,
      customWidth = "1",
    }))
  end

  return xml.element_raw("cols", table.concat(parts))
end

--- Generate the sheetViews XML with freeze pane support
--- @param self Worksheet
--- @param is_active boolean Whether this sheet is the active/selected sheet
--- @return string
function M._generate_sheet_views(self, is_active)
  local tab_selected = is_active and "1" or "0"

  if not self.freeze_pane then
    -- Simple sheet view without freeze panes
    return '<sheetViews><sheetView tabSelected="' .. tab_selected .. '" workbookViewId="0"/></sheetViews>'
  end

  local fp = self.freeze_pane
  local parts = {}

  table.insert(parts, '<sheetViews>')
  table.insert(parts, '<sheetView tabSelected="' .. tab_selected .. '" workbookViewId="0">')

  -- Calculate the top-left cell of the unfrozen region
  local top_left_row = fp.rows + 1
  local top_left_col = fp.cols + 1
  local top_left_cell = column_utils.make_ref(top_left_row, top_left_col)

  -- Determine the active pane and pane state
  local active_pane
  local pane_state = "frozen"

  if fp.rows > 0 and fp.cols > 0 then
    -- Both rows and columns frozen
    active_pane = "bottomRight"
  elseif fp.rows > 0 then
    -- Only rows frozen
    active_pane = "bottomLeft"
  else
    -- Only columns frozen
    active_pane = "topRight"
  end

  -- Generate pane element
  local pane_attrs = {
    state = pane_state,
    topLeftCell = top_left_cell,
    activePane = active_pane,
  }

  if fp.cols > 0 then
    pane_attrs.xSplit = fp.cols
  end
  if fp.rows > 0 then
    pane_attrs.ySplit = fp.rows
  end

  table.insert(parts, xml.empty_element("pane", pane_attrs))

  -- Generate selection elements for the panes
  if fp.rows > 0 and fp.cols > 0 then
    -- Four panes: need selections for topRight, bottomLeft, bottomRight
    table.insert(parts, xml.empty_element("selection", { pane = "topRight", activeCell = column_utils.make_ref(1, top_left_col), sqref = column_utils.make_ref(1, top_left_col) }))
    table.insert(parts, xml.empty_element("selection", { pane = "bottomLeft", activeCell = column_utils.make_ref(top_left_row, 1), sqref = column_utils.make_ref(top_left_row, 1) }))
    table.insert(parts, xml.empty_element("selection", { pane = "bottomRight", activeCell = top_left_cell, sqref = top_left_cell }))
  elseif fp.rows > 0 then
    -- Two panes (top/bottom)
    table.insert(parts, xml.empty_element("selection", { pane = "bottomLeft", activeCell = top_left_cell, sqref = top_left_cell }))
  else
    -- Two panes (left/right)
    table.insert(parts, xml.empty_element("selection", { pane = "topRight", activeCell = top_left_cell, sqref = top_left_cell }))
  end

  table.insert(parts, '</sheetView>')
  table.insert(parts, '</sheetViews>')

  return table.concat(parts)
end

--- Generate auto-filter XML
--- @param self Worksheet
--- @return string
function M._generate_auto_filter(self)
  if not self.auto_filter then
    return ""
  end
  return xml.empty_element("autoFilter", { ref = self.auto_filter.ref })
end

--- Generate data validations XML
--- @param self Worksheet
--- @return string
function M._generate_data_validations(self)
  if #self.data_validations == 0 then
    return ""
  end

  local parts = {}

  for _, dv in ipairs(self.data_validations) do
    local attrs = {
      type = dv.type,
      sqref = dv.ref,
      allowBlank = dv.allowBlank and "1" or "0",
      showErrorMessage = dv.showErrorMessage and "1" or "0",
      showInputMessage = dv.showInputMessage and "1" or "0",
    }

    -- Note: Excel uses showDropDown="1" to HIDE the dropdown (counter-intuitive)
    -- We expose showDropDown=true to SHOW the dropdown, so we invert it
    if dv.type == "list" and dv.showDropDown == false then
      attrs.showDropDown = "1"  -- Hide dropdown
    end

    if dv.operator then
      attrs.operator = dv.operator
    end
    if dv.errorStyle and dv.errorStyle ~= "stop" then
      attrs.errorStyle = dv.errorStyle
    end
    if dv.errorTitle then
      attrs.errorTitle = dv.errorTitle
    end
    if dv.error then
      attrs.error = dv.error
    end
    if dv.promptTitle then
      attrs.promptTitle = dv.promptTitle
    end
    if dv.prompt then
      attrs.prompt = dv.prompt
    end

    -- Build inner content (formulas)
    local inner = {}
    if dv.formula1 then
      table.insert(inner, xml.element("formula1", dv.formula1))
    end
    if dv.formula2 then
      table.insert(inner, xml.element("formula2", dv.formula2))
    end

    if #inner > 0 then
      table.insert(parts, xml.element_raw("dataValidation", table.concat(inner), attrs))
    else
      table.insert(parts, xml.empty_element("dataValidation", attrs))
    end
  end

  return xml.element_raw("dataValidations", table.concat(parts), { count = #self.data_validations })
end

--- Generate hyperlinks XML
--- @param self Worksheet
--- @param hyperlink_rels table Table to populate with relationship info for external links
--- @return string
function M._generate_hyperlinks(self, hyperlink_rels)
  if #self.hyperlinks == 0 then
    return ""
  end

  local parts = {}
  local rel_id = 1

  for _, link in ipairs(self.hyperlinks) do
    local attrs = { ref = link.ref }

    if link.is_external then
      -- External link needs a relationship
      local rid = "rId" .. (1000 + rel_id)  -- Use high IDs to avoid conflicts
      attrs["r:id"] = rid
      if link.tooltip then
        attrs.tooltip = link.tooltip
      end
      if link.display then
        attrs.display = link.display
      end
      table.insert(hyperlink_rels, {
        id = rid,
        target = link.target,
        type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        targetMode = "External"
      })
      rel_id = rel_id + 1
    else
      -- Internal link uses location attribute
      -- Excel internal links need the location without a leading # (Excel adds it)
      attrs.location = link.location
      if link.tooltip then
        attrs.tooltip = link.tooltip
      end
      if link.display then
        attrs.display = link.display
      end
    end

    table.insert(parts, xml.empty_element("hyperlink", attrs))
  end

  return xml.element_raw("hyperlinks", table.concat(parts))
end

--- Generate print settings XML (pageMargins and pageSetup)
--- @param self Worksheet
--- @return string
function M._generate_print_settings(self)
  if not self.print_settings then
    return ""
  end

  local parts = {}
  local ps = self.print_settings

  -- Page margins
  if ps.margins then
    local m = ps.margins
    table.insert(parts, xml.empty_element("pageMargins", {
      left = m.left or 0.7,
      right = m.right or 0.7,
      top = m.top or 0.75,
      bottom = m.bottom or 0.75,
      header = m.header or 0.3,
      footer = m.footer or 0.3,
    }))
  end

  -- Page setup
  local setup_attrs = {}
  local has_setup = false

  if ps.orientation then
    setup_attrs.orientation = ps.orientation
    has_setup = true
  end
  if ps.paperSize then
    setup_attrs.paperSize = ps.paperSize
    has_setup = true
  end
  if ps.scale then
    setup_attrs.scale = ps.scale
    has_setup = true
  end
  if ps.fitToWidth then
    setup_attrs.fitToWidth = ps.fitToWidth
    has_setup = true
  end
  if ps.fitToHeight then
    setup_attrs.fitToHeight = ps.fitToHeight
    has_setup = true
  end
  if ps.gridLines then
    setup_attrs.gridLines = "1"
    has_setup = true
  end
  if ps.headings then
    setup_attrs.headings = "1"
    has_setup = true
  end

  if has_setup then
    table.insert(parts, xml.empty_element("pageSetup", setup_attrs))
  end

  return table.concat(parts)
end

--- Get hyperlink relationships for this worksheet
--- @param self Worksheet
--- @return table[] Array of relationship info for external hyperlinks
function M.get_hyperlink_relationships(self)
  local rels = {}
  M._generate_hyperlinks(self, rels)
  return rels
end

--- Generate the complete worksheet XML
--- @param self Worksheet
--- @param is_active? boolean Whether this sheet is the active/selected sheet
--- @return string
function M.to_xml(self, is_active)
  local b = xml.builder()

  b:declaration()
  b:open("worksheet", {
    xmlns = templates.NS.SPREADSHEET,
    ["xmlns:r"] = templates.NS.RELATIONSHIPS,
  })

  -- Dimension
  b:empty("dimension", { ref = self:get_dimension() })

  -- Sheet views (with freeze pane support)
  b:raw(M._generate_sheet_views(self, is_active or false))

  -- Sheet format defaults
  b:empty("sheetFormatPr", { defaultRowHeight = "15" })

  -- Column widths
  local cols = M._generate_cols(self)
  if cols ~= "" then
    b:raw(cols)
  end

  -- Sheet data
  local sheet_data = M._generate_sheet_data(self)
  b:elem_raw("sheetData", sheet_data)

  -- Auto-filter (must come after sheetData)
  local auto_filter = M._generate_auto_filter(self)
  if auto_filter ~= "" then
    b:raw(auto_filter)
  end

  -- Merged cells (if any)
  if #self.merged_cells > 0 then
    local merge_parts = {}
    for _, merge in ipairs(self.merged_cells) do
      table.insert(merge_parts, xml.empty_element("mergeCell", { ref = merge }))
    end
    b:elem_raw("mergeCells", table.concat(merge_parts), { count = #self.merged_cells })
  end

  -- Data validations
  local data_validations = M._generate_data_validations(self)
  if data_validations ~= "" then
    b:raw(data_validations)
  end

  -- Hyperlinks
  local hyperlink_rels = {}
  local hyperlinks = M._generate_hyperlinks(self, hyperlink_rels)
  if hyperlinks ~= "" then
    b:raw(hyperlinks)
  end

  -- Print settings (pageMargins and pageSetup)
  local print_settings = M._generate_print_settings(self)
  if print_settings ~= "" then
    b:raw(print_settings)
  end

  b:close("worksheet")

  return b:to_string()
end

return M
