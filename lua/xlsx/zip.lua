--- ZIP operations for xlsx file handling
--- Uses external zip commands via vim.uv.spawn
--- @module xlsx.zip

local M = {}

local uv = vim.uv or vim.loop

--- Check if running on Windows
--- @return boolean
local function is_windows()
  return vim.fn.has("win32") == 1 or vim.fn.has("win64") == 1
end

--- Generate a unique temporary directory path
--- @return string
local function make_temp_dir()
  local temp_base = vim.fn.tempname()
  local temp_dir = temp_base .. "_xlsx_" .. os.time()
  vim.fn.mkdir(temp_dir, "p")
  return temp_dir
end

--- Remove a directory recursively
--- @param path string Directory path
local function remove_dir(path)
  -- Use vim.fn.delete with "rf" flag for recursive force delete
  vim.fn.delete(path, "rf")
end

--- Synchronously spawn a process and wait for completion
--- @param cmd string Command to run
--- @param args string[] Arguments
--- @param cwd? string Working directory
--- @return boolean success
--- @return string output Combined stdout/stderr
local function spawn_sync(cmd, args, cwd)
  local stdout_data = {}
  local stderr_data = {}
  local done = false
  local exit_code = -1

  local stdout = uv.new_pipe(false)
  local stderr = uv.new_pipe(false)

  local handle
  handle = uv.spawn(cmd, {
    args = args,
    cwd = cwd,
    stdio = { nil, stdout, stderr },
  }, function(code)
    exit_code = code
    done = true
    if handle then
      handle:close()
    end
  end)

  if not handle then
    if stdout then stdout:close() end
    if stderr then stderr:close() end
    return false, "Failed to spawn: " .. cmd
  end

  stdout:read_start(function(err, data)
    if data then
      table.insert(stdout_data, data)
    end
    if err then
      stdout:close()
    end
  end)

  stderr:read_start(function(err, data)
    if data then
      table.insert(stderr_data, data)
    end
    if err then
      stderr:close()
    end
  end)

  -- Wait for process to complete
  local timeout = 30000 -- 30 seconds
  local start = uv.now()
  while not done do
    uv.run("once")
    if uv.now() - start > timeout then
      if handle then handle:close() end
      if stdout then stdout:close() end
      if stderr then stderr:close() end
      return false, "Process timed out"
    end
  end

  if stdout then stdout:close() end
  if stderr then stderr:close() end

  local output = table.concat(stdout_data) .. table.concat(stderr_data)
  return exit_code == 0, output
end

--- Create a ZIP archive from a directory (Windows implementation)
--- @param source_dir string Source directory containing files
--- @param output_path string Output ZIP file path
--- @return boolean success
--- @return string? error_message
local function zip_windows(source_dir, output_path)
  -- Ensure output path is absolute
  output_path = vim.fn.fnamemodify(output_path, ":p")
  source_dir = vim.fn.fnamemodify(source_dir, ":p")

  -- Remove trailing slash for PowerShell compatibility
  if source_dir:sub(-1) == "\\" or source_dir:sub(-1) == "/" then
    source_dir = source_dir:sub(1, -2)
  end

  -- Delete existing file if present
  if vim.fn.filereadable(output_path) == 1 then
    vim.fn.delete(output_path)
  end

  -- PowerShell's Compress-Archive only accepts .zip extension
  -- Create as .zip then rename if needed
  local temp_zip = output_path
  local needs_rename = false
  if not output_path:lower():match("%.zip$") then
    temp_zip = output_path .. ".zip"
    needs_rename = true
    -- Delete temp zip if it exists
    if vim.fn.filereadable(temp_zip) == 1 then
      vim.fn.delete(temp_zip)
    end
  end

  -- PowerShell command to create ZIP
  -- We need to add all contents from source_dir into the zip
  local ps_script = string.format(
    [[$ErrorActionPreference = 'Stop'; ]] ..
    [[Compress-Archive -Path '%s\*' -DestinationPath '%s' -CompressionLevel Optimal]],
    source_dir:gsub("'", "''"),
    temp_zip:gsub("'", "''")
  )

  local success, output = spawn_sync("powershell", {
    "-NoProfile",
    "-NonInteractive",
    "-Command",
    ps_script,
  })

  if not success then
    return false, "Failed to create ZIP: " .. output
  end

  -- Rename to final output path if needed
  if needs_rename then
    local rename_ok = vim.fn.rename(temp_zip, output_path)
    if rename_ok ~= 0 then
      vim.fn.delete(temp_zip)
      return false, "Failed to rename zip to " .. output_path
    end
  end

  return true
end

--- Create a ZIP archive from a directory (Unix implementation)
--- @param source_dir string Source directory containing files
--- @param output_path string Output ZIP file path
--- @return boolean success
--- @return string? error_message
local function zip_unix(source_dir, output_path)
  -- Ensure output path is absolute
  output_path = vim.fn.fnamemodify(output_path, ":p")

  -- Delete existing file if present
  if vim.fn.filereadable(output_path) == 1 then
    vim.fn.delete(output_path)
  end

  -- Use zip command from within source directory
  local success, output = spawn_sync("zip", {
    "-r",
    "-q",
    output_path,
    ".",
  }, source_dir)

  if not success then
    return false, "Failed to create ZIP: " .. output
  end

  return true
end

--- Extract a ZIP archive to a directory (Windows implementation)
--- @param zip_path string ZIP file path
--- @param dest_dir string Destination directory
--- @return boolean success
--- @return string? error_message
local function unzip_windows(zip_path, dest_dir)
  zip_path = vim.fn.fnamemodify(zip_path, ":p")
  dest_dir = vim.fn.fnamemodify(dest_dir, ":p")

  -- Create destination directory if it doesn't exist
  vim.fn.mkdir(dest_dir, "p")

  local ps_script = string.format(
    [[$ErrorActionPreference = 'Stop'; ]] ..
    [[Expand-Archive -Path '%s' -DestinationPath '%s' -Force]],
    zip_path:gsub("'", "''"),
    dest_dir:gsub("'", "''")
  )

  local success, output = spawn_sync("powershell", {
    "-NoProfile",
    "-NonInteractive",
    "-Command",
    ps_script,
  })

  if not success then
    return false, "Failed to extract ZIP: " .. output
  end

  return true
end

--- Extract a ZIP archive to a directory (Unix implementation)
--- @param zip_path string ZIP file path
--- @param dest_dir string Destination directory
--- @return boolean success
--- @return string? error_message
local function unzip_unix(zip_path, dest_dir)
  zip_path = vim.fn.fnamemodify(zip_path, ":p")
  dest_dir = vim.fn.fnamemodify(dest_dir, ":p")

  -- Create destination directory if it doesn't exist
  vim.fn.mkdir(dest_dir, "p")

  local success, output = spawn_sync("unzip", {
    "-q",
    "-o",
    zip_path,
    "-d",
    dest_dir,
  })

  if not success then
    return false, "Failed to extract ZIP: " .. output
  end

  return true
end

--- Create a ZIP archive from a directory
--- @param source_dir string Source directory containing files to zip
--- @param output_path string Output ZIP file path
--- @return boolean success
--- @return string? error_message
function M.zip_directory(source_dir, output_path)
  if is_windows() then
    return zip_windows(source_dir, output_path)
  else
    return zip_unix(source_dir, output_path)
  end
end

--- Extract a ZIP archive to a directory
--- @param zip_path string ZIP file path
--- @param dest_dir string Destination directory
--- @return boolean success
--- @return string? error_message
function M.unzip_file(zip_path, dest_dir)
  if is_windows() then
    return unzip_windows(zip_path, dest_dir)
  else
    return unzip_unix(zip_path, dest_dir)
  end
end

--- Create a temporary directory for xlsx operations
--- @return string temp_dir_path
function M.create_temp_dir()
  return make_temp_dir()
end

--- Remove a temporary directory and all its contents
--- @param path string Directory path to remove
function M.cleanup_temp_dir(path)
  remove_dir(path)
end

--- Write a file with directory creation
--- @param filepath string Full file path
--- @param content string File content
--- @return boolean success
--- @return string? error_message
function M.write_file(filepath, content)
  -- Ensure parent directory exists
  local parent = vim.fn.fnamemodify(filepath, ":h")
  if vim.fn.isdirectory(parent) == 0 then
    vim.fn.mkdir(parent, "p")
  end

  local file, err = io.open(filepath, "wb")
  if not file then
    return false, "Failed to open file for writing: " .. (err or filepath)
  end

  local ok, write_err = file:write(content)
  file:close()

  if not ok then
    return false, "Failed to write file: " .. (write_err or filepath)
  end

  return true
end

--- Read a file's contents
--- @param filepath string Full file path
--- @return string? content File content or nil on error
--- @return string? error_message
function M.read_file(filepath)
  local file, err = io.open(filepath, "rb")
  if not file then
    return nil, "Failed to open file for reading: " .. (err or filepath)
  end

  local content = file:read("*all")
  file:close()

  return content
end

return M
