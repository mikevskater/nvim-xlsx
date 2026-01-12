--- XML processing modules for xlsx
--- @module nvim-xlsx.xml

local M = {}

M.writer = require("nvim-xlsx.xml.writer")
M.parser = require("nvim-xlsx.xml.parser")

return M
