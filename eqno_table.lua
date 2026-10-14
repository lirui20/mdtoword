-- eqno_table.lua —— 把带 \tag{X} 的块级公式转成“公式居中 + 编号右对齐”的三列表格
-- 用法：pandoc input.md -o out.docx --lua-filter=eqno_table.lua
-- 说明：pandoc 的 texmath 不支持 \tag，这里先剥掉 \tag，再包进三列表格
--       （第1列空、第2列居中放公式、第3列右对齐放编号）。

local TEMPLATE = "|  |  |  |\n|:--|:-:|--:|\n|  |  |  |\n"  -- 三列：左/中/右

local function build_numbered(math, tag)
  math.text = math.text:gsub("\\tag%s*{[^}]*}", ""):gsub("^%s+", ""):gsub("%s+$", "")
  local tbl = pandoc.read(TEMPLATE, "markdown").blocks[1]
  local row = tbl.bodies[1].body[1]
  row.cells[1].contents = {}
  row.cells[2].contents = { pandoc.Plain({ math }) }
  row.cells[3].contents = { pandoc.Plain({ pandoc.Str("(" .. tag .. ")") }) }
  tbl.colspecs = {
    { pandoc.AlignLeft,   0.08 },
    { pandoc.AlignCenter, 0.84 },
    { pandoc.AlignRight,  0.08 },
  }
  return tbl
end

function Para(el)
  if #el.content == 1 and el.content[1].t == "Math"
     and el.content[1].mathtype == "DisplayMath" then
    local m = el.content[1]
    local tag = m.text:match("\\tag%s*{([^}]*)}")
    if tag then
      return build_numbered(m, tag)
    end
  end
  return nil
end
