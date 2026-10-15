#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
md2word.py —— 图形化将 Markdown 批量转换为 Word (.docx) 文档

依赖：
    - Python 3.x
    - wxPython  (pip install wxPython)
    - pandoc    (源码运行时需要；打包版已内置)

运行：
    python md2word.py

说明：
    转换后端使用 pandoc 命令行，因此在调用 pandoc 时本程序不占用界面线程，
    批量转换过程中窗口不会卡死。所有文件默认输出为同名 .docx。

    公式：pandoc 会把 LaTeX 数学（$...$、$$...$$、\(...\)、\[...\]）转换成
    Word 原生公式（OMML，可在 Word 里直接编辑）。
    注意：pandoc 的 texmath 不支持 LaTeX 的 \tag{}（会删掉编号，甚至导致整条
    公式转换失败）。因此本程序内置一个 Lua 过滤器，把带 \tag{X} 的块级公式转成
    无边框三列表格（公式居中、编号 (X) 右对齐），版式接近正式报告。

    复制：左侧粘贴框“转换并复制”与右侧“复制内容”会把结果写入剪贴板，粘贴到
    Word/WPS 时为原生可编辑公式，而不再是纯文本。做法是把结果作为带 MathML 的
    HTML 写入剪贴板，由 Word/WPS 在粘贴时把 MathML 转成原生公式，无需安装 Word。
"""

import html
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import zipfile

import wx
import wx.dataview

# --------------------------------------------------------------------------- #
# 常量
# --------------------------------------------------------------------------- #
APP_TITLE = "Markdown → Word 批量转换器"
MD_EXTS = (".md", ".markdown", ".mdown", ".mkd", ".mkdn", ".mdwn", ".mdtxt", ".mdtext")


# --------------------------------------------------------------------------- #
# 视觉主题：冷灰白 + 墨蓝点缀
#
# 设计原则：低饱和中性底色、克制的细边框、统一的字阶与留白。
# 不使用渐变、霓虹或高饱和强调色，保持安静、耐看的工具气质。
# --------------------------------------------------------------------------- #
class T:
    """集中管理配色、圆角与内边距，便于整体微调。"""

    BG          = wx.Colour(0xEC, 0xEE, 0xF1)   # 窗口底色（冷灰）
    SURFACE     = wx.Colour(0xFF, 0xFF, 0xFF)   # 卡片
    SURFACE_ALT = wx.Colour(0xF5, 0xF6, 0xF8)   # 次级填充
    TRACK       = wx.Colour(0xE3, 0xE7, 0xEC)   # 进度轨道
    BORDER      = wx.Colour(0xDB, 0xDF, 0xE4)   # 细边框
    BORDER_HI   = wx.Colour(0xC2, 0xC9, 0xD1)   # 悬停边框
    TEXT        = wx.Colour(0x1F, 0x26, 0x2E)   # 主文字
    TEXT_MUTED  = wx.Colour(0x6B, 0x77, 0x85)   # 次要文字
    ACCENT      = wx.Colour(0x2E, 0x4A, 0x62)   # 墨蓝
    ACCENT_HI   = wx.Colour(0x25, 0x3C, 0x52)   # 墨蓝·悬停
    ACCENT_LO   = wx.Colour(0x1B, 0x2F, 0x44)   # 墨蓝·按下
    ACCENT_SOFT = wx.Colour(0xE8, 0xED, 0xF2)   # 墨蓝·浅底
    DISABLED    = wx.Colour(0xB3, 0xBC, 0xC5)   # 禁用
    FOCUS       = wx.Colour(0x54, 0x74, 0x93)   # 焦点描边

    OK   = wx.Colour(0x2F, 0x6B, 0x4F)          # 成功（墨绿）
    ERR  = wx.Colour(0xA3, 0x3A, 0x3A)          # 失败（暗红）
    SKIP = wx.Colour(0x8A, 0x6A, 0x2F)          # 跳过（赭石）

    ROW_ALT     = wx.Colour(0xF7, 0xF8, 0xFA)   # 列表交替行
    RADIUS_CARD = 6
    RADIUS_CTRL = 5
    PAD         = 14                            # 卡片内边距


COL_STATUS_OK = T.OK
COL_STATUS_ERR = T.ERR
COL_STATUS_SKIP = T.SKIP

_FACE_UI = None
_FACE_MONO = None


def _pick_face(names, default):
    """在系统已安装字体中挑选第一个可用的。"""
    try:
        available = set(wx.FontEnumerator.GetFacenames())
    except Exception:
        return default
    for name in names:
        if name in available:
            return name
    return default


def ui_font(pt=9.5, bold=False):
    """界面字体：优先微软雅黑 UI，跨平台依次回退。"""
    global _FACE_UI
    if _FACE_UI is None:
        _FACE_UI = _pick_face(
            ["Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC",
             "Noto Sans CJK SC", "Source Han Sans SC", "Segoe UI"],
            "Segoe UI")
    info = wx.FontInfo(pt).FaceName(_FACE_UI).Family(wx.FONTFAMILY_SWISS)
    if bold:
        info = info.Bold()
    return wx.Font(info)


def mono_font(pt=9.5):
    """等宽字体：用于 Markdown 输入框与日志。"""
    global _FACE_MONO
    if _FACE_MONO is None:
        _FACE_MONO = _pick_face(
            ["Cascadia Mono", "Cascadia Code", "Consolas", "JetBrains Mono",
             "Menlo", "DejaVu Sans Mono", "Courier New"],
            "Courier New")
    return wx.Font(wx.FontInfo(pt).FaceName(_FACE_MONO)
                   .Family(wx.FONTFAMILY_TELETYPE))


# --------------------------------------------------------------------------- #
# 自绘控件：wxPython 原生控件在 Windows 上外观陈旧且难以着色，这里用轻量的
# 自绘实现换取统一、克制的观感。
# --------------------------------------------------------------------------- #
class FlatButton(wx.Control):
    """扁平按钮。kind="primary" 为实心墨蓝，kind="secondary" 为白底细边框。"""

    def __init__(self, parent, label="", kind="secondary", size=(-1, 32)):
        super().__init__(parent, style=wx.BORDER_NONE)
        self._label = label
        self._kind = kind
        self._hover = False
        self._pressed = False
        self._focused = False
        self._min_height = size[1] if size and size[1] > 0 else 32
        self.SetBackgroundStyle(wx.BG_STYLE_PAINT)
        self.SetFont(ui_font(9.5))
        self.SetMinSize((-1, self._min_height))
        self.SetCursor(wx.Cursor(wx.CURSOR_HAND))
        self.Bind(wx.EVT_PAINT, self._on_paint)
        self.Bind(wx.EVT_ERASE_BACKGROUND, lambda _e: None)
        self.Bind(wx.EVT_ENTER_WINDOW, self._on_enter)
        self.Bind(wx.EVT_LEAVE_WINDOW, self._on_leave)
        self.Bind(wx.EVT_LEFT_DOWN, self._on_down)
        self.Bind(wx.EVT_LEFT_UP, self._on_up)
        self.Bind(wx.EVT_KEY_DOWN, self._on_key)
        self.Bind(wx.EVT_SET_FOCUS, self._on_focus)
        self.Bind(wx.EVT_KILL_FOCUS, self._on_blur)

    def DoGetBestSize(self):
        """按文字宽度自适应，左右各留 17px 内边距。"""
        try:
            tw, _th = self.GetTextExtent(self._label)
        except Exception:
            tw = 80
        return wx.Size(max(tw + 34, 64), self._min_height)

    # -- 外观 --
    def _palette(self):
        """返回 (底色, 文字色, 边框色)。边框为 None 表示不描边。"""
        if not self.IsEnabled():
            if self._kind == "primary":
                return T.DISABLED, T.SURFACE, None
            return T.SURFACE_ALT, T.DISABLED, T.BORDER
        if self._kind == "primary":
            if self._pressed:
                bg = T.ACCENT_LO
            elif self._hover:
                bg = T.ACCENT_HI
            else:
                bg = T.ACCENT
            return bg, T.SURFACE, None
        if self._pressed:
            bg = T.ACCENT_SOFT
        elif self._hover:
            bg = T.SURFACE_ALT
        else:
            bg = T.SURFACE
        border = T.FOCUS if self._focused else (T.BORDER_HI if self._hover else T.BORDER)
        return bg, T.TEXT, border

    def _on_paint(self, _evt):
        dc = wx.AutoBufferedPaintDC(self)
        dc.SetBackground(wx.Brush(self.GetParent().GetBackgroundColour()))
        dc.Clear()
        w, h = self.GetClientSize()
        bg, fg, border = self._palette()
        dc.SetBrush(wx.Brush(bg))
        dc.SetPen(wx.Pen(border, 1) if border else wx.TRANSPARENT_PEN)
        dc.DrawRoundedRectangle(0, 0, max(w - 1, 1), max(h - 1, 1), T.RADIUS_CTRL)
        dc.SetFont(self.GetFont())
        dc.SetTextForeground(fg)
        dc.DrawLabel(self._label, wx.Rect(0, 0, w, h), wx.ALIGN_CENTER)

    def _repaint(self):
        self.Refresh(False)

    # -- 交互 --
    def _on_enter(self, evt):
        self._hover = True
        self._repaint()
        evt.Skip()

    def _on_leave(self, evt):
        self._hover = False
        self._pressed = False
        self._repaint()
        evt.Skip()

    def _on_down(self, _evt):
        if not self.IsEnabled():
            return
        self._pressed = True
        if self.CanAcceptFocus():
            self.SetFocus()
        self._repaint()

    def _on_up(self, _evt):
        pressed = self._pressed
        self._pressed = False
        self._repaint()
        if pressed and self.IsEnabled():
            if self.GetClientRect().Contains(self.ScreenToClient(wx.GetMousePosition())):
                self._emit()

    def _on_key(self, evt):
        if evt.GetKeyCode() in (wx.WXK_SPACE, wx.WXK_RETURN, wx.WXK_NUMPAD_ENTER):
            if self.IsEnabled():
                self._emit()
        else:
            evt.Skip()

    def _on_focus(self, evt):
        self._focused = True
        self._repaint()
        evt.Skip()

    def _on_blur(self, evt):
        self._focused = False
        self._repaint()
        evt.Skip()

    def _emit(self):
        evt = wx.CommandEvent(wx.wxEVT_BUTTON, self.GetId())
        evt.SetEventObject(self)
        self.GetEventHandler().ProcessEvent(evt)

    # -- 兼容 wx.Button 接口 --
    def SetLabel(self, label):
        self._label = label
        self.InvalidateBestSize()
        self._repaint()

    def GetLabel(self):
        return self._label

    def Enable(self, enable=True):
        result = super().Enable(enable)
        self._repaint()
        return result


class Card(wx.Panel):
    """白色圆角面板，用作内容分组容器。"""

    def __init__(self, parent, radius=T.RADIUS_CARD):
        super().__init__(parent, style=wx.BORDER_NONE)
        self._radius = radius
        self.SetBackgroundColour(T.SURFACE)
        self.SetBackgroundStyle(wx.BG_STYLE_PAINT)
        self.Bind(wx.EVT_PAINT, self._on_paint)
        self.Bind(wx.EVT_ERASE_BACKGROUND, lambda _e: None)

    def _on_paint(self, _evt):
        dc = wx.AutoBufferedPaintDC(self)
        dc.SetBackground(wx.Brush(self.GetParent().GetBackgroundColour()))
        dc.Clear()
        w, h = self.GetClientSize()
        dc.SetBrush(wx.Brush(T.SURFACE))
        dc.SetPen(wx.Pen(T.BORDER, 1))
        dc.DrawRoundedRectangle(0, 0, max(w - 1, 1), max(h - 1, 1), self._radius)


class TextField(wx.Panel):
    """细边框输入容器，内部为无边框 TextCtrl（焦点时边框转为墨蓝）。"""

    def __init__(self, parent, multiline=False, readonly=False, mono=False,
                 hint=None):
        super().__init__(parent, style=wx.BORDER_NONE)
        self._focused = False
        self.SetBackgroundColour(T.SURFACE)
        self.SetBackgroundStyle(wx.BG_STYLE_PAINT)

        style = wx.BORDER_NONE
        if multiline:
            style |= wx.TE_MULTILINE | wx.TE_RICH2 | wx.HSCROLL
        if readonly:
            style |= wx.TE_READONLY

        self.text = wx.TextCtrl(self, style=style)
        self.text.SetBackgroundColour(T.SURFACE)
        self.text.SetForegroundColour(T.TEXT)
        self.text.SetFont(mono_font(9.5) if mono else ui_font(9.5))
        if hint:
            self.text.SetHint(hint)
        self.text.Bind(wx.EVT_SET_FOCUS, self._on_focus)
        self.text.Bind(wx.EVT_KILL_FOCUS, self._on_blur)

        inner = wx.BoxSizer(wx.HORIZONTAL)
        inner.Add(self.text, 1, wx.EXPAND | wx.LEFT | wx.RIGHT, 8)
        box = wx.BoxSizer(wx.VERTICAL)
        box.Add(inner, 1, wx.EXPAND | wx.TOP | wx.BOTTOM, 4)

        self.Bind(wx.EVT_PAINT, self._on_paint)
        self.Bind(wx.EVT_ERASE_BACKGROUND, lambda _e: None)
        self.SetSizer(box)
        if not multiline:
            self.SetMinSize((-1, 32))

    def _on_focus(self, evt):
        self._focused = True
        self.Refresh(False)
        evt.Skip()

    def _on_blur(self, evt):
        self._focused = False
        self.Refresh(False)
        evt.Skip()

    def _on_paint(self, _evt):
        dc = wx.AutoBufferedPaintDC(self)
        dc.SetBackground(wx.Brush(self.GetParent().GetBackgroundColour()))
        dc.Clear()
        w, h = self.GetClientSize()
        dc.SetBrush(wx.Brush(T.SURFACE))
        dc.SetPen(wx.Pen(T.FOCUS if self._focused else T.BORDER, 1))
        dc.DrawRoundedRectangle(0, 0, max(w - 1, 1), max(h - 1, 1), T.RADIUS_CTRL)


class ProgressBar(wx.Control):
    """细圆角进度条（高 6px），接口与 wx.Gauge 的常用子集一致。"""

    def __init__(self, parent, size=(-1, 6)):
        super().__init__(parent, size=size, style=wx.BORDER_NONE)
        self._range = 100
        self._value = 0
        self._height = size[1] if size and size[1] > 0 else 6
        self.SetBackgroundStyle(wx.BG_STYLE_PAINT)
        self.SetMinSize((-1, self._height))
        self.Bind(wx.EVT_PAINT, self._on_paint)
        self.Bind(wx.EVT_ERASE_BACKGROUND, lambda _e: None)

    def SetRange(self, value):
        self._range = max(int(value), 1)
        if self._value > self._range:
            self._value = self._range
        self.Refresh(False)

    def GetRange(self):
        return self._range

    def SetValue(self, value):
        self._value = max(0, min(int(value), self._range))
        self.Refresh(False)

    def GetValue(self):
        return self._value

    def _on_paint(self, _evt):
        dc = wx.AutoBufferedPaintDC(self)
        dc.SetBackground(wx.Brush(self.GetParent().GetBackgroundColour()))
        dc.Clear()
        w, h = self.GetClientSize()
        radius = h / 2.0
        dc.SetPen(wx.TRANSPARENT_PEN)
        dc.SetBrush(wx.Brush(T.TRACK))
        dc.DrawRoundedRectangle(0, 0, w, h, radius)
        filled = int(round(w * (self._value / float(self._range))))
        if filled > 0:
            dc.SetBrush(wx.Brush(T.ACCENT))
            dc.DrawRoundedRectangle(0, 0, max(filled, h), h, radius)

# 内置 Lua 过滤器：把带 \tag{X} 的块级公式转成“公式居中 + 编号右对齐”的无边框三列表格
LUA_EQNO = r"""
local TEMPLATE = "|  |  |  |\n|:--|:-:|--:|\n|  |  |  |\n"

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
  tbl.attr = pandoc.Attr("", { "md2word-equation" })
  return tbl
end

local function trim_inline_whitespace(inlines)
  while #inlines > 0 and
        (inlines[1].t == "Space" or inlines[1].t == "SoftBreak" or
         inlines[1].t == "LineBreak") do
    table.remove(inlines, 1)
  end
  while #inlines > 0 and
        (inlines[#inlines].t == "Space" or inlines[#inlines].t == "SoftBreak" or
         inlines[#inlines].t == "LineBreak") do
    table.remove(inlines)
  end
end

local function normalize_math_text(text)
  -- 兼容 AI/富文本编辑器常见的重复数学定界：$\(...\)$、$$\[...\]$$。
  -- Pandoc 已经根据外层 $ 生成 Math 节点，内层定界符必须去掉。
  local inner = text:match("^%s*\\%((.-)\\%)%s*$")
  if not inner then
    inner = text:match("^%s*\\%[(.-)\\%]%s*$")
  end
  if inner then
    text = inner
  end

  -- Markdown 转义有时会被带进公式，例如 \sum\_{i=1}^n；在数学环境中
  -- “下划线 + 花括号”应当是下标，而不是字面下划线。
  text = text:gsub("\\_(%b{})", "_%1")
  return text
end

function Math(el)
  local normalized = normalize_math_text(el.text)
  if normalized ~= el.text then
    el.text = normalized
    return el
  end
  return nil
end

function Para(el)
  -- Pandoc 允许 $$...$$ 与普通文字处于同一段。旧逻辑只处理“整段仅有
  -- 一个公式”的情况，导致“前文 $$...\tag{1}$$ 后文”中的 \tag 未被移除。
  -- 这里把这种混合段落拆成前文、编号公式表格、后文三个块。
  local blocks = {}
  local pending = {}
  local found = false

  local function flush_pending()
    trim_inline_whitespace(pending)
    if #pending > 0 then
      table.insert(blocks, pandoc.Para(pending))
    end
    pending = {}
  end

  for _, inline in ipairs(el.content) do
    if inline.t == "Math" and inline.mathtype == "DisplayMath" then
      local tag = inline.text:match("\\tag%s*{([^}]*)}")
      if tag then
        flush_pending()
        table.insert(blocks, build_numbered(inline, tag))
        found = true
      else
        table.insert(pending, inline)
      end
    else
      table.insert(pending, inline)
    end
  end

  if found then
    flush_pending()
    return blocks
  end
  return nil
end

function CodeBlock(el)
  -- 常见 AI/编辑器会用 ```math ... ``` 表示公式，但 Pandoc 默认把它当
  -- 代码块。将 class=math 的围栏代码块显式提升为 DisplayMath。
  local is_math = false
  for _, class in ipairs(el.classes) do
    if class == "math" then
      is_math = true
      break
    end
  end
  if not is_math then
    return nil
  end

  local text = normalize_math_text(el.text):gsub("^%s+", ""):gsub("%s+$", "")
  local math = pandoc.Math("DisplayMath", text)
  local tag = text:match("\\tag%s*{([^}]*)}")
  if tag then
    return build_numbered(math, tag)
  end
  return pandoc.Para({ math })
end
"""


def resource_path(rel):
    """返回随程序分发的资源文件的绝对路径。

    源码运行时相对本脚本所在目录；PyInstaller 打包后相对解包目录
    （sys._MEIPASS），这样同一份代码在两种方式下都能找到图标资源。
    """
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return os.path.join(base, rel)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), rel)


def find_pandoc():
    """优先使用打包内置的 pandoc，再查找系统安装。"""
    bundled = resource_path(os.path.join("pandoc", "pandoc.exe"))
    if os.path.isfile(bundled):
        return bundled
    p = shutil.which("pandoc")
    if p:
        return p
    candidates = []
    if sys.platform.startswith("win"):
        local = os.environ.get("LOCALAPPDATA", "")
        candidates = [
            os.path.join(local, "Pandoc", "pandoc.exe") if local else "",
            r"C:\Program Files\Pandoc\pandoc.exe",
            r"C:\Program Files (x86)\Pandoc\pandoc.exe",
        ]
    else:
        candidates = [
            "/usr/local/bin/pandoc",
            "/usr/bin/pandoc",
            "/opt/homebrew/bin/pandoc",
            "/snap/bin/pandoc",
        ]
    for c in candidates:
        if c and os.path.isfile(c):
            return c
    return None


# --------------------------------------------------------------------------- #
# 富文本剪贴板辅助：生成 Windows “HTML Format”(CF_HTML)，让 Word/WPS 粘贴时
# 把 MathML 转成原生可编辑公式。
# --------------------------------------------------------------------------- #
def _build_cf_html(html_doc):
    """把一段完整 HTML 文档封装成 CF_HTML 字节串（Windows 'HTML Format'）。"""
    marker_s = "<!--StartFragment-->"
    marker_e = "<!--EndFragment-->"

    doc = html_doc
    m = re.search(r"<body[^>]*>", doc, flags=re.I)
    if m:
        doc = doc[:m.end()] + marker_s + doc[m.end():]
    else:
        doc = marker_s + doc
    if re.search(r"</body>", doc, flags=re.I):
        doc = re.sub(r"</body>", marker_e + "</body>", doc, count=1, flags=re.I)
    else:
        doc = doc + marker_e

    header_fmt = ("Version:0.9\r\n"
                  "StartHTML:{:010d}\r\n"
                  "EndHTML:{:010d}\r\n"
                  "StartFragment:{:010d}\r\n"
                  "EndFragment:{:010d}\r\n")
    header_len = len(header_fmt.format(0, 0, 0, 0).encode("utf-8"))

    doc_bytes = doc.encode("utf-8")
    start_html = doc_bytes.find(b"<html")
    if start_html < 0:
        start_html = 0
    start_frag = doc_bytes.find(marker_s.encode("utf-8")) + len(marker_s)
    end_frag = doc_bytes.find(marker_e.encode("utf-8"))
    header = header_fmt.format(
        header_len + start_html,
        header_len + len(doc_bytes),
        header_len + start_frag,
        header_len + end_frag,
    )
    return header.encode("utf-8") + doc_bytes


def _html_to_text(html_doc):
    """把 HTML 粗略转成纯文本，作为剪贴板兜底格式。"""
    s = re.sub(r"(?is)<(script|style)[^>]*>.*?</\1>", "", html_doc)
    s = re.sub(r"(?is)<br\s*/?>", "\n", s)
    s = re.sub(r"(?is)</(p|div|h[1-6]|li|tr|table|blockquote|section)>", "\n", s)
    s = re.sub(r"(?is)<[^>]+>", "", s)
    s = html.unescape(s)
    s = re.sub(r"[ \t]+\n", "\n", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def _merge_html_docs(docs):
    """把多份 pandoc HTML 文档合并成一份（复用第一份的样式）。"""
    if not docs:
        return ""
    if len(docs) == 1:
        return docs[0]
    m = re.search(r"(?is)<style[^>]*>.*?</style>", docs[0])
    style = m.group(0) if m else ""
    parts = []
    for d in docs:
        mb = re.search(r"(?is)<body[^>]*>(.*?)</body>", d)
        parts.append(mb.group(1) if mb else d)
    sep = '<hr style="border:none;border-top:1px solid #ccc;margin:1.2em 0">'
    return ('<!DOCTYPE html><html><head><meta charset="utf-8">' + style
            + "</head><body>" + sep.join(parts) + "</body></html>")


def _normalize_equation_tables_html(html_doc):
    """固定富文本中公式编号表的宽度、边距及水平/垂直对齐。"""
    table_re = re.compile(
        r'(?is)<table(?P<attrs>[^>]*\bclass="[^"]*\bmd2word-equation\b[^"]*"[^>]*)>'
        r'(?P<body>.*?)</table>')

    def fix_table(match):
        attrs = match.group("attrs")
        body = match.group("body")
        style_match = re.search(r'(?is)\bstyle="([^"]*)"', attrs)
        forced = ("width:100%;table-layout:fixed;border-collapse:collapse;"
                  "margin-left:0;margin-right:0")
        if style_match:
            old = re.sub(
                r'(?i)(?:width|table-layout|border-collapse|margin-left|margin-right)'
                r'\s*:[^;]+;?', "", style_match.group(1)).rstrip("; ")
            style = (old + ";" if old else "") + forced
            attrs = (attrs[:style_match.start()] + 'style="' + style + '"'
                     + attrs[style_match.end():])
        else:
            attrs += ' style="' + forced + '"'

        cells = list(re.finditer(r'(?is)<td(?P<attrs>[^>]*)>', body))
        aligns = ("left", "center", "right")
        for i, cell in reversed(list(enumerate(cells[:3]))):
            cell_attrs = cell.group("attrs")
            css = "padding:0;vertical-align:middle;text-align:%s" % aligns[i]
            sm = re.search(r'(?is)\bstyle="([^"]*)"', cell_attrs)
            if sm:
                old = re.sub(r'(?i)(?:padding|vertical-align|text-align)\s*:[^;]+;?',
                             "", sm.group(1)).rstrip("; ")
                new_style = (old + ";" if old else "") + css
                cell_attrs = (cell_attrs[:sm.start()] + 'style="' + new_style + '"'
                              + cell_attrs[sm.end():])
            else:
                cell_attrs += ' style="' + css + '"'
            body = body[:cell.start()] + "<td" + cell_attrs + ">" + body[cell.end():]
        return "<table" + attrs + ">" + body + "</table>"

    return table_re.sub(fix_table, html_doc)


def _normalize_equation_tables_docx(path):
    """强制编号公式表在 Word/WPS 中按正文宽度精确对齐。

    Pandoc 会生成正确的三列比例，但参考文档中的 Table 样式仍可能注入表格
    缩进、单元格边距或顶部对齐。本函数只修改“中间含 OMML、右侧为编号”的
    单行三列表格，不碰用户自己的普通表格。
    """
    with zipfile.ZipFile(path, "r") as zin:
        infos = zin.infolist()
        files = {info.filename: zin.read(info.filename) for info in infos}

    name = "word/document.xml"
    raw = files.get(name)
    if not raw:
        return 0
    xml = raw.decode("utf-8")
    changed = 0

    def force_ppr(cell, align):
        def update(match):
            ppr = match.group(0)
            ppr = re.sub(r'<w:jc\b[^>]*/>', "", ppr)
            ppr = re.sub(r'<w:ind\b[^>]*/>', "", ppr)
            ppr = re.sub(r'<w:spacing\b[^>]*/>', "", ppr)
            forced = ('<w:jc w:val="%s" />'
                      '<w:ind w:left="0" w:right="0" w:firstLine="0" />'
                      '<w:spacing w:before="0" w:after="0" />') % align
            return ppr.replace("<w:pPr>", "<w:pPr>" + forced, 1)

        if re.search(r'<w:pPr\s*/>', cell):
            forced = ('<w:pPr><w:jc w:val="%s" />'
                      '<w:ind w:left="0" w:right="0" w:firstLine="0" />'
                      '<w:spacing w:before="0" w:after="0" /></w:pPr>') % align
            return re.sub(r'<w:pPr\s*/>', forced, cell, count=1)
        if "<w:pPr>" in cell:
            return re.sub(r'<w:pPr>.*?</w:pPr>', update, cell,
                          count=1, flags=re.S)
        return cell.replace("<w:p>",
                            '<w:p><w:pPr><w:jc w:val="%s" />'
                            '<w:ind w:left="0" w:right="0" w:firstLine="0" />'
                            '<w:spacing w:before="0" w:after="0" />'
                            '</w:pPr>' % align, 1)

    def force_cell(cell, width, align):
        props = ('<w:tcW w:w="%d" w:type="pct" />'
                 '<w:vAlign w:val="center" />') % width
        if re.search(r'<w:tcPr\s*/>', cell):
            cell = re.sub(r'<w:tcPr\s*/>', '<w:tcPr>' + props + '</w:tcPr>',
                          cell, count=1)
        elif "<w:tcPr>" in cell:
            cell = re.sub(r'<w:tcW\b[^>]*/>|<w:vAlign\b[^>]*/>', "", cell)
            cell = cell.replace("<w:tcPr>", "<w:tcPr>" + props, 1)
        else:
            cell = cell.replace("<w:tc>", "<w:tc><w:tcPr>" + props + "</w:tcPr>", 1)
        return force_ppr(cell, align)

    def fix_table(match):
        nonlocal changed
        table = match.group(0)
        cells = re.findall(r'<w:tc>.*?</w:tc>', table, flags=re.S)
        if (len(cells) != 3 or "<m:oMath" not in cells[1] or
                not re.search(r'<w:t[^>]*>\s*\([^<]+\)\s*</w:t>', cells[2])):
            return table

        tbl_props = (
            '<w:tblW w:w="5000" w:type="pct" />'
            '<w:jc w:val="center" />'
            '<w:tblInd w:w="0" w:type="dxa" />'
            '<w:tblLayout w:type="fixed" />'
            '<w:tblCellMar>'
            '<w:top w:w="0" w:type="dxa" />'
            '<w:left w:w="0" w:type="dxa" />'
            '<w:bottom w:w="0" w:type="dxa" />'
            '<w:right w:w="0" w:type="dxa" />'
            '</w:tblCellMar>')
        def clean_tblpr(tblpr_match):
            inner = tblpr_match.group(0)[9:-10]
            for prop in ("tblW", "jc", "tblInd", "tblLayout", "tblCellMar"):
                inner = re.sub(
                    r'<w:%s\b[^>]*/>|<w:%s\b[^>]*>.*?</w:%s>' %
                    (prop, prop, prop), "", inner, flags=re.S)
            return '<w:tblPr>' + tbl_props + inner + '</w:tblPr>'

        table = re.sub(r'<w:tblPr>.*?</w:tblPr>', clean_tblpr,
                       table, count=1, flags=re.S)

        old_cells = re.findall(r'<w:tc>.*?</w:tc>', table, flags=re.S)
        new_cells = [
            force_cell(old_cells[0], 400, "left"),
            force_cell(old_cells[1], 4200, "center"),
            force_cell(old_cells[2], 400, "right"),
        ]
        for old, new in zip(old_cells, new_cells):
            table = table.replace(old, new, 1)
        changed += 1
        return table

    xml = re.sub(r'<w:tbl>.*?</w:tbl>', fix_table, xml, flags=re.S)
    if not changed:
        return 0
    files[name] = xml.encode("utf-8")

    tmp = path + ".layout.tmp"
    try:
        with zipfile.ZipFile(tmp, "w") as zout:
            for info in infos:
                zout.writestr(info, files[info.filename])
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return changed


# --------------------------------------------------------------------------- #
# 主窗口
# --------------------------------------------------------------------------- #
class MainFrame(wx.Frame):
    def __init__(self):
        super().__init__(None, title=APP_TITLE, size=(1280, 820))
        self.pandoc = find_pandoc()
        # entries: [{"path": abspath, "rel": 相对子目录(用于保留结构，可为 "")}]
        self.entries = []
        self.busy = False
        self.last_outputs = []   # 本次成功生成的 .docx 路径
        self._dst_by_src = {}    # 源文件 -> 输出文件

        self._apply_icon()
        self._build_ui()
        self.Centre()

        if not self.pandoc:
            self.log("✗ 未找到 pandoc；源码运行请先安装，打包版请重新下载完整程序。",
                     COL_STATUS_ERR)
            wx.MessageBox(
                "未检测到 pandoc，转换功能不可用。\n"
                "若从源码运行，请安装 pandoc 并加入系统 PATH；\n"
                "若使用打包版，请重新下载完整的程序文件。",
                "缺少 pandoc", wx.OK | wx.ICON_ERROR)

    # ------------------------------------------------------------------ UI --
    def _apply_icon(self):
        """设置窗口图标：标题栏（左上角）、任务栏与 Alt+Tab 统一使用 logo。"""
        for rel in ("logo/logo.ico", "logo/logo.png"):
            path = resource_path(rel)
            if not os.path.isfile(path):
                continue
            kind = (wx.BITMAP_TYPE_ICO if path.lower().endswith(".ico")
                    else wx.BITMAP_TYPE_PNG)
            try:
                icon = wx.Icon(path, kind)
            except Exception:
                continue
            if icon.IsOk():
                self.SetIcon(icon)
                return

    def _build_ui(self):
        self.SetBackgroundColour(T.BG)
        panel = wx.Panel(self)
        panel.SetBackgroundColour(T.BG)

        outer = wx.BoxSizer(wx.HORIZONTAL)

        # ==================== 左栏：Markdown 粘贴输入 ====================
        left_card = Card(panel)
        lc = wx.BoxSizer(wx.VERTICAL)
        lc.Add(self._title(left_card, "Markdown 输入"), 0)
        hint = self._muted(
            left_card,
            "粘贴内容后点击下方按钮，结果将复制为可直接粘贴的富文本（含 Word 原生公式）")
        hint.Wrap(420)
        lc.Add(hint, 0, wx.TOP, 4)

        self._md_field = TextField(left_card, multiline=True, mono=True)
        self.txt_md = self._md_field.text
        lc.Add(self._md_field, 1, wx.EXPAND | wx.TOP, 12)

        md_tools = wx.BoxSizer(wx.HORIZONTAL)
        self.btn_md_clear = FlatButton(left_card, "清空")
        self.btn_md_paste = FlatButton(left_card, "粘贴")
        md_tools.Add(self.btn_md_clear, 1, wx.RIGHT, 8)
        md_tools.Add(self.btn_md_paste, 1)
        lc.Add(md_tools, 0, wx.EXPAND | wx.TOP, 12)

        self.btn_md_convert = FlatButton(
            left_card, "转换并复制（含公式）", kind="primary", size=(-1, 34))
        lc.Add(self.btn_md_convert, 0, wx.EXPAND | wx.TOP, 10)

        wrap = wx.BoxSizer(wx.VERTICAL)
        wrap.Add(lc, 1, wx.EXPAND | wx.ALL, T.PAD)
        left_card.SetSizer(wrap)
        outer.Add(left_card, 3, wx.EXPAND | wx.TOP | wx.BOTTOM | wx.LEFT, 12)

        # ==================== 右栏：文件批量转换 ====================
        right = wx.BoxSizer(wx.VERTICAL)

        # --- 待转换文件 ---
        files_card = Card(panel)
        fc = wx.BoxSizer(wx.VERTICAL)
        fc.Add(self._title(files_card, "待转换文件"), 0)

        bar = wx.BoxSizer(wx.HORIZONTAL)
        self.btn_add_files = FlatButton(files_card, "添加文件…")
        self.btn_add_dir = FlatButton(files_card, "添加文件夹…")
        self.btn_remove = FlatButton(files_card, "移除选中")
        self.btn_clear = FlatButton(files_card, "清空列表")
        self.chk_recursive = self._checkbox(files_card, "递归子文件夹", checked=True)
        bar.Add(self.btn_add_files, 0, wx.RIGHT, 8)
        bar.Add(self.btn_add_dir, 0, wx.RIGHT, 8)
        bar.AddSpacer(20)
        bar.Add(self.btn_remove, 0, wx.RIGHT, 8)
        bar.Add(self.btn_clear, 0)
        bar.AddStretchSpacer(1)
        bar.Add(self.chk_recursive, 0, wx.ALIGN_CENTER_VERTICAL)
        fc.Add(bar, 0, wx.EXPAND | wx.TOP, 12)

        self.list = wx.dataview.DataViewListCtrl(
            files_card, style=wx.dataview.DV_SINGLE | wx.dataview.DV_ROW_LINES)
        self.list.SetBackgroundColour(T.SURFACE)
        self.list.SetFont(ui_font(9.5))
        self.list.AppendTextColumn("文件名", width=210)
        self.list.AppendTextColumn("所在目录", width=380)
        self.list.AppendTextColumn("状态", width=100)
        try:
            self.list.SetAlternateRowColour(T.ROW_ALT)
        except Exception:
            pass
        fc.Add(self.list, 1, wx.EXPAND | wx.TOP, 12)

        wrap = wx.BoxSizer(wx.VERTICAL)
        wrap.Add(fc, 1, wx.EXPAND | wx.ALL, T.PAD)
        files_card.SetSizer(wrap)
        right.Add(files_card, 1, wx.EXPAND)

        # --- 转换选项 ---
        opt_card = Card(panel)
        oc = wx.BoxSizer(wx.VERTICAL)
        oc.Add(self._title(opt_card, "转换选项"), 0)

        grid = wx.FlexGridSizer(cols=3, vgap=10, hgap=10)
        grid.AddGrowableCol(1, 1)

        grid.Add(self._field_label(opt_card, "输出目录"), 0, wx.ALIGN_CENTER_VERTICAL)
        self._out_field = TextField(opt_card, hint="留空时请勾选“输出到源文件所在目录”")
        self.txt_outdir = self._out_field.text
        self.txt_outdir.Enable(False)
        grid.Add(self._out_field, 1, wx.EXPAND)
        self.btn_outdir = FlatButton(opt_card, "浏览…", size=(78, 32))
        self.btn_outdir.Enable(False)
        grid.Add(self.btn_outdir, 0)

        grid.Add(self._field_label(opt_card, "输入格式"), 0, wx.ALIGN_CENTER_VERTICAL)
        self._from_field = TextField(
            opt_card, hint="留空 = 自动识别（公式安全，默认按 markdown 处理）")
        self.txt_from = self._from_field.text
        grid.Add(self._from_field, 1, wx.EXPAND)
        grid.Add(self._blank(opt_card), 0)

        grid.Add(self._field_label(opt_card, "参考样式 docx"), 0, wx.ALIGN_CENTER_VERTICAL)
        self._ref_field = TextField(opt_card, hint="可选：用某个 .docx 作为样式模板")
        self.txt_ref = self._ref_field.text
        grid.Add(self._ref_field, 1, wx.EXPAND)
        self.btn_ref = FlatButton(opt_card, "浏览…", size=(78, 32))
        grid.Add(self.btn_ref, 0)

        grid.Add(self._field_label(opt_card, "pandoc 额外参数"), 0, wx.ALIGN_CENTER_VERTICAL)
        self._extra_field = TextField(opt_card, hint="可选：例如 --number-sections")
        self.txt_extra = self._extra_field.text
        grid.Add(self._extra_field, 1, wx.EXPAND)
        grid.Add(self._blank(opt_card), 0)

        oc.Add(grid, 0, wx.EXPAND | wx.TOP, 14)

        self.chk_math = self._checkbox(
            opt_card, "LaTeX 公式 → Word 原生公式（\\tag 编号右对齐）", checked=True)
        oc.Add(self.chk_math, 0, wx.TOP, 16)

        chkrow = wx.BoxSizer(wx.HORIZONTAL)
        self.chk_to_source = self._checkbox(opt_card, "输出到源文件所在目录")
        self.chk_keep_struct = self._checkbox(opt_card, "保留相对目录结构")
        self.chk_overwrite = self._checkbox(opt_card, "覆盖已存在的文件")
        self.chk_toc = self._checkbox(opt_card, "生成目录 (--toc)")
        for c in (self.chk_to_source, self.chk_keep_struct,
                  self.chk_overwrite, self.chk_toc):
            chkrow.Add(c, 0, wx.RIGHT, 22)
        oc.Add(chkrow, 0, wx.TOP, 12)

        wrap = wx.BoxSizer(wx.VERTICAL)
        wrap.Add(oc, 1, wx.EXPAND | wx.ALL, T.PAD)
        opt_card.SetSizer(wrap)
        right.Add(opt_card, 0, wx.EXPAND | wx.TOP, 12)

        # --- 转换动作与日志 ---
        run_card = Card(panel)
        rc = wx.BoxSizer(wx.VERTICAL)

        actions = wx.BoxSizer(wx.HORIZONTAL)
        self.btn_convert = FlatButton(run_card, "开始转换", kind="primary", size=(112, 32))
        self.btn_copy = FlatButton(run_card, "复制内容", size=(96, 32))
        self.btn_open = FlatButton(run_card, "打开输出目录", size=(116, 32))
        self.btn_copy.Enable(False)
        self.btn_open.Enable(False)
        actions.Add(self.btn_convert, 0, wx.RIGHT, 10)
        actions.Add(self.btn_copy, 0, wx.RIGHT, 10)
        actions.Add(self.btn_open, 0, wx.RIGHT, 18)
        self.gauge = ProgressBar(run_card, size=(-1, 6))
        actions.Add(self.gauge, 1, wx.ALIGN_CENTER_VERTICAL)
        rc.Add(actions, 0, wx.EXPAND)

        self._log_field = TextField(run_card, multiline=True, readonly=True, mono=True)
        self.log_ctrl = self._log_field.text
        self._log_field.SetMinSize((-1, 126))
        rc.Add(self._log_field, 0, wx.EXPAND | wx.TOP, 12)

        wrap = wx.BoxSizer(wx.VERTICAL)
        wrap.Add(rc, 1, wx.EXPAND | wx.ALL, T.PAD)
        run_card.SetSizer(wrap)
        right.Add(run_card, 0, wx.EXPAND | wx.TOP, 12)

        outer.Add(right, 5, wx.EXPAND | wx.ALL, 12)
        panel.SetSizer(outer)

        # --- 事件绑定 ---
        self.btn_add_files.Bind(wx.EVT_BUTTON, self.on_add_files)
        self.btn_add_dir.Bind(wx.EVT_BUTTON, self.on_add_dir)
        self.btn_remove.Bind(wx.EVT_BUTTON, self.on_remove)
        self.btn_clear.Bind(wx.EVT_BUTTON, self.on_clear)
        self.btn_outdir.Bind(wx.EVT_BUTTON, self.on_pick_outdir)
        self.btn_ref.Bind(wx.EVT_BUTTON, self.on_pick_ref)
        self.btn_convert.Bind(wx.EVT_BUTTON, self.on_convert)
        self.btn_open.Bind(wx.EVT_BUTTON, self.on_open_outdir)
        self.btn_copy.Bind(wx.EVT_BUTTON, self.on_copy_content)
        self.btn_md_convert.Bind(wx.EVT_BUTTON, self.on_md_convert)
        self.btn_md_clear.Bind(wx.EVT_BUTTON, self.on_md_clear)
        self.btn_md_paste.Bind(wx.EVT_BUTTON, self.on_md_paste)
        self.chk_to_source.Bind(wx.EVT_CHECKBOX, self.on_toggle_outdir)
        self.Bind(wx.EVT_CLOSE, self.on_close)

    # ------------------------------------------------------------ 视图辅助 --
    @staticmethod
    def _title(parent, text):
        st = wx.StaticText(parent, label=text)
        st.SetFont(ui_font(10.5, bold=True))
        st.SetForegroundColour(T.TEXT)
        return st

    @staticmethod
    def _muted(parent, text):
        st = wx.StaticText(parent, label=text)
        st.SetFont(ui_font(9))
        st.SetForegroundColour(T.TEXT_MUTED)
        return st

    @staticmethod
    def _field_label(parent, text):
        st = wx.StaticText(parent, label=text)
        st.SetFont(ui_font(9.5))
        st.SetForegroundColour(T.TEXT_MUTED)
        return st

    @staticmethod
    def _blank(parent):
        return wx.StaticText(parent, label="")

    @staticmethod
    def _checkbox(parent, text, checked=False):
        cb = wx.CheckBox(parent, label=text)
        cb.SetValue(checked)
        cb.SetFont(ui_font(9.5))
        cb.SetForegroundColour(T.TEXT)
        return cb

    # ------------------------------------------------------------- 工具方法 --
    def log(self, msg, colour=None):
        def _do():
            self.log_ctrl.SetDefaultStyle(
                wx.TextAttr(colour) if colour else wx.TextAttr())
            self.log_ctrl.AppendText(msg + "\n")
        wx.CallAfter(_do)

    @staticmethod
    def _is_md(path):
        return os.path.splitext(path)[1].lower() in MD_EXTS


    def _set_status(self, idx, text, colour=None):
        """更新列表中某一行的状态文本（需在主线程调用）。"""
        try:
            self.list.SetValue(text, idx, 2)
            if colour is not None:
                attr = wx.dataview.DataViewItemAttr()
                attr.SetColour(colour)
                self.list.SetAttr(idx, 2, attr)
        except Exception:
            pass
    def _add_entry(self, path, rel=""):
        path = os.path.abspath(path)
        if any(e["path"] == path for e in self.entries):
            return False
        self.entries.append({"path": path, "rel": rel})
        self.list.AppendItem([os.path.basename(path), os.path.dirname(path), "待转换"])
        return True

    def _set_busy(self, busy):
        self.busy = busy
        for w in (self.btn_add_files, self.btn_add_dir, self.btn_remove,
                  self.btn_clear, self.btn_convert, self.btn_md_convert,
                  self.btn_md_clear, self.btn_md_paste):
            w.Enable(not busy)
        self.chk_recursive.Enable(not busy)

    # ---------------------------------------------------------------- 事件 --
    def on_add_files(self, _evt):
        dlg = wx.FileDialog(
            self, "选择 Markdown 文件", wildcard="Markdown (*.md;*.markdown)|*.md;*.markdown|所有文件 (*.*)|*.*",
            style=wx.FD_OPEN | wx.FD_MULTIPLE)
        if dlg.ShowModal() == wx.ID_OK:
            added = 0
            for p in dlg.GetPaths():
                if self._add_entry(p):
                    added += 1
            self.log("已添加 %d 个文件。" % added)
        dlg.Destroy()

    def on_add_dir(self, _evt):
        dlg = wx.DirDialog(self, "选择包含 Markdown 的文件夹")
        if dlg.ShowModal() == wx.ID_OK:
            folder = dlg.GetPath()
            recursive = self.chk_recursive.GetValue()
            added = 0
            if recursive:
                for dirpath, _dirs, files in os.walk(folder):
                    for fn in files:
                        if self._is_md(fn):
                            full = os.path.join(dirpath, fn)
                            rel = os.path.relpath(dirpath, folder)
                            if rel == ".":
                                rel = ""
                            if self._add_entry(full, rel):
                                added += 1
            else:
                for fn in os.listdir(folder):
                    full = os.path.join(folder, fn)
                    if os.path.isfile(full) and self._is_md(fn):
                        if self._add_entry(full, ""):
                            added += 1
            self.log("从文件夹添加 %d 个文件。" % added)
        dlg.Destroy()

    def on_remove(self, _evt):
        idx = self.list.GetSelectedRow()
        if idx == wx.NOT_FOUND or idx < 0 or idx >= len(self.entries):
            return
        self.list.DeleteItem(idx)
        del self.entries[idx]

    def on_clear(self, _evt):
        self.list.DeleteAllItems()
        self.entries.clear()
        self.last_outputs = []
        self._dst_by_src = {}
        self.btn_copy.Enable(False)
        self.log("已清空列表。")

    def on_toggle_outdir(self, _evt):
        to_source = self.chk_to_source.GetValue()
        self.txt_outdir.Enable(not to_source and not self.busy)
        self.btn_outdir.Enable(not to_source and not self.busy)
        self._out_field.Refresh(False)

    def on_pick_outdir(self, _evt):
        dlg = wx.DirDialog(self, "选择输出目录")
        if dlg.ShowModal() == wx.ID_OK:
            self.txt_outdir.SetValue(dlg.GetPath())
        dlg.Destroy()

    def on_pick_ref(self, _evt):
        dlg = wx.FileDialog(
            self, "选择参考样式文档", wildcard="Word 文档 (*.docx)|*.docx",
            style=wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_OK:
            self.txt_ref.SetValue(dlg.GetPath())
        dlg.Destroy()

    def on_open_outdir(self, _evt):
        path = self._effective_base_outdir()
        if path and os.path.isdir(path):
            self._open_path(path)

    def _collect_copy_targets(self):
        """决定“复制内容”要复制哪些源 md。

        若列表中选中了某个已成功转换的文件，则只复制它；否则复制本次全部
        成功转换的文件。
        """
        outs = set(p for p in self.last_outputs if os.path.isfile(p))
        if not outs:
            return []
        idx = self.list.GetSelectedRow()
        if idx != wx.NOT_FOUND and 0 <= idx < len(self.entries):
            src = self.entries[idx]["path"]
            if self._dst_by_src.get(src) in outs:
                return [src]
        return [s for s, d in self._dst_by_src.items() if d in outs]

    def on_copy_content(self, _evt):
        if self.busy:
            return
        sources = self._collect_copy_targets()
        if not sources:
            wx.MessageBox("还没有可复制的转换结果。", "提示",
                          wx.OK | wx.ICON_INFORMATION)
            return
        self._start_rich_copy(sources)

    # ------------------------------------------------- 生成“可粘贴”富文本 --
    def _common_pandoc_opts(self):
        """HTML 输出复用的公共 pandoc 选项（在界面线程里取值）。"""
        opts = []
        from_fmt = self._resolve_from_format()
        if from_fmt:
            opts += ["-f", from_fmt]
        if self.chk_toc.GetValue():
            opts += ["--toc"]
        if self.chk_math.GetValue():
            opts += ["--lua-filter", self._ensure_eqno_filter()]
        extra = self.txt_extra.GetValue().strip()
        if extra:
            opts += self._split_args(extra)
        return opts

    def _run_pandoc_html(self, src, opts):
        """把单个 md 源文件转成带 MathML 的完整 HTML（在后台线程调用）。"""
        cmd = [self.pandoc, src, "-t", "html", "--standalone", "--mathml"] + opts
        kwargs = {}
        if sys.platform.startswith("win"):
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        proc = subprocess.run(cmd, capture_output=True, text=True,
                              encoding="utf-8", errors="replace", **kwargs)
        if proc.returncode != 0:
            err = (proc.stderr or proc.stdout or "").strip()
            raise RuntimeError(err or ("pandoc 返回 %d" % proc.returncode))
        return _normalize_equation_tables_html(proc.stdout)

    def _start_rich_copy(self, sources, cleanup_dir=None, button=None):
        opts = self._common_pandoc_opts()
        self._copy_button = button
        self._copy_label = button.GetLabel() if button is not None else None
        self._set_busy(True)
        if button is not None:
            button.SetLabel("转换中…")
        threading.Thread(target=self._copy_worker,
                         args=(sources, opts, cleanup_dir), daemon=True).start()

    def _copy_worker(self, sources, opts, cleanup_dir):
        docs, err = [], None
        try:
            for src in sources:
                docs.append(self._run_pandoc_html(src, opts))
        except Exception as ex:
            err = str(ex)
        finally:
            if cleanup_dir:
                shutil.rmtree(cleanup_dir, ignore_errors=True)
        if err:
            wx.CallAfter(self._copy_done, None, err)
            return
        merged = _merge_html_docs(docs)
        wx.CallAfter(self._copy_done, (merged, _html_to_text(merged)), None)

    def _copy_done(self, payload, err):
        self._set_busy(False)
        btn = getattr(self, "_copy_button", None)
        if btn is not None and getattr(self, "_copy_label", None):
            btn.SetLabel(self._copy_label)
        if err:
            self.log("✗ 生成可粘贴内容失败：%s" % err, COL_STATUS_ERR)
            wx.MessageBox("生成失败：\n%s" % err, "错误", wx.OK | wx.ICON_ERROR)
            return
        html_doc, plain = payload
        if self._set_clipboard_rich(html_doc, plain):
            self.log("✓ 已复制富文本（含 Word 公式）到剪贴板（%d 字符）。" % len(plain),
                     COL_STATUS_OK)
        else:
            wx.MessageBox("无法写入剪贴板。", "错误", wx.OK | wx.ICON_ERROR)

    def _set_clipboard_rich(self, html_doc, plain):
        """把 HTML(含 MathML) + 纯文本兜底写入剪贴板。"""
        data = _build_cf_html(html_doc)
        if not wx.TheClipboard.Open():
            return False
        try:
            composite = wx.DataObjectComposite()
            html_obj = wx.CustomDataObject(wx.DataFormat("HTML Format"))
            html_obj.SetData(data)
            composite.Add(html_obj, True)
            composite.Add(wx.TextDataObject(plain), False)
            wx.TheClipboard.SetData(composite)
            wx.TheClipboard.Flush()
        except Exception as ex:
            self.log("✗ 剪贴板写入失败：%s" % ex, COL_STATUS_ERR)
            return False
        finally:
            wx.TheClipboard.Close()
        return True

    def on_md_clear(self, _evt):
        """清空左侧 Markdown 输入框（误操作可用 Ctrl+Z 撤销）。"""
        if self.busy:
            return
        if not self.txt_md.GetValue():
            return
        self.txt_md.Clear()
        self.txt_md.SetFocus()

    def on_md_paste(self, _evt):
        """把剪贴板中的文本粘贴到输入框当前光标处。"""
        if self.busy:
            return
        text = self._clipboard_text()
        if not text:
            wx.MessageBox("剪贴板中没有可粘贴的文本。", "提示",
                          wx.OK | wx.ICON_INFORMATION)
            return
        self.txt_md.WriteText(text)
        self.txt_md.SetFocus()

    @staticmethod
    def _clipboard_text():
        """读取剪贴板中的纯文本；无法取到（含非文本内容）时返回空串。"""
        if not wx.TheClipboard.Open():
            return ""
        try:
            data = wx.TextDataObject()
            if wx.TheClipboard.GetData(data):
                return data.GetText()
        finally:
            wx.TheClipboard.Close()
        return ""

    def on_md_convert(self, _evt):
        """把左侧粘贴的 Markdown 转成富文本（含 Word 公式）并复制到剪贴板。"""
        if self.busy:
            return
        if not self.pandoc:
            wx.MessageBox("未找到 pandoc，无法转换。", "错误", wx.OK | wx.ICON_ERROR)
            return
        md = self.txt_md.GetValue()
        if not md.strip():
            wx.MessageBox("请先在左侧粘贴 Markdown 内容。", "提示",
                          wx.OK | wx.ICON_INFORMATION)
            return

        tmpdir = tempfile.mkdtemp(prefix="md2word_")
        md_path = os.path.join(tmpdir, "input.md")
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md)
        self._start_rich_copy([md_path], cleanup_dir=tmpdir, button=self.btn_md_convert)

    def on_close(self, evt):
        if self.busy:
            if wx.MessageBox("转换正在进行中，确定要退出吗？", "确认",
                             wx.YES_NO | wx.ICON_QUESTION) != wx.YES:
                return
        evt.Skip()

    # ------------------------------------------------------------ 转换逻辑 --
    def _effective_base_outdir(self):
        """用于“打开输出目录”的基准目录。"""
        if self.chk_to_source.GetValue():
            return os.path.dirname(self.entries[0]["path"]) if self.entries else ""
        return self.txt_outdir.GetValue().strip()

    def on_convert(self, _evt):
        if self.busy:
            return
        if not self.pandoc:
            wx.MessageBox("未找到 pandoc，无法转换。", "错误", wx.OK | wx.ICON_ERROR)
            return
        if not self.entries:
            wx.MessageBox("请先添加要转换的 Markdown 文件。", "提示", wx.OK | wx.ICON_INFORMATION)
            return

        to_source = self.chk_to_source.GetValue()
        outdir = self.txt_outdir.GetValue().strip()
        if not to_source:
            if not outdir:
                wx.MessageBox("请选择输出目录。", "提示", wx.OK | wx.ICON_INFORMATION)
                return
            if not os.path.isdir(outdir):
                try:
                    os.makedirs(outdir, exist_ok=True)
                except OSError as e:
                    wx.MessageBox("无法创建输出目录：\n%s" % e, "错误", wx.OK | wx.ICON_ERROR)
                    return

        ref = self.txt_ref.GetValue().strip()
        if ref and not os.path.isfile(ref):
            wx.MessageBox("参考样式文档不存在：\n%s" % ref, "错误", wx.OK | wx.ICON_ERROR)
            return

        jobs = self._build_jobs(to_source, outdir, ref)
        self._dst_by_src = {j["src"]: j["dst"] for j in jobs}
        self.last_outputs = []
        self.btn_copy.Enable(False)

        self._set_busy(True)
        self.gauge.SetValue(0)
        self.gauge.SetRange(len(jobs))
        self.log_ctrl.Clear()
        self.log("开始转换，共 %d 个文件…" % len(jobs), COL_STATUS_OK)
        self.log("输入格式：%s" % (self._resolve_from_format() or "自动识别"), COL_STATUS_OK)
        threading.Thread(target=self._worker, args=(jobs,), daemon=True).start()

    def _ensure_eqno_filter(self):
        """把内置的 \\tag 处理过滤器写入临时文件，返回其路径。"""
        path = os.path.join(tempfile.gettempdir(), "md2word_eqno.lua")
        with open(path, "w", encoding="utf-8") as f:
            f.write(LUA_EQNO)
        return path

    def _resolve_from_format(self):
        """根据“公式转换”选项，确保输入格式启用 LaTeX 数学解析。

        - 选项关闭：原样返回用户填写的输入格式。
        - 选项开启：在格式串上追加 tex_math_dollars / tex_math_single_backslash；
          若用户留空，则改用 markdown（其块级公式解析比 gfm 稳，gfm 会把
          单独一行的 `=` 误判为 setext 标题而打断 $$...$$）。
        """
        fmt = self.txt_from.GetValue().strip()
        if not self.chk_math.GetValue():
            return fmt
        base = fmt if fmt else "markdown"
        for ext in ("tex_math_dollars", "tex_math_single_backslash"):
            if ext not in base:
                base += "+" + ext
        return base

    def _build_jobs(self, to_source, outdir, ref):
        jobs = []
        keep_struct = (not to_source) and self.chk_keep_struct.GetValue()
        overwrite = self.chk_overwrite.GetValue()
        from_fmt = self._resolve_from_format()
        extra = self.txt_extra.GetValue().strip()
        toc = self.chk_toc.GetValue()
        eqno_lua = self._ensure_eqno_filter() if self.chk_math.GetValue() else None

        for i, e in enumerate(self.entries):
            src = e["path"]
            if to_source:
                dest_dir = os.path.dirname(src)
            else:
                dest_dir = os.path.join(outdir, e["rel"]) if (keep_struct and e["rel"]) else outdir
            out_name = os.path.splitext(os.path.basename(src))[0] + ".docx"
            dst = os.path.join(dest_dir, out_name)

            cmd = [self.pandoc, src, "-o", dst, "--standalone"]
            if from_fmt:
                cmd += ["-f", from_fmt]
            if ref:
                cmd += ["--reference-doc", ref]
            if toc:
                cmd += ["--toc"]
            if eqno_lua:
                cmd += ["--lua-filter", eqno_lua]
            if extra:
                # 粗略按空白切分，支持引号包裹的参数
                cmd += self._split_args(extra)

            jobs.append({"idx": i, "src": src, "dst": dst, "cmd": cmd,
                         "overwrite": overwrite})
        return jobs

    @staticmethod
    def _split_args(text):
        import shlex
        try:
            return shlex.split(text, posix=not sys.platform.startswith("win"))
        except ValueError:
            # 回退到简单空白切分
            return text.split()

    def _worker(self, jobs):
        ok = skip = fail = 0
        outputs = []
        for job in jobs:
            idx = job["idx"]
            dst = job["dst"]
            src = job["src"]
            wx.CallAfter(self._set_status, idx, "转换中…", T.TEXT_MUTED)

            # 目标目录
            try:
                os.makedirs(os.path.dirname(dst), exist_ok=True)
            except OSError as ex:
                fail += 1
                wx.CallAfter(self._set_status, idx, "失败", T.ERR)
                self.log("✗ 无法创建目录：%s (%s)" % (os.path.dirname(dst), ex), COL_STATUS_ERR)
                wx.CallAfter(self.gauge.SetValue, self.gauge.GetValue() + 1)
                continue

            # 覆盖判断
            if os.path.exists(dst) and not job["overwrite"]:
                skip += 1
                wx.CallAfter(self._set_status, idx, "已跳过", T.SKIP)
                self.log("• 已存在，跳过：%s" % os.path.basename(dst), COL_STATUS_SKIP)
                wx.CallAfter(self.gauge.SetValue, self.gauge.GetValue() + 1)
                continue

            kwargs = {}
            if sys.platform.startswith("win"):
                kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
            try:
                proc = subprocess.run(
                    job["cmd"], capture_output=True, text=True,
                    encoding="utf-8", errors="replace", **kwargs)
            except Exception as ex:
                fail += 1
                wx.CallAfter(self._set_status, idx, "失败", T.ERR)
                self.log("✗ %s：%s" % (os.path.basename(src), ex), COL_STATUS_ERR)
                wx.CallAfter(self.gauge.SetValue, self.gauge.GetValue() + 1)
                continue

            if proc.returncode == 0 and os.path.isfile(dst):
                try:
                    _normalize_equation_tables_docx(dst)
                except Exception as ex:
                    fail += 1
                    wx.CallAfter(self._set_status, idx, "失败", T.ERR)
                    self.log("✗ %s：公式版式修正失败：%s" %
                             (os.path.basename(src), ex), COL_STATUS_ERR)
                else:
                    ok += 1
                    outputs.append(dst)
                    wx.CallAfter(self._set_status, idx, "成功", T.OK)
                    self.log("✓ %s → %s" %
                             (os.path.basename(src), os.path.basename(dst)), COL_STATUS_OK)
            else:
                fail += 1
                wx.CallAfter(self._set_status, idx, "失败", T.ERR)
                err = (proc.stderr or proc.stdout or "").strip()
                self.log("✗ %s：\n%s" % (os.path.basename(src), err), COL_STATUS_ERR)

            wx.CallAfter(self.gauge.SetValue, self.gauge.GetValue() + 1)

        # 完成
        summary = "转换完成：成功 %d，跳过 %d，失败 %d。" % (ok, skip, fail)
        self.log(("—" * 30) + "\n" + summary,
                 COL_STATUS_OK if fail == 0 else COL_STATUS_ERR)
        wx.CallAfter(self._set_busy, False)
        wx.CallAfter(self.btn_open.Enable, True)

        def _enable_copy():
            self.last_outputs = outputs
            self.btn_copy.Enable(bool(outputs))
        wx.CallAfter(_enable_copy)

        def _done():
            wx.MessageBox(summary, "完成", wx.OK | wx.ICON_INFORMATION)
        wx.CallAfter(_done)

    # ------------------------------------------------------------ 打开目录 --
    @staticmethod
    def _open_path(path):
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])


# --------------------------------------------------------------------------- #
# 入口
# --------------------------------------------------------------------------- #
def main():
    app = wx.App(False)
    frame = MainFrame()
    frame.Show()
    app.MainLoop()


def _self_test_pandoc():
    """供打包验收使用：用内置 pandoc 实际生成一个临时 DOCX。"""
    pandoc = find_pandoc()
    if not pandoc:
        return 2
    kwargs = {}
    if sys.platform.startswith("win"):
        kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
    try:
        with tempfile.TemporaryDirectory(prefix="md2word_selftest_") as tmpdir:
            src = os.path.join(tmpdir, "input.md")
            dst = os.path.join(tmpdir, "output.docx")
            with open(src, "w", encoding="utf-8") as f:
                f.write("# Pandoc self test\n\n$y = ax + b$\n")
            proc = subprocess.run([pandoc, src, "-o", dst, "--standalone"],
                                  capture_output=True, timeout=30, **kwargs)
            if proc.returncode != 0 or not os.path.isfile(dst):
                return 4
            if os.path.getsize(dst) < 1000:
                return 5
    except Exception:
        return 3
    return 0


if __name__ == "__main__":
    if "--self-test-pandoc" in sys.argv:
        sys.exit(_self_test_pandoc())
    main()
