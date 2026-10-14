# Markdown → Word 批量转换器

一个 Windows 桌面小工具：把 Markdown 批量转成 Word（`.docx`），并且让数学公式在 Word 里保持**原生可编辑**，而不是变成图片或纯文本。

界面用 wxPython 自绘，冷灰白配色，安静不抢眼。

---

## 功能特性

- **批量转换**：添加单个文件，或整个文件夹（可递归子目录），保留相对目录结构
- **公式保真**：LaTeX 数学（`$...$`、`$$...$$`、`\(...\)`、`\[...\]`）转成 Word 原生公式（OMML），在 Word 里可以直接点开编辑
- **公式编号**：`\tag{X}` 自动排成「公式居中 + 编号右对齐」的无边框三列表格，接近正式报告版式
- **一键复制**：把转换结果作为带 MathML 的富文本写入剪贴板，粘贴到 Word / WPS 即为原生公式，无需先落盘
- **参考样式**：可指定任意 `.docx` 作为样式模板
- **输出控制**：自定义输出目录、是否覆盖已存在文件、是否生成目录（`--toc`）、追加任意 pandoc 参数

## 环境要求

| 依赖 | 说明 |
| --- | --- |
| Windows | 剪贴板富文本、打开目录等按 Windows 实现；其它平台可运行但未验证 |
| [pandoc](https://pandoc.org/installing.html) | **必需**。转换后端完全依赖它，请确保 `pandoc` 在 `PATH` 中 |
| Python 3.x + wxPython | 仅在从源码运行时需要；用打包好的 exe 则不需要 |

## 使用方式

### 直接使用打包好的 exe

运行 `dist/Markdown转Word批量转换器.exe`。注意 exe **不包含 pandoc**，目标机器仍需自行安装 pandoc。

### 从源码运行

```bash
pip install wxPython
python md2word.py
```

### 自行打包

```bash
pyinstaller --noconfirm md2word.spec
```

产物为 `dist/Markdown转Word批量转换器.exe`（单文件、无控制台）。`md2word.spec` 已排除 numpy、torch、tensorflow、PyQt、jupyter 等运行期用不到的库，产物约 17 MB。

## 界面说明

左栏是**粘贴区**，右栏是**批量转换区**。

**左栏**

- 把 Markdown 直接粘到输入框
- `清空` / `粘贴`：清空输入框 / 把剪贴板文本插入到光标处
- `转换并复制（含公式）`：转换并把结果以富文本形式复制到剪贴板

**右栏**

- `添加文件…` / `添加文件夹…`：加入待转换的 Markdown，文件夹可递归
- `转换选项`：输出目录、输入格式（留空则自动识别）、参考样式 docx、pandoc 额外参数，以及公式转换、输出位置、是否覆盖、是否生成目录等开关
- `开始转换`：批量转换，逐条显示状态与日志
- `复制内容` / `打开输出目录`：复制本次转换结果的富文本 / 打开输出目录

## 项目结构

```
md2word.py            主程序（单文件，包含界面、转换逻辑与内置 Lua 过滤器）
md2word.spec          PyInstaller 打包配置
logo/logo.png         原始 logo
logo/logo.ico         多尺寸图标（打包与窗口图标使用）
eqno_table.lua        独立的公式编号过滤器（主程序已内嵌等价实现，保留作参考）
test.md / test.docx    转换效果样例
```

## 实现说明

- **公式编号**：pandoc 的 texmath 不支持 LaTeX 的 `\tag{}`，会丢掉编号甚至让整条公式转换失败。程序内置一个 Lua 过滤器（见 `md2word.py` 中的 `LUA_EQNO`），在转换前把带 `\tag{X}` 的块级公式改写成三列表格。
- **富文本剪贴板**：把结果封装成 Windows 的 `HTML Format`（`CF_HTML`）写入剪贴板，由 Word / WPS 在粘贴时把 MathML 转成原生公式，因此不需要安装 Word 就能拿到可编辑公式。
- **输入格式**：勾选公式转换时会自动在输入格式上追加 `tex_math_dollars` / `tex_math_single_backslash`；留空则改用 `markdown` 而非 `gfm`（`gfm` 会把单独一行的 `=` 误判为 setext 标题，打断 `$$...$$`）。
- **界面**：按钮、卡片、输入框、进度条均为自绘控件，以绕开 Windows 原生控件陈旧且难以着色的限制。

## 已知限制

- 依赖外部 `pandoc`，未随程序分发
- 未在 macOS / Linux 上验证
