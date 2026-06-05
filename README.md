# Word Report Directory Tool

一个 Windows 小工具，用来自动整理 Word 科技报告中的主目录、插图清单和附表清单。工具会复制原文档，在原文件旁边生成新文件，不会覆盖原文件。

## 中文说明

### 适用场景

适合已经写好正文、但需要统一生成这些内容的 Word 科技报告：

- 主目录
- 插图清单
- 附表清单
- 正文标题样式
- 引言的无编号目录项

### 运行环境

- Windows
- 桌面版 Microsoft Word
- Python 3
- 第一次安装依赖时需要联网

如果还没安装 Python，请从 <https://www.python.org/downloads/> 安装，并勾选 `Add python.exe to PATH`。

### 快速使用

1. 下载或克隆本仓库。
2. 第一次使用前，双击 `Install-Dependencies.bat`。
3. 双击 `Run-WordReportTool.bat`。
4. 在弹出的窗口里选择要处理的 `.docx` 文件。
5. 等待完成。

输出文件会生成在原文档旁边，文件名类似：

```text
report.with-directories.docx
```

如果同名文件已经存在，工具会自动追加时间戳。原文档不会被覆盖。

### 文档格式要求

文档前面最好有这些独立段落：

```text
目录
插图清单
附表清单
引言
```

正文标题建议写成：

```text
引言
1 第一章标题
1.1 二级标题
1.1.1 三级标题
```

工具也兼容这种旧写法：

```text
1 引言
1.1 引言内部小标题
1.2 引言内部小标题
2 第一章标题
```

它会自动处理成：

```text
引言
1 第一章标题
```

其中 `1.1 / 1.2` 这类引言内部小标题不会进入主目录，后续章节编号会自动前移。

图题注和表题注应类似：

```text
图 1 示例图题注
表 1 示例表题注
```

### 自定义字体和格式

可以打开 `config\format-settings.json` 修改字体、字号、加粗和行距。保存后重新运行工具即可生效。

常用字段：

```text
front_title       目录/插图清单/附表清单标题
toc_1/toc_2/toc_3 主目录一级/二级/三级条目
heading_1/2/3     正文一级/二级/三级标题
intro             引言标题
intro_subheading  引言内部小标题，不进入目录
caption           正文图题注、表题注
list_entry        插图清单、附表清单条目
line_spacing      默认行距
```

每个格式项通常长这样：

```json
{
  "font": "宋体",
  "size": 12,
  "bold": null
}
```

说明：

- `font` 是字体名，例如 `宋体`、`黑体`、`仿宋`、`Times New Roman`。
- `size` 是字号，单位为 pt。
- `bold` 可填 `true`、`false` 或 `null`。`null` 表示不强制改加粗状态。

### PowerShell 手动运行

也可以打开 PowerShell，进入工具目录后运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\tools\Update-ReportDirectories.ps1" -Path "C:\path\to\report.docx"
```

使用自定义配置文件：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\tools\Update-ReportDirectories.ps1" -Path "C:\path\to\report.docx" -ConfigPath ".\config\format-settings.json"
```

### 常见问题

处理前请先关闭正在打开的同一个 Word 文档。

如果双击运行失败，先看桌面上的日志文件：

```text
WordReportTool-run-log.txt
```

如果提示找不到 Python，请安装 Python 3，并确认安装时勾选了 `Add python.exe to PATH`，然后重新运行 `Install-Dependencies.bat`。

如果提示找不到 Word COM 自动化，请确认电脑上安装的是桌面版 Microsoft Word，而不是网页版 Word。

## English

This is a small Windows utility for updating Word report directories: the main table of contents, list of figures, and list of tables. It works on a copy of the selected document and does not overwrite the original file.

### Requirements

- Windows
- Desktop Microsoft Word
- Python 3
- Internet access for first-time dependency installation

### Quick Start

1. Download or clone this repository.
2. Before first use, double-click `Install-Dependencies.bat`.
3. Double-click `Run-WordReportTool.bat`.
4. Select a `.docx` file in the file picker.
5. Wait for the tool to finish.

The output file is written next to the original document:

```text
report.with-directories.docx
```

The original document is not overwritten.

### Expected Document Format

The front matter should preferably contain:

```text
目录
插图清单
附表清单
引言
```

Body headings should look like:

```text
引言
1 Chapter title
1.1 Section title
1.1.1 Subsection title
```

If the source document uses `1 引言`, `1.1 introduction subheading`, and then `2 Chapter title`, the tool converts the introduction to an unnumbered TOC entry, excludes introduction subheadings from the TOC, and shifts later chapter numbers down by one.

Figure and table captions should look like:

```text
图 1 Example figure caption
表 1 Example table caption
```

### Custom Formatting

Edit `config\format-settings.json` to customize fonts, font sizes, bold settings, and line spacing. The tool reads this file automatically when it exists.

Common keys:

```text
front_title       TOC/List of Figures/List of Tables titles
toc_1/toc_2/toc_3 Main TOC level 1/2/3 entries
heading_1/2/3     Body heading level 1/2/3
intro             Introduction heading
intro_subheading  Introduction subheadings, excluded from the TOC
caption           Figure/table captions in the body
list_entry        List of figures/tables entries
line_spacing      Default line spacing
```

### Troubleshooting

Close the target Word document before processing it.

If the tool fails, check the desktop log file:

```text
WordReportTool-run-log.txt
```

If Python is not found, install Python 3 and tick `Add python.exe to PATH`, then run `Install-Dependencies.bat` again.

If Word COM automation is not available, make sure the desktop version of Microsoft Word is installed.
