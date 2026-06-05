# Word 科技报告目录自动化工具

一个 Windows 小工具，用来给 Word 科技报告自动整理主目录、插图清单和附表清单。工具会复制原文档，在原文件旁边生成新文件，不会覆盖原文件。

## 中文说明

### 最简单用法

1. 第一次使用前，双击 `Install-Dependencies.bat` 安装依赖。
2. 双击 `Run-WordReportTool.bat`。
3. 在弹出的窗口里选择要处理的 `.docx` 文件。
4. 等待完成。

默认输出文件名类似：

```text
报告.with-directories.docx
```

如果同名文件已经存在，会自动追加时间戳。

### 它会做什么

1. 把正文里的编号章节标题套用 Word 标题样式：`Heading 1`、`Heading 2`、`Heading 3`，并兼容中文 Word 样式；`引言` 会作为一级目录项单独进入目录。
2. 删除手写主目录，插入真正的 Word 自动目录，并刷新页码。
3. 更新 `插图清单` 和 `附表清单`，给正文图题注、表题注套用 `图题注`、`表题注` 样式。
4. 对中文用户名、中文文件夹名、中文文档路径做了编码兼容处理。

### 文档格式要求

文档前面最好有这些标题：

```text
目录
插图清单
附表清单
引言
```

正文标题应类似：

```text
引言
1 第一章标题
1.1 探索任务的技术定位
1.1.1 技术动机
```

图题注应类似：

```text
图 1 论文中的 Graph.hls 工作流对比
```

表题注应类似：

```text
表 1 报告二内容来源与证据组织
```

### 如果双击运行不了

先看桌面上的日志文件：

```text
WordReportTool-run-log.txt
```

如果提示找不到 Python，请先安装 Python 3，并在安装时勾选 `Add python.exe to PATH`，然后重新双击 `Install-Dependencies.bat`。

如果提示找不到 Word COM 自动化，请确认电脑上安装的是桌面版 Microsoft Word，而不是只有网页版 Word。

也可以打开 PowerShell，进入工具目录后运行：

```powershell
cd "解压后的工具目录"
powershell -NoProfile -ExecutionPolicy Bypass -File ".\tools\Update-ReportDirectories.ps1" -Path "C:\Users\YourName\Desktop\report.docx"
```

注意：处理前请先关闭正在打开的同一个 Word 文档。

## English

This is a small Windows utility for updating Word report directories: the main table of contents, list of figures, and list of tables. It works on a copy of the selected document and does not overwrite the original file.

### Quick Start

1. Before first use, double-click `Install-Dependencies.bat`.
2. Double-click `Run-WordReportTool.bat`.
3. Select a `.docx` file in the file picker.
4. Wait for the tool to finish.

The default output file name looks like this:

```text
report.with-directories.docx
```

If that file already exists, a timestamp is appended automatically.

### What It Does

1. Applies Word heading styles to numbered body headings: `Heading 1`, `Heading 2`, and `Heading 3`, with compatibility for Chinese Word style tables; `引言` is added as a separate level-1 TOC entry.
2. Replaces a manually typed main TOC with a real Word TOC field and refreshes page numbers.
3. Updates `插图清单` and `附表清单`, and applies figure/table caption styles to matching captions.
4. Handles Chinese usernames, Chinese folder names, and Chinese document paths more reliably.

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
1.1 技术定位
1.1.1 技术动机
```

Figure captions should look like:

```text
图 1 Example figure caption
```

Table captions should look like:

```text
表 1 Example table caption
```

### Troubleshooting

Check the log file on the desktop:

```text
WordReportTool-run-log.txt
```

If Python is not found, install Python 3 and tick `Add python.exe to PATH`, then run `Install-Dependencies.bat` again.

If Word COM automation is not available, make sure the desktop version of Microsoft Word is installed.

You can also run the main script manually from PowerShell:

```powershell
cd "path\to\word-report-directory-tool"
powershell -NoProfile -ExecutionPolicy Bypass -File ".\tools\Update-ReportDirectories.ps1" -Path "C:\Users\you\Desktop\report.docx"
```

Close the target Word document before processing it.
