# 人工智能导论实验指导书：Markdown 批量转换为 Word（Pandoc）

本项目包含三份实验的 Markdown 文件（`实验一.md`、`实验二.md`、`实验三.md`）。本文档介绍如何在 Windows + PowerShell 环境下使用 Pandoc 批量转换为 Word（`.docx`），并统一中文学术规范样式与多级自动编号。

## 环境要求
- Windows 10/11，PowerShell
- Pandoc（建议 2.x 及以上）
- Microsoft Word（用于参考模板样式与导出效果预览）

## 安装 Pandoc（PowerShell）
```powershell
winget install --id JohnMacFarlane.Pandoc -e
# 验证
pandoc -v
```
如无 `winget`，可前往 Pandoc 官网下载安装包。

## 目录结构（摘）
```
人工智能导论实验指导书/
  1.md
  2.md
  实验一.md
  实验二.md
  实验三.md
  1.assets/
  2.assets/
  ...
```
图片与 Markdown 位于同一根目录，Pandoc 会就近检索并内嵌。

## 生成中文学术规范模板（推荐）
为确保导出的 `.docx` 在字体、行距、标题、代码、表格、参考文献等方面符合规范，建议先生成参考模板 `reference-zh-academic.docx`。

### 方法一：使用已提供的 VBA 宏（一键生成模板）
文件：`BuildZhAcademicTemplateAndNumbering.bas`
- 打开 Word → 新建空白文档
- 按 Alt+F11 打开 VBA 编辑器 → File → Import File...
- 选择 `BuildZhAcademicTemplateAndNumbering.bas` 导入
- 按 Alt+F8 运行 `BuildZhAcademicTemplateAndNumbering`
- 在保存对话框中保存为项目根目录：`reference-zh-academic.docx`

宏会：
- 设置页面（A4、边距）、正文（宋体小四、1.5 倍行距、首行缩进）、标题 1/2/3（黑体）
- 配置代码块、图表题、列表、参考文献等样式
- 为“标题1/2/3”绑定多级自动编号（1、1.1、1.1.1）

提示：若被宏安全策略拦截 → 文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置 → “禁用所有宏并通知”，重新打开并允许运行。

### 方法二：使用 Pandoc 默认模板并手动调整（备选）
1) 导出 Pandoc 默认参考模板：
```powershell
cd 'D:\QSync\work\教学\人工智能导论实验指导书'
cmd /c "pandoc --print-default-data-file=reference.docx > reference-default.docx"
```
2) 打开 `reference-default.docx`，在 Word 中手动设置以下样式（建议）：
- 正文：宋体，小四（12pt），1.5 倍行距，首行缩进 2 字符
- 标题1/2/3：黑体（三号/小三/四号），加粗，单倍，段前 12pt 段后 6pt，与下段同页
- 代码：Consolas 10pt，单倍，左缩进 0.74 cm，前后 6pt，浅灰底 + 左边框
- 图表题（Caption）：宋体，五号 10.5pt，居中，单倍，前后 6pt
- 列表（List Paragraph）：左缩进 0.74 cm，1.5 倍
- 参考文献（Bibliography）：宋体，小四，单倍，悬挂缩进 2 字符，条目间 6pt
- 页面：A4；上/下 2.5 cm，左 3.0 cm，右 2.5 cm；页脚居中页码
3) 另存为 `reference-zh-academic.docx`。

提示：若需“1/1.1/1.1.1”自动编号，可在 Word 的“多级列表”中将编号级别分别链接到样式“标题1/2/3”。

## 批量转换命令
- 使用模板（推荐）：
```powershell
cd 'D:\QSync\work\教学\人工智能导论实验指导书'
Get-ChildItem -Filter '实验*.md' | ForEach-Object {
  pandoc $_.FullName -o ($_.BaseName + '.docx') --standalone --reference-doc '.\reference-zh-academic.docx' --resource-path=.
}
```

- 自动检测模板（模板缺失时退化为默认样式）：
```powershell
cd 'D:\QSync\work\教学\人工智能导论实验指导书'
$ref = Join-Path (Get-Location) 'reference-zh-academic.docx'
$useRef = Test-Path $ref
Get-ChildItem -Filter '实验*.md' | ForEach-Object {
  if ($useRef) {
    pandoc $_.FullName -o ($_.BaseName + '.docx') --standalone --reference-doc $ref --resource-path=.
  } else {
    pandoc $_.FullName -o ($_.BaseName + '.docx') --standalone --resource-path=.
  }
}
```
说明：`--standalone` 生成完整 Word 文档；`--reference-doc` 指定参考模板；`--resource-path=.` 在当前目录搜索图片与资源。

## 公式、代码与表格
- 公式：Markdown/LaTeX 公式会转换为 Office Math（Word 原生公式）
- 代码：高亮与样式由参考模板中“Code / Code Block”控制
- 表格：尽量使用标准 Markdown 表格；复杂表格可导出后在 Word 微调

## 常见问题（FAQ）
- 默认模板导出报“文件被占用”：关闭同名文件，改用新文件名重试；或前置 `cmd /c`
- PowerShell 出现 PSReadLine 异常：不影响 Pandoc 执行，可在新窗口重试
- 模板应用后标题未自动编号：确认模板中“多级列表”已绑定“标题1/2/3”
- 中文字体不一致：在模板中统一正文、标题、Caption、Bibliography 的字体（如宋体/黑体）
- 图片未显示：在项目根执行命令、保留 `--resource-path=.`，并检查图片相对路径

## 参考
- Pandoc 文档（DOCX 参考模板）：`https://pandoc.org/MANUAL.html#options-affecting-specific-writers`

完成以上步骤后，即可得到统一样式、自动编号的 Word 版实验文档：`实验一.docx / 实验二.docx / 实验三.docx`。