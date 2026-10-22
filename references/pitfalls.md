# 已知坑与技巧（pitfalls）

都是实战踩过的，按遇到频率排序。生成异常时先查这里。

## 1. compatibilityMode=14 → Word 显示“兼容性模式”

python-docx 默认模板的 `word/settings.xml` 里 `compatibilityMode` 是 14（Word 2010）。
**处理**：`build_all.py` 会自动对每份文档调用 `fix_compat`；单独修复用
`python scripts/fix_compat.py out/`。

## 2. zip 中文文件名乱码

外层 zip（作业包）多为 UTF-8，内层素材 zip 常为 GBK：

```python
import zipfile
try:
    zf = zipfile.ZipFile(f, metadata_encoding="gbk")
    _ = zf.namelist()
except Exception:
    zf = zipfile.ZipFile(f)  # 或 metadata_encoding="utf-8"
```

解出来目录名仍是乱码（如 `瀹為獙1`）说明编码猜错了。

## 3. 多节文档必须“顺序写”

`add_section()` 只是插入分节符，光标永远在文末——先建全部节再回头填内容，
内容会全部堆到文末。正确做法：写一节 → `add_section()` → 立刻写下一节。

## 4. 自动目录为空

TOC 域没问题，是标题没套样式：`doc.add_paragraph(style='Heading 1')`。
手动加粗放大字号的标题不进目录。“目 录”两个字本身不要用 Heading 样式。

## 5. 页眉页脚串页 / 封面也有页眉

页眉页脚按“节”继承：`section.header.is_linked_to_previous = False` 断开；
封面节要空白页眉就断开后清空段落，不是只删文字。

## 6. 正文页码从 1 重排

`restart_page_numbering(section, 1)`（写 `w:pgNumType w:start="1"`）。
论文封面/目录不编页码、正文从 1 开始是标配；单节文档用 `NUMPAGES`，
多节各自计数用 `SECTIONPAGES`。

## 7. 分栏报头被关进去 / 表格不能通栏

报头通栏：第 1 节 1 栏 → `add_section(WD_SECTION.CONTINUOUS)` → 第 2 节 `set_columns(2)`。
**表格在分栏节里不会自动通栏**，需要通栏表格就再切一个 CONTINUOUS 节。
分栏改 `w:cols w:num`，不要用分栏符。

## 8. 页面边框不显示

`w:pgBorders` 必须插在 `w:pgSz/w:pgMar` 之后、`w:cols/w:docGrid` 之前，
直接 `append` 会因元素顺序错误被 Word 忽略。`common.add_page_border()` 已用
`insert_element_before` 处理。

## 9. 中文字体只设 font.name 无效

必须写 `rPr/rFonts` 的 `w:eastAsia`。`common.set_font()` 一次设
`ascii/hAnsi/eastAsia` 三属性；`new_doc()` 已处理好 Normal 样式。

## 10. Word COM 截图的坑

- **PrintWindow 对 Word 返回空白**：Word 硬件渲染，用 `PIL.ImageGrab.grab()` 全屏抓。
- **FindWindow("OpusApp") 拿到隐藏窗口**：枚举 `EnumWindows` + `IsWindowVisible` + 标题非空。
- **宿主窗口遮挡 Word**：截图前 `SetWindowPos(HWND_TOPMOST)`，截图后撤掉
  （`word_shots.py` 的 `drop_topmost()`）。
- **深色主题下截图发灰**：先把 Word 切回浅色主题。
- **GoTo 翻页不可靠**：用 `Range.Find` 搜关键词 → `Select()` → `Collapse(1)`。
- **HiDPI 下抓屏尺寸错**：`ctypes.windll.shcore.SetProcessDpiAwareness(2)`。
- **页眉页脚视图**：`ActivePane.View.SeekView = 9/10/0`（页眉/页脚/主文档）。
- 更新域：`doc.Fields.Update()` + `doc.TablesOfContents.Item(1).Update()`；
  用户端打开还要右键“更新域”，所以 TOC 域的 cached 文本要有兜底说明。

## 11. python-docx 不能验证渲染结果

页数、目录、分页对不对，必须借本机 Word：`ComputeStatistics(2)` 数页，或
`ExportAsFixedFormat` 导 PDF 再数页。`check_docx.py` 已封装。

## 12. 杂项

- `~$` 开头的 Word 锁文件：批量处理跳过（脚本已处理）。
- `doc.add_table()` 默认样式随模板，务必显式 `set_table_borders` / `remove_table_borders`。
- 单元格宽度逐格设 `cell.width`，只设列宽会被 Word 重算。
- 表格跨页重复表头：`set_repeat_header(row)`。
- 图片缺失不中断构建，只打印警告并跳过该图（报告里则显示红色占位提示）。
- 字段值以 `content/*.json` 为准，脚本内不硬编码任何个人信息。

## 13. 交作业前的最后一遍

- 每份 docx 用 Word 打开，右键更新域，看目录/页码/共 Y 页；
- 报告截图 ≥3 张且为真实 Word 窗口（有标题栏、功能区）；
- 通读一遍语气：第一人称、有具体卡壳点、无排比空话（见 SKILL.md 的写作风格要求）。
