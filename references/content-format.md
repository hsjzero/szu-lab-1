# 内容文件格式说明

所有内容都是 JSON，放在工作目录的 `content/` 下；模板在 `assets/examples/`。
文本字段支持 `{占位符}`（见下），构建时替换为 config.json 的值。

## 通用占位符

`{name}` `{sid}` `{school}` `{college}` `{major}` `{grade}` `{teacher}`
`{teacher_title}` `{course}` `{email}` `{phone}` `{birth}` `{hometown}`
`{politics}` `{gender}` `{exp_name}` `{date_range}` `{due_date}` `{report_date}`

---

## config.json

```json
{
  "identity": {
    "name": "姓名", "sid": "学号", "name_pinyin": "ZHANG SAN",
    "gender": "男", "birth": "2008 年 01 月", "hometown": "籍贯",
    "politics": "群众", "school": "学校", "college": "学院",
    "major": "专业", "grade": "2026 级", "teacher": "指导教师",
    "teacher_title": "职称", "course": "课程名",
    "email": "someone@example.com", "phone": "138-0000-0000"
  },
  "assignment": {
    "exp_name": "实验1",
    "title": "文字信息处理 + AI 辅助工具使用",
    "date_range": "2026 年 10 月 15 日 — 10 月 24 日",
    "due_date": "2026 年 10 月 25 日",
    "report_date": "2026 年 10 月 25 日",
    "ai_tool": "腾讯元宝（yuanbao.tencent.com）",
    "ai_source": "AI文字处理原稿.docx"
  },
  "thesis":    { "title": "论文题目", "keywords": "关键词一；关键词二", "author_line": "" },
  "newsletter":{ "title": "小报主题（也是文件名）", "pinyin": "XIAOBAO ZHUTI", "issue": "第 1 期", "org": "" },
  "resume":    { "intent": "求职意向" }
}
```

- `thesis.title` 中含 `——` 时，封面题目自动分两行显示；
  `thesis.author_line` 留空则自动生成"学院（专业） 姓名"。
- `newsletter.title` 同时用作小报文件名。

## content/thesis.json（毕业论文）

顶层：`title`（留空用 config）、`abstract`、`keywords`（留空用 config）、
`references` 参考文献数组。`blocks` 正文块：

| type | 字段 | 说明 |
|---|---|---|
| `h1` / `h2` / `h3` | `text` | 标题（套 Heading 样式，进自动目录） |
| `p` | `text` | 正文段 |
| `table` | `headers` `rows` `widths` `caption` `font_size` | 表格；`widths` 为各列 cm |
| `figure` | `image` `width` `caption` | 插图，`image` 为 images/ 下文件名，width 默认 11.0 cm |
| `ref` | `text` | 单条参考文献（也可全部放顶层 `references`） |

## content/newsletter.json（小报）

`masthead`：`title` 报头大标题、`pinyin` 拼音行、`issue` 期号、
`org_line` 落款行（`org_line` 里的 `{issue}` 取 masthead.issue，其余占位符同通用规则）。
`blocks` 为块列表：

| type | 字段 | 说明 |
|---|---|---|
| `section` | `title` `sub` | 栏目标题，`sub` 为灰色小字副标题 |
| `p` | `text` `drop_cap` | 正文段；`drop_cap: true` 首字下沉 |
| `formula` | `text` | 居中加粗的公式/化学式行 |
| `figure` | `image` `caption` `width` | 插图，width 默认 7.8 cm |
| `table` | `headers` `rows` `caption` | 表格 |
| `list` | `items` | ◆ 列表 |
| `editorial` | `text` `sign` | 编后语（灰底块），`sign` 为署名行 |

## content/resume.json（个人简历）

| 字段 | 说明 |
|---|---|
| `intent` | 求职意向（留空用 config.resume.intent） |
| `info_rows` | 基本信息，4 列一组 `["标签","值","标签","值"]`；留空自动用身份信息 |
| `contact` | `{"phone": "...", "email": "..."}`，留空用身份信息 |
| `education` / `skills` / `campus` / `awards` | 两列 `[左列, 右列]`，右列支持 `\n` 换行 |
| `self_eval` | 自我评价正文 |
| `footer_note` | 页脚提示语 |

## content/ai_tasks.json（AI 文字处理结果）

| 字段 | 说明 |
|---|---|
| `tool` | AI 工具名（留空用 config.assignment.ai_tool） |
| `intro` | 实验说明段 |
| `tasks` | 任务数组，每项：`name`、`original` 原稿、`problems` 原稿问题、`rounds` 轮次数组、`compare` 对比分析 |
| `tasks[].rounds[]` | `title` 小标题、`prompt` 提示词、`output` 输出、`note` 结果分析 |
| `summary_table` | `headers` `rows` `widths` `caption` |
| `conclusions` | 小结段落数组 |
| `disclaimer` | 文末说明 |

## content/report_steps.json（实验报告）

| 字段 | 说明 |
|---|---|
| `goals` | 实验目的列表（留空用内置默认） |
| `tasks_overview` | 实验内容列表（留空用内置默认） |
| `topics` | 选题说明段（一句话交代四份文档各自选了什么题） |
| `intro_steps` | 第三节引言（留空用内置默认） |
| `steps` | 步骤数组，每项：`title`、`problem`、`method`、`result`、`shots` |
| `steps[].shots[]` | `image` 文件名、`caption` 图题、`width`（默认 11.2 cm） |
| `summary_table` | `headers` `rows` `caption` |
| `conclusions` | 体会段落数组 |
| `thinking` | 思考题 `[{"q": "题目", "a": ["答案段", ...]}]` |

注意：`problem` / `method` / `result` **不要**自己写"问题：""操作方法：""结果："前缀，
构建器会自动加。

---

## 截图计划 shots_plan.json（word_shots.py）

```json
{
  "out_dir": "shots",
  "docs": [
    {
      "file": "out/张三-2026000001-毕业论文.docx",
      "update_toc": true,
      "shots": [
        {"name": "s1_toc.png",    "find": "目  录"},
        {"name": "s2_header.png", "find": "家访是", "view": "header"},
        {"name": "s3_footer.png", "find": "家访是", "view": "footer"},
        {"name": "s4_cover.png",  "find": "本 科 毕 业 论 文", "show_all": true}
      ]
    },
    {
      "file": "out/小报主题.docx",
      "shots": [{"name": "s5_news.png", "find": "小报主题"}]
    },
    {
      "file": "out/张三-2026000001-个人简历.docx",
      "shots": [{"name": "s6_resume.png", "find": "基本信息", "table_gridlines": true}]
    }
  ]
}
```

| 字段 | 说明 |
|---|---|
| `out_dir` | 截图输出目录（默认 `shots`） |
| `docs[].file` | 要打开的 docx 路径（相对计划文件所在目录） |
| `docs[].update_toc` | 打开后更新所有域和目录 |
| `shots[].name` | 输出文件名 |
| `shots[].find` | 查找并滚动定位到的文本 |
| `shots[].view` | `main`（默认）/ `header` / `footer` |
| `shots[].show_all` | 显示编辑标记 |
| `shots[].table_gridlines` | 显示表格网格线 |

截图后用 `crop_shots.py` 裁掉桌面边缘（默认 box 适配 2560×1600 全屏，
其他分辨率先目视校准一次再批量），裁好的图放入 `shots_crop/`，
再在 `content/report_steps.json` 里引用并加 `--shots shots_crop --only report` 重建报告。
每张截图后人工核对内容是否正确（定位是否成功、页眉页脚是否切对）。
