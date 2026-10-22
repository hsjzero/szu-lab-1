---
name: word-lab-report
description: 大学计算机／文字信息处理类 Word 实验作业流水线。当用户需要完成"文字信息处理 + AI 辅助工具使用"实验，一次性提交实验报告、个人简历、Word 小报、毕业论文、AI 文字处理结果等多份 docx 交付物，或需要批量生成带封面／自动目录／分节页眉页脚／分栏小报／页面边框的规范 Word 文档时使用。用户只需提供个人信息与内容素材，即可由脚本产出全部交付文档。
agent_created: true
---

# word-lab-report：Word 实验作业交付流水线

## 用途

把"大学计算机／文字信息处理"类实验作业的整套 Word 交付物（实验报告、个人简历、
Word 小报、毕业论文、AI 文字处理结果，共 5 份）用参数化脚本一次产出。
所有个人信息与正文内容都来自用户提供的 JSON 配置，脚本本身不含任何个人数据，
可公开发布。

## 环境依赖

```bash
pip install python-docx pillow
# Windows 上若要做真实 Word 截图，另装：
pip install pywin32
```

## 工作流程

严格按以下顺序执行：

### 1. 收集用户信息（必做）

向用户收集并确认以下信息，缺项先用占位符、最后提醒替换：

- 身份：姓名、学号、姓名拼音（可选）、性别、出生年月、籍贯、政治面貌
- 学校：学校名、学院、专业、年级、指导教师及职称、课程名
- 联系：手机、邮箱
- 日期：实验起止日期、提交日期
- 选题：毕业论文题目与主题、小报主题、简历内容（教育背景／技能／校园经历／获奖／自我评价）
- AI 任务：使用哪个 AI 工具、待处理原稿、各自提示词与输出结果
- 图片：小报和论文需要的插图（AI 生成或用户提供），放入工作目录 `images/`

信息齐了再动手，不要用猜测值直接生成正式交付物。

### 2. 建立工作目录并写入配置

```bash
mkdir work && cd work
cp <skill目录>/assets/config.example.json config.json
cp -r <skill目录>/assets/examples content
# 插图放入 work/images/
```

按用户真实信息编辑 `config.json`（身份／作业／论文／小报／简历各部分）。
再从 `<skill目录>/assets/examples/` 复制示例内容到 `content/`，
按 `references/content-format.md` 的格式说明改写为用户的实际内容。
**config.json 与 content/ 放在工作目录，不要写回 skill 目录**——skill 目录必须保持无个人隐私状态。

### 3. 一键构建

```bash
python <skill目录>/scripts/build_all.py --config config.json --content content --out out --images images
```

产出 `out/` 下 5 份 docx（构建顺序：毕业论文 → 小报 → 简历 → AI → 报告，
最后自动修复 compatibilityMode）：

| 文件 | 说明 |
|---|---|
| `姓名-学号-实验N.docx` | 实验报告（含目的／内容／步骤截图／体会／思考题） |
| `姓名-学号-个人简历.docx` | 两页简历，无框线表格排版 |
| `小报主题.docx` | 双栏小报，通栏报头 + 页面边框 + 图 + 表 |
| `姓名-学号-毕业论文.docx` | 封面＋诚信声明＋自动目录＋摘要＋正文＋参考文献，正文节独立页眉页脚 |
| `姓名-学号-AI文字处理结果.docx` | 润色／扩写／简写提示词对比记录 |

只重建某一类：`--only thesis,newsletter,resume,ai,report`（逗号分隔）。

### 4. （可选但推荐）真实 Word 截图后重建报告

实验报告要求"关键步骤截图至少 3 项"时，用真实操作截图而不是网图：

1. 先完成第 3 步，得到 docx；
2. 编写截图计划 JSON（格式见 `references/content-format.md` 的"截图计划"一节），
   例如：打开毕业论文 → 定位到正文 → 切到页眉视图 → 截图；
3. 运行（plan 为位置参数）：

```bash
python <skill目录>/scripts/word_shots.py shots_plan.json --out shots
```

   日志写在 `word_shots.log`；
4. 裁剪掉桌面边缘只留 Word 纸面区域：

```bash
python <skill目录>/scripts/crop_shots.py shots shots_crop --box 0.25,0.205,0.76,0.965
```

   `--box` 为相对比例 (左,上,右,下)，不同分辨率先目视校准一次再批量；
5. 在 `content/report_steps.json` 的 steps 里引用截图文件名，只重建报告：

```bash
python <skill目录>/scripts/build_all.py --config config.json --content content --out out --images images --shots shots_crop --only report
```

### 5. 自检

```bash
python <skill目录>/scripts/check_docx.py out
```

检查每份文档的节数、标题样式、表格/图片数、域、占位符残留、兼容性模式与页数
（页数需本机 Word）。再人工确认：

- 用 Word 或 WPS 打开每份文档，目录处右键"更新域"；
- 核对封面／页眉／页脚无个人信息错误、无占位符残留；
- 页数是否满足要求（小报 ≥2 页、论文正文 ≥3 页）。

## 写作风格要求（避免被识别为纯 AI 生成）

- 第一人称口语化，像刚学会的大学生：短句、具体、写清楚"卡在哪里、怎么试出来的"；
- 不写排比、不写"总而言之"、不用"赋能／抓手／深度融合"类空话；
- 步骤部分写真实操作路径（"布局"→"页面设置"→"分栏"），不写抽象描述；
- 内容不用过度精致——能交差、逻辑顺、像真人即可。

## 隐私红线

- skill 目录（含示例文件）永远只放占位符（张三、示例学号、someone@example.com 之类）；
- 真实姓名／学号／邮箱／手机只允许出现在用户工作目录的 `config.json` 与 `content/` 中；
- 交付给用户或公开前，通读一遍产物确认没有真实个人信息误入示例文件。

## 常见问题

页眉页脚不显示、目录为空、图片跑栏、中文文件名解压乱码等已知坑与解法，
见 `references/pitfalls.md`，遇错先查它。
