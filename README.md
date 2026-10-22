# # szu-lab-1
大学计算机／文字信息处理类 Word 实验作业的一键交付流水线（面向深圳大学《大学计算机》实验一类作业，同样适用于任何"用 Word 交 5 份文档"的课程作业）。


输入一份 `config.json`（个人信息）和几份内容 JSON（正文素材），即可产出完整的 5 份交付文档：

| 交付物 | 内容 |
|---|---|
| 实验报告 | 封面 + 实验目的 + 实验内容 + 实验步骤与过程（可嵌入真实 Word 截图）+ 体会 + 思考题 |
| 个人简历 | 无框线表格排版的两页简历 |
| Word 小报 | 通栏报头 + 双栏正文 + 页面边框 + 插图 + 表格 |
| 毕业论文 | 封面 + 诚信声明 + 自动目录 + 摘要关键词 + 正文 + 参考文献，正文节独立页眉页脚 |
| AI 文字处理结果 | 润色／扩写／简写的提示词与输出对比记录 |

脚本只含排版逻辑，**不含任何个人数据**；所有姓名、学号、邮箱、正文都来自用户自己的配置。

## 环境要求

- Python 3.9+
- `pip install python-docx pillow`
- Windows + Microsoft Word / WPS（真实截图与页数统计另需 `pip install pywin32`）

## 快速开始

```bash
# 1. 准备工作目录
mkdir work && cd work
cp ../word-lab-report/assets/config.example.json config.json
cp -r ../word-lab-report/assets/examples content

# 2. 编辑 config.json 填入自己的信息；按 references/content-format.md 改写 content/ 下的内容
#    插图（小报/论文用）放入 work/images/

# 3. 一键构建
python ../word-lab-report/scripts/build_all.py \
    --config config.json --content content --out out --images images

# 4. 自检（节数/标题样式/域/占位符残留/兼容性模式/页数）
python ../word-lab-report/scripts/check_docx.py out
```

产物在 `out/` 目录。只重建某一类文档：`--only thesis,newsletter,resume,ai,report`。

## （可选）真实 Word 截图填充实验报告

```bash
# 1. 编写截图计划 shots_plan.json（格式见 references/content-format.md）
# 2. 驱动 Word 真实操作并截图
python ../word-lab-report/scripts/word_shots.py shots_plan.json --out shots
# 3. 裁剪掉桌面边缘只留 Word 纸面区域（box 为相对比例，先目视校准）
python ../word-lab-report/scripts/crop_shots.py shots shots_crop --box 0.25,0.205,0.76,0.965
# 4. 在 content/report_steps.json 的 steps 中引用截图，只重建报告
python ../word-lab-report/scripts/build_all.py \
    --config config.json --content content --out out --images images \
    --shots shots_crop --only report
```

各内容文件的字段格式见 [references/content-format.md](references/content-format.md)，
已知坑与技巧见 [references/pitfalls.md](references/pitfalls.md)。

## 作为 Agent Skill 使用

本仓库同时是一个 WorkBuddy/Claude 风格的 skill（入口为 `SKILL.md`）。
放入 skills 目录后，对助手说"帮我完成文字信息处理实验作业"即可触发完整流程。

## 目录结构

```
word-lab-report/
├── SKILL.md                    # skill 入口（触发条件与工作流程）
├── README.md
├── scripts/
│   ├── common.py               # python-docx 排版辅助（字体/域/边框/分栏等）
│   ├── config.py               # 配置加载与 {占位符} 替换
│   ├── build_thesis.py         # 毕业论文构建器
│   ├── build_newsletter.py     # 小报构建器
│   ├── build_resume.py         # 个人简历构建器
│   ├── build_ai.py             # AI 文字处理结果构建器
│   ├── build_report.py         # 实验报告构建器
│   ├── build_all.py            # 一键构建入口
│   ├── word_shots.py           # 驱动 Word 真实操作并截图（pywin32）
│   ├── crop_shots.py           # 按比例裁剪截图
│   ├── check_docx.py           # 交付前自检（页数/域/样式/占位符/兼容模式）
│   └── fix_compat.py           # 修复"兼容性模式"
├── references/
│   ├── content-format.md       # config 与 content JSON 的字段格式
│   └── pitfalls.md             # python-docx / Word 排版已知坑
└── assets/
    ├── config.example.json     # 配置模板（全部为占位符）
    └── examples/               # 五份文档的示例内容 JSON
```

## 隐私说明

`assets/` 与 `references/` 中只包含占位符示例（张三、示例学号、someone@example.com）。
真实信息只写在你自己工作目录的 `config.json` 与 `content/` 中，不要提交到本仓库。

## License

MIT
