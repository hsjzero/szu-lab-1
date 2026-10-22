# -*- coding: utf-8 -*-
"""一键构建全部 5 份交付文档。

用法：
    python build_all.py [--config config.json] [--content content] [--out out]
                        [--images images] [--shots shots_crop] [--only a,b,c]

--only 取值：thesis, newsletter, resume, ai, report（逗号分隔）。缺省全部构建。
构建顺序固定为 毕业论文 → 小报 → 个人简历 → AI文字处理 → 实验报告，
最后统一修复 compatibilityMode。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import argparse

import config as cfgmod
import build_thesis
import build_newsletter
import build_resume
import build_ai
import build_report
from fix_compat import fix_file

ORDER = ["thesis", "newsletter", "resume", "ai", "report"]
LABELS = {"thesis": "毕业论文", "newsletter": "Word 小报", "resume": "个人简历",
          "ai": "AI 文字处理结果", "report": "实验报告"}


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default="config.json")
    ap.add_argument("--content", default="content")
    ap.add_argument("--out", default="out")
    ap.add_argument("--images", default="images")
    ap.add_argument("--shots", default="shots_crop")
    ap.add_argument("--only", default="")
    args = ap.parse_args(argv)

    cfg = cfgmod.load(args.config)
    content_dir = args.content
    out_dir = args.out

    only = [x.strip() for x in args.only.split(",") if x.strip()] or ORDER
    for k in only:
        if k not in ORDER:
            print("未知 --only 取值：%s（可用：%s）" % (k, ",".join(ORDER)))
            return 2

    os.makedirs(out_dir, exist_ok=True)
    made = []
    for kind in ORDER:
        if kind not in only:
            continue
        print("== 构建 %s ==" % LABELS[kind])
        content = cfgmod.load_json(os.path.join(content_dir, {
            "thesis": "thesis.json", "newsletter": "newsletter.json",
            "resume": "resume.json", "ai": "ai_tasks.json",
            "report": "report_steps.json"}[kind]), default={})
        if kind == "thesis":
            path = build_thesis.build(cfg, content, out_dir, args.images)
        elif kind == "newsletter":
            path = build_newsletter.build(cfg, content, out_dir, args.images)
        elif kind == "resume":
            path = build_resume.build(cfg, content, out_dir, args.images)
        elif kind == "ai":
            path = build_ai.build(cfg, content, out_dir, args.images)
        elif kind == "report":
            path = build_report.build(cfg, content, out_dir, args.shots)
        made.append(path)

    print("== 修复兼容性模式 ==")
    for path in made:
        changed = fix_file(path)
        print("  %-40s %s" % (os.path.basename(path), "fixed" if changed else "ok"))

    print("== 完成，共 %d 份 ==" % len(made))
    for path in made:
        print("  " + path)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
