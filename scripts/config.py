# -*- coding: utf-8 -*-
"""加载用户配置 config.json，提供统一取值接口与 {占位符} 替换。"""
import json
import os

DEFAULTS = {
    "identity": {
        "name": "张三",
        "sid": "2026000000",
        "name_pinyin": "ZHANG SAN",
        "gender": "男",
        "birth": "2008 年 01 月",
        "hometown": "示例省示例市",
        "politics": "群众",
        "school": "示例大学",
        "college": "示例学院",
        "major": "中国语言文学类",
        "grade": "2026 级",
        "teacher": "李老师",
        "teacher_title": "讲师",
        "course": "大学计算机",
        "email": "someone@example.com",
        "phone": "138-0000-0000",
    },
    "assignment": {
        "exp_name": "实验1",
        "title": "文字信息处理 + AI 辅助工具使用",
        "date_range": "2026 年 10 月 15 日 — 10 月 24 日",
        "due_date": "2026 年 10 月 25 日",
        "report_date": "2026 年 10 月 25 日",
        "ai_tool": "腾讯元宝（yuanbao.tencent.com）",
        "ai_source": "AI文字处理原稿.docx",
    },
    "thesis": {
        "title": "论文题目（请修改）",
        "keywords": "关键词1；关键词2；关键词3",
        "author_line": "",          # 留空则自动生成
    },
    "newsletter": {
        "title": "小报主题（请修改）",
        "pinyin": "",
        "issue": "第 1 期",
        "org": "",                  # 留空则用 学校+学院
    },
    "resume": {
        "intent": "文案策划 / 新媒体编辑",
    },
}

# 支持在内容文本里使用的 {占位符}
PLACEHOLDER_KEYS = ["name", "sid", "school", "college", "major", "grade",
                    "teacher", "teacher_title", "course", "email", "phone",
                    "birth", "hometown", "politics", "gender",
                    "exp_name", "date_range", "due_date", "report_date"]


def _merge(base, override):
    out = dict(base)
    for k, v in (override or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _merge(out[k], v)
        else:
            out[k] = v
    return out


class Config(object):
    def __init__(self, data):
        self.data = data

    # ---- 取值 ----
    def get(self, dotted, default=None):
        node = self.data
        for part in dotted.split('.'):
            if not isinstance(node, dict) or part not in node:
                return default
            node = node[part]
        return node

    @property
    def identity(self):
        return self.data["identity"]

    @property
    def assignment(self):
        return self.data["assignment"]

    @property
    def thesis(self):
        return self.data["thesis"]

    @property
    def newsletter(self):
        return self.data["newsletter"]

    @property
    def resume(self):
        return self.data["resume"]

    # ---- 便捷属性 ----
    @property
    def name(self):
        return self.identity["name"]

    @property
    def sid(self):
        return self.identity["sid"]

    def fmt(self, text):
        """把内容文本中的 {name} {sid} {school} {date} 等占位符替换为配置值"""
        if not isinstance(text, str) or "{" not in text:
            return text
        values = {k: self.get("identity." + k) for k in PLACEHOLDER_KEYS
                  if self.get("identity." + k) is not None}
        values.update({k: self.get("assignment." + k) for k in PLACEHOLDER_KEYS
                       if self.get("assignment." + k) is not None})
        out = text
        for k, v in values.items():
            out = out.replace("{%s}" % k, str(v))
        return out

    # ---- 输出文件名 ----
    def out_name(self, kind):
        n, sid = self.name, self.sid
        if kind == "report":
            return "%s-%s-%s.docx" % (n, sid, self.assignment["exp_name"])
        if kind == "resume":
            return "%s-%s-个人简历.docx" % (n, sid)
        if kind == "thesis":
            return "%s-%s-毕业论文.docx" % (n, sid)
        if kind == "ai":
            return "%s-%s-AI文字处理结果.docx" % (n, sid)
        if kind == "newsletter":
            from common import safe_filename
            return safe_filename(self.newsletter["title"]) + ".docx"
        raise ValueError(kind)


def load(path):
    if not os.path.exists(path):
        raise FileNotFoundError("配置文件不存在：%s（可从 assets/config.example.json 复制）" % path)
    with open(path, encoding="utf-8") as f:
        raw = json.load(f)
    return Config(_merge(DEFAULTS, raw))


def load_json(path, default=None):
    """加载内容 JSON；文件不存在时返回 default"""
    if not os.path.exists(path):
        return default
    with open(path, encoding="utf-8") as f:
        return json.load(f)
