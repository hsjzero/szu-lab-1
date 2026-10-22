# -*- coding: utf-8 -*-
"""驱动本机 Word 做真实操作并截图（截图填充实验报告用）。

用法：
    python word_shots.py <plan.json> [--out shots]

plan.json 格式：
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

字段：find=查找定位文本；view=main|header|footer；show_all=显示编辑标记；
table_gridlines=显示表格网格线；update_toc=打开后更新域与目录。
依赖：pywin32、pillow；仅 Windows 有效。
"""
import json
import os
import sys
import time
import traceback

WD_SEEK_MAIN, WD_SEEK_HEADER, WD_SEEK_FOOTER = 0, 9, 10
VIEW_MAP = {"main": WD_SEEK_MAIN, "header": WD_SEEK_HEADER,
            "footer": WD_SEEK_FOOTER}


def log(lines, msg):
    lines.append(msg)
    print(msg, flush=True)


def find_word_hwnd():
    """枚举可见的 OpusApp 窗口（FindWindow 会拿到隐藏窗口）"""
    import win32gui
    found = []

    def cb(h, _):
        if win32gui.GetClassName(h) == "OpusApp" and win32gui.IsWindowVisible(h):
            if win32gui.GetWindowTextLength(h) > 0:
                found.append(h)
        return True

    win32gui.EnumWindows(cb, None)
    return found[0] if found else 0


def bring_front():
    import win32gui
    import win32con
    h = find_word_hwnd()
    if not h:
        return 0
    win32gui.ShowWindow(h, win32con.SW_MAXIMIZE)
    try:
        win32gui.SetWindowPos(h, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                              win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        win32gui.SetForegroundWindow(h)
    except Exception:
        pass
    return h


def drop_topmost():
    """截图后撤掉置顶，避免挡住其他窗口"""
    import win32gui
    import win32con
    h = find_word_hwnd()
    if h:
        try:
            win32gui.SetWindowPos(h, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                                  win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
        except Exception:
            pass


def find_text(app, needle):
    """用查找定位，保证视图滚动到目标处"""
    rng = app.ActiveDocument.Content
    f = rng.Find
    f.ClearFormatting()
    f.Text = needle
    f.Forward = True
    f.Wrap = 0  # wdFindStop
    ok = f.Execute()
    if ok:
        rng.Select()
        app.Selection.Collapse(1)  # 折叠到开头
    return bool(ok)


def run(plan_path, out_dir=None):
    lines = []
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        pass
    from PIL import ImageGrab
    import win32com.client as wc

    with open(plan_path, encoding="utf-8") as f:
        plan = json.load(f)
    out_dir = out_dir or plan.get("out_dir", "shots")
    os.makedirs(out_dir, exist_ok=True)

    app = wc.Dispatch("Word.Application")
    app.Visible = True
    app.DisplayAlerts = 0
    time.sleep(3)
    log(lines, "word %s docs=%d" % (app.Version, app.Documents.Count))
    while app.Documents.Count > 0:
        app.Documents.Item(1).Close(False)

    try:
        for entry in plan.get("docs", []):
            path = entry["file"]
            if not os.path.exists(path):
                log(lines, "  [warn] 文件不存在，跳过：%s" % path)
                continue
            doc = app.Documents.Open(os.path.abspath(path))
            time.sleep(3)
            doc.Activate()
            app.WindowState = 1
            log(lines, "%s pages=%d sections=%d" %
                (os.path.basename(path), doc.ComputeStatistics(2), doc.Sections.Count))
            if entry.get("update_toc"):
                try:
                    doc.Fields.Update()
                    doc.TablesOfContents.Item(1).Update()
                    time.sleep(2)
                    log(lines, "  toc updated, paragraphs=%d" %
                        doc.TablesOfContents.Item(1).Range.Paragraphs.Count)
                except Exception as e:
                    log(lines, "  toc upd: %s" % e)

            for sh in entry.get("shots", []):
                # 复位视图状态
                try:
                    app.ActiveWindow.ActivePane.View.SeekView = WD_SEEK_MAIN
                    app.ActiveWindow.View.ShowAll = False
                    app.ActiveWindow.View.TableGridlines = False
                except Exception:
                    pass
                time.sleep(0.4)
                if sh.get("find"):
                    ok = find_text(app, sh["find"])
                    log(lines, "  find %-14s -> %s" % (sh["find"][:14], ok))
                    time.sleep(0.6)
                if sh.get("show_all"):
                    app.ActiveWindow.View.ShowAll = True
                if sh.get("table_gridlines"):
                    app.ActiveWindow.View.TableGridlines = True
                view = sh.get("view", "main")
                if view != "main":
                    app.ActiveWindow.ActivePane.View.SeekView = VIEW_MAP[view]
                    time.sleep(1.2)
                bring_front()
                time.sleep(0.8)
                img = ImageGrab.grab()
                name = sh.get("name") or ("shot_%d.png" % int(time.time()))
                outp = os.path.join(out_dir, name)
                img.save(outp)
                log(lines, "  shot %-18s %dx%d %d bytes" %
                    (name, img.width, img.height, os.path.getsize(outp)))
                # 复位
                try:
                    app.ActiveWindow.ActivePane.View.SeekView = WD_SEEK_MAIN
                    app.ActiveWindow.View.ShowAll = False
                    app.ActiveWindow.View.TableGridlines = False
                except Exception:
                    pass

            doc.Save()
            doc.Close(False)
            time.sleep(1.5)

        app.Quit()
        log(lines, "ALL DONE")
    except Exception:
        lines.append(traceback.format_exc())
        try:
            app.Quit()
        except Exception:
            pass
    finally:
        drop_topmost()
        with open("word_shots.log", "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
    return out_dir


def main(argv):
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("plan")
    ap.add_argument("--out", default=None)
    args = ap.parse_args(argv)
    run(args.plan, args.out)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
