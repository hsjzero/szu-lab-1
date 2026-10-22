# -*- coding: utf-8 -*-
"""按比例裁掉全屏截图中的桌面边缘，只留 Word 纸面区域。

用法：
    python crop_shots.py <输入目录> <输出目录> [--box 0.25,0.205,0.76,0.965]

--box 为相对比例 (left, top, right, bottom)，默认适配 2560x1600 全屏截图，
其他分辨率先目视校准一次再批量。缺省处理输入目录下所有 .png。
"""
import argparse
import os
import sys

DEFAULT_BOX = (0.250, 0.205, 0.760, 0.965)


def crop_file(src, dst, box):
    from PIL import Image
    im = Image.open(src).convert("RGB")
    w, h = im.size
    b = (int(w * box[0]), int(h * box[1]), int(w * box[2]), int(h * box[3]))
    out = im.crop(b)
    out.save(dst, "PNG")
    return out.size


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("--box", default=",".join(str(x) for x in DEFAULT_BOX))
    args = ap.parse_args(argv)
    box = tuple(float(x) for x in args.box.split(","))
    if len(box) != 4:
        print("--box 需要 4 个值：left,top,right,bottom")
        return 2
    os.makedirs(args.dst, exist_ok=True)
    names = [f for f in sorted(os.listdir(args.src)) if f.lower().endswith(".png")]
    if not names:
        print("输入目录没有 png：%s" % args.src)
        return 1
    for name in names:
        src = os.path.join(args.src, name)
        dst = os.path.join(args.dst, name)
        size = crop_file(src, dst, box)
        print("%-20s -> %dx%d" % (name, size[0], size[1]))
    print("done, %d file(s)" % len(names))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
