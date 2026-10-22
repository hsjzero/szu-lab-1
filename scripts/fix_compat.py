# -*- coding: utf-8 -*-
"""修复 python-docx 生成文档的"兼容性模式"。

python-docx 默认模板的 compatibilityMode=14（Word 2010），用新版 Word 打开会
显示"兼容性模式"。改为 15（Word 2013+）即可消除。直接改 settings.xml 再重打包。
"""
import sys
import os
import zipfile
import shutil

OLD = ('w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"')
NEW = ('w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"')


def fix_file(path):
    """修复单个 docx；返回是否发生了替换"""
    tmp = path + ".tmp"
    changed = False
    zin = zipfile.ZipFile(path, "r")
    zout = zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED)
    try:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/settings.xml":
                text = data.decode("utf-8")
                if OLD in text:
                    text = text.replace(OLD, NEW)
                    changed = True
                data = text.encode("utf-8")
            zout.writestr(item, data)
    finally:
        zin.close()
        zout.close()
    shutil.move(tmp, path)
    return changed


def main(argv):
    if not argv:
        print("用法：python fix_compat.py <文件1.docx> [文件2.docx ...]")
        return 1
    for p in argv:
        if not os.path.exists(p):
            print("跳过（不存在）：%s" % p)
            continue
        ok = fix_file(p)
        print("%s %s" % ("已修复" if ok else "无需修复", p))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
