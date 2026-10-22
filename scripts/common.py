# -*- coding: utf-8 -*-
"""python-docx 排版辅助函数（不含任何个人数据）"""
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

SONG = "宋体"
HEI = "黑体"
KAI = "楷体"
FANG = "仿宋"


def set_font(run, cn=SONG, en="Times New Roman", size=12, bold=False,
             italic=False, color=None, underline=False):
    run.font.name = en
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    if color:
        run.font.color.rgb = RGBColor(*color)
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rpr.append(rf)
    rf.set(qn('w:ascii'), en)
    rf.set(qn('w:hAnsi'), en)
    rf.set(qn('w:eastAsia'), cn)
    return run


def para(doc, text="", cn=SONG, en="Times New Roman", size=12, bold=False,
         align=WD_ALIGN_PARAGRAPH.LEFT, first_indent=None, line=1.5,
         space_before=0, space_after=0, color=None, left_indent=None,
         right_indent=None, italic=False, keep_with_next=False):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing = line
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if first_indent is not None:
        pf.first_line_indent = Pt(first_indent)
    if left_indent is not None:
        pf.left_indent = Pt(left_indent)
    if right_indent is not None:
        pf.right_indent = Pt(right_indent)
    if keep_with_next:
        pf.keep_with_next = True
    if text:
        r = p.add_run(text)
        set_font(r, cn=cn, en=en, size=size, bold=bold, color=color, italic=italic)
    return p


def rich(doc, segments, align=WD_ALIGN_PARAGRAPH.LEFT, line=1.5,
         first_indent=None, space_before=0, space_after=0, left_indent=None):
    """segments: [(text, dict)]，dict 为 set_font 的关键字参数"""
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing = line
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    if first_indent is not None:
        pf.first_line_indent = Pt(first_indent)
    if left_indent is not None:
        pf.left_indent = Pt(left_indent)
    for text, opt in segments:
        r = p.add_run(text)
        set_font(r, **opt)
    return p


def add_field(p, instr, size=12, cn=SONG, bold=False, cached=""):
    """插入 Word 域（TOC / PAGE / NUMPAGES / SECTIONPAGES 等）。
    cached 为域未更新时显示的占位文本。"""
    r = p.add_run()
    set_font(r, cn=cn, size=size, bold=bold)
    fc = OxmlElement('w:fldChar')
    fc.set(qn('w:fldCharType'), 'begin')
    r._element.append(fc)
    r2 = p.add_run()
    set_font(r2, cn=cn, size=size, bold=bold)
    it = OxmlElement('w:instrText')
    it.set(qn('xml:space'), 'preserve')
    it.text = instr
    r2._element.append(it)
    r3 = p.add_run()
    set_font(r3, cn=cn, size=size, bold=bold)
    fc2 = OxmlElement('w:fldChar')
    fc2.set(qn('w:fldCharType'), 'separate')
    r3._element.append(fc2)
    if cached:
        r4 = p.add_run()
        set_font(r4, cn=cn, size=size, bold=bold)
        r4.text = cached
    r5 = p.add_run()
    set_font(r5, cn=cn, size=size, bold=bold)
    fc3 = OxmlElement('w:fldChar')
    fc3.set(qn('w:fldCharType'), 'end')
    r5._element.append(fc3)
    return p


def add_toc(doc, levels="1-3", size=12):
    p = doc.add_paragraph()
    add_field(p, ' TOC \\o "%s" \\h \\z \\u ' % levels, size=size,
              cached="（在 Word 中右键“更新域”即可生成目录）")
    return p


def set_cell_bg(cell, hexcolor):
    tcpr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hexcolor)
    tcpr.append(shd)


def set_cell_text(cell, text, cn=SONG, en="Times New Roman", size=10.5,
                  bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, color=None):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(1)
    lines = str(text).split("\n")
    for i, ln in enumerate(lines):
        if i > 0:
            p = cell.add_paragraph()
            p.alignment = align
            p.paragraph_format.line_spacing = 1.15
        r = p.add_run(ln)
        set_font(r, cn=cn, en=en, size=size, bold=bold, color=color)
    return cell


def set_cell_valign(cell, val="center"):
    tcpr = cell._tc.get_or_add_tcPr()
    va = OxmlElement('w:vAlign')
    va.set(qn('w:val'), val)
    tcpr.append(va)


def remove_table_borders(table, keep=None):
    """去掉表格全部边框（用于排版用隐形表格）"""
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        if keep and edge in keep:
            continue
        el = OxmlElement('w:' + edge)
        el.set(qn('w:val'), 'none')
        el.set(qn('w:sz'), '0')
        borders.append(el)
    tblPr.append(borders)


def set_table_borders(table, sz=8, color="000000", style="single"):
    tbl = table._tbl
    tblPr = tbl.tblPr
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        el = OxmlElement('w:' + edge)
        el.set(qn('w:val'), style)
        el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), '0')
        el.set(qn('w:color'), color)
        borders.append(el)
    tblPr.append(borders)


def add_page_border(section, style="single", sz=12, color="2E5C8A",
                    space=12):
    """给某一节加页面边框。注意：必须 insert_element_before，
    直接 append 会因元素顺序错误被 Word 忽略。"""
    sectPr = section._sectPr
    old = sectPr.find(qn('w:pgBorders'))
    if old is not None:
        sectPr.remove(old)
    pb = OxmlElement('w:pgBorders')
    pb.set(qn('w:offsetFrom'), 'page')
    for name in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement('w:' + name)
        el.set(qn('w:val'), style)
        el.set(qn('w:sz'), str(sz))
        el.set(qn('w:space'), str(space))
        el.set(qn('w:color'), color)
        pb.append(el)
    sectPr.insert_element_before(
        pb, 'w:lnNumType', 'w:pgNumType', 'w:cols', 'w:formProt', 'w:vAlign',
        'w:noEndnote', 'w:titlePg', 'w:textDirection', 'w:bidi', 'w:rtlGutter',
        'w:docGrid', 'w:printerSettings')


def add_bottom_line(p, size=8, color="000000"):
    """给段落加下边框线"""
    ppr = p._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(size))
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), color)
    pbdr.append(bottom)
    ppr.append(pbdr)


def add_shading(p, hexcolor):
    ppr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hexcolor)
    ppr.append(shd)


def set_columns(section, num=2, space_cm=1.0):
    sectPr = section._sectPr
    cols = sectPr.find(qn('w:cols'))
    if cols is None:
        cols = OxmlElement('w:cols')
        sectPr.append(cols)
    cols.set(qn('w:num'), str(num))
    cols.set(qn('w:space'), str(int(space_cm * 567)))
    cols.set(qn('w:equalWidth'), '1')


def add_page_number_footer(section, text_center="", size=9, cn=SONG):
    f = section.footer
    p = f.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if text_center:
        r = p.add_run(text_center + "   ")
        set_font(r, cn=cn, size=size)
    r = p.add_run("第 ")
    set_font(r, cn=cn, size=size)
    add_field(p, " PAGE ", size=size, cn=cn)
    r = p.add_run(" 页 共 ")
    set_font(r, cn=cn, size=size)
    add_field(p, " NUMPAGES ", size=size, cn=cn)
    r = p.add_run(" 页")
    set_font(r, cn=cn, size=size)
    return f


def new_doc(base_size=10.5, cn=SONG):
    """新建文档并设置 Normal 样式（中英文分开指定）"""
    from docx import Document
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(base_size)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), cn)
    return doc


def a4(section, margins=(2.2, 2.2, 1.8, 1.8)):
    """A4 页面；margins = (left, right, top, bottom) cm"""
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(margins[0])
    section.right_margin = Cm(margins[1])
    section.top_margin = Cm(margins[2])
    section.bottom_margin = Cm(margins[3])


def safe_filename(name):
    """去掉文件名中的非法字符"""
    for ch in '\\/:*?"<>|':
        name = name.replace(ch, '')
    return name.strip() or "untitled"


def unlink_header_footer(section):
    """断开本节页眉页脚与上一节的继承"""
    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False


def restart_page_numbering(section, start=1):
    """本节页码从 start 重新编号（论文正文节用）"""
    sectPr = section._sectPr
    pg = sectPr.find(qn('w:pgNumType'))
    if pg is None:
        pg = OxmlElement('w:pgNumType')
        sectPr.append(pg)
    pg.set(qn('w:start'), str(start))


def set_repeat_header(row):
    """表格跨页时重复表头行"""
    trPr = row._tr.get_or_add_trPr()
    th = OxmlElement('w:tblHeader')
    trPr.append(th)


def add_art_page_border(section, style="apples", sz="16", color="C0504D"):
    """艺术型页面边框（小报用）"""
    sectPr = section._sectPr
    old = sectPr.find(qn('w:pgBorders'))
    if old is not None:
        sectPr.remove(old)
    pb = OxmlElement('w:pgBorders')
    pb.set(qn('w:offsetFrom'), 'page')
    for name in ('top', 'left', 'bottom', 'right'):
        el = OxmlElement('w:' + name)
        el.set(qn('w:val'), style)
        el.set(qn('w:sz'), sz)
        el.set(qn('w:space'), '10')
        el.set(qn('w:color'), color)
        pb.append(el)
    sectPr.insert_element_before(
        pb, 'w:lnNumType', 'w:pgNumType', 'w:cols', 'w:formProt', 'w:vAlign',
        'w:noEndnote', 'w:titlePg', 'w:textDirection', 'w:bidi', 'w:rtlGutter',
        'w:docGrid', 'w:printerSettings')


def set_doc_grid(doc, lines=28):
    """文档网格：每页行数（论文每行字数控制）"""
    sectPr = doc.sections[0]._sectPr
    g = sectPr.find(qn('w:docGrid'))
    if g is None:
        g = OxmlElement('w:docGrid')
        sectPr.append(g)
    g.set(qn('w:type'), 'linesAndChars')
    g.set(qn('w:linePitch'), str(round(15.6 * 20)))
    g.set(qn('w:charSpace'), '0')
    g.set(qn('w:lines'), str(lines))


def add_section_page_number_footer(section, size=10.5, cn=SONG):
    """页脚：第 X 页 共 Y 页，Y 用 SECTIONPAGES（本节页数）"""
    f = section.footer
    p = f.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("第 ")
    set_font(r, cn=cn, size=size)
    add_field(p, " PAGE ", size=size, cn=cn)
    r = p.add_run(" 页  共 ")
    set_font(r, cn=cn, size=size)
    add_field(p, " SECTIONPAGES ", size=size, cn=cn)
    r = p.add_run(" 页")
    set_font(r, cn=cn, size=size)
    return f
