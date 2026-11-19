# -*- coding: utf-8 -*-
"""Markdown → 中文公文格式 Word（.docx）

排版：A4，页边距上下 2.54cm / 左右 3.17cm（可用宽度 14.66cm）
- 正文 宋体 12pt、行距 1.5；标题 黑体，一级 20pt / 二级 15pt / 三级 13.5pt
- 标题用真 Heading 样式（Word 导航窗格、目录可用），颜色强制为黑
- 表格 Table Grid 边框 + 表头底纹 DEEAF6 + 表头加粗，列宽按内容长度自适应
- 引用块：楷体 + 左侧灰线；页脚居中「第 X 页 共 Y 页」域
- 支持 **加粗**、`代码`、引用块、有序/无序列表、分隔线、**图片**（`![题注](路径)`，
  按页宽 14.4cm 居中插入，题注用楷体 10.5pt 居中；相对路径按 md 所在目录解析）

两个必须绕开的坑（已在本脚本内处理）：
1. **有序列表不要用 Word 自动编号**。同一 numId 会让文档里多个独立列表
   连续编号（如「七、处理建议」1-7 条之后，「八、核查说明」接着从 8 开始）。
   本脚本改用 markdown 里的**字面序号** + 悬挂缩进（first_line_indent 负值）。
2. **列宽必须算**。硬编码列宽很容易超过页面可用宽度 14.66cm 而被 Word 截断；
   本脚本按各列最大内容长度（中文按 2 字符宽折算）按比例分配，最小 1.5cm。
"""
import sys, io, re, os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

if __name__ == '__main__':
    # write_through=True：否则重包 stdout 变块缓冲，看不到进度
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8',
                                  errors='replace', write_through=True)

SONG = '宋体'
HEI = '黑体'
KAI = '楷体'
WEST = 'Times New Roman'


def set_font(run, ea=SONG, ascii_f=WEST, size=None, bold=None, color=None):
    run.font.name = ascii_f
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), ascii_f)
    rFonts.set(qn('w:hAnsi'), ascii_f)
    rFonts.set(qn('w:eastAsia'), ea)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hexcolor)
    tcPr.append(shd)


def add_runs(par, text, ea=SONG, size=12, bold=False, color=None):
    """解析 **加粗** 与 `代码`"""
    for part in re.split(r'(\*\*[^*]+\*\*|`[^`]+`)', text):
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            r = par.add_run(part[2:-2])
            set_font(r, ea=HEI if ea == SONG else ea, size=size, bold=True, color=color)
        elif part.startswith('`') and part.endswith('`'):
            r = par.add_run(part[1:-1])
            set_font(r, ea=KAI, ascii_f='Consolas', size=size - 0.5, bold=bold, color=color)
        else:
            r = par.add_run(part)
            set_font(r, ea=ea, size=size, bold=bold, color=color)


def set_body(par, space_before=0, space_after=6, line=1.5, indent=0):
    pf = par.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = line
    if indent:
        pf.left_indent = Cm(indent)


def cell_widths(table, rows, total_cm=14.4):
    """按各列最大内容长度按比例分配列宽（页面可用宽度约 14.66cm）"""
    ncol = len(table.columns)
    if ncol == 0:
        return
    # 中文按 2 个字符宽度折算，用于估权重
    def w(s):
        return sum(2 if ord(c) > 0x2E80 else 1 for c in s)

    weights = []
    for ci in range(ncol):
        mx = max((w(r[ci]) for r in rows if ci < len(r)), default=4)
        weights.append(min(max(mx, 4), 46))
    tot = sum(weights)
    ws = [total_cm * x / tot for x in weights]
    # 最小列宽 1.5cm，多出来的宽度按权重回补给其他列
    MIN = 1.5
    deficit = sum(MIN - x for x in ws if x < MIN)
    ws = [max(x, MIN) for x in ws]
    if deficit > 0:
        slack = [i for i, x in enumerate(ws) if x > MIN]
        if slack:
            per = deficit / len(slack)
            for i in slack:
                ws[i] -= per
    table.autofit = False
    for row in table.rows:
        for i, cw in enumerate(ws):
            if i < len(row.cells):
                row.cells[i].width = Cm(round(cw, 2))


def border_left(par, color='BFBFBF'):
    pPr = par._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    left = OxmlElement('w:left')
    left.set(qn('w:val'), 'single')
    left.set(qn('w:sz'), '18')
    left.set(qn('w:space'), '6')
    left.set(qn('w:color'), color)
    pbdr.append(left)
    pPr.append(pbdr)


def hr(par):
    pPr = par._p.get_or_add_pPr()
    pbdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '1')
    bottom.set(qn('w:color'), '999999')
    pbdr.append(bottom)
    pPr.append(pbdr)


def build_table(doc, rows, ncol_hint=None):
    ncol = max(len(r) for r in rows)
    rows = [r + [''] * (ncol - len(r)) for r in rows]
    t = doc.add_table(rows=len(rows), cols=ncol)
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for ri, row in enumerate(rows):
        for ci, txt in enumerate(row):
            cell = t.cell(ri, ci)
            cell.text = ''
            p = cell.paragraphs[0]
            set_body(p, space_before=2, space_after=2, line=1.15)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if (ri == 0 or ci == 0) else WD_ALIGN_PARAGRAPH.LEFT
            add_runs(p, txt, ea=HEI if ri == 0 else SONG,
                     size=10.5, bold=(ri == 0))
            if ri == 0:
                shade(cell, 'DEEAF6')
    return t


def add_page_number(sec):
    """页脚居中插入「第 X 页 / 共 Y 页」域"""
    p = sec.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_body(p, space_before=0, space_after=0, line=1.0)

    def field(instr):
        r = p.add_run()
        f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
        it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = instr
        f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'end')
        r._r.append(f1); r._r.append(it); r._r.append(f2)
        set_font(r, ea=SONG, size=9)

    r = p.add_run('第 '); set_font(r, ea=SONG, size=9)
    field(' PAGE ')
    r = p.add_run(' 页  共 '); set_font(r, ea=SONG, size=9)
    field(' NUMPAGES ')
    r = p.add_run(' 页'); set_font(r, ea=SONG, size=9)


def md_to_docx(md_path, docx_path):
    lines = open(md_path, encoding='utf-8').read().split('\n')
    doc = Document()

    st = doc.styles['Normal']
    st.font.name = WEST
    st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), SONG)

    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(2.54)
    sec.left_margin = sec.right_margin = Cm(3.17)
    add_page_number(sec)

    i, n = 0, len(lines)
    first_h1_done = False
    while i < n:
        ln = lines[i].rstrip()

        if not ln.strip():
            i += 1
            continue

        # 分隔线
        if re.fullmatch(r'-{3,}', ln.strip()):
            p = doc.add_paragraph()
            set_body(p, space_before=4, space_after=8, line=1.0)
            hr(p)
            i += 1
            continue

        # 图片：![题注](相对或绝对路径)
        _img = re.fullmatch(r'!\[(.*?)\]\((.+?)\)', ln.strip())
        if _img:
            cap, src = _img.group(1).strip(), _img.group(2).strip()
            if not os.path.isabs(src):
                src = os.path.join(os.path.dirname(os.path.abspath(md_path)), src)
            if os.path.exists(src):
                pw = doc.add_paragraph()
                pw.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_body(pw, space_before=6, space_after=2, line=1.0)
                pw.add_run().add_picture(src, width=Cm(14.4))
                if cap:
                    pc = doc.add_paragraph()
                    pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    set_body(pc, space_before=0, space_after=8, line=1.0)
                    add_runs(pc, cap, ea=KAI, size=10.5, color=RGBColor(0x40, 0x40, 0x40))
            else:
                _p = doc.add_paragraph()
                set_body(_p)
                add_runs(_p, '［缺图：%s］' % os.path.basename(src), ea=KAI, size=10.5,
                         color=RGBColor(0xC0, 0x00, 0x00))
            i += 1
            continue

        # 表格
        if ln.lstrip().startswith('|') and i + 1 < n and re.match(r'^\s*\|[\s:\-|]+\|\s*$', lines[i + 1]):
            block = []
            while i < n and lines[i].lstrip().startswith('|'):
                row = [c.strip() for c in lines[i].strip().strip('|').split('|')]
                if not re.match(r'^[\s:\-]+$', ''.join(row)):
                    block.append(row)
                i += 1
            t = build_table(doc, block)
            cell_widths(t, block)
            p = doc.add_paragraph()
            set_body(p, space_before=0, space_after=4, line=1.0)
            continue

        # 标题
        m = re.match(r'^(#{1,6})\s+(.*)$', ln)
        if m:
            level, txt = len(m.group(1)), m.group(2).strip()
            txt = txt.replace('**', '')
            if level == 1:
                p = doc.add_paragraph(style='Heading 1')
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                set_body(p, space_before=6, space_after=6, line=1.4)
                add_runs(p, txt, ea=HEI, size=20, bold=True, color=RGBColor(0, 0, 0))
                first_h1_done = True
            elif level == 2:
                p = doc.add_paragraph(style='Heading 2')
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if not first_h1_done else WD_ALIGN_PARAGRAPH.LEFT
                set_body(p, space_before=10, space_after=8, line=1.4)
                add_runs(p, txt, ea=HEI, size=15, bold=True, color=RGBColor(0, 0, 0))
                first_h1_done = True
            elif level == 3:
                p = doc.add_paragraph(style='Heading 3')
                set_body(p, space_before=10, space_after=6, line=1.4)
                add_runs(p, txt, ea=HEI, size=13.5, bold=True, color=RGBColor(0, 0, 0))
            else:
                p = doc.add_paragraph(style='Heading 4')
                set_body(p, space_before=8, space_after=4, line=1.4)
                add_runs(p, txt, ea=HEI, size=12, bold=True, color=RGBColor(0, 0, 0))
            i += 1
            continue

        # 引用
        if ln.lstrip().startswith('>'):
            buf = []
            while i < n and lines[i].lstrip().startswith('>'):
                buf.append(lines[i].lstrip()[1:].strip())
                i += 1
            for b in buf:
                p = doc.add_paragraph()
                set_body(p, space_before=2, space_after=2, line=1.4, indent=0.6)
                border_left(p)
                add_runs(p, b, ea=KAI, size=10.5, color=RGBColor(0x44, 0x44, 0x44))
            continue

        # 列表
        m = re.match(r'^(\s*)([-*]|\d+\.)\s+(.*)$', ln)
        if m:
            indent = len(m.group(1)) // 2
            ordered = bool(re.match(r'\d+\.', m.group(2)))
            # 用 markdown 里的字面序号，不用 Word 自动编号：
            # 同一 numId 会让「七、处理建议」和「八、核查说明」两个列表连号（1..7 后接 8..11）。
            marker = (m.group(2) if ordered else '•') + ' '
            txt = marker + m.group(3)
            p = doc.add_paragraph()
            set_body(p, space_before=1, space_after=3, line=1.4)
            pf = p.paragraph_format
            base = 0.74 + indent * 0.6
            pf.left_indent = Cm(base + 0.62)
            pf.first_line_indent = Cm(-0.62)      # 悬挂缩进，换行对齐
            add_runs(p, txt, size=11.5)
            i += 1
            continue

        # 普通段落
        p = doc.add_paragraph()
        set_body(p, space_before=2, space_after=5, line=1.5)
        add_runs(p, ln.strip())
        i += 1

    doc.save(docx_path)
    return docx_path


if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser(description='Markdown → 中文公文格式 Word')
    ap.add_argument('inputs', nargs='+', help='Markdown 文件')
    ap.add_argument('-o', '--outdir', default=None, help='输出目录')
    ap.add_argument('--flat', action='store_true', help='输出到 md 同目录，不建子目录')
    a = ap.parse_args()

    if len(a.inputs) == 2 and a.inputs[1].lower().endswith('.docx'):
        md_to_docx(a.inputs[0], a.inputs[1])
        print('✓', a.inputs[1])
    else:
        for src in a.inputs:
            if not os.path.exists(src):
                print('!! 缺失', src)
                continue
            base = os.path.splitext(os.path.basename(src))[0]
            if a.outdir:
                od = a.outdir
            elif a.flat:
                od = os.path.dirname(os.path.abspath(src))
            else:
                od = os.path.join(os.path.dirname(os.path.abspath(src)), 'Word版报告')
            os.makedirs(od, exist_ok=True)
            out = md_to_docx(src, os.path.join(od, base + '.docx'))
            print('✓', os.path.basename(out), os.path.getsize(out), 'bytes')
