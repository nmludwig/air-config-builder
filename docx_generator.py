"""
Generate a formatted Word doc matching the Dr. Liu AIR Configuration Playbook style.
Uses python-docx — pure Python, no Node.js required.
"""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Emu, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.section import WD_ORIENT
import re, io
from datetime import datetime

# ── Brand Colors ──────────────────────────────────────────
NAVY    = RGBColor(0x1E,0x3A,0x5F)
BLUE    = RGBColor(0x0F,0x62,0xFE)
GREEN   = RGBColor(0x16,0x65,0x34)
RED     = RGBColor(0x99,0x1B,0x1B)
AMBER   = RGBColor(0x92,0x40,0x0E)
SLATE   = RGBColor(0x47,0x55,0x69)
GRAY4   = RGBColor(0x94,0xA3,0xB8)
TEXT2   = RGBColor(0x33,0x41,0x55)
WHITE   = RGBColor(0xFF,0xFF,0xFF)

NAVY_H  = "1E3A5F"
BLUE_H  = "0F62FE"
BLUE_LT = "EFF6FF"
BLUE_MD = "DBEAFE"
BLUE_BD = "BFDBFE"
GRN_LT  = "DCFCE7"
GRN_BD  = "BBF7D0"
RED_LT  = "FEE2E2"
RED_BD  = "FECACA"
AMB_LT  = "FEF3C7"
AMB_BD  = "FDE68A"
GRAY_H  = "F3F4F6"
GRAY2_H = "E2E8F0"
WHITE_H = "FFFFFF"

# ── XML Helpers ───────────────────────────────────────────
def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    # Remove existing shd
    for s in tcPr.findall(qn('w:shd')):
        tcPr.remove(s)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_borders(cell, color="CCCCCC", sz="4"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for b in tcPr.findall(qn('w:tcBorders')):
        tcPr.remove(b)
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top','left','bottom','right']:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), sz)
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), color)
        tcBorders.append(b)
    tcPr.append(tcBorders)

def no_cell_borders(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for b in tcPr.findall(qn('w:tcBorders')):
        tcPr.remove(b)
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top','left','bottom','right']:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'none')
        b.set(qn('w:sz'), '0')
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), 'auto')
        tcBorders.append(b)
    tcPr.append(tcBorders)

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for m in tcPr.findall(qn('w:tcMar')):
        tcPr.remove(m)
    tcMar = OxmlElement('w:tcMar')
    for side, val in [('top',top),('bottom',bottom),('left',left),('right',right)]:
        m = OxmlElement(f'w:{side}')
        m.set(qn('w:w'), str(val))
        m.set(qn('w:type'), 'dxa')
        tcMar.append(m)
    tcPr.append(tcMar)

def set_row_height(row, height_twips):
    tr = row._tr
    trPr = tr.get_or_add_trPr()
    trHeight = OxmlElement('w:trHeight')
    trHeight.set(qn('w:val'), str(height_twips))
    trHeight.set(qn('w:hRule'), 'atLeast')
    trPr.append(trHeight)

def set_paragraph_border_bottom(para, color=BLUE_H, sz=12):
    pPr = para._p.get_or_add_pPr()
    for pb in pPr.findall(qn('w:pBdr')):
        pPr.remove(pb)
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), str(sz))
    bottom.set(qn('w:space'), '4')
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    pPr.append(pBdr)

def set_table_width(table, width_twips=9360):
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    for w in tblPr.findall(qn('w:tblW')):
        tblPr.remove(w)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), str(width_twips))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)

def set_col_width(cell, width_twips):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for w in tcPr.findall(qn('w:tcW')):
        tcPr.remove(w)
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(width_twips))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)

def para_spacing(para, before=0, after=0):
    pPr = para._p.get_or_add_pPr()
    for s in pPr.findall(qn('w:spacing')):
        pPr.remove(s)
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), str(before))
    sp.set(qn('w:after'), str(after))
    pPr.append(sp)

# ── Document Helpers ──────────────────────────────────────
def add_run(para, text, bold=False, italic=False, color=None, size=10, font='Arial'):
    run = para.add_run(str(text))
    run.bold = bold
    run.italic = italic
    run.font.name = font
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color
    return run

def add_h1(doc, text):
    p = doc.add_paragraph()
    para_spacing(p, before=240, after=120)
    set_paragraph_border_bottom(p, BLUE_H, sz=12)
    add_run(p, text, bold=True, color=NAVY, size=16)
    return p

def add_h2(doc, text):
    p = doc.add_paragraph()
    para_spacing(p, before=180, after=80)
    add_run(p, text, bold=True, color=NAVY, size=13)
    return p

def add_h3(doc, text):
    p = doc.add_paragraph()
    para_spacing(p, before=120, after=60)
    add_run(p, text, bold=True, color=BLUE, size=11)
    return p

def add_body(doc, text, color=None, bold=False, italic=False, size=10):
    p = doc.add_paragraph()
    para_spacing(p, before=40, after=40)
    add_run(p, text, bold=bold, italic=italic, color=color or TEXT2, size=size)
    return p

def add_bullet(doc, text, bold_pre=None, size=10):
    p = doc.add_paragraph(style='List Bullet')
    para_spacing(p, before=30, after=30)
    if bold_pre:
        add_run(p, bold_pre + ': ', bold=True, color=NAVY, size=size)
        add_run(p, text, size=size)
    else:
        add_run(p, text, size=size)
    return p

def add_numbered(doc, text, size=10):
    p = doc.add_paragraph(style='List Number')
    para_spacing(p, before=30, after=30)
    add_run(p, text, size=size)
    return p

def add_spacer(doc, size_pt=6):
    p = doc.add_paragraph()
    para_spacing(p, before=0, after=0)
    p.paragraph_format.line_spacing = Pt(size_pt)
    return p

# ── Callout Box ───────────────────────────────────────────
def add_callout(doc, icon, label, body_text, bg_color, border_color, text_color):
    """Two-column callout: icon col | label+body col"""
    table = doc.add_table(rows=1, cols=2)
    set_table_width(table, 9360)
    row = table.rows[0]

    # Icon cell
    c0 = row.cells[0]
    set_col_width(c0, 500)
    no_cell_borders(c0)
    set_cell_bg(c0, bg_color)
    set_cell_margins(c0, top=120, bottom=120, left=120, right=80)
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_spacing(p0, 0, 0)
    add_run(p0, icon, size=13, color=text_color)

    # Content cell
    c1 = row.cells[1]
    set_col_width(c1, 8860)
    no_cell_borders(c1)
    set_cell_bg(c1, bg_color)
    set_cell_margins(c1, top=120, bottom=120, left=100, right=140)
    p1 = c1.paragraphs[0]
    para_spacing(p1, 0, 0)
    if label:
        add_run(p1, label + ':  ', bold=True, color=text_color, size=10)
    add_run(p1, body_text, color=text_color, size=10)

    add_spacer(doc, 8)
    return table

def add_info_callout(doc, label, body):
    return add_callout(doc, 'i', label, body, BLUE_MD, BLUE_BD, RGBColor(0x1E,0x3A,0x8A))

def add_warn_callout(doc, label, body):
    return add_callout(doc, '!', label, body, AMB_LT, AMB_BD, AMBER)

def add_danger_callout(doc, label, body):
    return add_callout(doc, '!', label, body, RED_LT, RED_BD, RED)

def add_success_callout(doc, label, body):
    return add_callout(doc, '✓', label, body, GRN_LT, GRN_BD, GREEN)

# ── Copy Box (company description etc) ───────────────────
def add_copy_box(doc, text):
    table = doc.add_table(rows=1, cols=1)
    set_table_width(table, 9360)
    cell = table.cell(0, 0)
    set_col_width(cell, 9360)
    set_cell_bg(cell, BLUE_MD)
    set_cell_borders(cell, BLUE_BD, sz="4")
    set_cell_margins(cell, top=160, bottom=160, left=200, right=200)
    p = cell.paragraphs[0]
    para_spacing(p, 0, 0)
    run = p.add_run(text)
    run.font.name = 'Arial'
    run.font.size = Pt(10)
    run.font.color.rgb = NAVY
    add_spacer(doc, 8)

# ── Key-Value Table ───────────────────────────────────────
def add_kv_table(doc, rows, col1_w=2800, col2_w=6560):
    """Two-column label:value table"""
    total = col1_w + col2_w
    table = doc.add_table(rows=len(rows), cols=2)
    set_table_width(table, total)
    for i, (key, val) in enumerate(rows):
        bg = GRAY_H if i % 2 == 0 else WHITE_H
        c0, c1 = table.rows[i].cells
        set_col_width(c0, col1_w)
        set_col_width(c1, col2_w)
        set_cell_bg(c0, GRAY_H)
        set_cell_borders(c0, GRAY2_H)
        set_cell_bg(c1, bg)
        set_cell_borders(c1, GRAY2_H)
        set_cell_margins(c0, top=80, bottom=80, left=120, right=100)
        set_cell_margins(c1, top=80, bottom=80, left=120, right=120)
        p0 = c0.paragraphs[0]
        para_spacing(p0, 0, 0)
        add_run(p0, str(key), bold=True, color=NAVY, size=10)
        p1 = c1.paragraphs[0]
        para_spacing(p1, 0, 0)
        add_run(p1, str(val), size=10)
    add_spacer(doc, 8)

# ── FAQ Table ─────────────────────────────────────────────
def add_faq_table(doc, question, answer):
    C1, C2 = 3400, 5960
    table = doc.add_table(rows=2, cols=2)
    set_table_width(table, 9360)
    # Header row
    hr = table.rows[0]
    c0h, c1h = hr.cells
    set_col_width(c0h, C1); set_col_width(c1h, C2)
    set_cell_bg(c0h, NAVY_H); set_cell_borders(c0h, NAVY_H)
    set_cell_bg(c1h, NAVY_H); set_cell_borders(c1h, NAVY_H)
    set_cell_margins(c0h, 80,80,120,80); set_cell_margins(c1h, 80,80,120,120)
    p0h = c0h.paragraphs[0]; para_spacing(p0h,0,0)
    add_run(p0h, "QUESTION (what caller says)", bold=True, color=WHITE, size=9)
    p1h = c1h.paragraphs[0]; para_spacing(p1h,0,0)
    add_run(p1h, "ANSWER (what AIR says — copy exactly)", bold=True, color=WHITE, size=9)
    # Content row
    cr = table.rows[1]
    c0c, c1c = cr.cells
    set_col_width(c0c, C1); set_col_width(c1c, C2)
    set_cell_bg(c0c, BLUE_LT); set_cell_borders(c0c, GRAY2_H)
    set_cell_bg(c1c, GRAY_H); set_cell_borders(c1c, GRAY2_H)
    set_cell_margins(c0c, 100,100,120,80); set_cell_margins(c1c, 100,100,120,120)
    p0c = c0c.paragraphs[0]; para_spacing(p0c,0,0)
    add_run(p0c, question, italic=True, size=9)
    p1c = c1c.paragraphs[0]; para_spacing(p1c,0,0)
    add_run(p1c, answer, size=9)
    add_spacer(doc, 8)

# ── Routing Table ─────────────────────────────────────────
def add_routing_table(doc, rules):
    """rules = list of (keywords, route_to, notes)"""
    C = [3200, 2160, 4000]
    cols = len(C)
    table = doc.add_table(rows=1+len(rules), cols=cols)
    set_table_width(table, sum(C))
    headers = ["TRIGGER KEYWORDS / PHRASES", "ROUTE TO", "NOTES / COVERS"]
    hr = table.rows[0]
    for j, (cell, hdr) in enumerate(zip(hr.cells, headers)):
        set_col_width(cell, C[j])
        set_cell_bg(cell, NAVY_H)
        set_cell_borders(cell, NAVY_H)
        set_cell_margins(cell, 80,80,120,80)
        p = cell.paragraphs[0]; para_spacing(p,0,0)
        add_run(p, hdr, bold=True, color=WHITE, size=9)
    for i, rule in enumerate(rules):
        bg = BLUE_LT if i%2==0 else WHITE_H
        bg2 = GRAY_H if i%2==0 else WHITE_H
        row = table.rows[i+1]
        vals = [rule.get('keywords',''), rule.get('route',''), rule.get('note','')]
        bgs = [bg, bg2, bg2]
        for j, (cell, val, bg_) in enumerate(zip(row.cells, vals, bgs)):
            set_col_width(cell, C[j])
            set_cell_bg(cell, bg_)
            set_cell_borders(cell, GRAY2_H)
            set_cell_margins(cell, 80,80,120,80)
            p = cell.paragraphs[0]; para_spacing(p,0,0)
            add_run(p, val, italic=(j==0), bold=(j==1), size=9)
    add_spacer(doc, 8)

# ── Content Parser & Renderer ─────────────────────────────
def render_content(doc, content):
    lines = content.split('\n')
    i = 0
    current_faq = {}
    current_route = {}
    in_faq_block = False
    in_route_block = False
    route_rules = []

    def flush_faq():
        nonlocal current_faq, in_faq_block
        if current_faq.get('q') and current_faq.get('a'):
            add_faq_table(doc, current_faq['q'], current_faq['a'])
        current_faq = {}
        in_faq_block = False

    def flush_route():
        nonlocal current_route, route_rules
        if current_route.get('keywords'):
            route_rules.append(current_route)
        current_route = {}

    def flush_route_table():
        nonlocal route_rules
        if route_rules:
            add_routing_table(doc, route_rules)
            route_rules = []

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Empty line
        if not stripped:
            i += 1
            continue

        # ── Headings ────────────────────────────────────
        # Top-level section: ## 1. or # 1.
        if re.match(r'^#{1,2}\s+\d+[\.\s]', stripped):
            flush_faq()
            flush_route()
            flush_route_table()
            text = re.sub(r'^#+\s+', '', stripped).replace('**','')
            # Page break before major sections (not the first)
            if text and not text.startswith('1'):
                doc.add_page_break()
            add_h1(doc, text)
            i += 1; continue

        # Sub-section: ### or ##
        if re.match(r'^#{2,3}\s+', stripped) and not re.match(r'^#{1,2}\s+\d+[\.\s]', stripped):
            flush_faq()
            flush_route()
            text = re.sub(r'^#+\s+', '', stripped).replace('**','')
            add_h2(doc, text)
            i += 1; continue

        # Sub-sub-section: #### or FAQ entry title
        if re.match(r'^#{4,}\s+', stripped) or re.match(r'^###\s+FAQ', stripped):
            flush_faq()
            text = re.sub(r'^#+\s+', '', stripped).replace('**','')
            in_faq_block = True
            current_faq = {'title': text, 'q_parts': [], 'a': ''}
            add_h3(doc, text)
            i += 1; continue

        # ── Code blocks ─────────────────────────────────
        if stripped.startswith('```'):
            flush_faq()
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            i += 1
            text = '\n'.join(code_lines).strip()
            if text:
                add_copy_box(doc, text)
            continue

        # ── Horizontal rule ──────────────────────────────
        if re.match(r'^-{3,}$', stripped):
            i += 1; continue

        # ── Callout patterns ─────────────────────────────
        # ⚠️ warning
        if stripped.startswith('⚠️') or stripped.lower().startswith('warning:'):
            text = re.sub(r'^⚠️\s*', '', stripped).replace('**','')
            m = re.match(r'\*\*(.+?)\*\*[:\s]+(.+)', text)
            if m:
                add_warn_callout(doc, m.group(1), m.group(2))
            else:
                add_warn_callout(doc, 'Note', text)
            i += 1; continue

        # ❌ prohibition
        if stripped.startswith('❌') or stripped.startswith('🚫'):
            text = re.sub(r'^[❌🚫]\s*', '', stripped).replace('**','')
            add_danger_callout(doc, 'Must Never', text)
            i += 1; continue

        # ✅ success
        if stripped.startswith('✅'):
            text = re.sub(r'^✅\s*', '', stripped).replace('**','')
            add_success_callout(doc, 'Yes', text)
            i += 1; continue

        # ── Bullet points ────────────────────────────────
        if re.match(r'^[-*]\s+', stripped):
            text = re.sub(r'^[-*]\s+', '', stripped)
            # Strip markdown bold
            bold_match = re.match(r'\*\*(.+?)\*\*[:\s—-]+(.+)', text)
            if bold_match:
                add_bullet(doc, bold_match.group(2).replace('**',''), bold_pre=bold_match.group(1))
            else:
                add_bullet(doc, text.replace('**',''))
            i += 1; continue

        # ── Numbered lists ───────────────────────────────
        if re.match(r'^\d+\.\s+', stripped):
            text = re.sub(r'^\d+\.\s+', '', stripped).replace('**','')
            add_numbered(doc, text)
            i += 1; continue

        # ── FAQ parsing ──────────────────────────────────
        # Question variants section
        if re.match(r'^Question variants', stripped, re.I):
            i += 1
            q_parts = []
            while i < len(lines):
                ql = lines[i].strip()
                if not ql or re.match(r'^\*\*Answer', ql, re.I) or ql.startswith('---'):
                    break
                if re.match(r'^[-*"]\s*', ql):
                    q_parts.append(re.sub(r'^[-*"]\s*', '', ql).strip('"').strip())
                i += 1
            if in_faq_block:
                current_faq['q_parts'] = q_parts
                current_faq['q'] = ' / '.join(q_parts)
            continue

        # Answer line
        answer_match = re.match(r'^\*\*Answer[^:]*:\*\*\s*(.*)', stripped, re.I)
        if answer_match:
            ans = answer_match.group(1).strip().strip('"')
            if not ans:
                # Answer on next lines
                i += 1
                ans_lines = []
                while i < len(lines):
                    al = lines[i].strip()
                    if not al or re.match(r'^---', al) or re.match(r'^###', al) or re.match(r'^\*\*Rule', al):
                        break
                    ans_lines.append(al.strip('"').strip('*'))
                    i += 1
                ans = ' '.join(ans_lines).strip('"')
            if in_faq_block:
                current_faq['a'] = ans
                flush_faq()
            continue

        # ── Transfer by Context parsing ──────────────────
        # **Rule N: Department**
        rule_match = re.match(r'^\*\*Rule\s+\d+[:\s]+(.+)\*\*', stripped)
        if rule_match:
            flush_route()
            current_route = {'title': rule_match.group(1).strip(), 'keywords':'', 'route':'', 'note':''}
            i += 1; continue

        kw_match = re.match(r'^Keywords?[:\s]+(.+)', stripped, re.I)
        if kw_match and current_route is not None:
            current_route['keywords'] = kw_match.group(1).replace('**','').strip()
            i += 1; continue

        route_match = re.match(r'^Route\s+to[:\s]+(.+)', stripped, re.I)
        if route_match and current_route is not None:
            current_route['route'] = route_match.group(1).replace('**','').strip()
            i += 1; continue

        covers_match = re.match(r'^(?:Covers|Note|AIR holding message)[:\s]+(.+)', stripped, re.I)
        if covers_match and current_route is not None:
            current_route['note'] = covers_match.group(1).replace('**','').strip('"').strip()
            i += 1; continue

        # ── Bold label: value ────────────────────────────
        bold_kv = re.match(r'^\*\*(.+?)\*\*[:\s]+(.+)', stripped)
        if bold_kv:
            key = bold_kv.group(1).strip()
            val = bold_kv.group(2).replace('**','').strip()
            p = doc.add_paragraph()
            para_spacing(p, before=40, after=40)
            add_run(p, key + ': ', bold=True, color=NAVY, size=10)
            add_run(p, val, size=10)
            i += 1; continue

        # ── Bold-only line ───────────────────────────────
        if re.match(r'^\*\*.+\*\*$', stripped):
            text = stripped.replace('**','')
            p = doc.add_paragraph()
            para_spacing(p, before=60, after=40)
            add_run(p, text, bold=True, color=NAVY, size=10)
            i += 1; continue

        # ── Regular body ─────────────────────────────────
        text = stripped.replace('**','').replace('*','')
        if text:
            p = doc.add_paragraph()
            para_spacing(p, before=40, after=40)
            add_run(p, text, size=10, color=TEXT2)
        i += 1

    # Flush any remaining
    flush_faq()
    flush_route()
    flush_route_table()


# ── Main Generator ────────────────────────────────────────
def generate_docx(biz_name, content, prepared_by="RingCentral SE"):
    doc = Document()

    # Page setup
    for section in doc.sections:
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1)
        section.right_margin  = Inches(1)

    # Default font
    doc.styles['Normal'].font.name = 'Arial'
    doc.styles['Normal'].font.size = Pt(10)

    today = datetime.now().strftime('%B %Y')

    # ── COVER PAGE ─────────────────────────────────────────
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_spacing(p, before=480, after=80)
    add_run(p, "RingCentral AI Receptionist", bold=True, color=NAVY, size=26)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_spacing(p, before=0, after=80)
    add_run(p, "Configuration Playbook", color=BLUE, size=20)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_spacing(p, before=80, after=40)
    add_run(p, biz_name, bold=True, color=SLATE, size=16)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_spacing(p, before=40, after=40)
    add_run(p, "Ready-to-paste configuration for every AIR field", italic=True, color=GRAY4, size=10)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_spacing(p, before=20, after=20)
    add_run(p, f"Prepared by: {prepared_by}  |  {today}", italic=True, color=GRAY4, size=9)

    doc.add_page_break()

    # ── MAIN CONTENT ───────────────────────────────────────
    render_content(doc, content)

    # ── FOOTER LINE ────────────────────────────────────────
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para_spacing(p, before=240, after=60)
    set_paragraph_border_bottom(p, GRAY2_H, sz=4)
    add_run(p, f"RingCentral AI Receptionist  |  {biz_name}  |  {today}",
            italic=True, color=GRAY4, size=8)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()
