"""
Generate a formatted Word doc matching the Dr. Liu AIR Configuration Playbook style exactly.
"""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re, io
from datetime import datetime

# ── Colors ────────────────────────────────────────────────
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
    for s in tcPr.findall(qn('w:shd')): tcPr.remove(s)
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_cell_borders(cell, color="CCCCCC", sz="4"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for b in tcPr.findall(qn('w:tcBorders')): tcPr.remove(b)
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top','left','bottom','right']:
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), sz)
        b.set(qn('w:space'), '0')
        b.set(qn('w:color'), color)
        tcBorders.append(b)
    tcPr.append(tcBorders)

def no_borders(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for b in tcPr.findall(qn('w:tcBorders')): tcPr.remove(b)
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
    for m in tcPr.findall(qn('w:tcMar')): tcPr.remove(m)
    tcMar = OxmlElement('w:tcMar')
    for side, val in [('top',top),('bottom',bottom),('left',left),('right',right)]:
        m = OxmlElement(f'w:{side}')
        m.set(qn('w:w'), str(val))
        m.set(qn('w:type'), 'dxa')
        tcMar.append(m)
    tcPr.append(tcMar)

def set_col_width(cell, w):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for x in tcPr.findall(qn('w:tcW')): tcPr.remove(x)
    tcW = OxmlElement('w:tcW')
    tcW.set(qn('w:w'), str(w))
    tcW.set(qn('w:type'), 'dxa')
    tcPr.append(tcW)

def set_table_width(table, w=9360):
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    for x in tblPr.findall(qn('w:tblW')): tblPr.remove(x)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), str(w))
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)

def pspacing(para, before=0, after=0):
    pPr = para._p.get_or_add_pPr()
    for s in pPr.findall(qn('w:spacing')): pPr.remove(s)
    sp = OxmlElement('w:spacing')
    sp.set(qn('w:before'), str(before))
    sp.set(qn('w:after'), str(after))
    pPr.append(sp)

def para_border_bottom(para, color=BLUE_H, sz=12):
    pPr = para._p.get_or_add_pPr()
    for pb in pPr.findall(qn('w:pBdr')): pPr.remove(pb)
    pBdr = OxmlElement('w:pBdr')
    b = OxmlElement('w:bottom')
    b.set(qn('w:val'), 'single')
    b.set(qn('w:sz'), str(sz))
    b.set(qn('w:space'), '4')
    b.set(qn('w:color'), color)
    pBdr.append(b)
    pPr.append(pBdr)

# ── Run helper ────────────────────────────────────────────
def run(para, text, bold=False, italic=False, color=None, size=10):
    r = para.add_run(str(text))
    r.bold = bold
    r.italic = italic
    r.font.name = 'Arial'
    r.font.size = Pt(size)
    if color: r.font.color.rgb = color
    return r

# ── Block builders ────────────────────────────────────────
def add_h1(doc, text):
    p = doc.add_paragraph()
    pspacing(p, 280, 120)
    para_border_bottom(p, BLUE_H, 12)
    run(p, text, bold=True, color=NAVY, size=16)

def add_h2(doc, text):
    p = doc.add_paragraph()
    pspacing(p, 200, 80)
    run(p, text, bold=True, color=NAVY, size=13)

def add_h3(doc, text):
    p = doc.add_paragraph()
    pspacing(p, 140, 60)
    run(p, text, bold=True, color=BLUE, size=11)

def add_body(doc, text, bold=False, italic=False, color=None, size=10):
    p = doc.add_paragraph()
    pspacing(p, 40, 40)
    run(p, text, bold=bold, italic=italic, color=color or TEXT2, size=size)
    return p

def add_label_value(doc, label, value, size=10):
    p = doc.add_paragraph()
    pspacing(p, 40, 40)
    run(p, label + ': ', bold=True, color=NAVY, size=size)
    run(p, value, size=size)

def add_bullet(doc, text, bold_pre=None):
    p = doc.add_paragraph(style='List Bullet')
    pspacing(p, 30, 30)
    if bold_pre:
        run(p, bold_pre + ': ', bold=True, color=NAVY, size=10)
        run(p, text, size=10)
    else:
        run(p, text, size=10)

def add_numbered(doc, text):
    p = doc.add_paragraph(style='List Number')
    pspacing(p, 30, 30)
    run(p, text, size=10)

def add_spacer(doc):
    p = doc.add_paragraph()
    pspacing(p, 0, 0)
    p.paragraph_format.line_spacing = Pt(4)

def add_copy_box(doc, text):
    """Blue-tinted monospace box for copy-paste content"""
    table = doc.add_table(rows=1, cols=1)
    set_table_width(table, 9360)
    cell = table.cell(0,0)
    set_col_width(cell, 9360)
    set_cell_bg(cell, BLUE_MD)
    set_cell_borders(cell, BLUE_BD, "4")
    set_cell_margins(cell, 160, 160, 200, 200)
    p = cell.paragraphs[0]
    pspacing(p, 0, 0)
    r = p.add_run(text)
    r.font.name = 'Courier New'
    r.font.size = Pt(9)
    r.font.color.rgb = NAVY
    add_spacer(doc)

def add_callout(doc, icon, label, body_text, bg, bdr_color, text_color):
    table = doc.add_table(rows=1, cols=2)
    set_table_width(table, 9360)
    row = table.rows[0]
    c0, c1 = row.cells[0], row.cells[1]

    set_col_width(c0, 480)
    no_borders(c0)
    set_cell_bg(c0, bg)
    set_cell_margins(c0, 120, 120, 120, 80)
    p0 = c0.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pspacing(p0, 0, 0)
    run(p0, icon, size=12, color=text_color)

    set_col_width(c1, 8880)
    no_borders(c1)
    set_cell_bg(c1, bg)
    set_cell_margins(c1, 120, 120, 100, 140)
    p1 = c1.paragraphs[0]
    pspacing(p1, 0, 0)
    if label:
        run(p1, label + ':  ', bold=True, color=text_color, size=10)
    run(p1, body_text, color=text_color, size=10)
    add_spacer(doc)

def info_box(doc, label, text):
    add_callout(doc, 'i', label, text, BLUE_MD, BLUE_BD, RGBColor(0x1E,0x3A,0x8A))

def warn_box(doc, label, text):
    add_callout(doc, '!', label, text, AMB_LT, AMB_BD, AMBER)

def danger_box(doc, label, text):
    add_callout(doc, '!', label, text, RED_LT, RED_BD, RED)

def add_faq_table(doc, title, question_variants, answer):
    """Full FAQ block: h3 title + 2-col table with navy header"""
    add_h3(doc, title)
    C1, C2 = 3400, 5960
    table = doc.add_table(rows=2, cols=2)
    set_table_width(table, 9360)

    # Header row
    hr = table.rows[0]
    h0, h1 = hr.cells
    for cell, txt, w in [(h0,"QUESTION (what caller says)",C1),(h1,"ANSWER (what AIR says — copy exactly)",C2)]:
        set_col_width(cell, w)
        set_cell_bg(cell, NAVY_H)
        set_cell_borders(cell, NAVY_H)
        set_cell_margins(cell, 80, 80, 120, 80)
        p = cell.paragraphs[0]
        pspacing(p, 0, 0)
        run(p, txt, bold=True, color=WHITE, size=9)

    # Content row
    cr = table.rows[1]
    c0, c1 = cr.cells
    set_col_width(c0, C1); set_col_width(c1, C2)
    set_cell_bg(c0, BLUE_LT); set_cell_borders(c0, GRAY2_H)
    set_cell_bg(c1, GRAY_H);  set_cell_borders(c1, GRAY2_H)
    set_cell_margins(c0, 100, 100, 120, 80)
    set_cell_margins(c1, 100, 100, 120, 120)

    p0 = c0.paragraphs[0]; pspacing(p0, 0, 0)
    run(p0, question_variants, italic=True, size=9)

    p1 = c1.paragraphs[0]; pspacing(p1, 0, 0)
    run(p1, answer, size=9)

    add_spacer(doc)

def add_routing_table(doc, rules):
    """Routing rules table"""
    if not rules: return
    C = [3200, 2160, 4000]
    table = doc.add_table(rows=1+len(rules), cols=3)
    set_table_width(table, sum(C))

    for j, (hdr, w) in enumerate(zip(
        ["TRIGGER KEYWORDS / PHRASES", "ROUTE TO", "COVERS / NOTES"], C)):
        cell = table.rows[0].cells[j]
        set_col_width(cell, w)
        set_cell_bg(cell, NAVY_H)
        set_cell_borders(cell, NAVY_H)
        set_cell_margins(cell, 80, 80, 120, 80)
        p = cell.paragraphs[0]; pspacing(p, 0, 0)
        run(p, hdr, bold=True, color=WHITE, size=9)

    for i, rule in enumerate(rules):
        bg1 = BLUE_LT if i%2==0 else WHITE_H
        bg2 = GRAY_H  if i%2==0 else WHITE_H
        vals  = [rule.get('keywords',''), rule.get('route',''), rule.get('note','')]
        bgs   = [bg1, bg2, bg2]
        bolds = [False, True, False]
        iters = [True, False, False]
        row = table.rows[i+1]
        for j, (val, bg, bold_, italic_) in enumerate(zip(vals, bgs, bolds, iters)):
            cell = row.cells[j]
            set_col_width(cell, C[j])
            set_cell_bg(cell, bg)
            set_cell_borders(cell, GRAY2_H)
            set_cell_margins(cell, 80, 80, 120, 80)
            p = cell.paragraphs[0]; pspacing(p, 0, 0)
            run(p, val, bold=bold_, italic=italic_, size=9)

    add_spacer(doc)

def add_kv_table(doc, rows):
    """Two-column label/value info table"""
    table = doc.add_table(rows=len(rows), cols=2)
    set_table_width(table, 9360)
    for i, (k, v) in enumerate(rows):
        bg = GRAY_H if i%2==0 else WHITE_H
        c0, c1 = table.rows[i].cells
        set_col_width(c0, 2800); set_col_width(c1, 6560)
        set_cell_bg(c0, GRAY_H); set_cell_borders(c0, GRAY2_H)
        set_cell_bg(c1, bg);     set_cell_borders(c1, GRAY2_H)
        set_cell_margins(c0, 80,80,120,80)
        set_cell_margins(c1, 80,80,120,120)
        p0 = c0.paragraphs[0]; pspacing(p0,0,0)
        run(p0, str(k), bold=True, color=NAVY, size=10)
        p1 = c1.paragraphs[0]; pspacing(p1,0,0)
        run(p1, str(v), size=10)
    add_spacer(doc)

# ── Content Renderer ──────────────────────────────────────
def render(doc, content):
    lines = content.split('\n')
    n = len(lines)
    i = 0
    route_rules = []
    faq_blocks = []   # collect (title, q, a)
    in_section5 = False

    def flush_routing():
        nonlocal route_rules
        if route_rules:
            add_routing_table(doc, route_rules)
            route_rules = []

    while i < n:
        line = lines[i]
        s = line.strip()

        if not s:
            i += 1; continue

        # ── Major section headers ──────────────────────
        m = re.match(r'^#{1,2}\s+(\d+)\.\s+(.+)', s)
        if m:
            flush_routing()
            num = int(m.group(1))
            text = m.group(2).replace('**','')
            if num > 1:
                doc.add_page_break()
            add_h1(doc, f"{num}. {text}")
            in_section5 = (num == 5)
            i += 1; continue

        # ── Sub-section ────────────────────────────────
        if re.match(r'^#{2,4}\s+', s) and not re.match(r'^#{1,2}\s+\d+\.', s):
            flush_routing()
            text = re.sub(r'^#+\s+', '', s).replace('**','')
            # FAQ entry heading
            if re.match(r'^FAQ\s+\d+', text, re.I):
                add_h3(doc, text)
            else:
                add_h2(doc, text)
            i += 1; continue

        # ── Code block ─────────────────────────────────
        if s.startswith('```'):
            code = []
            i += 1
            while i < n and not lines[i].strip().startswith('```'):
                code.append(lines[i])
                i += 1
            i += 1
            t = '\n'.join(code).strip()
            if t: add_copy_box(doc, t)
            continue

        # ── Divider ────────────────────────────────────
        if re.match(r'^-{3,}$', s):
            i += 1; continue

        # ── FAQ structured block ────────────────────────
        # **FAQ N: Title**
        faq_h = re.match(r'^\*\*FAQ\s+(\d+)[:\s]+(.+?)\*\*$', s, re.I)
        if faq_h:
            title = f"FAQ {faq_h.group(1)}: {faq_h.group(2)}"
            # Look ahead for Question variants and Answer
            i += 1
            q_line = ''; a_line = ''
            while i < n:
                ls = lines[i].strip()
                qm = re.match(r'^Question variants?[:\s]+(.+)', ls, re.I)
                am = re.match(r'^Answer[:\s]+"?(.+)"?$', ls, re.I)
                am2 = re.match(r'^\*\*Answer[^:]*:\*\*\s*"?(.+)"?$', ls, re.I)
                if qm:
                    q_line = qm.group(1).strip().strip('"')
                    i += 1; continue
                if am:
                    a_line = am.group(1).strip().strip('"')
                    i += 1; break
                if am2:
                    a_line = am2.group(1).strip().strip('"')
                    i += 1; break
                if re.match(r'^---$', ls) or re.match(r'^\*\*FAQ', ls) or re.match(r'^#{2,}', ls):
                    break
                i += 1
            if q_line and a_line:
                add_faq_table(doc, title, q_line, a_line)
            elif title:
                add_h3(doc, title)
            continue

        # ── Transfer by Context rule ────────────────────
        rule_m = re.match(r'^\*\*Rule\s+\d+[:\s]+(.+?)\*\*', s)
        if rule_m:
            rule = {'title': rule_m.group(1), 'keywords':'', 'route':'', 'note':''}
            i += 1
            while i < n:
                ls = lines[i].strip()
                km = re.match(r'^Keywords?[:\s]+(.+)', ls, re.I)
                rm = re.match(r'^Route\s+to[:\s]+(.+)', ls, re.I)
                nm = re.match(r'^(?:Covers|Note|AIR holding message)[:\s]+(.+)', ls, re.I)
                if km:
                    rule['keywords'] = km.group(1).replace('**','').strip()
                    i += 1; continue
                if rm:
                    rule['route'] = rm.group(1).replace('**','').strip()
                    i += 1; continue
                if nm:
                    rule['note'] = nm.group(1).replace('**','').strip().strip('"')
                    i += 1; continue
                if re.match(r'^\*\*Rule', ls) or re.match(r'^#{2,}', ls) or re.match(r'^---$', ls):
                    break
                i += 1
            route_rules.append(rule)
            continue

        # ── Callouts ────────────────────────────────────
        if s.startswith('⚠️'):
            t = re.sub(r'^⚠️\s*', '', s).replace('**','')
            m2 = re.match(r'\*\*(.+?)\*\*[:\s]+(.+)', s[2:])
            if m2: warn_box(doc, m2.group(1), m2.group(2).replace('**',''))
            else: warn_box(doc, 'Note', t)
            i += 1; continue

        if s.startswith('❌') or s.startswith('🚫'):
            t = re.sub(r'^[❌🚫]\s*', '', s).replace('**','')
            danger_box(doc, 'Must Never', t)
            i += 1; continue

        # ── Bullets ─────────────────────────────────────
        if re.match(r'^[-*]\s+', s):
            text = re.sub(r'^[-*]\s+', '', s)
            bm = re.match(r'\*\*(.+?)\*\*[:\s—-]+(.+)', text)
            if bm:
                add_bullet(doc, bm.group(2).replace('**',''), bm.group(1))
            else:
                add_bullet(doc, text.replace('**',''))
            i += 1; continue

        # ── Numbered ────────────────────────────────────
        if re.match(r'^\d+\.\s+', s):
            text = re.sub(r'^\d+\.\s+', '', s).replace('**','')
            add_numbered(doc, text)
            i += 1; continue

        # ── Bold label: value ────────────────────────────
        bkv = re.match(r'^\*\*(.+?)\*\*[:\s]+(.+)', s)
        if bkv:
            k = bkv.group(1).strip()
            v = bkv.group(2).replace('**','').strip()
            add_label_value(doc, k, v)
            i += 1; continue

        # ── Bold only ────────────────────────────────────
        if re.match(r'^\*\*.+\*\*$', s):
            add_body(doc, s.replace('**',''), bold=True, color=NAVY)
            i += 1; continue

        # ── Markdown table ───────────────────────────────────
        if s.startswith('|') and s.endswith('|'):
            tbl_lines = []
            while i < n and lines[i].strip().startswith('|'):
                row_line = lines[i].strip()
                # skip separator rows like |---|---|
                if not re.match(r'^[|\s\-]+$', row_line):
                    tbl_lines.append(row_line)
                i += 1
            if len(tbl_lines) >= 1:
                def parse_row(line):
                    cells = [c.strip() for c in line.strip('|').split('|')]
                    return [re.sub(r'\*\*(.+?)\*\*', r'\1', c).strip() for c in cells]
                rows = [parse_row(l) for l in tbl_lines]
                if not rows:
                    continue
                headers = rows[0]
                data = rows[1:] if len(rows) > 1 else []
                
                # Coverage plan: 4 cols, # and Handles
                if len(headers) == 4 and headers[0].strip() in ('#','') and any('Handles' in h for h in headers):
                    coverage_rows = [(r[0],r[1],r[2],r[3]) for r in data if len(r)==4]
                    if coverage_rows: add_coverage_table(doc, coverage_rows)
                
                # FAQ: 2 cols, QUESTION header
                elif len(headers) == 2 and 'QUESTION' in headers[0].upper():
                    for r in data:
                        if len(r) == 2: add_faq_table(doc, '', r[0], r[1])
                
                # Routing: 3 cols, TRIGGER/KEYWORD header
                elif len(headers) == 3 and any(k in headers[0].upper() for k in ['TRIGGER','KEYWORD']):
                    rules = [{'keywords':r[0],'route':r[1],'note':r[2]} for r in data if len(r)==3]
                    if rules: add_routing_table(doc, rules)
                
                # KV table: 2 cols
                elif len(headers) == 2:
                    kv = [(r[0],r[1]) for r in data if len(r)==2]
                    if kv: add_kv_table(doc, kv)
                    elif len(rows)==1: add_copy_box(doc, rows[0][0])
                
                # Fallback
                else:
                    for r in data:
                        if r: add_body(doc, ' | '.join(r))
            continue

        # ── Plain text ───────────────────────────────────
        text = s.replace('**','').replace('*','')
        if text:
            add_body(doc, text)
        i += 1

    flush_routing()


# ── Main ─────────────────────────────────────────────────
def generate_docx(biz_name, content, prepared_by="RingCentral SE"):
    doc = Document()
    for section in doc.sections:
        section.top_margin = section.bottom_margin = Inches(1)
        section.left_margin = section.right_margin = Inches(1)

    doc.styles['Normal'].font.name = 'Arial'
    doc.styles['Normal'].font.size = Pt(10)

    today = datetime.now().strftime('%B %Y')

    # Cover
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pspacing(p, 480, 80)
    run(p, "RingCentral AI Receptionist", bold=True, color=NAVY, size=26)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pspacing(p, 0, 80)
    run(p, "Configuration Playbook", color=BLUE, size=20)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pspacing(p, 80, 40)
    run(p, biz_name, bold=True, color=SLATE, size=16)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pspacing(p, 20, 20)
    run(p, "Ready-to-paste configuration for every AIR field", italic=True, color=GRAY4, size=10)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pspacing(p, 20, 20)
    run(p, f"Prepared by: {prepared_by}  |  {today}", italic=True, color=GRAY4, size=9)

    doc.add_page_break()

    render(doc, content)

    # Footer
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pspacing(p, 240, 60)
    para_border_bottom(p, GRAY2_H, 4)
    run(p, f"RingCentral AI Receptionist  |  {biz_name}  |  {today}",
        italic=True, color=GRAY4, size=8)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()

# ── Coverage Plan Table (YES/PARTIAL/ROUTE color coding) ──
def add_coverage_table(doc, rows):
    """4-column table: # | Call Reason | AIR Handles? | How"""
    C = [360, 3000, 2400, 3600]
    table = doc.add_table(rows=1+len(rows), cols=4)
    set_table_width(table, sum(C))
    
    headers = ["#", "Call Reason", "AIR Handles?", "How"]
    hr = table.rows[0]
    for j, (cell, hdr, w) in enumerate(zip(hr.cells, headers, C)):
        set_col_width(cell, w)
        set_cell_bg(cell, NAVY_H)
        set_cell_borders(cell, NAVY_H)
        set_cell_margins(cell, 80,80,120,80)
        p = cell.paragraphs[0]; pspacing(p,0,0)
        run(p, hdr, bold=True, color=WHITE, size=9)
    
    for i, row_data in enumerate(rows):
        bg = GRAY_H if i%2==0 else WHITE_H
        row = table.rows[i+1]
        num, reason, handles, how = row_data
        
        # Determine color for "AIR Handles?" cell
        handles_upper = handles.upper()
        if 'YES' in handles_upper:
            handles_bg = "1A5C2B"  # dark green
            handles_color = RGBColor(0xFF,0xFF,0xFF)
        elif 'PARTIAL' in handles_upper:
            handles_bg = "7C4A00"  # dark amber  
            handles_color = RGBColor(0xFF,0xFF,0xFF)
        else:  # ROUTE
            handles_bg = "8B1A1A"  # dark red
            handles_color = RGBColor(0xFF,0xFF,0xFF)
        
        vals = [str(num), reason, handles, how]
        bgs  = [bg, bg, handles_bg, bg]
        colors = [None, None, handles_color, None]
        bolds = [False, False, True, False]
        
        for j, (cell, val, bg_, col, bld) in enumerate(zip(row.cells, vals, bgs, colors, bolds)):
            set_col_width(cell, C[j])
            set_cell_bg(cell, bg_)
            set_cell_borders(cell, GRAY2_H)
            set_cell_margins(cell, 80,80,120,80)
            p = cell.paragraphs[0]; pspacing(p,0,0)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if j==0 else WD_ALIGN_PARAGRAPH.LEFT
            run(p, val, bold=bld, color=col, size=9)
    
    add_spacer(doc)

def add_dark_copy_box(doc, text):
    """Dark navy box with white italic text — for greetings"""
    table = doc.add_table(rows=1, cols=1)
    set_table_width(table, 9360)
    cell = table.cell(0,0)
    set_col_width(cell, 9360)
    set_cell_bg(cell, "1E3A5F")
    set_cell_borders(cell, "2D4E78", "4")
    set_cell_margins(cell, 160, 160, 200, 200)
    p = cell.paragraphs[0]
    pspacing(p, 0, 0)
    r = p.add_run(text)
    r.font.name = 'Arial'
    r.font.size = Pt(10)
    r.font.color.rgb = WHITE
    r.italic = True
    add_spacer(doc)
