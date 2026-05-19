"""
Generate a formatted Word doc from AIR configuration text.
Called directly from app.py — no Node.js required.
"""
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re
import io

# Colors
NAVY   = RGBColor(0x1E, 0x3A, 0x5F)
BLUE   = RGBColor(0x0F, 0x62, 0xFE)
GRAY5  = RGBColor(0x47, 0x55, 0x69)
TEXT2  = RGBColor(0x33, 0x41, 0x55)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)

def hex_to_rgb_str(r, g, b):
    return f"{r:02X}{g:02X}{b:02X}"

NAVY_HEX  = "1E3A5F"
BLUE_HEX  = "0F62FE"
BLUE_LT   = "EFF6FF"
BLUE_MD   = "DBEAFE"
GRAY_HEX  = "F3F4F6"
GRAY2_HEX = "E2E8F0"
RED_LT    = "FEE2E2"
AMB_LT    = "FEF3C7"
GRN_LT    = "DCFCE7"
WHITE_HEX = "FFFFFF"


def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


def set_cell_borders(cell, color="CCCCCC"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ['top', 'left', 'bottom', 'right']:
        border = OxmlElement(f'w:{side}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), color)
        tcBorders.append(border)
    tcPr.append(tcBorders)


def set_paragraph_border_bottom(para, color=BLUE_HEX):
    pPr = para._p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    bottom = OxmlElement('w:bottom')
    bottom.set(qn('w:val'), 'single')
    bottom.set(qn('w:sz'), '6')
    bottom.set(qn('w:space'), '4')
    bottom.set(qn('w:color'), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def add_run(para, text, bold=False, italic=False, color=None, size=11):
    run = para.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color
    return run


def generate_docx(biz_name, content, prepared_by="RingCentral SE"):
    doc = Document()

    # Page margins — 1 inch all around
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Default font
    doc.styles['Normal'].font.name = 'Arial'
    doc.styles['Normal'].font.size = Pt(11)

    lines = content.split('\n')
    i = 0

    # ── COVER PAGE ─────────────────────────────────────────
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(48)
    add_run(p, "RingCentral AI Receptionist", bold=True, color=NAVY, size=26)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(p, "Configuration Playbook", color=BLUE, size=20)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(12)
    add_run(p, biz_name, bold=True, color=GRAY5, size=16)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    add_run(p, f"Prepared by: {prepared_by}", italic=True, color=GRAY5, size=10)

    doc.add_page_break()

    # ── PARSE AND RENDER CONTENT ───────────────────────────
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # Skip empty
        if not stripped:
            i += 1
            continue

        # H1 — ## 1. or # 1.
        if re.match(r'^#{1,2}\s+\d+\.', stripped):
            text = re.sub(r'^#+\s+', '', stripped).replace('**', '')
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(18)
            p.paragraph_format.space_after = Pt(6)
            set_paragraph_border_bottom(p, BLUE_HEX)
            add_run(p, text, bold=True, color=NAVY, size=16)
            i += 1
            continue

        # H2 — ### or ##
        if re.match(r'^#{2,3}\s+', stripped):
            text = re.sub(r'^#+\s+', '', stripped).replace('**', '')
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(4)
            add_run(p, text, bold=True, color=NAVY, size=13)
            i += 1
            continue

        # H4 — ####
        if re.match(r'^#{4,}\s+', stripped):
            text = re.sub(r'^#+\s+', '', stripped).replace('**', '')
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            add_run(p, text, bold=True, color=BLUE, size=11)
            i += 1
            continue

        # Code block — render as blue-tinted box
        if stripped.startswith('```'):
            code_lines = []
            i += 1
            while i < len(lines) and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            i += 1  # skip closing ```
            code_text = '\n'.join(code_lines).strip()
            if code_text:
                table = doc.add_table(rows=1, cols=1)
                table.style = 'Table Grid'
                cell = table.cell(0, 0)
                set_cell_bg(cell, BLUE_MD)
                set_cell_borders(cell, "BFDBFE")
                p = cell.paragraphs[0]
                run = p.add_run(code_text)
                run.font.name = 'Arial'
                run.font.size = Pt(10)
                run.font.color.rgb = NAVY
                doc.add_paragraph()
            continue

        # Horizontal rule
        if re.match(r'^-{3,}$', stripped):
            i += 1
            continue

        # Bullet point
        if re.match(r'^[-*]\s+', stripped):
            text = re.sub(r'^[-*]\s+', '', stripped)
            text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
            p = doc.add_paragraph(style='List Bullet')
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            # Check for bold prefix pattern
            bold_match = re.match(r'\*\*(.+?)\*\*[:\s—-]+(.+)', stripped[2:])
            if bold_match:
                add_run(p, bold_match.group(1) + ': ', bold=True, size=10)
                add_run(p, bold_match.group(2), size=10)
            else:
                add_run(p, text, size=10)
            i += 1
            continue

        # Numbered list
        if re.match(r'^\d+\.\s+', stripped):
            text = re.sub(r'^\d+\.\s+', '', stripped).replace('**', '')
            p = doc.add_paragraph(style='List Number')
            p.paragraph_format.space_before = Pt(2)
            add_run(p, text, size=10)
            i += 1
            continue

        # ⚠️ warning callout
        if stripped.startswith('⚠️') or stripped.startswith('⛔'):
            text = re.sub(r'^[⚠️⛔]\s*', '', stripped).replace('**', '')
            table = doc.add_table(rows=1, cols=2)
            table.style = 'Table Grid'
            c0, c1 = table.cell(0, 0), table.cell(0, 1)
            c0.width = Emu(457200)  # ~0.5 inch
            set_cell_bg(c0, AMB_LT)
            set_cell_borders(c0, "FDE68A")
            set_cell_bg(c1, AMB_LT)
            set_cell_borders(c1, "FDE68A")
            p0 = c0.paragraphs[0]
            p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_run(p0, "⚠️", size=11)
            p1 = c1.paragraphs[0]
            add_run(p1, text, size=10, color=RGBColor(0x92, 0x40, 0x0E))
            doc.add_paragraph()
            i += 1
            continue

        # Bold label: value pattern
        bold_label = re.match(r'^\*\*(.+?)\*\*[:\s]+(.+)', stripped)
        if bold_label:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            add_run(p, bold_label.group(1) + ': ', bold=True, color=NAVY, size=10)
            add_run(p, bold_label.group(2).replace('**', ''), size=10)
            i += 1
            continue

        # Bold only line
        if re.match(r'^\*\*(.+?)\*\*$', stripped):
            text = stripped.replace('**', '')
            p = doc.add_paragraph()
            add_run(p, text, bold=True, color=NAVY, size=11)
            i += 1
            continue

        # Regular body text
        text = stripped.replace('**', '')
        if text:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            add_run(p, text, size=10)
        i += 1

    # Footer line
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(24)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_run(p, f"RingCentral AIR Configuration  |  {biz_name}", italic=True,
            color=RGBColor(0x94, 0xA3, 0xB8), size=8)

    # Return as bytes
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()
