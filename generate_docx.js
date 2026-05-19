#!/usr/bin/env node
// Called by Flask: node generate_docx.js <json_input_file> <output_file>
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageBreak
} = require('docx');
const fs = require('fs');

const input = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
const { bizName, content, preparedBy } = input;

// ── Colors ──────────────────────────────────────────────
const NAVY   = "1E3A5F";
const BLUE   = "0F62FE";
const GREEN  = "166534";
const GBKG   = "DCFCE7";
const GBDR   = "BBF7D0";
const RED    = "991B1B";
const RBKG   = "FEE2E2";
const RBDR   = "FECACA";
const AMBER  = "92400E";
const ABKG   = "FEF3C7";
const ABDR   = "FDE68A";
const GRAY   = "F3F4F6";
const GRAY2  = "E2E8F0";
const WHITE  = "FFFFFF";
const BLUE_LT= "EFF6FF";
const BLUE_MD= "DBEAFE";

// ── Helpers ──────────────────────────────────────────────
const bdr = (c="CCCCCC") => ({ style: BorderStyle.SINGLE, size: 1, color: c });
const borders = (c="CCCCCC") => ({ top:bdr(c), bottom:bdr(c), left:bdr(c), right:bdr(c) });
const noBdr = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top:noBdr, bottom:noBdr, left:noBdr, right:noBdr };

function cell(text, opts={}) {
  const w = opts.w || 4680;
  return new TableCell({
    borders: opts.noBorder ? noBorders : borders(opts.bdrColor || "CCCCCC"),
    width: { size: w, type: WidthType.DXA },
    shading: { fill: opts.fill || WHITE, type: ShadingType.CLEAR },
    margins: { top: opts.topMar||100, bottom: opts.botMar||100, left: 140, right: 140 },
    children: [new Paragraph({
      alignment: opts.align || AlignmentType.LEFT,
      children: [new TextRun({
        text: String(text||""),
        bold: opts.bold||false,
        size: opts.size||20,
        color: opts.color||"000000",
        font: "Arial",
        italics: opts.italics||false,
      })]
    })]
  });
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 160 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE, space: 4 } },
    children: [new TextRun({ text, bold: true, size: 40, color: NAVY, font: "Arial" })]
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 280, after: 120 },
    children: [new TextRun({ text, bold: true, size: 28, color: NAVY, font: "Arial" })]
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, bold: true, size: 24, color: BLUE, font: "Arial" })]
  });
}
function body(text, opts={}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({ text, size: 20, font: "Arial", color: opts.color||"1E293B", italics: opts.italics||false, bold: opts.bold||false })]
  });
}
function bullet(text, boldPre) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: boldPre
      ? [new TextRun({ text: boldPre+": ", bold: true, size: 20, font: "Arial" }),
         new TextRun({ text, size: 20, font: "Arial" })]
      : [new TextRun({ text, size: 20, font: "Arial" })]
  });
}
function sp(n=1) {
  return Array.from({length:n}, () => new Paragraph({ spacing:{before:40,after:40}, children:[new TextRun("")] }));
}
function pb() { return new Paragraph({ children: [new PageBreak()] }); }

function callout(icon, label, bodyText, bg, textColor) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [520, 8840],
    rows: [new TableRow({ children: [
      new TableCell({
        borders: noBorders,
        width: { size: 520, type: WidthType.DXA },
        shading: { fill: bg, type: ShadingType.CLEAR },
        margins: { top:120, bottom:120, left:120, right:60 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: icon, size: 22, font: "Arial", color: textColor })] })]
      }),
      new TableCell({
        borders: noBorders,
        width: { size: 8840, type: WidthType.DXA },
        shading: { fill: bg, type: ShadingType.CLEAR },
        margins: { top:120, bottom:120, left:100, right:140 },
        children: [new Paragraph({ children: [
          new TextRun({ text: label+":  ", bold: true, size: 20, font: "Arial", color: textColor }),
          new TextRun({ text: bodyText, size: 20, font: "Arial", color: textColor })
        ]})]
      })
    ]})]
  });
}

function faqTable(question, answer) {
  const C1=3400, C2=5960;
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [C1, C2],
    rows: [
      new TableRow({ children: [
        cell("QUESTION (what caller says)", { w:C1, bold:true, fill:NAVY, color:WHITE, size:18 }),
        cell("ANSWER (what AIR says — copy exactly)", { w:C2, bold:true, fill:NAVY, color:WHITE, size:18 }),
      ]}),
      new TableRow({ children: [
        cell(question, { w:C1, fill:BLUE_LT, italics:true, size:18 }),
        cell(answer,   { w:C2, fill:GRAY,    size:18 }),
      ]}),
    ]
  });
}

function routeTable(rows) {
  const C=[3200, 2160, 4000];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: C,
    rows: [
      new TableRow({ children: [
        cell("TRIGGER KEYWORDS / PHRASES", { w:C[0], bold:true, fill:NAVY, color:WHITE, size:18 }),
        cell("ROUTE TO",                   { w:C[1], bold:true, fill:NAVY, color:WHITE, size:18 }),
        cell("NOTES",                      { w:C[2], bold:true, fill:NAVY, color:WHITE, size:18 }),
      ]}),
      ...rows.map((r,i) => new TableRow({ children: [
        cell(r[0], { w:C[0], fill: i%2===0?"EFF6FF":WHITE, italics:true, size:18 }),
        cell(r[1], { w:C[1], fill: i%2===0?GRAY:WHITE, bold:true, size:18 }),
        cell(r[2], { w:C[2], fill: i%2===0?GRAY:WHITE, size:18 }),
      ]}))
    ]
  });
}

// ── Parse sections from markdown-style text ──────────────
function parseSections(text) {
  const lines = text.split('\n');
  const sections = {};
  let current = null;
  let buf = [];

  for (const line of lines) {
    const h1 = line.match(/^#{1,2}\s+\d+\.\s+(.+)/);
    if (h1) {
      if (current) sections[current] = buf.join('\n').trim();
      current = h1[1].trim();
      buf = [];
    } else if (current) {
      buf.push(line);
    }
  }
  if (current) sections[current] = buf.join('\n').trim();
  return sections;
}

// ── Extract subsections ──────────────────────────────────
function extractFAQs(text) {
  const faqs = [];
  // Match ### FAQ #N: Title ... Question variants ... Answer
  const blocks = text.split(/\n###\s+/);
  for (const block of blocks.slice(1)) {
    const titleMatch = block.match(/^(?:FAQ\s+[#\d]+[:\.]?\s*)?(.+)/);
    const title = titleMatch ? titleMatch[1].trim() : 'FAQ';
    const qMatch = block.match(/Question variants[^:]*:\s*([\s\S]*?)(?=\*\*Answer|Answer[\s(])/i);
    const aMatch = block.match(/\*\*Answer[^:]*:\*\*\s*([\s\S]*?)(?=\n---|\n###|$)/i) ||
                   block.match(/Answer[^:]*:\s*"([\s\S]*?)"/i) ||
                   block.match(/Answer[^:]*:\s*([\s\S]*?)(?=\n---|\n###|$)/i);
    if (qMatch && aMatch) {
      const questions = qMatch[1].replace(/^[-*"'\s]+/gm, '').replace(/["']\s*$/gm, '').trim();
      const answer = aMatch[1].replace(/^["'\s*]+/, '').replace(/["'\s*]+$/, '').trim();
      faqs.push({ title, questions, answer });
    }
  }
  return faqs;
}

function extractTransferRules(text) {
  const rules = [];
  const blocks = text.split(/\n\*\*Rule\s+/);
  for (const block of blocks.slice(1)) {
    const titleMatch = block.match(/^[#\d]*[:\.]?\s*(.+)/);
    const title = titleMatch ? titleMatch[1].replace(/\*\*/g,'').trim() : 'Rule';
    const kwMatch = block.match(/Keywords?:\s*(.+?)(?=\nRoute|\nCovers|\n\*\*|$)/is);
    const routeMatch = block.match(/Route\s+to:\s*(.+?)(?=\nCovers|\nKeywords|\n\*\*|$)/is);
    const noteMatch = block.match(/(?:Covers|AIR holding|Note):\s*(.+?)(?=\n\*\*Rule|\n---|\n##|$)/is);
    if (kwMatch && routeMatch) {
      rules.push({
        title,
        keywords: kwMatch[1].replace(/\*\*/g,'').trim(),
        route: routeMatch[1].replace(/\*\*/g,'').trim(),
        note: noteMatch ? noteMatch[1].replace(/\*\*/g,'').trim() : ''
      });
    }
  }
  return rules;
}

// ── Build document children ──────────────────────────────
function buildDoc(bizName, content, preparedBy) {
  const today = new Date().toLocaleDateString('en-US', {month:'long', year:'numeric'});
  const children = [];

  // ── COVER PAGE ────────────────────────────────────────
  children.push(...sp(2));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before:480, after:60 },
    children: [new TextRun({ text: "RingCentral AI Receptionist", bold:true, size:52, color:NAVY, font:"Arial" })]
  }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before:0, after:60 },
    children: [new TextRun({ text: "Configuration Playbook", size:40, color:BLUE, font:"Arial" })]
  }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before:60, after:40 },
    children: [new TextRun({ text: bizName, bold:true, size:32, color:"475569", font:"Arial" })]
  }));
  children.push(...sp(1));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before:40, after:40 },
    children: [new TextRun({ text: "Ready-to-paste configuration for every AIR field", size:20, italics:true, color:"94A3B8", font:"Arial" })]
  }));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before:20, after:20 },
    children: [new TextRun({ text: `Prepared by: ${preparedBy||"RingCentral SE"}  |  ${today}`, size:18, italics:true, color:"94A3B8", font:"Arial" })]
  }));
  children.push(...sp(3));

  // ── MAIN CONTENT ──────────────────────────────────────
  // Split content into lines and render section by section
  const lines = content.split('\n');
  let inFAQSection = false;
  let inTransferSection = false;
  let inCodeBlock = false;
  let codeLines = [];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Page break on major sections
    if (line.match(/^#{1,2}\s+\d+\./)) {
      if (i > 0) children.push(pb());
    }

    // Headings
    if (line.match(/^# /)) {
      const text = line.replace(/^#+\s+/, '').replace(/\*\*/g,'');
      children.push(h1(text));
      inFAQSection = text.toLowerCase().includes('faq');
      inTransferSection = text.toLowerCase().includes('transfer');
    } else if (line.match(/^## /)) {
      const text = line.replace(/^#+\s+/, '').replace(/\*\*/g,'');
      children.push(h2(text));
    } else if (line.match(/^### /)) {
      const text = line.replace(/^#+\s+/, '').replace(/\*\*/g,'');
      children.push(h3(text));
    }
    // Code blocks (company description etc)
    else if (line.startsWith('```')) {
      if (!inCodeBlock) {
        inCodeBlock = true;
        codeLines = [];
      } else {
        inCodeBlock = false;
        const codeText = codeLines.join('\n');
        // Render as a single-cell table (copy-box style)
        children.push(new Table({
          width: { size: 9360, type: WidthType.DXA },
          columnWidths: [9360],
          rows: [new TableRow({ children: [new TableCell({
            borders: borders("BFDBFE"),
            width: { size: 9360, type: WidthType.DXA },
            shading: { fill: BLUE_MD, type: ShadingType.CLEAR },
            margins: { top:160, bottom:160, left:200, right:200 },
            children: [new Paragraph({ children: [new TextRun({ text: codeText, size:20, font:"Arial", color:"1E3A5F" })] })]
          })] })]
        }));
        children.push(...sp(1));
        codeLines = [];
      }
    } else if (inCodeBlock) {
      codeLines.push(line);
    }
    // Horizontal rules → spacer
    else if (line.match(/^---+$/)) {
      children.push(...sp(1));
    }
    // Bullet points
    else if (line.match(/^[-*]\s+/)) {
      const text = line.replace(/^[-*]\s+/, '').replace(/\*\*/g,'');
      const boldMatch = text.match(/^\*\*(.+?)\*\*[:\s—-]+(.+)/);
      if (boldMatch) children.push(bullet(boldMatch[2], boldMatch[1]));
      else children.push(bullet(text));
    }
    // Numbered lists
    else if (line.match(/^\d+\.\s+/)) {
      const text = line.replace(/^\d+\.\s+/, '').replace(/\*\*/g,'');
      children.push(new Paragraph({
        numbering: { reference: "numbers", level: 0 },
        spacing: { before:40, after:40 },
        children: [new TextRun({ text, size:20, font:"Arial" })]
      }));
    }
    // Bold labels (field: value pattern)
    else if (line.match(/^\*\*[^*]+\*\*[:\s]/)) {
      const boldMatch = line.match(/^\*\*(.+?)\*\*[:\s]+(.+)/);
      if (boldMatch) {
        children.push(new Paragraph({
          spacing: { before:60, after:60 },
          children: [
            new TextRun({ text: boldMatch[1]+": ", bold:true, size:20, font:"Arial", color:NAVY }),
            new TextRun({ text: boldMatch[2].replace(/\*\*/g,''), size:20, font:"Arial" })
          ]
        }));
      } else {
        const text = line.replace(/\*\*/g,'');
        children.push(body(text, { bold:true }));
      }
    }
    // ⚠️ warning / callout lines
    else if (line.match(/^⚠️|^⛔|^🔴/)) {
      const text = line.replace(/^[⚠️⛔🔴]\s*/, '');
      children.push(...sp(0));
      children.push(callout("⚠️", "Note", text.replace(/\*\*/g,''), ABKG, AMBER));
      children.push(...sp(1));
    }
    // ❌ prohibition items
    else if (line.match(/^❌/)) {
      const text = line.replace(/^❌\s*/, '').replace(/\*\*/g,'');
      children.push(bullet("❌ " + text));
    }
    // ✅ positive items
    else if (line.match(/^✅/)) {
      const text = line.replace(/^✅\s*/, '').replace(/\*\*/g,'');
      children.push(bullet("✓ " + text));
    }
    // Empty lines
    else if (line.trim() === '') {
      // skip — handled by spacing
    }
    // Regular body text
    else if (line.trim()) {
      const text = line.replace(/\*\*/g,'').replace(/^\*|\*$/g,'');
      children.push(body(text));
    }
  }

  // Footer
  children.push(...sp(2));
  children.push(new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before:240, after:60 },
    border: { top: { style: BorderStyle.SINGLE, size: 2, color: "CBD5E1", space: 6 } },
    children: [new TextRun({ text: `RingCentral AI Receptionist Configuration Playbook  |  ${bizName}  |  ${today}`, size:16, italics:true, color:"94A3B8", font:"Arial" })]
  }));

  return children;
}

const children = buildDoc(bizName, content, preparedBy);

const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level:0, format:LevelFormat.BULLET, text:"•", alignment:AlignmentType.LEFT,
          style: { paragraph: { indent: { left:720, hanging:360 } } } }] },
      { reference: "numbers", levels: [{ level:0, format:LevelFormat.DECIMAL, text:"%1.", alignment:AlignmentType.LEFT,
          style: { paragraph: { indent: { left:720, hanging:360 } } } }] },
    ]
  },
  styles: {
    default: { document: { run: { font:"Arial", size:20 } } },
    paragraphStyles: [
      { id:"Heading1", name:"Heading 1", basedOn:"Normal", next:"Normal", quickFormat:true,
        run: { size:40, bold:true, font:"Arial", color:NAVY },
        paragraph: { spacing:{before:360,after:160}, outlineLevel:0 } },
      { id:"Heading2", name:"Heading 2", basedOn:"Normal", next:"Normal", quickFormat:true,
        run: { size:28, bold:true, font:"Arial", color:NAVY },
        paragraph: { spacing:{before:280,after:120}, outlineLevel:1 } },
      { id:"Heading3", name:"Heading 3", basedOn:"Normal", next:"Normal", quickFormat:true,
        run: { size:24, bold:true, font:"Arial", color:BLUE },
        paragraph: { spacing:{before:200,after:80}, outlineLevel:2 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width:12240, height:15840 },
        margin: { top:1260, right:1260, bottom:1260, left:1260 }
      }
    },
    children
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync(process.argv[3], buf);
  console.log('OK');
}).catch(err => {
  console.error(err);
  process.exit(1);
});
