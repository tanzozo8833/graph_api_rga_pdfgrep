"""Generate graph_grep_en.pptx and docs_report_en.docx — English documentation.

Run: venv\\Scripts\\python.exe make_docs_en.py
"""

import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt as DocxPt
from docx.shared import RGBColor as DocxRGB

# ---------- Palette ----------
BLUE = RGBColor(0x00, 0x78, 0xD4)
DARK_BLUE = RGBColor(0x00, 0x47, 0x82)
GREEN = RGBColor(0x10, 0x7C, 0x41)
PURPLE = RGBColor(0x62, 0x64, 0xA7)
TEAL = RGBColor(0x03, 0x83, 0x87)
ORANGE = RGBColor(0xD8, 0x3B, 0x01)
DARK = RGBColor(0x32, 0x31, 0x30)
LIGHT_BG = RGBColor(0xF3, 0xF2, 0xF1)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
YELLOW = RGBColor(0xFF, 0xB9, 0x00)
LIGHT_BLUE = RGBColor(0xDE, 0xEC, 0xF9)
GRAY = RGBColor(0x60, 0x5E, 0x5C)


def fill(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color


def no_line(shape):
    shape.line.fill.background()


def set_text(tf, text, *, size=18, bold=False, color=DARK, align=PP_ALIGN.LEFT, font="Segoe UI"):
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = font


def add_para(tf, text, *, size=14, bold=False, color=DARK, align=PP_ALIGN.LEFT, font="Segoe UI"):
    p = tf.add_paragraph()
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = font


def blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def title_bar(slide, title, slide_w):
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, slide_w, Inches(0.85))
    fill(bar, BLUE)
    no_line(bar)
    tf = bar.text_frame
    tf.margin_left = Inches(0.5)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    set_text(tf, title, size=24, bold=True, color=WHITE)


def textbox(slide, left, top, width, height, text, **kw):
    box = slide.shapes.add_textbox(left, top, width, height)
    set_text(box.text_frame, text, **kw)
    return box


def card(slide, left, top, width, height, header, body, color):
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    fill(box, color)
    no_line(box)
    tf = box.text_frame
    tf.margin_left = Inches(0.18)
    tf.margin_right = Inches(0.18)
    tf.margin_top = Inches(0.15)
    tf.margin_bottom = Inches(0.15)
    tf.word_wrap = True
    set_text(tf, header, size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_para(tf, body, size=13, color=WHITE, align=PP_ALIGN.CENTER)
    return box


def arrow_right(slide, left, top, width, height, color=BLUE):
    a = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, left, top, width, height)
    fill(a, color)
    no_line(a)
    return a


def arrow_down(slide, left, top, width, height, color=BLUE):
    a = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, left, top, width, height)
    fill(a, color)
    no_line(a)
    return a


def make_pptx(path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    SW = prs.slide_width
    SH = prs.slide_height

    # ===== 1. Title =====
    s = blank_slide(prs)
    bg = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    fill(bg, BLUE)
    no_line(bg)
    side = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.4), SH)
    fill(side, YELLOW)
    no_line(side)
    textbox(s, Inches(1), Inches(2.4), Inches(11.5), Inches(1.2),
            "Microsoft 365 Keyword-Search Agent",
            size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    textbox(s, Inches(1), Inches(3.7), Inches(11.5), Inches(0.8),
            "Question answering over OneDrive — without a vector database",
            size=22, color=WHITE, align=PP_ALIGN.CENTER)
    textbox(s, Inches(1), Inches(5.0), Inches(11.5), Inches(0.5),
            "Inspired by 'Keyword search is all you need' (arxiv 2602.23368)",
            size=14, color=LIGHT_BG, align=PP_ALIGN.CENTER)
    textbox(s, Inches(1), Inches(6.5), Inches(11.5), Inches(0.4),
            "Microsoft Graph API  •  Azure OpenAI  •  LangChain Agent",
            size=12, color=LIGHT_BG, align=PP_ALIGN.CENTER)

    # ===== 2. Problem =====
    s = blank_slide(prs)
    title_bar(s, "The Problem — Traditional RAG is too complex", SW)
    items = [
        ("Chunking", "Splits documents — context lost mid-sentence or table", BLUE),
        ("Embedding", "Costly vector computation for every chunk", PURPLE),
        ("Vector DB", "Another service to deploy and maintain", TEAL),
        ("Exact match", "Can miss product codes / rare technical terms", ORANGE),
    ]
    for i, (h, t, c) in enumerate(items):
        x = Inches(0.6 + (i % 2) * 6.3)
        y = Inches(1.4 + (i // 2) * 2.7)
        card(s, x, y, Inches(6.0), Inches(2.3), h, t, c)

    # ===== 3. Solution =====
    s = blank_slide(prs)
    title_bar(s, "Our Approach — Agentic Keyword Search", SW)
    textbox(s, Inches(0.7), Inches(1.2), Inches(12), Inches(0.6),
            "Let the LLM drive the search — no fixed pipeline",
            size=20, bold=True, color=DARK_BLUE)
    feats = [
        ("LLM-generated keywords", "Synonyms, abbreviations, regex, OR-joined", BLUE),
        ("Microsoft Graph", "Built-in full-text index for OneDrive", GREEN),
        ("Download on demand", "Extract text via pdfplumber / docx / xlsx / pptx", PURPLE),
        ("Grep with context", "Regex multi-pattern, plus or minus N lines per hit", TEAL),
        ("Self-retry strategy", "LLM tries new keywords if nothing found", ORANGE),
    ]
    for i, (h, t, c) in enumerate(feats):
        x = Inches(0.6 + i * 2.5)
        y = Inches(2.5)
        card(s, x, y, Inches(2.4), Inches(2.7), h, t, c)
    textbox(s, Inches(0.7), Inches(5.7), Inches(12), Inches(0.6),
            "Microsoft already indexes your content — no need to rebuild it",
            size=16, color=GRAY, align=PP_ALIGN.CENTER)

    # ===== 4. Actors =====
    s = blank_slide(prs)
    title_bar(s, "Actors in the System", SW)

    card(s, Inches(0.5), Inches(1.3), Inches(2.4), Inches(1.1),
         "User", "Asks the question", DARK_BLUE)
    arrow_right(s, Inches(2.95), Inches(1.65), Inches(0.4), Inches(0.4))
    card(s, Inches(3.4), Inches(1.3), Inches(2.4), Inches(1.1),
         "Browser", "Chat panel + SSE", BLUE)
    arrow_right(s, Inches(5.85), Inches(1.65), Inches(0.4), Inches(0.4))
    card(s, Inches(6.3), Inches(1.3), Inches(2.4), Inches(1.1),
         "Flask", "/ask route + MSAL", PURPLE)
    arrow_right(s, Inches(8.75), Inches(1.65), Inches(0.4), Inches(0.4))
    card(s, Inches(9.2), Inches(1.3), Inches(3.5), Inches(1.1),
         "LangChain Agent", "Tool-use loop", ORANGE)

    arrow_down(s, Inches(10.7), Inches(2.5), Inches(0.4), Inches(0.5))

    card(s, Inches(2.0), Inches(3.3), Inches(3.3), Inches(1.1),
         "Azure OpenAI", "gpt-4.1-mini decides", BLUE)
    card(s, Inches(5.5), Inches(3.3), Inches(3.3), Inches(1.1),
         "Microsoft Graph", "Search and Download", GREEN)
    card(s, Inches(9.0), Inches(3.3), Inches(3.3), Inches(1.1),
         "Local Modules", "Extractors and Cache", TEAL)

    arrow_down(s, Inches(6.5), Inches(4.5), Inches(0.4), Inches(0.5))

    card(s, Inches(1.5), Inches(5.2), Inches(3.0), Inches(1.4),
         "graph_search", "Find files by keyword", BLUE)
    card(s, Inches(5.0), Inches(5.2), Inches(3.0), Inches(1.4),
         "fetch_file_text", "Download and extract", GREEN)
    card(s, Inches(8.5), Inches(5.2), Inches(3.0), Inches(1.4),
         "grep_context", "Regex with context window", PURPLE)

    # ===== 5. 3 Tools =====
    s = blank_slide(prs)
    title_bar(s, "The Agent's Three Tools", SW)

    cols = [
        ("graph_search", BLUE,
         ["Parameter: query (string)",
          "Calls POST /search/query",
          "Returns file list with snippets",
          "OR-joined: 'k1 OR k2 OR k3'"]),
        ("fetch_file_text", GREEN,
         ["Parameter: item_id",
          "Downloads bytes from Graph CDN",
          "Extracts: PDF / DOCX / XLSX / PPTX",
          "Caches text in RAM"]),
        ("grep_context", PURPLE,
         ["Parameters: item_id, patterns, plus/minus N",
          "Regex IGNORECASE on cached text",
          "Merges overlapping windows",
          "Returns plus/minus 5 lines per hit"]),
    ]
    for i, (h, c, bullets) in enumerate(cols):
        x = Inches(0.5 + i * 4.3)
        hdr = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  x, Inches(1.2), Inches(4.0), Inches(0.9))
        fill(hdr, c)
        no_line(hdr)
        set_text(hdr.text_frame, h, size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        body = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   x, Inches(2.2), Inches(4.0), Inches(4.5))
        fill(body, LIGHT_BG)
        no_line(body)
        btf = body.text_frame
        btf.margin_left = Inches(0.2)
        btf.margin_right = Inches(0.2)
        btf.margin_top = Inches(0.2)
        btf.word_wrap = True
        set_text(btf, "- " + bullets[0], size=14, color=DARK)
        for b in bullets[1:]:
            add_para(btf, "- " + b, size=14, color=DARK)

    # ===== 6. Workflow =====
    s = blank_slide(prs)
    title_bar(s, "End-to-End Workflow", SW)

    steps = [
        ("1", "User asks", "Types question in chat panel", BLUE),
        ("2", "Flask /ask", "Initializes agent with token", PURPLE),
        ("3", "LLM generates keywords", "OR-joined synonyms", ORANGE),
        ("4", "graph_search", "Microsoft Graph returns files + snippets", GREEN),
        ("5", "Enough?", "Are snippets sufficient?", TEAL),
        ("6", "fetch + grep", "Read deeper if needed", DARK_BLUE),
        ("7", "Final answer", "Synthesis with citations", BLUE),
    ]
    for i, (n, h, t, c) in enumerate(steps):
        x = Inches(0.3 + i * 1.85)
        y = Inches(1.7)
        num = s.shapes.add_shape(MSO_SHAPE.OVAL, x + Inches(0.55), y, Inches(0.6), Inches(0.6))
        fill(num, c)
        no_line(num)
        set_text(num.text_frame, n, size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        bx = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                 x, Inches(2.5), Inches(1.7), Inches(2.0))
        fill(bx, LIGHT_BG)
        no_line(bx)
        btf = bx.text_frame
        btf.margin_left = Inches(0.08)
        btf.margin_right = Inches(0.08)
        btf.margin_top = Inches(0.1)
        btf.word_wrap = True
        set_text(btf, h, size=13, bold=True, color=c, align=PP_ALIGN.CENTER)
        add_para(btf, t, size=10, color=DARK, align=PP_ALIGN.CENTER)
        if i < len(steps) - 1:
            arrow_right(s, x + Inches(1.72), y + Inches(0.18), Inches(0.13), Inches(0.25), color=GRAY)

    textbox(s, Inches(0.5), Inches(5.0), Inches(12.5), Inches(0.6),
            "The LLM decides the next step — not a hard-coded pipeline",
            size=18, bold=True, color=DARK_BLUE, align=PP_ALIGN.CENTER)
    textbox(s, Inches(0.5), Inches(5.7), Inches(12.5), Inches(0.6),
            "Each step is streamed to the browser via Server-Sent Events (SSE)",
            size=14, color=GRAY, align=PP_ALIGN.CENTER)

    # ===== 7. Search step =====
    s = blank_slide(prs)
    title_bar(s, "Search Step — find files by keyword", SW)

    card(s, Inches(0.7), Inches(2.0), Inches(3.5), Inches(2.0),
         "User Query", '"What is the metrics\nfor RAG?"', BLUE)
    arrow_right(s, Inches(4.3), Inches(2.85), Inches(0.5), Inches(0.4))
    card(s, Inches(4.9), Inches(2.0), Inches(3.5), Inches(2.0),
         "LLM Output", '"metric OR evaluation\nOR BLEU OR ROUGE"', ORANGE)
    arrow_right(s, Inches(8.5), Inches(2.85), Inches(0.5), Inches(0.4))
    card(s, Inches(9.1), Inches(2.0), Inches(3.5), Inches(2.0),
         "Microsoft Graph", "POST /search/query\nreturns file list + snippets", GREEN)

    textbox(s, Inches(0.5), Inches(4.7), Inches(12.5), Inches(0.5),
            "Microsoft has already indexed the content — no need to rebuild",
            size=16, bold=True, color=DARK_BLUE, align=PP_ALIGN.CENTER)
    textbox(s, Inches(0.5), Inches(5.5), Inches(12.5), Inches(0.5),
            "Each file comes with a ~240 char snippet around the matched keyword",
            size=14, color=GRAY, align=PP_ALIGN.CENTER)

    # ===== 8. Fetch =====
    s = blank_slide(prs)
    title_bar(s, "Fetch Step — download and extract text", SW)

    textbox(s, Inches(0.5), Inches(1.2), Inches(12.5), Inches(0.5),
            "Only when the LLM judges the snippet insufficient",
            size=16, color=GRAY, align=PP_ALIGN.CENTER)

    card(s, Inches(0.7), Inches(2.2), Inches(2.6), Inches(0.9), "PDF", "pdfplumber", BLUE)
    card(s, Inches(0.7), Inches(3.3), Inches(2.6), Inches(0.9), "DOCX", "python-docx", PURPLE)
    card(s, Inches(0.7), Inches(4.4), Inches(2.6), Inches(0.9), "XLSX", "openpyxl", GREEN)
    card(s, Inches(0.7), Inches(5.5), Inches(2.6), Inches(0.9), "PPTX", "python-pptx", ORANGE)

    arrow_right(s, Inches(3.5), Inches(3.7), Inches(0.6), Inches(0.5))

    proc = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                               Inches(4.3), Inches(2.5), Inches(4.5), Inches(3.0))
    fill(proc, LIGHT_BLUE)
    no_line(proc)
    tf = proc.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.2)
    tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.25)
    set_text(tf, "Extract Text", size=20, bold=True, color=DARK_BLUE, align=PP_ALIGN.CENTER)
    add_para(tf, "", size=10)
    add_para(tf, "Bytes loaded into RAM via io.BytesIO", size=13, color=DARK, align=PP_ALIGN.CENTER)
    add_para(tf, "Parse per file format", size=13, color=DARK, align=PP_ALIGN.CENTER)
    add_para(tf, "Markers [PAGE N] / [SHEET] / [SLIDE]", size=13, color=DARK, align=PP_ALIGN.CENTER)
    add_para(tf, "Line-oriented plain text", size=13, color=DARK, align=PP_ALIGN.CENTER)

    arrow_right(s, Inches(9.0), Inches(3.7), Inches(0.6), Inches(0.5))

    card(s, Inches(9.8), Inches(2.5), Inches(3.0), Inches(3.0),
         "RAM Cache",
         "item_id -> text\n\nSubsequent greps\nskip the download",
         TEAL)

    # ===== 9. Grep =====
    s = blank_slide(prs)
    title_bar(s, "Grep Step — extract context windows", SW)

    textbox(s, Inches(0.5), Inches(1.1), Inches(12.5), Inches(0.5),
            "Multi-pattern regex on the cached text",
            size=16, bold=True, color=DARK_BLUE, align=PP_ALIGN.CENTER)

    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              Inches(0.7), Inches(1.9), Inches(5.5), Inches(1.2))
    fill(box, ORANGE)
    no_line(box)
    set_text(box.text_frame, "LLM-generated pattern", size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_para(box.text_frame, "(metric|BLEU|ROUGE|recall|F1)", size=12, color=WHITE, align=PP_ALIGN.CENTER)

    arrow_down(s, Inches(3.0), Inches(3.2), Inches(0.5), Inches(0.4))

    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              Inches(0.7), Inches(3.7), Inches(5.5), Inches(3.3))
    fill(box, LIGHT_BG)
    no_line(box)
    tf = box.text_frame
    tf.margin_left = Inches(0.2)
    tf.margin_top = Inches(0.15)
    tf.word_wrap = True
    set_text(tf, "Output — chunks with plus/minus 5 lines", size=14, bold=True, color=DARK_BLUE)
    add_para(tf, "   L11: Section 4. Evaluation", size=11, color=GRAY)
    add_para(tf, "   L12: ", size=11, color=GRAY)
    add_para(tf, ">> L13: We use BLEU as primary metric", size=11, bold=True, color=BLUE)
    add_para(tf, "   L14: and ROUGE-L for generation.", size=11, color=GRAY)
    add_para(tf, "   L15: ...", size=11, color=GRAY)

    card(s, Inches(6.8), Inches(1.9), Inches(6.0), Inches(1.1),
         "Merge overlapping windows",
         "Two near-by hits become one chunk — no duplicates",
         GREEN)
    card(s, Inches(6.8), Inches(3.1), Inches(6.0), Inches(1.1),
         "Cap at 10 chunks",
         "Avoid blowing up the LLM context window",
         PURPLE)
    card(s, Inches(6.8), Inches(4.3), Inches(6.0), Inches(1.1),
         "Log to debug/grep.log",
         "Full context saved for inspection later",
         TEAL)
    card(s, Inches(6.8), Inches(5.5), Inches(6.0), Inches(1.1),
         "Cite line numbers",
         "LLM references exact lines in its answer",
         BLUE)

    # ===== 10. Final =====
    s = blank_slide(prs)
    title_bar(s, "Final Step — synthesize and cite", SW)

    card(s, Inches(0.7), Inches(1.6), Inches(3.8), Inches(4.5),
         "LLM inputs",
         "- User query\n\n- Search snippets\n\n- Grep chunks\n\n- File metadata\n\n(No full text!)",
         PURPLE)

    arrow_right(s, Inches(4.7), Inches(3.6), Inches(0.5), Inches(0.5))

    box = s.shapes.add_shape(MSO_SHAPE.OVAL,
                              Inches(5.3), Inches(2.6), Inches(2.8), Inches(2.8))
    fill(box, ORANGE)
    no_line(box)
    tf = box.text_frame
    tf.word_wrap = True
    set_text(tf, "LLM", size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_para(tf, "synthesizes", size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    arrow_right(s, Inches(8.3), Inches(3.6), Inches(0.5), Inches(0.5))

    card(s, Inches(8.9), Inches(1.6), Inches(3.9), Inches(4.5),
         "Final Answer",
         "Concise reply\n\n+ Citations:\n[file.pdf, L42]\n[doc.docx, PAGE 3]\n\nUser can click\nto open source file",
         GREEN)

    textbox(s, Inches(0.5), Inches(6.4), Inches(12.5), Inches(0.6),
            "LLM must cite — if nothing is found, it says so without inventing",
            size=14, color=DARK_BLUE, align=PP_ALIGN.CENTER)

    # ===== 11. Comparison =====
    s = blank_slide(prs)
    title_bar(s, "Comparison — Traditional RAG vs This Agent", SW)

    headers = ["Aspect", "Traditional RAG", "Agentic Keyword"]
    rows = [
        ["Infrastructure", "Dedicated vector DB", "None required"],
        ["Chunking", "Mandatory upfront", "Not needed"],
        ["Exact match", "Often misses", "Perfect"],
        ["Document updates", "Re-index everything", "Instant (Microsoft indexes)"],
        ["Snippet shape", "Fixed-size chunk", "Dynamic plus/minus N lines"],
        ["Cost", "Embedding + storage", "LLM calls only"],
    ]
    table = s.shapes.add_table(len(rows) + 1, 3,
                                Inches(0.7), Inches(1.3),
                                Inches(11.9), Inches(5.4)).table
    table.columns[0].width = Inches(3.0)
    table.columns[1].width = Inches(4.45)
    table.columns[2].width = Inches(4.45)

    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = BLUE
        set_text(cell.text_frame, h, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    for i, row in enumerate(rows, start=1):
        for j, v in enumerate(row):
            cell = table.cell(i, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHT_BG if i % 2 == 0 else WHITE
            color = DARK if j == 0 else (ORANGE if j == 1 else GREEN)
            set_text(cell.text_frame, v, size=14,
                     bold=(j == 0), color=color, align=PP_ALIGN.LEFT)

    # ===== 12. Conclusion =====
    s = blank_slide(prs)
    title_bar(s, "Key Takeaways", SW)

    points = [
        ("1", "Simpler", "No vector DB, no re-indexing"),
        ("2", "Instant", "Microsoft already indexes — query and go"),
        ("3", "Exact match", "Never misses product codes or technical terms"),
        ("4", "LLM-driven", "Generates keywords, retries strategies, cites sources"),
        ("5", "Transparent", "Every step streamed live and fully logged"),
    ]
    for i, (ico, h, t) in enumerate(points):
        y = Inches(1.4 + i * 1.05)
        ic = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.8), y, Inches(0.8), Inches(0.8))
        fill(ic, BLUE)
        no_line(ic)
        set_text(ic.text_frame, ico, size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

        textbox(s, Inches(2.0), y + Inches(0.05), Inches(3.0), Inches(0.5),
                h, size=18, bold=True, color=DARK_BLUE)
        textbox(s, Inches(5.2), y + Inches(0.1), Inches(7.5), Inches(0.5),
                t, size=14, color=GRAY)

    textbox(s, Inches(0.5), Inches(6.9), Inches(12.5), Inches(0.5),
            "Microsoft 365 Keyword-Search Agent",
            size=12, bold=True, color=BLUE, align=PP_ALIGN.CENTER)

    prs.save(path)
    print(f"[OK] Saved {path} ({len(prs.slides)} slides)")


def add_heading(doc, text, level=1, color=None):
    h = doc.add_heading(text, level=level)
    if color is not None:
        for run in h.runs:
            run.font.color.rgb = color
    return h


def add_body(doc, text, bold=False, italic=False, size=11):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = DocxPt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = "Calibri"
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(text, style="List Bullet")
    for run in p.runs:
        run.font.size = DocxPt(11)
        run.font.name = "Calibri"
    return p


def make_docx(path):
    doc = Document()

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("Microsoft 365 Keyword-Search Agent")
    r.font.size = DocxPt(24)
    r.font.bold = True
    r.font.color.rgb = DocxRGB(0x00, 0x78, 0xD4)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run("Process Documentation")
    r.font.size = DocxPt(14)
    r.font.italic = True
    r.font.color.rgb = DocxRGB(0x60, 0x5E, 0x5C)

    doc.add_paragraph()

    # ===== 1. Introduction =====
    add_heading(doc, "1. Introduction", level=1, color=DocxRGB(0x00, 0x78, 0xD4))
    add_body(doc,
        "This web application lets users query their OneDrive documents in "
        "natural language. Inspired by the paper \"Keyword search is all you "
        "need: Achieving RAG-Level Performance without vector databases using "
        "agentic tool use\" (arxiv 2602.23368), the system replaces the "
        "conventional Retrieval-Augmented Generation (RAG) architecture with "
        "an intelligent LLM agent that leverages Microsoft Graph's built-in "
        "keyword search.")

    # ===== 2. Problem =====
    add_heading(doc, "2. Problems with Traditional RAG", level=1, color=DocxRGB(0x00, 0x78, 0xD4))
    add_body(doc, "A traditional RAG pipeline requires several complex stages:")
    for b in [
        "Chunking — splits documents into small pieces. Context is easily lost when splits fall mid-sentence or mid-table.",
        "Embedding — computes a vector for every chunk. Significant compute and storage cost.",
        "Vector database — adds another service to deploy (FAISS, Pinecone, Weaviate, etc.).",
        "Re-indexing — every document update forces costly re-embedding.",
        "Weak exact match — semantic similarity may miss product codes or rare technical terms.",
    ]:
        add_bullet(doc, b)

    # ===== 3. Solution =====
    add_heading(doc, "3. Proposed Solution", level=1, color=DocxRGB(0x00, 0x78, 0xD4))
    add_body(doc,
        "Instead of building a RAG pipeline, the system uses an LLM agent "
        "that drives the search loop itself. The core insights are:")
    for b in [
        "Microsoft already provides a full-text index for OneDrive / SharePoint via the Graph Search API — no rebuild needed.",
        "The LLM generates many keyword variants (synonyms, abbreviations, regex) — compensating for literal-match-only indexes.",
        "Files are downloaded and extracted on demand only when snippets are insufficient — preserving original document context with no pre-chunking.",
        "Regex grep with a plus/minus N line context window — effectively a 'dynamic chunk' placed exactly around real hits.",
        "The LLM evaluates results and retries with different strategies when nothing useful comes back.",
    ]:
        add_bullet(doc, b)

    # ===== 4. Actors =====
    add_heading(doc, "4. Actors in the System", level=1, color=DocxRGB(0x00, 0x78, 0xD4))

    add_heading(doc, "4.1. User", level=2)
    add_body(doc, "Signs in with a Microsoft 365 account and types a natural-language question into the dashboard chat panel.")

    add_heading(doc, "4.2. Web UI (Browser)", level=2)
    add_body(doc,
        "Dashboard page displays the user's OneDrive, Teams, and SharePoint counts. "
        "The JavaScript chat panel sends requests to the server and renders each agent "
        "step in real time via Server-Sent Events (SSE).")

    add_heading(doc, "4.3. Flask Backend", level=2)
    add_body(doc,
        "Python web server. Handles OAuth2 sign-in via Microsoft Entra ID, manages the "
        "session, and exposes the /ask route that boots the agent. Passes the "
        "user's access token through to the agent so it can call Microsoft Graph "
        "on the user's behalf.")

    add_heading(doc, "4.4. LangChain Agent (langgraph)", level=2)
    add_body(doc,
        "The orchestrator of the tool-use loop. Feeds the question and the system "
        "prompt to the LLM, receives its tool-call decisions, executes each tool, "
        "feeds the result back. Loops up to 30 iterations until the LLM produces a "
        "final answer.")

    add_heading(doc, "4.5. LLM (Azure OpenAI)", level=2)
    add_body(doc,
        "The decision-making brain. Reads the user question and the conversation "
        "history to decide which tool to call with which arguments, or when the "
        "answer is ready. It generates OR-joined keyword variants for search "
        "and multi-pattern regex for grep. Default model: gpt-4.1-mini.")

    add_heading(doc, "4.6. Microsoft Graph API", level=2)
    add_body(doc,
        "Microsoft's official API for 365 data. Two endpoints are essential:")
    for b in [
        "POST /search/query — full-text search over OneDrive, SharePoint, and Teams files. Returns metadata plus a snippet around the matched keyword.",
        "GET /drives/{driveId}/items/{id}/content — downloads the binary content of a file.",
    ]:
        add_bullet(doc, b)

    add_heading(doc, "4.7. The Three Agent Tools", level=2)
    add_body(doc, "Each tool is a Python function registered with LangChain:")
    for b in [
        "graph_search(query) — calls Microsoft Graph Search and returns matching files.",
        "fetch_file_text(item_id) — downloads a file, extracts its text, caches it in RAM.",
        "grep_context(item_id, patterns, context_lines) — regex-searches the cached text and returns chunks with plus/minus N lines of surrounding context.",
    ]:
        add_bullet(doc, b)

    add_heading(doc, "4.8. File Extractors", level=2)
    add_body(doc, "Four modules convert binary content into line-oriented plain text:")
    for b in [
        "PDF — pdfplumber, marks each page with [PAGE N].",
        "DOCX — python-docx, captures paragraphs and tables.",
        "XLSX — openpyxl, marks each sheet with [SHEET name], joins cells with the pipe character.",
        "PPTX — python-pptx, marks each slide with [SLIDE N], including speaker notes.",
    ]:
        add_bullet(doc, b)

    add_heading(doc, "4.9. Cache and Drive Index", level=2)
    add_body(doc, "Two in-memory dictionaries improve efficiency:")
    for b in [
        "Text cache (cache.py) — once a file is extracted, its text is held in RAM. Later grep calls skip re-downloading.",
        "Drive index (drive_index.py) — maps item_id to driveId so fetch uses the right endpoint for personal, shared, SharePoint, or Teams files.",
    ]:
        add_bullet(doc, b)

    # ===== 5. End-to-End =====
    add_heading(doc, "5. End-to-End Workflow", level=1, color=DocxRGB(0x00, 0x78, 0xD4))

    steps = [
        ("Step 1 — User submits a question",
         "The user types a natural-language question into the chat panel, "
         "for example \"What is the metrics for RAG?\", and clicks Send."),
        ("Step 2 — Browser sends the request",
         "JavaScript prevents the default form submission, calls fetch POST "
         "/ask with the question as JSON, and accepts a text/event-stream response."),
        ("Step 3 — Flask initializes the agent",
         "The server pulls the access_token from the session (already obtained "
         "from the MSAL OAuth2 flow). It builds a LangChain agent with three "
         "tools that capture the token via closures, then opens an SSE stream."),
        ("Step 4 — LLM generates search keywords",
         "The LLM reads the question and the system prompt. The prompt mandates "
         "multi-pattern OR queries (never a single isolated word). The LLM emits "
         "a graph_search tool call like \"metric OR evaluation OR BLEU OR ROUGE "
         "OR recall\"."),
        ("Step 5 — Microsoft Graph finds files",
         "The graph_search tool sends POST /search/query with entityTypes "
         "[driveItem]. Microsoft's index returns up to 15 matching files, each "
         "with a ~240 character snippet centered on the matched keyword."),
        ("Step 6 — LLM evaluates the snippets",
         "The LLM reviews the returned files and snippets. Two paths: "
         "(a) the snippets already answer the question — jump to Step 9; "
         "(b) more depth is needed — proceed to Step 7."),
        ("Step 7 — Fetch and extract files",
         "The LLM calls fetch_file_text for the 1-3 most relevant items. The "
         "tool looks up the driveId in the drive index, downloads the file "
         "through the Graph endpoint (with auto-follow of the CDN redirect), "
         "extracts text using the right extractor, and caches it. The LLM "
         "receives only a short preview to avoid blowing up the context window."),
        ("Step 8 — Grep with context",
         "The LLM calls grep_context with a multi-pattern regex and "
         "context_lines=5. The tool searches the cached text, merges "
         "overlapping windows, and returns up to 10 chunks. Each chunk shows "
         "hit lines prefixed with '>>' and surrounding context lines with "
         "two spaces. The full result is also written to debug/grep.log."),
        ("Step 9 — LLM synthesizes the answer",
         "With enough context gathered, the LLM produces a final answer with "
         "citations like [filename, L42] or [filename, PAGE 3]. If nothing "
         "useful was found, the LLM must say so plainly — it is forbidden "
         "from making things up."),
        ("Step 10 — Stream back to the browser",
         "Each step (tool call, tool result, final answer, summary) is yielded "
         "as an SSE event. The browser receives and renders each event in "
         "real time on the chat panel."),
    ]
    for h, t in steps:
        add_heading(doc, h, level=3)
        add_body(doc, t)

    # ===== 6. SSE =====
    add_heading(doc, "6. Server-Sent Events (SSE)", level=1, color=DocxRGB(0x00, 0x78, 0xD4))
    add_body(doc,
        "SSE lets the server push events continuously to the browser over a "
        "single HTTP connection. In this application, every agent step (LLM "
        "decides to call a tool, tool returns a result, LLM produces the final "
        "answer) is sent the moment it occurs. The browser renders each event "
        "live, letting users observe the agent's reasoning rather than waiting "
        "for the entire workflow to finish.")

    # ===== 7. Cache =====
    add_heading(doc, "7. The Role of the Cache and Drive Index", level=1, color=DocxRGB(0x00, 0x78, 0xD4))
    add_body(doc,
        "Each file is downloaded once. After extraction, the text is stored in "
        "an in-memory dictionary keyed by the Graph item_id. Subsequent "
        "grep_context calls read straight from RAM with no network round-trip. "
        "The drive index stores item_id to driveId because OneDrive files do "
        "not always live under /me/drive — they may be shared, hosted by a "
        "SharePoint site, or owned by a Teams group drive.")

    # ===== 8. Comparison =====
    add_heading(doc, "8. Comparison with Traditional RAG", level=1, color=DocxRGB(0x00, 0x78, 0xD4))
    table = doc.add_table(rows=7, cols=3)
    table.style = "Light Grid Accent 1"
    headers = ["Aspect", "Traditional RAG", "Agentic Keyword (this system)"]
    rows = [
        ["Infrastructure", "Dedicated vector DB (FAISS, Pinecone, ...)", "None — uses Microsoft's existing index"],
        ["Chunking", "Required, fixed upfront", "Not required — dynamic context window per hit"],
        ["Exact match", "May miss (semantic similarity)", "Perfect (literal regex)"],
        ["Document updates", "Re-index everything", "Instant (Microsoft indexes automatically)"],
        ["Cost", "Embedding + storage + LLM", "LLM calls only"],
        ["Transparency", "Hard to debug similarity scores", "Every step logged and traceable"],
    ]
    cells = table.rows[0].cells
    for i, h in enumerate(headers):
        cells[i].text = ""
        run = cells[i].paragraphs[0].add_run(h)
        run.font.bold = True
        run.font.size = DocxPt(11)
    for i, row in enumerate(rows, start=1):
        for j, v in enumerate(row):
            table.rows[i].cells[j].text = ""
            run = table.rows[i].cells[j].paragraphs[0].add_run(v)
            run.font.size = DocxPt(10)

    # ===== 9. Limitations =====
    add_heading(doc, "9. Current Limitations", level=1, color=DocxRGB(0x00, 0x78, 0xD4))
    for b in [
        "Scanned PDFs (image-only) — pdfplumber cannot OCR; the extracted text is empty. Azure Document Intelligence or Docling can be added for OCR.",
        "25 MB per-file limit — larger files are refused to protect RAM and avoid timeouts.",
        "Latency depends on LLM call count — each question typically triggers 5-15 ReAct loop iterations.",
        "Accuracy depends on the model — gpt-4.1-mini works well; weaker models may fail to generate enough keyword variants.",
        "Microsoft search index lag — newly uploaded files can take a few minutes to become searchable.",
    ]:
        add_bullet(doc, b)

    # ===== 10. Conclusion =====
    add_heading(doc, "10. Conclusion", level=1, color=DocxRGB(0x00, 0x78, 0xD4))
    add_body(doc,
        "The system demonstrates that with a capable LLM and the right keyword "
        "search toolkit, one can match RAG-level performance without "
        "building or operating a vector database. By riding on Microsoft "
        "Graph's existing search index and letting the LLM drive the search, "
        "the architecture becomes simpler, cheaper, easier to debug, and "
        "offers perfect exact-match support for technical documents.")

    doc.add_paragraph()
    foot = doc.add_paragraph()
    foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = foot.add_run("— Microsoft 365 Keyword-Search Agent —")
    r.font.size = DocxPt(10)
    r.font.italic = True
    r.font.color.rgb = DocxRGB(0x60, 0x5E, 0x5C)

    doc.save(path)
    print(f"[OK] Saved {path}")


if __name__ == "__main__":
    here = os.path.dirname(os.path.abspath(__file__))
    make_pptx(os.path.join(here, "graph_grep_en.pptx"))
    make_docx(os.path.join(here, "docs_report_en.docx"))
