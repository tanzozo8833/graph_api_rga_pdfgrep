"""Generate graph_grep.pptx and docs_report.docx — project documentation.

Run: venv\\Scripts\\python.exe make_docs.py
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


# ---------- pptx helpers ----------
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


# ---------- pptx generation ----------
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
            "Hỏi đáp tài liệu OneDrive — không cần vector database",
            size=22, color=WHITE, align=PP_ALIGN.CENTER)
    textbox(s, Inches(1), Inches(5.0), Inches(11.5), Inches(0.5),
            "Cảm hứng: paper 'Keyword search is all you need' (arxiv 2602.23368)",
            size=14, color=LIGHT_BG, align=PP_ALIGN.CENTER)
    textbox(s, Inches(1), Inches(6.5), Inches(11.5), Inches(0.4),
            "Microsoft Graph API  •  Azure OpenAI  •  LangChain Agent",
            size=12, color=LIGHT_BG, align=PP_ALIGN.CENTER)

    # ===== 2. Vấn đề =====
    s = blank_slide(prs)
    title_bar(s, "Vấn đề — RAG truyền thống quá phức tạp", SW)
    items = [
        ("🧩  Chunking", "Cắt nhỏ tài liệu, dễ mất ngữ cảnh giữa câu/bảng", BLUE),
        ("💰  Embedding", "Tốn chi phí tính vector cho mọi chunk", PURPLE),
        ("🗄  Vector DB", "Phải vận hành thêm một service riêng", TEAL),
        ("🎯  Exact match", "Có thể miss mã sản phẩm / từ kỹ thuật hiếm", ORANGE),
    ]
    for i, (h, t, c) in enumerate(items):
        x = Inches(0.6 + (i % 2) * 6.3)
        y = Inches(1.4 + (i // 2) * 2.7)
        card(s, x, y, Inches(6.0), Inches(2.3), h, t, c)

    # ===== 3. Giải pháp =====
    s = blank_slide(prs)
    title_bar(s, "Giải pháp — Agentic Keyword Search", SW)
    textbox(s, Inches(0.7), Inches(1.2), Inches(12), Inches(0.6),
            "Để LLM tự lái việc tìm kiếm — không pipeline cố định",
            size=20, bold=True, color=DARK_BLUE)
    feats = [
        ("LLM tự sinh keyword", "Synonyms, viết tắt, regex, OR-joined", BLUE),
        ("Microsoft Graph", "Index full-text có sẵn cho OneDrive", GREEN),
        ("Tải file khi cần", "Extract text bằng pdfplumber / docx / xlsx / pptx", PURPLE),
        ("Grep với context", "Regex multi-pattern, ±N dòng quanh hit", TEAL),
        ("LLM tự retry", "Đổi chiến lược nếu chưa tìm thấy", ORANGE),
    ]
    for i, (h, t, c) in enumerate(feats):
        x = Inches(0.6 + i * 2.5)
        y = Inches(2.5)
        card(s, x, y, Inches(2.4), Inches(2.7), h, t, c)
    textbox(s, Inches(0.7), Inches(5.7), Inches(12), Inches(0.6),
            "Microsoft đã có sẵn keyword index — không cần build lại",
            size=16, color=GRAY, align=PP_ALIGN.CENTER)

    # ===== 4. Tác nhân =====
    s = blank_slide(prs)
    title_bar(s, "Các tác nhân tham gia", SW)

    # User
    card(s, Inches(0.5), Inches(1.3), Inches(2.4), Inches(1.1),
         "👤 User", "Người hỏi", DARK_BLUE)
    arrow_right(s, Inches(2.95), Inches(1.65), Inches(0.4), Inches(0.4))
    # Browser
    card(s, Inches(3.4), Inches(1.3), Inches(2.4), Inches(1.1),
         "🌐 Browser", "Chat panel + SSE", BLUE)
    arrow_right(s, Inches(5.85), Inches(1.65), Inches(0.4), Inches(0.4))
    # Flask
    card(s, Inches(6.3), Inches(1.3), Inches(2.4), Inches(1.1),
         "🔥 Flask", "Route /ask + MSAL", PURPLE)
    arrow_right(s, Inches(8.75), Inches(1.65), Inches(0.4), Inches(0.4))
    # Agent
    card(s, Inches(9.2), Inches(1.3), Inches(3.5), Inches(1.1),
         "🤖 LangChain Agent", "Vòng lặp tool-use", ORANGE)

    # Down arrow from agent
    arrow_down(s, Inches(10.7), Inches(2.5), Inches(0.4), Inches(0.5))

    # Three branches under agent
    card(s, Inches(2.0), Inches(3.3), Inches(3.3), Inches(1.1),
         "🧠 Azure OpenAI", "gpt-4.1-mini quyết định", BLUE)
    card(s, Inches(5.5), Inches(3.3), Inches(3.3), Inches(1.1),
         "☁ Microsoft Graph", "Search & Download", GREEN)
    card(s, Inches(9.0), Inches(3.3), Inches(3.3), Inches(1.1),
         "📦 Local Modules", "Extractors + Cache", TEAL)

    # Down arrow
    arrow_down(s, Inches(6.5), Inches(4.5), Inches(0.4), Inches(0.5))

    # Tools row
    card(s, Inches(1.5), Inches(5.2), Inches(3.0), Inches(1.4),
         "🔎 graph_search", "Tìm file theo keyword", BLUE)
    card(s, Inches(5.0), Inches(5.2), Inches(3.0), Inches(1.4),
         "📥 fetch_file_text", "Download + extract", GREEN)
    card(s, Inches(8.5), Inches(5.2), Inches(3.0), Inches(1.4),
         "🔬 grep_context", "Regex + context window", PURPLE)

    # ===== 5. 3 Tools =====
    s = blank_slide(prs)
    title_bar(s, "Ba công cụ của Agent", SW)

    cols = [
        ("🔎  graph_search", BLUE,
         ["Tham số: query (string)",
          "Gọi POST /search/query",
          "Trả về list file + snippet",
          "OR-joined: 'k1 OR k2 OR k3'"]),
        ("📥  fetch_file_text", GREEN,
         ["Tham số: item_id",
          "Download bytes từ Graph CDN",
          "Extract: PDF / DOCX / XLSX / PPTX",
          "Cache text vào RAM"]),
        ("🔬  grep_context", PURPLE,
         ["Tham số: item_id, patterns, ±N",
          "Regex IGNORECASE trên cache",
          "Merge overlapping windows",
          "Trả về ±5 dòng quanh hit"]),
    ]
    for i, (h, c, bullets) in enumerate(cols):
        x = Inches(0.5 + i * 4.3)
        # header
        hdr = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  x, Inches(1.2), Inches(4.0), Inches(0.9))
        fill(hdr, c)
        no_line(hdr)
        set_text(hdr.text_frame, h, size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        # body
        body = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   x, Inches(2.2), Inches(4.0), Inches(4.5))
        fill(body, LIGHT_BG)
        no_line(body)
        btf = body.text_frame
        btf.margin_left = Inches(0.2)
        btf.margin_right = Inches(0.2)
        btf.margin_top = Inches(0.2)
        btf.word_wrap = True
        set_text(btf, "• " + bullets[0], size=14, color=DARK)
        for b in bullets[1:]:
            add_para(btf, "• " + b, size=14, color=DARK)

    # ===== 6. Workflow =====
    s = blank_slide(prs)
    title_bar(s, "Quy trình từ đầu đến cuối", SW)

    steps = [
        ("1", "User hỏi", "Gõ câu hỏi vào chat panel", BLUE),
        ("2", "Flask /ask", "Khởi tạo agent với token", PURPLE),
        ("3", "LLM sinh keyword", "OR-joined synonyms", ORANGE),
        ("4", "graph_search", "Microsoft Graph trả file + snippet", GREEN),
        ("5", "Đủ?", "Snippet đã trả lời được?", TEAL),
        ("6", "fetch + grep", "Đọc sâu nếu cần", DARK_BLUE),
        ("7", "Final answer", "Tổng hợp + citation", BLUE),
    ]
    for i, (n, h, t, c) in enumerate(steps):
        x = Inches(0.3 + i * 1.85)
        y = Inches(1.7)
        # Number circle
        num = s.shapes.add_shape(MSO_SHAPE.OVAL, x + Inches(0.55), y, Inches(0.6), Inches(0.6))
        fill(num, c)
        no_line(num)
        set_text(num.text_frame, n, size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        # Card
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
        # Arrow except last
        if i < len(steps) - 1:
            arrow_right(s, x + Inches(1.72), y + Inches(0.18), Inches(0.13), Inches(0.25), color=GRAY)

    textbox(s, Inches(0.5), Inches(5.0), Inches(12.5), Inches(0.6),
            "LLM tự quyết định bước tiếp theo — không phải pipeline cố định",
            size=18, bold=True, color=DARK_BLUE, align=PP_ALIGN.CENTER)
    textbox(s, Inches(0.5), Inches(5.7), Inches(12.5), Inches(0.6),
            "Mỗi bước được stream về browser qua Server-Sent Events (SSE)",
            size=14, color=GRAY, align=PP_ALIGN.CENTER)

    # ===== 7. Bước Search =====
    s = blank_slide(prs)
    title_bar(s, "Bước Search — tìm file qua keyword", SW)

    card(s, Inches(0.7), Inches(2.0), Inches(3.5), Inches(2.0),
         "User Query", '"What is the metrics\nfor RAG?"', BLUE)
    arrow_right(s, Inches(4.3), Inches(2.85), Inches(0.5), Inches(0.4))
    card(s, Inches(4.9), Inches(2.0), Inches(3.5), Inches(2.0),
         "LLM Output", '"metric OR evaluation\nOR BLEU OR ROUGE"', ORANGE)
    arrow_right(s, Inches(8.5), Inches(2.85), Inches(0.5), Inches(0.4))
    card(s, Inches(9.1), Inches(2.0), Inches(3.5), Inches(2.0),
         "Microsoft Graph", "POST /search/query\n→ list file + snippet", GREEN)

    textbox(s, Inches(0.5), Inches(4.7), Inches(12.5), Inches(0.5),
            "Microsoft đã index full-text — không cần ta tự build",
            size=16, bold=True, color=DARK_BLUE, align=PP_ALIGN.CENTER)
    textbox(s, Inches(0.5), Inches(5.5), Inches(12.5), Inches(0.5),
            "Mỗi file kèm snippet ~240 ký tự quanh keyword khớp",
            size=14, color=GRAY, align=PP_ALIGN.CENTER)

    # ===== 8. Fetch & Extract =====
    s = blank_slide(prs)
    title_bar(s, "Bước Fetch — tải file và extract text", SW)

    textbox(s, Inches(0.5), Inches(1.2), Inches(12.5), Inches(0.5),
            "Chỉ khi LLM cho rằng snippet chưa đủ trả lời",
            size=16, color=GRAY, align=PP_ALIGN.CENTER)

    # Input formats column
    card(s, Inches(0.7), Inches(2.2), Inches(2.6), Inches(0.9), "📄 PDF", "pdfplumber", BLUE)
    card(s, Inches(0.7), Inches(3.3), Inches(2.6), Inches(0.9), "📝 DOCX", "python-docx", PURPLE)
    card(s, Inches(0.7), Inches(4.4), Inches(2.6), Inches(0.9), "📊 XLSX", "openpyxl", GREEN)
    card(s, Inches(0.7), Inches(5.5), Inches(2.6), Inches(0.9), "🎯 PPTX", "python-pptx", ORANGE)

    arrow_right(s, Inches(3.5), Inches(3.7), Inches(0.6), Inches(0.5))

    # Process
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
    add_para(tf, "→ Bytes về RAM qua io.BytesIO", size=13, color=DARK, align=PP_ALIGN.CENTER)
    add_para(tf, "→ Parse theo từng định dạng", size=13, color=DARK, align=PP_ALIGN.CENTER)
    add_para(tf, "→ Markers [PAGE N] / [SHEET] / [SLIDE]", size=13, color=DARK, align=PP_ALIGN.CENTER)
    add_para(tf, "→ Plain text dòng-by-dòng", size=13, color=DARK, align=PP_ALIGN.CENTER)

    arrow_right(s, Inches(9.0), Inches(3.7), Inches(0.6), Inches(0.5))

    # Cache
    card(s, Inches(9.8), Inches(2.5), Inches(3.0), Inches(3.0),
         "💾 Cache RAM",
         "item_id → text\n\nGrep lần sau\nkhông download lại",
         TEAL)

    # ===== 9. Grep =====
    s = blank_slide(prs)
    title_bar(s, "Bước Grep — trích xuất ngữ cảnh", SW)

    textbox(s, Inches(0.5), Inches(1.1), Inches(12.5), Inches(0.5),
            "Regex multi-pattern trên text đã cache",
            size=16, bold=True, color=DARK_BLUE, align=PP_ALIGN.CENTER)

    # Input pattern
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              Inches(0.7), Inches(1.9), Inches(5.5), Inches(1.2))
    fill(box, ORANGE)
    no_line(box)
    set_text(box.text_frame, "LLM sinh pattern", size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_para(box.text_frame, "(metric|BLEU|ROUGE|recall|F1)", size=12, color=WHITE, align=PP_ALIGN.CENTER)

    arrow_down(s, Inches(3.0), Inches(3.2), Inches(0.5), Inches(0.4))

    # Match + window
    box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              Inches(0.7), Inches(3.7), Inches(5.5), Inches(3.3))
    fill(box, LIGHT_BG)
    no_line(box)
    tf = box.text_frame
    tf.margin_left = Inches(0.2)
    tf.margin_top = Inches(0.15)
    tf.word_wrap = True
    set_text(tf, "Output — chunk ±5 dòng", size=14, bold=True, color=DARK_BLUE)
    add_para(tf, "   L11: Section 4. Evaluation", size=11, color=GRAY)
    add_para(tf, "   L12: ", size=11, color=GRAY)
    add_para(tf, ">> L13: We use BLEU as primary metric", size=11, bold=True, color=BLUE)
    add_para(tf, "   L14: and ROUGE-L for generation.", size=11, color=GRAY)
    add_para(tf, "   L15: ...", size=11, color=GRAY)

    # Right side: features
    card(s, Inches(6.8), Inches(1.9), Inches(6.0), Inches(1.1),
         "Merge overlapping windows",
         "2 hit gần nhau → 1 chunk duy nhất, tránh trùng lặp",
         GREEN)
    card(s, Inches(6.8), Inches(3.1), Inches(6.0), Inches(1.1),
         "Cap 10 chunks tối đa",
         "Tránh nuốt hết context window của LLM",
         PURPLE)
    card(s, Inches(6.8), Inches(4.3), Inches(6.0), Inches(1.1),
         "Log vào debug/grep.log",
         "Toàn bộ context được lưu để inspect",
         TEAL)
    card(s, Inches(6.8), Inches(5.5), Inches(6.0), Inches(1.1),
         "Cite line number",
         "LLM tham chiếu được dòng cụ thể trong câu trả lời",
         BLUE)

    # ===== 10. Final Answer =====
    s = blank_slide(prs)
    title_bar(s, "Bước Final — tổng hợp và trích dẫn", SW)

    # Input observations
    card(s, Inches(0.7), Inches(1.6), Inches(3.8), Inches(4.5),
         "Inputs vào LLM",
         "• User query\n\n• Search snippets\n\n• Grep chunks\n\n• File metadata\n\n(Không có full text)",
         PURPLE)

    arrow_right(s, Inches(4.7), Inches(3.6), Inches(0.5), Inches(0.5))

    # LLM brain
    box = s.shapes.add_shape(MSO_SHAPE.OVAL,
                              Inches(5.3), Inches(2.6), Inches(2.8), Inches(2.8))
    fill(box, ORANGE)
    no_line(box)
    tf = box.text_frame
    tf.word_wrap = True
    set_text(tf, "🧠", size=48, color=WHITE, align=PP_ALIGN.CENTER)
    add_para(tf, "LLM tổng hợp", size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    arrow_right(s, Inches(8.3), Inches(3.6), Inches(0.5), Inches(0.5))

    # Answer
    card(s, Inches(8.9), Inches(1.6), Inches(3.9), Inches(4.5),
         "Final Answer",
         "Câu trả lời ngắn gọn\n\n+ Citation:\n[file.pdf, L42]\n[doc.docx, PAGE 3]\n\nNgười dùng click\nmở file gốc trên OneDrive",
         GREEN)

    textbox(s, Inches(0.5), Inches(6.4), Inches(12.5), Inches(0.6),
            "LLM phải cite — nếu không tìm thấy thì nói rõ, không bịa",
            size=14, color=DARK_BLUE, align=PP_ALIGN.CENTER)

    # ===== 11. So sánh =====
    s = blank_slide(prs)
    title_bar(s, "So sánh — RAG truyền thống vs Agent này", SW)

    headers = ["Tiêu chí", "RAG truyền thống", "Agentic Keyword"]
    rows = [
        ["Hạ tầng", "Vector DB riêng", "Không cần"],
        ["Chunking", "Bắt buộc", "Không cần"],
        ["Exact match", "Có thể miss", "Hoàn hảo"],
        ["Cập nhật", "Re-index toàn bộ", "Tức thì (Microsoft index)"],
        ["Snippet", "Chunk cố định", "Context động ±N dòng"],
        ["Chi phí", "Embedding + storage", "Chỉ LLM calls"],
    ]
    table = s.shapes.add_table(len(rows) + 1, 3,
                                Inches(0.7), Inches(1.3),
                                Inches(11.9), Inches(5.4)).table
    table.columns[0].width = Inches(3.0)
    table.columns[1].width = Inches(4.45)
    table.columns[2].width = Inches(4.45)

    # Header row
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.fill.solid()
        cell.fill.fore_color.rgb = BLUE
        set_text(cell.text_frame, h, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    # Data rows
    for i, row in enumerate(rows, start=1):
        for j, v in enumerate(row):
            cell = table.cell(i, j)
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHT_BG if i % 2 == 0 else WHITE
            color = DARK if j == 0 else (ORANGE if j == 1 else GREEN)
            set_text(cell.text_frame, v, size=14,
                     bold=(j == 0), color=color, align=PP_ALIGN.LEFT)

    # ===== 12. Kết luận =====
    s = blank_slide(prs)
    title_bar(s, "Kết luận", SW)

    points = [
        ("✨", "Đơn giản hơn", "Không cần vector DB, không re-index"),
        ("⚡", "Tức thời", "Microsoft index sẵn — query là có ngay"),
        ("🎯", "Exact match", "Không bao giờ miss mã sản phẩm / từ kỹ thuật"),
        ("🤖", "LLM tự lái", "Sinh keyword, retry chiến lược, cite source"),
        ("👁", "Minh bạch", "Mỗi bước stream realtime + log đầy đủ"),
    ]
    for i, (ico, h, t) in enumerate(points):
        y = Inches(1.4 + i * 1.05)
        ic = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.8), y, Inches(0.8), Inches(0.8))
        fill(ic, BLUE)
        no_line(ic)
        set_text(ic.text_frame, ico, size=22, color=WHITE, align=PP_ALIGN.CENTER)

        textbox(s, Inches(2.0), y + Inches(0.05), Inches(3.0), Inches(0.5),
                h, size=18, bold=True, color=DARK_BLUE)
        textbox(s, Inches(5.2), y + Inches(0.1), Inches(7.5), Inches(0.5),
                t, size=14, color=GRAY)

    textbox(s, Inches(0.5), Inches(6.9), Inches(12.5), Inches(0.5),
            "Microsoft 365 Keyword-Search Agent",
            size=12, bold=True, color=BLUE, align=PP_ALIGN.CENTER)

    prs.save(path)
    print(f"[OK] Saved {path} ({len(prs.slides)} slides)")


# ---------- docx helpers ----------
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


def add_bullet(doc, text, level=0):
    p = doc.add_paragraph(text, style="List Bullet")
    if level:
        p.paragraph_format.left_indent = DocxPt(18 * (level + 1))
    for run in p.runs:
        run.font.size = DocxPt(11)
        run.font.name = "Calibri"
    return p


# ---------- docx generation ----------
def make_docx(path):
    doc = Document()

    # ====== Title ======
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("Microsoft 365 Keyword-Search Agent")
    r.font.size = DocxPt(24)
    r.font.bold = True
    r.font.color.rgb = DocxRGB(0x00, 0x78, 0xD4)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = sub.add_run("Báo cáo quy trình hoạt động")
    r.font.size = DocxPt(14)
    r.font.italic = True
    r.font.color.rgb = DocxRGB(0x60, 0x5E, 0x5C)

    doc.add_paragraph()

    # ====== 1. Giới thiệu ======
    add_heading(doc, "1. Giới thiệu", level=1, color=DocxRGB(0x00, 0x78, 0xD4))
    add_body(doc,
        "Ứng dụng web cho phép người dùng hỏi đáp tài liệu OneDrive bằng "
        "ngôn ngữ tự nhiên. Lấy cảm hứng từ paper \"Keyword search is all "
        "you need: Achieving RAG-Level Performance without vector databases "
        "using agentic tool use\" (arxiv 2602.23368), hệ thống thay thế "
        "kiến trúc Retrieval-Augmented Generation (RAG) truyền thống bằng "
        "một LLM agent thông minh sử dụng keyword search có sẵn của "
        "Microsoft Graph API.")

    # ====== 2. Vấn đề ======
    add_heading(doc, "2. Vấn đề với RAG truyền thống", level=1, color=DocxRGB(0x00, 0x78, 0xD4))
    add_body(doc, "RAG truyền thống đòi hỏi một chuỗi xử lý phức tạp:")
    for b in [
        "Chunking — cắt tài liệu thành các đoạn nhỏ. Dễ mất ngữ cảnh khi cắt giữa câu hoặc bảng.",
        "Embedding — tạo vector cho mỗi chunk. Tốn chi phí tính toán và lưu trữ.",
        "Vector database — phải vận hành thêm một service riêng (FAISS, Pinecone, Weaviate...).",
        "Re-indexing — khi tài liệu cập nhật phải tái tạo embedding tốn kém.",
        "Exact match yếu — semantic similarity có thể bỏ qua mã sản phẩm, từ kỹ thuật hiếm.",
    ]:
        add_bullet(doc, b)

    # ====== 3. Giải pháp ======
    add_heading(doc, "3. Giải pháp đề xuất", level=1, color=DocxRGB(0x00, 0x78, 0xD4))
    add_body(doc,
        "Thay vì xây dựng pipeline RAG, hệ thống dùng một agent LLM tự "
        "điều khiển việc tìm kiếm. Cốt lõi của ý tưởng là:")
    for b in [
        "Microsoft đã có sẵn full-text index cho OneDrive/SharePoint — tận dụng qua Graph Search API.",
        "LLM tự sinh nhiều biến thể keyword (synonyms, abbreviations, regex) — bù đắp việc index chỉ match literal.",
        "Tải file về và extract text khi snippet không đủ — giữ nguyên ngữ cảnh gốc, không chunk trước.",
        "Regex grep với context window ±N dòng — mô phỏng \"chunk động\" đặt trên hit thực tế.",
        "LLM tự đánh giá kết quả và retry với chiến lược khác nếu chưa tìm thấy.",
    ]:
        add_bullet(doc, b)

    # ====== 4. Tác nhân ======
    add_heading(doc, "4. Các tác nhân tham gia", level=1, color=DocxRGB(0x00, 0x78, 0xD4))

    add_heading(doc, "4.1. User (Người dùng)", level=2)
    add_body(doc, "Đăng nhập tài khoản Microsoft 365 và gõ câu hỏi tự nhiên vào chat panel trên dashboard web.")

    add_heading(doc, "4.2. Web UI (Browser)", level=2)
    add_body(doc,
        "Trang dashboard hiển thị thông tin OneDrive/Teams/SharePoint của user. "
        "Chat panel JavaScript gửi request đến server và hiển thị từng bước "
        "agent thực hiện theo thời gian thực qua Server-Sent Events (SSE).")

    add_heading(doc, "4.3. Flask Backend", level=2)
    add_body(doc,
        "Server Python xử lý đăng nhập OAuth2 (Microsoft Entra ID), quản lý "
        "session, cung cấp route /ask để khởi tạo agent. Truyền access_token "
        "đến agent để gọi Microsoft Graph thay user.")

    add_heading(doc, "4.4. LangChain Agent (langgraph)", level=2)
    add_body(doc,
        "Bộ điều phối vòng lặp tool-use. Đưa câu hỏi và system prompt cho LLM, "
        "nhận quyết định gọi tool nào, chạy tool đó, đưa kết quả lại cho LLM. "
        "Lặp tối đa 30 vòng đến khi LLM sinh câu trả lời cuối.")

    add_heading(doc, "4.5. LLM (Azure OpenAI)", level=2)
    add_body(doc,
        "Bộ não quyết định. Đọc câu hỏi và lịch sử tương tác để quyết định: "
        "gọi tool nào, với tham số gì, hoặc đã đủ để trả lời. Sinh các "
        "biến thể keyword OR-joined cho tìm kiếm và regex multi-pattern cho grep. "
        "Mô hình mặc định: gpt-4.1-mini.")

    add_heading(doc, "4.6. Microsoft Graph API", level=2)
    add_body(doc,
        "Endpoint chính thức của Microsoft cho 365 data. Hai endpoint quan trọng:")
    for b in [
        "POST /search/query — full-text search trên OneDrive, SharePoint, Teams files. Trả về metadata + snippet quanh keyword.",
        "GET /drives/{driveId}/items/{id}/content — tải nội dung binary của một file.",
    ]:
        add_bullet(doc, b)

    add_heading(doc, "4.7. Ba công cụ của Agent", level=2)
    add_body(doc, "Mỗi công cụ là một hàm Python được đăng ký với LangChain:")
    for b in [
        "graph_search(query) — gọi Microsoft Graph Search, trả danh sách file phù hợp với keyword.",
        "fetch_file_text(item_id) — tải binary file, extract text, lưu cache trong RAM.",
        "grep_context(item_id, patterns, context_lines) — regex tìm pattern trên text đã cache, trả các chunk ±N dòng quanh hit.",
    ]:
        add_bullet(doc, b)

    add_heading(doc, "4.8. File Extractors", level=2)
    add_body(doc, "Bốn module chuyển binary sang plain text dòng-by-dòng:")
    for b in [
        "PDF — dùng pdfplumber, đánh dấu [PAGE N] mỗi trang.",
        "DOCX — dùng python-docx, lấy paragraph + bảng.",
        "XLSX — dùng openpyxl, đánh dấu [SHEET name] mỗi sheet, mỗi row join bằng dấu |.",
        "PPTX — dùng python-pptx, đánh dấu [SLIDE N], gồm cả speaker notes.",
    ]:
        add_bullet(doc, b)

    add_heading(doc, "4.9. Cache và Drive Index", level=2)
    add_body(doc,
        "Hai dict in-memory giúp tối ưu hiệu năng:")
    for b in [
        "Cache text (cache.py) — sau khi extract một file, lưu text vào RAM. Lần grep sau không phải tải lại.",
        "Drive index (drive_index.py) — lưu mapping item_id → driveId, giúp fetch dùng đúng endpoint cho file shared / SharePoint / Teams.",
    ]:
        add_bullet(doc, b)

    # ====== 5. Quy trình end-to-end ======
    add_heading(doc, "5. Quy trình từ đầu đến cuối", level=1, color=DocxRGB(0x00, 0x78, 0xD4))

    steps = [
        ("Bước 1 — User nhập câu hỏi",
         "User gõ câu hỏi tự nhiên vào chat panel trên dashboard, ví dụ "
         "\"What is the metrics for RAG?\" và bấm Send."),
        ("Bước 2 — Browser gửi request",
         "JavaScript chặn submit form, gọi fetch POST /ask với JSON "
         "chứa câu hỏi. Chấp nhận response dạng text/event-stream."),
        ("Bước 3 — Flask khởi tạo agent",
         "Server lấy access_token từ session đã có sẵn (từ luồng OAuth2 "
         "với MSAL). Khởi tạo LangChain agent với 3 tools đã \"bind\" "
         "token qua closure. Mở stream SSE."),
        ("Bước 4 — LLM tự sinh keyword tìm kiếm",
         "LLM đọc câu hỏi và system prompt. Hệ thống prompt bắt buộc LLM "
         "phải dùng multi-pattern OR (không bao giờ search một từ đơn). "
         "LLM sinh tool call graph_search với query kiểu \"metric OR "
         "evaluation OR BLEU OR ROUGE OR recall\"."),
        ("Bước 5 — Microsoft Graph tìm file",
         "Tool graph_search gọi POST /search/query với entityTypes "
         "[driveItem]. Microsoft index trả về tối đa 15 file matching, "
         "mỗi file kèm snippet ~240 ký tự quanh keyword khớp."),
        ("Bước 6 — LLM đánh giá snippet",
         "LLM nhận lại danh sách file + snippet. Hai khả năng: "
         "(a) snippet đã đủ trả lời — chuyển sang Bước 9; "
         "(b) cần đọc sâu — chuyển sang Bước 7."),
        ("Bước 7 — Tải và extract file",
         "LLM gọi fetch_file_text cho 1-3 file phù hợp nhất. Tool tra "
         "drive_index để biết driveId, tải file qua endpoint Graph "
         "(redirect tự động đến CDN của Microsoft), extract text bằng "
         "extractor tương ứng, lưu cache. LLM nhận về preview ngắn để "
         "không nuốt context window."),
        ("Bước 8 — Grep context",
         "LLM gọi grep_context với regex multi-pattern và context_lines=5. "
         "Tool tìm match trên text đã cache, merge các window chồng lấp, "
         "trả về tối đa 10 chunk. Mỗi chunk có các dòng đánh dấu \">>\" "
         "(hit) và \"  \" (context). Đồng thời ghi toàn bộ vào debug/grep.log."),
        ("Bước 9 — LLM tổng hợp câu trả lời",
         "Sau khi đã có đủ ngữ cảnh, LLM sinh câu trả lời cuối kèm "
         "trích dẫn dạng [filename, L42] hoặc [filename, PAGE 3]. "
         "Nếu không tìm thấy, LLM phải nói rõ \"không tìm thấy\" — "
         "không được bịa."),
        ("Bước 10 — Stream về browser",
         "Mỗi bước (tool call, tool result, final answer, summary) được "
         "yield ra dưới dạng SSE event. Browser nhận và render từng event "
         "theo thời gian thực trên chat panel."),
    ]
    for h, t in steps:
        add_heading(doc, h, level=3)
        add_body(doc, t)

    # ====== 6. SSE ======
    add_heading(doc, "6. Server-Sent Events (SSE)", level=1, color=DocxRGB(0x00, 0x78, 0xD4))
    add_body(doc,
        "SSE cho phép server đẩy event liên tục về browser trên một "
        "kết nối HTTP duy nhất. Trong ứng dụng này, mỗi bước của agent "
        "(LLM quyết định gọi tool, tool trả về kết quả, LLM sinh câu trả "
        "lời cuối) được gửi ngay khi xảy ra. Browser hiển thị live giúp "
        "người dùng thấy được \"agent đang suy nghĩ gì\" mà không phải "
        "chờ toàn bộ quy trình kết thúc.")

    # ====== 7. Cache ======
    add_heading(doc, "7. Vai trò của cache và drive index", level=1, color=DocxRGB(0x00, 0x78, 0xD4))
    add_body(doc,
        "Mỗi file chỉ được tải về một lần. Sau khi extract, text được "
        "lưu vào dict in-memory keyed bằng item_id của Graph. Các lần "
        "grep_context tiếp theo đọc thẳng từ RAM, không gọi mạng. "
        "Drive index lưu mapping từ item_id sang driveId vì file trong "
        "OneDrive không phải lúc nào cũng thuộc /me/drive — có thể là "
        "file được share, từ SharePoint site, hoặc Teams group drive.")

    # ====== 8. So sánh ======
    add_heading(doc, "8. So sánh với RAG truyền thống", level=1, color=DocxRGB(0x00, 0x78, 0xD4))
    table = doc.add_table(rows=7, cols=3)
    table.style = "Light Grid Accent 1"
    headers = ["Tiêu chí", "RAG truyền thống", "Agentic Keyword (hệ thống này)"]
    rows = [
        ["Hạ tầng", "Vector DB riêng (FAISS, Pinecone...)", "Không cần — dùng index có sẵn của Microsoft"],
        ["Chunking", "Bắt buộc, cố định trước", "Không cần — context window động theo hit"],
        ["Exact match", "Có thể miss (semantic similarity)", "Hoàn hảo (regex literal)"],
        ["Cập nhật tài liệu", "Re-index toàn bộ", "Tức thì (Microsoft tự index)"],
        ["Chi phí", "Embedding + storage + LLM", "Chỉ LLM calls"],
        ["Minh bạch", "Khó debug similarity score", "Mọi bước log đầy đủ, trace được"],
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

    # ====== 9. Giới hạn ======
    add_heading(doc, "9. Giới hạn hiện tại", level=1, color=DocxRGB(0x00, 0x78, 0xD4))
    for b in [
        "PDF scan (image-only) — pdfplumber không OCR được, sẽ trả về text rỗng. Cần Azure Document Intelligence hoặc Docling cho hướng OCR.",
        "Giới hạn 25 MB mỗi file — file lớn hơn bị từ chối để tránh quá tải RAM và timeout.",
        "Tốc độ phụ thuộc số lần gọi LLM — mỗi câu hỏi thường tốn 5-15 lần gọi qua vòng lặp ReAct.",
        "Độ chính xác phụ thuộc chất lượng model — gpt-4.1-mini hoạt động tốt, model yếu hơn có thể không sinh đủ biến thể keyword.",
        "Search index lag của Microsoft — file vừa upload có thể mất vài phút mới được index.",
    ]:
        add_bullet(doc, b)

    # ====== 10. Kết luận ======
    add_heading(doc, "10. Kết luận", level=1, color=DocxRGB(0x00, 0x78, 0xD4))
    add_body(doc,
        "Hệ thống chứng minh rằng với một LLM đủ thông minh và bộ công cụ "
        "keyword search phù hợp, ta có thể đạt chất lượng tương đương RAG "
        "mà không cần xây dựng và vận hành vector database. Bằng cách tận "
        "dụng Microsoft Graph Search có sẵn và để LLM tự lái quá trình tìm "
        "kiếm, kiến trúc trở nên đơn giản, ít tốn kém, dễ debug, và hỗ trợ "
        "exact match hoàn hảo cho các tài liệu kỹ thuật.")

    # Footer
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
    make_pptx(os.path.join(here, "graph_grep.pptx"))
    make_docx(os.path.join(here, "docs_report.docx"))
