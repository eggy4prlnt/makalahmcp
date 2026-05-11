import os
import re
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from fpdf import FPDF


# =============================================================================
# Shared: Markdown parser
# =============================================================================

def _parse_markdown_to_blocks(content: str) -> list[dict]:
    """Parse markdown content into structured blocks."""
    blocks = []
    lines = content.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].strip()

        if not line:
            i += 1
            continue

        # Heading: # ## ###
        heading_match = re.match(r"^(#{1,3})\s+(.*)", line)
        if heading_match:
            level = len(heading_match.group(1))
            blocks.append({"type": "heading", "level": level, "text": heading_match.group(2)})
            i += 1
            continue

        # Image: ![caption](path)
        img_match = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)", line)
        if img_match:
            blocks.append({
                "type": "image",
                "caption": img_match.group(1),
                "path": img_match.group(2),
                "level": 0,
                "text": "",
            })
            i += 1
            continue

        # List item: - or *
        list_match = re.match(r"^[-*]\s+(.*)", line)
        if list_match:
            blocks.append({"type": "list_item", "level": 0, "text": list_match.group(1)})
            i += 1
            continue

        # Numbered list: 1. 2. etc
        num_list_match = re.match(r"^(\d+)\.\s+(.*)", line)
        if num_list_match:
            blocks.append({"type": "numbered_item", "level": 0, "text": num_list_match.group(2), "number": num_list_match.group(1)})
            i += 1
            continue

        # Regular paragraph - collect consecutive non-empty lines
        para_lines = [line]
        i += 1
        while i < len(lines) and lines[i].strip() and not re.match(r"^(#{1,3}\s|!\[|[-*]\s|\d+\.\s)", lines[i].strip()):
            para_lines.append(lines[i].strip())
            i += 1
        blocks.append({"type": "paragraph", "level": 0, "text": " ".join(para_lines)})

    return blocks


def _extract_headings(blocks: list[dict]) -> list[dict]:
    """Extract headings from blocks for table of contents."""
    return [b for b in blocks if b["type"] == "heading"]


# =============================================================================
# DOCX generation
# =============================================================================

def _setup_document() -> Document:
    """Create a new Document with proper page setup."""
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(3.5)
    section.right_margin = Cm(3.0)
    section.top_margin = Cm(3.0)
    section.bottom_margin = Cm(3.0)

    # Set default font
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Times New Roman"
    font.size = Pt(12)
    pf = style.paragraph_format
    pf.line_spacing = 1.5

    return doc


def _add_centered_run(doc, text, bold=False, italic=False, size=12, spacing_after=0):
    """Helper to add a centered paragraph with styled run."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(spacing_after)
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    return p


def _add_cover_page(doc, title, title_en="", lecturer="", author="",
                     nim="", program_studi="", fakultas="",
                     universitas="", year="", logo_path=""):
    """Add a formatted cover page."""
    # "MAKALAH" header
    _add_centered_run(doc, "MAKALAH", bold=True, size=14, spacing_after=12)

    # Title (Indonesian) - uppercase, bold, 14pt
    _add_centered_run(doc, title.upper(), bold=True, size=14, spacing_after=6)

    # Title (English) - italic, 14pt
    if title_en:
        _add_centered_run(doc, title_en, italic=True, size=14, spacing_after=6)

    # Lecturer
    if lecturer:
        _add_centered_run(doc, f"Dosen Pengampu : {lecturer}", size=12, spacing_after=6)

    # Spacing before logo
    doc.add_paragraph()

    # Logo
    if logo_path and os.path.exists(logo_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run()
        run.add_picture(logo_path, width=Cm(4))

    # Spacing
    doc.add_paragraph()

    # "Di susun oleh:"
    _add_centered_run(doc, "Di susun oleh:", size=12, spacing_after=6)

    # Author info in table-like format
    if author or nim or lecturer:
        # Use a table for alignment like the example
        table = doc.add_table(rows=0, cols=2)
        table.alignment = WD_ALIGN_PARAGRAPH.CENTER

        def _add_info_row(label, value):
            row = table.add_row()
            cell_label = row.cells[0]
            cell_value = row.cells[1]
            cell_label.width = Cm(4)
            cell_value.width = Cm(7)

            p_label = cell_label.paragraphs[0]
            run_l = p_label.add_run(label)
            run_l.font.name = "Times New Roman"
            run_l.font.size = Pt(12)

            p_value = cell_value.paragraphs[0]
            run_v = p_value.add_run(f": {value}")
            run_v.font.name = "Times New Roman"
            run_v.font.size = Pt(12)

        if author:
            _add_info_row("Nama", author)
        if nim:
            _add_info_row("Nim", nim)
        if lecturer:
            _add_info_row("Dosen pengampu", lecturer)

        # Remove table borders
        from docx.oxml.ns import qn
        for row in table.rows:
            for cell in row.cells:
                tc = cell._tc
                tcPr = tc.get_or_add_tcPr()
                tcBorders = tcPr.find(qn('w:tcBorders'))
                if tcBorders is not None:
                    tcPr.remove(tcBorders)

    # Spacing
    doc.add_paragraph()
    doc.add_paragraph()

    # Program Studi
    if program_studi:
        _add_centered_run(doc, program_studi.upper(), bold=True, size=12, spacing_after=4)
    if fakultas:
        _add_centered_run(doc, fakultas.upper(), bold=True, size=12, spacing_after=4)
    if universitas:
        _add_centered_run(doc, universitas.upper(), bold=True, size=12, spacing_after=4)
    if year:
        _add_centered_run(doc, year, size=12)

    doc.add_page_break()


def _add_kata_pengantar(doc, title, author="", nim="", city="", year=""):
    """Add Kata Pengantar page."""
    _add_centered_run(doc, "KATA PENGANTAR", bold=True, size=14, spacing_after=12)

    # Standard kata pengantar text
    text1 = (
        "Puji syukur kehadirat Tuhan Yang Maha Esa atas segala rahmatNYA "
        "sehingga makalah ini dapat tersusun hingga selesai. Tidak lupa kami juga "
        "mengucapkan banyak terimakasih atas bantuan dari pihak yang telah "
        "berkontribusi dengan memberikan sumbangan baik materi maupun pikirannya."
    )
    text2 = (
        f"Dan harapan kami semoga makalah tentang \"{title}\" ini dapat menambah pengetahuan "
        "dan pengalaman bagi para pembaca. Untuk ke depannya dapat memperbaiki "
        "bentuk maupun menambah isi makalah agar menjadi lebih baik lagi."
    )
    text3 = (
        "Karena keterbatasan pengetahuan maupun pengalaman kami, Kami "
        "yakin masih banyak kekurangan dalam makalah ini. Oleh karena itu kami "
        "sangat mengharapkan saran dan kritik yang membangun dari pembaca demi "
        "kesempurnaan makalah ini."
    )

    for text in [text1, text2, text3]:
        p = doc.add_paragraph(text)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(1.25)
        p.paragraph_format.line_spacing = 1.5
        for run in p.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(12)

    # Date and author
    doc.add_paragraph()
    doc.add_paragraph()

    date_str = f"{city}, {datetime.now().strftime('%B')} {year}" if city else f"{datetime.now().strftime('%B')} {year}"
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(date_str)
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run("Penyusun")
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)

    doc.add_paragraph()
    doc.add_paragraph()

    if author:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run(author)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        run.underline = True

    doc.add_page_break()


def _add_daftar_isi(doc, headings):
    """Add Daftar Isi (Table of Contents) page."""
    _add_centered_run(doc, "DAFTAR ISI", bold=True, size=14, spacing_after=12)

    # Front matter entries
    front_items = ["KATA PENGANTAR", "DAFTAR ISI"]
    for item in front_items:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.line_spacing = 1.5
        run = p.add_run(item)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)

    # Content headings
    for h in headings:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.5

        text = h["text"]
        level = h["level"]

        if level == 1:
            # BAB heading - no indent
            run = p.add_run(text.upper())
            run.bold = True
        elif level == 2:
            # Sub-heading - slight indent
            p.paragraph_format.left_indent = Cm(1.0)
            run = p.add_run(text)
        else:
            # Sub-sub-heading
            p.paragraph_format.left_indent = Cm(2.0)
            run = p.add_run(text)

        run.font.name = "Times New Roman"
        run.font.size = Pt(12)

    doc.add_page_break()


def _render_blocks_to_doc(doc, blocks, image_counter=1):
    """Render parsed blocks into the document. Returns updated image counter."""
    for block in blocks:
        if block["type"] == "heading":
            level = block["level"]
            text = block["text"]

            if level == 1:
                # Check if it's a BAB heading (starts with roman numeral pattern)
                bab_match = re.match(r"^([IVX]+)\.\s*(.*)", text)
                if bab_match:
                    # BAB number centered
                    _add_centered_run(doc, f"BAB {bab_match.group(1)}", bold=True, size=14, spacing_after=6)
                    # BAB title centered
                    if bab_match.group(2):
                        _add_centered_run(doc, bab_match.group(2).upper(), bold=True, size=14, spacing_after=12)
                else:
                    # Non-BAB heading (like "Daftar Pustaka")
                    _add_centered_run(doc, text.upper(), bold=True, size=14, spacing_after=12)
            else:
                # Sub-headings: left-aligned, bold
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(12)
                p.paragraph_format.space_after = Pt(6)
                p.paragraph_format.line_spacing = 1.5
                run = p.add_run(text)
                run.bold = True
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)

        elif block["type"] == "paragraph":
            p = doc.add_paragraph(block["text"])
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(1.25)
            p.paragraph_format.line_spacing = 1.5
            for run in p.runs:
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)

        elif block["type"] == "image":
            path = block["path"]
            if os.path.exists(path):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run()
                run.add_picture(path, width=Cm(12))

                caption = block["caption"] or f"Gambar {image_counter}"
                p_cap = doc.add_paragraph()
                p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run_cap = p_cap.add_run(f"Gambar {image_counter}: {caption}")
                run_cap.font.name = "Times New Roman"
                run_cap.font.size = Pt(10)
                run_cap.italic = True
                image_counter += 1

        elif block["type"] == "list_item":
            p = doc.add_paragraph(block["text"], style="List Bullet")
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.5
            for run in p.runs:
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)

        elif block["type"] == "numbered_item":
            p = doc.add_paragraph(block["text"], style="List Number")
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.5
            for run in p.runs:
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)

    return image_counter


def save_as_docx(content, title, output_path, title_en="", lecturer="",
                  author="", nim="", program_studi="", fakultas="",
                  universitas="", year="", logo_path=""):
    """Convert markdown content to a styled DOCX file. Returns the file path."""
    if not year:
        year = str(datetime.now().year)

    doc = _setup_document()

    # 1. Cover page
    _add_cover_page(doc, title=title, title_en=title_en, lecturer=lecturer,
                     author=author, nim=nim, program_studi=program_studi,
                     fakultas=fakultas, universitas=universitas,
                     year=year, logo_path=logo_path)

    # 2. Kata Pengantar
    _add_kata_pengantar(doc, title=title, author=author, nim=nim, year=year)

    # 3. Daftar Isi
    blocks = _parse_markdown_to_blocks(content)
    headings = _extract_headings(blocks)
    _add_daftar_isi(doc, headings)

    # 4. Content
    _render_blocks_to_doc(doc, blocks)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    doc.save(output_path)
    return output_path


# =============================================================================
# PDF generation
# =============================================================================

class MakalahPDF(FPDF):
    """Custom PDF class for makalah formatting."""

    def __init__(self):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.set_margins(35, 30, 30)
        self.set_auto_page_break(auto=True, margin=30)
        self._page_number_style = "roman"  # "roman" for front matter, "arabic" for content
        self._content_start_page = 1

    def footer(self):
        """Add page numbers at bottom center."""
        self.set_y(-15)
        self.set_font("Times", "I", 10)
        page = self.page_no()
        if self._page_number_style == "roman" and page < self._content_start_page:
            # Roman numeral for front matter
            roman_map = {1: "i", 2: "ii", 3: "iii", 4: "iv", 5: "v",
                         6: "vi", 7: "vii", 8: "viii", 9: "ix", 10: "x"}
            # Cover has no page number, kata pengantar starts at ii
            front_page = page - 1  # subtract cover page
            if front_page > 0:
                num_str = roman_map.get(front_page, str(front_page))
                self.cell(0, 10, num_str, align="C")
        else:
            content_page = page - self._content_start_page + 1
            if content_page > 0:
                self.cell(0, 10, str(content_page), align="C")

    def _set_font_safe(self, family="Times", style="", size=12):
        """Set font with fallback."""
        try:
            self.set_font(family, style, size)
        except Exception:
            self.set_font("Times", style, size)


def _build_cover_pdf(pdf, title, title_en="", lecturer="", author="",
                      nim="", program_studi="", fakultas="",
                      universitas="", year="", logo_path=""):
    """Add cover page to PDF."""
    pdf.add_page()
    pdf.ln(15)

    # "MAKALAH"
    pdf._set_font_safe("Times", "B", 14)
    pdf.cell(0, 8, "MAKALAH", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(8)

    # Title - uppercase, bold
    pdf._set_font_safe("Times", "B", 14)
    pdf.multi_cell(0, 8, title.upper(), align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    # English title - italic
    if title_en:
        pdf._set_font_safe("Times", "I", 14)
        pdf.multi_cell(0, 8, title_en, align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    # Lecturer line
    if lecturer:
        pdf._set_font_safe("Times", "", 12)
        pdf.multi_cell(0, 7, f"Dosen Pengampu : {lecturer}", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(8)

    # Logo
    if logo_path and os.path.exists(logo_path):
        page_w = pdf.w - pdf.l_margin - pdf.r_margin
        img_w = 40
        x = pdf.l_margin + (page_w - img_w) / 2
        pdf.image(logo_path, x=x, w=img_w)
        pdf.ln(8)

    pdf.ln(8)

    # "Di susun oleh:"
    pdf._set_font_safe("Times", "", 12)
    pdf.cell(0, 7, "Di susun oleh:", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    # Author info
    info_x = pdf.l_margin + 20
    if author:
        pdf._set_font_safe("Times", "", 12)
        pdf.set_x(info_x)
        pdf.cell(35, 7, "Nama")
        pdf.cell(0, 7, f": {author}", new_x="LMARGIN", new_y="NEXT")
    if nim:
        pdf.set_x(info_x)
        pdf.cell(35, 7, "Nim")
        pdf.cell(0, 7, f": {nim}", new_x="LMARGIN", new_y="NEXT")
    if lecturer:
        pdf.set_x(info_x)
        pdf.cell(35, 7, "Dosen pengampu")
        pdf.cell(0, 7, f": {lecturer}", new_x="LMARGIN", new_y="NEXT")

    pdf.ln(15)

    # Institution
    if program_studi:
        pdf._set_font_safe("Times", "B", 12)
        pdf.cell(0, 7, program_studi.upper(), align="C", new_x="LMARGIN", new_y="NEXT")
    if fakultas:
        pdf._set_font_safe("Times", "B", 12)
        pdf.cell(0, 7, fakultas.upper(), align="C", new_x="LMARGIN", new_y="NEXT")
    if universitas:
        pdf._set_font_safe("Times", "B", 12)
        pdf.cell(0, 7, universitas.upper(), align="C", new_x="LMARGIN", new_y="NEXT")
    if year:
        pdf._set_font_safe("Times", "", 12)
        pdf.ln(3)
        pdf.cell(0, 7, year, align="C", new_x="LMARGIN", new_y="NEXT")


def _build_kata_pengantar_pdf(pdf, title, author="", year=""):
    """Add Kata Pengantar page to PDF."""
    pdf.add_page()

    pdf._set_font_safe("Times", "B", 14)
    pdf.cell(0, 8, "KATA PENGANTAR", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(10)

    texts = [
        "Puji syukur kehadirat Tuhan Yang Maha Esa atas segala rahmatNYA "
        "sehingga makalah ini dapat tersusun hingga selesai. Tidak lupa kami juga "
        "mengucapkan banyak terimakasih atas bantuan dari pihak yang telah "
        "berkontribusi dengan memberikan sumbangan baik materi maupun pikirannya.",

        f"Dan harapan kami semoga makalah tentang \"{title}\" ini dapat menambah pengetahuan "
        "dan pengalaman bagi para pembaca. Untuk ke depannya dapat memperbaiki "
        "bentuk maupun menambah isi makalah agar menjadi lebih baik lagi.",

        "Karena keterbatasan pengetahuan maupun pengalaman kami, Kami "
        "yakin masih banyak kekurangan dalam makalah ini. Oleh karena itu kami "
        "sangat mengharapkan saran dan kritik yang membangun dari pembaca demi "
        "kesempurnaan makalah ini.",
    ]

    pdf._set_font_safe("Times", "", 12)
    for text in texts:
        # Indent first line
        pdf.set_x(pdf.l_margin + 12.5)
        pdf.multi_cell(0, 7, text, align="J", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(3)

    pdf.ln(15)

    # Date and author (right-aligned)
    date_str = f"{datetime.now().strftime('%B')} {year}"
    pdf._set_font_safe("Times", "", 12)
    pdf.cell(0, 7, date_str, align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(8)
    pdf.cell(0, 7, "Penyusun", align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(15)
    if author:
        pdf._set_font_safe("Times", "U", 12)
        pdf.cell(0, 7, author, align="R", new_x="LMARGIN", new_y="NEXT")


def _build_daftar_isi_pdf(pdf, headings):
    """Add Daftar Isi page to PDF."""
    pdf.add_page()

    pdf._set_font_safe("Times", "B", 14)
    pdf.cell(0, 8, "DAFTAR ISI", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(10)

    pdf._set_font_safe("Times", "", 12)

    # Front matter
    pdf.cell(0, 7, "KATA PENGANTAR", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 7, "DAFTAR ISI", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    for h in headings:
        level = h["level"]
        text = h["text"]

        if level == 1:
            pdf._set_font_safe("Times", "B", 12)
            pdf.cell(0, 7, text.upper(), new_x="LMARGIN", new_y="NEXT")
        elif level == 2:
            pdf._set_font_safe("Times", "", 12)
            pdf.set_x(pdf.l_margin + 10)
            pdf.cell(0, 7, text, new_x="LMARGIN", new_y="NEXT")
        else:
            pdf._set_font_safe("Times", "", 12)
            pdf.set_x(pdf.l_margin + 20)
            pdf.cell(0, 7, text, new_x="LMARGIN", new_y="NEXT")


def _render_blocks_to_pdf(pdf, blocks, image_counter=1):
    """Render parsed markdown blocks into PDF. Returns updated image counter."""
    for block in blocks:
        if block["type"] == "heading":
            level = block["level"]
            text = block["text"]

            if level == 1:
                bab_match = re.match(r"^([IVX]+)\.\s*(.*)", text)
                if bab_match:
                    # New page for each BAB
                    pdf.add_page()
                    pdf._set_font_safe("Times", "B", 14)
                    pdf.cell(0, 8, f"BAB {bab_match.group(1)}", align="C", new_x="LMARGIN", new_y="NEXT")
                    pdf.ln(3)
                    if bab_match.group(2):
                        pdf.cell(0, 8, bab_match.group(2).upper(), align="C", new_x="LMARGIN", new_y="NEXT")
                    pdf.ln(8)
                else:
                    pdf.add_page()
                    pdf._set_font_safe("Times", "B", 14)
                    pdf.cell(0, 8, text.upper(), align="C", new_x="LMARGIN", new_y="NEXT")
                    pdf.ln(8)
            else:
                pdf._set_font_safe("Times", "B", 12)
                pdf.ln(5)
                pdf.multi_cell(0, 7, text, new_x="LMARGIN", new_y="NEXT")
                pdf.ln(3)

        elif block["type"] == "paragraph":
            pdf._set_font_safe("Times", "", 12)
            # First line indent
            pdf.set_x(pdf.l_margin + 12.5)
            pdf.multi_cell(0, 7, block["text"], align="J", new_x="LMARGIN", new_y="NEXT")
            pdf.ln(3)

        elif block["type"] == "image":
            path = block["path"]
            if os.path.exists(path):
                page_w = pdf.w - pdf.l_margin - pdf.r_margin
                img_w = min(120, page_w)
                x = pdf.l_margin + (page_w - img_w) / 2
                pdf.image(path, x=x, w=img_w)
                pdf.ln(3)

                caption = block["caption"] or f"Gambar {image_counter}"
                pdf._set_font_safe("Times", "I", 10)
                pdf.cell(0, 6, f"Gambar {image_counter}: {caption}", align="C", new_x="LMARGIN", new_y="NEXT")
                pdf.ln(3)
                image_counter += 1

        elif block["type"] == "list_item":
            pdf._set_font_safe("Times", "", 12)
            pdf.set_x(pdf.l_margin + 10)
            pdf.multi_cell(0, 7, f"- {block['text']}", align="J", new_x="LMARGIN", new_y="NEXT")

        elif block["type"] == "numbered_item":
            pdf._set_font_safe("Times", "", 12)
            pdf.set_x(pdf.l_margin + 10)
            pdf.multi_cell(0, 7, f"{block['number']}. {block['text']}", align="J", new_x="LMARGIN", new_y="NEXT")

    return image_counter


def save_as_pdf(content, title, output_path, title_en="", lecturer="",
                 author="", nim="", program_studi="", fakultas="",
                 universitas="", year="", logo_path=""):
    """Convert markdown content to a styled PDF file. Returns the file path."""
    if not year:
        year = str(datetime.now().year)

    pdf = MakalahPDF()

    # 1. Cover page (no page number)
    _build_cover_pdf(pdf, title=title, title_en=title_en, lecturer=lecturer,
                      author=author, nim=nim, program_studi=program_studi,
                      fakultas=fakultas, universitas=universitas,
                      year=year, logo_path=logo_path)

    # 2. Kata Pengantar
    _build_kata_pengantar_pdf(pdf, title=title, author=author, year=year)

    # 3. Daftar Isi
    blocks = _parse_markdown_to_blocks(content)
    headings = _extract_headings(blocks)
    _build_daftar_isi_pdf(pdf, headings)

    # Mark where content starts for page numbering
    pdf._content_start_page = pdf.page_no() + 1
    pdf._page_number_style = "arabic"

    # 4. Content
    _render_blocks_to_pdf(pdf, blocks)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    pdf.output(output_path)
    return output_path
