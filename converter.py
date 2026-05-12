import os
import re
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from fpdf import FPDF


# =============================================================================
# Shared: Markdown parser
# =============================================================================

def _strip_md_formatting(text: str) -> str:
    """Remove markdown bold/italic markers from text."""
    text = re.sub(r'\*\*\*(.+?)\*\*\*', r'\1', text)  # ***bold italic***
    text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)      # **bold**
    text = re.sub(r'\*(.+?)\*', r'\1', text)           # *italic*
    text = re.sub(r'__(.+?)__', r'\1', text)           # __bold__
    text = re.sub(r'_(.+?)_', r'\1', text)             # _italic_
    return text


def _parse_markdown_to_blocks(content: str) -> list[dict]:
    """Parse markdown content into structured blocks."""
    blocks = []
    lines = content.split("\n")
    i = 0
    in_daftar_pustaka = False
    found_first_heading = False  # Skip ALL content before first heading

    while i < len(lines):
        line = lines[i].strip()
        if not line:
            i += 1
            continue

        heading_match = re.match(r"^(#{1,6})\s+(.*)", line)
        if heading_match:
            level = min(len(heading_match.group(1)), 3)
            text = heading_match.group(2)
            clean_text = _strip_md_formatting(text)
            # Force level 1 for BAB headings and Daftar Pustaka regardless of # count
            if _parse_bab_heading(clean_text) or "daftar pustaka" in clean_text.lower():
                level = 1
            # Skip front matter headings
            if clean_text.lower().strip() not in _SKIP_HEADINGS:
                blocks.append({"type": "heading", "level": level, "text": clean_text})
                found_first_heading = True
            in_daftar_pustaka = "daftar pustaka" in text.lower()
            i += 1
            continue

        # SKIP everything before the first real heading
        if not found_first_heading:
            i += 1
            continue

        img_match = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)", line)
        if img_match:
            blocks.append({"type": "image", "caption": img_match.group(1), "path": img_match.group(2), "level": 0, "text": ""})
            i += 1
            continue

        list_match = re.match(r"^[-*]\s+(.*)", line)
        if list_match:
            blocks.append({"type": "list_item", "level": 0, "text": _strip_md_formatting(list_match.group(1))})
            i += 1
            continue

        num_list_match = re.match(r"^(\d+)\.\s+(.*)", line)
        if num_list_match:
            blocks.append({"type": "numbered_item", "level": 0, "text": _strip_md_formatting(num_list_match.group(2)), "number": num_list_match.group(1)})
            i += 1
            continue

        if in_daftar_pustaka:
            blocks.append({"type": "reference", "level": 0, "text": line})
            i += 1
            continue

        # Regular paragraph
        para_lines = [line]
        i += 1
        while i < len(lines) and lines[i].strip() and not re.match(r"^(#{1,6}\s|!\[|[-*]\s|\d+\.\s)", lines[i].strip()):
            para_lines.append(lines[i].strip())
            i += 1
        para_text = _strip_md_formatting(" ".join(para_lines))
        blocks.append({"type": "paragraph", "level": 0, "text": para_text})

    return blocks


_SKIP_HEADINGS = {"kata pengantar", "daftar isi", "cover"}

# Patterns that indicate kata pengantar content to skip
_KATA_PENGANTAR_PATTERNS = [
    "puji syukur", "puji dan syukur", "segala rahmat", "kehadirat tuhan",
    "kehadirat allah", "panjatkan kehadirat", "panjatkan puji",
    "makalah ini disusun untuk memenuhi", "makalah ini disusun sebagai",
    "penulis menyadari bahwa makalah ini", "penulis mengucapkan terima kasih",
    "penulis mengucapkan terimakasih", "kritik dan saran yang membangun",
    "saran dan kritik yang membangun", "jauh dari sempurna",
    "kesempurnaan makalah ini", "demi perbaikan di masa",
    "terima kasih kepada semua pihak", "terimakasih kepada semua pihak",
    "bantuan dari pihak yang telah", "semoga makalah ini dapat",
    "makalah ini dapat tersusun", "menyelesaikan makalah",
    "membantu dalam penyusunan", "bahan pembelajaran",
]

# Short standalone lines to skip (exact or near-exact match)
_SKIP_SHORT_LINES = {"penulis", "penyusun", "hormat kami", "wassalam"}

_BULAN_ID = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}


def _tanggal_indonesia():
    """Return current date in Indonesian format: 'Kota, DD Bulan YYYY'"""
    now = datetime.now()
    return f"{now.day} {_BULAN_ID[now.month]} {now.year}"


def _extract_headings(blocks):
    return [b for b in blocks if b["type"] == "heading"
            and b["text"].lower().strip() not in _SKIP_HEADINGS]


def _make_bookmark_id(text):
    return f"_bm_{re.sub(r'[^a-zA-Z0-9]', '_', text)[:40]}"


# =============================================================================
# DOCX
# =============================================================================

def _setup_document():
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(3.5)
    section.right_margin = Cm(3.0)
    section.top_margin = Cm(3.0)
    section.bottom_margin = Cm(3.0)

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    style.paragraph_format.line_spacing = 1.5
    style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    return doc


def _centered(doc, text, bold=False, italic=False, size=12, after=0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(0)
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    return p


def _add_bookmark(paragraph, name):
    tag = paragraph._p
    bid = str(abs(hash(name)) % 100000)
    bs = OxmlElement('w:bookmarkStart')
    bs.set(qn('w:id'), bid)
    bs.set(qn('w:name'), name)
    be = OxmlElement('w:bookmarkEnd')
    be.set(qn('w:id'), bid)
    tag.insert(0, bs)
    tag.append(be)


# --- Cover ---
def _add_cover_page(doc, title, title_en="", lecturer="", author="",
                     nim="", program_studi="", fakultas="",
                     universitas="", year="", logo_path=""):
    _centered(doc, "MAKALAH", bold=True, size=14, after=18)
    _centered(doc, title.upper(), bold=True, size=14, after=6)
    if title_en:
        _centered(doc, title_en, italic=True, size=14, after=6)
    if lecturer:
        _centered(doc, f"Dosen Pengampu : {lecturer}", size=12, after=6)

    doc.add_paragraph()  # spacing

    if logo_path and os.path.exists(logo_path):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(logo_path, width=Cm(4))

    doc.add_paragraph()  # spacing
    _centered(doc, "Disusun Oleh :", size=12, after=12)

    if author and nim:
        _centered(doc, f"{author} {nim}", size=12, after=4)
    elif author:
        _centered(doc, author, size=12, after=4)

    doc.add_paragraph()
    doc.add_paragraph()

    if program_studi:
        _centered(doc, program_studi.upper(), bold=True, size=12, after=4)
    if fakultas:
        _centered(doc, fakultas.upper(), bold=True, size=12, after=4)
    if universitas:
        _centered(doc, universitas.upper(), bold=True, size=12, after=4)
    if year:
        _centered(doc, year, size=12)

    doc.add_page_break()


# --- Kata Pengantar ---
def _add_kata_pengantar(doc, title, author="", nim="", year=""):
    p = _centered(doc, "KATA PENGANTAR", bold=True, size=14, after=18)
    _add_bookmark(p, "_bm_kata_pengantar")

    texts = [
        "Puji syukur kehadirat Tuhan Yang Maha Esa atas segala rahmat-Nya "
        "sehingga makalah ini dapat tersusun hingga selesai . Tidak lupa kami juga "
        "mengucapkan banyak terimakasih atas bantuan dari pihak yang telah "
        "berkontribusi dengan memberikan sumbangan baik materi maupun pikirannya.",

        f"Dan harapan kami semoga makalah ini dapat menambah pengetahuan "
        "dan pengalaman bagi para pembaca, Untuk ke depannya dapat memperbaiki "
        "bentuk maupun menambah isi makalah agar menjadi lebih baik lagi.",

        "Karena keterbatasan pengetahuan maupun pengalaman kami, Kami "
        "yakin masih banyak kekurangan dalam makalah ini, Oleh karena itu kami "
        "sangat mengharapkan saran dan kritik yang membangun dari pembaca demi "
        "kesempurnaan makalah ini.",
    ]
    for t in texts:
        p = doc.add_paragraph(t)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(1.25)
        p.paragraph_format.line_spacing = 1.5
        for run in p.runs:
            run.font.name = "Times New Roman"
            run.font.size = Pt(12)

    doc.add_paragraph()
    doc.add_paragraph()

    # Right-aligned date
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run(f"{_tanggal_indonesia()}")
    r.font.name = "Times New Roman"; r.font.size = Pt(12)

    doc.add_paragraph()

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = p.add_run("Penyusun")
    r.font.name = "Times New Roman"; r.font.size = Pt(12)

    doc.add_paragraph()
    doc.add_paragraph()

    if author:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run(author)
        r.font.name = "Times New Roman"; r.font.size = Pt(12); r.underline = True
    if nim:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r = p.add_run(f"Nim. {nim}")
        r.font.name = "Times New Roman"; r.font.size = Pt(12)

    doc.add_page_break()


# --- Daftar Isi ---
def _toc_entry(doc, text, page_str, indent_cm=0, bold=False, bookmark=None):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    if indent_cm > 0:
        p.paragraph_format.left_indent = Cm(indent_cm)

    tab_pos = 14.5 - indent_cm
    pPr = p._p.get_or_add_pPr()
    tabs = OxmlElement('w:tabs')
    tab = OxmlElement('w:tab')
    tab.set(qn('w:val'), 'right')
    tab.set(qn('w:leader'), 'dot')
    tab.set(qn('w:pos'), str(int(tab_pos * 567)))
    tabs.append(tab)
    pPr.append(tabs)

    if bookmark:
        hl = OxmlElement('w:hyperlink')
        hl.set(qn('w:anchor'), bookmark)
        r_el = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')
        rF = OxmlElement('w:rFonts')
        rF.set(qn('w:ascii'), 'Times New Roman'); rF.set(qn('w:hAnsi'), 'Times New Roman')
        rPr.append(rF)
        sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '24'); rPr.append(sz)
        if bold:
            rPr.append(OxmlElement('w:b'))
        col = OxmlElement('w:color'); col.set(qn('w:val'), '000000'); rPr.append(col)
        r_el.append(rPr)
        t = OxmlElement('w:t'); t.text = text; t.set(qn('xml:space'), 'preserve')
        r_el.append(t)
        hl.append(r_el)
        # tab
        tab_r = OxmlElement('w:r'); tab_r.append(OxmlElement('w:tab')); hl.append(tab_r)
        # page number
        pg_r = OxmlElement('w:r')
        pg_rPr = OxmlElement('w:rPr')
        pf = OxmlElement('w:rFonts'); pf.set(qn('w:ascii'), 'Times New Roman'); pf.set(qn('w:hAnsi'), 'Times New Roman')
        pg_rPr.append(pf)
        psz = OxmlElement('w:sz'); psz.set(qn('w:val'), '24'); pg_rPr.append(psz)
        pg_r.append(pg_rPr)
        pg_t = OxmlElement('w:t'); pg_t.text = page_str; pg_r.append(pg_t)
        hl.append(pg_r)
        p._p.append(hl)
    else:
        run = p.add_run(text); run.bold = bold
        run.font.name = "Times New Roman"; run.font.size = Pt(12)
        p.add_run("\t").font.name = "Times New Roman"
        run2 = p.add_run(page_str)
        run2.font.name = "Times New Roman"; run2.font.size = Pt(12)


def _toc_label(doc, text):
    """BAB heading line in TOC — no dots, no page number, not bold."""
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)


def _add_daftar_isi(doc, headings):
    _centered(doc, "DAFTAR ISI", bold=True, size=14, after=18)

    # Fixed front matter
    p = doc.add_paragraph("Cover")
    p.paragraph_format.space_after = Pt(0)
    for r in p.runs:
        r.font.name = "Times New Roman"; r.font.size = Pt(12)

    _toc_entry(doc, "KATA PENGANTAR", "ii", bookmark="_bm_kata_pengantar")
    _toc_entry(doc, "DAFTAR ISI", "iii")

    # Content headings
    current_page = 1
    for h in headings:
        text = h["text"]
        level = h["level"]
        bm = _make_bookmark_id(text)

        if level == 1:
            bab = _parse_bab_heading(text)
            is_dafpus = "daftar pustaka" in text.lower()
            if is_dafpus:
                _toc_entry(doc, text.upper(), str(current_page), bookmark=bm)
                current_page += 1
            elif bab:
                bab_num, bab_title = bab
                label = f"BAB {bab_num} {bab_title.upper()}" if bab_title else f"BAB {bab_num}"
                _toc_label(doc, label)
                current_page += 2
            else:
                _toc_label(doc, text.upper())
                current_page += 2
        elif level == 2:
            _toc_entry(doc, text, str(current_page), indent_cm=1.0, bookmark=bm)
            current_page += 1
        else:
            _toc_entry(doc, text, str(current_page), indent_cm=2.5, bookmark=bm)


# --- Content ---
def _parse_bab_heading(text):
    """Parse BAB heading text. Returns (bab_number, bab_title) or None."""
    # Match "I. Pendahuluan" or "II. Pembahasan"
    m = re.match(r"^([IVX]+)\.\s*(.*)", text)
    if m:
        return m.group(1), m.group(2)
    # Match "BAB I PENDAHULUAN" or "BAB II Pembahasan"
    m = re.match(r"^BAB\s+([IVX]+)\s*(.*)", text, re.IGNORECASE)
    if m:
        return m.group(1), m.group(2)
    return None


def _render_blocks_to_doc(doc, blocks, image_counter=1):
    for block in blocks:
        if block["type"] == "heading":
            level = block["level"]
            text = block["text"]

            if text.lower().strip() in _SKIP_HEADINGS:
                continue

            bm = _make_bookmark_id(text)

            if level == 1:
                # ALWAYS page break before every BAB/h1
                doc.add_page_break()
                bab = _parse_bab_heading(text)
                if bab:
                    bab_num, bab_title = bab
                    p = _centered(doc, f"BAB {bab_num}", bold=True, size=14, after=6)
                    _add_bookmark(p, bm)
                    if bab_title:
                        _centered(doc, bab_title.upper(), bold=True, size=14, after=18)
                else:
                    p = _centered(doc, text.upper(), bold=True, size=14, after=18)
                    _add_bookmark(p, bm)
            else:
                # Sub-heading: "1.1    Latar belakang masalah" — left-aligned, bold
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(12)
                p.paragraph_format.space_after = Pt(6)
                p.paragraph_format.line_spacing = 1.5
                run = p.add_run(text)
                run.bold = True
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)
                _add_bookmark(p, bm)

        elif block["type"] == "paragraph":
            p = doc.add_paragraph(block["text"])
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(1.25)
            p.paragraph_format.line_spacing = 1.5
            for run in p.runs:
                run.font.name = "Times New Roman"; run.font.size = Pt(12)

        elif block["type"] == "image":
            path = block["path"]
            if os.path.exists(path):
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(path, width=Cm(12))
                cap = block["caption"] or ""
                # Avoid duplicate "Gambar N:" prefix
                if re.match(r"^Gambar\s+\d+", cap, re.IGNORECASE):
                    cap_text = cap
                else:
                    cap_text = f"Gambar {image_counter}: {cap}" if cap else f"Gambar {image_counter}"
                pc = doc.add_paragraph()
                pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = pc.add_run(cap_text)
                r.font.name = "Times New Roman"; r.font.size = Pt(10); r.italic = True
                image_counter += 1

        elif block["type"] == "list_item":
            p = doc.add_paragraph(block["text"], style="List Bullet")
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.5
            for run in p.runs:
                run.font.name = "Times New Roman"; run.font.size = Pt(12)

        elif block["type"] == "numbered_item":
            p = doc.add_paragraph(block["text"], style="List Number")
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.5
            for run in p.runs:
                run.font.name = "Times New Roman"; run.font.size = Pt(12)

        elif block["type"] == "reference":
            # Daftar pustaka entry: hanging indent (first line flush, rest indented)
            p = doc.add_paragraph(block["text"])
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.left_indent = Cm(1.0)
            p.paragraph_format.first_line_indent = Cm(-1.0)  # hanging indent
            p.paragraph_format.space_after = Pt(6)
            for run in p.runs:
                run.font.name = "Times New Roman"; run.font.size = Pt(12)

    return image_counter


def save_as_docx(content, title, output_path, title_en="", lecturer="",
                  author="", nim="", program_studi="", fakultas="",
                  universitas="", year="", logo_path=""):
    if not year:
        year = str(datetime.now().year)
    doc = _setup_document()
    _add_cover_page(doc, title=title, title_en=title_en, lecturer=lecturer,
                     author=author, nim=nim, program_studi=program_studi,
                     fakultas=fakultas, universitas=universitas, year=year, logo_path=logo_path)
    _add_kata_pengantar(doc, title=title, author=author, nim=nim, year=year)
    blocks = _parse_markdown_to_blocks(content)
    headings = _extract_headings(blocks)
    _add_daftar_isi(doc, headings)
    _render_blocks_to_doc(doc, blocks)
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    doc.save(output_path)
    return output_path


# =============================================================================
# PDF
# =============================================================================

class MakalahPDF(FPDF):
    def __init__(self):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.set_margins(35, 30, 30)
        self.set_auto_page_break(auto=True, margin=30)
        self._page_number_style = "roman"
        self._content_start_page = 1

    def footer(self):
        self.set_y(-15)
        self.set_font("Times", "", 10)
        page = self.page_no()
        if page == 1:
            return  # no number on cover
        if page < self._content_start_page:
            roman = {1: "i", 2: "ii", 3: "iii", 4: "iv", 5: "v",
                     6: "vi", 7: "vii", 8: "viii", 9: "ix", 10: "x"}
            n = page - 1  # cover is page 1
            self.cell(0, 10, roman.get(n, str(n)), align="R")
        else:
            n = page - self._content_start_page + 1
            self.cell(0, 10, str(n), align="R")

    def _font(self, style="", size=12):
        self.set_font("Times", style, size)


# --- PDF Cover ---
def _pdf_cover(pdf, title, title_en="", lecturer="", author="",
               nim="", program_studi="", fakultas="",
               universitas="", year="", logo_path=""):
    pdf.add_page()
    pdf.ln(15)
    pdf._font("B", 14)
    pdf.cell(0, 8, "MAKALAH", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(10)
    pdf.multi_cell(0, 8, title.upper(), align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)
    if title_en:
        pdf._font("I", 14)
        pdf.multi_cell(0, 8, title_en, align="C", new_x="LMARGIN", new_y="NEXT")
    if lecturer:
        pdf.ln(3)
        pdf._font("", 12)
        pdf.multi_cell(0, 7, f"Dosen Pengampu : {lecturer}", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(8)
    if logo_path and os.path.exists(logo_path):
        pw = pdf.w - pdf.l_margin - pdf.r_margin
        iw = 40
        pdf.image(logo_path, x=pdf.l_margin + (pw - iw) / 2, w=iw)
        pdf.ln(8)
    pdf.ln(8)
    pdf._font("", 12)
    pdf.cell(0, 7, "Disusun Oleh :", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(8)
    if author and nim:
        pdf.cell(0, 7, f"{author} {nim}", align="C", new_x="LMARGIN", new_y="NEXT")
    elif author:
        pdf.cell(0, 7, author, align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(15)
    for txt in [program_studi, fakultas, universitas]:
        if txt:
            pdf._font("B", 12)
            pdf.cell(0, 7, txt.upper(), align="C", new_x="LMARGIN", new_y="NEXT")
    if year:
        pdf._font("", 12); pdf.ln(3)
        pdf.cell(0, 7, year, align="C", new_x="LMARGIN", new_y="NEXT")


# --- PDF Kata Pengantar ---
def _pdf_kata_pengantar(pdf, title, author="", nim="", year=""):
    pdf.add_page()
    pdf._font("B", 14)
    pdf.cell(0, 8, "KATA PENGANTAR", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(10)
    texts = [
        "Puji syukur kehadirat Tuhan Yang Maha Esa atas segala rahmat-Nya "
        "sehingga makalah ini dapat tersusun hingga selesai . Tidak lupa kami juga "
        "mengucapkan banyak terimakasih atas bantuan dari pihak yang telah "
        "berkontribusi dengan memberikan sumbangan baik materi maupun pikirannya.",
        "Dan harapan kami semoga makalah ini dapat menambah pengetahuan "
        "dan pengalaman bagi para pembaca, Untuk ke depannya dapat memperbaiki "
        "bentuk maupun menambah isi makalah agar menjadi lebih baik lagi.",
        "Karena keterbatasan pengetahuan maupun pengalaman kami, Kami "
        "yakin masih banyak kekurangan dalam makalah ini, Oleh karena itu kami "
        "sangat mengharapkan saran dan kritik yang membangun dari pembaca demi "
        "kesempurnaan makalah ini.",
    ]
    pdf._font("", 12)
    for t in texts:
        pdf.set_x(pdf.l_margin + 12.5)
        pdf.multi_cell(0, 7, t, align="J", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(3)
    pdf.ln(15)
    pdf.cell(0, 7, f"{_tanggal_indonesia()}", align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(10)
    pdf.cell(0, 7, "Penyusun", align="R", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(15)
    if author:
        pdf._font("U", 12)
        pdf.cell(0, 7, author, align="R", new_x="LMARGIN", new_y="NEXT")
    if nim:
        pdf._font("", 12)
        pdf.cell(0, 7, f"Nim. {nim}", align="R", new_x="LMARGIN", new_y="NEXT")


# --- PDF Daftar Isi ---
def _pdf_toc_entry(pdf, text, page_str, indent=0, link=None):
    pdf._font("", 12)
    x0 = pdf.l_margin + indent
    x_right = pdf.w - pdf.r_margin
    # Fixed-width cell for page number, right-aligned so all digits end at same x
    pg_cell_w = 10  # enough for 2-3 digit page numbers
    pg_x = x_right - pg_cell_w  # left edge of page number cell

    y0 = pdf.get_y()
    tw = pdf.get_string_width(text)

    # Available space for text + dots (before page number cell)
    avail = pg_x - x0 - 1  # -1mm gap before page number

    if tw > avail:
        # Text too long — show full text, put ". N" right after
        # No truncation, just minimal dots
        pdf.set_x(x0)
        pdf.multi_cell(avail, 7, text, new_x="RIGHT", new_y="LAST")
        # Dots (minimal)
        dw = pdf.get_string_width(".")
        remaining = pg_x - pdf.get_x()
        if remaining > dw:
            num_dots = int(remaining / dw)
            pdf.cell(remaining, 7, "." * num_dots)
    else:
        # Normal: text + dots filling the space
        dw = pdf.get_string_width(".")
        dots_space = avail - tw
        num_dots = max(0, int(dots_space / dw)) if dw > 0 else 0

        pdf.set_x(x0)
        pdf.cell(tw, 7, text)
        pdf.cell(dots_space, 7, "." * num_dots)

    # Page number: fixed-width cell, right-aligned — guarantees alignment
    pdf.set_x(pg_x)
    pdf.cell(pg_cell_w, 7, page_str, align="R", new_x="LMARGIN", new_y="NEXT")

    if link is not None:
        pdf.link(x0, y0, x_right - x0, 7, link)


def _pdf_toc_label(pdf, text):
    """BAB line in TOC — no dots, no page number, not bold."""
    pdf._font("", 12)
    pdf.cell(0, 7, text, new_x="LMARGIN", new_y="NEXT")


def _pdf_daftar_isi(pdf, headings, page_map=None, link_map=None):
    pdf.add_page()
    pdf._font("B", 14)
    pdf.cell(0, 8, "DAFTAR ISI", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(10)
    pdf._font("", 12)
    pdf.cell(0, 7, "Cover", new_x="LMARGIN", new_y="NEXT")
    _pdf_toc_entry(pdf, "KATA PENGANTAR", "ii")
    _pdf_toc_entry(pdf, "DAFTAR ISI", "iii")
    pdf.ln(2)

    est = 1
    for i, h in enumerate(headings):
        level = h["level"]; text = h["text"]
        pg = str(page_map.get(i, est)) if page_map else str(est)
        lnk = link_map.get(i) if link_map else None

        if level == 1:
            bab = _parse_bab_heading(text)
            is_dafpus = "daftar pustaka" in text.lower()
            if is_dafpus:
                _pdf_toc_entry(pdf, text.upper(), pg, link=lnk)
                est += 1
            elif bab:
                bab_num, bab_title = bab
                label = f"BAB {bab_num} {bab_title.upper()}" if bab_title else f"BAB {bab_num}"
                _pdf_toc_label(pdf, label)
                est += 2
            else:
                _pdf_toc_label(pdf, text.upper())
                est += 2
        elif level == 2:
            _pdf_toc_entry(pdf, text, pg, indent=10, link=lnk)
            est += 1
        else:
            _pdf_toc_entry(pdf, text, pg, indent=25, link=lnk)


# --- PDF Content ---
def _pdf_render(pdf, blocks, image_counter=1, link_map=None):
    heading_idx = 0
    for block in blocks:
        if block["type"] == "heading":
            level = block["level"]; text = block["text"]

            # Skip front matter headings
            if text.lower().strip() in _SKIP_HEADINGS:
                continue

            if level == 1:
                pdf.add_page()
                if link_map and heading_idx in link_map:
                    pdf.set_link(link_map[heading_idx], y=0, page=pdf.page_no())
                bab = _parse_bab_heading(text)
                if bab:
                    bab_num, bab_title = bab
                    pdf._font("B", 14)
                    pdf.cell(0, 8, f"BAB {bab_num}", align="C", new_x="LMARGIN", new_y="NEXT")
                    pdf.ln(3)
                    if bab_title:
                        pdf.multi_cell(0, 8, bab_title.upper(), align="C", new_x="LMARGIN", new_y="NEXT")
                    pdf.ln(10)
                else:
                    pdf._font("B", 14)
                    pdf.multi_cell(0, 8, text.upper(), align="C", new_x="LMARGIN", new_y="NEXT")
                    pdf.ln(10)
            else:
                if link_map and heading_idx in link_map:
                    pdf.set_link(link_map[heading_idx], y=pdf.get_y(), page=pdf.page_no())
                pdf._font("B", 12)
                pdf.ln(5)
                pdf.multi_cell(0, 7, text, new_x="LMARGIN", new_y="NEXT")
                pdf.ln(3)
            heading_idx += 1

        elif block["type"] == "paragraph":
            pdf._font("", 12)
            pdf.set_x(pdf.l_margin + 12.5)
            pdf.multi_cell(0, 7, block["text"], align="J", new_x="LMARGIN", new_y="NEXT")
            pdf.ln(3)

        elif block["type"] == "image":
            path = block["path"]
            if os.path.exists(path):
                pw = pdf.w - pdf.l_margin - pdf.r_margin
                iw = min(120, pw)
                pdf.image(path, x=pdf.l_margin + (pw - iw) / 2, w=iw)
                pdf.ln(3)
                cap = block["caption"] or ""
                if re.match(r"^Gambar\s+\d+", cap, re.IGNORECASE):
                    cap_text = cap
                else:
                    cap_text = f"Gambar {image_counter}: {cap}" if cap else f"Gambar {image_counter}"
                pdf._font("I", 10)
                pdf.cell(0, 6, cap_text, align="C", new_x="LMARGIN", new_y="NEXT")
                pdf.ln(3); image_counter += 1

        elif block["type"] in ("list_item", "numbered_item"):
            pdf._font("", 12)
            prefix = f"{block.get('number', '-')}. " if block["type"] == "numbered_item" else "- "
            pdf.set_x(pdf.l_margin + 10)
            pdf.multi_cell(0, 7, f"{prefix}{block['text']}", align="J", new_x="LMARGIN", new_y="NEXT")

        elif block["type"] == "reference":
            # Daftar pustaka entry: hanging indent
            pdf._font("", 12)
            text = block["text"]
            # First line at left margin, continuation lines indented
            indent = 10  # mm for hanging indent
            pdf.set_x(pdf.l_margin)
            # Use multi_cell with left margin adjustment for hanging effect
            # Save x, print first part, then set indent for wrap
            pdf.set_left_margin(pdf.l_margin + indent)
            pdf.set_x(pdf.l_margin - indent)  # first line starts at original margin
            pdf.multi_cell(0, 7, text, align="J", new_x="LMARGIN", new_y="NEXT")
            pdf.set_left_margin(35)  # reset to original left margin
            pdf.ln(1)

    return image_counter


def save_as_pdf(content, title, output_path, title_en="", lecturer="",
                 author="", nim="", program_studi="", fakultas="",
                 universitas="", year="", logo_path=""):
    if not year:
        year = str(datetime.now().year)

    blocks = _parse_markdown_to_blocks(content)
    headings = _extract_headings(blocks)
    kw = dict(title=title, title_en=title_en, lecturer=lecturer, author=author,
              nim=nim, program_studi=program_studi, fakultas=fakultas,
              universitas=universitas, year=year, logo_path=logo_path)

    # Track page numbers for TOC
    page_map = {}
    pt = MakalahPDF()
    _pdf_cover(pt, **kw)
    _pdf_kata_pengantar(pt, title=title, author=author, nim=nim, year=year)
    _pdf_daftar_isi(pt, headings)
    pt._content_start_page = pt.page_no() + 1
    hidx = 0
    for block in blocks:
        if block["type"] == "heading":
            if block["text"].lower().strip() in _SKIP_HEADINGS:
                continue
            if block["level"] == 1:
                pt.add_page()
            if hidx < len(headings):
                page_map[hidx] = pt.page_no() - pt._content_start_page + 1
                hidx += 1
            if block["level"] == 1:
                pt._font("B", 14)
                pt.cell(0, 8, "X", align="C", new_x="LMARGIN", new_y="NEXT")
            else:
                pt._font("B", 12)
                pt.multi_cell(0, 7, block["text"], new_x="LMARGIN", new_y="NEXT")
        elif block["type"] == "paragraph":
            pt._font("", 12)
            pt.set_x(pt.l_margin + 12.5)
            pt.multi_cell(0, 7, block["text"], align="J", new_x="LMARGIN", new_y="NEXT")
            pt.ln(3)
        elif block["type"] in ("list_item", "numbered_item"):
            pt._font("", 12)
            pt.set_x(pt.l_margin + 10)
            pt.multi_cell(0, 7, f"- {block['text']}", new_x="LMARGIN", new_y="NEXT")
        elif block["type"] == "reference":
            pt._font("", 12)
            pt.multi_cell(0, 7, block["text"], new_x="LMARGIN", new_y="NEXT")

    # Pass 2: final render with links
    pdf = MakalahPDF()
    link_map = {}
    for i in range(len(headings)):
        lnk = pdf.add_link()
        pdf.set_link(lnk, page=1)  # temp
        link_map[i] = lnk

    _pdf_cover(pdf, **kw)
    _pdf_kata_pengantar(pdf, title=title, author=author, nim=nim, year=year)
    _pdf_daftar_isi(pdf, headings, page_map=page_map, link_map=link_map)
    pdf._content_start_page = pdf.page_no() + 1
    pdf._page_number_style = "arabic"
    _pdf_render(pdf, blocks, link_map=link_map)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    pdf.output(output_path)
    return output_path
