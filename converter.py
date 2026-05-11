import os
import re
from datetime import datetime
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT


def _setup_document() -> Document:
    """Create a new Document with proper page setup: A4, margins, Times New Roman."""
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


def _add_cover_page(
    doc: Document,
    title: str,
    title_en: str = "",
    lecturer: str = "",
    author: str = "",
    nim: str = "",
    program_studi: str = "",
    fakultas: str = "",
    universitas: str = "",
    year: str = "",
    logo_path: str = "",
) -> None:
    """Add a formatted cover page to the document."""
    # Title (Indonesian) - uppercase, bold, 14pt
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title.upper())
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(14)

    # Title (English) - italic, 14pt
    if title_en:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(title_en)
        run.italic = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(14)

    # Lecturer
    if lecturer:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"Dosen Pengampu : {lecturer}")
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)

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

    # "Disusun Oleh:"
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Disusun Oleh:")
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)

    # Author name - uppercase, bold
    if author:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(author.upper())
        run.bold = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)

    # NIM
    if nim:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"NIM: {nim}")
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)

    # Spacing
    doc.add_paragraph()
    doc.add_paragraph()

    # Program Studi - uppercase, bold
    if program_studi:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(program_studi.upper())
        run.bold = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)

    # Fakultas - uppercase, bold
    if fakultas:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(fakultas.upper())
        run.bold = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)

    # Universitas - uppercase, bold
    if universitas:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(universitas.upper())
        run.bold = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)

    # Year
    if year:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(year)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)

    # Page break after cover
    doc.add_page_break()


def _parse_markdown_to_blocks(content: str) -> list[dict]:
    """Parse markdown content into structured blocks.

    Returns a list of dicts with keys:
    - type: "heading", "paragraph", "image", "list_item"
    - level: heading level (1-3) or 0
    - text: the content text
    - path: image path (for image type)
    - caption: image caption (for image type)
    """
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
        num_list_match = re.match(r"^\d+\.\s+(.*)", line)
        if num_list_match:
            blocks.append({"type": "list_item", "level": 0, "text": num_list_match.group(1)})
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


def _add_styled_heading(doc: Document, text: str, level: int) -> None:
    """Add a heading with Times New Roman font."""
    heading = doc.add_heading(text, level=level)
    for run in heading.runs:
        run.font.name = "Times New Roman"
        run.font.size = Pt(14) if level == 1 else Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)


def _render_blocks_to_doc(doc: Document, blocks: list[dict], image_counter: int = 1) -> int:
    """Render parsed blocks into the document. Returns updated image counter."""
    for block in blocks:
        if block["type"] == "heading":
            _add_styled_heading(doc, block["text"], block["level"])

        elif block["type"] == "paragraph":
            p = doc.add_paragraph(block["text"])
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

                # Caption below image
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
            for run in p.runs:
                run.font.name = "Times New Roman"
                run.font.size = Pt(12)

    return image_counter


def save_as_docx(
    content: str,
    title: str,
    output_path: str,
    title_en: str = "",
    lecturer: str = "",
    author: str = "",
    nim: str = "",
    program_studi: str = "",
    fakultas: str = "",
    universitas: str = "",
    year: str = "",
    logo_path: str = "",
) -> str:
    """Convert markdown content to a styled DOCX file. Returns the file path."""
    doc = _setup_document()

    _add_cover_page(
        doc,
        title=title,
        title_en=title_en,
        lecturer=lecturer,
        author=author,
        nim=nim,
        program_studi=program_studi,
        fakultas=fakultas,
        universitas=universitas,
        year=year or str(datetime.now().year),
        logo_path=logo_path,
    )

    blocks = _parse_markdown_to_blocks(content)
    _render_blocks_to_doc(doc, blocks)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    doc.save(output_path)
    return output_path
