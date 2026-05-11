# MakalahMCP Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build an MCP server with 2 tools — `research_topic` (Google scraping) and `save_makalah` (Markdown → DOCX/PDF) — for generating Indonesian academic papers.

**Architecture:** MCP server (FastMCP, stdio transport) exposes two tools. `scraper.py` handles Google Search scraping, page content extraction, and image downloading. `converter.py` handles Markdown parsing and conversion to styled DOCX (python-docx) and PDF (fpdf2) following Indonesian academic formatting guidelines.

**Tech Stack:** Python 3.13+, mcp[cli], httpx, beautifulsoup4, python-docx, fpdf2

---

## File Structure

| File | Responsibility |
|------|---------------|
| `main.py` | MCP server entry point, tool definitions (`research_topic`, `save_makalah`) |
| `scraper.py` | Google Search scraping, page content extraction, image downloading |
| `converter.py` | Markdown → DOCX and PDF conversion with academic styling |
| `pyproject.toml` | Project config and dependencies |

---

### Task 1: Add Dependencies

**Files:**
- Modify: `pyproject.toml`

- [ ] **Step 1: Add all required dependencies to pyproject.toml**

```toml
[project]
name = "makalahmcp"
version = "0.1.0"
description = "MCP server for generating Indonesian academic papers (makalah)"
readme = "README.md"
requires-python = ">=3.13"
dependencies = [
    "httpx>=0.28.1",
    "mcp[cli]>=1.27.1",
    "beautifulsoup4>=4.13.4",
    "python-docx>=1.1.2",
    "fpdf2>=2.8.3",
]
```

- [ ] **Step 2: Install dependencies**

Run: `cd /Users/aslan/Project/makalahmcp && uv sync`
Expected: All packages installed successfully, `uv.lock` updated.

- [ ] **Step 3: Commit**

```bash
git add pyproject.toml uv.lock
git commit -m "feat: add project dependencies for MCP makalah server"
```

---

### Task 2: Build Scraper Module — Google Search

**Files:**
- Create: `scraper.py`

This task builds the Google Search scraping function. It searches Google, extracts result links, fetches page content and images.

- [ ] **Step 1: Create scraper.py with google_search function**

```python
import httpx
import os
import re
import tempfile
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "id-ID,id;q=0.9,en-US;q=0.8,en;q=0.7",
}

IMAGE_DIR = os.path.join(tempfile.gettempdir(), "makalahmcp", "images")


async def google_search(query: str, num_results: int = 5) -> list[str]:
    """Search Google and return a list of result URLs."""
    url = "https://www.google.com/search"
    params = {"q": query, "num": num_results + 5, "hl": "id"}
    async with httpx.AsyncClient(headers=HEADERS, follow_redirects=True, timeout=15) as client:
        resp = await client.get(url, params=params)
        resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    links: list[str] = []
    for a_tag in soup.select("a[href]"):
        href = a_tag["href"]
        if isinstance(href, list):
            href = href[0]
        if href.startswith("/url?q="):
            real_url = href.split("/url?q=")[1].split("&")[0]
            if not any(x in real_url for x in ["google.com", "youtube.com", "webcache"]):
                links.append(real_url)
        if len(links) >= num_results:
            break
    return links
```

- [ ] **Step 2: Add fetch_page_content function**

Append to `scraper.py`:

```python
MIN_IMAGE_SIZE = 10000  # 10KB minimum to skip icons/logos


async def _download_image(client: httpx.AsyncClient, img_url: str, idx: int) -> dict | None:
    """Download a single image and return its metadata, or None on failure."""
    try:
        resp = await client.get(img_url, timeout=10)
        resp.raise_for_status()
        content_type = resp.headers.get("content-type", "")
        if "image" not in content_type:
            return None
        if len(resp.content) < MIN_IMAGE_SIZE:
            return None
        ext = "png"
        if "jpeg" in content_type or "jpg" in content_type:
            ext = "jpg"
        elif "gif" in content_type:
            ext = "gif"
        os.makedirs(IMAGE_DIR, exist_ok=True)
        filename = f"img_{idx:03d}.{ext}"
        filepath = os.path.join(IMAGE_DIR, filename)
        with open(filepath, "wb") as f:
            f.write(resp.content)
        return {"url": img_url, "caption": "", "local_path": filepath}
    except Exception:
        return None


async def fetch_page_content(url: str, image_start_idx: int = 0) -> dict:
    """Fetch a web page and extract its text content, metadata, and images."""
    result = {
        "title": "",
        "url": url,
        "author": "",
        "year": "",
        "content": "",
        "images": [],
    }
    try:
        async with httpx.AsyncClient(headers=HEADERS, follow_redirects=True, timeout=15) as client:
            resp = await client.get(url)
            resp.raise_for_status()
            soup = BeautifulSoup(resp.text, "html.parser")

            # Title
            title_tag = soup.find("title")
            result["title"] = title_tag.get_text(strip=True) if title_tag else ""

            # Author from meta tags
            author_meta = soup.find("meta", attrs={"name": "author"})
            if author_meta and author_meta.get("content"):
                result["author"] = author_meta["content"]

            # Year from meta date or page content
            date_meta = soup.find("meta", attrs={"property": "article:published_time"})
            if date_meta and date_meta.get("content"):
                result["year"] = date_meta["content"][:4]
            else:
                time_tag = soup.find("time")
                if time_tag:
                    dt = time_tag.get("datetime", time_tag.get_text())
                    year_match = re.search(r"(20\d{2}|19\d{2})", str(dt))
                    if year_match:
                        result["year"] = year_match.group(1)

            # Remove script, style, nav, footer, header elements
            for tag in soup.find_all(["script", "style", "nav", "footer", "header", "aside"]):
                tag.decompose()

            # Extract main content
            main = soup.find("article") or soup.find("main") or soup.find("body")
            if main:
                paragraphs = main.find_all("p")
                text_parts = [p.get_text(strip=True) for p in paragraphs if len(p.get_text(strip=True)) > 30]
                result["content"] = "\n\n".join(text_parts)

            # Extract images
            if main:
                img_tags = main.find_all("img", src=True)
                img_idx = image_start_idx
                for img in img_tags:
                    src = img.get("src", "")
                    if not src or src.startswith("data:"):
                        continue
                    img_url = urljoin(url, src)
                    downloaded = await _download_image(client, img_url, img_idx)
                    if downloaded:
                        alt = img.get("alt", "")
                        downloaded["caption"] = alt
                        result["images"].append(downloaded)
                        img_idx += 1
                    if len(result["images"]) >= 5:
                        break
    except Exception:
        pass
    return result
```

- [ ] **Step 3: Add research_topic orchestrator function**

Append to `scraper.py`:

```python
async def research_topic(title: str, num_results: int = 5) -> dict:
    """Research a topic by searching Google and fetching content from results."""
    urls = await google_search(title, num_results)
    if not urls:
        return {"references": [], "error": "Tidak dapat menemukan hasil pencarian. Coba judul yang berbeda."}

    references = []
    image_idx = 0
    for url in urls:
        page = await fetch_page_content(url, image_start_idx=image_idx)
        if page["content"]:
            references.append(page)
            image_idx += len(page["images"])

    return {"references": references}
```

- [ ] **Step 4: Verify module imports work**

Run: `cd /Users/aslan/Project/makalahmcp && uv run python -c "from scraper import research_topic; print('OK')"`
Expected: `OK`

- [ ] **Step 5: Commit**

```bash
git add scraper.py
git commit -m "feat: add scraper module for Google Search and content extraction"
```

---

### Task 3: Build Converter Module — DOCX Generation

**Files:**
- Create: `converter.py`

This task builds the DOCX converter with proper Indonesian academic formatting (Times New Roman, margins, cover page, numbered headings).

- [ ] **Step 1: Create converter.py with document setup and cover page**

```python
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
```

- [ ] **Step 2: Add markdown parsing and content rendering functions**

Append to `converter.py`:

```python
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
```

- [ ] **Step 3: Add the main save_as_docx function**

Append to `converter.py`:

```python
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
```

- [ ] **Step 4: Verify DOCX generation works**

Run: `cd /Users/aslan/Project/makalahmcp && uv run python -c "
from converter import save_as_docx
content = '# I. Pendahuluan\n\nIni adalah contoh paragraf.\n\n## 1.1 Latar Belakang\n\nLorem ipsum dolor sit amet.'
path = save_as_docx(content, title='Contoh Makalah', output_path='/tmp/test_makalah.docx', author='Test User', universitas='Universitas Test', year='2026')
print(f'Created: {path}')
"`
Expected: `Created: /tmp/test_makalah.docx`

- [ ] **Step 5: Commit**

```bash
git add converter.py
git commit -m "feat: add DOCX converter with academic styling and cover page"
```

---

### Task 4: Build Converter Module — PDF Generation

**Files:**
- Modify: `converter.py`

Add PDF generation capability using fpdf2.

- [ ] **Step 1: Add PDF generation function to converter.py**

Append to `converter.py`:

```python
from fpdf import FPDF


class MakalahPDF(FPDF):
    """Custom PDF class for makalah formatting."""

    def __init__(self):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.set_margins(35, 30, 30)
        self.set_auto_page_break(auto=True, margin=30)

    def _set_font_safe(self, family: str = "Times", style: str = "", size: int = 12):
        """Set font with fallback."""
        try:
            self.set_font(family, style, size)
        except Exception:
            self.set_font("Times", style, size)


def _build_cover_pdf(
    pdf: MakalahPDF,
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
    """Add cover page to PDF."""
    pdf.add_page()
    pdf.ln(20)

    # Title - uppercase, bold, 14pt
    pdf._set_font_safe("Times", "B", 14)
    pdf.multi_cell(0, 8, title.upper(), align="C")
    pdf.ln(3)

    # English title - italic
    if title_en:
        pdf._set_font_safe("Times", "I", 14)
        pdf.multi_cell(0, 8, title_en, align="C")
    pdf.ln(3)

    # Lecturer
    if lecturer:
        pdf._set_font_safe("Times", "", 12)
        pdf.multi_cell(0, 7, f"Dosen Pengampu : {lecturer}", align="C")
    pdf.ln(10)

    # Logo
    if logo_path and os.path.exists(logo_path):
        page_w = pdf.w - pdf.l_margin - pdf.r_margin
        img_w = 40
        x = pdf.l_margin + (page_w - img_w) / 2
        pdf.image(logo_path, x=x, w=img_w)
        pdf.ln(10)

    pdf.ln(10)

    # "Disusun Oleh:"
    pdf._set_font_safe("Times", "", 12)
    pdf.cell(0, 7, "Disusun Oleh:", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    # Author
    if author:
        pdf._set_font_safe("Times", "B", 12)
        pdf.cell(0, 7, author.upper(), align="C", new_x="LMARGIN", new_y="NEXT")

    # NIM
    if nim:
        pdf._set_font_safe("Times", "", 12)
        pdf.cell(0, 7, f"NIM: {nim}", align="C", new_x="LMARGIN", new_y="NEXT")

    pdf.ln(20)

    # Program Studi
    if program_studi:
        pdf._set_font_safe("Times", "B", 12)
        pdf.cell(0, 7, program_studi.upper(), align="C", new_x="LMARGIN", new_y="NEXT")

    # Fakultas
    if fakultas:
        pdf._set_font_safe("Times", "B", 12)
        pdf.cell(0, 7, fakultas.upper(), align="C", new_x="LMARGIN", new_y="NEXT")

    # Universitas
    if universitas:
        pdf._set_font_safe("Times", "B", 12)
        pdf.cell(0, 7, universitas.upper(), align="C", new_x="LMARGIN", new_y="NEXT")

    # Year
    if year:
        pdf._set_font_safe("Times", "", 12)
        pdf.ln(5)
        pdf.cell(0, 7, year, align="C", new_x="LMARGIN", new_y="NEXT")


def _render_blocks_to_pdf(pdf: MakalahPDF, blocks: list[dict], image_counter: int = 1) -> int:
    """Render parsed markdown blocks into PDF. Returns updated image counter."""
    for block in blocks:
        if block["type"] == "heading":
            level = block["level"]
            size = 14 if level == 1 else 12
            pdf._set_font_safe("Times", "B", size)
            pdf.ln(5)
            pdf.multi_cell(0, 8, block["text"])
            pdf.ln(3)

        elif block["type"] == "paragraph":
            pdf._set_font_safe("Times", "", 12)
            pdf.multi_cell(0, 7, block["text"])
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
            x = pdf.get_x()
            pdf.cell(5, 7, chr(8226))  # bullet
            pdf.multi_cell(0, 7, block["text"])

    return image_counter


def save_as_pdf(
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
    """Convert markdown content to a styled PDF file. Returns the file path."""
    pdf = MakalahPDF()

    _build_cover_pdf(
        pdf,
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

    # Content pages
    pdf.add_page()
    blocks = _parse_markdown_to_blocks(content)
    _render_blocks_to_pdf(pdf, blocks)

    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    pdf.output(output_path)
    return output_path
```

- [ ] **Step 2: Verify PDF generation works**

Run: `cd /Users/aslan/Project/makalahmcp && uv run python -c "
from converter import save_as_pdf
content = '# I. Pendahuluan\n\nIni adalah contoh paragraf.\n\n## 1.1 Latar Belakang\n\nLorem ipsum dolor sit amet.'
path = save_as_pdf(content, title='Contoh Makalah', output_path='/tmp/test_makalah.pdf', author='Test User', universitas='Universitas Test', year='2026')
print(f'Created: {path}')
"`
Expected: `Created: /tmp/test_makalah.pdf`

- [ ] **Step 3: Commit**

```bash
git add converter.py
git commit -m "feat: add PDF generation with academic formatting"
```

---

### Task 5: Build MCP Server — Tool Definitions

**Files:**
- Modify: `main.py`

This is the main entry point. Define the MCP server with `research_topic` and `save_makalah` tools.

- [ ] **Step 1: Replace main.py with MCP server and research_topic tool**

```python
import json
import os
from datetime import datetime

from mcp.server.fastmcp import FastMCP

from scraper import research_topic as do_research
from converter import save_as_docx, save_as_pdf

mcp = FastMCP(
    "MakalahMCP",
    description="MCP server for generating Indonesian academic papers (makalah)",
)


@mcp.tool()
async def research_topic(title: str, num_results: int = 5) -> str:
    """Research a topic for an academic paper (makalah).

    Searches Google and fetches content from top results to gather references.

    Args:
        title: The topic/title of the makalah to research.
        num_results: Number of search results to fetch (default: 5).

    Returns:
        JSON string containing references with title, url, author, year, content, and images.
    """
    result = await do_research(title, num_results)
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
async def save_makalah(
    content: str,
    title: str,
    format: str = "both",
    title_en: str = "",
    lecturer: str = "",
    author: str = "",
    nim: str = "",
    program_studi: str = "",
    fakultas: str = "",
    universitas: str = "",
    year: str = "",
    logo_path: str = "",
    output_dir: str = "",
) -> str:
    """Save a makalah (academic paper) as DOCX and/or PDF file.

    Takes markdown content and converts it to properly formatted academic documents
    following Indonesian academic paper guidelines (Times New Roman, proper margins,
    cover page, numbered headings).

    Args:
        content: The makalah content in markdown format.
        title: The makalah title (Indonesian).
        format: Output format - "docx", "pdf", or "both" (default: "both").
        title_en: The makalah title in English (optional).
        lecturer: Name of the supervising lecturer (optional).
        author: Author's name (optional).
        nim: Student ID number / NIM (optional).
        program_studi: Study program name (optional).
        fakultas: Faculty name (optional).
        universitas: University name (optional).
        year: Year of publication (default: current year).
        logo_path: Path to institution logo image file (optional).
        output_dir: Output directory (default: ~/Documents).

    Returns:
        JSON string with list of created file paths.
    """
    if not output_dir:
        output_dir = os.path.expanduser("~/Documents")
    os.makedirs(output_dir, exist_ok=True)

    safe_title = "".join(c if c.isalnum() or c in " -_" else "" for c in title)[:80].strip()
    base_name = f"Makalah - {safe_title}"
    if not year:
        year = str(datetime.now().year)

    files = []
    kwargs = dict(
        content=content,
        title=title,
        title_en=title_en,
        lecturer=lecturer,
        author=author,
        nim=nim,
        program_studi=program_studi,
        fakultas=fakultas,
        universitas=universitas,
        year=year,
        logo_path=logo_path,
    )

    if format in ("docx", "both"):
        docx_path = os.path.join(output_dir, f"{base_name}.docx")
        save_as_docx(output_path=docx_path, **kwargs)
        files.append(docx_path)

    if format in ("pdf", "both"):
        pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
        save_as_pdf(output_path=pdf_path, **kwargs)
        files.append(pdf_path)

    return json.dumps({"files": files}, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    mcp.run(transport="stdio")
```

- [ ] **Step 2: Verify MCP server starts without errors**

Run: `cd /Users/aslan/Project/makalahmcp && uv run python -c "from main import mcp; print('Server OK, tools:', [t.name for t in mcp._tool_manager.list_tools()])"`
Expected: Output shows `Server OK, tools:` with `research_topic` and `save_makalah` listed.

- [ ] **Step 3: Commit**

```bash
git add main.py
git commit -m "feat: add MCP server with research_topic and save_makalah tools"
```

---

### Task 6: Integration Test

**Files:**
- None (testing only)

End-to-end verification that both tools work together.

- [ ] **Step 1: Test research_topic tool**

Run: `cd /Users/aslan/Project/makalahmcp && uv run python -c "
import asyncio, json
from scraper import research_topic
result = asyncio.run(research_topic('dampak media sosial terhadap remaja', 3))
print(f'Found {len(result.get(\"references\", []))} references')
for ref in result.get('references', []):
    print(f'  - {ref[\"title\"][:60]}... ({len(ref[\"content\"])} chars, {len(ref[\"images\"])} images)')
"`
Expected: Shows 1-3 references with content and possibly images.

- [ ] **Step 2: Test save_makalah tool via MCP**

Run: `cd /Users/aslan/Project/makalahmcp && uv run python -c "
import asyncio, json
from main import save_makalah

content = '''# I. Abstrak

Makalah ini membahas dampak media sosial terhadap remaja di era digital.

# II. Pendahuluan

## 1.1 Latar Belakang

Media sosial telah menjadi bagian tak terpisahkan dari kehidupan remaja modern.

## 1.2 Rumusan Masalah

Bagaimana dampak media sosial terhadap perilaku dan kesehatan mental remaja?

## 1.3 Tujuan

Menganalisis pengaruh media sosial terhadap perkembangan remaja.

# III. Pembahasan

Media sosial memberikan dampak positif dan negatif bagi remaja (Anderson 2021).

# IV. Simpulan

Media sosial memiliki peran signifikan dalam kehidupan remaja.

# Daftar Pustaka

Anderson, M. 2021. Teens, Social Media and Technology. Washington: Pew Research Center.
'''

result = asyncio.run(save_makalah(
    content=content,
    title='Dampak Media Sosial Terhadap Remaja',
    title_en='The Impact of Social Media on Teenagers',
    author='Puan Maharanti',
    nim='C1C020082',
    program_studi='Program Studi Akuntansi',
    fakultas='Fakultas Ekonomi dan Bisnis',
    universitas='Universitas Jambi',
    year='2026',
    output_dir='/tmp/makalahmcp_test',
))
data = json.loads(result)
for f in data['files']:
    size = __import__('os').path.getsize(f)
    print(f'{f} ({size} bytes)')
"`
Expected: Shows 2 files created (DOCX and PDF) with non-zero sizes.

- [ ] **Step 3: Test MCP server with mcp dev (manual)**

Run: `cd /Users/aslan/Project/makalahmcp && uv run mcp dev main.py`

This opens the MCP inspector in the browser. Manually verify:
1. Both tools appear in the tool list
2. `research_topic` works with a test title
3. `save_makalah` works with test content

- [ ] **Step 4: Final commit with any fixes**

Only if fixes were needed during integration testing:

```bash
git add -A
git commit -m "fix: integration test fixes"
```
