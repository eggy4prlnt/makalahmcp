import json
import os
from datetime import datetime

from mcp.server.fastmcp import FastMCP

from scraper import research_topic as do_research
from converter import save_as_docx, save_as_pdf

mcp = FastMCP(
    "MakalahMCP",
    instructions="MCP server for generating Indonesian academic papers (makalah)",
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
