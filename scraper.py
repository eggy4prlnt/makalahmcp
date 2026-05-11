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


async def web_search(query: str, num_results: int = 5) -> list[str]:
    """Search the web using DuckDuckGo HTML and return a list of result URLs."""
    url = "https://html.duckduckgo.com/html/"
    data = {"q": query}
    async with httpx.AsyncClient(headers=HEADERS, follow_redirects=True, timeout=15) as client:
        resp = await client.post(url, data=data)
        resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    links: list[str] = []
    for a_tag in soup.select("a.result__a[href]"):
        href = a_tag.get("href", "")
        if isinstance(href, list):
            href = href[0]
        # DuckDuckGo wraps URLs in a redirect, extract the actual URL
        if "uddg=" in href:
            from urllib.parse import unquote, parse_qs, urlparse as _urlparse
            parsed = _urlparse(href)
            qs = parse_qs(parsed.query)
            if "uddg" in qs:
                href = unquote(qs["uddg"][0])
        if href.startswith("http") and "duckduckgo.com" not in href:
            links.append(href)
        if len(links) >= num_results:
            break
    return links


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


async def research_topic(title: str, num_results: int = 5) -> dict:
    """Research a topic by searching the web and fetching content from results."""
    urls = await web_search(title, num_results)
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
