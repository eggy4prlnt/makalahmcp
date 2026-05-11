import json
import os
from datetime import datetime

from mcp.server.fastmcp import FastMCP

from scraper import research_topic as do_research, search_university_logo as do_search_logo
from converter import save_as_docx, save_as_pdf

mcp = FastMCP(
    "MakalahMCP",
    instructions="""MCP server for generating Indonesian academic papers (makalah).

FLOW WAJIB saat user minta buat makalah:
1. Tanyakan JUDUL makalah (jika belum ada)
2. Tanyakan data penulis (satu per satu atau sekaligus):
   - Nama lengkap (wajib)
   - NIM (wajib)
   - Universitas (wajib)
   - Program Studi (wajib)
   - Fakultas (wajib)
   - Dosen Pengampu (opsional, boleh dikosongkan)
3. Cari logo universitas dengan tool search_logo
4. Tahun OTOMATIS pakai tahun sekarang, JANGAN tanya ke user
5. Lakukan research_topic dengan judul
6. Generate konten makalah menggunakan prompt generate_makalah
7. Simpan dengan save_makalah (format "both" untuk DOCX + PDF)
""",
)


@mcp.tool()
async def research_topic(title: str, num_results: int = 5) -> str:
    """Research a topic for an academic paper (makalah).

    Searches the web and fetches content from top results to gather references.
    Each reference includes a URL that MUST be included in the daftar pustaka (bibliography).

    After calling this tool, use the 'generate_makalah' prompt to get detailed
    instructions for writing comprehensive, long-form academic content from these references.

    Args:
        title: The topic/title of the makalah to research.
        num_results: Number of search results to fetch (default: 5).

    Returns:
        JSON string containing references with title, url, author, year, content, and images.
        The URL field is important for bibliography entries.
    """
    result = await do_research(title, num_results)
    return json.dumps(result, ensure_ascii=False, indent=2)


@mcp.tool()
async def search_logo(university_name: str) -> str:
    """Search and download a university logo image for the makalah cover page.

    Searches the web for the university logo, downloads it, and returns the local file path.
    Call this after getting the university name from the user.

    Args:
        university_name: The full name of the university (e.g. "Universitas Jambi").

    Returns:
        JSON string with found status and local file path of the downloaded logo.
    """
    result = await do_search_logo(university_name)
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
        year: Year of publication (default: current year, DO NOT ask user).
        logo_path: Path to institution logo image file from search_logo tool.
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


@mcp.prompt()
def generate_makalah(title: str, references_json: str) -> str:
    """Generate a complete makalah (academic paper) from research references.

    Use this prompt after calling research_topic to get the proper format
    and instructions for generating comprehensive makalah content.

    Args:
        title: The makalah title.
        references_json: JSON string from research_topic tool output.
    """
    return f"""Kamu adalah penulis akademik. Buat makalah lengkap dan komprehensif dengan judul:
"{title}"

Gunakan referensi berikut sebagai sumber:
{references_json}

## ATURAN PENTING

### Panjang Konten — INI SANGAT PENTING
- Makalah harus SANGAT PANJANG dan SANGAT DETAIL. Target minimal 5000-8000 kata.
- Setiap sub-bab harus memiliki minimal 4-6 paragraf PANJANG
- Setiap paragraf harus 5-8 kalimat, bukan 1-2 kalimat
- JANGAN pernah menulis paragraf pendek. Setiap paragraf harus penuh penjelasan.
- Elaborasi setiap poin dengan: definisi, penjelasan mendalam, contoh konkret, data/angka, dampak/implikasi, dan hubungan dengan teori
- Gunakan banyak referensi dan kutipan dalam setiap bab
- Tulis seolah-olah ini adalah makalah untuk tugas akhir kuliah yang sangat penting

### Cara Menulis Pembahasan yang Detail
Contoh BURUK (terlalu singkat):
"Dashboard interaktif terbukti efektif dalam menyajikan data operasional."

Contoh BAIK (detail dan mendalam):
"Penerapan dashboard interaktif dalam konteks operasional teknologi informasi telah menunjukkan efektivitas yang sangat signifikan dalam menyajikan data operasional secara komprehensif. Berdasarkan hasil analisis yang telah dilakukan, dashboard yang dibangun menggunakan Power BI mampu mengintegrasikan berbagai sumber data seperti log server, tiket helpdesk, dan laporan performa jaringan ke dalam satu tampilan visual yang koheren dan mudah dipahami oleh berbagai tingkatan pemangku kepentingan. Kemampuan drill-down yang disediakan oleh dashboard memungkinkan pengguna untuk menelusuri data dari tingkat agregat hingga tingkat detail yang paling granular, sehingga analisis root cause terhadap permasalahan operasional dapat dilakukan dengan lebih cepat dan akurat. Hal ini sejalan dengan temuan penelitian sebelumnya yang menunjukkan bahwa visualisasi data interaktif mampu meningkatkan kecepatan pengambilan keputusan hingga 5 kali lipat dibandingkan metode pelaporan konvensional berbasis spreadsheet (Silvianingsih 2024)."

Tulis SEMUA paragraf dengan level detail seperti contoh BAIK di atas.

### Struktur (gunakan heading markdown #)
PENTING: Kata Pengantar dan Daftar Isi TIDAK perlu ditulis dalam konten markdown.
Keduanya akan di-generate OTOMATIS oleh sistem saat konversi ke DOCX/PDF.
Mulai langsung dari Abstrak.

# I. Abstrak
(Dalam Bahasa Indonesia DAN Bahasa Inggris, masing-masing 200-300 kata)

# II. Pendahuluan
## 1.1 Latar Belakang (minimal 5 paragraf panjang, jelaskan konteks luas sampai spesifik)
## 1.2 Rumusan Masalah (jelaskan dulu konteksnya, baru list pertanyaan)
## 1.3 Tujuan Penulisan (jelaskan tujuan umum dulu, baru list tujuan khusus)
## 1.4 Manfaat Penulisan (manfaat teoritis DAN praktis, masing-masing 1 paragraf)
## 1.5 Tinjauan Pustaka (minimal 4 paragraf, bahas setiap referensi secara mendalam, bandingkan temuan antar referensi)

# III. Metode Penelitian
(Minimal 4 paragraf: pendekatan, tahapan, alat/tools, serta justifikasi pemilihan metode)

# IV. Hasil
(Minimal 5 paragraf: sajikan setiap temuan dengan data spesifik, angka, persentase, tren. Pisahkan per domain/kategori temuan. Jelaskan juga visualisasi yang dihasilkan.)

# V. Pembahasan
(Minimal 6 paragraf: analisis mendalam setiap temuan, hubungkan dengan teori dan referensi, bandingkan dengan penelitian terdahulu, jelaskan implikasi praktis, identifikasi keterbatasan, berikan rekomendasi. INI BAGIAN TERPENTING — harus paling panjang dan paling detail.)

# VI. Simpulan
(Minimal 3 paragraf: ringkasan temuan utama, kontribusi penelitian, saran untuk penelitian selanjutnya dengan detail)

# Daftar Pustaka

### Format Kutipan Dalam Teks
- Satu penulis: (Thompson 1990)
- Dua penulis: (Becker and Seligman 1996)
- Tiga atau lebih: (Barakat et al. 1995)
- Multiple referensi: (Abbott 1991; Kelso and Smith 1998)
- Kutipan langsung dengan halaman: (Wiegand and Gouws 2013: 273)

### Format Daftar Pustaka
- Urutkan alfabetis
- Untuk sumber website, WAJIB sertakan URL lengkap
- Format buku: Nama, Inisial. Tahun. Judul. Kota: Penerbit.
- Format website: Nama/Organisasi. Tahun. "Judul Artikel". URL. Diakses pada tanggal.
- Contoh website: Kompas. 2024. "Dampak Media Sosial pada Remaja". https://lifestyle.kompas.com/read/2024/dampak-media-sosial. Diakses pada 11 Mei 2026.
- Contoh buku: Arief, M. Sarief. 2010. Politik Film di Hindia Belanda. Jakarta: Komunitas Bambu.
- SEMUA referensi dari research harus masuk daftar pustaka dengan URL-nya

### Paraphrase
- JANGAN copy-paste dari referensi, tulis ulang dengan kalimat sendiri
- Gunakan sinonim dan struktur kalimat yang berbeda
- Tetap pertahankan makna dan akurasi informasi
- Setiap klaim dari referensi harus disertai kutipan

### Gambar
- Jika ada gambar dari referensi, sisipkan dengan format: ![keterangan gambar](path_file)
- Beri keterangan yang relevan pada setiap gambar

### Bahasa
- Gunakan bahasa akademik yang formal dan baku
- Hindari bahasa sehari-hari atau slang
- Gunakan kalimat pasif jika diperlukan untuk gaya akademik
"""


if __name__ == "__main__":
    mcp.run(transport="stdio")
