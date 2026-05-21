import json
import os
import sys
import logging
import warnings
from datetime import datetime
from pypdf import PdfReader

from mcp.server.fastmcp import FastMCP

from scraper import research_topic as do_research, search_university_logo as do_search_logo
from converter import save_as_docx, save_as_pdf

# Suppress pypdf encoding warnings and errors globally
warnings.filterwarnings('ignore', category=UserWarning, module='pypdf')
logging.getLogger('pypdf').setLevel(logging.CRITICAL)
logging.getLogger('pypdf._utils').setLevel(logging.CRITICAL)

PEDOMAN_DIR = os.path.join(os.path.expanduser("~"), ".makalahmcp_pedoman")
os.makedirs(PEDOMAN_DIR, exist_ok=True)

mcp = FastMCP(
    "MakalahMCP",
    instructions="""MCP server for generating Indonesian academic papers (makalah).

FLOW WAJIB saat user minta buat makalah:

## STEP 0: DETEKSI OTOMATIS - LAKUKAN INI PERTAMA KALI
Sebelum melakukan apapun, CEK prompt user untuk:

### A. Deteksi Link Referensi
- Cari pattern URL: https://, http://, www.
- Cari keyword: "Referensi:", "referensi:", "sumber:", "link:"
- Jika TIDAK ADA URL sama sekali → set flag: needs_research = True
- Jika ADA minimal 1 URL → set flag: needs_research = False, has_references = True

### B. Deteksi Data Mahasiswa
Cek apakah prompt mengandung SEMUA data berikut:
- Nama (cari: "Nama:", "nama:", atau setelah "oleh")
- NIM (cari: "NIM:", "nim:", atau angka 8-12 digit)
- Program Studi (cari: "Program Studi:", "Prodi:", "program studi:")
- Fakultas (cari: "Fakultas:", "fakultas:")
- Universitas (cari: "Universitas:", "universitas:", atau nama universitas)

Jika TIDAK LENGKAP (kurang 1 atau lebih) → set flag: needs_student_data = True
Jika LENGKAP (semua ada) → set flag: needs_student_data = False

### C. Action Berdasarkan Flag
1. Jika needs_student_data = True:
   - TANYA data yang kurang ke user
   - JANGAN lanjut sebelum dapat semua data
   - Format: "Saya perlu data berikut untuk membuat makalah: [list data yang kurang]. Mohon berikan informasinya."

2. Jika needs_research = True (tidak ada referensi):
   - Setelah dapat data mahasiswa, panggil research_topic(title)
   - Gunakan hasil research untuk generate_makalah

3. Jika has_references = True (ada referensi):
   - Extract semua URL dari prompt
   - Format jadi JSON: [{"title": "...", "url": "...", "content": "..."}]
   - Skip research_topic, langsung ke generate_makalah

## Jika user kasih semua data sekaligus (contoh: "Buat makalah dengan judul X, Nama: Y, NIM: Z, ..."):
1. EXTRACT semua informasi dari prompt user:
   - Judul makalah (cari kata "judul", "tentang", atau setelah "makalah")
   - Nama (cari "Nama:", "nama:", atau setelah "oleh")
   - NIM (cari "NIM:", "nim:", atau angka panjang)
   - Program Studi (cari "Program Studi:", "Prodi:", "program studi:")
   - Fakultas (cari "Fakultas:", "fakultas:")
   - Universitas (cari "Universitas:", "universitas:", biasanya di akhir)
   - Dosen Pengampu (cari "Dosen:", "dosen:", "pengampu:", opsional)
2. KONFIRMASI data yang di-extract: "Saya akan buat makalah dengan data: [list semua data]. Apakah sudah benar?"
3. TUNGGU konfirmasi dari user (ya/benar/ok/lanjut/betul)
4. JANGAN langsung proses sebelum user konfirmasi!
5. Jika user bilang ada yang salah, tanya data yang benar
6. Setelah user konfirmasi, lanjut ke step 7
7. Cari logo universitas dengan tool search_logo
8. Tahun OTOMATIS pakai tahun sekarang (2026), JANGAN tanya ke user
9. Cek apakah ada pedoman dengan get_pedoman — jika ada, WAJIB ikuti struktur dari pedoman
10. JIKA user TIDAK kasih referensi sendiri, lakukan research_topic dengan judul
11. JIKA user SUDAH kasih referensi, skip research_topic dan gunakan referensi user
12. Generate konten makalah:
    - Panggil prompt generate_makalah dengan 2 argument WAJIB:
      * title: judul makalah (string)
      * references_json: hasil dari research_topic (JSON string) ATAU referensi user yang sudah diformat jadi JSON
    - Contoh: generate_makalah(title="Implementasi AI", references_json='[{"title":"...","url":"...","content":"..."}]')
13. TUNJUKKAN konten makalah ke user dan TANYA: "Apakah konten sudah sesuai? Ada yang ingin diubah?"
14. TUNGGU feedback dari user, JANGAN langsung save!
15. Jika user minta edit/ubah, lakukan perubahan sesuai permintaan user lalu tunjukkan lagi
16. Jika user sudah setuju (bilang "ok", "sudah", "lanjut", "simpan", dll), baru simpan dengan save_makalah (format "both")
17. Setelah disimpan, berikan link download (BUKAN path file lokal):
    - Format: http://mcp.asln.dev/makalah/download/[filename dengan URL encoding]
    - Contoh: http://mcp.asln.dev/makalah/download/Makalah%20-%20Implementasi%20AI%20di%20Sistem%20HRIS.docx
    - JANGAN tampilkan path seperti /root/Documents/... atau ~/Documents/...
    - Gunakan URL encoding untuk spasi (%20) dan karakter khusus
    - Tampilkan sebagai clickable link untuk user
18. Tanyakan: "File sudah disimpan. Ada yang ingin diubah lagi?"

## Jika user TIDAK kasih data lengkap:
1. Tanyakan JUDUL makalah (jika belum ada)
2. Tanyakan data penulis (satu per satu atau sekaligus):
   - Nama lengkap (wajib)
   - NIM (wajib)
   - Universitas (wajib)
   - Program Studi (wajib)
   - Fakultas (wajib)
   - Dosen Pengampu (opsional, boleh dikosongkan)
3. KONFIRMASI semua data: "Data yang saya terima: [list]. Apakah sudah benar?"
4. TUNGGU konfirmasi user sebelum lanjut
5. Lanjut ke step 7 di atas

## Jika user kasih REFERENSI sendiri:
- User bisa kasih referensi dalam format:
  * Dengan keyword: "Referensi: https://url1.com, https://url2.com"
  * Tanpa keyword: Langsung URL "https://url1.com" atau list URL
  * List judul + URL: "Referensi: 1. Judul (https://url1.com) 2. Judul (https://url2.com)"
  * Text biasa: User paste text/kutipan dari sumber
- Detect referensi dengan cara:
  * Cari keyword "Referensi:" atau "referensi:" ATAU
  * Detect URL pattern (https://..., http://...) di prompt user
  * Jika ada 1+ URL di prompt, anggap itu referensi user
- JANGAN panggil research_topic jika user sudah kasih referensi (ada URL atau keyword)
- Format referensi user menjadi JSON untuk generate_makalah prompt:
  * Format JSON: [{"title": "judul/deskripsi", "url": "https://...", "author": "Unknown", "year": "2026", "content": "ringkasan/kutipan"}]
  * Jika user kasih URL saja, buat title dari URL atau "Referensi 1", "Referensi 2", dst
  * Jika user kasih judul + URL, gunakan judul yang diberikan
  * Jika user kasih text/kutipan, masukkan ke field "content"
- Konfirmasi: "Saya akan gunakan referensi yang Anda berikan (X referensi). Lanjut generate?"
- Lalu panggil generate_makalah(title="...", references_json='[...]') dengan JSON yang sudah diformat

## Jika user kasih file PDF sebagai contoh/pedoman:
- Langsung panggil set_pedoman dengan path PDF tersebut
- JANGAN tanya user apakah mau set pedoman, langsung simpan saja
- Konfirmasi ke user: "Pedoman sudah disimpan, struktur makalah selanjutnya akan mengikuti format ini."

## PENTING - Parsing Data dari Prompt:
- Nama bisa setelah "Nama:", "nama:", "Nama :", atau "oleh"
- NIM bisa setelah "NIM:", "nim:", "NIM :", atau angka panjang (8-12 digit)
- Program Studi bisa "Program Studi:", "Prodi:", "program studi:", atau singkatan seperti "PJJ Informatika"
- Fakultas bisa setelah "Fakultas:", "fakultas:", "Fakultas :"
- Universitas biasanya di akhir, setelah "Universitas:", "universitas:", atau nama lengkap universitas
- Judul makalah biasanya dalam tanda petik, setelah "judul", "tentang", atau di awal prompt

## CRITICAL - JANGAN LANGSUNG PROSES:
- SELALU konfirmasi data mahasiswa dulu sebelum research/generate
- SELALU tunjukkan preview konten sebelum save
- TUNGGU user bilang "ok"/"simpan"/"lanjut" sebelum save_makalah
- Jika user tidak konfirmasi, JANGAN lanjut ke step berikutnya
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
async def set_pedoman(pdf_path: str) -> str:
    """Simpan PDF sebagai pedoman/template struktur makalah.

    Bisa menyimpan beberapa pedoman sekaligus. Sistem akan mempelajari
    struktur dari semua pedoman yang tersimpan.

    Args:
        pdf_path: Path absolut ke file PDF pedoman.

    Returns:
        JSON string dengan struktur yang diekstrak dan konfirmasi.
    """
    if not os.path.exists(pdf_path):
        return json.dumps({"error": f"File tidak ditemukan: {pdf_path}"})

    try:
        reader = PdfReader(pdf_path)
        all_text = []
        for i, page in enumerate(reader.pages):
            try:
                text = page.extract_text()
                if text and text.strip():
                    all_text.append(f"[Halaman {i + 1}]\n{text.strip()}")
            except Exception as e:
                # Skip pages with encoding issues
                all_text.append(f"[Halaman {i + 1}]\n[Error extracting text: {str(e)}]")
                continue

        full_text = "\n\n".join(all_text)

        # Extract structure
        lines = full_text.split("\n")
        structure = []
        for line in lines:
            line = line.strip()
            if not line:
                continue
            upper = line.upper()
            if upper.startswith("BAB ") or upper in ("DAFTAR PUSTAKA", "KATA PENGANTAR", "DAFTAR ISI", "KESIMPULAN", "PENUTUP", "PENDAHULUAN", "PEMBAHASAN"):
                structure.append({"type": "heading", "level": 1, "text": line})
            elif len(line) < 200 and line[0].isdigit() and "." in line[:5]:
                structure.append({"type": "heading", "level": 2, "text": line})

        # Save with unique name based on filename
        base_name = os.path.splitext(os.path.basename(pdf_path))[0]
        safe_name = "".join(c if c.isalnum() or c in " -_" else "" for c in base_name)[:50].strip()
        pedoman_file = os.path.join(PEDOMAN_DIR, f"{safe_name}.json")

        pedoman = {
            "source_file": pdf_path,
            "name": base_name,
            "total_pages": len(reader.pages),
            "structure": structure,
            "full_text": full_text[:10000],
        }

        with open(pedoman_file, "w") as f:
            json.dump(pedoman, f, ensure_ascii=False, indent=2)

        # Count total pedoman
        total = len([f for f in os.listdir(PEDOMAN_DIR) if f.endswith(".json")])

        summary = f"Pedoman '{base_name}' disimpan ({len(reader.pages)} halaman, {len(structure)} heading).\n"
        summary += f"Total pedoman tersimpan: {total}\n\nStruktur:\n"
        for s in structure:
            prefix = "  " if s["level"] == 2 else ""
            summary += f"{prefix}{s['text']}\n"

        return json.dumps({"success": True, "summary": summary}, ensure_ascii=False, indent=2)

    except Exception as e:
        return json.dumps({"error": f"Gagal membaca PDF: {str(e)}"})


@mcp.tool()
async def get_pedoman() -> str:
    """Lihat semua pedoman/template yang tersimpan.

    Returns:
        JSON string dengan daftar pedoman dan strukturnya.
    """
    files = [f for f in os.listdir(PEDOMAN_DIR) if f.endswith(".json")] if os.path.exists(PEDOMAN_DIR) else []

    if not files:
        return json.dumps({"has_pedoman": False, "message": "Belum ada pedoman. Gunakan set_pedoman untuk menambah template."})

    pedoman_list = []
    for f in files:
        with open(os.path.join(PEDOMAN_DIR, f)) as fh:
            data = json.load(fh)
            pedoman_list.append({
                "name": data.get("name", f),
                "source": data.get("source_file", ""),
                "pages": data.get("total_pages", 0),
                "structure": data.get("structure", []),
            })

    return json.dumps({"has_pedoman": True, "count": len(pedoman_list), "pedoman": pedoman_list}, ensure_ascii=False, indent=2)


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
        output_dir: Output directory (default: /tmp).

    Returns:
        JSON string with list of download URLs (format: http://mcp.asln.dev/makalah/download/filename).
        Each entry contains: filename, url (download link), and local_path.
    """
    if not output_dir:
        output_dir = "/tmp"
    os.makedirs(output_dir, exist_ok=True)

    safe_title = "".join(c if c.isalnum() or c in " -_" else "" for c in title)[:80].strip()
    base_name = f"Makalah - {safe_title}"
    if not year:
        year = str(datetime.now().year)

    if format not in ("docx", "pdf", "both"):
        return json.dumps({"error": f"Format '{format}' tidak valid. Gunakan 'docx', 'pdf', atau 'both'."})

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

    # Generate download URLs instead of local paths
    from urllib.parse import quote
    download_urls = []
    for file_path in files:
        filename = os.path.basename(file_path)
        encoded_filename = quote(filename)
        download_url = f"http://mcp.asln.dev/makalah/download/{encoded_filename}"
        download_urls.append({
            "filename": filename,
            "url": download_url,
            "local_path": file_path
        })

    return json.dumps({"files": download_urls}, ensure_ascii=False, indent=2)


@mcp.prompt()
def generate_makalah(title: str, references_json: str = "[]") -> str:
    """Generate a complete makalah (academic paper) from research references.

    Use this prompt after calling research_topic to get the proper format
    and instructions for generating comprehensive makalah content.

    Args:
        title: The makalah title.
        references_json: JSON string from research_topic tool output (default: empty list).
    """
    # Load all pedoman if any exist
    pedoman_section = ""
    if os.path.exists(PEDOMAN_DIR):
        files = [f for f in os.listdir(PEDOMAN_DIR) if f.endswith(".json")]
        if files:
            pedoman_section = "\n## PEDOMAN/TEMPLATE YANG HARUS DIIKUTI\n"
            pedoman_section += "Berikut adalah contoh-contoh struktur makalah dari pedoman yang tersimpan.\n"
            pedoman_section += "Pelajari SEMUA pedoman ini, lalu buat struktur yang KONSISTEN mengikuti pola yang sama.\n\n"
            for fname in files:
                with open(os.path.join(PEDOMAN_DIR, fname)) as fh:
                    data = json.load(fh)
                name = data.get("name", fname)
                structure = data.get("structure", [])
                if structure:
                    pedoman_section += f"### Pedoman: {name}\n"
                    for s in structure:
                        prefix = "  " if s["level"] == 2 else ""
                        pedoman_section += f"{prefix}{s['text']}\n"
                    pedoman_section += "\n"
            pedoman_section += "Gunakan struktur pedoman di atas sebagai kerangka. Sesuaikan judul sub-bab dengan topik makalah, tapi PERTAHANKAN jumlah BAB, urutan, dan pola yang sama.\n"

    return f"""Kamu adalah penulis akademik. Buat makalah lengkap dan komprehensif dengan judul:
"{title}"

Gunakan referensi berikut sebagai sumber:
{references_json}
{pedoman_section}
## ATURAN PENTING

### Panjang Konten — INI ATURAN PALING PENTING, TIDAK BOLEH DILANGGAR
- Target MINIMAL 6000-10000 kata. Makalah yang kurang dari 5000 kata TIDAK DITERIMA.
- Setiap sub-bab (##) harus memiliki MINIMAL 3-5 paragraf PANJANG
- Setiap paragraf WAJIB terdiri dari MINIMAL 5-8 kalimat panjang
- DILARANG KERAS menulis paragraf yang hanya 1-3 kalimat
- Setiap sub-bab harus MINIMAL 300-500 kata
- Elaborasi setiap poin dengan: definisi lengkap, penjelasan mendalam, contoh konkret, data/angka spesifik, dampak/implikasi, hubungan dengan teori, dan perbandingan dengan penelitian terdahulu
- Gunakan BANYAK referensi dan kutipan dalam SETIAP paragraf, bukan hanya di akhir sub-bab
- Tulis seolah-olah ini adalah makalah untuk tugas akhir kuliah yang nilainya menentukan kelulusan

### Aturan Per Sub-bab — SETIAP sub-bab WAJIB mengikuti pola ini:
1. Paragraf pembuka: jelaskan konteks dan relevansi topik sub-bab (5-8 kalimat)
2. Paragraf inti 1: jelaskan konsep/temuan utama dengan detail, kutip referensi (5-8 kalimat)
3. Paragraf inti 2: berikan contoh konkret, data spesifik, atau perbandingan (5-8 kalimat)
4. Paragraf inti 3: hubungkan dengan teori atau penelitian terdahulu, kutip referensi (5-8 kalimat)
5. Paragraf penutup: simpulkan dan transisi ke sub-bab berikutnya (3-5 kalimat)

### Contoh Penulisan
Contoh BURUK (DILARANG — terlalu singkat, hanya 1-2 kalimat):
"Dominasi masalah jaringan sebesar 35 persen memberikan informasi strategis bagi manajemen."

Contoh BAIK (WAJIB — detail, panjang, banyak kutipan):
"Dominasi masalah jaringan sebesar 35 persen dari total tiket layanan bantuan merupakan temuan yang sangat signifikan dan memerlukan perhatian khusus dari pihak manajemen organisasi. Angka ini menunjukkan bahwa lebih dari sepertiga seluruh permasalahan TI yang dilaporkan oleh pengguna berkaitan dengan konektivitas dan infrastruktur jaringan, yang mencakup permasalahan seperti putusnya koneksi internet, lambatnya kecepatan akses, gangguan pada perangkat jaringan, serta masalah konfigurasi yang menyebabkan ketidakstabilan layanan. Temuan ini sejalan dengan hasil penelitian Silvianingsih (2024) yang juga mengidentifikasi bahwa permasalahan infrastruktur jaringan merupakan kategori dominan dalam data operasional perusahaan penyedia layanan internet di Indonesia. Berdasarkan kerangka kerja ITIL yang dikembangkan oleh Axelos (2019), organisasi perlu melakukan analisis akar masalah secara sistematis terhadap kategori permasalahan yang paling dominan untuk mengidentifikasi faktor-faktor penyebab utama dan merancang tindakan perbaikan yang efektif. Implikasi praktis dari temuan ini adalah bahwa manajemen perlu mengalokasikan proporsi anggaran yang lebih besar untuk pemeliharaan dan peningkatan infrastruktur jaringan, termasuk investasi dalam perangkat jaringan yang lebih handal, peningkatan kapasitas bandwidth, serta pelatihan berkala bagi tim teknis yang menangani permasalahan jaringan. Selain itu, organisasi juga perlu mempertimbangkan implementasi sistem pemantauan jaringan secara waktu nyata yang dapat mendeteksi potensi gangguan sebelum berdampak pada pengguna akhir, sebagaimana direkomendasikan oleh Turban dkk. (2018) dalam konteks penerapan inteligensi bisnis untuk operasional TI."

Tulis SEMUA paragraf dengan level detail dan panjang seperti contoh BAIK di atas. JANGAN PERNAH menulis paragraf pendek.

### FORMAT MARKDOWN — IKUTI PERSIS, JANGAN DIUBAH

DILARANG KERAS:
- JANGAN tulis Kata Pengantar, Daftar Isi, atau Cover. Semua itu OTOMATIS oleh sistem.
- JANGAN tulis apapun sebelum # I. Abstrak
- JANGAN pakai #### (4 hash atau lebih). HANYA gunakan # (BAB) dan ## (sub-bab) dan ### (sub-sub-bab)
- JANGAN pakai **bold** atau *italic* di dalam teks. Tulis teks biasa saja tanpa formatting markdown.
- JANGAN pakai format heading lain selain yang ditentukan di bawah ini

FORMAT HEADING YANG BENAR (COPY PERSIS, nomor sub-bab ikut nomor BAB):
```
# I. Abstrak
# II. Pendahuluan
## 2.1 Latar Belakang
## 2.2 Rumusan Masalah
## 2.3 Tujuan Penulisan
## 2.4 Manfaat Penulisan
## 2.5 Tinjauan Pustaka
# III. Metode Penelitian
## 3.1 Pendekatan Penelitian
## 3.2 Tahapan Penelitian
## 3.3 Alat dan Bahan
## 3.4 Teknik Analisis Data
# IV. Hasil
## 4.1 [Judul sub-bab sesuai topik]
## 4.2 [Judul sub-bab sesuai topik]
## 4.3 [Judul sub-bab sesuai topik]
# V. Pembahasan
## 5.1 [Judul sub-bab sesuai topik]
## 5.2 [Judul sub-bab sesuai topik]
## 5.3 [Judul sub-bab sesuai topik]
## 5.4 Implikasi Praktis
## 5.5 Keterbatasan Penelitian
# VI. Simpulan
## 6.1 Kesimpulan
## 6.2 Saran
# Daftar Pustaka
```

ATURAN HEADING:
- BAB menggunakan # (satu hash) diikuti angka romawi: # I. , # II. , # III. , dst.
- Sub-bab menggunakan ## (dua hash) diikuti nomor desimal: ## 1.1 , ## 1.2 , ## 3.1 , dst.
- Sub-sub-bab menggunakan ### (tiga hash) jika diperlukan: ### 3.1.1 , ### 3.1.2 , dst.
- SETIAP BAB WAJIB punya minimal 2 sub-bab (##)
- JANGAN pakai heading level 4 atau lebih (####, #####)

# Daftar Pustaka

### Format Kutipan Dalam Teks
- Satu penulis: (Thompson 1990)
- Dua penulis: (Becker and Seligman 1996)
- Tiga atau lebih: (Barakat et al. 1995)
- Multiple referensi: (Abbott 1991; Kelso and Smith 1998)
- Kutipan langsung dengan halaman: (Wiegand and Gouws 2013: 273)

### Format Daftar Pustaka — HARUS BANYAK (minimal 8-15 referensi) dan SEMUA WAJIB ADA URL

ATURAN DAFTAR PUSTAKA YANG TIDAK BOLEH DILANGGAR:
1. Urutkan alfabetis
2. SETIAP referensi WAJIB diakhiri dengan URL dan "Diakses pada [tanggal]". TIDAK ADA PENGECUALIAN.
3. Referensi TANPA URL akan DITOLAK oleh sistem.

FORMAT WAJIB setiap entry daftar pustaka (perhatikan URL di akhir):
```
Nama, Inisial. Tahun. "Judul". Penerbit/Jurnal. URL. Diakses pada [tanggal].
```

CONTOH YANG BENAR:
```
Sugiyono. 2019. "Metode Penelitian Kuantitatif, Kualitatif, dan R&D". Bandung: Alfabeta. https://opac.perpusnas.go.id/DetailOpac.aspx?id=1133109. Diakses pada 12 Mei 2026.
Laudon, K. C. dan Laudon, J. P. 2020. "Management Information Systems". New York: Pearson. https://books.google.co.id/books?id=o1yuDwAAQBAJ. Diakses pada 12 Mei 2026.
Wibowo, A. 2020. "Analisis Deskriptif Data Operasional TI". Jurnal Teknologi Informasi, 12(3), 45-58. https://jurnal.example.ac.id/index.php/jti/article/view/123. Diakses pada 12 Mei 2026.
```

CONTOH YANG SALAH (JANGAN SEPERTI INI — TIDAK ADA URL):
```
Laudon, K. C. dan Laudon, J. P. 2020. Management Information Systems. New York: Pearson.
```

SUMBER URL:
- Buku: gunakan https://books.google.co.id/books?id=... atau https://opac.perpusnas.go.id/...
- Jurnal: gunakan URL portal jurnal (https://jurnal.*.ac.id/..., https://scholar.google.com/..., dll)
- Website: gunakan URL asli dari research

ATURAN TAMBAHAN:
- SEMUA referensi dari research WAJIB masuk dengan URL asli-nya
- TAMBAHKAN minimal 5 referensi pendukung dari buku/jurnal (HARUS punya URL)
- Setiap referensi HARUS dikutip minimal 1 kali di dalam teks
- JANGAN buat referensi fiktif

### Paraphrase
- JANGAN copy-paste dari referensi, tulis ulang dengan kalimat sendiri
- Gunakan sinonim dan struktur kalimat yang berbeda
- Tetap pertahankan makna dan akurasi informasi
- Setiap klaim dari referensi harus disertai kutipan

### Gambar — WAJIB ADA GAMBAR
- SETIAP makalah WAJIB menyertakan gambar yang relevan dengan topik
- Gambar dari hasil research (yang sudah di-download oleh research_topic) WAJIB dimasukkan
- Lihat field "images" di setiap referensi dari research_topic — sisipkan semua gambar yang ada
- Format: ![Keterangan gambar yang jelas dan deskriptif](local_path_dari_images)
- Contoh: ![Gambar 1: Arsitektur sistem dashboard interaktif](/tmp/makalahmcp/images/img_001.png)
- Letakkan gambar di bagian yang relevan (BAB Hasil atau BAB Pembahasan)
- Setiap gambar HARUS memiliki keterangan yang menjelaskan isi gambar
- Penomoran gambar menggunakan angka arab: Gambar 1, Gambar 2, dst.
- Sumber gambar harus dicantumkan jika berasal dari referensi

### Bahasa — FULL BAHASA INDONESIA
- SELURUH makalah WAJIB ditulis dalam Bahasa Indonesia yang baku dan formal
- KECUALI bagian abstrak bahasa Inggris, SEMUA konten harus Bahasa Indonesia
- Heading dan sub-heading dalam Bahasa Indonesia
- Istilah asing boleh digunakan jika sudah umum (misal: dashboard, software, hardware) tapi berikan penjelasan dalam Bahasa Indonesia saat pertama kali muncul
- Gunakan bahasa akademik formal, hindari bahasa sehari-hari atau slang
- Gunakan kalimat pasif jika diperlukan untuk gaya akademik
- JANGAN campur bahasa Inggris di dalam paragraf kecuali istilah teknis yang memang tidak ada padanannya

### Detail Konten — SETIAP POIN HARUS DIELABORASI
- Jangan hanya menyebutkan poin, tapi JELASKAN secara mendalam
- Setiap poin dalam rumusan masalah harus dibahas tuntas di BAB Pembahasan
- Setiap sub-bab harus memiliki pengantar, isi detail, dan penutup/transisi ke sub-bab berikutnya
- Jika menyebutkan data/angka, jelaskan konteks, signifikansi, dan implikasinya
- Jika menyebutkan teori, jelaskan definisi, penemu, dan relevansinya dengan topik
- Jika menyebutkan hasil, jelaskan proses, temuan spesifik, dan interpretasinya
- BAB Pembahasan harus paling panjang — analisis mendalam setiap temuan, bandingkan dengan teori dan penelitian terdahulu
"""


@mcp.custom_route("/download/{filename}", methods=["GET"])
async def download_file(request):
    """Download generated makalah file"""
    from starlette.responses import FileResponse, JSONResponse

    filename = request.path_params['filename']
    output_dir = "/tmp"
    file_path = os.path.join(output_dir, filename)

    if not os.path.exists(file_path):
        return JSONResponse({"error": "File not found", "path": file_path}, status_code=404)

    # Determine media type
    if filename.endswith(".docx"):
        media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif filename.endswith(".pdf"):
        media_type = "application/pdf"
    else:
        media_type = "application/octet-stream"

    return FileResponse(
        path=file_path,
        media_type=media_type,
        filename=filename
    )


@mcp.custom_route("/files", methods=["GET"])
async def list_files(request):
    """List all generated makalah files"""
    from starlette.responses import JSONResponse

    output_dir = "/tmp"

    if not os.path.exists(output_dir):
        return JSONResponse({"files": [], "count": 0})

    files = []
    for filename in os.listdir(output_dir):
        if filename.startswith("Makalah - ") and (filename.endswith(".docx") or filename.endswith(".pdf")):
            file_path = os.path.join(output_dir, filename)
            stat = os.stat(file_path)
            files.append({
                "filename": filename,
                "size": stat.st_size,
                "created": datetime.fromtimestamp(stat.st_ctime).isoformat(),
                "download_url": f"/download/{filename}"
            })

    return JSONResponse({"files": files, "count": len(files)})


if __name__ == "__main__":
    import sys
    if "--http" in sys.argv:
        mcp.settings.host = os.environ.get("HOST", "0.0.0.0")
        mcp.settings.port = int(os.environ.get("PORT", "8000"))
        mcp.run(transport="streamable-http")
    else:
        mcp.run(transport="stdio")
