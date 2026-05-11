# MakalahMCP — Design Spec

## Overview

MCP server yang membantu generate makalah akademik. User memberikan judul, sistem melakukan research via Google Search scraping, Claude menyusun konten berdasarkan referensi (dengan paraphrase), lalu MCP menyimpan hasil ke file DOCX dan/atau PDF.

## Arsitektur

MCP Server dengan 2 tools, Claude sebagai orchestrator:

```
User → "buatkan makalah tentang X"
  → Claude panggil research_topic(title="X")
  → MCP scrape Google Search, fetch konten + gambar dari top results
  → MCP return referensi terstruktur (teks, gambar, metadata)
  → Claude generate makalah lengkap (paraphrase, struktur sesuai pedoman)
  → Claude panggil save_makalah(content="...", format="both", ...)
  → MCP convert markdown → DOCX (styled) → PDF
  → MCP return path file yang disimpan
```

## Tool 1: `research_topic`

### Input

| Parameter     | Type   | Required | Default | Description                    |
|---------------|--------|----------|---------|--------------------------------|
| `title`       | string | yes      | —       | Judul makalah untuk di-research |
| `num_results` | int    | no       | 5       | Jumlah hasil pencarian          |

### Proses

1. Build search query dari judul
2. HTTP GET ke Google Search (pakai httpx, set User-Agent browser)
3. Parse HTML hasil pencarian dengan BeautifulSoup, extract top N links
4. Untuk setiap link:
   - Fetch halaman
   - Extract konten utama (teks paragraf)
   - Extract gambar yang relevan (filter: ukuran min, skip icon/logo)
   - Extract metadata: judul halaman, author (jika ada), tahun (jika ada)
   - Download gambar ke direktori temporary
5. Return daftar referensi

### Output

```json
{
  "references": [
    {
      "title": "Judul Artikel",
      "url": "https://example.com/artikel",
      "author": "Nama Penulis",
      "year": "2024",
      "content": "Konten utama artikel yang di-extract...",
      "images": [
        {
          "url": "https://example.com/gambar.png",
          "caption": "Deskripsi gambar",
          "local_path": "/tmp/makalahmcp/images/img_001.png"
        }
      ]
    }
  ]
}
```

### Error Handling

- Jika Google Search gagal (rate limit, block): return error message
- Jika link gagal di-fetch: skip, lanjut ke link berikutnya
- Jika gambar gagal di-download: skip gambar tersebut, tetap return referensi teks

## Tool 2: `save_makalah`

### Input

| Parameter       | Type   | Required | Default          | Description                          |
|-----------------|--------|----------|------------------|--------------------------------------|
| `content`       | string | yes      | —                | Konten makalah dalam markdown        |
| `title`         | string | yes      | —                | Judul makalah (Indonesia)            |
| `title_en`      | string | no       | —                | Judul makalah (Inggris)              |
| `lecturer`      | string | no       | —                | Nama dosen pengampu                  |
| `author`        | string | no       | —                | Nama penulis                         |
| `nim`           | string | no       | —                | NIM penulis                          |
| `program_studi` | string | no       | —                | Program studi                        |
| `fakultas`      | string | no       | —                | Fakultas                             |
| `universitas`   | string | no       | —                | Nama universitas                     |
| `year`          | string | no       | tahun sekarang   | Tahun                                |
| `logo_path`     | string | no       | —                | Path ke file logo institusi          |
| `format`        | string | no       | "both"           | "docx", "pdf", atau "both"           |
| `output_dir`    | string | no       | "~/Documents"    | Direktori output                     |

### Proses

1. Parse markdown content
2. Generate DOCX dengan python-docx:
   - Buat cover page sesuai format
   - Apply styling sesuai pedoman penulisan
   - Sisipkan gambar dengan keterangan
   - Format daftar pustaka
3. Jika format PDF diminta: convert DOCX → PDF
4. Simpan file

### Output

```json
{
  "files": [
    "/path/to/Makalah - Judul.docx",
    "/path/to/Makalah - Judul.pdf"
  ]
}
```

## Format Cover

Layout cover page (center-aligned, single page):

```
[JUDUL MAKALAH - uppercase, bold, Times New Roman 14]
[Judul dalam Bahasa Inggris - italic, Times New Roman 14]
[Dosen Pengampu: nama dosen - Times New Roman 12]

[Logo Institusi - gambar, centered]

Disusun Oleh:
[NAMA PENULIS - uppercase, bold, Times New Roman 12]
NIM: [nomor - Times New Roman 12]

[PROGRAM STUDI - uppercase, bold, Times New Roman 12]
[FAKULTAS - uppercase, bold, Times New Roman 12]
[UNIVERSITAS - uppercase, bold, Times New Roman 12]
[Tahun - Times New Roman 12]
```

## Pedoman Penulisan

### Font & Formatting

- Font: Times New Roman seluruhnya
- Judul: bold, size 14
- Nama penulis: bold, size 12
- Institusi & email: size 11
- Body text: size 12, spasi 1.5
- Kertas: A4
- Margin: kiri 3.5cm, kanan 3cm, atas 3cm, bawah 3cm

### Struktur Isi Makalah

1. Cover
2. Daftar Isi
3. Kata Pengantar
4. **I. Abstrak** — Bahasa Indonesia dan Bahasa Inggris
5. **II. Pendahuluan** — Latar belakang, masalah, tujuan, tinjauan pustaka
6. **III. Metode Penelitian**
7. **IV. Hasil**
8. **V. Pembahasan**
9. **VI. Simpulan**
10. Catatan Akhir (jika ada)
11. Daftar Pustaka

### Penomoran Subbab

- Bab: angka romawi (I, II, III, ...)
- Sub-bab: `1.1`, `1.1.1` (maks 3 digit)

### Format Kutipan

- Dalam teks: `(Thompson 1990)`, `(Becker and Seligman 1996)`
- Multiple: `(Abbott 1991; Barakat et al. 1995; Kelso and Smith 1998)`
- Dengan halaman: `(Wiegand and Gouws 2013: 273)`

### Tabel dan Gambar

- Penomoran arab (Gambar 1, Tabel 1, ...)
- Setiap tabel/gambar diberi keterangan
- Sumber kutipan ditulis sesuai format daftar pustaka

### Daftar Pustaka

- Diurutkan alfabetis dan kronologis
- Format: `Nama, Inisial. Tahun. Judul. Kota: Penerbit.`
- Contoh: `Arief, M. Sarief. 2010. Politik Film di Hindia Belanda. Jakarta: Komunitas Bambu.`

## Struktur File Project

```
makalahmcp/
├── main.py           # MCP server entry point, tool definitions
├── scraper.py        # Google search + web scraping + image download
├── converter.py      # Markdown → DOCX/PDF conversion with styling
├── pyproject.toml    # Dependencies
└── uv.lock
```

## Dependencies

```
mcp[cli]>=1.27.1      # MCP server framework
httpx>=0.28.1          # HTTP client untuk scraping
beautifulsoup4         # HTML parsing
python-docx            # Generate DOCX
fpdf2                  # Generate PDF (lightweight, no system deps)
```

## Batasan & Catatan

- Google Search scraping bisa kena rate limit. Jika terjadi, tool return error message yang jelas.
- Gambar yang di-download disimpan di `/tmp/makalahmcp/images/` dan di-embed ke DOCX/PDF.
- PDF di-generate dari DOCX untuk konsistensi format.
- Bahasa output makalah tergantung instruksi user saat prompt.
