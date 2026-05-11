# MakalahMCP

MCP server untuk generate makalah akademik Indonesia secara otomatis. Sistem akan research topik dari web, generate konten lengkap dengan paraphrase, dan simpan sebagai file DOCX dan PDF dengan format akademik yang benar.

## Fitur

- **Research otomatis** — cari referensi dari web (DuckDuckGo)
- **Logo universitas otomatis** — cari dan download logo universitas untuk cover
- **Format akademik** — Times New Roman, margin sesuai pedoman, cover page, kata pengantar, daftar isi
- **Daftar isi clickable** — klik langsung navigate ke section
- **Paraphrase** — konten ditulis ulang dari referensi, bukan copy-paste
- **Daftar pustaka lengkap** — 8-15 referensi dengan URL sumber web
- **Output DOCX + PDF** — kedua format sekaligus

## Struktur Makalah

1. Cover (judul, logo, nama, NIM, dosen, institusi)
2. Kata Pengantar (auto-generated)
3. Daftar Isi (auto-generated, clickable, dengan titik-titik)
4. BAB I - Abstrak (Indonesia + Inggris)
5. BAB II - Pendahuluan (latar belakang, rumusan masalah, tujuan, manfaat, tinjauan pustaka)
6. BAB III - Metode Penelitian
7. BAB IV - Hasil
8. BAB V - Pembahasan
9. BAB VI - Simpulan
10. Daftar Pustaka

## Prasyarat

- Python 3.13+
- [uv](https://docs.astral.sh/uv/) (package manager)

## Instalasi

```bash
# Clone repository
git clone <repo-url>
cd makalahmcp

# Install dependencies
uv sync
```

## Cara Pakai

### 1. Test dengan MCP Inspector

```bash
uv run mcp dev main.py
```

Buka browser ke URL yang ditampilkan, lalu test tools di inspector.

### 2. Tambah ke Claude Desktop

Edit file config Claude Desktop:

- **macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
- **Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

Tambahkan:

```json
{
  "mcpServers": {
    "makalahmcp": {
      "command": "uv",
      "args": ["--directory", "/path/ke/makalahmcp", "run", "main.py"]
    }
  }
}
```

Ganti `/path/ke/makalahmcp` dengan path absolut ke folder project.

### 3. Tambah ke Claude Code

Edit file `~/.claude/settings.json` atau jalankan:

```bash
claude mcp add makalahmcp -- uv --directory /path/ke/makalahmcp run main.py
```

## Penggunaan

Setelah MCP server terhubung, cukup minta Claude:

> "Buatkan makalah tentang Dampak Media Sosial Terhadap Remaja"

Claude akan otomatis:
1. Tanya data kamu (nama, NIM, universitas, prodi, fakultas, dosen pengampu)
2. Cari logo universitas
3. Research topik dari web
4. Generate konten makalah lengkap
5. Simpan ke DOCX dan PDF di `~/Documents/`

## Tools

| Tool | Fungsi |
|------|--------|
| `research_topic` | Research topik dari web, ambil referensi |
| `search_logo` | Cari dan download logo universitas |
| `save_makalah` | Simpan makalah ke DOCX/PDF |

## Prompt

| Prompt | Fungsi |
|--------|--------|
| `generate_makalah` | Instruksi lengkap untuk generate konten makalah |

## Struktur Project

```
makalahmcp/
├── main.py           # MCP server, tool & prompt definitions
├── scraper.py        # Web search, content extraction, logo search
├── converter.py      # Markdown → DOCX/PDF conversion
├── pyproject.toml    # Dependencies
└── README.md
```

## Format Output

### DOCX
- Times New Roman 12pt, spasi 1.5
- Margin: kiri 3.5cm, kanan 3cm, atas 3cm, bawah 3cm
- Kertas A4
- Teks justify dengan indent paragraf pertama
- Heading BAB centered, sub-heading left-aligned bold
- Daftar pustaka hanging indent

### PDF
- Format sama dengan DOCX
- Nomor halaman: romawi (ii, iii) untuk front matter, arab (1, 2, 3) untuk konten
- Daftar isi clickable (internal links)
