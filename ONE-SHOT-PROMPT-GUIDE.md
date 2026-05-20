# Cara Pakai MakalahMCP - One-Shot Prompt

MCP server sekarang bisa handle prompt lengkap dalam satu kali input!

## ✅ **Format Prompt yang Didukung:**

### **Format 1: Lengkap dengan Label**
```
Buat makalah dengan judul "Implementasi AI di sistem HRIS"
Nama: Egi Aprilianto
NIM: 12973702883
Program Studi: PJJ Informatika
Fakultas: Sistem Informasi
Universitas: Siber Asia
```

### **Format 2: Inline**
```
Buat makalah tentang "Implementasi AI di sistem HRIS" oleh Egi Aprilianto (NIM: 12973702883), Program Studi PJJ Informatika, Fakultas Sistem Informasi, Universitas Siber Asia
```

### **Format 3: Mixed**
```
Buatkan makalah "Implementasi AI di sistem HRIS"
nama: Egi Aprilianto nim: 12973702883 
prodi: PJJ Informatika fakultas: Sistem Informasi 
universitas: Siber Asia
```

---

## 🤖 **Apa yang Akan Terjadi:**

### **Step 1: Extract Data**
MCP akan otomatis extract:
- ✅ **Judul**: "Implementasi AI di sistem HRIS"
- ✅ **Nama**: Egi Aprilianto
- ✅ **NIM**: 12973702883
- ✅ **Program Studi**: PJJ Informatika
- ✅ **Fakultas**: Sistem Informasi
- ✅ **Universitas**: Siber Asia
- ✅ **Tahun**: 2026 (otomatis)

### **Step 2: Konfirmasi**
```
Saya akan buat makalah dengan data:
- Judul: Implementasi AI di sistem HRIS
- Nama: Egi Aprilianto
- NIM: 12973702883
- Program Studi: PJJ Informatika
- Fakultas: Sistem Informasi
- Universitas: Siber Asia
- Tahun: 2026

Apakah sudah benar?
```

### **Step 3: User Konfirmasi**
User reply: `ya` / `benar` / `ok` / `lanjut`

### **Step 4: Generate**
MCP akan:
1. ✅ Search logo Universitas Siber Asia
2. ✅ Research topic dari web
3. ✅ Generate konten makalah
4. ✅ Show preview ke user
5. ✅ Save ke DOCX + PDF
6. ✅ Return download links

### **Step 5: Download**
```
✅ Makalah berhasil dibuat!

Download:
- DOCX: http://mcp.asln.dev/makalah/download/Makalah%20-%20Implementasi%20AI%20di%20sistem%20HRIS.docx
- PDF: http://mcp.asln.dev/makalah/download/Makalah%20-%20Implementasi%20AI%20di%20sistem%20HRIS.pdf

Ada yang ingin diubah?
```

---

## 📝 **Parsing Rules:**

MCP akan cari pattern ini:

| Field | Pattern yang Dicari |
|-------|---------------------|
| **Judul** | Dalam tanda petik, setelah "judul", "tentang", atau di awal |
| **Nama** | Setelah "Nama:", "nama:", "oleh" |
| **NIM** | Setelah "NIM:", "nim:", atau angka 8-12 digit |
| **Program Studi** | Setelah "Program Studi:", "Prodi:", "program studi:" |
| **Fakultas** | Setelah "Fakultas:", "fakultas:" |
| **Universitas** | Setelah "Universitas:", "universitas:", atau nama lengkap |
| **Dosen** | Setelah "Dosen:", "dosen:", "pengampu:" (opsional) |

---

## 🎯 **Contoh Lengkap di Hermes Agent:**

### **User Input:**
```
Buat makalah dengan judul "Implementasi AI di sistem HRIS"
Nama: Egi Aprilianto
NIM: 12973702883
Program Studi: PJJ Informatika
Fakultas: Sistem Informasi
Universitas: Siber Asia
```

### **Hermes Agent Response:**
```
Saya akan buat makalah dengan data:
- Judul: Implementasi AI di sistem HRIS
- Nama: Egi Aprilianto
- NIM: 12973702883
- Program Studi: PJJ Informatika
- Fakultas: Sistem Informasi
- Universitas: Siber Asia
- Tahun: 2026

Apakah sudah benar?
```

### **User:** 
```
ya
```

### **Hermes Agent:**
```
[Searching logo...]
[Researching topic...]
[Generating content...]

Preview konten makalah:

# I. Abstrak
[konten...]

# II. Pendahuluan
[konten...]

...

Apakah konten sudah sesuai? Ada yang ingin diubah?
```

### **User:**
```
ok simpan
```

### **Hermes Agent:**
```
✅ Makalah berhasil dibuat!

Download:
- DOCX: http://mcp.asln.dev/makalah/download/Makalah%20-%20Implementasi%20AI%20di%20sistem%20HRIS.docx
- PDF: http://mcp.asln.dev/makalah/download/Makalah%20-%20Implementasi%20AI%20di%20sistem%20HRIS.pdf

File sudah tersimpan. Ada yang ingin diubah lagi?
```

---

## 🔧 **Update Server:**

Untuk apply perubahan ini di server:

```bash
cd /root/makalahmcp
./update-server.sh
```

---

## ✅ **Testing:**

Test dengan Hermes Agent menggunakan prompt:

```
Buat makalah dengan judul "Implementasi AI di sistem HRIS"
Nama: Egi Aprilianto
NIM: 12973702883
Program Studi: PJJ Informatika
Fakultas: Sistem Informasi
Universitas: Siber Asia
```

Hermes Agent seharusnya bisa extract semua data dan proses otomatis! 🚀

---

**Last Updated:** 2026-05-20
