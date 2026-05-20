# User-Provided References Guide

MakalahMCP sekarang bisa menerima referensi dari user langsung (skip auto-research)!

## ✅ **Format Referensi yang Didukung:**

### **Format 1: List URL**
```
Buat makalah dengan judul "Implementasi AI di sistem HRIS"
Nama: Egi Aprilianto NIM: 12973702883
Program Studi: PJJ Informatika Fakultas: Sistem Informasi
Universitas: Siber Asia

Referensi:
- https://journal.example.com/ai-hris-2024
- https://research.example.com/hr-automation
- https://tech.example.com/ai-implementation
```

### **Format 2: Judul + URL**
```
Referensi:
1. "AI in Human Resource Management" (https://journal.example.com/ai-hris)
2. "Automation in HR Systems" (https://research.example.com/hr-auto)
3. "Implementation Guide for AI HRIS" (https://tech.example.com/guide)
```

### **Format 3: Text/Kutipan**
```
Referensi:
Menurut penelitian Smith (2024), implementasi AI dalam sistem HRIS dapat meningkatkan efisiensi hingga 40%. Sistem AI dapat mengotomasi proses rekrutmen, evaluasi kinerja, dan manajemen data karyawan.

Sumber: Smith, J. (2024). "AI in HRIS". Journal of HR Technology.
https://journal.example.com/smith-2024
```

### **Format 4: Mixed (URL + Text)**
```
Referensi:
1. https://journal.example.com/ai-hris
   - Membahas tentang implementasi AI
   - Studi kasus di 50 perusahaan
   
2. https://research.example.com/hr-automation
   - Fokus pada automation
   - ROI analysis
```

---

## 🤖 **Apa yang Akan Terjadi:**

### **Dengan Referensi User:**

**User Input:**
```
Buat makalah "Implementasi AI di sistem HRIS"
Nama: Egi Aprilianto NIM: 12973702883
Program Studi: PJJ Informatika Fakultas: Sistem Informasi
Universitas: Siber Asia

Referensi:
- https://journal.example.com/ai-hris-2024
- https://research.example.com/hr-automation
- https://tech.example.com/ai-implementation
```

**Agent Response:**
```
Saya akan buat makalah dengan data:
- Judul: Implementasi AI di sistem HRIS
- Nama: Egi Aprilianto
- NIM: 12973702883
- Program Studi: PJJ Informatika
- Fakultas: Sistem Informasi
- Universitas: Siber Asia
- Tahun: 2026

Referensi yang akan digunakan:
1. https://journal.example.com/ai-hris-2024
2. https://research.example.com/hr-automation
3. https://tech.example.com/ai-implementation

Apakah sudah benar?
```

**User:** `ya`

**Agent:**
```
✅ Data dikonfirmasi
🔍 Mencari logo Universitas Siber Asia...
⏭️ Skip auto-research (menggunakan referensi Anda)
📝 Generating konten makalah...

[Preview konten...]

Apakah konten sudah sesuai?
```

---

### **Tanpa Referensi User (Auto-Research):**

**User Input:**
```
Buat makalah "Implementasi AI di sistem HRIS"
Nama: Egi Aprilianto NIM: 12973702883
Program Studi: PJJ Informatika Fakultas: Sistem Informasi
Universitas: Siber Asia
```

**Agent Response:**
```
[Konfirmasi data...]

Apakah sudah benar?
```

**User:** `ya`

**Agent:**
```
✅ Data dikonfirmasi
🔍 Mencari logo Universitas Siber Asia...
🌐 Researching topic dari web... (auto-research)
📝 Generating konten makalah...

[Preview konten...]
```

---

## 🎯 **Keuntungan User-Provided References:**

| Feature | Auto-Research | User References |
|---------|---------------|-----------------|
| **Speed** | ~10-15s | ⚡ ~2-3s (skip research) |
| **Quality** | Good (web scraping) | ⭐ Better (user curated) |
| **Control** | Limited | ⭐ Full control |
| **Specific Sources** | Random from web | ⭐ Exact sources |
| **Academic Sources** | Mixed | ⭐ Can be all academic |

---

## 📝 **Use Cases:**

### **1. Academic Paper dengan Jurnal Spesifik**
```
Referensi:
- IEEE Xplore: https://ieeexplore.ieee.org/document/123456
- ACM Digital Library: https://dl.acm.org/doi/10.1145/123456
- Springer: https://link.springer.com/article/10.1007/123456
```

### **2. Company Internal Research**
```
Referensi:
- Internal Report 2024: https://company.internal/reports/ai-hris
- Case Study: https://company.internal/cases/implementation
- Best Practices: https://company.internal/docs/best-practices
```

### **3. Specific Authors/Papers**
```
Referensi:
1. Smith, J. (2024). "AI in HRIS". https://journal.com/smith-2024
2. Johnson, K. (2023). "HR Automation". https://research.com/johnson-2023
3. Lee, M. (2024). "Implementation Guide". https://tech.com/lee-2024
```

### **4. Mixed Sources (Web + Books + Papers)**
```
Referensi:
- Web: https://techcrunch.com/ai-hris-trends
- Book: "AI for HR" by Smith (2024) - ISBN: 978-1234567890
- Paper: https://arxiv.org/abs/2024.12345
- Blog: https://medium.com/@expert/ai-hris-guide
```

---

## 🔄 **Flow Comparison:**

### **Without User References:**
```
1. Confirm data
2. Search logo
3. 🌐 Auto-research from web (10-15s)
4. Generate content
5. Preview
6. Save
```

### **With User References:**
```
1. Confirm data + references
2. Search logo
3. ⏭️ Skip research (use user refs)
4. Generate content (faster!)
5. Preview
6. Save
```

---

## 🧪 **Testing:**

### **Test 1: With References**
```
Buat makalah "Implementasi AI di sistem HRIS"
Nama: Egi Aprilianto NIM: 12973702883
Program Studi: PJJ Informatika Fakultas: Sistem Informasi
Universitas: Siber Asia

Referensi:
- https://journal.example.com/ai-hris
- https://research.example.com/hr-automation
```

Expected: Agent skip research, use provided URLs

### **Test 2: Without References**
```
Buat makalah "Implementasi AI di sistem HRIS"
Nama: Egi Aprilianto NIM: 12973702883
Program Studi: PJJ Informatika Fakultas: Sistem Informasi
Universitas: Siber Asia
```

Expected: Agent do auto-research from web

---

## 📊 **Reference Format in Daftar Pustaka:**

Agent akan format referensi user menjadi:

```
Daftar Pustaka

Smith, J. 2024. "AI in Human Resource Management". Journal of HR Technology. 
https://journal.example.com/ai-hris. Diakses pada 20 Mei 2026.

Johnson, K. 2023. "Automation in HR Systems". Research Quarterly. 
https://research.example.com/hr-automation. Diakses pada 20 Mei 2026.

Lee, M. 2024. "Implementation Guide for AI HRIS". Tech Review. 
https://tech.example.com/guide. Diakses pada 20 Mei 2026.
```

---

## ✅ **Update Server:**

```bash
cd /root/makalahmcp
./update-server.sh
```

---

## 🎉 **Summary:**

Sekarang user bisa:
- ✅ Kasih referensi sendiri (skip auto-research)
- ✅ Atau biarkan kosong (auto-research dari web)
- ✅ Mix format (URL, text, kutipan)
- ✅ Full control atas sumber referensi
- ✅ Lebih cepat (skip research step)

Perfect untuk academic papers yang butuh jurnal spesifik! 📚

---

**Last Updated:** 2026-05-20
