# Download File ke Local User

Ada 3 cara untuk user download file makalah ke local mereka:

## 🎯 **Cara 1: Download via HTTP API (Recommended)**

### **Setup:**
Server sudah ditambahkan 2 endpoint baru:
- `GET /files` - List semua file yang tersedia
- `GET /download/{filename}` - Download file

### **Usage:**

#### **1. List Available Files:**
```bash
curl http://mcp.asln.dev/makalah/files
```

Response:
```json
{
  "files": [
    {
      "filename": "Makalah - Analisis AI.docx",
      "size": 45678,
      "created": "2026-05-20T08:30:00",
      "download_url": "/download/Makalah - Analisis AI.docx"
    },
    {
      "filename": "Makalah - Analisis AI.pdf",
      "size": 123456,
      "created": "2026-05-20T08:30:00",
      "download_url": "/download/Makalah - Analisis AI.pdf"
    }
  ],
  "count": 2
}
```

#### **2. Download File:**
```bash
# Download DOCX
curl -O http://mcp.asln.dev/makalah/download/Makalah%20-%20Analisis%20AI.docx

# Download PDF
curl -O http://mcp.asln.dev/makalah/download/Makalah%20-%20Analisis%20AI.pdf
```

Atau buka di browser:
```
http://mcp.asln.dev/makalah/download/Makalah%20-%20Analisis%20AI.docx
```

---

## 🎯 **Cara 2: Return Base64 di Response**

Modifikasi `save_makalah` tool untuk return file sebagai base64:

```python
# Di save_makalah tool, tambahkan:
import base64

# After saving files
file_data = {}
for file_path in files:
    with open(file_path, 'rb') as f:
        file_data[os.path.basename(file_path)] = base64.b64encode(f.read()).decode()

return json.dumps({
    "files": files,
    "file_data": file_data  # Base64 encoded files
})
```

User bisa decode base64 dan save ke local.

---

## 🎯 **Cara 3: Shared Volume (Docker)**

Mount shared volume antara server dan user:

```yaml
# docker-compose.yml
services:
  makalahmcp:
    volumes:
      - /shared/makalah:/output
```

User akses via:
- SMB/CIFS share
- NFS mount
- Cloud storage (S3, Google Drive)

---

## 📝 **Cara Pakai di Hermes Agent:**

### **Workflow:**

```
1. User: "Buatkan makalah tentang AI"

2. Hermes Agent:
   - Call research_topic
   - Call generate_makalah prompt
   - Call save_makalah tool
   - Get response: {"files": ["Makalah - AI.docx", "Makalah - AI.pdf"]}

3. Hermes Agent response:
   "Makalah sudah dibuat! Download di:
   - DOCX: http://mcp.asln.dev/makalah/download/Makalah%20-%20AI.docx
   - PDF: http://mcp.asln.dev/makalah/download/Makalah%20-%20AI.pdf"

4. User: Click link untuk download
```

---

## 🌐 **Update Nginx Config:**

Tambahkan di nginx config untuk handle download:

```nginx
server {
    listen 80;
    server_name mcp.asln.dev;

    location /makalah/ {
        proxy_pass http://localhost:8000/;
        proxy_set_header Host $host;
        
        # Increase timeout for large files
        proxy_read_timeout 300s;
        proxy_send_timeout 300s;
        
        # Increase buffer for large files
        proxy_buffering on;
        proxy_buffer_size 4k;
        proxy_buffers 8 4k;
        proxy_busy_buffers_size 8k;
    }
}
```

---

## 🔧 **Update Instructions di MCP:**

Update instructions di `main.py` untuk include download info:

```python
mcp = FastMCP(
    "MakalahMCP",
    instructions="""...

11. Setelah disimpan, berikan link download ke user:
    - DOCX: http://mcp.asln.dev/makalah/download/[filename].docx
    - PDF: http://mcp.asln.dev/makalah/download/[filename].pdf
    
User bisa klik link untuk download file ke local mereka.
""",
)
```

---

## ✅ **Testing:**

```bash
# 1. Generate makalah via Hermes Agent
# 2. Get filename from response
# 3. Test download:

curl -O http://mcp.asln.dev/makalah/download/Makalah%20-%20Test.docx

# 4. Check file downloaded
ls -lh Makalah*.docx
```

---

## 📊 **Pros & Cons:**

| Method | Pros | Cons |
|--------|------|------|
| **HTTP Download** | ✅ Simple<br>✅ Works everywhere<br>✅ Direct download | ⚠️ Files stored on server |
| **Base64 Response** | ✅ No storage needed<br>✅ Immediate | ❌ Large response size<br>❌ Need decode |
| **Shared Volume** | ✅ Real-time access<br>✅ No API needed | ❌ Complex setup<br>❌ Security concerns |

**Recommended:** HTTP Download (Cara 1)

---

Mau saya commit perubahan ini ke repository?
