# Quickstart: Cara Menjalankan Project

Ikuti langkah-langkah di bawah ini untuk memastikan project berjalan dengan benar di lingkungan lokal Anda.

## 1. Persiapan Lingkungan (Setup)

Buka terminal di folder project (`D:\abdul\certificate-generator`) dan jalankan perintah berikut secara berurutan untuk memperbaiki *path* yang rusak:

```powershell
# 1. Pastikan Anda tidak sedang menjalankan aplikasi (tekan Ctrl+C jika masih berjalan)

# 2. Hapus environment lama yang rusak
Remove-Item -Recurse -Force .\venv

# 3. Buat environment baru (akan otomatis menggunakan path D:\...)
python -m venv venv

# 4. Aktifkan environment
.\venv\Scripts\Activate.ps1

# 5. Instal dependensi
python -m pip install -r requirements.txt
```

## 2. Menjalankan Aplikasi

Setelah langkah setup selesai, selalu gunakan cara ini untuk menjalankan server:

```powershell
# Jalankan menggunakan modul python untuk menghindari error path launcher
python -m uvicorn main:app --reload
```

Server akan berjalan di: **http://127.0.0.1:8000**

## 3. Catatan Penting
- **Microsoft Word:** Project ini menggunakan `docx2pdf` untuk konversi dokumen. **Aplikasi Microsoft Word harus terinstal** di komputer Anda agar proses konversi berjalan.
- **Folder Upload:** Jika folder `uploads` atau `output` terhapus, aplikasi akan membuatnya secara otomatis saat pertama kali ada *request* generate dokumen.
