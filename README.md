# 🎓 Auto-Generate PPT Wisuda dari Excel

Script VBA sakti untuk mengotomatisasi pembuatan slide PowerPoint (kartu wisudawan) mengambil data langsung dari Excel, lengkap dengan fitur *download* foto otomatis dari Google Drive! 

## 🚀 Fitur Utama
- Tarik data teks (Nama, NIM, Judul Skripsi) dari Excel ke PPT menembus *Group Shape*.
- Tembus proteksi link Google Drive (menggunakan WinHTTP & Thumbnail API anti-blokir).
- *Auto-Resize* pasfoto secara presisi tanpa menumpuk di atas foto lama.

## 🛠️ Cara Penggunaan
1. Siapkan file Excel berisi data mahasiswa dan link foto Google Drive.
2. Pastikan link foto Google Drive sudah di-setting **"Anyone with the link / Siapa saja yang memiliki tautan"**.
3. Siapkan file template PowerPoint. Kosongkan area fotonya di Slide 1, dan pastikan *placeholder* teksnya sesuai (contoh: `[NAMA]`, `[NIM]`).
4. Buka VBA di Excel (`ALT + F11`), *paste* kode dari file `GeneratorPPT_Wisuda.vba`.
5. **PENTING:** Matikan fitur *Preview Pane* di folder Windows Explorer agar script tidak membentur PowerPoint bayangan.
6. Buka file PPT dan Excel secara bersamaan, lalu klik *Run* / *Play* di VBA Excel.
7. Tunggu script bekerja dan selesai! Jika ada link Drive yang di-private, akan muncul kotak peringatan merah di area foto.

---
*Dibuat untuk mempermudah urusan IT Kampus biar bisa cepat ngopi wkwk.*
