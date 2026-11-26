# TML HIBAH – GRAND DOCUMENTATION · MODE TERBUKA
Aplikasi web internal **PT Tri Mandiri Logistik** untuk menyiapkan **Commercial Invoice** dan **Packing List** pengiriman hibah dari dokumen sumber (B/L, MoU, NPWP), dengan verifikasi, konfirmasi, dan arsip.

Alamat: `https://zulkiflisyahbana-art.github.io/TML-HIBAH-Grand-Documentation/`

> ⚠ **MODE TERBUKA.** Template resmi, **stempel GO&DO, tanda tangan**, dan data contoh berada di `templates-data.js` tanpa enkripsi. Karena GitHub Pages gratis memerlukan repository **Public**, **siapa pun yang tahu alamatnya dapat membuka aplikasi dan mengunduh gambar stempel**. Untuk menjaga kerahasiaan, gunakan paket **MODE RAHASIA** (terenkripsi, dibuka dengan kode akses).

---
## 1. Alur kerja (3 tahap) dan penanda status
| Tab | Isi |
|---|---|
| **1. Isi Data** | **Unggah dokumen sumber** (penanda B/L, MoU, NPWP) → isi data proyek, shipper/consignee, unit (▾ untuk memilih data lama) → **💾 Simpan & Lanjut ke Verifikasi**. |
| **2. Verifikasi** | Laporan verifikasi langsung tampil (baris berwarna menurut status), tombol **Edit** per baris, jejak aktivitas. Jika sesuai: **💾 Simpan** → **✅ Telah Dikonfirmasi**. |
| **3. Excel** | Pratinjau CI/PL (dengan stempel dan tanda tangan), **Unduh .xlsx**, **Cetak**, **Ekspor PDF**, dan **Arsip proyek**. Mengubah data setelah konfirmasi mengunci tab ini lagi. |

**Penanda status (agar jelas mana yang sudah dikerjakan):**
- Tombol berputar saat diproses, lalu **berubah hijau dengan centang dan jam** (✓ Tersimpan 10.24, ✓ Terkonfirmasi, ✓ CI terunduh, dan seterusnya).
- Tiap tab punya lencana: **●** oranye = ada perubahan belum disimpan, **✓** hijau = selesai, **○** = belum, **🔒** = terkunci sampai dikonfirmasi.
- Setiap perubahan data otomatis **membatalkan konfirmasi** dan mengembalikan tombol ke warna semula.
- Kartu unit menampilkan **✓ Lengkap / ⚠ Belum lengkap / ✎ Sedang diedit**; tombol **Edit** pada laporan menyorot kolom yang dituju.

### Baca dokumen otomatis (tab 1)
Setelah B/L, MoU, dan/atau NPWP diunggah, klik **🔍 Baca dokumen otomatis** (±30 detik). Aplikasi membaca dokumen **di browser Anda** (tidak ada yang dikirim ke server), menampilkan **hasil bacaan beserta sumbernya** untuk ditinjau, lalu **✔ Isi ke Data Proyek & Unit** mengisi form (bisa diedit). Kolom **kuning** = dari dokumen, **biru** = dari template contoh (cylinder/engine/color tidak ada di dokumen). Laporan verifikasi memuat kolom **Sumber dokumen** dan pemeriksaan silang (NPWP B/L ↔ kartu NPWP, chassis B/L ↔ MoU).

| Dokumen | Cara baca | Keandalan |
|---|---|---|
| **B/L** (PDF dengan teks) | Teks asli PDF | Tinggi. Diuji pada format Heung-A Line; format pelayaran lain perlu dicek. |
| **MoU** (scan) | OCR | Sedang: nomor, tanggal, chassis, merk, tahun, jarak tempuh terbaca; total unit dan total nilai terbaca. |
| **Tabel lampiran MoU** (harga, berat, dimensi) | OCR | **Rendah pada scan kecil.** Angka hanya diisi otomatis bila lolos cek jumlah total; jika tidak, dibiarkan kosong dan diisi manual sambil membuka MoU (tombol Lihat). |
| **Kartu NPWP** | OCR | Sedang: nomor dan nama terbaca. |

Aturan keselamatan: nomor chassis hanya dikoreksi bila selisih **1 karakter** dan tidak menimpa kecocokan persis. Berat per unit = total B/L ÷ jumlah unit (aturan TML). Departure date dan vessel tidak diisi otomatis (keputusan TML sebelumnya) kecuali dicentang. **Selalu periksa hasil dengan dokumen asli.**

## 2. Isi paket (semua berkas sejajar dengan index.html)
```
index.html  style.css  script.js  preview.js  baca.js
baca/                 ← pembaca PDF + OCR (7 berkas, ±12 MB). UNGGAH FOLDER INI JUGA
templates-data.js     ← template resmi + stempel + data contoh (TIDAK terenkripsi)
vault.enc.js          ← kosong pada mode ini
buat-vault.html       ← alat membuat vault terenkripsi (untuk beralih ke MODE RAHASIA)
jszip.min.js  html2canvas.min.js  jspdf.umd.min.js
logo.png  favicon.png  README.md  .nojekyll
optional-backend/Code.gs
```

## 3. Beralih ke mode rahasia nanti
Unggah `vault.enc.js` dan `templates-data.js` dari paket **MODE RAHASIA** (menimpa keduanya). Layar kunci akan muncul. **Catatan:** stempel yang pernah dipublikasikan tanpa enkripsi tetap ada di riwayat repository; untuk menghapusnya hapus repository lalu buat ulang dengan nama yang sama.

## 4. Deploy ke GitHub Pages
1. **Menimpa versi yang sedang berjalan:** buka repository → **Add file → Upload files** → ekstrak ZIP, **Ctrl+A** lalu seret **semua isi folder** (bukan folder induknya) → **Commit changes**. File bernama sama otomatis menimpa.
2. Pastikan semua berkas di bagian 2 terlihat di daftar paling atas repository (bukan dalam subfolder). Folder `assets` dan `templates` dari versi lama tidak dipakai lagi dan boleh dihapus.
3. Tunggu 1–3 menit (tab **Actions** hijau), buka alamat website, tekan **Ctrl+F5**.
4. Jika ada berkas yang kurang, website menampilkan kotak peringatan merah yang menyebut namanya. Folder **`baca`** harus ikut terunggah (berisi 7 berkas); tanpa folder itu aplikasi tetap jalan, hanya tombol Baca dokumen otomatis yang tidak aktif.
5. Pembaca dokumen butuh alamat https (GitHub Pages) atau `localhost`; tidak jalan jika `index.html` dibuka langsung dari folder komputer.

## 5. Arsip proyek ("folder khusus TML HIBAH")
Struktur: `TML HIBAH / tahun / NAMA_PROYEK /` berisi `01_Dokumen_Sumber`, `02_Excel`, `03_PDF`, `04_Laporan_Verifikasi`, `data-proyek.json`, `jejak-aktivitas.txt`, `00_RINGKASAN.txt`.
- **📦 Unduh Paket Arsip (.zip):** semua perangkat.
- **💾 Simpan ke Folder TML HIBAH:** tulis langsung ke folder di komputer (Chrome/Edge di PC/laptop). Pilih folder sekali lewat **Pilih folder…**.
- **☁ Kirim ke Google Drive (eksperimental):** muncul bila URL Apps Script diisi. **Belum diuji ke Google.**
- Lampiran dokumen sumber tersimpan di browser yang mengunggah; paket arsip menyertakannya.

## 6. Penyimpanan bersama dan arsip Drive (Apps Script)
1. Google Sheet → **Extensions → Apps Script** → tempel `optional-backend/Code.gs` → simpan.
2. **Project Settings → Script properties** → tambah `ACCESS_CODE` = kode rahasia.
3. **Deploy → Manage deployments → ✏ → Version: New version → Deploy**. Setujui izin Google Drive.
4. Di website: **⚙ Pengaturan** → URL `/exec` + **Kode akses** yang sama → **Simpan pengaturan**.
Tanpa `ACCESS_CODE`, siapa pun yang tahu URL bisa menulis ke Sheet/Drive Anda.

## 7. Fitur yang TIDAK ada
- Pembacaan yang selalu akurat: tabel lampiran MoU pada scan kecil tidak terbaca andal (harga, berat, dimensi diisi manual). OCR dapat salah baca; hasil selalu ditinjau operator.
- Login per orang dan peran (Operator/Approver). Perlu backend berotentikasi, bukan hosting statis gratis.
- Nama operator di jejak aktivitas diketik sendiri (belum terotentikasi).

## 8. Lisensi pihak ketiga
JSZip 3.10.1 (MIT/GPLv3), html2canvas 1.4.1 (MIT), jsPDF 2.5.1 (MIT).
