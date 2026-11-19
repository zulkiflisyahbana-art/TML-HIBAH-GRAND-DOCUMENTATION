# TML HIBAH – GRAND DOCUMENTATION

Website mandiri (HTML + CSS + JavaScript, tanpa server) milik **PT Tri Mandiri Logistik** untuk menyiapkan
**Commercial Invoice** dan **Packing List** pengiriman barang hibah dalam format Excel resmi TML.

Target alamat: `https://zulkiflisyahbana-art.github.io/TML-HIBAH-Grand-Documentation/`

---

## 1. Isi paket

```
TML-HIBAH-Grand-Documentation/
├── index.html              ← halaman utama
├── style.css               ← tampilan (mengikuti TML Order Hub: navy, oranye, krem; responsif)
├── script.js               ← seluruh logika aplikasi
├── README.md               ← panduan ini
├── .nojekyll               ← (opsional) agar GitHub tidak memproses ulang dengan Jekyll
├── assets/
│   ├── logo.png            ← logo PT Tri Mandiri Logistik (dipotong bulat dari file logo)
│   ├── favicon.png
│   ├── templates-data.js   ← template Excel resmi TML (tertanam, base64)
│   └── vendor/jszip.min.js ← library JSZip 3.10.1 (MIT), disalin lokal, tanpa CDN
├── templates/
│   ├── TML_COMMERCIAL_INVOICE_TEMPLATE.xlsx   ← salinan template (rujukan/arsip)
│   └── TML_PACKING_LIST_TEMPLATE.xlsx
└── optional-backend/
    └── Code.gs             ← OPSIONAL: penyimpanan bersama via Google Apps Script
```

Semua jalur memakai **path relatif**, jadi aman di alamat berawalan nama repository (`/TML-HIBAH-Grand-Documentation/`).

## 2. Fitur: apa yang jalan, apa yang tidak

### Berjalan penuh di browser (sudah diuji)
| Fitur | Keterangan |
|---|---|
| Tampilan 3 tab (Isi Data, Verifikasi, Excel) | Gaya mengikuti TML Order Hub: header navy dengan logo TML, tab bergaris oranye, kartu putih di latar krem, tombol utama oranye. |
| Commercial Invoice dan Packing List | Mengisi template resmi TML (stempel, tanda tangan, merged cells, area cetak tetap). Jumlah unit bebas (1 sampai banyak); 3–5 unit otomatis 1 halaman. |
| Unduh Excel | Langsung dari browser, siap dibuka dan dicetak di Excel. |
| Droplist semua kolom | Klik ▾ untuk memilih data yang pernah dipakai, atau ketik baru. |
| Unit: tambah, edit, hapus | Tombol **+ Unit**, **Edit**, **Hapus** (dua langkah). |
| Pemeriksaan otomatis | VIN 17 karakter valid (tanpa I, O, Q), VIN kembar, total berat vs B/L, jumlah unit vs B/L, harga terisi. |
| Dokumen pembanding | Ketik ulang chassis B/L dan data MoU (chassis + berat) agar dicocokkan dengan form. |
| Laporan verifikasi | Tabel No / Informasi / Data / Status / Catatan dengan status VERIFIED, MISSING, CONFLICT, NEED CONFIRMATION. Bisa diunduh (.html) dan dicetak. |
| Impor dari Excel | Membaca kembali file CI/PL **buatan aplikasi ini** untuk melanjutkan atau mengoreksi. |
| Simpan proyek | Di browser perangkat ini (localStorage), plus ekspor dan impor `.json` untuk berbagi antar perangkat. |

### Tidak termasuk (jujur)
| Fitur | Alasan |
|---|---|
| Unggah PDF lalu **ekstraksi otomatis** isi B/L, MoU, NPWP | Dokumen tersebut berupa hasil scan (gambar). Membacanya butuh OCR atau AI. Itu tetap dikerjakan di chat Claude Project; hasilnya diketik atau diimpor ke aplikasi ini. |
| Pembacaan gambar dokumen oleh Claude | Hanya tersedia di dalam tampilan claude.ai. Di GitHub Pages butuh API berbayar dan kunci rahasia yang **tidak boleh** ditanam di JavaScript publik. |
| Pencocokan otomatis ke file dokumen asli | Pencocokan dilakukan terhadap data yang diketik pengguna, bukan terhadap file. |
| Nomor mesin dan dimensi di Excel | Template tidak punya kolomnya. Keduanya hanya muncul di laporan verifikasi (opsional). |

### Opsional, belum diuji
- **Penyimpanan bersama tim** lewat Google Apps Script (`optional-backend/Code.gs`), memakai pola JSONP seperti TML Order Hub. Kodenya ada dan terintegrasi, tetapi **belum diuji terhadap Google**. Uji dulu dengan 1 proyek sebelum dipakai tim.

## 3. Mencoba di komputer sendiri
Klik dua kali `index.html` (berjalan juga tanpa internet), atau jalankan `python -m http.server 8000` di folder ini lalu buka `http://localhost:8000`.

## 4. Deploy ke GitHub Pages (tanpa instal apa pun)

> Akun: `zulkiflisyahbana-art` · Repository: `TML-HIBAH-Grand-Documentation` (huruf persis seperti ini)

**A. Buat repository**
1. Login ke https://github.com sebagai `zulkiflisyahbana-art`.
2. Klik **+** (kanan atas) → **New repository**.
3. **Repository name:** `TML-HIBAH-Grand-Documentation`
4. Pilih **Public** (GitHub Pages gratis memerlukan repository publik).
5. **Jangan** centang "Add a README" (kita sudah punya).
6. Klik **Create repository**.

**B. Unggah file**
1. Ekstrak file ZIP di komputer. Buka foldernya hingga terlihat `index.html`, `style.css`, `script.js`, `assets`, dan seterusnya.
2. Di halaman repository baru, klik **uploading an existing file**.
3. Seret **seluruh isi** folder (semua file **dan** folder `assets`, `templates`, `optional-backend`) ke jendela GitHub. Jangan menyeret folder induknya, supaya `index.html` berada di akar repository.
4. Tunggu proses unggah selesai, lalu klik **Commit changes**.
5. Pastikan di halaman utama repository terlihat `index.html`, `style.css`, `script.js` di daftar paling atas (bukan di dalam subfolder).

**C. Aktifkan GitHub Pages**
1. Buka tab **Settings** → menu **Pages** (kolom kiri).
2. **Build and deployment → Source:** pilih **Deploy from a branch**.
3. **Branch:** pilih `main` dan folder `/ (root)`, klik **Save**.
4. Tunggu 1–3 menit. Buka tab **Actions** dan tunggu "pages build and deployment" berstatus centang hijau.
5. Kembali ke **Settings → Pages**. Akan muncul tulisan **"Your site is live at …"**.

**D. Pastikan alamat berhasil dibuka**
1. Buka `https://zulkiflisyahbana-art.github.io/TML-HIBAH-Grand-Documentation/` di Google Chrome.
2. Cek berurutan: judul "TML HIBAH – GRAND DOCUMENTATION" tampil → klik **Isi contoh Party 2** → klik **Unduh Commercial Invoice** → buka file di Excel dan pastikan stempel serta tanda tangan muncul.
3. Buka di HP; tampilan harus menyesuaikan layar.

**Jika bermasalah**
| Gejala | Penyebab dan solusi |
|---|---|
| Halaman 404 | Tunggu beberapa menit; pastikan Pages aktif dan `index.html` ada di akar repository; pastikan nama repository sama persis. |
| Pages tidak bisa diaktifkan | Repository masih Private. Ubah ke Public di **Settings → General → Danger Zone**. |
| Tampilan polos atau tombol tidak jalan | Muat ulang paksa (Ctrl+F5). Pastikan folder `assets` ikut terunggah lengkap. |
| Unduhan tidak muncul | Cek pengaturan unduhan Chrome; coba tab baru tanpa ekstensi pemblokir. |

**Memperbarui aplikasi:** di repository klik **Add file → Upload files**, unggah file yang berubah (nama sama menimpa file lama), lalu **Commit**. Situs diperbarui otomatis dalam 1–3 menit.

**Cara baris perintah (opsional):**
```bash
git init && git add . && git commit -m "TML HIBAH Grand Documentation"
git branch -M main
git remote add origin https://github.com/zulkiflisyahbana-art/TML-HIBAH-Grand-Documentation.git
git push -u origin main
```

## 5. Penyimpanan bersama tim (opsional, belum diuji)
1. Buat Google Sheet baru. **Extensions → Apps Script**.
2. Tempel isi `optional-backend/Code.gs`, simpan.
3. **Deploy → New deployment → Web app**. *Execute as:* Me. *Who has access:* Anyone. Salin URL `/exec`.
4. Di aplikasi klik **⚙ Pengaturan**, tempel URL, **Simpan pengaturan**.
5. Simpan satu proyek uji, lalu buka dari perangkat lain.

Catatan: data dikirim lewat alamat URL, jadi proyek dengan sangat banyak unit bisa melewati batas panjang URL. "Anyone" berarti siapa pun yang tahu URL itu bisa membaca dan menulis data. Jangan dibagikan di luar tim.

## 6. Keamanan dan privasi (baca sebelum meng-publish)
- **Repository publik berarti isinya terlihat siapa saja.** Template Excel memuat **gambar stempel GO&DO dan tanda tangan DAON Trading**. Siapa pun yang membuka repository dapat mengunduhnya. Pertimbangkan hal ini sebelum memilih Public, atau minta persetujuan pihak terkait.
- Tidak ada API key, password, atau data rahasia di dalam kode.
- Data proyek yang disimpan lokal ada di browser masing-masing dan tidak dikirim ke mana pun, kecuali Anda mengaktifkan penyimpanan bersama.
- Situs ini bukan halaman login; siapa pun yang tahu alamatnya dapat membukanya.

## 7. Mengganti template resmi
Ganti file di `templates/`, lalu buat ulang `assets/templates-data.js`:
```bash
python3 - <<'PY'
import base64
ci=base64.b64encode(open('templates/TML_COMMERCIAL_INVOICE_TEMPLATE.xlsx','rb').read()).decode()
pl=base64.b64encode(open('templates/TML_PACKING_LIST_TEMPLATE.xlsx','rb').read()).decode()
open('assets/templates-data.js','w').write('window.TML_TEMPLATES={CI:"%s",PL:"%s"};\n'%(ci,pl))
PY
```
Struktur template (sel isian, 2 blok barang baris 27–40, total baris 41) harus tetap sama seperti Party 2 of 4, karena mesin pengisi mengacu ke posisi sel tersebut.

## 8. Lisensi pihak ketiga
JSZip 3.10.1 — MIT/GPLv3 (https://stuk.github.io/jszip/).
