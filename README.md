# TML HIBAH – GRAND DOCUMENTATION

Website mandiri (HTML + CSS + JavaScript, tanpa server) milik **PT Tri Mandiri Logistik** untuk menyiapkan
**Commercial Invoice** dan **Packing List** pengiriman barang hibah dalam format Excel resmi TML.

Alamat: `https://zulkiflisyahbana-art.github.io/TML-HIBAH-Grand-Documentation/`

---

## 1. Alur kerja di website (3 tahap)

| Tab | Yang dilakukan |
|---|---|
| **1. Isi Data** | **Mulai dengan Unggah dokumen** B/L, NPWP, MoU, dan lainnya (penanda centang menunjukkan dokumen wajib yang sudah terlampir). Lalu isi data proyek, shipper/consignee, dan unit kendaraan sambil membuka dokumen lewat tombol Lihat (klik ▾ untuk memilih data lama, atau ketik baru). Isian tersimpan otomatis sebagai draf. Klik **💾 Simpan & Lanjut ke Verifikasi**. |
| **2. Verifikasi** | Laporan verifikasi langsung tampil (VERIFIED / MISSING / CONFLICT / NEED CONFIRMATION). Klik **Edit** pada baris untuk memperbaiki datanya, atau **✏️ Edit Data**. Dokumen lampiran bisa dibuka lewat **Lihat**. Jika sudah sesuai, klik **💾 Simpan**, lalu **✅ Telah Dikonfirmasi**. |
| **3. Excel** | Terbuka setelah dikonfirmasi. **Pratinjau** Commercial Invoice dan Packing List, lalu **Unduh Excel (.xlsx)**, **Cetak**, **Ekspor PDF**, atau **Simpan**. Mengubah data setelah konfirmasi akan mengunci tab ini lagi sampai dikonfirmasi ulang. |

## 2. Isi paket (semua berkas pendukung sejajar dengan index.html)

```
index.html            ← halaman utama
style.css             ← tampilan (mengikuti TML Order Hub)
script.js             ← logika aplikasi
preview.js            ← pratinjau Excel
templates-data.js     ← template Excel resmi TML (tertanam)
jszip.min.js          ← library JSZip 3.10.1
html2canvas.min.js    ← library html2canvas 1.4.1 (untuk PDF)
jspdf.umd.min.js      ← library jsPDF 2.5.1 (untuk PDF)
logo.png, favicon.png ← logo PT Tri Mandiri Logistik
README.md, .nojekyll
templates/            ← salinan template .xlsx (arsip, tidak dipakai saat berjalan)
optional-backend/     ← Code.gs untuk penyimpanan bersama (opsional)
```

**Penting:** 11 berkas di baris atas harus berada **di level yang sama dengan `index.html`**, bukan di dalam folder. Jika ada yang tidak terunggah, website menampilkan **kotak peringatan merah** yang menyebut nama berkas yang kurang.

## 3. Fitur: apa yang jalan, apa yang tidak

### Berjalan penuh di browser (sudah diuji)
- Pengisian template resmi TML (stempel, tanda tangan, merged cells, area cetak tetap), 1 sampai banyak unit.
- Pratinjau, unduh `.xlsx`, cetak, dan ekspor PDF (CI di kertas Legal, PL di A4, sesuai template).
- Droplist semua kolom, tambah/edit/hapus unit, laporan verifikasi dengan tombol Edit, konfirmasi. Data MoU (chassis, berat, harga) dicocokkan per unit; harga yang sama pada beberapa unit diberi peringatan.
- Unggah dokumen sebagai lampiran untuk dilihat (PDF, gambar) dan impor Excel CI/PL buatan aplikasi.
- Draf otomatis, simpan proyek di perangkat, ekspor/impor `.json`.

### Tidak termasuk (jujur)
| Fitur | Alasan |
|---|---|
| **Membaca isi PDF/gambar otomatis** (ekstraksi B/L, MoU, NPWP) | Dokumen berupa scan; butuh OCR atau AI. Itu tetap dikerjakan di chat Claude Project; hasilnya diketik atau diimpor. Lampiran hanya disimpan dan ditampilkan untuk dicocokkan secara visual. |
| Fitur Claude/AI di website | Butuh API berbayar dan kunci rahasia yang tidak boleh ada di JavaScript publik. |
| Pencocokan otomatis ke file dokumen | Pencocokan dilakukan terhadap data yang diketik pengguna. |
| Nomor mesin dan dimensi di Excel | Template tidak punya kolomnya; keduanya hanya muncul di laporan. |

### Catatan teknis
- **Lampiran dokumen** disimpan di browser perangkat itu (IndexedDB), tidak ikut penyimpanan online dan tidak ikut ekspor `.json`.
- **Pratinjau dan PDF** adalah gambar ulang template; sangat mirip, tetapi hasil cetak resmi tetap dari file `.xlsx` di Excel. PDF berupa gambar (bukan teks yang bisa dipilih). Untuk PDF teks, klik **Cetak** lalu pilih *Save as PDF*.
- **Penyimpanan bersama tim** (Google Apps Script, folder `optional-backend`) sudah berhasil dipakai menyimpan proyek pada percobaan pengguna. Uji pembukaan dari perangkat lain sebelum dipakai tim.

## 4. Deploy ke GitHub Pages (tanpa instal apa pun)

> Akun: `zulkiflisyahbana-art` · Repository: `TML-HIBAH-Grand-Documentation` · harus **Public**

**A. Buat repository** (lewati jika sudah ada): GitHub → **+** → **New repository** → nama `TML-HIBAH-Grand-Documentation` → **Public** → jangan centang README → **Create repository**.

**B. Unggah (atau perbarui) file**
1. Ekstrak ZIP. Buka foldernya sampai terlihat `index.html`.
2. Di repository: **Add file → Upload files**.
3. Tekan **Ctrl+A** di folder hasil ekstrak, lalu seret **semua isinya** ke GitHub (termasuk folder `templates` dan `optional-backend`).
4. Tunggu semua file selesai, lalu **Commit changes**. File bernama sama otomatis menimpa yang lama.
5. Buka halaman utama repository dan pastikan **11 berkas di atas terlihat langsung di daftar paling atas**.
6. Folder `assets` dari versi lama sudah tidak dipakai. Boleh dibiarkan atau dihapus (buka folder → **⋯ → Delete directory**).

**C. Aktifkan Pages** (sekali saja): **Settings → Pages → Source: Deploy from a branch → Branch `main` / `(root)` → Save**. Tunggu 1–3 menit; di tab **Actions** statusnya harus centang hijau.

**D. Pastikan berhasil**
1. Buka alamat website, lalu tekan **Ctrl+F5**.
2. Logo TML tampil di header dan **tidak ada kotak peringatan merah**.
3. Klik **Isi contoh Party 2** → **Simpan & Lanjut** → tabel verifikasi tampil → **Telah Dikonfirmasi** (klik dua kali bila ada CONFLICT) → pratinjau Excel tampil → coba **Unduh Excel**, **Ekspor PDF**, dan **Cetak**.

**Jika bermasalah**
| Gejala | Solusi |
|---|---|
| Kotak merah "Berkas pendukung belum terunggah" | Unggah berkas yang disebut ke akar repository (sejajar `index.html`), tunggu 1–3 menit, Ctrl+F5. |
| Tampilan lama masih muncul | Ctrl+F5; GitHub Pages bisa menyimpan cache beberapa menit. |
| Halaman 404 | Pages belum aktif, `index.html` tidak di akar repository, atau nama repository tidak sama. |
| Pages tidak bisa diaktifkan | Repository masih Private; ubah ke Public. |
| Unduhan/PDF tidak muncul | Cek izin unduhan di Chrome; coba tab tanpa ekstensi pemblokir. |

## 5. Penyimpanan bersama tim (opsional)
1. Buat Google Sheet → **Extensions → Apps Script** → tempel isi `optional-backend/Code.gs` → simpan.
2. **Deploy → New deployment → Web app** → *Execute as:* Me, *Who has access:* Anyone → salin URL `/exec`.
3. Di website klik ⚙, tempel URL, **Simpan pengaturan** (status berubah menjadi "Tersimpan online").

"Anyone" berarti siapa pun yang tahu URL bisa membaca dan menulis data proyek; jangan dibagikan di luar tim. Proyek dengan sangat banyak unit bisa melewati batas panjang URL.

## 6. Keamanan dan privasi
- **Repository publik terlihat siapa saja.** Template memuat **stempel GO&DO dan tanda tangan DAON Trading**; siapa pun yang membuka repository dapat mengunduhnya. Pertimbangkan sebelum memilih Public.
- Tidak ada API key atau password di dalam kode. Data proyek lokal hanya ada di browser masing-masing kecuali penyimpanan bersama diaktifkan.
- Situs ini bukan halaman login.

## 7. Mengganti template resmi
Ganti file di `templates/`, lalu buat ulang `templates-data.js`:
```bash
python3 - <<'PY'
import base64
ci=base64.b64encode(open('templates/TML_COMMERCIAL_INVOICE_TEMPLATE.xlsx','rb').read()).decode()
pl=base64.b64encode(open('templates/TML_PACKING_LIST_TEMPLATE.xlsx','rb').read()).decode()
open('templates-data.js','w').write('window.TML_TEMPLATES={CI:"%s",PL:"%s"};\n'%(ci,pl))
PY
```
Struktur template (sel isian, 2 blok barang baris 27–40, total baris 41) harus tetap seperti Party 2 of 4.

## 8. Lisensi pihak ketiga
JSZip 3.10.1 (MIT/GPLv3), html2canvas 1.4.1 (MIT), jsPDF 2.5.1 (MIT).
