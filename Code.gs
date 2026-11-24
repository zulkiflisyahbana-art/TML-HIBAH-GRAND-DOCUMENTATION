/**
 * TML HIBAH – penyimpanan bersama + arsip Google Drive (Google Apps Script).
 *
 * STATUS: penyimpanan proyek (list/save/delete) sudah berhasil dipakai pengguna.
 *         Arsip Drive (putfile / archive_check) dan kode akses BELUM diuji ke Google oleh pembuat paket ini.
 *
 * LANGKAH PERBARUI (wajib setelah mengganti kode):
 *  1. Tempel seluruh isi file ini ke Apps Script, simpan.
 *  2. Project Settings > Script properties > tambah ACCESS_CODE = (kode rahasia pilihan Anda).
 *     Isi kode yang sama di website: ⚙ Pengaturan > Kode akses. Tanpa ACCESS_CODE, siapa pun yang tahu URL dapat menulis.
 *  3. Deploy > Manage deployments > ikon pensil > Version: New version > Deploy (URL tetap sama).
 *  4. Saat diminta, setujui izin akses Google Drive.
 */
const SHEET_NAME = 'Projects';   // kolom: id | name | updated | data
const ROOT_NAME  = 'TML HIBAH';  // folder arsip di Google Drive pemilik skrip

function codeOk_(p) {
  const c = PropertiesService.getScriptProperties().getProperty('ACCESS_CODE');
  return !c || String(p.code || '') === c;
}
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(['id', 'name', 'updated', 'data']); }
  return sh;
}
function out_(cb, obj) {
  const body = (cb || 'callback') + '(' + JSON.stringify(obj) + ')';
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JAVASCRIPT);
}
function rootFolder_() {
  const it = DriveApp.getFoldersByName(ROOT_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(ROOT_NAME);
}
function folderPath_(parts, create) {
  let f = rootFolder_();
  for (const n of parts) {
    const it = f.getFoldersByName(n);
    if (it.hasNext()) f = it.next();
    else if (create) f = f.createFolder(n);
    else return null;
  }
  return f;
}
function countFiles_(f) {
  let n = 0;
  const fs = f.getFiles(); while (fs.hasNext()) { fs.next(); n++; }
  const ds = f.getFolders(); while (ds.hasNext()) n += countFiles_(ds.next());
  return n;
}
function cleanParts_(path) {
  const parts = String(path).split('/').filter(Boolean);
  if (parts[0] === ROOT_NAME) parts.shift();
  return parts;
}

function doGet(e) {
  const p = e.parameter, cb = p.callback;
  try {
    if (!codeOk_(p)) return out_(cb, { ok: false, error: 'kode akses salah' });
    if (p.action === 'archive_check') {
      const f = folderPath_(cleanParts_(p.path), false);
      return out_(cb, f ? { ok: true, count: countFiles_(f), url: f.getUrl() } : { ok: false, error: 'folder belum ada' });
    }
    const sh = sheet_(), rows = sh.getDataRange().getValues().slice(1);
    if (p.action === 'list') {
      // updated disimpan sebagai teks agar tidak terkena konversi angka Sheets
      return out_(cb, { ok: true, items: rows.filter(r => r[0]).map(r => ({ id: String(r[0]), name: String(r[1]), updated: String(r[2]), data: String(r[3]) })) });
    }
    const idx = rows.findIndex(r => String(r[0]) === p.id);
    if (p.action === 'save') {
      const rec = [p.id, p.name, String(p.updated), p.data];
      if (idx >= 0) sh.getRange(idx + 2, 1, 1, 4).setValues([rec]); else sh.appendRow(rec);
      return out_(cb, { ok: true });
    }
    if (p.action === 'delete') {
      if (idx >= 0) sh.deleteRow(idx + 2);
      return out_(cb, { ok: true });
    }
    return out_(cb, { ok: false, error: 'aksi tidak dikenal' });
  } catch (err) { return out_(cb, { ok: false, error: String(err) }); }
}

// Menerima satu berkas arsip per permintaan (POST teks JSON dari website).
function doPost(e) {
  try {
    const p = JSON.parse(e.postData.contents);
    if (!codeOk_(p)) return ContentService.createTextOutput('denied');
    if (p.action !== 'putfile') return ContentService.createTextOutput('unknown');
    const parts = cleanParts_(p.path), name = parts.pop(), f = folderPath_(parts, true);
    const old = f.getFilesByName(name); while (old.hasNext()) old.next().setTrashed(true); // timpa versi lama
    f.createFile(Utilities.newBlob(Utilities.base64Decode(p.b64), p.mime || 'application/octet-stream', name));
    return ContentService.createTextOutput('ok');
  } catch (err) { return ContentService.createTextOutput('error: ' + err); }
}
