/**
 * OPSIONAL – penyimpanan proyek bersama untuk TML HIBAH (Google Apps Script + Google Sheets).
 * Pola JSONP sama seperti TML Order Hub / Finance Hub. BELUM DIUJI oleh pembuat paket ini.
 * Cara pakai: lihat README.md bagian "Penyimpanan bersama tim".
 */
const SHEET_NAME = 'Projects'; // kolom: id | name | updated | data

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
function doGet(e) {
  const p = e.parameter, cb = p.callback;
  try {
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
