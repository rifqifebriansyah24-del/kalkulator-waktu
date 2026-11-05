/**
 * Kalkulator Waktu — Google Sheets sync endpoint.
 *
 * Setup:
 * 1. Buat Google Sheet baru (boleh kosong).
 * 2. Buka menu Extensions > Apps Script, hapus isi default, lalu tempel
 *    seluruh isi file ini.
 * 3. Di Apps Script editor: klik ikon gerigi "Project Settings" di kiri,
 *    scroll ke "Script Properties", klik "Add script property", isi:
 *      Property: SYNC_TOKEN
 *      Value:    (bikin sendiri, string acak yang cuma kamu tahu)
 *    Simpan. Token ini sama yang nanti dipakai di pengaturan sinkronisasi
 *    di app Kalkulator Waktu.
 * 4. Klik Deploy > New deployment > pilih tipe "Web app".
 *      Execute as: Me
 *      Who has access: Anyone
 *    Klik Deploy, izinkan akses yang diminta Google.
 * 5. Salin URL Web App yang muncul (diakhiri "/exec"). Itu yang kamu
 *    tempel ke kolom "URL Web App" di pengaturan sinkronisasi app.
 *
 * Setiap kali kamu ganti isi file ini, kamu perlu Deploy > Manage
 * deployments > edit (pensil) > New version, supaya URL yang sama
 * pakai kode terbaru.
 */

var SHEET_NAME = "Momen";

function doGet(e) {
  var token = e.parameter.token;
  var check = checkToken_(token);
  if (check) return jsonOutput_(check);

  var sheet = getSheet_();
  var rows = sheet.getDataRange().getValues();
  var entries = [];
  for (var i = 1; i < rows.length; i++) {
    var r = rows[i];
    if (!r[0]) continue;
    entries.push({ id: String(r[0]), label: String(r[1] || ""), date: String(r[2] || "") });
  }
  return jsonOutput_({ entries: entries, updatedAt: getUpdatedAt_() });
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOutput_({ error: "bad_request" });
  }

  var check = checkToken_(body.token);
  if (check) return jsonOutput_(check);

  var entries = Array.isArray(body.entries) ? body.entries : [];
  var sheet = getSheet_();
  sheet.clearContents();

  // Force the whole range to plain text BEFORE writing anything, otherwise
  // Sheets auto-detects the "Tanggal" column as a date/time and silently
  // rewrites it in its own format, which the app can no longer parse back
  // into a datetime-local value.
  sheet.getRange(1, 1, entries.length + 1, 4).setNumberFormat("@");

  sheet.getRange(1, 1, 1, 4).setValues([["ID", "Nama Momen", "Tanggal", "Diperbarui"]]);

  var now = new Date().toISOString();
  if (entries.length) {
    var values = entries.map(function (en) {
      return [String(en.id || ""), String(en.label || ""), String(en.date || ""), now];
    });
    sheet.getRange(2, 1, values.length, 4).setValues(values);
  }
  setUpdatedAt_(now);
  return jsonOutput_({ ok: true, updatedAt: now });
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 4).setValues([["ID", "Nama Momen", "Tanggal", "Diperbarui"]]);
  }
  return sheet;
}

function checkToken_(token) {
  var expected = PropertiesService.getScriptProperties().getProperty("SYNC_TOKEN");
  if (!expected) return { error: "server_misconfigured", message: "SYNC_TOKEN belum diatur di Script Properties." };
  if (!token || token !== expected) return { error: "unauthorized" };
  return null;
}

function getUpdatedAt_() {
  return PropertiesService.getScriptProperties().getProperty("UPDATED_AT") || "";
}

function setUpdatedAt_(v) {
  PropertiesService.getScriptProperties().setProperty("UPDATED_AT", v);
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
