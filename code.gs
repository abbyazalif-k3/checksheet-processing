/**
 * BACKEND CEK SHEET DIGITAL — Google Apps Script (v2 — format sesuai kertas)
 * ------------------------------------------------------------------------
 * PENTING kalau kamu upgrade dari versi lama: sheet lama HARUS ditambah kolom baru.
 *
 * Sheet "Areas"     -> id | name | code | active                     (tidak berubah)
 * Sheet "Templates" -> id | areaId | section | label | type | order  (TAMBAH kolom "section" setelah areaId)
 * Sheet "Entries"    -> id | areaId | areaName | shift | operator | datetime | valuesJSON | correctiveActionsJSON | confirmed
 *                       (TAMBAH kolom "correctiveActionsJSON" dan "confirmed" di akhir)
 *
 * Cara tambah kolom di sheet yang sudah ada:
 * - Buka sheet Templates, klik kanan kolom B (areaId) -> Insert 1 column right -> beri nama header "section"
 * - Buka sheet Entries, tambahkan 2 kolom baru di akhir (setelah valuesJSON): "correctiveActionsJSON" dan "confirmed"
 *
 * Setelah itu: tempel ulang SEMUA isi file ini ke Apps Script (ganti yang lama),
 * lalu Deploy > Manage deployments > Edit (ikon pensil) > New version > Deploy.
 * (Pakai "Manage deployments" supaya URL Web App TIDAK berubah, tidak perlu ganti API_URL di HTML.)
 */

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = {
    areas: sheetToObjects(ss.getSheetByName('Areas')),
    templates: sheetToObjects(ss.getSheetByName('Templates')),
    entries: sheetToObjects(ss.getSheetByName('Entries')).map(function (en) {
      try { en.values = JSON.parse(en.valuesJSON || '{}'); } catch (err) { en.values = {}; }
      try { en.correctiveActions = JSON.parse(en.correctiveActionsJSON || '[]'); } catch (err) { en.correctiveActions = []; }
      return en;
    })
  };
  return jsonResponse(data);
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ ok: false, error: 'Payload tidak valid' });
  }

  try {
    switch (body.action) {
      case 'loginLHForeman': {
  var userId = String(body.data.userId || '').trim();
  var pin = String(body.data.pin || '').trim();

  return jsonResponse(
    loginLHForeman(ss, userId, pin)
  );
}
      case 'addArea': {
        var auth = requireLHForeman(body);
        if (!auth.ok) return jsonResponse(auth);
        appendRow(ss, 'Areas', [body.data.id, body.data.name, body.data.code, true]);
        break;
      }
      case 'toggleArea': {
        var auth = requireLHForeman(body);
        if (!auth.ok) return jsonResponse(auth);
        updateCellByRowId(ss, 'Areas', body.data.id, 'active', body.data.active);
        break;
      }
      case 'addField': {
        var auth = requireLHForeman(body);
        if (!auth.ok) return jsonResponse(auth);
        appendRow(ss, 'Templates', [
          body.data.id, body.data.areaId, body.data.section || 'Umum',
          body.data.label, body.data.type, body.data.order
        ]);
        break;
      }
      case 'removeField': {
        var auth = requireLHForeman(body);
        if (!auth.ok) return jsonResponse(auth);
        deleteRowById(ss, 'Templates', body.data.id);
        break;
      }
case 'addEntry':
  appendRow(ss, 'Entries', [
    body.data.id,
    body.data.areaId,
    body.data.areaName,
    body.data.shift,
    body.data.crew || '',
    body.data.operator,
    body.data.lhForeman || '',
    body.data.datetime,
    JSON.stringify(body.data.values),
    JSON.stringify(body.data.correctiveActions || []),
    !!body.data.confirmed,
    body.data.lhComment || ''
  ]);
  break;
  case 'confirmEntry': {
    var auth = requireLHForeman(body);
    if (!auth.ok) {
      return jsonResponse(auth);
    }

  var entryId = String(body.data.entryId || '').trim();
  var lhForeman = String(body.data.lhForeman || '').trim();
  var lhComment = String(body.data.lhComment || '').trim();
  if (!entryId) {
    return jsonResponse({
      ok: false,
      error: 'entryId wajib diisi'
    });
  }

  if (!lhForeman) {
    return jsonResponse({
      ok: false,
      error: 'Nama LH/Foreman wajib diisi'
    });
  }

  var entriesSheet = ss.getSheetByName('Entries');
  var rowIndex = findRowIndexById(entriesSheet, entryId);

  if (rowIndex < 0) {
    return jsonResponse({
      ok: false,
      error: 'Pemeriksaan tidak ditemukan'
    });
  }

var headers = entriesSheet
  .getDataRange()
  .getValues()[0];

var normalizedHeaders = headers.map(function(header) {
  return String(header || '').trim().toLowerCase();
});

var confirmedCol = normalizedHeaders.indexOf('confirmed') + 1;
var lhForemanCol = normalizedHeaders.indexOf('lhforeman') + 1;
var lhCommentCol = normalizedHeaders.indexOf('lhcomment') + 1;

if (
  confirmedCol <= 0 ||
  lhForemanCol <= 0 ||
  lhCommentCol <= 0
) {
  return jsonResponse({
    ok: false,
    error: 'Kolom lhforeman, confirmed, atau lhComment tidak ditemukan'
  });
}

  var currentConfirmed =
    entriesSheet.getRange(rowIndex, confirmedCol).getValue();

  if (currentConfirmed === true ||
      String(currentConfirmed).toUpperCase() === 'TRUE') {
    return jsonResponse({
      ok: false,
      error: 'Pemeriksaan sudah dikonfirmasi'
    });
  }

entriesSheet
  .getRange(rowIndex, lhForemanCol)
  .setValue(lhForeman);

entriesSheet
  .getRange(rowIndex, lhCommentCol)
  .setValue(lhComment);

entriesSheet
  .getRange(rowIndex, confirmedCol)
  .setValue(true);

SpreadsheetApp.flush();

return jsonResponse({
  ok: true,
  message: 'Pemeriksaan berhasil dikonfirmasi',
  entryId: entryId,
  lhForeman: lhForeman,
  lhComment: lhComment,
  confirmed: true
});
}
  break;
      case 'logoutLHForeman': {
        var auth = requireLHForeman(body);
        if (!auth.ok) return jsonResponse(auth);
        var authToken = String(body.data.authToken || '').trim();
        CacheService.getScriptCache().remove('lh_session_' + authToken);
        return jsonResponse({ ok: true });
      }
      case 'deleteEntry':
        var auth = requireLHForeman(body);
        if (!auth.ok) return jsonResponse(auth);
        deleteRowById(ss, 'Entries', body.data.id);
        break;
      default:
        return jsonResponse({ ok: false, error: 'Aksi tidak dikenal: ' + body.action });
    }
    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ---------- Helper functions ----------

function loginLHForeman(ss, userId, pin) {

  userId = String(userId || '').trim();
  pin = String(pin || '').trim();

  if (!userId || !pin) {
    return {
      ok: false,
      error: 'ID dan PIN wajib diisi'
    };
  }

  var usersSheet = ss.getSheetByName('Users');

  if (!usersSheet) {
    return {
      ok: false,
      error: 'Sheet Users tidak ditemukan'
    };
  }

  var data = usersSheet.getDataRange().getValues();

  if (data.length < 2) {
    return {
      ok: false,
      error: 'Belum ada pengguna terdaftar'
    };
  }

  var headers = data[0].map(function(header) {
    return String(header || '').trim().toLowerCase();
  });

  var userIdCol = headers.indexOf('userid');
  var namaCol = headers.indexOf('nama');
  var roleCol = headers.indexOf('role');
  var pinCol = headers.indexOf('pin');
  var aktifCol = headers.indexOf('aktif');

  if (
    userIdCol < 0 ||
    namaCol < 0 ||
    roleCol < 0 ||
    pinCol < 0 ||
    aktifCol < 0
  ) {
    return {
      ok: false,
      error: 'Kolom Users tidak lengkap'
    };
  }

  for (var i = 1; i < data.length; i++) {

    var row = data[i];

    var rowUserId = String(row[userIdCol] || '').trim();
    var rowPin = String(row[pinCol] || '').trim();
    var rowRole = String(row[roleCol] || '').trim().toUpperCase();
    var rowAktif = String(row[aktifCol] || '').trim().toUpperCase();

    if (rowUserId === userId && rowPin === pin) {

      if (rowAktif !== 'TRUE') {
        return {
          ok: false,
          error: 'Pengguna tidak aktif'
        };
      }

      if (rowRole !== 'LH' && rowRole !== 'FOREMAN') {
        return {
          ok: false,
          error: 'Pengguna tidak memiliki akses LH/Foreman'
        };
      }

      var token = createLHSession(
        rowUserId,
        String(row[namaCol] || '').trim(),
        rowRole
      );

      return {
        ok: true,
        userId: rowUserId,
        nama: String(row[namaCol] || '').trim(),
        role: rowRole,
        authToken: token
      };
    }
  }

  return {
    ok: false,
    error: 'ID atau PIN salah'
  };
}

function createLHSession(userId, nama, role) {
  var token = Utilities.getUuid();

  var session = {
    userId: userId,
    nama: nama,
    role: role,
    createdAt: new Date().getTime()
  };

  CacheService
    .getScriptCache()
    .put(
      'lh_session_' + token,
      JSON.stringify(session),
      21600
    );

  return token;
}

function getLHSession(token) {
  token = String(token || '').trim();

  if (!token) return null;

  var raw = CacheService
    .getScriptCache()
    .get('lh_session_' + token);

  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}

function requireLHForeman(body) {
  var token = body && body.data && body.data.authToken
    ? String(body.data.authToken).trim()
    : '';

  var session = getLHSession(token);

  if (!session) {
    return {
      ok: false,
      error: 'Sesi LH/Foreman tidak valid atau sudah berakhir'
    };
  }

  var role = String(session.role || '')
    .trim()
    .toUpperCase();

  if (role !== 'LH' && role !== 'FOREMAN') {
    return {
      ok: false,
      error: 'Tidak memiliki akses LH/Foreman'
    };
  }

  return {
    ok: true,
    session: session
  };
}
function sheetToObjects(sheet) {
  if (!sheet) return [];
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  var headers = values[0];
  var rows = values.slice(1);
  return rows
    .filter(function (r) { return r[0] !== '' && r[0] !== null; })
    .map(function (r) {
      var obj = {};
      headers.forEach(function (h, i) { obj[h] = r[i]; });
      return obj;
    });
}

function appendRow(ss, sheetName, rowArray) {
  var sheet = ss.getSheetByName(sheetName);
  sheet.appendRow(rowArray);
}

function findRowIndexById(sheet, id) {
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

function deleteRowById(ss, sheetName, id) {
  var sheet = ss.getSheetByName(sheetName);
  var rowIndex = findRowIndexById(sheet, id);
  if (rowIndex > 0) sheet.deleteRow(rowIndex);
}

function updateCellByRowId(ss, sheetName, id, columnName, value) {
  var sheet = ss.getSheetByName(sheetName);
  var headers = sheet.getDataRange().getValues()[0];
  var colIndex = headers.indexOf(columnName) + 1;
  var rowIndex = findRowIndexById(sheet, id);
  if (rowIndex > 0 && colIndex > 0) sheet.getRange(rowIndex, colIndex).setValue(value);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


function isiChecksheetElution() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Templates");
  const areasSheet = ss.getSheetByName("Areas");

  if (!sheet) {
    throw new Error('Sheet "Templates" tidak ditemukan.');
  }

  if (!areasSheet) {
    throw new Error('Sheet "Areas" tidak ditemukan.');
  }

  // =========================================================
  // CARI AREA ELUTION
  // =========================================================

  const areaData = areasSheet.getDataRange().getValues();

  if (areaData.length < 2) {
    throw new Error("Sheet Areas belum mempunyai data area.");
  }

  const areaHeaders = areaData[0];

  const idCol = areaHeaders.indexOf("id");
  const nameCol = areaHeaders.indexOf("name");
  const codeCol = areaHeaders.indexOf("code");

  if (idCol < 0 || nameCol < 0 || codeCol < 0) {
    throw new Error(
      'Header Areas harus mempunyai: id | name | code | active'
    );
  }

  let areaId = null;
  let areaName = null;
  let areaCode = null;

  for (let i = 1; i < areaData.length; i++) {
    const row = areaData[i];

    const name = String(row[nameCol] || "").trim();
    const code = String(row[codeCol] || "").trim();

    if (
      name.toLowerCase() === "elution area" ||
      code === "TT-PRO-FO-62-001"
    ) {
      areaId = row[idCol];
      areaName = name;
      areaCode = code;
      break;
    }
  }

  if (!areaId) {
    throw new Error(
      'Area "Elution Area" dengan kode "TT-PRO-FO-62-001" tidak ditemukan di sheet Areas.'
    );
  }

  // =========================================================
  // DATA CHECKSHEET SESUAI KERTAS
  // =========================================================

  const TEMPLATE_ID = "ELUTION";

  const data = [

    // -------------------------------------------------------
    // 1. GENERAL
    // -------------------------------------------------------

    ["ELU-1.1", areaId, "GENERAL",
      "Periksa kerapihan-kebersihan semua area, termasuk area bund, catwalks, grid mesh dan akses jalan terbebas dari bahaya tersandung, sampah, dan semua selang yang tidak digunakan telah digulung rapi",
      "rating", "1.1"],

    ["ELU-1.2", areaId, "GENERAL",
      "Periksa kondisi grating floor",
      "rating", "1.2"],

    ["ELU-1.3", areaId, "GENERAL",
      "Periksa apakah Safety Showers/Eyewash berfungsi",
      "rating", "1.3"],

    ["ELU-1.4", areaId, "GENERAL",
      "Periksa tempat sampah tidak penuh",
      "rating", "1.4"],

    ["ELU-1.5", areaId, "GENERAL",
      "Periksa apakah Fire Extinguisher ada pada tempatnya",
      "rating", "1.5"],

    ["ELU-1.6", areaId, "GENERAL",
      "Periksa suplai air decant di area elution",
      "rating", "1.6"],

    ["ELU-1.7", areaId, "GENERAL",
      "Periksa kondisi sump pump di area elution",
      "rating", "1.7"],

    ["ELU-1.8", areaId, "GENERAL",
      "Periksa tumpahan oli di lantai heater dan bersihkan jika ada",
      "rating", "1.8"],

    ["ELU-1.9", areaId, "GENERAL",
      "Periksa pH meter portable berfungsi dengan baik dan disimpan secara tepat (probe terendam air)",
      "rating", "1.9"],

    ["ELU-1.10", areaId, "GENERAL",
      "Periksa kondisi instrument pengukur seperti pressure gauge, level sensor, dan flow sensor.",
      "rating", "1.10"],


    // -------------------------------------------------------
    // 2. CARBON SCREEN
    // -------------------------------------------------------

    ["ELU-2.1", areaId, "CARBON SCREEN",
      "Periksa kebersihan dan ketersediaan semua carbon de-watering screens",
      "rating", "2.1"],

    ["ELU-2.2", areaId, "CARBON SCREEN",
      "Periksa kebocoran pipa dan sumbatan di water sprays",
      "rating", "2.2"],

    ["ELU-2.3", areaId, "CARBON SCREEN",
      "Periksa panel screen tidak tersumbat",
      "rating", "2.3"],

    ["ELU-2.4", areaId, "CARBON SCREEN",
      "Periksa baut pada vibrator screen",
      "rating", "2.4"],


    // -------------------------------------------------------
    // 3. ACID COLUMN
    // -------------------------------------------------------

    ["ELU-3.1", areaId, "ACID COLUMN",
      "Periksa kebocoran di semua tangki, pipa dan area bunded",
      "rating", "3.1"],

    ["ELU-3.2", areaId, "ACID COLUMN",
      "Periksa kebocoran dari strainer di dalam acid column",
      "rating", "3.2"],

    ["ELU-3.3", areaId, "ACID COLUMN",
      "Periksa kondisi valve di top dan bottom column",
      "rating", "3.3"],


    // -------------------------------------------------------
    // 4. ELUTION COLUMN
    // -------------------------------------------------------

    ["ELU-4.1", areaId, "ELUTION COLUMN",
      "Periksa status dan kesiapan elution column",
      "rating", "4.1"],

    ["ELU-4.2", areaId, "ELUTION COLUMN",
      "Periksa level oli di elution heater",
      "rating", "4.2"],

    ["ELU-4.3", areaId, "ELUTION COLUMN",
      "Periksa kondisi drat penahan cover strainer",
      "rating", "4.3"],

    ["ELU-4.4", areaId, "ELUTION COLUMN",
      "Periksa kondisi O-ring basket strainer",
      "rating", "4.4"],

    ["ELU-4.5", areaId, "ELUTION COLUMN",
      "Periksa kebocoran oli di elution heater",
      "rating", "4.5"],

    ["ELU-4.6", areaId, "ELUTION COLUMN",
      "Periksa kebocoran di line sianida",
      "rating", "4.6"],

    ["ELU-4.7", areaId, "ELUTION COLUMN",
      "Periksa kebocoran di line caustic soda",
      "rating", "4.7"],

    ["ELU-4.8", areaId, "ELUTION COLUMN",
      "Periksa kondisi filter",
      "rating", "4.8"],

    ["ELU-4.9", areaId, "ELUTION COLUMN",
      "Periksa dan pastikan tidak ada kebocoran di elution circuit",
      "rating", "4.9"],

    ["ELU-4.10", areaId, "ELUTION COLUMN",
      "Periksa kondisi relief valve di elution column 1 dan 2",
      "rating", "4.10"],

    ["ELU-4.11", areaId, "ELUTION COLUMN",
      "Periksa kebocoran di heat exchanger",
      "rating", "4.11"],

    ["ELU-4.12", areaId, "ELUTION COLUMN",
      "Periksa kondisi pompa sulphamic acid dan unit tangki mixing",
      "rating", "4.12"],

    ["ELU-4.13", areaId, "ELUTION COLUMN",
      "Apakah ada injeksi sulphamic acid ke Heat Exchanger",
      "rating", "4.13"],


    // -------------------------------------------------------
    // 5. ELECTROWINNING
    // -------------------------------------------------------

    ["ELU-5.1", areaId, "ELECTROWINNING",
      "Periksa kebocoran di semua tangki, pipa dan area bunded",
      "rating", "5.1"],

    ["ELU-5.2", areaId, "ELECTROWINNING",
      "Periksa rectifier berfungsi normal",
      "rating", "5.2"],


    // -------------------------------------------------------
    // 6. REGEN KILN
    // -------------------------------------------------------

    ["ELU-6.1", areaId, "REGEN KILN",
      "Periksa status dan kesiapan kiln",
      "rating", "6.1"],

    ["ELU-6.2", areaId, "REGEN KILN",
      "Periksa kebocoran pada semua pipa dan pompa",
      "rating", "6.2"],

    ["ELU-6.3", areaId, "REGEN KILN",
      "Periksa level karbon di kiln hopper dan quench tank",
      "rating", "6.3"],

    ["ELU-6.4", areaId, "REGEN KILN",
      "Periksa kondisi pompa transfer karbon di quench tank",
      "rating", "6.4"],

    ["ELU-6.5", areaId, "REGEN KILN",
      "Periksa kondisi air flushing",
      "rating", "6.5"],

    ["ELU-6.6", areaId, "REGEN KILN",
      "Periksa kondisi pompa 62PU009",
      "rating", "6.6"],

    ["ELU-6.7", areaId, "REGEN KILN",
      "Periksa kondisi pompa 62PU008",
      "rating", "6.7"],

    ["ELU-6.8", areaId, "REGEN KILN",
      "Periksa temperatur kiln",
      "rating", "6.8"],

    ["ELU-6.9", areaId, "REGEN KILN",
      "Periksa kondisi screw feeder",
      "rating", "6.9"],

    ["ELU-6.10", areaId, "REGEN KILN",
      "Periksa kondisi burner",
      "rating", "6.10"],

    ["ELU-6.11", areaId, "REGEN KILN",
      "Periksa kondisi sistem sprinkler",
      "rating", "6.11"],


    // -------------------------------------------------------
    // 7. REAGEN DAN PENGGUNAANNYA
    // -------------------------------------------------------

    ["ELU-7.1", areaId, "REAGEN DAN PENGGUNAANNYA",
      "Periksa level tangki storage acid, kebocoran pipa, kebersihan bund dan kesiapan pompa untuk dioperasikan",
      "rating", "7.1"],

    ["ELU-7.2", areaId, "REAGEN DAN PENGGUNAANNYA",
      "Periksa level tangki solar dan level burner oil",
      "rating", "7.2"],

    ["ELU-7.3", areaId, "REAGEN DAN PENGGUNAANNYA",
      "Periksa dan pastikan air di dalam prep tank dalam kondisi penuh",
      "rating", "7.3"]
  ];

  // =========================================================
  // CEK HEADER TEMPLATE
  // =========================================================

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const requiredHeaders = [
    "id",
    "areaId",
    "section",
    "label",
    "type",
    "order"
  ];

  requiredHeaders.forEach(function(header) {
    if (headers.indexOf(header) === -1) {
      throw new Error(
        'Kolom "' + header + '" tidak ditemukan di sheet Templates.'
      );
    }
  });

  // =========================================================
  // HAPUS CHECKSHEET ELUTION YANG SUDAH ADA
  // =========================================================

  const values = sheet.getDataRange().getValues();

  const areaIdIndex = headers.indexOf("areaId");

  for (let r = values.length - 1; r >= 1; r--) {

    if (String(values[r][areaIdIndex]) === String(areaId)) {
      sheet.deleteRow(r + 1);
    }

  }

  // =========================================================
  // SUSUN DATA SESUAI HEADER SHEET
  // =========================================================

  const rows = data.map(function(item) {

    const row = new Array(headers.length).fill("");

    row[headers.indexOf("id")] = item[0];
    row[headers.indexOf("areaId")] = item[1];
    row[headers.indexOf("section")] = item[2];
    row[headers.indexOf("label")] = item[3];
    row[headers.indexOf("type")] = item[4];
    row[headers.indexOf("order")] = item[5];

    return row;

  });

  // =========================================================
// PAKSA KOLOM ORDER MENJADI TEKS
// Supaya 1.1, 1.2, 1.10 tidak diubah Google Sheets
// menjadi tanggal.
// =========================================================

const orderColumn = headers.indexOf("order") + 1;

sheet
  .getRange(1, orderColumn, sheet.getMaxRows(), 1)
  .setNumberFormat("@");


// =========================================================
// TULIS SEMUA 46 POIN
// =========================================================

const startRow = sheet.getLastRow() + 1;

sheet
  .getRange(startRow, 1, rows.length, headers.length)
  .setValues(rows);

  // =========================================================
  // FORMAT
  // =========================================================

  sheet
    .getRange(startRow, 1, rows.length, headers.length)
    .setVerticalAlignment("middle");

  SpreadsheetApp.flush();

  Logger.log(
    "Checksheet Elution berhasil dimasukkan."
  );

  Logger.log(
    "Area      : " + areaName
  );

  Logger.log(
    "Kode      : " + areaCode
  );

  Logger.log(
    "Jumlah    : " + rows.length + " poin"
  );

  Logger.log(
    "Template  : Elution Area Prestart Check"
  );

  return "Berhasil memasukkan " + rows.length + " poin checksheet Elution Area.";
}
