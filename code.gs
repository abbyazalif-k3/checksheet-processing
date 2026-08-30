/**
 * BACKEND CHECKSHEET DIGITAL — Google Apps Script
 * ------------------------------------------------------------------------
 * Struktur Sheet:
 *
 * Areas     -> id | name | code | active | pin
 * Templates -> id | areaId | section | label | type | order
 * Users     -> userid | nama | role | pin | aktif
 * Entries   -> id | areaId | areaName | shift | crew | operator |
 *             lhForeman | datetime | valuesJSON | correctiveActionsJSON |
 *             confirmed | lhComment
 *
 * Catatan:
 * - Kolom Areas.pin dipakai untuk login Operator berdasarkan Area ID + PIN.
 * - Users.role harus LH atau FOREMAN untuk akses verifikasi/kelola.
 * - PIN tidak pernah dikirim ke frontend dan tidak disimpan di sessionStorage.
 * - Frontend menggunakan POST action=getData, bukan GET, sehingga data tidak
 *   terbuka hanya karena seseorang mengetahui URL Web App.
 *
 * Setelah mengganti Code.gs:
 * Deploy > Manage deployments > Edit > New version > Deploy.
 */

var SESSION_TTL_SECONDS = 21600; // 6 jam

function doGet(e) {
  // Jangan expose data melalui GET. Frontend menggunakan POST + session token.
  return jsonResponse({
    ok: false,
    error: 'Endpoint aktif. Gunakan POST API dari aplikasi.'
  });
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var body;

  try {
    body = JSON.parse(e && e.postData ? e.postData.contents : '{}');
  } catch (err) {
    return jsonResponse({ ok: false, error: 'Payload tidak valid' });
  }

  body = body || {};
  body.data = body.data || {};

  try {
    switch (body.action) {
      case 'loginAreaOperator':
        return jsonResponse(loginAreaOperator(ss, body.data.areaId, body.data.pin));

      case 'loginLHForeman':
        return jsonResponse(loginLHForeman(ss, body.data.userId, body.data.pin));

      case 'getData': {
        var dataAuth = requireAnySession(body);
        if (!dataAuth.ok) return jsonResponse(dataAuth);
        return jsonResponse(buildDataResponse(ss, dataAuth.session));
      }

      case 'addEntry': {
        var operatorAuth = requireOperator(body);
        if (!operatorAuth.ok) return jsonResponse(operatorAuth);
        return jsonResponse(addEntryForOperator(ss, body.data, operatorAuth.session));
      }

      case 'confirmEntry': {
        var lhAuth = requireLHForeman(body);
        if (!lhAuth.ok) return jsonResponse(lhAuth);
        return jsonResponse(confirmEntryForLH(ss, body.data, lhAuth.session));
      }

      case 'addArea': {
        var addAreaAuth = requireLHForeman(body);
        if (!addAreaAuth.ok) return jsonResponse(addAreaAuth);

        var areasForAdd = ss.getSheetByName('Areas');
        ensureColumn(areasForAdd, 'pin');

        var newAreaPin = String(body.data.pin || '').trim();
        if (!newAreaPin) {
          return jsonResponse({ ok: false, error: 'PIN Operator wajib diisi' });
        }

        appendObjectRow(areasForAdd, {
          id: body.data.id,
          name: body.data.name,
          code: body.data.code,
          active: true,
          pin: newAreaPin
        });
        return jsonResponse({ ok: true });
      }

      case 'toggleArea': {
        var toggleAuth = requireLHForeman(body);
        if (!toggleAuth.ok) return jsonResponse(toggleAuth);
        updateCellByRowId(ss, 'Areas', body.data.id, 'active', !!body.data.active);
        return jsonResponse({ ok: true });
      }

      case 'addField': {
        var addFieldAuth = requireLHForeman(body);
        if (!addFieldAuth.ok) return jsonResponse(addFieldAuth);
        appendObjectRow(ss.getSheetByName('Templates'), {
          id: body.data.id,
          areaId: body.data.areaId,
          section: body.data.section || 'Umum',
          label: body.data.label,
          type: body.data.type,
          order: body.data.order
        });
        return jsonResponse({ ok: true });
      }

      case 'removeField': {
        var removeFieldAuth = requireLHForeman(body);
        if (!removeFieldAuth.ok) return jsonResponse(removeFieldAuth);
        deleteRowById(ss, 'Templates', body.data.id);
        return jsonResponse({ ok: true });
      }

      case 'getEntries': {
        var entriesAuth = requireLHForeman(body);
        if (!entriesAuth.ok) return jsonResponse(entriesAuth);
        return jsonResponse({ ok: true, entries: getFullEntries(ss) });
      }

      case 'deleteEntry': {
        var deleteAuth = requireLHForeman(body);
        if (!deleteAuth.ok) return jsonResponse(deleteAuth);
        deleteRowById(ss, 'Entries', body.data.id);
        return jsonResponse({ ok: true });
      }

      case 'logoutLHForeman': {
        var logoutLHAuth = requireLHForeman(body);
        if (!logoutLHAuth.ok) return jsonResponse(logoutLHAuth);
        removeSession('lh_session_', body.data.authToken);
        return jsonResponse({ ok: true });
      }

      case 'logoutOperator': {
        var logoutOpAuth = requireOperator(body);
        if (!logoutOpAuth.ok) return jsonResponse(logoutOpAuth);
        removeSession('operator_session_', body.data.areaToken);
        return jsonResponse({ ok: true });
      }

      default:
        return jsonResponse({
          ok: false,
          error: 'Aksi tidak dikenal: ' + String(body.action || '')
        });
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ---------- DATA ----------

function buildDataResponse(ss, session) {
  var allAreas = sheetToObjects(ss.getSheetByName('Areas'));
  var allTemplates = sheetToObjects(ss.getSheetByName('Templates'));
  var allEntries = getFullEntries(ss);

  var areas;
  var templates;
  var entries;

  if (session.role === 'OPERATOR') {
    areas = allAreas
      .filter(function (a) { return String(a.id) === String(session.areaId); })
      .map(sanitizeArea);

    templates = allTemplates.filter(function (t) {
      return String(t.areaId) === String(session.areaId);
    });

    entries = allEntries
      .filter(function (entry) {
        return String(entry.areaId) === String(session.areaId);
      })
      .map(sanitizeEntryForOperator);
  } else {
    areas = allAreas.map(sanitizeArea);
    templates = allTemplates;
    entries = allEntries;
  }

  return {
    ok: true,
    role: session.role,
    areas: areas,
    templates: templates,
    entries: entries
  };
}

function sanitizeArea(area) {
  return {
    id: area.id,
    name: area.name,
    code: area.code,
    active: area.active === true || String(area.active).toUpperCase() === 'TRUE'
  };
}

function getFullEntries(ss) {
  return sheetToObjects(ss.getSheetByName('Entries')).map(function (entry) {
    try {
      entry.values = JSON.parse(entry.valuesJSON || '{}');
    } catch (err) {
      entry.values = {};
    }

    try {
      entry.correctiveActions = JSON.parse(entry.correctiveActionsJSON || '[]');
    } catch (err) {
      entry.correctiveActions = [];
    }

    delete entry.valuesJSON;
    delete entry.correctiveActionsJSON;
    return entry;
  });
}

function sanitizeEntryForOperator(entry) {
  return {
    id: entry.id,
    areaId: entry.areaId,
    areaName: entry.areaName,
    shift: entry.shift,
    crew: entry.crew,
    operator: entry.operator,
    lhForeman: entry.lhForeman || '',
    datetime: entry.datetime,
    confirmed: entry.confirmed,
    lhComment: entry.lhComment || '',
    values: entry.values || {},
    correctiveActions: entry.correctiveActions || []
  };
}

// ---------- LOGIN OPERATOR ----------

function loginAreaOperator(ss, areaId, pin) {
  areaId = String(areaId || '').trim();
  pin = String(pin || '').trim();

  if (!areaId || !pin) {
    return { ok: false, error: 'Area ID dan PIN wajib diisi' };
  }

  var sheet = ss.getSheetByName('Areas');
  if (!sheet) {
    return { ok: false, error: 'Sheet Areas tidak ditemukan' };
  }

  var data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    return { ok: false, error: 'Belum ada area terdaftar' };
  }

  var headers = normalizeHeaders(data[0]);
  var idCol = headers.indexOf('id');
  var nameCol = headers.indexOf('name');
  var activeCol = headers.indexOf('active');
  var pinCol = headers.indexOf('pin');

  if (idCol < 0 || nameCol < 0 || activeCol < 0 || pinCol < 0) {
    return {
      ok: false,
      error: 'Kolom Areas belum lengkap. Tambahkan kolom pin pada sheet Areas.'
    };
  }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowId = String(row[idCol] || '').trim();

    if (rowId !== areaId) continue;

    var active = String(row[activeCol] || '').trim().toUpperCase();
    if (active !== 'TRUE') {
      return { ok: false, error: 'Area tidak aktif' };
    }

    var storedPin = String(row[pinCol] || '').trim();
    if (!storedPin || storedPin !== pin) {
      return { ok: false, error: 'Area ID atau PIN salah' };
    }

    var areaName = String(row[nameCol] || '').trim();
    var token = createOperatorSession(rowId, areaName);

    return {
      ok: true,
      areaToken: token,
      resolvedAreaId: rowId,
      areaName: areaName
    };
  }

  return { ok: false, error: 'Area ID atau PIN salah' };
}

function createOperatorSession(areaId, areaName) {
  var token = Utilities.getUuid();
  var session = {
    role: 'OPERATOR',
    areaId: String(areaId),
    areaName: String(areaName || ''),
    createdAt: new Date().getTime()
  };

  CacheService.getScriptCache().put(
    'operator_session_' + token,
    JSON.stringify(session),
    SESSION_TTL_SECONDS
  );

  return token;
}

function getOperatorSession(token) {
  return getSessionFromCache('operator_session_', token);
}

function requireOperator(body) {
  var token = body && body.data && body.data.areaToken
    ? String(body.data.areaToken).trim()
    : '';
  var session = getOperatorSession(token);

  if (!session || String(session.role).toUpperCase() !== 'OPERATOR') {
    return {
      ok: false,
      error: 'Sesi Operator tidak valid atau sudah berakhir'
    };
  }

  return { ok: true, session: session };
}

// ---------- LOGIN LH / FOREMAN ----------

function loginLHForeman(ss, userId, pin) {
  userId = String(userId || '').trim();
  pin = String(pin || '').trim();

  if (!userId || !pin) {
    return { ok: false, error: 'ID dan PIN wajib diisi' };
  }

  var usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    return { ok: false, error: 'Sheet Users tidak ditemukan' };
  }

  var data = usersSheet.getDataRange().getValues();
  if (data.length < 2) {
    return { ok: false, error: 'Belum ada pengguna terdaftar' };
  }

  var headers = normalizeHeaders(data[0]);
  var userIdCol = headers.indexOf('userid');
  var namaCol = headers.indexOf('nama');
  var roleCol = headers.indexOf('role');
  var pinCol = headers.indexOf('pin');
  var aktifCol = headers.indexOf('aktif');

  if (userIdCol < 0 || namaCol < 0 || roleCol < 0 || pinCol < 0 || aktifCol < 0) {
    return { ok: false, error: 'Kolom Users tidak lengkap' };
  }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowUserId = String(row[userIdCol] || '').trim();
    var rowPin = String(row[pinCol] || '').trim();
    var rowRole = String(row[roleCol] || '').trim().toUpperCase();
    var rowAktif = String(row[aktifCol] || '').trim().toUpperCase();

    if (rowUserId !== userId || rowPin !== pin) continue;

    if (rowAktif !== 'TRUE') {
      return { ok: false, error: 'Pengguna tidak aktif' };
    }

    if (rowRole !== 'LH' && rowRole !== 'FOREMAN') {
      return { ok: false, error: 'Pengguna tidak memiliki akses LH/Foreman' };
    }

    var nama = String(row[namaCol] || '').trim();
    var token = createLHSession(rowUserId, nama, rowRole);

    return {
      ok: true,
      userId: rowUserId,
      nama: nama,
      role: rowRole,
      authToken: token
    };
  }

  return { ok: false, error: 'ID atau PIN salah' };
}

function createLHSession(userId, nama, role) {
  var token = Utilities.getUuid();
  var session = {
    role: String(role || '').toUpperCase(),
    userId: String(userId || ''),
    nama: String(nama || ''),
    createdAt: new Date().getTime()
  };

  CacheService.getScriptCache().put(
    'lh_session_' + token,
    JSON.stringify(session),
    SESSION_TTL_SECONDS
  );

  return token;
}

function getLHSession(token) {
  return getSessionFromCache('lh_session_', token);
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

  var role = String(session.role || '').trim().toUpperCase();
  if (role !== 'LH' && role !== 'FOREMAN') {
    return { ok: false, error: 'Tidak memiliki akses LH/Foreman' };
  }

  return { ok: true, session: session };
}

function requireAnySession(body) {
  var data = body && body.data ? body.data : {};
  var areaToken = String(data.areaToken || '').trim();
  var authToken = String(data.authToken || '').trim();

  if (areaToken) return requireOperator(body);
  if (authToken) return requireLHForeman(body);

  return { ok: false, error: 'Sesi tidak ditemukan. Silakan login kembali.' };
}

function getSessionFromCache(prefix, token) {
  token = String(token || '').trim();
  if (!token) return null;

  var raw = CacheService.getScriptCache().get(prefix + token);
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}

function removeSession(prefix, token) {
  token = String(token || '').trim();
  if (token) CacheService.getScriptCache().remove(prefix + token);
}

// ---------- ENTRY ----------

function addEntryForOperator(ss, data, session) {
  var areaId = String(data.areaId || '').trim();
  var operator = String(data.operator || '').trim();

  if (!areaId) return { ok: false, error: 'Area wajib diisi' };
  if (!operator) return { ok: false, error: 'Nama operator wajib diisi' };

  if (String(session.areaId) !== areaId) {
    return { ok: false, error: 'Operator tidak memiliki akses ke area ini' };
  }

  var area = findObjectById(ss.getSheetByName('Areas'), areaId);
  if (!area) return { ok: false, error: 'Area tidak ditemukan' };

  if (String(area.active).toUpperCase() !== 'TRUE' && area.active !== true) {
    return { ok: false, error: 'Area tidak aktif' };
  }

  var values = data.values && typeof data.values === 'object' ? data.values : {};
  var correctiveActions = Array.isArray(data.correctiveActions) ? data.correctiveActions : [];

  appendObjectRow(ss.getSheetByName('Entries'), {
    id: data.id || Utilities.getUuid(),
    areaId: areaId,
    areaName: String(area.name || ''),
    shift: String(data.shift || ''),
    crew: String(data.crew || ''),
    operator: operator,
    lhForeman: '',
    datetime: data.datetime || new Date().toISOString(),
    valuesJSON: JSON.stringify(values),
    correctiveActionsJSON: JSON.stringify(correctiveActions),
    confirmed: false,
    lhComment: ''
  });

  return { ok: true };
}

function confirmEntryForLH(ss, data, session) {
  var entryId = String(data.entryId || '').trim();
  var lhComment = String(data.lhComment || '').trim();

  if (!entryId) return { ok: false, error: 'entryId wajib diisi' };

  var entriesSheet = ss.getSheetByName('Entries');
  if (!entriesSheet) return { ok: false, error: 'Sheet Entries tidak ditemukan' };

  var rowIndex = findRowIndexById(entriesSheet, entryId);
  if (rowIndex < 0) return { ok: false, error: 'Pemeriksaan tidak ditemukan' };

  var headers = normalizeHeaders(entriesSheet.getDataRange().getValues()[0]);
  var confirmedCol = headers.indexOf('confirmed') + 1;
  var lhForemanCol = headers.indexOf('lhforeman') + 1;
  var lhCommentCol = headers.indexOf('lhcomment') + 1;

  if (confirmedCol <= 0 || lhForemanCol <= 0 || lhCommentCol <= 0) {
    return {
      ok: false,
      error: 'Kolom Entries belum lengkap. Pastikan ada lhForeman, confirmed, dan lhComment.'
    };
  }

  var currentConfirmed = entriesSheet.getRange(rowIndex, confirmedCol).getValue();
  if (currentConfirmed === true || String(currentConfirmed).toUpperCase() === 'TRUE') {
    return { ok: false, error: 'Pemeriksaan sudah dikonfirmasi' };
  }

  // Nama LH/Foreman berasal dari session server, bukan dari input client.
  var lhForeman = String(session.nama || '').trim();
  if (!lhForeman) return { ok: false, error: 'Nama LH/Foreman pada akun belum diisi' };

  entriesSheet.getRange(rowIndex, lhForemanCol).setValue(lhForeman);
  entriesSheet.getRange(rowIndex, lhCommentCol).setValue(lhComment);
  entriesSheet.getRange(rowIndex, confirmedCol).setValue(true);
  SpreadsheetApp.flush();

  return {
    ok: true,
    message: 'Pemeriksaan berhasil dikonfirmasi',
    entryId: entryId,
    lhForeman: lhForeman,
    lhComment: lhComment,
    confirmed: true
  };
}

// ---------- HELPERS ----------

function normalizeHeaders(headers) {
  return headers.map(function (header) {
    return String(header || '').trim().toLowerCase();
  });
}

function sheetToObjects(sheet) {
  if (!sheet) return [];

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  var headers = values[0];
  return values.slice(1)
    .filter(function (row) {
      return row[0] !== '' && row[0] !== null;
    })
    .map(function (row) {
      var obj = {};
      headers.forEach(function (header, index) {
        obj[header] = row[index];
      });
      return obj;
    });
}


function ensureColumn(sheet, columnName) {
  if (!sheet) throw new Error('Sheet tidak ditemukan');

  var headers = sheet.getDataRange().getValues()[0];
  var normalized = normalizeHeaders(headers);
  if (normalized.indexOf(String(columnName).toLowerCase()) >= 0) return;

  var newColumn = sheet.getLastColumn() + 1;
  sheet.getRange(1, newColumn).setValue(columnName);
}

function appendObjectRow(sheet, data) {
  if (!sheet) throw new Error('Sheet tidak ditemukan');

  var range = sheet.getDataRange();
  var headers = range.getValues()[0];
  var normalized = normalizeHeaders(headers);
  var row = new Array(headers.length).fill('');

  Object.keys(data).forEach(function (key) {
    var index = normalized.indexOf(String(key).toLowerCase());
    if (index >= 0) row[index] = data[key];
  });

  sheet.appendRow(row);
}

function findObjectById(sheet, id) {
  var rows = sheetToObjects(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].id) === String(id)) return rows[i];
  }
  return null;
}

function findRowIndexById(sheet, id) {
  if (!sheet) return -1;
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
  if (!sheet) throw new Error('Sheet ' + sheetName + ' tidak ditemukan');

  var headers = sheet.getDataRange().getValues()[0];
  var normalized = normalizeHeaders(headers);
  var colIndex = normalized.indexOf(String(columnName).toLowerCase()) + 1;
  var rowIndex = findRowIndexById(sheet, id);

  if (rowIndex > 0 && colIndex > 0) {
    sheet.getRange(rowIndex, colIndex).setValue(value);
    return true;
  }

  throw new Error('Data atau kolom tidak ditemukan: ' + columnName);
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
