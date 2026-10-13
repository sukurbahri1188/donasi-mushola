/**
 * ============================================================
 * BACKEND GOOGLE APPS SCRIPT
 * Mushola SDN 13 Selat Penuguan — Aplikasi Donasi
 * ============================================================
 * CARA SETUP:
 * 1. Buka https://script.google.com → Buat project baru
 * 2. Hapus kode default, tempel SELURUH kode ini
 * 3. Deploy → New deployment → Web app
 *    Execute as: Me | Who has access: Anyone
 * 4. Salin URL Web App → sudah terpasang di file HTML
 * ============================================================
 */

var SPREADSHEET_ID = '';
var SHEET_DONORS   = 'Donatur';
var SHEET_GALLERY  = 'Galeri';
var SHEET_BERITA   = 'Berita';
var SHEET_CONFIG   = 'Config';
var SHEET_PENDING  = 'Konfirmasi_Pending'; // ← sheet konfirmasi dari pengunjung

// ============================================================
// ENTRYPOINT
// ============================================================

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    var action   = (e.parameter && e.parameter.action)   ? e.parameter.action   : '';
    var callback = (e.parameter && e.parameter.callback) ? e.parameter.callback : '';
    var data     = null;

    if (e.parameter && e.parameter.payload) {
      try { data = JSON.parse(decodeURIComponent(e.parameter.payload)); } catch(err) {}
    }
    if (!action && e.postData && e.postData.contents) {
      try {
        var body = JSON.parse(e.postData.contents);
        if (body.action) action = body.action;
        if (body.data)   data   = body.data;
      } catch(err) {}
    }

    var ss = getOrCreateSpreadsheet();
    var result;

    switch (action) {
      case 'ping':
        result = { ok: true, message: 'Server aktif', timestamp: new Date().toISOString() };
        break;

      // ---- DONATUR ----
      case 'getDonors':
        result = getDonors(ss);
        break;
      case 'saveDonors':
        result = saveDonors(ss, data || []);
        break;

      // ---- GALERI ----
      case 'getGallery':
        result = getGallery(ss);
        break;
      case 'saveGallery':
        result = saveGallery(ss, data || []);
        break;

      // ---- BERITA ----
      case 'getBerita':
        result = getBerita(ss);
        break;
      case 'saveBerita':
        result = saveBerita(ss, data || []);
        break;

      // ---- PENDING KONFIRMASI (dari pengunjung lintas perangkat) ----
      case 'addPending':
        // Pengunjung kirim konfirmasi donasi → simpan ke sheet Pending
        result = addPending(ss, data || {});
        break;
      case 'getPending':
        // Admin baca semua konfirmasi yang belum diproses
        result = getPending(ss);
        break;
      case 'updatePending':
        // Admin setujui atau tolak: { id, status: 'approved'|'rejected' }
        result = updatePending(ss, data || {});
        break;

      // ---- SEMUA DATA ----
      case 'getAllData':
        result = {
          donors:  getDonors(ss).data,
          gallery: getGallery(ss).data,
          berita:  getBerita(ss).data
        };
        break;

      default:
        result = { ok: false, error: 'Action tidak dikenal: ' + action };
    }

    var jsonStr = JSON.stringify(Object.assign({ ok: true }, result));

    if (callback) {
      return ContentService
        .createTextOutput(callback + '(' + jsonStr + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(jsonStr)
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    var errStr = JSON.stringify({ ok: false, error: err.message });
    var cb = (e && e.parameter && e.parameter.callback) ? e.parameter.callback : '';
    if (cb) {
      return ContentService
        .createTextOutput(cb + '(' + errStr + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(errStr)
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// SPREADSHEET HELPER
// ============================================================

function getOrCreateSpreadsheet() {
  var props   = PropertiesService.getScriptProperties();
  var savedId = props.getProperty('SPREADSHEET_ID');

  if (SPREADSHEET_ID) {
    try { return SpreadsheetApp.openById(SPREADSHEET_ID); } catch(e) {}
  }
  if (savedId) {
    try {
      var ss = SpreadsheetApp.openById(savedId);
      SPREADSHEET_ID = savedId;
      return ss;
    } catch(e) {}
  }

  var ss = SpreadsheetApp.create('Database Donasi Mushola SDN 13 Selat Penuguan');
  SPREADSHEET_ID = ss.getId();
  props.setProperty('SPREADSHEET_ID', SPREADSHEET_ID);
  initSheets(ss);
  return ss;
}

function initSheets(ss) {
  var defaultSheet = ss.getSheets()[0];
  defaultSheet.setName(SHEET_CONFIG);
  defaultSheet.appendRow(['key', 'value', 'updated_at']);
  defaultSheet.appendRow(['created', new Date().toISOString(), new Date().toISOString()]);

  var sheetDonors = ss.insertSheet(SHEET_DONORS);
  sheetDonors.appendRow(['id', 'name', 'amount', 'date', 'message', 'pending', 'ts']);

  var sheetGallery = ss.insertSheet(SHEET_GALLERY);
  sheetGallery.appendRow(['id', 'title', 'date', 'type', 'fileId', 'driveUrl', 'ts']);

  var sheetBerita = ss.insertSheet(SHEET_BERITA);
  sheetBerita.appendRow(['id', 'judul', 'isi', 'tanggal', 'icon', 'ts']);

  var sheetPending = ss.insertSheet(SHEET_PENDING);
  sheetPending.appendRow(['id', 'name', 'amount', 'date', 'message', 'status', 'ts', 'device']);

  var allSheets = [sheetDonors, sheetGallery, sheetBerita, sheetPending, defaultSheet];
  allSheets.forEach(function(sh) {
    sh.getRange(1, 1, 1, sh.getLastColumn())
      .setFontWeight('bold')
      .setBackground('#064e3b')
      .setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });

  ss.getSheets().forEach(function(sh) {
    if (sh.getLastColumn() > 0) sh.autoResizeColumns(1, sh.getLastColumn());
  });
}

function getOrCreateSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#064e3b')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ============================================================
// PENDING KONFIRMASI CRUD
// ============================================================

function addPending(ss, item) {
  var sheet = getOrCreateSheet(ss, SHEET_PENDING,
    ['id', 'name', 'amount', 'date', 'message', 'status', 'ts', 'device']);

  var id = 'P' + Date.now();
  sheet.appendRow([
    id,
    item.name    || '',
    item.amount  || 0,
    item.date    || '',
    item.message || '',
    'pending',           // status: pending | approved | rejected
    item._ts     || Date.now(),
    item.device  || ''
  ]);

  // Kirim notif email ke admin (opsional — aktif jika ada email terdaftar)
  try {
    var props    = PropertiesService.getScriptProperties();
    var adminEmail = props.getProperty('ADMIN_EMAIL');
    if (adminEmail) {
      MailApp.sendEmail({
        to:      adminEmail,
        subject: '🕌 Konfirmasi Donasi Baru — Mushola SDN 13',
        body:    'Ada konfirmasi donasi baru dari ' + (item.name || 'Pengunjung') +
                 '\nNominal: Rp ' + (item.amount || 0) +
                 '\nTanggal: ' + (item.date || '-') +
                 '\nPesan: ' + (item.message || '-') +
                 '\n\nSilakan buka aplikasi untuk menyetujui.'
      });
    }
  } catch(e2) { /* abaikan jika email gagal */ }

  return { saved: true, id: id };
}

function getPending(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_PENDING,
    ['id', 'name', 'amount', 'date', 'message', 'status', 'ts', 'device']);
  var rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { data: [] };

  var items = rows.slice(1).map(function(r) {
    return {
      _sheetId: String(r[0] || ''),
      name:     String(r[1] || ''),
      amount:   Number(r[2]) || 0,
      date:     r[3] ? String(r[3]).substring(0, 10) : '',
      message:  String(r[4] || ''),
      status:   String(r[5] || 'pending'),
      _ts:      Number(r[6]) || 0,
      device:   String(r[7] || ''),
      _pending: true
    };
  }).filter(function(d) {
    return d.name && d.status === 'pending'; // hanya yang belum diproses
  });

  items.sort(function(a, b) { return b._ts - a._ts; });
  return { data: items };
}

function updatePending(ss, item) {
  // item: { _sheetId: 'P123...', status: 'approved' | 'rejected' }
  var sheet = ss.getSheetByName(SHEET_PENDING);
  if (!sheet) return { ok: false, error: 'Sheet tidak ditemukan' };

  var rows = sheet.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(item._sheetId)) {
      sheet.getRange(i + 1, 6).setValue(item.status); // kolom status = kolom 6
      return { updated: true };
    }
  }
  return { ok: false, error: 'ID tidak ditemukan' };
}

// ============================================================
// DONATUR CRUD
// ============================================================

function getDonors(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_DONORS, ['id','name','amount','date','message','pending','ts']);
  var rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { data: [] };

  var donors = rows.slice(1).map(function(r) {
    return {
      id:       String(r[0] || ''),
      name:     String(r[1] || ''),
      amount:   Number(r[2]) || 0,
      date:     r[3] ? String(r[3]).substring(0, 10) : '',
      message:  String(r[4] || ''),
      _pending: r[5] === true || r[5] === 'true',
      _ts:      Number(r[6]) || 0
    };
  }).filter(function(d) { return d.name; });

  donors.sort(function(a, b) { return b._ts - a._ts; });
  return { data: donors };
}

function saveDonors(ss, donors) {
  var sheet   = getOrCreateSheet(ss, SHEET_DONORS, ['id','name','amount','date','message','pending','ts']);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  if (donors.length > 0) {
    var rows = donors.map(function(d) {
      return [
        d.id    || Date.now(),
        d.name  || '',
        d.amount || 0,
        d.date  || '',
        d.message || '',
        d._pending ? 'true' : 'false',
        d._ts   || Date.now()
      ];
    });
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }
  return { saved: donors.length };
}

// ============================================================
// GALERI CRUD
// ============================================================

function getGallery(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_GALLERY, ['id','title','date','type','fileId','driveUrl','ts']);
  var rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { data: [] };

  var items = rows.slice(1).map(function(r) {
    return {
      id:       String(r[0] || ''),
      title:    String(r[1] || ''),
      date:     r[2] ? String(r[2]).substring(0, 10) : '',
      type:     String(r[3] || 'foto'),
      fileId:   String(r[4] || ''),
      driveUrl: String(r[5] || ''),
      _ts:      Number(r[6]) || 0
    };
  }).filter(function(g) { return g.fileId || g.title; });

  items.sort(function(a, b) { return b._ts - a._ts; });
  return { data: items };
}

function saveGallery(ss, items) {
  var sheet   = getOrCreateSheet(ss, SHEET_GALLERY, ['id','title','date','type','fileId','driveUrl','ts']);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  if (items.length > 0) {
    var rows = items.map(function(g) {
      return [
        g.id       || Date.now(),
        g.title    || '',
        g.date     || '',
        g.type     || 'foto',
        g.fileId   || '',
        g.driveUrl || '',
        g._ts      || Date.now()
      ];
    });
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }
  return { saved: items.length };
}

// ============================================================
// BERITA CRUD
// ============================================================

function getBerita(ss) {
  var sheet = getOrCreateSheet(ss, SHEET_BERITA, ['id','judul','isi','tanggal','icon','ts']);
  var rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return { data: [] };

  var items = rows.slice(1).map(function(r) {
    return {
      id:      String(r[0] || ''),
      judul:   String(r[1] || ''),
      isi:     String(r[2] || ''),
      tanggal: r[3] ? String(r[3]).substring(0, 10) : '',
      icon:    String(r[4] || 'fa-star'),
      _ts:     Number(r[5]) || 0
    };
  }).filter(function(b) { return b.judul; });

  items.sort(function(a, b) { return b._ts - a._ts; });
  return { data: items };
}

function saveBerita(ss, items) {
  var sheet   = getOrCreateSheet(ss, SHEET_BERITA, ['id','judul','isi','tanggal','icon','ts']);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  if (items.length > 0) {
    var rows = items.map(function(b) {
      return [
        b.id      || Date.now(),
        b.judul   || '',
        b.isi     || '',
        b.tanggal || '',
        b.icon    || 'fa-star',
        b._ts     || Date.now()
      ];
    });
    sheet.getRange(2, 1, rows.length, 6).setValues(rows);
  }
  return { saved: items.length };
}
