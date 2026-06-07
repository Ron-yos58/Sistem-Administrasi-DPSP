function getSpreadsheet_() {
  return SpreadsheetApp.openById(APP_CONFIG.SPREADSHEET_ID);
}

function getSheet_(key, required) {
  const descriptor = APP_CONFIG.SHEETS[key];
  if (!descriptor) throw new Error('Konfigurasi sheet tidak dikenal: ' + key);

  const ss = getSpreadsheet_();
  const names = [descriptor.name].concat(descriptor.aliases || []);
  for (let i = 0; i < names.length; i++) {
    const sheet = ss.getSheetByName(names[i]);
    if (sheet) return sheet;
  }

  if (required === false) return null;
  throw new Error('Sheet "' + descriptor.name + '" tidak ditemukan. Jalankan setupSystem().');
}

function getCurrentUser_() {
  return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
}

function assertAuthorized_() {
  const email = getCurrentUser_();
  if (!email) {
    throw new Error(
      'Identitas pengguna tidak tersedia. Deploy web app sebagai "User accessing the web app" dan batasi akses ke domain.'
    );
  }

  const accessSheet = getSheet_('ACCESS', false);
  if (!accessSheet || accessSheet.getLastRow() < 2) {
    throw new Error('Config_Access belum berisi pengguna aktif. Jalankan setupSystem() dengan akun admin awal.');
  }

  const values = accessSheet.getRange(2, 1, accessSheet.getLastRow() - 1, 3).getDisplayValues();
  const match = values.find(function(row) {
    return String(row[0]).trim().toLowerCase() === email && isAccessActive_(row[1]);
  });
  if (!match) throw new Error('Akun ' + email + ' tidak memiliki akses.');
  return { email: email, role: normalizeRole_(match[2]) };
}

function isAccessActive_(value) {
  const state = String(value == null ? '' : value).trim().toLowerCase();
  return ['false', 'tidak', 'no', '0', 'inactive', 'nonaktif', 'disabled'].indexOf(state) === -1;
}

function normalizeRole_(value) {
  const role = String(value || '').trim().toUpperCase();
  if (['ADMIN', 'OPERATOR'].indexOf(role) === -1) {
    throw new Error('Role akses tidak valid: ' + role);
  }
  return role;
}

function assertAdmin_(user) {
  if ((user || {}).role !== 'ADMIN') throw new Error('Hanya Admin dapat menjalankan operasi ini.');
}

function withScriptLock_(callback) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Sistem sedang memproses perubahan lain. Coba kembali beberapa detik.');
  }
  try {
    const result = callback();
    SpreadsheetApp.flush();
    return result;
  } finally {
    lock.releaseLock();
  }
}

function text_(value, maxLength) {
  const output = value == null ? '' : String(value).trim();
  return maxLength ? output.slice(0, maxLength) : output;
}

function emailList_(value) {
  const source = Array.isArray(value) ? value : String(value || '').split(/[\n,;]+/);
  const seen = {};
  return source.map(function(item) {
    return String(item || '').trim().toLowerCase();
  }).filter(function(item) {
    if (!item || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(item) || seen[item]) return false;
    seen[item] = true;
    return true;
  });
}

function uniqueTextList_(values) {
  const seen = {};
  return values.map(function(value) {
    return text_(value);
  }).filter(function(value) {
    if (!value || seen[value]) return false;
    seen[value] = true;
    return true;
  });
}

function formatDate_(value, pattern) {
  if (!value) return '';
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return text_(value);
  return Utilities.formatDate(date, APP_CONFIG.TIME_ZONE, pattern || 'yyyy-MM-dd');
}

function parseIsoDate_(value) {
  const clean = text_(value);
  if (!clean) return null;
  const match = clean.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) throw new Error('Format tanggal harus YYYY-MM-DD: ' + clean);
  const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12, 0, 0);
  if (
    date.getFullYear() !== Number(match[1]) ||
    date.getMonth() !== Number(match[2]) - 1 ||
    date.getDate() !== Number(match[3])
  ) {
    throw new Error('Tanggal tidak valid: ' + clean);
  }
  return date;
}

function formatIndonesianDate_(value) {
  if (!value) return '';
  const date = value instanceof Date ? value : parseIsoDate_(value);
  const months = [
    'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
  ];
  return date.getDate() + ' ' + months[date.getMonth()] + ' ' + date.getFullYear();
}

function formatDateRange_(startIso, endIso) {
  if (!startIso) return '';
  const start = parseIsoDate_(startIso);
  const end = endIso ? parseIsoDate_(endIso) : start;
  if (end.getTime() < start.getTime()) throw new Error('Tanggal selesai tidak boleh sebelum tanggal mulai.');

  const months = [
    'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
  ];
  if (start.getTime() === end.getTime()) return formatIndonesianDate_(start);
  if (start.getFullYear() === end.getFullYear() && start.getMonth() === end.getMonth()) {
    return start.getDate() + ' - ' + end.getDate() + ' ' +
      months[start.getMonth()] + ' ' + start.getFullYear();
  }
  return start.getDate() + ' ' + months[start.getMonth()] + ' ' + start.getFullYear() +
    ' - ' + end.getDate() + ' ' + months[end.getMonth()] + ' ' + end.getFullYear();
}

function formatDayRange_(startIso, endIso) {
  if (!startIso) return '';
  const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  const start = parseIsoDate_(startIso);
  const end = endIso ? parseIsoDate_(endIso) : start;
  return start.getTime() === end.getTime()
    ? days[start.getDay()]
    : days[start.getDay()] + ' - ' + days[end.getDay()];
}

function serializeValue_(value) {
  if (value instanceof Date) return formatDate_(value, 'yyyy-MM-dd');
  if (Array.isArray(value)) return value.map(serializeValue_);
  if (value && typeof value === 'object') {
    return Object.keys(value).reduce(function(result, key) {
      result[key] = serializeValue_(value[key]);
      return result;
    }, {});
  }
  return value == null ? '' : value;
}

function rowToObject_(headers, row) {
  return headers.reduce(function(result, header, index) {
    result[header] = serializeValue_(row[index]);
    return result;
  }, {});
}

function getJsonCache_(key) {
  try {
    const cached = CacheService.getScriptCache().get(key);
    return cached ? JSON.parse(cached) : null;
  } catch (error) {
    return null;
  }
}

function putJsonCache_(key, value) {
  try {
    CacheService.getScriptCache().put(key, JSON.stringify(value), APP_CONFIG.CACHE_SECONDS);
  } catch (error) {
    // Oversized cache payloads should not break the main request path.
  }
  return value;
}

function findRowById_(sheet, id, idColumn) {
  const column = idColumn || 1;
  const lastRow = lastNonEmptyRowInColumn_(sheet, column);
  if (lastRow < 2) return 0;
  const values = sheet.getRange(2, column, lastRow - 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(id).trim()) return i + 2;
  }
  return 0;
}

function appendDataRow_(sheet, row, idColumn) {
  const column = idColumn || 1;
  const lastUsed = lastNonEmptyRowInColumn_(sheet, column);
  const targetRow = Math.max(2, lastUsed + 1);
  if (targetRow > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), targetRow - sheet.getMaxRows());
  }
  if (row.length > sheet.getMaxColumns()) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), row.length - sheet.getMaxColumns());
  }
  sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  return targetRow;
}

function lastNonEmptyRowInColumn_(sheet, column) {
  const bottom = sheet.getRange(sheet.getMaxRows(), column);
  const candidate = bottom.getNextDataCell(SpreadsheetApp.Direction.UP);
  return text_(candidate.getDisplayValue()) ? candidate.getRow() : 1;
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function escapeRegex_(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function generateRequestId_() {
  const stamp = Utilities.formatDate(new Date(), APP_CONFIG.TIME_ZONE, 'yyyyMMdd');
  const suffix = Utilities.getUuid().replace(/-/g, '').slice(0, 6).toUpperCase();
  return 'DPSP-' + stamp + '-' + suffix;
}

function logAudit_(action, entityId, success, details) {
  try {
    const sheet = getSheet_('AUDIT', false);
    if (!sheet) return;
    appendDataRow_(sheet, [
      Utilities.getUuid(),
      new Date(),
      getCurrentUser_(),
      action,
      entityId || '',
      Boolean(success),
      JSON.stringify(details || {})
    ]);
  } catch (error) {
    console.error('Audit log gagal: ' + error.message);
  }
}

function clearAppCache_() {
  CacheService.getScriptCache().removeAll(['bootstrap', 'references', 'signers', 'employeeCatalog']);
}
