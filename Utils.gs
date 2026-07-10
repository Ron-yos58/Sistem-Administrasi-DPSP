function getSpreadsheet_() {
  const propId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return SpreadsheetApp.openById(propId || APP_CONFIG.SPREADSHEET_ID);
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
  var email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email) {
    email = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  }
  return email;
}

function assertAuthorized_() {
  const email = getCurrentUser_();
  if (!email) {
    throw new Error(
      'Identitas pengguna tidak tersedia. Pastikan Anda telah masuk (login) dengan Google Account dan deploy web app dikonfigurasi dengan benar.'
    );
  }

  const accessSheet = getSheet_('ACCESS', false);
  if (!accessSheet || accessSheet.getLastRow() < 2) {
    throw new Error('Config_Access belum berisi pengguna aktif. Jalankan setupSystem() dengan akun admin awal.');
  }

  const values = accessSheet.getRange(2, 1, accessSheet.getLastRow() - 1, 3).getDisplayValues();
  const match = values.find(function(row) {
    return String(row[0]).trim().toLowerCase() === email;
  });

  if (match) {
    if (!isAccessActive_(match[1])) {
      throw new Error('Akun ' + email + ' telah dinonaktifkan.');
    }
    return { email: email, role: normalizeRole_(match[2]) };
  }

  // Default ke role OPERATOR untuk semua pengguna unlisted agar bisa mengakses sistem secara penuh kecuali menu admin
  return { email: email, role: 'OPERATOR' };
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
  try {
    lock.waitLock(25000);
  } catch (error) {
    throw new Error('Sistem sedang sibuk memproses permintaan lain. Silakan coba beberapa saat lagi.');
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
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return [...new Set(
    source.map(function(item) { return String(item || '').trim().toLowerCase(); })
          .filter(function(item) { return item && re.test(item); })
  )];
}

function uniqueTextList_(values) {
  return [...new Set((values || []).map(function(v) { return text_(v); }).filter(Boolean))];
}

function canonicalOrganizationUnit_(value) {
  const raw = text_(value, 200).trim();
  if (!raw) return '';

  function key_(item) {
    return String(item || '')
      .toLowerCase()
      .replace(/&/g, ' dan ')
      .replace(/[^a-z0-9]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  const aliases = {
    fe: 'Fakultas Ekonomi',
    ff: 'Fakultas Filsafat',
    fh: 'Fakultas Hukum',
    fisip: 'Fakultas Ilmu Sosial dan Ilmu Politik',
    fk: 'Fakultas Kedokteran',
    fkip: 'Fakultas Keguruan dan Ilmu Pendidikan',
    fs: 'Fakultas Sains',
    ft: 'Fakultas Teknik',
    ftr: 'Fakultas Teknologi Rekayasa',
    fv: 'Fakultas Vokasi'
  };
  const rawKey = key_(raw);
  if (aliases[rawKey]) return aliases[rawKey];

  const canonical = (APP_CONFIG.FACULTIES || []).find(function(item) {
    return key_(item) === rawKey;
  });
  return canonical || raw;
}

function isDate_(value) {
  return value && (value instanceof Date || Object.prototype.toString.call(value) === '[object Date]' || (typeof value.getMonth === 'function'));
}

function formatDate_(value, pattern) {
  if (!value) return '';
  const date = isDate_(value) ? value : new Date(value);
  if (isNaN(date.getTime())) return text_(value);
  return Utilities.formatDate(date, APP_CONFIG.TIME_ZONE, pattern || 'yyyy-MM-dd');
}

function parseIsoDate_(value) {
  if (isDate_(value)) return value;
  const clean = text_(value);
  if (!clean) return null;
  const match = clean.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) {
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
  const fallbackDate = new Date(clean);
  if (!isNaN(fallbackDate.getTime())) {
    return fallbackDate;
  }
  throw new Error('Format tanggal harus YYYY-MM-DD: ' + clean);
}

function formatIndonesianDate_(value) {
  if (!value) return '';
  const date = isDate_(value) ? value : parseIsoDate_(value);
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

function joinIndonesian_(values) {
  const items = (values || []).filter(Boolean);
  if (items.length < 2) return items[0] || '';
  if (items.length === 2) return items[0] + ' dan ' + items[1];
  return items.slice(0, -1).join(', ') + ', dan ' + items[items.length - 1];
}

function normalizeTimeValue_(value) {
  if (value == null || value === '') return '';
  if (isDate_(value) && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, APP_CONFIG.TIME_ZONE, 'HH:mm');
  }
  if (typeof value === 'number' && isFinite(value)) {
    const minutes = Math.round(((value % 1) + 1) % 1 * 24 * 60) % (24 * 60);
    return String(Math.floor(minutes / 60)).padStart(2, '0') + ':' +
      String(minutes % 60).padStart(2, '0');
  }

  const input = text_(value);
  const match = input.match(/(?:^|\s)(\d{1,2})[.:](\d{2})(?::\d{2})?/);
  if (!match) return '';
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  if (hour > 23 || minute > 59) return '';
  return String(hour).padStart(2, '0') + ':' + String(minute).padStart(2, '0');
}

function formatTimeRange_(startTime, endTime) {
  const start = normalizeTimeValue_(startTime).replace(':', '.');
  const end = normalizeTimeValue_(endTime).replace(':', '.');
  if (!start && !end) return '';
  if (!start) return end + ' WIB';
  if (!end || end === start) return start + ' WIB';
  return start + '\u2013' + end + ' WIB';
}

function parseTimeRange_(value) {
  if (!value) return { startTime: '', endTime: '' };
  if (isDate_(value) && !isNaN(value.getTime())) {
    return {
      startTime: Utilities.formatDate(value, APP_CONFIG.TIME_ZONE, 'HH:mm'),
      endTime: ''
    };
  }

  const input = text_(value)
    .replace(/\b(?:wib|wita|wit)\b/gi, '')
    .replace(/\bpukul\b/gi, '')
    .replace(/GMT[+-]\d{2}:?\d{2}/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
  const matches = Array.from(input.matchAll(/(\d{1,2})[.:](\d{2})(?::\d{2})?/g));
  if (!matches.length) return { startTime: '', endTime: '' };

  function clock_(hour, minute) {
    const hours = Number(hour);
    const minutes = Number(minute);
    if (hours > 23 || minutes > 59) return '';
    return String(hours).padStart(2, '0') + ':' + String(minutes).padStart(2, '0');
  }

  return {
    startTime: clock_(matches[0][1], matches[0][2]),
    endTime: matches.length > 1
      ? clock_(matches[matches.length - 1][1], matches[matches.length - 1][2])
      : ''
  };
}

function formatScheduleDateList_(items) {
  if (!items.length) return '';
  const singleDates = items.every(function(item) {
    return !item.endDate || item.endDate === item.startDate;
  });
  const first = parseIsoDate_(items[0].startDate);
  const sameMonth = singleDates && items.every(function(item) {
    const date = parseIsoDate_(item.startDate);
    return date.getFullYear() === first.getFullYear() && date.getMonth() === first.getMonth();
  });
  if (!sameMonth) {
    return joinIndonesian_(uniqueTextList_(items.map(function(item) {
      return formatDateRange_(item.startDate, item.endDate || item.startDate);
    })));
  }

  const months = [
    'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
  ];
  const days = uniqueTextList_(items.map(function(item) {
    return String(parseIsoDate_(item.startDate).getDate());
  }));
  return joinIndonesian_(days) + ' ' + months[first.getMonth()] + ' ' + first.getFullYear();
}

function formatScheduleItem_(schedule) {
  const item = schedule || {};
  const startDate = text_(item.startDate);
  const endDate = text_(item.endDate) || startDate;
  const dateLabel = startDate
    ? formatDayRange_(startDate, endDate).replace(/ - /g, '\u2013') + ', ' +
      formatDateRange_(startDate, endDate).replace(/ - /g, '\u2013')
    : '';
  return [
    dateLabel,
    formatTimeRange_(item.startTime, item.endTime),
    text_(item.place)
  ].filter(Boolean).join(' | ');
}

function buildScheduleSummary_(schedules) {
  const sourceItems = (schedules || []).slice();
  const items = sourceItems.filter(function(item) {
    return text_(item.startDate);
  }).sort(function(a, b) {
    return String(a.startDate).localeCompare(String(b.startDate)) ||
      Number(a.sequence || 0) - Number(b.sequence || 0);
  });
  const places = uniqueTextList_(sourceItems.map(function(item) {
    return item.place;
  }).filter(Boolean));
  if (!items.length) {
    return {
      startDate: '',
      endDate: '',
      dayDisplay: '',
      dateDisplay: '',
      timeDisplay: sourceItems.length === 1
        ? formatTimeRange_(sourceItems[0].startTime, sourceItems[0].endTime)
        : '',
      placeDisplay: places.join('\n')
    };
  }

  const dayLabels = uniqueTextList_(items.map(function(item) {
    return formatDayRange_(item.startDate, item.endDate || item.startDate);
  }));
  const timeLabels = items.map(function(item) {
    const time = formatTimeRange_(item.startTime, item.endTime);
    return items.length > 1 && time
      ? formatDateRange_(item.startDate, item.endDate || item.startDate) + ': ' + time
      : time;
  }).filter(Boolean);
  const placeLabels = places.length <= 1
    ? places
    : items.map(function(item) {
        return item.place
          ? formatDateRange_(item.startDate, item.endDate || item.startDate) + ': ' + item.place
          : '';
      }).filter(Boolean);

  return {
    startDate: items[0].startDate,
    endDate: items.reduce(function(latest, item) {
      const candidate = item.endDate || item.startDate;
      return candidate > latest ? candidate : latest;
    }, items[0].endDate || items[0].startDate),
    dayDisplay: joinIndonesian_(dayLabels),
    dateDisplay: formatScheduleDateList_(items),
    timeDisplay: timeLabels.join('\n'),
    placeDisplay: placeLabels.join('\n')
  };
}

function serializeValue_(value) {
  if (isDate_(value)) return formatDate_(value, 'yyyy-MM-dd');
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
    const stringified = JSON.stringify(value);
    if (stringified.length < 100000) {
      CacheService.getScriptCache().put(key, stringified, APP_CONFIG.CACHE_SECONDS);
    } else {
      console.warn('Cache payload too large to store: ' + key + ' (' + stringified.length + ' bytes)');
    }
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
    let detailsStr = JSON.stringify(details || {});
    if (detailsStr.length > 5000) {
      detailsStr = detailsStr.slice(0, 4970) + '...[TRUNCATED_DPSP]';
    }
    appendDataRow_(sheet, [
      Utilities.getUuid(),
      new Date(),
      getCurrentUser_(),
      action,
      entityId || '',
      Boolean(success),
      detailsStr
    ]);
  } catch (error) {
    console.error('Audit log gagal: ' + error.message);
  }
}

const APP_CACHE_KEYS = [
  'bootstrap', 'bootstrap_requests', 'bootstrap_summary',
  'references', 'signers', 'employeeCatalog', 'template_configs'
];

function clearAppCache_() {
  const cache = CacheService.getScriptCache();
  cache.removeAll(APP_CACHE_KEYS);

  try {
    const idsKey = 'cached_request_ids';
    const idsVal = cache.get(idsKey);
    if (idsVal) {
      const ids = JSON.parse(idsVal);
      cache.removeAll(ids.flatMap(function(id) {
        return ['req_detail_int_' + id, 'req_detail_pub_' + id, 'fin_rdns_' + id];
      }));
      cache.remove(idsKey);
    }
  } catch (e) {
    // ignore
  }
}

function trackCachedRequestId_(requestId) {
  try {
    const cache = CacheService.getScriptCache();
    const idsKey = 'cached_request_ids';
    const idsVal = cache.get(idsKey);
    const ids = idsVal ? JSON.parse(idsVal) : [];
    if (ids.indexOf(requestId) === -1) {
      ids.push(requestId);
      if (ids.length > 50) ids.shift();
      cache.put(idsKey, JSON.stringify(ids), 21600);
    }
  } catch (e) {
    // ignore
  }
}


/**
 * Membuat nilai placeholder join pegawai/narasumber untuk template dokumen dan email.
 * Mendukung placeholder baru ({textJoinNomor}) dan placeholder lama Autocrat
 * (<<Text Join Nomor>>, <<Text Join Nama>>, dst.).
 */
function buildEmployeeJoinPlaceholders_(employees) {
  const rows = Array.isArray(employees) ? employees : [];
  const numbers = rows.map(function(_, index) { return String(index + 1); });
  const names = rows.map(function(item) { return text_(item && item.name); });
  const identifiers = rows.map(function(item) { return text_(item && item.identifier); });
  const roles = rows.map(function(item) { return text_(item && item.rank); });
  const units = rows.map(function(item) { return text_(item && item.unit); });
  const emails = rows.map(function(item) { return text_(item && item.email); });

  const narasumber = rows.map(function(item, index) {
    const lines = [(index + 1) + '. ' + text_(item && item.name)];
    if (text_(item && item.identifier)) lines.push('   NIK/NPM: ' + text_(item.identifier));
    if (text_(item && item.rank)) lines.push('   Jabatan: ' + text_(item.rank));
    if (text_(item && item.unit)) lines.push('   Prodi/Unit: ' + text_(item.unit));
    if (text_(item && item.email)) lines.push('   Email: ' + text_(item.email));
    return lines.join('\n');
  }).join('\n\n');

  const values = {
    textJoinNomor: numbers.join('\n'),
    textJoinNama: names.join('\n'),
    textJoinNikNpm: identifiers.join('\n'),
    textJoinJabatan: roles.join('\n'),
    textJoinProdi: units.join('\n'),
    textJoinFakultas: units.join('\n'),
    textJoinEmail: emails.filter(Boolean).join('\n'),
    namaPegawai: names.join('\n'),
    nipPegawai: identifiers.join('\n'),
    jabatanPegawai: roles.join('\n'),
    prodiPegawai: units.join('\n'),
    fakultasPegawai: units.join('\n'),
    emailPegawai: emails.filter(Boolean).join('\n'),
    narasumber: narasumber
  };

  const aliases = {
    'Text Join Nomor': values.textJoinNomor,
    'Text Join Nama': values.textJoinNama,
    'Text Join NIK/NPM': values.textJoinNikNpm,
    'Text Join Jabatan': values.textJoinJabatan,
    'Text Join Prodi': values.textJoinProdi,
    'Text Join Fakultas': values.textJoinFakultas,
    'Text Join Email': values.textJoinEmail,
    'Nomor Urut Pegawai': values.textJoinNomor,
    'Nama Pegawai': values.textJoinNama,
    'NIP/NPM Pegawai': values.textJoinNikNpm,
    'Jabatan Pegawai': values.textJoinJabatan,
    'Prodi/Unit Pegawai': values.textJoinProdi,
    'Email Pegawai': values.textJoinEmail
  };
  Object.keys(aliases).forEach(function(key) { values[key] = aliases[key]; });
  return values;
}

function assertRateLimit_(action, cooldownSeconds) {
  const key = 'rl_' + action + '_' + getCurrentUser_();
  const props = PropertiesService.getUserProperties();
  const last = Number(props.getProperty(key) || 0);
  const now = Date.now();
  if (now - last < cooldownSeconds * 1000) {
    throw new Error('Terlalu banyak permintaan. Mohon tunggu ' + cooldownSeconds + ' detik sebelum mencoba kembali.');
  }
  props.setProperty(key, String(now));
}

function normalizeStatusFromSheet_(s) {
  var clean = String(s || 'DRAFT').toUpperCase().trim();
  if (clean === 'ARCHIVED' || clean === 'ARCHIEVED' || clean === 'SELESAI') return 'ARCHIVED';
  if (clean === 'READY' || clean === 'SIAP DIPROSES') return 'READY';
  return 'DRAFT';
}

function normalizeStatusToSheet_(status) {
  var s = String(status || 'DRAFT').toUpperCase().trim();
  if (s === 'ARCHIVED' || s === 'SELESAI') return 'Selesai';
  if (s === 'READY' || s === 'SIAP DIPROSES') return 'Siap Diproses';
  return 'Draft';
}

function isArchivedStatus_(status) {
  return normalizeStatusFromSheet_(status) === 'ARCHIVED';
}
