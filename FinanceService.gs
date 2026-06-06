function listTravelData(filters) {
  assertAuthorized_();
  const query = text_((filters || {}).query).toLowerCase();
  const rows = readDataRows_(getSheet_('TRAVEL'), TRAVEL_HEADERS.length)
    .map(travelRowToDto_)
    .filter(function(item) {
      if (!query) return true;
      return [item.requestId, item.name, item.identifier, item.place]
        .join(' ').toLowerCase().indexOf(query) !== -1;
    });
  return serializeValue_({ items: rows, total: rows.length });
}

function syncTravelData(requestId) {
  const user = assertAuthorized_();
  assertCanWrite_(user);
  return withScriptLock_(function() {
    const result = syncTravelDataInternal_(text_(requestId));
    logAudit_('SYNC_TRAVEL', requestId || 'ALL', true, result);
    return result;
  });
}

function syncTravelDataInternal_(requestId) {
  const masterSheet = getSheet_('MASTER');
  const employeeSheet = getSheet_('EMPLOYEES');
  const travelSheet = getSheet_('TRAVEL');
  const existingRows = readDataRows_(travelSheet, TRAVEL_HEADERS.length);
  const existingByKey = {};

  existingRows.forEach(function(row) {
    const key = text_(row[19]) || inferParticipantKey_(row[0], row[2], row[1]);
    row[19] = key;
    existingByKey[key] = row;
  });

  const masterRows = readDataRows_(masterSheet, MASTER_HEADERS.length)
    .map(masterRowToDto_)
    .filter(function(item) {
      return item.id &&
        item.status !== 'ARCHIVED' &&
        item.travel === 'Ya' &&
        (!requestId || item.id === requestId);
    });
  const masterById = {};
  masterRows.forEach(function(item) { masterById[item.id] = item; });

  const targetRows = readDataRows_(employeeSheet, EMPLOYEE_HEADERS.length)
    .filter(function(row) { return Boolean(masterById[text_(row[0])]); })
    .map(function(employeeRow) {
      const request = masterById[text_(employeeRow[0])];
      const key = text_(employeeRow[8]) ||
        inferParticipantKey_(request.id, employeeRow[2], employeeRow[1]);
      const old = existingByKey[key];
      const costs = old ? old.slice(8, 19) : new Array(11).fill(0);
      return [
        request.id,
        employeeRow[1],
        employeeRow[2],
        firstDocumentNumber_(request.id, 'Surat Tugas'),
        employeeRow[6],
        employeeRow[7],
        request.dateDisplay,
        request.activityPlace
      ].concat(costs).concat([key]);
    });

  let finalRows;
  if (requestId) {
    const untouched = existingRows.filter(function(row) { return text_(row[0]) !== requestId; });
    finalRows = untouched.concat(targetRows);
  } else {
    finalRows = targetRows;
  }
  rewriteDataRows_(travelSheet, finalRows, TRAVEL_HEADERS.length);
  return { ok: true, synced: targetRows.length };
}

function saveTravelCosts(payload) {
  const user = assertAuthorized_();
  assertCanWrite_(user);
  const participantKey = text_((payload || {}).participantKey);
  const costs = (payload || {}).costs || [];
  if (!participantKey) throw new Error('Participant Key wajib diisi.');
  if (costs.length !== 11) throw new Error('Jumlah komponen biaya harus 11.');

  const normalized = costs.map(function(value) {
    const number = Number(value || 0);
    if (!isFinite(number) || number < 0 || number > 1000000000000) {
      throw new Error('Nominal biaya tidak valid.');
    }
    return number;
  });

  return withScriptLock_(function() {
    const sheet = getSheet_('TRAVEL');
    const rows = readDataRows_(sheet, TRAVEL_HEADERS.length);
    const index = rows.findIndex(function(row) { return text_(row[19]) === participantKey; });
    if (index === -1) throw new Error('Data perjadin tidak ditemukan. Jalankan sinkronisasi.');
    for (let i = 0; i < normalized.length; i++) rows[index][8 + i] = normalized[i];
    rewriteDataRows_(sheet, rows, TRAVEL_HEADERS.length);
    logAudit_('SAVE_TRAVEL_COSTS', rows[index][0], true, { participantKey: participantKey });
    return { ok: true, item: travelRowToDto_(rows[index]) };
  });
}

function getTravelByRequestInternal_(requestId) {
  return readDataRows_(getSheet_('TRAVEL'), TRAVEL_HEADERS.length)
    .filter(function(row) { return text_(row[0]) === requestId; })
    .map(travelRowToDto_);
}

function travelRowToDto_(row) {
  const costs = row.slice(8, 19).map(function(value) { return Number(value || 0); });
  return {
    requestId: row[0],
    name: row[1],
    identifier: row[2],
    letterNumber: row[3],
    rank: row[4],
    category: row[5],
    date: row[6],
    place: row[7],
    costs: costs,
    total: costs.reduce(function(sum, value) { return sum + value; }, 0),
    participantKey: row[19]
  };
}

function generateFinanceSheet(requestId, kind) {
  const user = assertAuthorized_();
  assertCanWrite_(user);
  const id = text_(requestId);
  const reportKind = text_(kind).toUpperCase();
  if (['HONOR', 'PERJADIN'].indexOf(reportKind) === -1) {
    throw new Error('Jenis laporan harus HONOR atau PERJADIN.');
  }

  return withScriptLock_(function() {
    return generateFinanceSheetInternal_(id, reportKind);
  });
}

function generateFinanceSheetInternal_(requestId, reportKind) {
  const detail = getRequestDetailInternal_(requestId);
  if (reportKind === 'HONOR' && detail.request.honor !== 'Ya') {
    throw new Error('Permohonan ini tidak ditandai memiliki honor.');
  }
  if (reportKind === 'PERJADIN' && detail.request.travel !== 'Ya') {
    throw new Error('Permohonan ini tidak ditandai memiliki perjalanan dinas.');
  }

  const ss = getSpreadsheet_();
  const sheetName = generatedSheetName_(reportKind, requestId);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.addDeveloperMetadata('DPSP_GENERATED', reportKind + '|' + requestId);
  } else {
    sheet.getDataRange().breakApart();
    sheet.clear();
  }

  if (reportKind === 'HONOR') {
    renderHonorSheet_(sheet, detail);
  } else {
    renderTravelSheet_(sheet, detail);
  }
  logAudit_('GENERATE_FINANCE_SHEET', requestId, true, {
    kind: reportKind,
    sheetId: sheet.getSheetId()
  });
  return {
    ok: true,
    requestId: requestId,
    kind: reportKind,
    sheetName: sheetName,
    sheetId: sheet.getSheetId(),
    url: ss.getUrl() + '#gid=' + sheet.getSheetId()
  };
}

function previewGeneratedSheetCleanup() {
  const user = assertAuthorized_();
  if (user.role !== 'ADMIN') throw new Error('Hanya Admin dapat membersihkan sheet keluaran.');
  const sheets = getSpreadsheet_().getSheets().filter(isGeneratedSheet_);
  return sheets.map(function(sheet) {
    return { name: sheet.getName(), sheetId: sheet.getSheetId() };
  });
}

function deleteGeneratedSheets(sheetNames) {
  const user = assertAuthorized_();
  if (user.role !== 'ADMIN') throw new Error('Hanya Admin dapat membersihkan sheet keluaran.');
  const targets = uniqueTextList_(sheetNames || []);

  return withScriptLock_(function() {
    const ss = getSpreadsheet_();
    const deleted = [];
    targets.forEach(function(name) {
      const sheet = ss.getSheetByName(name);
      if (!sheet || !isGeneratedSheet_(sheet)) {
        throw new Error('Sheet bukan artefak sistem dan tidak boleh dihapus: ' + name);
      }
      ss.deleteSheet(sheet);
      deleted.push(name);
    });
    logAudit_('DELETE_GENERATED_SHEETS', '', true, { deleted: deleted });
    return { ok: true, deleted: deleted };
  });
}

function renderHonorSheet_(sheet, detail) {
  const request = detail.request;
  const employees = detail.employees;
  sheet.getRange('A1:G1').merge().setValue('DAFTAR HONOR')
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 1, 4, 2).setValues([
    ['Kegiatan', request.activityName],
    ['Tipe', request.activityType + (request.speakerSubtype ? ' - ' + request.speakerSubtype : '')],
    ['Mitra', request.partnerName],
    ['Tanggal', request.dateDisplay]
  ]);
  const headerRow = 8;
  sheet.getRange(headerRow, 1, 1, 7).setValues([[
    'No.', 'Nama', 'NIP/NPM', 'Kategori', 'Honor', 'Transport', 'Jumlah'
  ]]);
  const rows = employees.map(function(employee, index) {
    return [index + 1, employee.name, employee.identifier, employee.category, 0, 0, 0];
  });
  if (rows.length) {
    sheet.getRange(headerRow + 1, 1, rows.length, 7).setValues(rows);
    rows.forEach(function(_, index) {
      const row = headerRow + 1 + index;
      sheet.getRange(row, 7).setFormula('=E' + row + '+F' + row);
    });
  }
  styleFinanceSheet_(sheet, headerRow, Math.max(rows.length, 1), 7);
}

function renderTravelSheet_(sheet, detail) {
  const request = detail.request;
  const travel = detail.travel;
  sheet.getRange('A1:F1').merge().setValue('RINGKASAN PERJALANAN DINAS')
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange(2, 1, 4, 2).setValues([
    ['Kegiatan', request.activityName],
    ['Mitra', request.partnerName],
    ['Tanggal', request.dateDisplay],
    ['Tempat', request.activityPlace]
  ]);
  const headerRow = 8;
  sheet.getRange(headerRow, 1, 1, 6).setValues([[
    'No.', 'Nama', 'NIP/NPM', 'Pangkat', 'Kategori', 'Total'
  ]]);
  const rows = travel.map(function(item, index) {
    return [index + 1, item.name, item.identifier, item.rank, item.category, item.total];
  });
  if (rows.length) sheet.getRange(headerRow + 1, 1, rows.length, 6).setValues(rows);
  styleFinanceSheet_(sheet, headerRow, Math.max(rows.length, 1), 6);
  if (rows.length) sheet.getRange(headerRow + 1, 6, rows.length, 1).setNumberFormat('"Rp"#,##0');
}

function styleFinanceSheet_(sheet, headerRow, rowCount, width) {
  sheet.getRange(headerRow, 1, 1, width)
    .setBackground('#015850').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setWrap(true);
  sheet.getRange(headerRow, 1, rowCount + 1, width)
    .setBorder(true, true, true, true, true, true);
  sheet.setFrozenRows(headerRow);
  sheet.autoResizeColumns(1, width);
}

function generatedSheetName_(kind, requestId) {
  const safeId = requestId.replace(/[^A-Za-z0-9_-]/g, '-').slice(-40);
  return ('GEN-' + kind + '-' + safeId).slice(0, 99);
}

function isGeneratedSheet_(sheet) {
  if (sheet.getName().indexOf('GEN-') !== 0) return false;
  return sheet.getDeveloperMetadata().some(function(metadata) {
    return metadata.getKey() === 'DPSP_GENERATED';
  });
}

function inferParticipantKey_(requestId, identifier, name) {
  return buildParticipantKey_(requestId, {
    identifier: identifier,
    name: name,
    email: ''
  });
}

function firstDocumentNumber_(requestId, type) {
  const document = getDocumentsByRequest_(requestId).find(function(item) {
    return item.type === type;
  });
  return document ? document.number : '';
}

function getRequestDetailInternal_(requestId) {
  const master = getSheet_('MASTER');
  const rowNumber = findRowById_(master, requestId, 1);
  if (!rowNumber) throw new Error('Permohonan tidak ditemukan: ' + requestId);
  const row = master.getRange(rowNumber, 1, 1, MASTER_HEADERS.length).getValues()[0];
  return {
    request: masterRowToDto_(row),
    documents: getDocumentsByRequest_(requestId),
    employees: getEmployeesByRequest_(requestId),
    travel: getTravelByRequestInternal_(requestId)
  };
}
