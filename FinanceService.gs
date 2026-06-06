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
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
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
  if (request.activityType === 'Penugasan Narasumber') {
    renderSpeakerHonorSheet_(sheet, detail);
    return;
  }
  renderEduCampusHonorSheet_(sheet, detail);
}

function renderTravelSheet_(sheet, detail) {
  const request = detail.request;
  const travel = detail.travel;
  sheet.clearFormats();
  sheet.clearNotes();
  sheet.setHiddenGridlines(false);
  if (!travel.length) {
    sheet.getRange('A1:F1').merge().setValue('PERMOHONAN PENCAIRAN PERJALANAN DINAS')
      .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
    sheet.getRange('A3:F3').merge().setValue('Belum ada data perjadin untuk permohonan ini.')
      .setHorizontalAlignment('center');
    return;
  }

  setTravelSheetColumnWidths_(sheet);
  let startRow = 1;
  travel.forEach(function(item, index) {
    startRow = renderTravelParticipantSection_(sheet, startRow, request, item, index + 1, travel.length);
  });
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

function renderEduCampusHonorSheet_(sheet, detail) {
  const request = detail.request;
  const employees = detail.employees;
  const formulaSeparator = financeFormulaSeparator_();
  sheet.clearFormats();
  sheet.clearNotes();

  setMergedRichLabel_(sheet, 1, 1, 6, 'Daftar Honor: ', request.activityName || '');
  setMergedRichLabel_(sheet, 2, 1, 6, 'Tipe Kegiatan: ', request.activityType || '');
  setMergedRichLabel_(sheet, 3, 1, 6, 'Nama Mitra: ', request.partnerName || '');
  setMergedRichLabel_(sheet, 4, 1, 6, 'Alamat Kegiatan: ', request.activityPlace || '');
  setMergedRichLabel_(sheet, 5, 1, 6, 'Tanggal Kegiatan: ', request.dateDisplay || '');

  const headerRow = 8;
  const headers = [['No.', 'Nama Pegawai / Mahasiswa', 'NIP/NPM', 'Jam Mulai', 'Jam Akhir', 'Honor', 'Transport', 'Uang Makan', 'Jumlah']];
  sheet.getRange(headerRow, 1, 1, headers[0].length)
    .setValues(headers)
    .setFontWeight('bold')
    .setBackground('#d9d9d9')
    .setBorder(true, true, true, true, true, true)
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 120);
  sheet.setColumnWidth(8, 120);
  sheet.setColumnWidth(9, 120);

  const startRow = headerRow + 1;
  employees.forEach(function(employee, index) {
    const rowNumber = startRow + index;
    sheet.getRange(rowNumber, 1, 1, 5).setValues([[
      index + 1,
      employee.name || '',
      employee.identifier || '',
      '',
      ''
    ]]);
    sheet.getRange(rowNumber, 6).setFormula(buildEduCampusHonorFormula_(rowNumber, formulaSeparator));
    sheet.getRange(rowNumber, 9).setFormula('=F' + rowNumber + '+G' + rowNumber + '+H' + rowNumber);
    sheet.getRange(rowNumber, 1, 1, 9)
      .setBorder(true, true, true, true, true, true)
      .setWrap(true)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.getRange(rowNumber, 2).setHorizontalAlignment('left');
    sheet.getRange(rowNumber, 6, 1, 4).setNumberFormat('"Rp"#,##0').setHorizontalAlignment('right');
  });

  if (employees.length) {
    setHourDropdown_(sheet, startRow, employees.length);
  }

  const lastDataRow = employees.length ? startRow + employees.length - 1 : startRow - 1;
  const totalRow = lastDataRow + 2;
  sheet.getRange(totalRow, 1, 1, 8).merge()
    .setValue('TOTAL')
    .setFontWeight('bold')
    .setBorder(true, true, true, false, true, true)
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');
  sheet.getRange(totalRow, 9)
    .setFormula(employees.length ? '=SUM(I' + startRow + ':I' + lastDataRow + ')' : '=0')
    .setFontWeight('bold')
    .setBackground('#d9d9d9')
    .setBorder(true, true, true, true, true, true)
    .setNumberFormat('"Rp"#,##0')
    .setHorizontalAlignment('right');

  addSignatureBlock_(sheet, totalRow + 3, 7, 3, request);
}

function renderSpeakerHonorSheet_(sheet, detail) {
  const request = detail.request;
  const employees = detail.employees;
  const formulaSeparator = financeFormulaSeparator_();
  sheet.clearFormats();
  sheet.clearNotes();

  setMergedRichLabel_(sheet, 1, 1, 7, 'Daftar Honor: ', request.activityName || '');
  setMergedRichLabel_(sheet, 2, 1, 7, 'Tipe Kegiatan: ',
    [request.activityType, request.speakerSubtype].filter(Boolean).join(' - '));
  setMergedRichLabel_(sheet, 3, 1, 7, 'Nama Mitra: ', request.partnerName || '');
  setMergedRichLabel_(sheet, 4, 1, 7, 'Alamat Kegiatan: ', request.activityPlace || '');
  setMergedRichLabel_(sheet, 5, 1, 7, 'Tanggal Kegiatan: ', request.dateDisplay || '');

  const headerRow = 8;
  const headers = [['No.', 'Nama Pegawai', 'NIP', 'Makalah', 'Honor', 'Transport', 'Jumlah']];
  sheet.getRange(headerRow, 1, 1, headers[0].length)
    .setValues(headers)
    .setFontWeight('bold')
    .setBackground('#d9d9d9')
    .setBorder(true, true, true, true, true, true)
    .setWrap(true)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 164);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 120);

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Dengan Makalah', 'Tidak Dengan Makalah'], true)
    .setAllowInvalid(false)
    .build();
  const startRow = headerRow + 1;
  employees.forEach(function(employee, index) {
    const rowNumber = startRow + index;
    sheet.getRange(rowNumber, 1, 1, 7).setValues([[
      index + 1,
      employee.name || '',
      employee.identifier || '',
      'Dengan Makalah',
      '',
      '',
      ''
    ]]);
    sheet.getRange(rowNumber, 4).setDataValidation(rule);
    sheet.getRange(rowNumber, 5).setFormula(buildSpeakerHonorFormula_(rowNumber, formulaSeparator));
    sheet.getRange(rowNumber, 7).setFormula('=E' + rowNumber + '+F' + rowNumber);
    sheet.getRange(rowNumber, 1, 1, 7)
      .setBorder(true, true, true, true, true, true)
      .setWrap(true)
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.getRange(rowNumber, 2).setHorizontalAlignment('left');
    sheet.getRange(rowNumber, 5, 1, 3).setNumberFormat('"Rp"#,##0').setHorizontalAlignment('right');
  });

  const lastDataRow = employees.length ? startRow + employees.length - 1 : startRow - 1;
  const totalRow = lastDataRow + 2;
  sheet.getRange(totalRow, 1, 1, 6).merge()
    .setValue('TOTAL')
    .setFontWeight('bold')
    .setBorder(true, true, true, false, true, true)
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');
  sheet.getRange(totalRow, 7)
    .setFormula(employees.length ? '=SUM(G' + startRow + ':G' + lastDataRow + ')' : '=0')
    .setFontWeight('bold')
    .setBackground('#d9d9d9')
    .setBorder(true, true, true, true, true, true)
    .setNumberFormat('"Rp"#,##0')
    .setHorizontalAlignment('right');

  addSignatureBlock_(sheet, totalRow + 3, 5, 3, request);
}

function renderTravelParticipantSection_(sheet, startRow, request, item, sequence, totalSections) {
  const titleRow = startRow;
  const componentLabels = [
    'Uang kegiatan', 'Uang saku', 'Uang makan', 'Uang penginapan',
    'Uang transportasi dalam kota', 'Uang transportasi antar kota',
    'Uang transportasi antar negara pp', 'Uang aplikasi visa',
    'Uang asuransi perjalanan', 'Uang fiskal & pajak bandara', 'Uang harian'
  ];
  const requestDate = request.letterDate ? formatIndonesianDate_(request.letterDate) : '';
  const officialDays = countTravelDays_(request.startDate, request.endDate);

  sheet.getRange(titleRow, 1, 1, 6).merge()
    .setValue('PERMOHONAN PENCAIRAN PERJALANAN DINAS')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontSize(16);
  sheet.getRange(titleRow + 1, 1, 1, 6).merge()
    .setValue('Tanggal: ' + requestDate)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  const detailLabels = [
    'Nama',
    'NIP/NPM',
    'Surat Tugas',
    'Dosen/Pangkat Penunjang',
    'Kategori penerima tugas',
    'Berangkat ke',
    'Tanggal Kegiatan',
    'Tanggal Berangkat',
    'Untuk keperluan',
    'Jml. Hari kegiatan resmi',
    'Tambahan hari lamanya perjalanan'
  ];
  const detailValues = [
    item.name || '',
    item.identifier || '',
    item.letterNumber || '',
    item.rank || '',
    item.category || '',
    item.place || request.activityPlace || '',
    request.dateDisplay || '',
    '',
    request.activityType + (request.speakerSubtype ? ' - ' + request.speakerSubtype : ''),
    officialDays ? officialDays + ' hari' : '',
    ''
  ];

  for (let index = 0; index < detailLabels.length; index++) {
    const row = titleRow + 3 + index;
    sheet.getRange(row, 1, 1, 2).merge().setValue(detailLabels[index]).setFontWeight('bold');
    sheet.getRange(row, 3, 1, 4).merge()
      .setValue(detailValues[index] ? ': ' + detailValues[index] : ': ')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle');
    if (index === 7 || index === 10) {
      sheet.getRange(row, 3, 1, 4).setBackground('#d9d9d9');
    }
  }

  const tableStartRow = titleRow + 15;
  sheet.getRange(tableStartRow, 1, 1, 4).setValues([[
    'No', 'Komponen Uang Perjalanan Dinas', 'Jumlah yang disetujui', 'Keterangan'
  ]])
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#d9d9d9')
    .setBorder(true, true, true, true, true, true)
    .setWrap(true);
  sheet.getRange(tableStartRow, 4, 1, 3).merge();

  componentLabels.forEach(function(label, idx) {
    const row = tableStartRow + 1 + idx;
    sheet.getRange(row, 1, 1, 5).setBorder(true, true, true, true, true, true);
    sheet.getRange(row, 1).setValue(idx + 1).setHorizontalAlignment('center');
    sheet.getRange(row, 2).setValue(label);
    sheet.getRange(row, 3).setValue(Number(item.costs[idx] || 0)).setBackground('#d9d9d9').setNumberFormat('"Rp"#,##0');
    sheet.getRange(row, 4, 1, 3).merge().setBackground('#d9d9d9');
  });

  const totalRow = tableStartRow + 1 + componentLabels.length;
  sheet.getRange(totalRow, 1, 1, 2).merge()
    .setValue('TOTAL')
    .setFontWeight('bold')
    .setHorizontalAlignment('right')
    .setBackground('#d9d9d9')
    .setBorder(true, true, true, true, true, true)
    .setWrap(true);
  sheet.getRange(totalRow, 3)
    .setFormula('=SUM(C' + (tableStartRow + 1) + ':C' + (totalRow - 1) + ')')
    .setFontWeight('bold')
    .setNumberFormat('"Rp"#,##0')
    .setHorizontalAlignment('right')
    .setBackground('#d9d9d9')
    .setBorder(true, true, true, true, true, true);
  sheet.getRange(totalRow, 4, 1, 3).merge().setBackground('#d9d9d9').setBorder(true, true, true, true, true, true);

  const signatureRow = totalRow + 3;
  sheet.getRange(signatureRow, 1, 1, 2).merge().setValue('Bandung, ' + requestDate).setHorizontalAlignment('left').setWrap(true);
  sheet.getRange(signatureRow, 3, 1, 2).merge().setValue('Tanggal,').setHorizontalAlignment('left').setWrap(true);
  sheet.getRange(signatureRow, 5, 1, 2).merge().setValue('Pemeriksa').setHorizontalAlignment('left').setWrap(true);
  sheet.getRange(signatureRow + 5, 1, 1, 2).merge()
    .setValue(request.signerName || 'Penandatangan')
    .setHorizontalAlignment('left').setWrap(true).setFontWeight('bold');
  sheet.getRange(signatureRow + 5, 3, 1, 2).merge()
    .setValue('Kepala BIKEU')
    .setHorizontalAlignment('left').setWrap(true).setFontWeight('bold');
  sheet.getRange(signatureRow + 5, 5, 1, 2).merge()
    .setValue('Staff BIKEU')
    .setHorizontalAlignment('left').setWrap(true).setFontWeight('bold');
  sheet.getRange(signatureRow + 6, 1, 1, 2).merge()
    .setValue(request.signerRole || '')
    .setHorizontalAlignment('left').setWrap(true);
  sheet.getRange(signatureRow + 6, 3, 1, 2).merge()
    .setValue('Jabatan Kepala BIKEU')
    .setHorizontalAlignment('left').setWrap(true);
  sheet.getRange(signatureRow + 6, 5, 1, 2).merge()
    .setValue('Jabatan Staff BIKEU')
    .setHorizontalAlignment('left').setWrap(true);

  const table2StartRow = signatureRow + 9;
  sheet.getRange(table2StartRow, 1, 1, 3).setValues([[
    'No', 'Komponen Uang Perjalanan Dinas', 'Pencairan sebelum Pelaksanaan Tugas '
  ]])
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#d9d9d9')
    .setBorder(true, true, true, true, true, true)
    .setWrap(true);
  componentLabels.forEach(function(label, idx) {
    const row = table2StartRow + 1 + idx;
    sheet.getRange(row, 1, 1, 3).setBorder(true, true, true, true, true, true);
    sheet.getRange(row, 1).setValue(idx + 1).setHorizontalAlignment('center');
    sheet.getRange(row, 2).setValue(label);
    sheet.getRange(row, 3).setValue(Number(item.costs[idx] || 0)).setBackground('#d9d9d9').setNumberFormat('"Rp"#,##0');
  });
  const totalRow2 = table2StartRow + 1 + componentLabels.length;
  sheet.getRange(totalRow2, 1, 1, 2).merge()
    .setValue('TOTAL')
    .setFontWeight('bold')
    .setHorizontalAlignment('right')
    .setBackground('#d9d9d9')
    .setBorder(true, true, true, true, true, true)
    .setWrap(true);
  sheet.getRange(totalRow2, 3)
    .setFormula('=SUM(C' + (table2StartRow + 1) + ':C' + (totalRow2 - 1) + ')')
    .setFontWeight('bold')
    .setNumberFormat('"Rp"#,##0')
    .setHorizontalAlignment('right')
    .setBackground('#d9d9d9')
    .setBorder(true, true, true, true, true, true)
    .setWrap(true);

  const receiverRow = table2StartRow + 29;
  sheet.getRange(receiverRow, 5, 1, 2).merge().setValue('Tanggal, ').setHorizontalAlignment('left').setWrap(true);
  sheet.getRange(receiverRow + 5, 5, 1, 2).merge().setValue('Penerima').setHorizontalAlignment('left').setWrap(true);
  sheet.getRange(receiverRow + 6, 5, 1, 2).merge()
    .setValue(item.name || '')
    .setHorizontalAlignment('left')
    .setWrap(true)
    .setFontWeight('bold');

  const table3StartRow = totalRow2 + 3;
  sheet.getRange(table3StartRow, 1, 1, 5).setValues([[
    'No',
    'Komponen Uang Perjalanan Dinas',
    'Pencairan Setelah Pelaksanaan Tugas',
    'Penggantian*)',
    'Reimbursement/Tanpa Penggantian*)'
  ]])
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#d9d9d9')
    .setBorder(true, true, true, true, true, true)
    .setWrap(true);
  sheet.getRange(table3StartRow, 5, 1, 2).merge()
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setFontWeight('bold')
    .setBackground('#d9d9d9');
  componentLabels.forEach(function(label, idx) {
    const row = table3StartRow + 1 + idx;
    sheet.getRange(row, 1, 1, 6).setBorder(true, true, true, true, true, true);
    sheet.getRange(row, 1).setValue(idx + 1).setHorizontalAlignment('center');
    sheet.getRange(row, 2).setValue(label);
    sheet.getRange(row, 3).setValue(0).setBackground('#d9d9d9').setNumberFormat('"Rp"#,##0');
    sheet.getRange(row, 4).setBackground('#d9d9d9');
    sheet.getRange(row, 5, 1, 2).merge().setBackground('#d9d9d9');
  });
  const totalRow3 = table3StartRow + 1 + componentLabels.length;
  sheet.getRange(totalRow3, 1, 1, 2).merge()
    .setValue('TOTAL')
    .setFontWeight('bold')
    .setHorizontalAlignment('right')
    .setBackground('#d9d9d9')
    .setBorder(true, true, true, true, true, true)
    .setWrap(true);
  sheet.getRange(totalRow3, 3)
    .setFormula('=SUM(C' + (table3StartRow + 1) + ':C' + (totalRow3 - 1) + ')')
    .setFontWeight('bold')
    .setNumberFormat('"Rp"#,##0')
    .setHorizontalAlignment('right')
    .setBackground('#d9d9d9')
    .setBorder(true, true, true, true, true, true)
    .setWrap(true);
  sheet.getRange(totalRow3, 4).setBackground('#d9d9d9').setBorder(true, true, true, true, true, true);
  sheet.getRange(totalRow3, 5, 1, 2).merge().setBackground('#d9d9d9').setBorder(true, true, true, true, true, true);
  sheet.getRange(totalRow3 + 1, 1, 1, 2).merge()
    .setValue('*) diisi setelah pelaksanaan tugas')
    .setHorizontalAlignment('left')
    .setWrap(true)
    .setFontSize(10)
    .setFontStyle('italic');
  sheet.getRange(totalRow3 + 2, 1, 1, 2).merge()
    .setValue('**) untuk perjalanan dinas ke luar negeri')
    .setHorizontalAlignment('left')
    .setWrap(true)
    .setFontSize(10)
    .setFontStyle('italic');

  if (sequence < totalSections) {
    sheet.getRange(totalRow3 + 4, 1, 1, 6).merge().setValue('');
  }
  return totalRow3 + 6;
}

function setTravelSheetColumnWidths_(sheet) {
  sheet.setColumnWidth(1, 40);
  sheet.setColumnWidth(2, 270);
  sheet.setColumnWidth(3, 165);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 100);
  sheet.setColumnWidth(6, 100);
}

function addSignatureBlock_(sheet, startRow, startColumn, width, request) {
  const signDate = request.letterDate ? formatIndonesianDate_(request.letterDate) : '';
  sheet.getRange(startRow, startColumn, 1, width).merge()
    .setValue('Bandung, ' + signDate)
    .setWrap(true);
  sheet.getRange(startRow + 5, startColumn, 1, width).merge()
    .setValue(request.signerName || 'Penandatangan')
    .setWrap(true)
    .setFontWeight('bold');
  sheet.getRange(startRow + 6, startColumn, 1, width).merge()
    .setValue(request.signerRole || '')
    .setWrap(true);
}

function setMergedRichLabel_(sheet, row, startColumn, width, label, value) {
  const range = sheet.getRange(row, startColumn, 1, width);
  range.merge();
  range.setWrap(true);
  const text = String(label || '') + String(value || '');
  const richText = SpreadsheetApp.newRichTextValue()
    .setText(text)
    .setTextStyle(0, String(label || '').length, SpreadsheetApp.newTextStyle().setBold(true).build())
    .build();
  range.setRichTextValue(richText);
}

function financeFormulaSeparator_() {
  const ss = getSpreadsheet_();
  const locale = typeof ss.getSpreadsheetLocale === 'function'
    ? ss.getSpreadsheetLocale()
    : Session.getActiveUserLocale();
  return locale && String(locale).toLowerCase().indexOf('en') === 0 ? ',' : ';';
}

function buildEduCampusHonorFormula_(rowNumber, separator) {
  const subEnd = 'SUBSTITUTE(E' + rowNumber + separator + '"."' + separator + '":")';
  const subStart = 'SUBSTITUTE(D' + rowNumber + separator + '"."' + separator + '":")';
  const timeDiff = 'MOD(TIMEVALUE(' + subEnd + ')-TIMEVALUE(' + subStart + ')' + separator + '1)*24*9000';
  const andPart = 'AND(LEN(D' + rowNumber + ')' + separator + 'LEN(E' + rowNumber + '))';
  return '=IF(' + andPart + separator + timeDiff + separator + '0)';
}

function buildSpeakerHonorFormula_(rowNumber, separator) {
  return '=IF(D' + rowNumber + '="Dengan Makalah"' + separator + '1000000' + separator +
    'IF(D' + rowNumber + '="Tidak Dengan Makalah"' + separator + '750000' + separator + '0))';
}

function setHourDropdown_(sheet, startRow, rowCount) {
  const times = [];
  for (let hour = 0; hour < 24; hour++) {
    times.push((hour < 10 ? '0' + hour : String(hour)) + ':00');
  }
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(times, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(startRow, 4, rowCount, 1).setDataValidation(rule);
  sheet.getRange(startRow, 5, rowCount, 1).setDataValidation(rule);
}

function countTravelDays_(startDate, endDate) {
  const startIso = text_(startDate);
  if (!startIso) return 0;
  const start = parseIsoDate_(startIso);
  const end = endDate ? parseIsoDate_(text_(endDate)) : start;
  return Math.max(1, Math.floor((end.getTime() - start.getTime()) / (24 * 60 * 60 * 1000)) + 1);
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
