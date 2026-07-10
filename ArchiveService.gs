function previewYearlyArchive(year) {
  const user = assertAuthorized_();
  assertAdmin_(user);
  return serializeValue_(buildYearlyArchivePlan_(year));
}

function archivePreviousYearData(year, confirmText) {
  const user = assertAuthorized_();
  assertAdmin_(user);
  const targetYear = resolveArchiveYear_(year);
  if (text_(confirmText) !== 'ARCHIVE-' + targetYear) {
    throw new Error('Konfirmasi arsip harus tepat: ARCHIVE-' + targetYear);
  }

  return serializeValue_(withScriptLock_(function() {
    const plan = buildYearlyArchivePlan_(targetYear);
    if (!plan.eligibleRequests.length) {
      logAudit_('ARCHIVE_YEARLY_DATA', String(targetYear), true, {
        year: targetYear,
        movedRequests: 0,
        message: 'Tidak ada permohonan selesai untuk diarsipkan.'
      });
      return {
        ok: true,
        year: targetYear,
        movedRequests: 0,
        message: 'Tidak ada permohonan selesai untuk diarsipkan.',
        archiveSpreadsheetUrl: plan.archiveSpreadsheetUrl,
        archiveSpreadsheetId: plan.archiveSpreadsheetId
      };
    }

    const archiveSpreadsheet = openArchiveSpreadsheet_(targetYear);
    const idsToWrite = plan.missingRequestIds;
    if (idsToWrite.length) {
      appendArchiveRows_(archiveSpreadsheet, 'MASTER', MASTER_HEADERS, plan.masterRows, 0);
      appendArchiveRows_(archiveSpreadsheet, 'DOCUMENTS', DOCUMENT_HEADERS, plan.documentRows, 0);
      appendArchiveRows_(archiveSpreadsheet, 'SCHEDULES', SCHEDULE_HEADERS, plan.scheduleRows, 1);
      appendArchiveRows_(archiveSpreadsheet, 'EMPLOYEES', EMPLOYEE_HEADERS, plan.employeeRows, 0);
      appendArchiveRows_(archiveSpreadsheet, 'TRAVEL', TRAVEL_HEADERS, plan.travelRows, 0);
      appendArchiveRows_(archiveSpreadsheet, 'FILES', GENERATED_FILE_HEADERS, plan.fileRows, 0);
    }

    const verified = buildYearlyArchivePlan_(targetYear);
    if (verified.missingRequestIds.length) {
      throw new Error('Verifikasi arsip gagal. Masih ada permohonan yang belum tersalin: ' + verified.missingRequestIds.join(', '));
    }

    removeArchivedRowsFromSource_(verified.eligibleRequestIds);
    clearAppCache_();
    logAudit_('ARCHIVE_YEARLY_DATA', String(targetYear), true, {
      year: targetYear,
      movedRequests: verified.eligibleRequestIds.length,
      archiveSpreadsheetId: verified.archiveSpreadsheetId,
      archiveSpreadsheetUrl: verified.archiveSpreadsheetUrl
    });

    return {
      ok: true,
      year: targetYear,
      movedRequests: verified.eligibleRequestIds.length,
      archiveSpreadsheetId: verified.archiveSpreadsheetId,
      archiveSpreadsheetUrl: verified.archiveSpreadsheetUrl,
      message: verified.eligibleRequestIds.length + ' permohonan selesai dipindahkan ke arsip tahun ' + targetYear + '.'
    };
  }));
}

function buildYearlyArchivePlan_(year) {
  const targetYear = resolveArchiveYear_(year);
  const archiveSpreadsheet = openArchiveSpreadsheet_(targetYear);
  const source = collectArchiveSourceData_(targetYear);
  const archiveMasterIds = readArchiveRequestIds_(archiveSpreadsheet);
  const eligibleRequestIds = source.requests.map(function(item) { return item.id; });
  const missingRequestIds = eligibleRequestIds.filter(function(id) {
    return archiveMasterIds.indexOf(id) === -1;
  });
  const idsToWriteMap = toLookup_(missingRequestIds);

  return {
    year: targetYear,
    archiveSpreadsheetId: archiveSpreadsheet.getId(),
    archiveSpreadsheetUrl: archiveSpreadsheet.getUrl(),
    eligibleRequests: source.requests.map(function(item) {
      return {
        id: item.id,
        activityName: item.activityName,
        year: item.archiveYear,
        status: item.status
      };
    }),
    eligibleRequestIds: eligibleRequestIds,
    alreadyArchivedRequestIds: eligibleRequestIds.filter(function(id) {
      return archiveMasterIds.indexOf(id) !== -1;
    }),
    missingRequestIds: missingRequestIds,
    masterRows: source.masterRows.filter(function(row) { return idsToWriteMap[text_(row[0])]; }),
    documentRows: source.documentRows.filter(function(row) { return idsToWriteMap[text_(row[1])]; }),
    scheduleRows: source.scheduleRows.filter(function(row) { return idsToWriteMap[text_(row[1])]; }),
    employeeRows: source.employeeRows.filter(function(row) { return idsToWriteMap[text_(row[0])]; }),
    travelRows: source.travelRows.filter(function(row) { return idsToWriteMap[text_(row[0])]; }),
    fileRows: source.fileRows.filter(function(row) { return idsToWriteMap[text_(row[0])]; }),
    counts: {
      requests: eligibleRequestIds.length,
      requestsToWrite: missingRequestIds.length,
      documents: source.documentRows.length,
      schedules: source.scheduleRows.length,
      employees: source.employeeRows.length,
      travel: source.travelRows.length,
      files: source.fileRows.length
    }
  };
}

function collectArchiveSourceData_(targetYear) {
  const masterRows = readDataRows_(getSheet_('MASTER'), MASTER_HEADERS.length);
  const eligibleMasterRows = [];
  const eligibleRequests = [];
  const requestIds = {};

  masterRows.forEach(function(row) {
    const dto = masterRowToDto_(row);
    if (!dto.id || dto.status !== 'ARCHIVED') return;
    const archiveYear = inferArchiveYearFromRequest_(dto);
    if (archiveYear !== targetYear) return;
    eligibleMasterRows.push(row);
    eligibleRequests.push({
      id: dto.id,
      activityName: dto.activityName,
      archiveYear: archiveYear,
      status: dto.status
    });
    requestIds[dto.id] = true;
  });

  return {
    requests: eligibleRequests,
    masterRows: eligibleMasterRows,
    documentRows: readDataRows_(getSheet_('DOCUMENTS'), DOCUMENT_HEADERS.length).filter(function(row) {
      return requestIds[text_(row[1])];
    }),
    scheduleRows: readDataRows_(getSheet_('SCHEDULES', false), SCHEDULE_HEADERS.length).filter(function(row) {
      return requestIds[text_(row[1])];
    }),
    employeeRows: readDataRows_(getSheet_('EMPLOYEES'), EMPLOYEE_HEADERS.length).filter(function(row) {
      return requestIds[text_(row[0])];
    }),
    travelRows: readDataRows_(getSheet_('TRAVEL'), TRAVEL_HEADERS.length).filter(function(row) {
      return requestIds[text_(row[0])];
    }),
    fileRows: readDataRows_(getSheet_('FILES'), GENERATED_FILE_HEADERS.length).filter(function(row) {
      return requestIds[text_(row[0])];
    })
  };
}

function resolveArchiveYear_(year) {
  const numeric = Number(year || new Date().getFullYear() - 1);
  if (!numeric || !isFinite(numeric) || numeric < 2000 || numeric > 2999) {
    throw new Error('Tahun arsip tidak valid.');
  }
  return numeric;
}

function inferArchiveYearFromRequest_(request) {
  const candidates = [request.letterDate, request.createdAt, request.updatedAt, request.endDate, request.startDate];
  for (let i = 0; i < candidates.length; i++) {
    const year = extractYearFromValue_(candidates[i]);
    if (year) return year;
  }
  return 0;
}

function extractYearFromValue_(value) {
  if (!value) return 0;
  if (isDate_(value)) return value.getFullYear();
  const raw = text_(value);
  const match = raw.match(/^(\d{4})[-/]/) || raw.match(/(\d{4})$/);
  return match ? Number(match[1]) : 0;
}

function openArchiveSpreadsheet_(year) {
  const name = 'Arsip DPSP ' + year;
  const mainSpreadsheet = getSpreadsheet_();
  const parent = firstParentFolderOfFile_(mainSpreadsheet.getId());
  const files = DriveApp.getFilesByName(name);
  while (files.hasNext()) {
    const file = files.next();
    const spreadsheet = SpreadsheetApp.openById(file.getId());
    if (!parent || fileIsInsideFolder_(file, parent.getId())) {
      ensureArchiveStructure_(spreadsheet);
      return spreadsheet;
    }
  }

  const spreadsheet = SpreadsheetApp.create(name);
  const file = DriveApp.getFileById(spreadsheet.getId());
  if (parent) {
    parent.addFile(file);
    try {
      DriveApp.getRootFolder().removeFile(file);
    } catch (error) {}
  }
  ensureArchiveStructure_(spreadsheet);
  return spreadsheet;
}

function ensureArchiveStructure_(spreadsheet) {
  ensureArchiveSheet_(spreadsheet, 'MASTER', MASTER_HEADERS);
  ensureArchiveSheet_(spreadsheet, 'DOCUMENTS', DOCUMENT_HEADERS);
  ensureArchiveSheet_(spreadsheet, 'SCHEDULES', SCHEDULE_HEADERS);
  ensureArchiveSheet_(spreadsheet, 'EMPLOYEES', EMPLOYEE_HEADERS);
  ensureArchiveSheet_(spreadsheet, 'TRAVEL', TRAVEL_HEADERS);
  ensureArchiveSheet_(spreadsheet, 'FILES', GENERATED_FILE_HEADERS);
}

function ensureArchiveSheet_(spreadsheet, key, headers) {
  const descriptor = APP_CONFIG.SHEETS[key];
  let sheet = spreadsheet.getSheetByName(descriptor.name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(descriptor.name);
  }
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function readArchiveRequestIds_(spreadsheet) {
  const sheet = ensureArchiveSheet_(spreadsheet, 'MASTER', MASTER_HEADERS);
  return readDataRows_(sheet, MASTER_HEADERS.length).map(function(row) {
    return text_(row[0]);
  }).filter(Boolean);
}

function appendArchiveRows_(spreadsheet, key, headers, rows, idColumnIndex) {
  if (!rows.length) return;
  const sheet = ensureArchiveSheet_(spreadsheet, key, headers);
  const existingIds = readExistingArchiveIds_(sheet, headers.length, idColumnIndex);
  const rowsToAppend = rows.filter(function(row) {
    return !existingIds[text_(row[idColumnIndex])];
  });
  if (!rowsToAppend.length) return;
  const startRow = Math.max(sheet.getLastRow(), 1) + 1;
  sheet.getRange(startRow, 1, rowsToAppend.length, headers.length).setValues(rowsToAppend);
}

function readExistingArchiveIds_(sheet, width, idColumnIndex) {
  return readDataRows_(sheet, width).reduce(function(result, row) {
    const id = text_(row[idColumnIndex]);
    if (id) result[id] = true;
    return result;
  }, {});
}

function removeArchivedRowsFromSource_(requestIds) {
  const lookup = toLookup_(requestIds);
  rewriteDataRows_(getSheet_('MASTER'), readDataRows_(getSheet_('MASTER'), MASTER_HEADERS.length).filter(function(row) {
    return !lookup[text_(row[0])];
  }), MASTER_HEADERS.length);
  rewriteDataRows_(getSheet_('DOCUMENTS'), readDataRows_(getSheet_('DOCUMENTS'), DOCUMENT_HEADERS.length).filter(function(row) {
    return !lookup[text_(row[1])];
  }), DOCUMENT_HEADERS.length);
  const scheduleSheet = getSheet_('SCHEDULES', false);
  if (scheduleSheet) {
    rewriteDataRows_(scheduleSheet, readDataRows_(scheduleSheet, SCHEDULE_HEADERS.length).filter(function(row) {
      return !lookup[text_(row[1])];
    }), SCHEDULE_HEADERS.length);
  }
  rewriteDataRows_(getSheet_('EMPLOYEES'), readDataRows_(getSheet_('EMPLOYEES'), EMPLOYEE_HEADERS.length).filter(function(row) {
    return !lookup[text_(row[0])];
  }), EMPLOYEE_HEADERS.length);
  rewriteDataRows_(getSheet_('TRAVEL'), readDataRows_(getSheet_('TRAVEL'), TRAVEL_HEADERS.length).filter(function(row) {
    return !lookup[text_(row[0])];
  }), TRAVEL_HEADERS.length);
  rewriteDataRows_(getSheet_('FILES'), readDataRows_(getSheet_('FILES'), GENERATED_FILE_HEADERS.length).filter(function(row) {
    return !lookup[text_(row[0])];
  }), GENERATED_FILE_HEADERS.length);
}

function toLookup_(values) {
  return (values || []).reduce(function(result, value) {
    if (value) result[value] = true;
    return result;
  }, {});
}

function firstParentFolderOfFile_(fileId) {
  try {
    const parents = DriveApp.getFileById(fileId).getParents();
    return parents.hasNext() ? parents.next() : null;
  } catch (error) {
    return null;
  }
}

function fileIsInsideFolder_(file, folderId) {
  const parents = file.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === folderId) return true;
  }
  return false;
}