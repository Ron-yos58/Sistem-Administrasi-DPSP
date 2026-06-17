function previewLegacyMigration() {
  const user = assertAuthorized_();
  assertAdmin_(user);
  const sheet = getSheet_('MASTER');
  const rows = readLegacyMasterRows_(sheet, LEGACY_MASTER_HEADERS.length);
  const groups = {};
  const missingIds = [];
  const generatedIdRows = [];
  const dateWarnings = [];
  const unsupportedDocuments = [];
  const schemaErrors = [];

  try {
    validateLegacyAutocratColumns_(sheet);
  } catch (error) {
    schemaErrors.push(text_(error && error.message ? error.message : error));
  }

  rows.forEach(function(row, index) {
    const idInfo = resolveLegacyRequestId_(row, index);
    const id = idInfo.id;
    if (!id) {
      missingIds.push(index + 2);
      return;
    }
    if (idInfo.generated) {
      generatedIdRows.push({ row: index + 2, id: id });
    }
    if (!groups[id]) groups[id] = [];
    groups[id].push(row);
    if (row[17] && !parseLegacyDateRange_(row[17])) {
      dateWarnings.push({ row: index + 2, value: String(row[17]) });
    }
    const descriptor = normalizeDocumentDescriptor_({
      activityType: row[1],
      type: row[2],
      speakerSubtype: row[3],
      speakerStatus: row[4]
    });
    if (descriptor.activityType === 'Penugasan Narasumber' && descriptor.type === 'Surat Tugas' && !descriptor.speakerSubtype) {
      const workshopDocId = text_(row[42]); // Merged Doc ID - Penugasan Narasumber (col 43, index 42)
      const promotionDocId = text_(row[46]); // Merged Doc ID - Penugasan Narasumber (Promosi) (col 47, index 46)
      if (promotionDocId && !workshopDocId) {
        descriptor.speakerSubtype = 'Promosi';
      } else {
        descriptor.speakerSubtype = 'Workshop';
      }
    }
    try {
      resolveTemplateKey_(descriptor);
    } catch (error) {
      unsupportedDocuments.push({
        row: index + 2,
        descriptor: descriptor,
        message: text_(error && error.message ? error.message : error)
      });
    }
  });

  const duplicateIds = Object.keys(groups).filter(function(id) {
    return groups[id].length > 1;
  }).map(function(id) {
    return { id: id, rows: groups[id].length };
  });
  const result = serializeValue_({
    user: user.email,
    sourceSheet: sheet.getName(),
    sourceRows: rows.length,
    uniqueRequests: Object.keys(groups).length,
    duplicateIds: duplicateIds,
    generatedIdRows: generatedIdRows,
    missingIdRows: missingIds,
    dateWarnings: dateWarnings,
    unsupportedDocuments: unsupportedDocuments,
    schemaErrors: schemaErrors,
    autocratColumnsToRemove: schemaErrors.length ? 0 : AUTOCRAT_HEADERS.length,
    ready: missingIds.length === 0 && unsupportedDocuments.length === 0 && schemaErrors.length === 0
  });

  Logger.log('previewLegacyMigration summary: %s', JSON.stringify({
    ready: result.ready,
    sourceRows: result.sourceRows,
    uniqueRequests: result.uniqueRequests,
    duplicateIdCount: result.duplicateIds.length,
    generatedIdCount: result.generatedIdRows.length,
    missingIdCount: result.missingIdRows.length,
    dateWarningCount: result.dateWarnings.length,
    unsupportedCount: result.unsupportedDocuments.length,
    schemaErrorCount: result.schemaErrors.length,
    autocratColumnsToRemove: result.autocratColumnsToRemove
  }));

  if (!result.ready) {
    Logger.log('previewLegacyMigration blockers: %s', JSON.stringify({
      generatedIdRows: result.generatedIdRows.slice(0, 20),
      missingIdRows: result.missingIdRows,
      unsupportedDocuments: result.unsupportedDocuments.slice(0, 20),
      schemaErrors: result.schemaErrors
    }));
  }

  return result;
}

function migrateLegacyData(confirmText) {
  const user = assertAuthorized_();
  assertAdmin_(user);
  if (text_(confirmText) !== 'MIGRATE') {
    throw new Error('Konfirmasi migrasi harus tepat: MIGRATE');
  }

  return withScriptLock_(function() {
    const documentSheet = getSheet_('DOCUMENTS');
    const scheduleSheet = getSheet_('SCHEDULES');
    if (documentSheet.getLastRow() > 1) {
      throw new Error('Sheet Dokumen Permohonan sudah berisi data. Migrasi dibatalkan agar tidak overwrite.');
    }

    const ss = getSpreadsheet_();
    const masterSheet = getSheet_('MASTER');
    const sourceRows = readLegacyMasterRows_(masterSheet, LEGACY_MASTER_HEADERS.length);
    if (!sourceRows.length) return { ok: true, migrated: 0, message: 'Tidak ada data lama.' };

    const preview = previewLegacyMigration();
    if (!preview.ready) {
      throw new Error(
        'Migrasi belum aman. missingIdRows=' + preview.missingIdRows.length +
        ', unsupportedDocuments=' + preview.unsupportedDocuments.length +
        ', schemaErrors=' + preview.schemaErrors.length +
        '. Jalankan previewLegacyMigration lalu cek Execution log (summary dan blockers).'
      );
    }

    const backupStamp = Utilities.formatDate(
      new Date(),
      APP_CONFIG.TIME_ZONE,
      'yyyyMMdd-HHmmss'
    );
    const backupSheets = [
      createMigrationBackup_(ss, masterSheet, 'Master', backupStamp),
      createMigrationBackup_(ss, getSheet_('EMPLOYEES'), 'Pegawai', backupStamp),
      createMigrationBackup_(ss, getSheet_('TRAVEL'), 'Perjadin', backupStamp),
      createMigrationBackup_(ss, documentSheet, 'Dokumen', backupStamp),
      createMigrationBackup_(ss, scheduleSheet, 'Jadwal', backupStamp)
    ];

    const groups = {};
    const legacyRequestIdMap = {};
    sourceRows.forEach(function(row, index) {
      const explicitId = text_(row[0]);
      const id = resolveLegacyRequestId_(row, index).id;
      if (!groups[id]) groups[id] = [];
      groups[id].push(row);
      if (explicitId && explicitId !== id) legacyRequestIdMap[explicitId] = id;
    });

    const employeeRows = readDataRows_(getSheet_('EMPLOYEES'), EMPLOYEE_HEADERS.length);
    const employeesByRequest = {};
    employeeRows.forEach(function(row) {
      const originalId = text_(row[0]);
      const participantKey = text_(row[8]);
      const participantRequestId = participantKey.indexOf('|') !== -1 ? participantKey.split('|')[0] : '';
      const mappedId = legacyRequestIdMap[originalId] || legacyRequestIdMap[participantRequestId] || originalId;
      const id = text_(mappedId);
      if (id && id !== originalId) row[0] = id;
      if (!employeesByRequest[id]) employeesByRequest[id] = [];
      if (!row[8]) row[8] = inferParticipantKey_(id, row[2], row[1]);
      else if (participantRequestId && participantRequestId !== id) row[8] = normalizeParticipantKeyForRequest_(id, row[8]);
      employeesByRequest[id].push({
        name: row[1],
        identifier: row[2],
        role: row[3],
        unit: row[4],
        email: row[5],
        rank: row[6],
        category: row[7],
        participantKey: row[8]
      });
    });
    rewriteDataRows_(getSheet_('EMPLOYEES'), employeeRows, EMPLOYEE_HEADERS.length);

    const masterOutput = [];
    const documentOutput = [];
    const scheduleOutput = [];
    const now = new Date();

    Object.keys(groups).forEach(function(id) {
      const sourceGroup = groups[id];
      const row = compactLegacyMasterRow_(sourceGroup[0]);
      const documentTypes = [];
      const documentNumbers = [];

      sourceGroup.forEach(function(source) {
        const descriptor = normalizeDocumentDescriptor_({
          activityType: source[1],
          type: source[2],
          speakerSubtype: source[3],
          speakerStatus: source[4]
        });
        if (descriptor.activityType === 'Penugasan Narasumber' && descriptor.type === 'Surat Tugas' && !descriptor.speakerSubtype) {
          const workshopDocId = text_(source[42]); // Merged Doc ID - Penugasan Narasumber (col 43, index 42)
          const promotionDocId = text_(source[46]); // Merged Doc ID - Penugasan Narasumber (Promosi) (col 47, index 46)
          if (promotionDocId && !workshopDocId) {
            descriptor.speakerSubtype = 'Promosi';
          } else {
            descriptor.speakerSubtype = 'Workshop';
          }
        }
        const templateKey = resolveTemplateKey_(descriptor);
        const columns = APP_CONFIG.LEGACY_DOCUMENT_COLUMNS[templateKey];
        const docId = 'DOC-' + Utilities.getUuid().slice(0, 12).toUpperCase();
        const generatedDocId = columns ? text_(source[columns[0] - 1]) : '';
        const generatedDocUrl = columns ? text_(source[columns[1] - 1]) : '';
        const generatedStatus = columns ? text_(source[columns[3] - 1]) : '';

        documentTypes.push(descriptor.type);
        if (source[6]) documentNumbers.push(descriptor.type + ': ' + source[6]);
        documentOutput.push([
          docId,
          id,
          descriptor.type,
          descriptor.speakerSubtype,
          descriptor.speakerStatus,
          source[6],
          templateKey,
          generatedDocId ? 'GENERATED' : (generatedStatus || 'PENDING'),
          generatedDocId,
          generatedDocUrl,
          '',
          '',
          '',
          legacyMasterValue_(source, 'Email Status') ? 'DRAFTED' : '',
          1,
          now,
          now,
          '',
          '',
          ''
        ]);

      });

      const range = parseLegacyDateRange_(sourceGroup[0][17]);
      if (range) {
        const timeRange = parseTimeRange_(sourceGroup[0][18]);
        const list = range.list || [{ start: range.start, end: range.end }];
        list.forEach(function(item, itemIndex) {
          scheduleOutput.push([
            'SCH-' + Utilities.getUuid().slice(0, 12).toUpperCase(),
            id,
            parseIsoDate_(item.start),
            parseIsoDate_(item.end || item.start),
            timeRange.startTime,
            timeRange.endTime,
            text_(sourceGroup[0][19]),
            itemIndex + 1
          ]);
        });
      }
      const employees = employeesByRequest[id] || [];
      const requestForRouting = {
        documents: documentOutput.filter(function(docRow) { return docRow[1] === id; })
          .map(function(docRow) { return { type: docRow[2] }; }),
        partnerEmail: masterValue_(row, 'Email Mitra'),
        partnerName: masterValue_(row, 'Nama Mitra'),
        faculties: text_(masterValue_(row, 'Fakultas Asal Narasumber')).split(/\s*,\s*/).filter(Boolean),
        speakerStatus: masterValue_(row, 'Status Narasumber')
      };

      // Keep Master value validation-safe (single enum); full document list is stored in DOCUMENTS sheet.
      const inferredSubtype = uniqueTextList_(
        documentOutput.filter(function(docRow) { return docRow[1] === id; }).map(function(docRow) { return docRow[3]; })
      )[0] || '';
      const normFirst = normalizeDocumentDescriptor_({
        activityType: sourceGroup[0][1],
        type: sourceGroup[0][2],
        speakerSubtype: sourceGroup[0][3],
        speakerStatus: sourceGroup[0][4]
      });
      setMaster_(row, 'Tipe Kegiatan', normFirst.activityType);
      setMaster_(row, 'Status Narasumber', normFirst.speakerStatus);
      setMaster_(row, 'Sub-Tipe Kegiatan', inferredSubtype);
      setMaster_(row, 'Jenis Surat', uniqueTextList_(documentTypes)[0] || '');
      setMaster_(row, 'ID Permohonan', id);
      setMaster_(row, 'Nomor Surat', documentNumbers.join('\n'));
      setMaster_(row, 'Status Permohonan', 'DRAFT');
      setMaster_(row, 'Dibuat Oleh', user.email);
      setMaster_(row, 'Dibuat Pada', now);
      setMaster_(row, 'Diubah Oleh', user.email);
      setMaster_(row, 'Diubah Pada', now);
      setMaster_(row, 'Client Token', Utilities.getUuid());
      setMaster_(row, 'Tanggal Mulai ISO', range ? range.start : '');
      setMaster_(row, 'Tanggal Selesai ISO', range ? range.end : '');
      setMaster_(row, 'Revision', 1);
      applyEmployeeAndRoutingColumns_(row, employees, requestForRouting);
      masterOutput.push(row);
    });

    const removedAutocratColumns = removeLegacyAutocratColumns_(masterSheet);
    rewriteDataRows_(masterSheet, masterOutput, MASTER_HEADERS.length);
    masterSheet.getRange(1, 1, 1, MASTER_HEADERS.length).setValues([MASTER_HEADERS]);
    rewriteDataRows_(documentSheet, documentOutput, DOCUMENT_HEADERS.length);
    rewriteDataRows_(scheduleSheet, scheduleOutput, SCHEDULE_HEADERS.length);
    syncTravelDataInternal_('');
    clearAppCache_();
    logAudit_('MIGRATE_LEGACY_DATA', '', true, {
      sourceRows: sourceRows.length,
      requests: masterOutput.length,
      documents: documentOutput.length,
      schedules: scheduleOutput.length,
      backupSheets: backupSheets,
      removedAutocratColumns: removedAutocratColumns
    });
    return {
      ok: true,
      sourceRows: sourceRows.length,
      migrated: masterOutput.length,
      documents: documentOutput.length,
      schedules: scheduleOutput.length,
      backupSheets: backupSheets,
      removedAutocratColumns: removedAutocratColumns
    };
  });
}

function resolveLegacyRequestId_(row, index) {
  const explicitId = text_((row || [])[0]);
  if (explicitId) return { id: explicitId, generated: false };

  const fields = [
    row[1], row[3], row[4], row[5], row[7], row[8], row[9], row[10], row[11], row[12], row[14], row[17], row[19]
  ].map(function(value) {
    return text_(value).toLowerCase().replace(/\s+/g, ' ').trim();
  });
  const fingerprint = fields.join('|');
  if (!fingerprint) return { id: '', generated: false };

  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, fingerprint);
  const suffix = digest.slice(0, 6).map(function(byte) {
    const normalized = byte < 0 ? byte + 256 : byte;
    return normalized.toString(16).padStart(2, '0');
  }).join('').toUpperCase();

  return { id: 'LEGACY-' + suffix, generated: true };
}

function migrateLegacyDataConfirmed() {
  return migrateLegacyData('MIGRATE');
}

function cleanupMigratedAutocratColumns() {
  const user = assertAuthorized_();
  assertAdmin_(user);

  return withScriptLock_(function() {
    const ss = getSpreadsheet_();
    const masterSheet = getSheet_('MASTER');
    const schema = inspectMasterSchema_(masterSheet);
    if (schema === 'ACTIVE') {
      return {
        ok: true,
        removedAutocratColumns: 0,
        message: 'Kolom Autocrat sudah tidak ada pada Master Permohonan.'
      };
    }
    if (schema !== 'LEGACY') {
      throw new Error(
        'Struktur Master Permohonan tidak dikenali. Cleanup dibatalkan tanpa menghapus kolom.'
      );
    }

    const stamp = Utilities.formatDate(new Date(), APP_CONFIG.TIME_ZONE, 'yyyyMMdd-HHmmss');
    const backupSheet = createMigrationBackup_(ss, masterSheet, 'Master Pre-Cleanup', stamp);
    const removedAutocratColumns = removeLegacyAutocratColumns_(masterSheet);
    masterSheet.getRange(1, 1, 1, MASTER_HEADERS.length).setValues([MASTER_HEADERS]);
    clearAppCache_();

    logAudit_('CLEANUP_AUTOCRAT_COLUMNS', '', true, {
      backupSheet: backupSheet,
      removedAutocratColumns: removedAutocratColumns,
      user: user.email
    });
    return {
      ok: true,
      backupSheet: backupSheet,
      removedAutocratColumns: removedAutocratColumns,
      message: 'Kolom Autocrat AM-BV berhasil dihapus dari Master Permohonan.'
    };
  });
}

function repairMigratedMasterIds() {
  const user = assertAuthorized_();
  assertAdmin_(user);

  return withScriptLock_(function() {
    const masterSheet = getSheet_('MASTER');
    const rows = readLegacyMasterRows_(masterSheet, MASTER_HEADERS.length);
    if (!rows.length) return { ok: true, updated: 0, message: 'Master kosong.' };

    let updated = 0;
    const usedIds = {};

    rows.forEach(function(row, index) {
      const existingId = text_(row[0]);
      if (existingId) {
        usedIds[existingId] = true;
        return;
      }

      const idInfo = resolveLegacyRequestId_(row, index);
      let id = idInfo.id;
      if (!id) return;

      // Prevent accidental collision when multiple blank-ID rows resolve to same fingerprint.
      if (usedIds[id]) id = id + '-' + String(index + 2);
      row[0] = id;
      usedIds[id] = true;
      updated += 1;
    });

    if (!updated) return { ok: true, updated: 0, message: 'Tidak ada ID kosong pada Master.' };

    rewriteDataRows_(masterSheet, rows, MASTER_HEADERS.length);
    clearAppCache_();
    logAudit_('REPAIR_MIGRATED_MASTER_IDS', '', true, {
      updated: updated,
      totalRows: rows.length
    });

    return {
      ok: true,
      updated: updated,
      totalRows: rows.length,
      message: 'ID Master berhasil diperbaiki.'
    };
  });
}

function repairMigratedEmployeeIds() {
  const user = assertAuthorized_();
  assertAdmin_(user);

  return withScriptLock_(function() {
    const employeeSheet = getSheet_('EMPLOYEES');
    const rows = readDataRows_(employeeSheet, EMPLOYEE_HEADERS.length);
    if (!rows.length) return { ok: true, updated: 0, message: 'Data Pegawai kosong.' };

    const mapping = buildLegacyEmployeeRequestIdMap_();
    const validRequestIds = {};
    readDataRows_(getSheet_('MASTER'), MASTER_HEADERS.length).forEach(function(row) {
      const id = text_(row[0]);
      if (id) validRequestIds[id] = true;
    });

    let updated = 0;
    rows.forEach(function(row) {
      const currentId = text_(row[0]);
      const participantKey = text_(row[8]);
      const participantRequestId = participantKey.indexOf('|') !== -1 ? participantKey.split('|')[0] : '';
      const repairedId = validRequestIds[currentId]
        ? currentId
        : (mapping[currentId] || mapping[participantRequestId] || '');

      if (!repairedId || repairedId === currentId) {
        if (participantRequestId && validRequestIds[currentId] && participantRequestId !== currentId) {
          row[8] = normalizeParticipantKeyForRequest_(currentId, participantKey);
          updated += 1;
        }
        return;
      }

      row[0] = repairedId;
      row[8] = participantKey ? normalizeParticipantKeyForRequest_(repairedId, participantKey) : inferParticipantKey_(repairedId, row[2], row[1]);
      updated += 1;
    });

    if (!updated) {
      return {
        ok: true,
        updated: 0,
        totalRows: rows.length,
        message: 'Tidak ada ID pegawai yang perlu diperbaiki.'
      };
    }

    rewriteDataRows_(employeeSheet, rows, EMPLOYEE_HEADERS.length);
    clearAppCache_();
    logAudit_('REPAIR_MIGRATED_EMPLOYEE_IDS', '', true, {
      updated: updated,
      totalRows: rows.length
    });

    return {
      ok: true,
      updated: updated,
      totalRows: rows.length,
      message: 'ID Data Pegawai berhasil diperbaiki.'
    };
  });
}

function buildLegacyEmployeeRequestIdMap_() {
  const backupSheet = findLatestMigrationBackupSheet_('Backup Master ');
  if (!backupSheet) return {};

  const rows = readLegacyMasterRows_(backupSheet, LEGACY_MASTER_HEADERS.length);
  return rows.reduce(function(map, row, index) {
    const oldId = text_(row[0]);
    const resolvedId = resolveLegacyRequestId_(row, index).id;
    if (oldId && resolvedId && oldId !== resolvedId) map[oldId] = resolvedId;
    return map;
  }, {});
}

function findLatestMigrationBackupSheet_(prefix) {
  const sheets = getSpreadsheet_().getSheets().filter(function(sheet) {
    return sheet.getName().indexOf(prefix) === 0;
  }).sort(function(a, b) {
    return b.getName().localeCompare(a.getName());
  });
  return sheets.length ? sheets[0] : null;
}

function compactLegacyMasterRow_(legacyRow) {
  return MASTER_HEADERS.map(function(header) {
    const legacyIndex = LEGACY_MASTER_HEADERS.indexOf(header);
    return legacyIndex === -1 ? '' : legacyRow[legacyIndex];
  });
}

function legacyMasterValue_(row, header) {
  const index = LEGACY_MASTER_HEADERS.indexOf(header);
  return index === -1 ? '' : row[index];
}

function removeLegacyAutocratColumns_(sheet) {
  const startColumn = 39; // AM
  const columnCount = AUTOCRAT_HEADERS.length; // AM:BV
  validateLegacyAutocratColumns_(sheet);
  sheet.deleteColumns(startColumn, columnCount);
  return columnCount;
}

function validateLegacyAutocratColumns_(sheet) {
  const startColumn = 39; // AM
  const columnCount = AUTOCRAT_HEADERS.length; // AM:BV
  if (sheet.getMaxColumns() < startColumn + columnCount - 1) {
    throw new Error('Kolom Autocrat AM-BV tidak lengkap. Migrasi dihentikan tanpa menghapus kolom.');
  }

  const actualHeaders = sheet.getRange(1, startColumn, 1, columnCount).getDisplayValues()[0];
  const mismatch = actualHeaders.findIndex(function(header, index) {
    return text_(header) !== AUTOCRAT_HEADERS[index];
  });
  if (mismatch !== -1) {
    throw new Error(
      'Header Autocrat tidak cocok di kolom ' + (startColumn + mismatch) +
      ': expected="' + AUTOCRAT_HEADERS[mismatch] +
      '", actual="' + text_(actualHeaders[mismatch]) + '". Kolom tidak dihapus.'
    );
  }
  return true;
}

function inspectMasterSchema_(sheet) {
  if (!sheet) return 'UNKNOWN';
  const maxColumns = sheet.getMaxColumns();
  if (maxColumns >= LEGACY_MASTER_HEADERS.length) {
    const legacyHeaders = sheet.getRange(1, 1, 1, LEGACY_MASTER_HEADERS.length).getDisplayValues()[0];
    if (headersMatch_(legacyHeaders, LEGACY_MASTER_HEADERS)) return 'LEGACY';
  }
  if (maxColumns >= MASTER_HEADERS.length) {
    const activeHeaders = sheet.getRange(1, 1, 1, MASTER_HEADERS.length).getDisplayValues()[0];
    if (headersMatch_(activeHeaders, MASTER_HEADERS)) return 'ACTIVE';
  }
  return 'UNKNOWN';
}

function headersMatch_(actual, expected) {
  return expected.every(function(header, index) {
    return text_(actual[index]) === header;
  });
}

function createMigrationBackup_(ss, sourceSheet, label, stamp) {
  const name = ('Backup ' + label + ' ' + stamp).slice(0, 99);
  const backup = sourceSheet.copyTo(ss).setName(name);
  backup.hideSheet();
  return name;
}

function readLegacyMasterRows_(sheet, width) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, width).getValues()
    .filter(function(row) {
      return row.some(function(value) { return text_(value) !== ''; });
    });
}

function parseLegacyDateRange_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) {
    const iso = formatDate_(value, 'yyyy-MM-dd');
    return { start: iso, end: iso };
  }

  const input = text_(value).replace(/\s+/g, ' ');
  const months = {
    januari: 1, februari: 2, maret: 3, april: 4, mei: 5, juni: 6,
    juli: 7, agustus: 8, september: 9, oktober: 10, november: 11, desember: 12
  };
  function iso(day, monthName, year) {
    const month = months[String(monthName).toLowerCase()];
    if (!month) return '';
    const candidate = year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
    try {
      parseIsoDate_(candidate);
      return candidate;
    } catch (error) {
      return '';
    }
  }

  let match = input.match(/^(\d{1,2})\s+(?:dan|&)\s+(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/i);
  if (match) {
    const start = iso(Number(match[1]), match[3], Number(match[4]));
    const end = iso(Number(match[2]), match[3], Number(match[4]));
    if (start && end) {
      return {
        start: start,
        end: end,
        list: [
          { start: start, end: start },
          { start: end, end: end }
        ]
      };
    }
  }

  match = input.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})\s*(?:-|s\/d|s\.d\.?|sd|sampai)\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/i);
  if (match) {
    const start = iso(Number(match[1]), match[2], Number(match[3]));
    const end = iso(Number(match[4]), match[5], Number(match[6]));
    return start && end ? { start: start, end: end } : null;
  }
  match = input.match(/^(\d{1,2})\s+([A-Za-z]+)\s*(?:-|s\/d|s\.d\.?|sd|sampai)\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/i);
  if (match) {
    const start = iso(Number(match[1]), match[2], Number(match[5]));
    const end = iso(Number(match[3]), match[4], Number(match[5]));
    return start && end ? { start: start, end: end } : null;
  }
  match = input.match(/^(\d{1,2})\s*(?:-|s\/d|s\.d\.?|sd|sampai)\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/i);
  if (match) {
    const start = iso(Number(match[1]), match[3], Number(match[4]));
    const end = iso(Number(match[2]), match[3], Number(match[4]));
    return start && end ? { start: start, end: end } : null;
  }
  match = input.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/i);
  if (match) {
    const date = iso(Number(match[1]), match[2], Number(match[3]));
    return date ? { start: date, end: date } : null;
  }
  match = input.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) return { start: input, end: input };
  return null;
}
