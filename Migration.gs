function previewLegacyMigration() {
  const user = assertAuthorized_();
  assertCanWrite_(user);
  const sheet = getSheet_('MASTER');
  const rows = readLegacyMasterRows_(sheet, MASTER_HEADERS.length);
  const groups = {};
  const missingIds = [];
  const generatedIdRows = [];
  const dateWarnings = [];
  const unsupportedDocuments = [];

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
    const descriptor = {
      activityType: text_(row[1]),
      type: text_(row[2]),
      speakerSubtype: text_(row[3]),
      speakerStatus: text_(row[4])
    };
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
    ready: missingIds.length === 0 && unsupportedDocuments.length === 0
  });

  Logger.log('previewLegacyMigration summary: %s', JSON.stringify({
    ready: result.ready,
    sourceRows: result.sourceRows,
    uniqueRequests: result.uniqueRequests,
    duplicateIdCount: result.duplicateIds.length,
    generatedIdCount: result.generatedIdRows.length,
    missingIdCount: result.missingIdRows.length,
    dateWarningCount: result.dateWarnings.length,
    unsupportedCount: result.unsupportedDocuments.length
  }));

  if (!result.ready) {
    Logger.log('previewLegacyMigration blockers: %s', JSON.stringify({
      generatedIdRows: result.generatedIdRows.slice(0, 20),
      missingIdRows: result.missingIdRows,
      unsupportedDocuments: result.unsupportedDocuments.slice(0, 20)
    }));
  }

  return result;
}

function migrateLegacyData(confirmText) {
  const user = assertAuthorized_();
  assertCanWrite_(user);
  if (text_(confirmText) !== 'MIGRATE') {
    throw new Error('Konfirmasi migrasi harus tepat: MIGRATE');
  }

  return withScriptLock_(function() {
    const documentSheet = getSheet_('DOCUMENTS');
    if (documentSheet.getLastRow() > 1) {
      throw new Error('Sheet Dokumen Permohonan sudah berisi data. Migrasi dibatalkan agar tidak overwrite.');
    }

    const ss = getSpreadsheet_();
    const masterSheet = getSheet_('MASTER');
    const sourceRows = readLegacyMasterRows_(masterSheet, MASTER_HEADERS.length);
    if (!sourceRows.length) return { ok: true, migrated: 0, message: 'Tidak ada data lama.' };

    const preview = previewLegacyMigration();
    if (!preview.ready) {
      throw new Error(
        'Migrasi belum aman. missingIdRows=' + preview.missingIdRows.length +
        ', unsupportedDocuments=' + preview.unsupportedDocuments.length +
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
      createMigrationBackup_(ss, documentSheet, 'Dokumen', backupStamp)
    ];

    const groups = {};
    sourceRows.forEach(function(row, index) {
      const id = resolveLegacyRequestId_(row, index).id;
      if (!groups[id]) groups[id] = [];
      groups[id].push(row);
    });

    const employeeRows = readDataRows_(getSheet_('EMPLOYEES'), EMPLOYEE_HEADERS.length);
    const employeesByRequest = {};
    employeeRows.forEach(function(row) {
      const id = text_(row[0]);
      if (!employeesByRequest[id]) employeesByRequest[id] = [];
      if (!row[8]) row[8] = inferParticipantKey_(id, row[2], row[1]);
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
    const now = new Date();

    Object.keys(groups).forEach(function(id) {
      const sourceGroup = groups[id];
      const row = sourceGroup[0].slice();
      const documentTypes = [];
      const documentNumbers = [];

      sourceGroup.forEach(function(source) {
        const descriptor = {
          activityType: text_(source[1]),
          type: text_(source[2]),
          speakerSubtype: text_(source[3]),
          speakerStatus: text_(source[4])
        };
        const templateKey = resolveTemplateKey_(descriptor);
        const columns = APP_CONFIG.DOCUMENT_COLUMNS[templateKey];
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
          source[74] ? 'DRAFTED' : '',
          1,
          now,
          now
        ]);

        if (columns) {
          row[columns[0] - 1] = source[columns[0] - 1];
          row[columns[1] - 1] = source[columns[1] - 1];
          row[columns[2] - 1] = source[columns[2] - 1];
          row[columns[3] - 1] = source[columns[3] - 1];
        }
      });

      const range = parseLegacyDateRange_(row[17]);
      const employees = employeesByRequest[id] || [];
      const requestForRouting = {
        documents: documentOutput.filter(function(docRow) { return docRow[1] === id; })
          .map(function(docRow) { return { type: docRow[2] }; }),
        partnerEmail: row[14],
        partnerName: row[12],
        faculties: text_(row[5]).split(/\s*,\s*/).filter(Boolean),
        speakerStatus: row[4]
      };

      // Keep Master value validation-safe (single enum); full document list is stored in DOCUMENTS sheet.
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

    rewriteDataRows_(masterSheet, masterOutput, MASTER_HEADERS.length);
    rewriteDataRows_(documentSheet, documentOutput, DOCUMENT_HEADERS.length);
    syncTravelDataInternal_('');
    clearAppCache_();
    logAudit_('MIGRATE_LEGACY_DATA', '', true, {
      sourceRows: sourceRows.length,
      requests: masterOutput.length,
      documents: documentOutput.length,
      backupSheets: backupSheets
    });
    return {
      ok: true,
      sourceRows: sourceRows.length,
      migrated: masterOutput.length,
      documents: documentOutput.length,
      backupSheets: backupSheets
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

function repairMigratedMasterIds() {
  const user = assertAuthorized_();
  assertCanWrite_(user);

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

  let match = input.match(/^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/i);
  if (match) {
    const date = iso(Number(match[1]), match[2], Number(match[3]));
    return date ? { start: date, end: date } : null;
  }
  match = input.match(/^(\d{1,2})\s*-\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/i);
  if (match) {
    const start = iso(Number(match[1]), match[3], Number(match[4]));
    const end = iso(Number(match[2]), match[3], Number(match[4]));
    return start && end ? { start: start, end: end } : null;
  }
  match = input.match(/^(\d{1,2})\s+([A-Za-z]+)\s*-\s*(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$/i);
  if (match) {
    const start = iso(Number(match[1]), match[2], Number(match[5]));
    const end = iso(Number(match[3]), match[4], Number(match[5]));
    return start && end ? { start: start, end: end } : null;
  }
  match = input.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) return { start: input, end: input };
  return null;
}
