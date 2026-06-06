function previewLegacyMigration() {
  const user = assertAuthorized_();
  assertCanWrite_(user);
  const sheet = getSheet_('MASTER');
  const rows = readDataRows_(sheet, MASTER_HEADERS.length);
  const groups = {};
  const missingIds = [];
  const dateWarnings = [];
  const unsupportedDocuments = [];

  rows.forEach(function(row, index) {
    const id = text_(row[0]);
    if (!id) {
      missingIds.push(index + 2);
      return;
    }
    if (!groups[id]) groups[id] = [];
    groups[id].push(row);
    if (row[17] && !parseLegacyDateRange_(row[17])) {
      dateWarnings.push({ row: index + 2, value: String(row[17]) });
    }
    try {
      resolveTemplateKey_({
        activityType: text_(row[1]),
        type: text_(row[2]),
        speakerSubtype: text_(row[3]),
        speakerStatus: text_(row[4])
      });
    } catch (error) {
      unsupportedDocuments.push({ row: index + 2, type: text_(row[2]) });
    }
  });

  const duplicateIds = Object.keys(groups).filter(function(id) {
    return groups[id].length > 1;
  }).map(function(id) {
    return { id: id, rows: groups[id].length };
  });
  return serializeValue_({
    user: user.email,
    sourceRows: rows.length,
    uniqueRequests: Object.keys(groups).length,
    duplicateIds: duplicateIds,
    missingIdRows: missingIds,
    dateWarnings: dateWarnings,
    unsupportedDocuments: unsupportedDocuments,
    ready: missingIds.length === 0 && unsupportedDocuments.length === 0
  });
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
    const sourceRows = readDataRows_(masterSheet, MASTER_HEADERS.length);
    if (!sourceRows.length) return { ok: true, migrated: 0, message: 'Tidak ada data lama.' };

    const preview = previewLegacyMigration();
    if (!preview.ready) {
      throw new Error(
        'Migrasi belum aman. Perbaiki ID kosong atau jenis dokumen tidak didukung melalui previewLegacyMigration().'
      );
    }

    const backupName = 'Backup Master ' +
      Utilities.formatDate(new Date(), APP_CONFIG.TIME_ZONE, 'yyyyMMdd-HHmmss');
    const backup = masterSheet.copyTo(ss).setName(backupName);
    backup.hideSheet();

    const groups = {};
    sourceRows.forEach(function(row) {
      const id = text_(row[0]);
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

      setMaster_(row, 'Jenis Surat', uniqueTextList_(documentTypes).join('\n'));
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
      backupSheet: backupName
    });
    return {
      ok: true,
      sourceRows: sourceRows.length,
      migrated: masterOutput.length,
      documents: documentOutput.length,
      backupSheet: backupName
    };
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
