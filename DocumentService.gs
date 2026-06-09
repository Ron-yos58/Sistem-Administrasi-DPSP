function generateDocument(documentId, force) {
  const user = assertAuthorized_();
  assertCanWrite_(user);
  const id = text_(documentId);

  return withScriptLock_(function() {
    try {
      const result = generateDocumentInternal_(id, Boolean(force), user);
      logAudit_('GENERATE_DOCUMENT', result.requestId, true, {
        documentId: id,
        reused: result.reused
      });
      return serializeValue_(result);
    } catch (error) {
      markDocumentError_(id, error.message);
      logAudit_('GENERATE_DOCUMENT', id, false, { error: error.message });
      throw error;
    }
  });
}

function generateAllDocuments(requestId, force) {
  const user = assertAuthorized_();
  assertCanWrite_(user);
  const id = text_(requestId);

  return withScriptLock_(function() {
    const documents = getDocumentsByRequest_(id);
    if (!documents.length) throw new Error('Tidak ada dokumen untuk permohonan ' + id);
    const results = [];
    const errors = [];

    documents.forEach(function(document) {
      try {
        results.push(generateDocumentInternal_(document.id, Boolean(force), user));
      } catch (error) {
        markDocumentError_(document.id, error.message);
        errors.push({ documentId: document.id, type: document.type, error: error.message });
      }
    });
    logAudit_('GENERATE_ALL_DOCUMENTS', id, errors.length === 0, {
      generated: results.length,
      errors: errors
    });
    return serializeValue_({ ok: errors.length === 0, results: results, errors: errors });
  });
}

function generateDocumentInternal_(documentId, force, user) {
  const documentSheet = getSheet_('DOCUMENTS');
  const rowNumber = findRowById_(documentSheet, documentId, 1);
  if (!rowNumber) throw new Error('Dokumen tidak ditemukan: ' + documentId);
  const row = documentSheet.getRange(rowNumber, 1, 1, DOCUMENT_HEADERS.length).getValues()[0];
  const document = documentRowToDto_(row);
  const detail = getRequestDetailInternal_(document.requestId);
  if (detail.request.status !== 'READY') {
    throw new Error(
      detail.request.status === 'ARCHIVED'
        ? 'Permohonan sudah selesai dan tidak dapat diproses ulang.'
        : 'Tandai permohonan sebagai Siap Diproses sebelum membuat dokumen.'
    );
  }

  validateDocumentGeneration_(detail.request, document, detail.employees);
  if (
    !force &&
    document.status === 'GENERATED' &&
    document.docId &&
    document.pdfId &&
    driveFileExists_(document.docId) &&
    driveFileExists_(document.pdfId)
  ) {
    return {
      ok: true,
      reused: true,
      requestId: document.requestId,
      document: document
    };
  }

  const templateId = APP_CONFIG.TEMPLATES[document.templateKey];
  if (!templateId) throw new Error('Template ID tidak ditemukan: ' + document.templateKey);
  const folder = getOutputFolder_();
  const fileName = buildDocumentFileName_(detail.request, document);
  const templateFile = DriveApp.getFileById(templateId);
  const docFile = templateFile.makeCopy(fileName, folder);
  const googleDoc = DocumentApp.openById(docFile.getId());

  try {
    replaceDocumentPlaceholders_(googleDoc, buildDocumentPlaceholders_(detail, document));
    googleDoc.saveAndClose();
  } catch (error) {
    docFile.setTrashed(true);
    throw new Error('Gagal mengisi template: ' + error.message);
  }

  const pdfFile = folder.createFile(
    DriveApp.getFileById(docFile.getId()).getBlob().getAs(MimeType.PDF)
      .setName(fileName + '.pdf')
  );
  const now = new Date();
  row[7] = 'GENERATED';
  row[8] = docFile.getId();
  row[9] = docFile.getUrl();
  row[10] = pdfFile.getId();
  row[11] = pdfFile.getUrl();
  row[16] = now;
  documentSheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);

  updateMasterDocumentLinks_(detail.request.id, docFile, pdfFile);
  recordGeneratedArtifact_(detail.request.id, document.templateKey, document.revision, 'DOC', docFile, {
    documentId: document.id
  });
  recordGeneratedArtifact_(detail.request.id, document.templateKey, document.revision, 'PDF', pdfFile, {
    documentId: document.id
  });

  return {
    ok: true,
    reused: false,
    requestId: document.requestId,
    document: documentRowToDto_(row)
  };
}

function validateDocumentGeneration_(request, document, employees) {
  const errors = [];
  if (!document.number) errors.push('Nomor surat belum diisi.');
  if (!request.letterDate) errors.push('Tanggal surat belum diisi.');
  if (!request.activityName) errors.push('Nama kegiatan belum diisi.');
  if (!request.partnerName) errors.push('Nama mitra belum diisi.');
  if (!detailSchedulesAvailable_(request)) errors.push('Jadwal kegiatan belum diisi.');
  if (document.type === 'Surat Tugas' && !employees.length) {
    errors.push('Surat Tugas membutuhkan minimal satu penerima tugas.');
  }
  if (
    document.type === 'Surat Permohonan Narasumber kepada Dekan' &&
    request.speakerStatus === 'Tidak Dicarikan' &&
    !employees.length
  ) {
    errors.push('Data narasumber belum diisi.');
  }
  if (errors.length) throw new Error(errors.join('\n'));
}

function detailSchedulesAvailable_(request) {
  return Boolean(request.startDate && request.activityPlace);
}

function buildDocumentPlaceholders_(detail, document) {
  const request = detail.request;
  const employees = detail.employees;
  const routing = computeEmailRouting_({
    documents: detail.documents,
    partnerEmail: request.partnerEmail,
    partnerName: request.partnerName,
    faculties: request.faculties,
    speakerStatus: request.speakerStatus
  }, employees, document.type);
  const people = employees.map(function(employee, index) {
    const id = employee.identifier ? ' (' + employee.identifier + ')' : '';
    return (index + 1) + '. ' + employee.name + id;
  }).join('\n');

  const values = {
    idPermohonan: request.id,
    tipeKegiatan: request.activityType,
    jenisSurat: document.type,
    subTipe: request.speakerSubtype,
    statusNarasumber: request.speakerStatus,
    fakultas: request.faculties.join(', '),
    nomorSurat: document.number,
    nomorSuratMasuk: request.incomingNumber,
    tanggalSuratMasuk: request.incomingDate ? formatIndonesianDate_(request.incomingDate) : '',
    pengirimSuratMasuk: request.incomingSender,
    perihalSuratMasuk: request.incomingSubject,
    namaKegiatan: request.activityName,
    namaMitra: request.partnerName,
    alamatMitra: request.partnerAddress,
    emailMitra: request.partnerEmail,
    tanggalSurat: request.letterDate ? formatIndonesianDate_(request.letterDate) : '',
    hari: request.day,
    tanggal: request.dateDisplay,
    waktuKegiatan: request.activityTime,
    tempat: request.activityPlace,
    namaPenandatangan: request.signerName,
    nikPenandatangan: request.signerId,
    jabatanPenandatangan: request.signerRole,
    honor: request.honor,
    perjalananDinas: request.travel,
    namaPegawai: employees.map(function(item) { return item.name; }).join('\n'),
    nipPegawai: employees.map(function(item) { return item.identifier; }).join('\n'),
    jabatanPegawai: employees.map(function(item) { return item.role; }).join('\n'),
    prodiPegawai: employees.map(function(item) { return item.unit; }).join('\n'),
    fakultasPegawai: employees.map(function(item) { return item.unit; }).join('\n'),
    emailPegawai: employees.map(function(item) { return item.email; }).join('\n'),
    narasumber: people,
    kepadaYth: routing.toRoles.join('\n'),
    tembusan: routing.ccRoles.join('\n')
  };

  const aliases = {
    'ID Permohonan': values.idPermohonan,
    'Tipe Kegiatan': values.tipeKegiatan,
    'Sub-Tipe': values.jenisSurat,
    'Nomor Surat': values.nomorSurat,
    'Nomor Surat Masuk': values.nomorSuratMasuk,
    'Tanggal Surat Masuk': values.tanggalSuratMasuk,
    'Pengirim Surat Masuk': values.pengirimSuratMasuk,
    'Perihal Surat Masuk': values.perihalSuratMasuk,
    'Nama Kegiatan': values.namaKegiatan,
    'Nama Mitra': values.namaMitra,
    'Alamat Mitra': values.alamatMitra,
    'Email Mitra': values.emailMitra,
    'Tanggal Surat dibuat': values.tanggalSurat,
    'Hari Kegiatan': values.hari,
    'Tanggal Kegiatan': values.tanggal,
    'Waktu Kegiatan': values.waktuKegiatan,
    'Tempat Kegiatan': values.tempat,
    'Nama Penandatangan Surat Tugas': values.namaPenandatangan,
    'NIK Penandatangan Surat Tugas': values.nikPenandatangan,
    'Jabatan Penandatangan Surat Tugas': values.jabatanPenandatangan,
    'Text Join Nama': values.namaPegawai,
    'Text Join NIK/NPM': values.nipPegawai,
    'Text Join Jabatan': values.jabatanPegawai,
    'Text Join Prodi': values.prodiPegawai,
    'Text Join Fakultas': values.fakultasPegawai,
    'Text Join Email': values.emailPegawai
  };
  Object.keys(aliases).forEach(function(key) { values[key] = aliases[key]; });
  return values;
}

function replaceDocumentPlaceholders_(doc, values) {
  const sections = [doc.getBody()];
  const header = doc.getHeader();
  const footer = doc.getFooter();
  if (header) sections.push(header);
  if (footer) sections.push(footer);

  Object.keys(values).forEach(function(key) {
    const replacement = String(values[key] == null ? '' : values[key])
      .replace(/\\/g, '\\\\')
      .replace(/\$/g, '\\$');
    const patterns = ['{{' + key + '}}', '{' + key + '}', '<<' + key + '>>'];
    sections.forEach(function(section) {
      patterns.forEach(function(pattern) {
        section.replaceText(escapeRegex_(pattern), replacement);
      });
    });
  });
}

function buildDocumentFileName_(request, document) {
  const rawNumber = text_(document && document.number);
  const rawType = text_(document && document.type);
  const safeNumber = rawNumber.replace(/[^A-Za-z0-9_-]+/g, '-');
  const safeType = rawType.replace(/[^A-Za-z0-9 ]+/g, '').replace(/\s+/g, '-');
  return [request.id, safeType, safeNumber].filter(Boolean).join('_').slice(0, 180);
}

function getOutputFolder_() {
  const exportSheet = getSheet_('EXPORT');
  const configured = exportSheet.getLastRow() >= 2
    ? text_(exportSheet.getRange(2, 2).getValue())
    : '';
  return DriveApp.getFolderById(configured || APP_CONFIG.OUTPUT_FOLDER_ID);
}

function driveFileExists_(fileId) {
  try {
    return !DriveApp.getFileById(fileId).isTrashed();
  } catch (error) {
    return false;
  }
}

function updateMasterDocumentLinks_(requestId, docFile, pdfFile) {
  const sheet = getSheet_('MASTER');
  const rowNumber = findRowById_(sheet, requestId, 1);
  if (!rowNumber) return;
  sheet.getRange(rowNumber, masterColumn_('Edit Surat')).setValue(docFile.getUrl());
  sheet.getRange(rowNumber, masterColumn_('Download PDF Surat')).setValue(pdfFile.getUrl());
}

function recordGeneratedArtifact_(requestId, artifactKey, revision, type, file, metadata) {
  recordGeneratedArtifactData_(
    requestId,
    artifactKey,
    revision,
    type,
    file.getId(),
    file.getUrl(),
    metadata
  );
}

function recordGeneratedArtifactData_(requestId, artifactKey, revision, type, fileId, url, metadata) {
  const sheet = getSheet_('FILES');
  appendDataRow_(sheet, [
    requestId,
    artifactKey,
    revision,
    type,
    fileId,
    url,
    new Date(),
    getCurrentUser_(),
    'ACTIVE',
    JSON.stringify(metadata || {})
  ]);
}

function supersedeGeneratedArtifacts_(requestId, artifactKey, types) {
  const sheet = getSheet_('FILES', false);
  if (!sheet) return;
  const allowedTypes = (types || []).map(function(type) {
    return text_(type).toUpperCase();
  });
  const rows = readDataRows_(sheet, GENERATED_FILE_HEADERS.length);
  let changed = false;
  rows.forEach(function(row) {
    if (
      text_(row[0]) === requestId &&
      text_(row[1]) === artifactKey &&
      text_(row[8]).toUpperCase() === 'ACTIVE' &&
      (!allowedTypes.length || allowedTypes.indexOf(text_(row[3]).toUpperCase()) !== -1)
    ) {
      row[8] = 'SUPERSEDED';
      changed = true;
    }
  });
  if (changed) rewriteDataRows_(sheet, rows, GENERATED_FILE_HEADERS.length);
}

function getGeneratedFilesByRequest_(requestId) {
  return readDataRows_(getSheet_('FILES'), GENERATED_FILE_HEADERS.length)
    .filter(function(row) { return text_(row[0]) === requestId; })
    .map(function(row) {
      return {
        requestId: row[0],
        artifactKey: row[1],
        revision: Number(row[2] || 0),
        type: row[3],
        fileId: row[4],
        url: row[5],
        createdAt: serializeValue_(row[6]),
        createdBy: row[7],
        status: row[8],
        metadata: parseJsonSafe_(row[9])
      };
    });
}

function getFinanceArtifactUrls_(requestId) {
  const result = {
    honorSheetUrl: '',
    honorPdfUrl: '',
    honorExcelUrl: '',
    perjadinSheetUrl: '',
    perjadinPdfUrl: '',
    perjadinExcelUrl: ''
  };
  getGeneratedFilesByRequest_(requestId).forEach(function(file) {
    if (text_(file.status).toUpperCase() !== 'ACTIVE') return;
    const kind = file.artifactKey === 'FINANCE_HONOR'
      ? 'honor'
      : file.artifactKey === 'FINANCE_PERJADIN'
        ? 'perjadin'
        : '';
    if (!kind) return;
    if (file.type === 'SHEET') result[kind + 'SheetUrl'] = file.url;
    const fileAvailable = typeof DriveApp === 'undefined' || driveFileExists_(file.fileId);
    if (file.type === 'PDF' && fileAvailable) {
      result[kind + 'PdfUrl'] = driveDownloadUrl_(file.fileId);
    }
    if (file.type === 'XLSX' && fileAvailable) {
      result[kind + 'ExcelUrl'] = driveDownloadUrl_(file.fileId);
    }
  });
  if (typeof SpreadsheetApp !== 'undefined') {
    [
      { kind: 'HONOR', field: 'honorSheetUrl' },
      { kind: 'PERJADIN', field: 'perjadinSheetUrl' }
    ].forEach(function(descriptor) {
      if (result[descriptor.field]) return;
      const ss = getSpreadsheet_();
      const sheet = ss.getSheetByName(generatedSheetName_(descriptor.kind, requestId));
      if (!sheet || !isGeneratedSheet_(sheet)) return;
      const metadata = sheet.getDeveloperMetadata().find(function(item) {
        return item.getKey() === 'DPSP_GENERATED';
      });
      if (metadata && metadata.getValue() === descriptor.kind + '|' + requestId) {
        result[descriptor.field] = ss.getUrl() + '#gid=' + sheet.getSheetId();
      }
    });
  }
  return result;
}

function driveDownloadUrl_(fileId) {
  const id = text_(fileId);
  return id ? 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(id) : '';
}

function parseJsonSafe_(value) {
  try {
    return value ? JSON.parse(value) : {};
  } catch (error) {
    return {};
  }
}

function markDocumentError_(documentId, message) {
  const sheet = getSheet_('DOCUMENTS');
  const rowNumber = findRowById_(sheet, documentId, 1);
  if (!rowNumber) return;
  sheet.getRange(rowNumber, 8).setValue('ERROR');
  sheet.getRange(rowNumber, 17).setValue(new Date());
  console.error('Dokumen ' + documentId + ': ' + message);
}
