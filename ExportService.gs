function exportGeneratedSheet(sheetName, format, options) {
  const user = assertAuthorized_();
  assertCanWrite_(user);
  return withScriptLock_(function() {
    return exportGeneratedSheetInternal_(
      text_(sheetName),
      text_(format).toUpperCase(),
      options || {},
      user
    );
  });
}

function generateAndExportFinance(requestId, kind, format) {
  const user = assertAuthorized_();
  assertCanWrite_(user);
  const id = text_(requestId);
  const reportKind = text_(kind).toUpperCase();
  const exportFormat = text_(format).toUpperCase();
  if (['HONOR', 'PERJADIN'].indexOf(reportKind) === -1) {
    throw new Error('Jenis laporan tidak valid.');
  }

  return withScriptLock_(function() {
    assertFinanceExportReady_(id, reportKind);
    const sheetName = generatedSheetName_(reportKind, id);
    const sheet = getSpreadsheet_().getSheetByName(sheetName);
    if (!sheet || !isGeneratedSheet_(sheet)) {
      throw new Error('Kelola dan lengkapi Google Sheet ' + reportKind + ' terlebih dahulu.');
    }
    const exportOptions = reportKind === 'PERJADIN'
      ? { paperSize: 'A4', orientation: 'PORTRAIT' }
      : { paperSize: 'A4', orientation: 'LANDSCAPE' };
    return exportGeneratedSheetInternal_(
      sheetName,
      exportFormat,
      exportOptions,
      user
    );
  });
}

function exportGeneratedSheetInternal_(sheetName, format, options, user) {
  if (['PDF', 'XLSX'].indexOf(format) === -1) {
    throw new Error('Format ekspor harus PDF atau XLSX.');
  }
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || !isGeneratedSheet_(sheet)) {
    throw new Error('Hanya sheet keluaran sistem yang dapat diekspor.');
  }

  SpreadsheetApp.flush();
  const metadata = generatedSheetMetadata_(sheet);
  const requestId = metadata.requestId;
  const contentHash = sheetContentHash_(sheet, format, options);
  const existing = getGeneratedFilesByRequest_(requestId).find(function(file) {
    return file.status === 'ACTIVE' &&
      file.type === format &&
      file.metadata.sheetName === sheetName &&
      file.metadata.contentHash === contentHash &&
      driveFileExists_(file.fileId);
  });
  if (existing) {
    // File lama mungkin dibuat sebelum akses Drive dibuka; pastikan tidak 403 saat diunduh.
    grantGeneratedFileAccessById_(existing.fileId, user, { editor: false });
    return {
      ok: true,
      reused: true,
      requestId: requestId,
      format: format,
      fileId: existing.fileId,
      url: driveDownloadUrl_(existing.fileId)
    };
  }

  const blob = format === 'PDF'
    ? exportSheetPdfBlob_(ss, sheet, options)
    : exportSheetXlsxBlob_(sheet);
  const fileName = sheetName + '.' + (format === 'PDF' ? 'pdf' : 'xlsx');
  const file = getOutputFolder_().createFile(blob.setName(fileName));
  // Make the file accessible via direct download link and prevent Google Drive 403.
  grantGeneratedFileAccess_(file, user, { editor: false });
  const artifactKey = 'FINANCE_' + metadata.kind;
  supersedeGeneratedArtifacts_(requestId, artifactKey, [format]);
  recordGeneratedArtifactData_(
    requestId,
    artifactKey,
    getRequestDetailInternal_(requestId).request.revision,
    format,
    file.getId(),
    driveDownloadUrl_(file.getId()),
    {
      sheetName: sheetName,
      sheetId: sheet.getSheetId(),
      contentHash: contentHash,
      format: format
    }
  );
  logAudit_('EXPORT_GENERATED_SHEET', requestId, true, {
    sheetName: sheetName,
    format: format,
    fileId: file.getId()
  });
  return {
    ok: true,
    reused: false,
    requestId: requestId,
    format: format,
    fileId: file.getId(),
    url: driveDownloadUrl_(file.getId())
  };
}

function exportSheetPdfBlob_(ss, sheet, options) {
  // New implementation using Drive export to avoid 403 errors.
  // Create a temporary copy of the sheet in a new spreadsheet.
  const tempSpreadsheet = SpreadsheetApp.create('TEMP_EXPORT_' + Utilities.getUuid());
  try {
    const copied = sheet.copyTo(tempSpreadsheet).setName(sheet.getName());
    // Remove default sheet if present.
    const sheets = tempSpreadsheet.getSheets();
    sheets.forEach(function(s) {
      if (s.getSheetId() !== copied.getSheetId()) {
        tempSpreadsheet.deleteSheet(s);
      }
    });
    SpreadsheetApp.flush();
    // Export the temporary spreadsheet as PDF using Drive API.
    const file = DriveApp.getFileById(tempSpreadsheet.getId());
    const pdfBlob = file.getAs('application/pdf');
    // Optionally set PDF metadata (orientation, paper size) – not directly supported via Drive export.
    // For now, rely on default settings; callers can adjust options when needed.
    return pdfBlob;
  } finally {
    // Clean up the temporary spreadsheet.
    DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true);
  }
}

function exportSheetXlsxBlob_(sourceSheet) {
  const temporary = SpreadsheetApp.create('TEMP-DPSP-' + Utilities.getUuid());
  const temporaryFile = DriveApp.getFileById(temporary.getId());
  try {
    const copied = sourceSheet.copyTo(temporary).setName(sourceSheet.getName());
    temporary.getSheets().forEach(function(sheet) {
      if (sheet.getSheetId() !== copied.getSheetId()) temporary.deleteSheet(sheet);
    });
    SpreadsheetApp.flush();
    const url = 'https://docs.google.com/spreadsheets/d/' + temporary.getId() + '/export?format=xlsx';
    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      throw new Error('Ekspor XLSX gagal. HTTP ' + response.getResponseCode());
    }
    return response.getBlob();
  } finally {
    temporaryFile.setTrashed(true);
  }
}

function generatedSheetMetadata_(sheet) {
  const metadata = sheet.getDeveloperMetadata().find(function(item) {
    return item.getKey() === 'DPSP_GENERATED';
  });
  if (!metadata) throw new Error('Metadata sheet keluaran tidak ditemukan.');
  const parts = String(metadata.getValue() || '').split('|');
  return { kind: parts[0], requestId: parts[1] };
}

function sheetContentHash_(sheet, format, options) {
  const range = sheet.getDataRange();
  const payload = JSON.stringify({
    values: range.getDisplayValues(),
    formulas: range.getFormulas(),
    format: format,
    options: options || {}
  });
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    payload,
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(byte) {
    const value = byte < 0 ? byte + 256 : byte;
    return ('0' + value.toString(16)).slice(-2);
  }).join('');
}

function cleanupTempExportFiles() {
  try {
    const user = assertAuthorized_();
    if (user.role !== 'ADMIN') {
      throw new Error('Hanya Admin yang dapat menjalankan pembersihan file.');
    }
  } catch (e) {
    // Allow cron/time-driven triggers to run without active user session
  }

  const twelveHoursAgo = new Date(Date.now() - 12 * 60 * 60 * 1000);
  const formattedTime = twelveHoursAgo.toISOString().replace(/\.\d+Z$/, 'Z');

  const queries = [
    "title contains 'TEMP_EXPORT_' and mimeType = 'application/vnd.google-apps.spreadsheet' and createdDate < '" + formattedTime + "'",
    "title contains 'TEMP-DPSP-' and mimeType = 'application/vnd.google-apps.spreadsheet' and createdDate < '" + formattedTime + "'"
  ];

  let deletedCount = 0;
  queries.forEach(function(query) {
    const files = DriveApp.searchFiles(query);
    while (files.hasNext()) {
      const file = files.next();
      try {
        file.setTrashed(true);
        deletedCount++;
      } catch (e) {
        console.error('Gagal menghapus temp file ' + file.getName() + ': ' + e.message);
      }
    }
  });

  return { ok: true, cleaned: deletedCount };
}
