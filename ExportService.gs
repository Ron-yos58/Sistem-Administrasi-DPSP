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
    const generated = generateFinanceSheetInternal_(id, reportKind);
    return exportGeneratedSheetInternal_(
      generated.sheetName,
      exportFormat,
      { paperSize: 'A4', orientation: 'LANDSCAPE' },
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
    return file.type === format &&
      file.metadata.sheetName === sheetName &&
      file.metadata.contentHash === contentHash &&
      driveFileExists_(file.fileId);
  });
  if (existing) {
    return {
      ok: true,
      reused: true,
      requestId: requestId,
      format: format,
      fileId: existing.fileId,
      url: existing.url
    };
  }

  const blob = format === 'PDF'
    ? exportSheetPdfBlob_(ss, sheet, options)
    : exportSheetXlsxBlob_(sheet);
  const fileName = sheetName + '.' + (format === 'PDF' ? 'pdf' : 'xlsx');
  const file = getOutputFolder_().createFile(blob.setName(fileName));
  recordGeneratedArtifact_(
    requestId,
    'FINANCE_' + metadata.kind,
    1,
    format,
    file,
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
    url: file.getUrl()
  };
}

function exportSheetPdfBlob_(ss, sheet, options) {
  const orientation = text_(options.orientation).toUpperCase() || 'LANDSCAPE';
  const paperSize = text_(options.paperSize).toUpperCase() || 'A4';
  const params = {
    format: 'pdf',
    gid: sheet.getSheetId(),
    size: paperSize,
    portrait: orientation === 'PORTRAIT',
    fitw: true,
    sheetnames: false,
    printtitle: false,
    pagenumbers: true,
    gridlines: false,
    fzr: false,
    top_margin: 0.5,
    bottom_margin: 0.5,
    left_margin: 0.5,
    right_margin: 0.5
  };
  const query = Object.keys(params).map(function(key) {
    return encodeURIComponent(key) + '=' + encodeURIComponent(params[key]);
  }).join('&');
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?' + query;
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() !== 200) {
    throw new Error('Ekspor PDF gagal. HTTP ' + response.getResponseCode());
  }
  return response.getBlob();
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
