function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_CONFIG.APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSystemStatus() {
  const user = assertAuthorized_();
  const checks = {};
  const addCheck = function(key, callback) {
    try {
      checks[key] = Object.assign({ ok: true }, callback());
    } catch (error) {
      checks[key] = { ok: false, error: error.message };
    }
  };

  addCheck('SPREADSHEET', function() {
    const ss = getSpreadsheet_();
    return { name: ss.getName(), id: ss.getId() };
  });

  addCheck('OUTPUT_FOLDER', function() {
    const folder = getOutputFolder_();
    return { name: folder.getName(), id: folder.getId() };
  });

  Object.keys(APP_CONFIG.SHEETS).forEach(function(key) {
    addCheck('SHEET_' + key, function() {
      const sheet = getSheet_(key, false);
      if (!sheet) throw new Error('Sheet tidak ditemukan: ' + APP_CONFIG.SHEETS[key].name);
      return { name: sheet.getName(), rows: sheet.getLastRow() };
    });
  });

  addCheck('ACCESS_ROLES', function() {
    const rows = readDataRows_(getSheet_('ACCESS'), ACCESS_HEADERS.length);
    let admins = 0;
    let operators = 0;
    rows.forEach(function(row) {
      if (!text_(row[0]) || !isAccessActive_(row[1])) return;
      const role = normalizeRole_(row[2]);
      if (role === 'ADMIN') admins += 1;
      if (role === 'OPERATOR') operators += 1;
    });
    if (!admins) throw new Error('Config_Access harus memiliki minimal satu ADMIN aktif.');
    return { admins: admins, operators: operators };
  });

  addCheck('EMAIL_TEMPLATES', function() {
    const rows = readDataRows_(getSheet_('EMAIL_TEMPLATE'), 5);
    const types = uniqueTextList_(rows.map(function(row) { return row[0]; }));
    const required = [
      'Surat Tugas',
      'Surat Balasan Campus Visit',
      'Surat Rekomendasi Campus Visit - SU',
      'Surat izin pimpinan - Campus Visit',
      'Surat Permohonan Narasumber kepada Dekan'
    ];
    const missing = required.filter(function(type) { return types.indexOf(type) === -1; });
    if (missing.length) throw new Error('Template email belum lengkap: ' + missing.join(', '));
    return { rows: rows.length, types: types.length };
  });

  Object.keys(APP_CONFIG.TEMPLATES).forEach(function(key) {
    addCheck('TEMPLATE_' + key, function() {
      const file = DriveApp.getFileById(APP_CONFIG.TEMPLATES[key]);
      return { name: file.getName(), id: file.getId() };
    });
  });

  return serializeValue_({
    ok: Object.keys(checks).every(function(key) { return checks[key].ok; }),
    user: user,
    version: APP_CONFIG.APP_VERSION,
    checks: checks
  });
}
