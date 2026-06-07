function setupSystem() {
  const user = getSetupUser_();
  return withScriptLock_(function() {
    const ss = getSpreadsheet_();
    const report = [];

    ensureSheet_(ss, 'MASTER', MASTER_HEADERS, report);
    ensureSheet_(ss, 'DOCUMENTS', DOCUMENT_HEADERS, report);
    ensureSheet_(ss, 'EMPLOYEES', EMPLOYEE_HEADERS, report);
    ensureSheet_(ss, 'TRAVEL', TRAVEL_HEADERS, report);
    ensureSheet_(ss, 'CC', CC_HEADERS, report);
    ensureSheet_(ss, 'EMAIL_TEMPLATE', ['Jenis Surat', 'Status Narasumber', 'Subject Email', 'Body Email', 'Place Holder'], report);
    ensureSheet_(ss, 'EXPORT', ['Peruntukan', 'Folder_ID', 'Keterangan'], report);
    ensureSheet_(ss, 'AUDIT', AUDIT_HEADERS, report);
    ensureSheet_(ss, 'FILES', GENERATED_FILE_HEADERS, report);
    ensureSheet_(ss, 'ACCESS', ACCESS_HEADERS, report);
    ensureSheet_(ss, 'SIGNATURE', SIGNATURE_HEADERS, report);

    seedInitialAccessUser_(user.email, report);
    seedExportFolder_();
    seedEmailTemplates_();
    formatSystemSheets_();
    clearAppCache_();
    logAudit_('SETUP_SYSTEM', '', true, { report: report, user: user.email });
    return { ok: true, report: report };
  });
}

function getSetupUser_() {
  const email = getCurrentUser_();
  if (!email) {
    throw new Error(
      'Identitas pengguna tidak tersedia. Deploy sebagai "User accessing the web app" sebelum setup.'
    );
  }

  const accessSheet = getSheet_('ACCESS', false);
  if (!accessSheet || accessSheet.getLastRow() < 2) {
    return { email: email, role: 'ADMIN' };
  }

  const user = assertAuthorized_();
  assertAdmin_(user);
  return user;
}

function seedInitialAccessUser_(email, report) {
  const sheet = getSheet_('ACCESS');
  const rows = sheet.getLastRow() >= 2
    ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getDisplayValues()
    : [];
  const hasActiveUser = rows.some(function(row) {
    return text_(row[0]) && isAccessActive_(row[1]);
  });
  if (hasActiveUser) return;

  sheet.getRange(2, 1, 1, 3).setValues([[email, 'TRUE', 'ADMIN']]);
  report.push('Menambahkan admin awal ke Config_Access: ' + email);
}

function ensureSheet_(ss, key, headers, report) {
  const descriptor = APP_CONFIG.SHEETS[key];
  let sheet = getSheet_(key, false);
  if (!sheet) {
    sheet = ss.insertSheet(descriptor.name);
    report.push('Membuat sheet: ' + descriptor.name);
  } else if (sheet.getName() !== descriptor.name && key === 'EMPLOYEES') {
    report.push('Memakai alias lama: ' + sheet.getName());
  }

  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

function seedExportFolder_() {
  const sheet = getSheet_('EXPORT');
  if (sheet.getLastRow() < 2 || !text_(sheet.getRange(2, 2).getValue())) {
    sheet.getRange(2, 1, 1, 3).setValues([[
      'Dokumen DPSP',
      APP_CONFIG.OUTPUT_FOLDER_ID,
      'Folder output dokumen, PDF, dan ekspor website'
    ]]);
  }
}

function seedEmailTemplates_() {
  const sheet = getSheet_('EMAIL_TEMPLATE');
  if (sheet.getLastRow() > 1) return;

  const placeholders = [
    '{kepadaYth}', '{jenisSurat}', '{subTipe}', '{statusNarasumber}',
    '{namaKegiatan}', '{namaMitra}', '{narasumber}', '{hari}', '{tanggal}',
    '{tempat}', '{tembusan}', '{nomorSurat}', '{nomorSuratMasuk}',
    '{tanggalSuratMasuk}', '{pengirimSuratMasuk}', '{perihalSuratMasuk}',
    '{alamatMitra}', '{waktuKegiatan}'
  ].join(', ');

  const commonClosing = '<br><br>Hormat kami,<br><strong>Direktur Perencanaan Strategis dan Pemasaran</strong>';
  const rows = [
    [
      'Surat Tugas', '',
      'Surat Tugas | {jenisSurat}',
      'Kepada Yth.<br><strong>{kepadaYth}</strong><br><br>Bersama ini kami sampaikan Surat Tugas Nomor <strong>{nomorSurat}</strong> untuk kegiatan <strong>{namaKegiatan}</strong> pada <strong>{tanggal}</strong>.' + commonClosing,
      placeholders
    ],
    [
      'Surat Balasan Campus Visit', '',
      'Surat Balasan Campus Visit | {namaMitra}',
      'Kepada Yth.<br><strong>{kepadaYth}</strong><br><br>Menanggapi surat Nomor <strong>{nomorSuratMasuk}</strong> tanggal <strong>{tanggalSuratMasuk}</strong>, kami menyampaikan penerimaan kegiatan <strong>{namaKegiatan}</strong> pada {hari}, {tanggal}, pukul {waktuKegiatan}, di {tempat}.' + commonClosing + '<br><br>Tembusan:<br>{tembusan}',
      placeholders
    ],
    [
      'Surat Rekomendasi Campus Visit - SU', '',
      'Rekomendasi Campus Visit | {namaMitra}',
      'Kepada Yth.<br><strong>{kepadaYth}</strong><br><br>Sehubungan dengan Campus Visit <strong>{namaMitra}</strong> pada {hari}, {tanggal}, pukul {waktuKegiatan}, kami merekomendasikan kegiatan dapat dilaksanakan sesuai rencana.' + commonClosing + '<br><br>Tembusan:<br>{tembusan}',
      placeholders
    ],
    [
      'Surat izin pimpinan - Campus Visit', '',
      'Permohonan Izin Campus Visit | {namaMitra}',
      'Kepada Yth.<br><strong>{kepadaYth}</strong><br><br>Sehubungan dengan Campus Visit <strong>{namaMitra}</strong> pada {hari}, {tanggal}, pukul {waktuKegiatan}, kami memohon izin pelaksanaan kegiatan di lingkungan Universitas Katolik Parahyangan.' + commonClosing + '<br><br>Tembusan:<br>{tembusan}',
      placeholders
    ],
    [
      'Surat Permohonan Narasumber kepada Dekan', 'Tidak Dicarikan',
      'Permohonan Penugasan Narasumber | {namaMitra}',
      'Kepada Yth.<br><strong>{kepadaYth}</strong><br><br>Kami mohon dapat menugaskan:<br><br>{narasumber}<br><br>sebagai narasumber kegiatan <strong>{namaKegiatan}</strong> pada {hari}, {tanggal}, di {tempat}.' + commonClosing,
      placeholders
    ],
    [
      'Surat Permohonan Narasumber kepada Dekan', 'Dicarikan',
      'Permohonan Narasumber | {namaMitra}',
      'Kepada Yth.<br><strong>{kepadaYth}</strong><br><br>Kami mohon dapat menunjuk dosen yang sesuai sebagai narasumber kegiatan <strong>{namaKegiatan}</strong> pada {hari}, {tanggal}, di {tempat}.' + commonClosing,
      placeholders
    ]
  ];
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function formatSystemSheets_() {
  ['MASTER', 'DOCUMENTS', 'EMPLOYEES', 'TRAVEL', 'CC', 'EMAIL_TEMPLATE', 'EXPORT', 'AUDIT', 'FILES', 'ACCESS', 'SIGNATURE']
    .forEach(function(key) {
      const sheet = getSheet_(key);
      const width = Math.max(1, sheet.getLastColumn());
      sheet.getRange(1, 1, 1, width)
        .setBackground('#015850')
        .setFontColor('#ffffff')
        .setFontWeight('bold')
        .setWrap(true);
      sheet.setFrozenRows(1);
    });
}
