function setupSystem() {
  const user = getSetupUser_();
  return withScriptLock_(function() {
    const ss = getSpreadsheet_();
    const report = [];

    ensureSheet_(ss, 'MASTER', MASTER_HEADERS, report);
    ensureSheet_(ss, 'DOCUMENTS', DOCUMENT_HEADERS, report);
    ensureSheet_(ss, 'SCHEDULES', SCHEDULE_HEADERS, report);
    ensureSheet_(ss, 'EMPLOYEES', EMPLOYEE_HEADERS, report);
    ensureSheet_(ss, 'TRAVEL', TRAVEL_HEADERS, report);
    ensureSheet_(ss, 'CC', CC_HEADERS, report);
    ensureSheet_(ss, 'EMAIL_TEMPLATE', ['Jenis Surat', 'Status Narasumber', 'Subject Email', 'Body Email', 'Place Holder'], report);
    ensureSheet_(ss, 'EXPORT', ['Peruntukan', 'Folder_ID', 'Keterangan'], report);
    ensureSheet_(ss, 'AUDIT', AUDIT_HEADERS, report);
    ensureSheet_(ss, 'FILES', GENERATED_FILE_HEADERS, report);
    ensureSheet_(ss, 'ACCESS', ACCESS_HEADERS, report);
    ensureSheet_(ss, 'SIGNATURE', SIGNATURE_HEADERS, report);
    
    // Auto-migrate TEMPLATE_CONFIG columns if legacy structure detected
    const tempSheet = getSheet_('TEMPLATE_CONFIG', false);
    if (tempSheet && tempSheet.getLastRow() >= 1) {
      const currentHeaders = tempSheet.getRange(1, 1, 1, Math.max(1, tempSheet.getLastColumn())).getValues()[0].map(function(h) {
        return String(h).trim();
      });
      if (currentHeaders.indexOf('Default CC') === -1) {
        const oldRows = tempSheet.getRange(1, 1, tempSheet.getLastRow(), Math.max(1, tempSheet.getLastColumn())).getValues();
        tempSheet.clear();
        const newRows = oldRows.map(function(row, idx) {
          if (idx === 0) return TEMPLATE_CONFIG_HEADERS;
          return [row[0], row[1], row[2], row[3], '', '', '', '', row[4], row[5]];
        });
        tempSheet.getRange(1, 1, newRows.length, TEMPLATE_CONFIG_HEADERS.length).setValues(newRows);
        report.push('Migrasi struktur TEMPLATE_CONFIG ke 10 kolom selesai.');
      }
    }
    
    ensureSheet_(ss, 'TEMPLATE_CONFIG', TEMPLATE_CONFIG_HEADERS, report);

    seedInitialAccessUser_(user.email, report);
    seedExportFolder_();
    seedEmailTemplates_();
    seedTemplateConfigs_();
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

  const targetHeaders = key === 'MASTER' && isLegacyMasterSheet_(sheet)
    ? LEGACY_MASTER_HEADERS
    : headers;
  if (sheet.getMaxColumns() < targetHeaders.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), targetHeaders.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
  sheet.setFrozenRows(1);
}

function isLegacyMasterSheet_(sheet) {
  if (!sheet || sheet.getMaxColumns() < 75) return false;
  return text_(sheet.getRange(1, 39).getDisplayValue()) === AUTOCRAT_HEADERS[0] &&
    text_(sheet.getRange(1, 74).getDisplayValue()) === AUTOCRAT_HEADERS[AUTOCRAT_HEADERS.length - 1] &&
    text_(sheet.getRange(1, 75).getDisplayValue()) === 'Email Status';
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
    '{alamatMitra}', '{waktuKegiatan}', '{textJoinNomor}', '{textJoinNama}',
    '{textJoinNikNpm}', '{textJoinJabatan}', '{textJoinProdi}', '{textJoinEmail}',
    '<<Text Join Nomor>>', '<<Text Join Nama>>', '<<Text Join NIK/NPM>>',
    '<<Text Join Jabatan>>', '<<Text Join Prodi>>', '<<Text Join Email>>'
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
  ['MASTER', 'DOCUMENTS', 'SCHEDULES', 'EMPLOYEES', 'TRAVEL', 'CC', 'EMAIL_TEMPLATE', 'EXPORT', 'AUDIT', 'FILES', 'ACCESS', 'SIGNATURE', 'TEMPLATE_CONFIG']
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

function seedTemplateConfigs_() {
  const sheet = getSheet_('TEMPLATE_CONFIG');
  if (sheet.getLastRow() > 1) return;

  const rows = [
    ['EDU_FAIR_TASK', 'Surat Tugas - Edu Fair', 'Surat Tugas', 'ST', '', '', '', '', '1GxHt4CYcsmKHjuMlLwzfpolhblG6GUtOOazaX70zuws', 'TRUE'],
    ['SPEAKER_WORKSHOP_TASK', 'Surat Tugas - Narasumber (Workshop)', 'Surat Tugas', 'ST', '', '', '', '', '1PcVlBo6Q81x9FjU6hpOECgI5Q5Pio0wD3-mwHsA4wRc', 'TRUE'],
    ['SPEAKER_PROMOTION_TASK', 'Surat Tugas - Narasumber (Promosi)', 'Surat Tugas', 'ST', '', '', '', '', '1uaEvHJxPqdhgcNTqpjG-LcssPo-7bAboaa_rsrglrQI', 'TRUE'],
    ['CAMPUS_VISIT_TASK', 'Surat Tugas - Campus Visit', 'Surat Tugas', 'ST', '', '', '', '', '1FB-0BchuGmPDEda7NyAdtGhxWv07VrMWz2YHbgR_L5o', 'TRUE'],
    ['CAMPUS_VISIT_PERMISSION', 'Surat Izin Pimpinan - Campus Visit', 'Surat izin pimpinan - Campus Visit', 'I', 'Sekretaris Universitas', 'Wakil Rektor Bidang Kerjasama, Alumni, Inovasi dan Bisnis, Manajer Aset dan Sarana Prasarana, Manajer Kemahasiswaan, Koordinator Kebersihan, Keamanan dan Ketertiban, Koordinator Kelas dan Fasilitas Pendukung', '', '', '1yjEBTJ4d-KODomgJqmFCeu6jeDhlEbF0E26Qx3mGoNQ', 'TRUE'],
    ['CAMPUS_VISIT_RECOMMENDATION', 'Surat Rekomendasi Campus Visit - SU', 'Surat Rekomendasi Campus Visit - SU', 'I', 'Sekretaris Universitas', 'Wakil Rektor Bidang Kerjasama, Alumni, Inovasi dan Bisnis', '', '', '1K_yPfDoq0ePXlJtYiLzpSBCxdOGDgubTsrtj4CeQGRA', 'TRUE'],
    ['CAMPUS_VISIT_REPLY', 'Surat Balasan Campus Visit', 'Surat Balasan Campus Visit', 'E', '', 'Wakil Rektor Bidang Kerjasama, Alumni, Inovasi dan Bisnis', '', '', '1kK5MORz9alFtAct2gAkygivqkpDTpN1MLXaGhkqVOMM', 'TRUE'],
    ['SPEAKER_REQUEST_SEARCH', 'Surat Permohonan Narasumber (Dicarikan)', 'Surat Permohonan Narasumber kepada Dekan', 'I', '', '', '', '', '1ZhLP76Ysse6IKMlO_mN-61IwMZc0iWcQqk6OMyFdSK8', 'TRUE'],
    ['SPEAKER_REQUEST_KNOWN', 'Surat Permohonan Narasumber (Sudah Ada)', 'Surat Permohonan Narasumber kepada Dekan', 'I', '', '', '', '', '1K5aik94SZW20X0TPLqIXZEmJ0p-r5ZoxJy9FO4Ywd24', 'TRUE']
  ];
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}
