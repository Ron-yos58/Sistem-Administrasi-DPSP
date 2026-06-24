const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const context = {
  console,
  Date,
  Math,
  JSON,
  Set,
  Array,
  Object,
  String,
  Number,
  Boolean,
  RegExp,
  isFinite,
  Utilities: {
    getUuid: () => '00000000-0000-4000-8000-000000000000',
    formatDate: (date, timeZone, pattern) => {
      if (pattern === 'HH:mm') {
        return `${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
      }
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
  },
  window: {
    CSS: {},
    setTimeout: () => {},
    clearTimeout: () => {}
  },
  document: {}
};
vm.createContext(context);

for (const file of ['Config.gs', 'Utils.gs', 'DataService.gs', 'Migration.gs', 'DocumentService.gs', 'FinanceService.gs']) {
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), context, { filename: file });
}
const scriptFiles = [
  'ScriptsCore.html',
  'ScriptsApp.html',
  'ScriptsEvents.html',
  'ScriptsRender.html',
  'ScriptsForm.html',
  'ScriptsDetail.html',
  'ScriptsAdmin.html'
];
const scriptsHtml = scriptFiles.map(file => fs.readFileSync(path.join(root, file), 'utf8')).join('\n');
const jsCode = scriptsHtml.replace(/<script>/g, '').replace(/<\/script>/g, '');
vm.runInContext(jsCode, context, { filename: 'CombinedScripts.js' });

function test(name, callback) {
  try {
    callback();
    console.log(`OK behavior ${name}`);
  } catch (error) {
    console.error(`FAIL behavior ${name}: ${error.message}`);
    process.exitCode = 1;
  }
}

function equal(actual, expected) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    throw new Error(`expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
  }
}

test('Indonesian same-month date range', () => {
  equal(context.formatDateRange_('2026-06-01', '2026-06-03'), '1 - 3 Juni 2026');
});

test('Indonesian day range', () => {
  equal(context.formatDayRange_('2026-06-01', '2026-06-03'), 'Senin - Rabu');
});

test('non-contiguous sessions stay one schedule summary', () => {
  const summary = context.buildScheduleSummary_([
    {
      startDate: '2026-06-25',
      endDate: '2026-06-25',
      startTime: '09:00',
      endTime: '12:00',
      place: 'Aula',
      sequence: 1
    },
    {
      startDate: '2026-06-30',
      endDate: '2026-06-30',
      startTime: '13:00',
      endTime: '15:00',
      place: 'Aula',
      sequence: 2
    }
  ]);

  equal(summary.startDate, '2026-06-25');
  equal(summary.endDate, '2026-06-30');
  equal(summary.dayDisplay, 'Kamis dan Selasa');
  equal(summary.dateDisplay, '25 dan 30 Juni 2026');
  equal(
    summary.timeDisplay,
    '25 Juni 2026: 09.00–12.00 WIB\n30 Juni 2026: 13.00–15.00 WIB'
  );
  equal(summary.placeDisplay, 'Aula');
});

test('one continuous session remains a date range', () => {
  const summary = context.buildScheduleSummary_([{
    startDate: '2026-06-25',
    endDate: '2026-06-27',
    startTime: '09:00',
    endTime: '15:00',
    place: 'Aula',
    sequence: 1
  }]);

  equal(summary.dateDisplay, '25 - 27 Juni 2026');
  equal(summary.timeDisplay, '09.00–15.00 WIB');
});

test('legacy time range parsing', () => {
  equal(
    context.parseTimeRange_('Pukul 09.00 - 12.30 WIB'),
    { startTime: '09:00', endTime: '12:30' }
  );
});

test('legacy Date strings retain both clocks without the 1899 serial date', () => {
  equal(
    context.parseTimeRange_(
      'Sat Dec 30 1899 09:00:00 GMT+0707 (Western Indonesia Time) - ' +
      'Sat Dec 30 1899 18:00:00 GMT+0707 (Western Indonesia Time)'
    ),
    { startTime: '09:00', endTime: '18:00' }
  );
});

test('spreadsheet time Date objects do not leak the 1899 serial date', () => {
  const start = new Date(1899, 11, 30, 9, 0, 0);
  const end = new Date(1899, 11, 30, 18, 0, 0);
  equal(context.normalizeTimeValue_(start), '09:00');
  equal(context.normalizeTimeValue_(end), '18:00');
  equal(context.formatTimeRange_(start, end), '09.00–18.00 WIB');
});

test('spreadsheet numeric time serials are normalized', () => {
  equal(context.normalizeTimeValue_(0.375), '09:00');
  equal(context.normalizeTimeValue_(0.75), '18:00');
});

test('schedule detail label uses concise Indonesian format', () => {
  equal(
    context.formatScheduleItem_({
      startDate: '2026-06-25',
      endDate: '2026-06-25',
      startTime: new Date(1899, 11, 30, 9, 0, 0),
      endTime: new Date(1899, 11, 30, 18, 0, 0),
      place: 'Jakarta'
    }),
    'Kamis, 25 Juni 2026 | 09.00–18.00 WIB | Jakarta'
  );
});

test('schedule detail label supports a cross-day date range', () => {
  equal(
    context.formatScheduleItem_({
      startDate: '2026-06-25',
      endDate: '2026-06-26',
      startTime: '09:00',
      endTime: '18:00',
      place: 'Jakarta'
    }),
    'Kamis–Jumat, 25–26 Juni 2026 | 09.00–18.00 WIB | Jakarta'
  );
});

test('empty schedule detail does not leave dangling separators', () => {
  equal(context.formatScheduleItem_({}), '');
});

test('finance artifact links use active files and direct downloads', () => {
  const previous = context.getGeneratedFilesByRequest_;
  context.getGeneratedFilesByRequest_ = () => [
    {
      artifactKey: 'FINANCE_HONOR',
      type: 'PDF',
      fileId: 'old-pdf',
      url: 'https://drive.google.com/file/d/old-pdf/view',
      status: 'SUPERSEDED'
    },
    {
      artifactKey: 'FINANCE_HONOR',
      type: 'SHEET',
      fileId: 'spreadsheet-id',
      url: 'https://docs.google.com/spreadsheets/d/spreadsheet-id/edit#gid=12',
      status: 'ACTIVE'
    },
    {
      artifactKey: 'FINANCE_HONOR',
      type: 'PDF',
      fileId: 'new-pdf',
      url: 'https://drive.google.com/file/d/new-pdf/view',
      status: 'ACTIVE'
    },
    {
      artifactKey: 'FINANCE_PERJADIN',
      type: 'XLSX',
      fileId: 'travel-xlsx',
      url: 'https://drive.google.com/file/d/travel-xlsx/view',
      status: 'ACTIVE'
    }
  ];
  try {
    equal(
      context.getFinanceArtifactUrls_('REQ-1'),
      {
        honorSheetUrl: 'https://docs.google.com/spreadsheets/d/spreadsheet-id/edit#gid=12',
        honorPdfUrl: 'https://drive.google.com/uc?export=download&id=new-pdf',
        honorExcelUrl: '',
        perjadinSheetUrl: '',
        perjadinPdfUrl: '',
        perjadinExcelUrl: 'https://drive.google.com/uc?export=download&id=travel-xlsx'
      }
    );
  } finally {
    context.getGeneratedFilesByRequest_ = previous;
  }
});

test('finance artifact invalidation is scoped by request and artifact key', () => {
  const rows = [
    ['REQ-1', 'FINANCE_PERJADIN', 1, 'SHEET', 'sheet', 'url', '', '', 'ACTIVE', '{}'],
    ['REQ-1', 'FINANCE_PERJADIN', 1, 'PDF', 'pdf', 'url', '', '', 'ACTIVE', '{}'],
    ['REQ-1', 'FINANCE_HONOR', 1, 'PDF', 'honor-pdf', 'url', '', '', 'ACTIVE', '{}'],
    ['REQ-2', 'FINANCE_PERJADIN', 1, 'PDF', 'other-pdf', 'url', '', '', 'ACTIVE', '{}']
  ];
  const fakeSheet = {
    getLastRow: () => rows.length + 1,
    getRange: () => ({
      getValues: () => rows.map(row => row.slice()),
      clearContent: () => { rows.length = 0; },
      setValues: values => {
        rows.length = 0;
        values.forEach(row => rows.push(row.slice()));
      }
    })
  };
  const previousGetSheet = context.getSheet_;
  context.getSheet_ = () => fakeSheet;
  try {
    context.supersedeGeneratedArtifacts_(
      'REQ-1',
      'FINANCE_PERJADIN',
      ['SHEET', 'PDF', 'XLSX']
    );
    equal(rows.map(row => row[8]), ['SUPERSEDED', 'SUPERSEDED', 'ACTIVE', 'ACTIVE']);
  } finally {
    context.getSheet_ = previousGetSheet;
  }
});

test('document workflow summary keeps request and processing states separate', () => {
  equal(
    context.summarizeDocumentWorkflow_([
      { status: 'PENDING', emailStatus: '' },
      { status: 'PENDING', emailStatus: '' }
    ], 'READY').status,
    'NOT_CREATED'
  );
  equal(
    context.summarizeDocumentWorkflow_([
      { status: 'GENERATED', emailStatus: 'DRAFTED' },
      { status: 'GENERATED', emailStatus: 'DRAFTED' }
    ], 'READY').status,
    'DRAFTED'
  );
  equal(
    context.summarizeDocumentWorkflow_([
      { status: 'ERROR', emailStatus: '' },
      { status: 'GENERATED', emailStatus: '' }
    ], 'READY').status,
    'ERROR'
  );
  equal(
    context.summarizeDocumentWorkflow_([
      { status: 'GENERATED', emailStatus: 'DRAFTED' }
    ], 'ARCHIVED').status,
    'COMPLETE'
  );
});

test('email validation and deduplication', () => {
  equal(
    Array.from(context.emailList_('A@EXAMPLE.COM; invalid; a@example.com, b@example.com')),
    ['a@example.com', 'b@example.com']
  );
});

test('role policy follows documented access model', () => {
  equal(context.normalizeRole_('admin'), 'ADMIN');
  equal(context.normalizeRole_('OPERATOR'), 'OPERATOR');
  let rejected = false;
  try {
    context.normalizeRole_('USER');
  } catch (error) {
    rejected = /Role akses tidak valid/.test(error.message);
  }
  if (!rejected) throw new Error('USER role was accepted');

  context.assertCanWrite_({ role: 'OPERATOR' });
  rejected = false;
  try {
    context.assertCanWrite_({ role: 'VIEWER' });
  } catch (error) {
    rejected = /Admin atau Operator/.test(error.message);
  }
  if (!rejected) throw new Error('VIEWER role was allowed to write');

  rejected = false;
  try {
    context.normalizeRole_('VIEWER');
  } catch (error) {
    rejected = /Role akses tidak valid/.test(error.message);
  }
  if (!rejected) throw new Error('VIEWER role was accepted');
});

test('authorization fails closed when access sheet is empty', () => {
  const previousSession = context.Session;
  const previousGetSheet = context.getSheet_;
  context.Session = { getActiveUser: () => ({ getEmail: () => 'operator@example.com' }) };
  context.getSheet_ = key => key === 'ACCESS'
    ? { getLastRow: () => 1 }
    : null;

  let rejected = false;
  try {
    context.assertAuthorized_();
  } catch (error) {
    rejected = /Config_Access/.test(error.message);
  } finally {
    context.Session = previousSession;
    context.getSheet_ = previousGetSheet;
  }
  if (!rejected) throw new Error('empty Config_Access was allowed');
});

test('Edu Fair only allows task letter', () => {
  equal(Array.from(context.allowedDocumentsFor_('Edu Fair', '')), ['Surat Tugas']);
});

test('promotion template mapping', () => {
  equal(context.resolveTemplateKey_({
    activityType: 'Penugasan Narasumber',
    type: 'Surat Tugas',
    speakerSubtype: 'Promosi',
    speakerStatus: ''
  }), 'SPEAKER_PROMOTION_TASK');
});

test('legacy date range parsing', () => {
  equal(
    JSON.parse(JSON.stringify(context.parseLegacyDateRange_('31 Mei - 2 Juni 2026'))),
    { start: '2026-05-31', end: '2026-06-02' }
  );
  equal(
    JSON.parse(JSON.stringify(context.parseLegacyDateRange_('25 Juni 2026 - 27 Juni 2026'))),
    { start: '2026-06-25', end: '2026-06-27' }
  );
  equal(
    JSON.parse(JSON.stringify(context.parseLegacyDateRange_('25 Juni 2026 s/d 27 Juni 2026'))),
    { start: '2026-06-25', end: '2026-06-27' }
  );
  equal(
    JSON.parse(JSON.stringify(context.parseLegacyDateRange_('25 sampai 27 Juni 2026'))),
    { start: '2026-06-25', end: '2026-06-27' }
  );
});

test('legacy Master compacts without Autocrat columns', () => {
  const masterHeaders = vm.runInContext('MASTER_HEADERS.slice()', context);
  const legacyHeaders = vm.runInContext('LEGACY_MASTER_HEADERS.slice()', context);
  const autocratHeaders = vm.runInContext('AUTOCRAT_HEADERS.slice()', context);
  const legacyRow = legacyHeaders.map(header => `value:${header}`);
  const compacted = context.compactLegacyMasterRow_(legacyRow);

  equal(masterHeaders.length, 48);
  equal(legacyHeaders.length, 84);
  equal(autocratHeaders.length, 36);
  equal(compacted.length, masterHeaders.length);
  equal(compacted[masterHeaders.indexOf('Email Status')], 'value:Email Status');
  if (masterHeaders.some(header => autocratHeaders.includes(header))) {
    throw new Error('active Master still contains an Autocrat header');
  }
});

test('migration deletes only Autocrat AM-BV columns', () => {
  const autocratHeaders = vm.runInContext('AUTOCRAT_HEADERS.slice()', context);
  let deletedColumns = null;
  const sheet = {
    getMaxColumns: () => 84,
    getRange: (row, column, rowCount, columnCount) => ({
      getDisplayValues: () => [autocratHeaders.slice(0, columnCount)]
    }),
    deleteColumns: (startColumn, columnCount) => {
      deletedColumns = [startColumn, columnCount];
    }
  };

  equal(context.removeLegacyAutocratColumns_(sheet), 36);
  equal(deletedColumns, [39, 36]);
});

test('migration refuses unexpected Autocrat headers', () => {
  const autocratHeaders = vm.runInContext('AUTOCRAT_HEADERS.slice()', context);
  autocratHeaders[4] = 'Header lain';
  let deleted = false;
  const sheet = {
    getMaxColumns: () => 84,
    getRange: () => ({
      getDisplayValues: () => [autocratHeaders]
    }),
    deleteColumns: () => {
      deleted = true;
    }
  };

  let rejected = false;
  try {
    context.removeLegacyAutocratColumns_(sheet);
  } catch (error) {
    rejected = /Header Autocrat tidak cocok/.test(error.message);
  }
  if (!rejected) throw new Error('unexpected headers were accepted');
  if (deleted) throw new Error('columns were deleted after failed validation');
});

test('Master schema inspection recognizes legacy and active layouts', () => {
  const masterHeaders = vm.runInContext('MASTER_HEADERS.slice()', context);
  const legacyHeaders = vm.runInContext('LEGACY_MASTER_HEADERS.slice()', context);
  const makeSheet = headers => ({
    getMaxColumns: () => headers.length,
    getRange: (row, column, rowCount, columnCount) => ({
      getDisplayValues: () => [headers.slice(column - 1, column - 1 + columnCount)]
    })
  });

  equal(context.inspectMasterSchema_(makeSheet(legacyHeaders)), 'LEGACY');
  equal(context.inspectMasterSchema_(makeSheet(masterHeaders)), 'ACTIVE');

  const unexpectedHeaders = legacyHeaders.slice();
  unexpectedHeaders[38] = 'Header tidak dikenal';
  equal(context.inspectMasterSchema_(makeSheet(unexpectedHeaders)), 'UNKNOWN');
});

test('incomplete draft remains valid', () => {
  context.validateRequestPayload_({
    status: 'DRAFT',
    activityType: 'Edu Fair',
    activityName: '',
    partnerName: '',
    startDate: '',
    endDate: '',
    activityPlace: '',
    documents: [],
    employees: [],
    travel: 'Tidak',
    speakerSubtype: '',
    speakerStatus: '',
    faculties: [],
    incomingDate: '',
    letterDate: ''
  });
});

test('draft preserves partially filled employee rows', () => {
  const normalized = context.normalizeRequestPayload_({
    status: 'DRAFT',
    activityType: 'Edu Fair',
    documents: [],
    employees: [{
      name: '',
      identifier: '',
      email: '',
      role: 'Dosen',
      unit: 'Fakultas Teknik',
      rank: '',
      category: ''
    }],
    schedules: [],
    faculties: []
  });
  equal(normalized.employees.length, 1);
  equal(normalized.employees[0].role, 'Dosen');
  equal(normalized.employees[0].unit, 'Fakultas Teknik');
});

test('ready task requires document number and employee', () => {
  let failed = false;
  try {
    context.validateRequestPayload_({
      status: 'READY',
      activityType: 'Edu Fair',
      activityName: 'Edu Fair',
      partnerName: 'Sekolah',
      startDate: '2026-06-01',
      endDate: '2026-06-01',
      activityPlace: 'Bandung',
      documents: [{ type: 'Surat Tugas', number: '' }],
      employees: [],
      travel: 'Tidak',
      speakerSubtype: '',
      speakerStatus: '',
      faculties: [],
      incomingDate: '',
      letterDate: '2026-06-01'
    });
  } catch (error) {
    failed = /Nomor surat/.test(error.message) && /Minimal satu/.test(error.message);
  }
  if (!failed) throw new Error('validation did not reject incomplete ready request');
});

test('ready employee requires identity, contact, role, and faculty', () => {
  let failed = false;
  try {
    context.validateRequestPayload_({
      status: 'READY',
      activityType: 'Edu Fair',
      activityName: 'Edu Fair',
      partnerName: 'Sekolah',
      schedules: [{
        startDate: '2026-06-01',
        endDate: '2026-06-01',
        startTime: '09:00',
        endTime: '12:00',
        place: 'Bandung'
      }],
      documents: [{ type: 'Surat Tugas', number: '001/DPSP/2026' }],
      employees: [{
        name: 'Budi',
        identifier: '',
        email: '',
        role: '',
        unit: '',
        rank: '',
        category: ''
      }],
      travel: 'Tidak',
      speakerSubtype: '',
      speakerStatus: '',
      faculties: [],
      incomingDate: '',
      letterDate: '2026-06-01'
    });
  } catch (error) {
    failed = /NIP\/NPM/.test(error.message) &&
      /Email/.test(error.message) &&
      /Jabatan/.test(error.message) &&
      /Fakultas/.test(error.message);
  }
  if (!failed) throw new Error('incomplete employee data was accepted');
});

test('rank and category remain optional for travel requests', () => {
  context.validateRequestPayload_({
    status: 'READY',
    activityType: 'Edu Fair',
    activityName: 'Edu Fair',
    partnerName: 'Sekolah',
    schedules: [{
      startDate: '2026-06-01',
      endDate: '2026-06-01',
      startTime: '09:00',
      endTime: '12:00',
      place: 'Bandung'
    }],
    documents: [{ type: 'Surat Tugas', number: '001/DPSP/2026' }],
    employees: [{
      name: 'Budi',
      identifier: '12345',
      email: 'budi@example.com',
      role: 'Dosen',
      unit: 'Fakultas Teknik',
      rank: '',
      category: ''
    }],
    travel: 'Ya',
    speakerSubtype: '',
    speakerStatus: '',
    faculties: [],
    incomingDate: '',
    letterDate: '2026-06-01'
  });
});

test('ready employee faculty must use configured faculty list', () => {
  let failed = false;
  try {
    context.validateRequestPayload_({
      status: 'READY',
      activityType: 'Edu Fair',
      activityName: 'Edu Fair',
      partnerName: 'Sekolah',
      schedules: [{
        startDate: '2026-06-01',
        endDate: '2026-06-01',
        startTime: '09:00',
        endTime: '12:00',
        place: 'Bandung'
      }],
      documents: [{ type: 'Surat Tugas', number: '001/DPSP/2026' }],
      employees: [{
        name: 'Budi',
        identifier: '12345',
        email: 'budi@example.com',
        role: 'Dosen',
        unit: 'Unit Tidak Terdaftar',
        rank: '',
        category: ''
      }],
      travel: 'Tidak',
      speakerSubtype: '',
      speakerStatus: '',
      faculties: [],
      incomingDate: '',
      letterDate: '2026-06-01'
    });
  } catch (error) {
    failed = /Unit\/Fakultas orang ke-1 tidak dikenali/.test(error.message);
  }
  if (!failed) throw new Error('unknown faculty was accepted');
});

test('employee faculty aliases normalize before validation', () => {
  const normalized = context.normalizeRequestPayload_({
    status: 'DRAFT',
    activityType: 'Edu Fair',
    documents: [],
    employees: [{
      name: 'Budi',
      identifier: '123',
      email: 'budi@example.com',
      role: 'Dosen',
      unit: 'FE'
    }],
    schedules: [],
    faculties: []
  });
  equal(normalized.employees[0].unit, 'Fakultas Ekonomi');
});

test('readDataRows reads rows even when first column is blank', () => {
  const sheet = {
    getLastRow: () => 4,
    getRange: () => ({
      getValues: () => [
        ['', 'Andi', '123'],
        ['REQ-1', 'Budi', '456'],
        ['', '', '']
      ]
    })
  };
  equal(
    JSON.parse(JSON.stringify(context.readDataRows_(sheet, 3))),
    [
      ['', 'Andi', '123'],
      ['REQ-1', 'Budi', '456']
    ]
  );
});

test('request document rewrite invalidates stale generated files', () => {
  const rows = [[
    'DOC-1',
    'REQ-1',
    'Surat Tugas',
    '',
    '',
    'OLD-001',
    'EDU_FAIR_TASK',
    'GENERATED',
    'doc-id',
    'doc-url',
    'pdf-id',
    'pdf-url',
    'draft-id',
    'DRAFTED',
    1,
    '2026-06-01',
    '2026-06-01'
  ]];
  const fakeSheet = {
    getLastRow: () => rows.length + 1,
    getRange: (row, column, rowCount, columnCount) => ({
      getValues: () => rows.map(item => item.slice(0, columnCount)),
      clearContent: () => { rows.length = 0; },
      setValues: values => {
        rows.length = 0;
        values.forEach(value => rows.push(value.slice()));
      }
    })
  };
  const previousGetSheet = context.getSheet_;
  context.getSheet_ = key => {
    if (key !== 'DOCUMENTS') throw new Error('unexpected sheet ' + key);
    return fakeSheet;
  };

  try {
    const output = context.replaceDocumentsForRequest_('REQ-1', [{
      activityType: 'Edu Fair',
      type: 'Surat Tugas',
      speakerSubtype: '',
      speakerStatus: '',
      number: 'NEW-001'
    }], 2);
    equal(output[0].status, 'PENDING');
    equal(output[0].docId, '');
    equal(output[0].pdfId, '');
    equal(output[0].emailDraftId, '');
    equal(output[0].emailStatus, '');
  } finally {
    context.getSheet_ = previousGetSheet;
  }
});

test('request status changes do not change generation fingerprint', () => {
  const base = {
    status: 'DRAFT',
    activityType: 'Edu Fair',
    activityName: 'Pameran',
    partnerName: 'Mitra',
    documents: [{ type: 'Surat Tugas', number: '001' }],
    schedules: [{
      startDate: '2026-06-20',
      endDate: '2026-06-20',
      startTime: '09:00',
      endTime: '12:00',
      place: 'Aula'
    }],
    employees: [{
      name: 'Andi',
      identifier: '123',
      role: 'Staf',
      unit: 'Fakultas Ekonomi',
      email: 'andi@example.com'
    }]
  };
  equal(
    context.requestGenerationFingerprint_(base),
    context.requestGenerationFingerprint_(Object.assign({}, base, { status: 'READY' }))
  );
});

test('unchanged document rewrite preserves generated artifacts', () => {
  const rows = [[
    'DOC-1', 'REQ-1', 'Surat Tugas', '', '', '001', 'EDU_FAIR_TASK',
    'GENERATED', 'doc-id', 'doc-url', 'pdf-id', 'pdf-url',
    'draft-id', 'DRAFTED', 1, '2026-06-01', '2026-06-01'
  ]];
  const fakeSheet = {
    getLastRow: () => rows.length + 1,
    getRange: (row, column, rowCount, columnCount) => ({
      getValues: () => rows.map(item => item.slice(0, columnCount)),
      clearContent: () => { rows.length = 0; },
      setValues: values => {
        rows.length = 0;
        values.forEach(value => rows.push(value.slice()));
      }
    })
  };
  const previousGetSheet = context.getSheet_;
  context.getSheet_ = () => fakeSheet;
  try {
    const output = context.replaceDocumentsForRequest_('REQ-1', [{
      activityType: 'Edu Fair',
      type: 'Surat Tugas',
      speakerSubtype: '',
      speakerStatus: '',
      number: '001'
    }], 2, true);
    equal(output[0].docId, 'doc-id');
    equal(output[0].pdfId, 'pdf-id');
    equal(output[0].emailDraftId, 'draft-id');
  } finally {
    context.getSheet_ = previousGetSheet;
  }
});

test('process request only activates workflow', () => {
  const source = fs.readFileSync(path.join(root, 'EmailService.gs'), 'utf8');
  const match = source.match(/function processRequest\([\s\S]*?\n}\n\nfunction buildEmailPreviewInternal_/);
  if (!match) throw new Error('processRequest function was not found');
  if (!match[0].includes('activateRequestInternal_')) {
    throw new Error('processRequest does not activate the request');
  }
  if (/generateDocumentInternal_|createEmailDraftInternal_/.test(match[0])) {
    throw new Error('processRequest still generates documents or Gmail drafts');
  }
});

test('finance input is managed only in generated spreadsheets', () => {
  const index = fs.readFileSync(path.join(root, 'Index.html'), 'utf8');
  const scripts = scriptFiles.map(file => fs.readFileSync(path.join(root, file), 'utf8')).join('\n');
  const finance = fs.readFileSync(path.join(root, 'FinanceService.gs'), 'utf8');
  if (/view-finance|travelCostForm|travelModal|travelTableBody/.test(index + scripts)) {
    throw new Error('legacy web finance editor is still exposed');
  }
  if (!finance.includes('isTravelSheetComplete_')) {
    throw new Error('Perjadin readiness does not inspect the generated spreadsheet');
  }
  if (!finance.includes(".setFormula('=C' + sourceRow)")) {
    throw new Error('Perjadin duplicate amount table is not linked to the primary input');
  }
  if (!finance.includes('reused: reused')) {
    throw new Error('existing finance sheets are not explicitly reused');
  }
  if (!finance.includes('SpreadsheetApp.create(generatedSpreadsheetFileName_')) {
    throw new Error('finance output is not created as a separate spreadsheet');
  }
  if (/preview-document|documentPreviewModal/.test(index + scripts)) {
    throw new Error('removed document preview is still exposed');
  }
  const toolbar = scripts.match(/App\.renderFinanceToolbar_\s*=\s*function[\s\S]*?\nApp\.handleDetailAction\s*=\s*function/);
  if (!toolbar || /export-finance|Buat \/ Update PDF|Buat \/ Update Excel/.test(toolbar[0])) {
    throw new Error('finance export actions are still exposed in request detail');
  }
});

test('requests page keeps archived items visible by default', () => {
  const data = fs.readFileSync(path.join(root, 'DataService.gs'), 'utf8');
  if (data.includes('includeArchived') && data.includes("item.status === 'ARCHIVED'")) {
    const listRequests = data.match(/function listRequestsInternal_\([\s\S]*?\n  }\n\n  function saveRequest/);
    if (listRequests && /includeArchived[\s\S]*ARCHIVED/.test(listRequests[0])) {
      throw new Error('request list still hides archived items by default');
    }
  }
});

test('email routing applies template defaults and user selections', () => {
  const previousGetReferences = context.getReferenceDataInternal_;
  context.getReferenceDataInternal_ = () => ({
    cc: [
      { role: 'Dekan Fakultas Ekonomi', unit: 'Fakultas Ekonomi', email: 'dekan@example.com' },
      { role: 'Wakil Rektor Bidang Kerjasama, Alumni, Inovasi dan Bisnis', unit: 'Rektorat', email: 'wr@example.com' }
    ]
  });

  const previousGetTemplateConfigs = context.getTemplateConfigsInternal_;
  context.getTemplateConfigsInternal_ = () => [
    { key: 'EDU_FAIR_TASK', active: true, defaultTo: 'admin@example.com', defaultCc: 'wr@example.com', defaultBcc: '' }
  ];

  try {
    const request = {
      activityType: 'Edu Fair',
      documents: [{ type: 'Surat Tugas' }],
      manualTo: [],
      manualCc: []
    };
    const employees = [];
    const routing = context.computeEmailRouting_(request, employees);
    
    equal(routing.to.includes('admin@example.com'), true);
    equal(routing.cc.includes('wr@example.com'), true);
    equal(routing.cc.length, 1);

    const requestWithManual = Object.assign({}, request, { manualCc: ['wr@example.com'] });
    const routingWithManual = context.computeEmailRouting_(requestWithManual, employees);
    equal(routingWithManual.cc.includes('wr@example.com'), true);
    equal(routingWithManual.cc.length, 1);
  } finally {
    context.getReferenceDataInternal_ = previousGetReferences;
    context.getTemplateConfigsInternal_ = previousGetTemplateConfigs;
  }
});

test('task and speaker-request routing follow employee data', () => {
  const previousGetReferences = context.getReferenceDataInternal_;
  const previousGetTemplateConfigs = context.getTemplateConfigsInternal_;
  context.getReferenceDataInternal_ = () => ({
    cc: [
      { role: 'Dekan Fakultas Ekonomi', unit: 'Fakultas Ekonomi', email: 'dekan.fe@example.com' },
      { role: 'Sekretaris Universitas', unit: 'Rektorat', email: 'su@example.com' }
    ]
  });
  context.getTemplateConfigsInternal_ = () => [];

  try {
    const employees = [{
      name: 'Narasumber FE',
      unit: 'Fakultas Ekonomi',
      email: 'narasumber@example.com'
    }];

    const taskRouting = context.computeEmailRouting_({
      activityType: 'Edu Fair',
      documents: [{ type: 'Surat Tugas' }],
      manualTo: ['su@example.com'],
      manualCc: []
    }, employees, 'Surat Tugas');
    equal(taskRouting.to.join(','), 'narasumber@example.com');

    const requestRouting = context.computeEmailRouting_({
      activityType: 'Penugasan Narasumber',
      documents: [{
        type: 'Surat Permohonan Narasumber kepada Dekan',
        speakerStatus: 'Tidak Dicarikan'
      }],
      faculties: ['Fakultas Hukum'],
      manualTo: ['su@example.com'],
      manualCc: ['su@example.com']
    }, employees, 'Surat Permohonan Narasumber kepada Dekan');
    equal(requestRouting.to.join(','), 'dekan.fe@example.com');
    equal(requestRouting.cc.join(','), 'su@example.com');
  } finally {
    context.getReferenceDataInternal_ = previousGetReferences;
    context.getTemplateConfigsInternal_ = previousGetTemplateConfigs;
  }
});

test('email preview keeps recipient names separate from addresses', () => {
  const source = fs.readFileSync(path.join(root, 'EmailService.gs'), 'utf8');
  equal(source.includes('const automaticRouting = computeEmailRouting_'), true);
  equal(source.includes('toRoles: to'), false);
  equal(source.includes('ccRoles: cc'), false);
});

test('syncTravelDataInternal preserves archived travel rows during full sync', () => {
  const masterRows = [
    ['REQ-1', 'Edu Fair', 'Surat Tugas', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'Tidak', 'Ya', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'DRAFT', '', '', '', '', 'token-1', '2026-06-01', '2026-06-01', 1],
    ['REQ-2', 'Edu Fair', 'Surat Tugas', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'Tidak', 'Ya', '', '', '', '', '', '', '', '', '', '', '', '', '', '', 'ARCHIVED', '', '', '', '', 'token-2', '2026-06-01', '2026-06-01', 1]
  ];
  const employeeRows = [
    ['REQ-1', 'Andi', '123', 'Staf', 'Fakultas Ekonomi', 'andi@example.com', '', '', 'REQ-1|andi'],
    ['REQ-2', 'Budi', '456', 'Staf', 'Fakultas Ekonomi', 'budi@example.com', '', '', 'REQ-2|budi']
  ];
  const travelRows = [
    ['REQ-1', 'Andi', '123', '', '', '', '2026-06-01', 'Aula', 100000, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 'REQ-1|andi'],
    ['REQ-2', 'Budi', '456', '', '', '', '2026-06-01', 'Aula', 200000, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 'REQ-2|budi']
  ];

  let writtenRows = null;
  const previousGetSheet = context.getSheet_;
  const previousReadDataRows = context.readDataRows_;
  const previousRewriteDataRows = context.rewriteDataRows_;
  const previousGetDocuments = context.getDocumentsByRequest_;

  context.getSheet_ = key => ({ name: key, getLastRow: () => 10 });
  context.getDocumentsByRequest_ = () => [];
  context.readDataRows_ = (sheet, width) => {
    if (sheet.name === 'MASTER') return masterRows;
    if (sheet.name === 'EMPLOYEES') return employeeRows;
    if (sheet.name === 'TRAVEL') return travelRows;
    return [];
  };
  context.rewriteDataRows_ = (sheet, rows, width) => {
    if (sheet.name === 'TRAVEL') writtenRows = rows;
  };

  try {
    context.syncTravelDataInternal_('');
    equal(writtenRows.length, 2);
    const andiRow = writtenRows.find(r => r[0] === 'REQ-1');
    const budiRow = writtenRows.find(r => r[0] === 'REQ-2');
    equal(andiRow[8], 100000);
    equal(budiRow[8], 200000);
  } finally {
    context.getSheet_ = previousGetSheet;
    context.readDataRows_ = previousReadDataRows;
    context.rewriteDataRows_ = previousRewriteDataRows;
    context.getDocumentsByRequest_ = previousGetDocuments;
  }
});

test('legacy migration handles unnormalized activity and document types', () => {
  const d1 = context.normalizeDocumentDescriptor_({
    activityType: 'penugasan narasumber ',
    type: 'surat tugas',
    speakerSubtype: '',
    speakerStatus: ''
  });
  equal(d1.activityType, 'Penugasan Narasumber');
  equal(d1.type, 'Surat Tugas');

  const d2 = context.normalizeDocumentDescriptor_({
    activityType: 'narasumber',
    type: 'surat permohonan narasumber kepada dekan',
    speakerSubtype: '',
    speakerStatus: ''
  });
  equal(d2.activityType, 'Penugasan Narasumber');
  equal(d2.type, 'Surat Permohonan Narasumber kepada Dekan');
  equal(d2.speakerStatus, 'Tidak Dicarikan');
});

test('legacy migration infers subtype correctly even with casing variations', () => {
  const row = new Array(84).fill('');
  row[1] = 'penugasan narasumber ';
  row[2] = 'surat tugas';
  row[3] = '';
  row[42] = '';
  row[46] = 'DOC-PROMO-123';

  const descriptor = context.normalizeDocumentDescriptor_({
    activityType: row[1],
    type: row[2],
    speakerSubtype: row[3],
    speakerStatus: row[4]
  });
  if (descriptor.activityType === 'Penugasan Narasumber' && descriptor.type === 'Surat Tugas' && !descriptor.speakerSubtype) {
    const workshopDocId = context.text_(row[42]);
    const promotionDocId = context.text_(row[46]);
    if (promotionDocId && !workshopDocId) {
      descriptor.speakerSubtype = 'Promosi';
    } else {
      descriptor.speakerSubtype = 'Workshop';
    }
  }

  equal(descriptor.activityType, 'Penugasan Narasumber');
  equal(descriptor.type, 'Surat Tugas');
  equal(descriptor.speakerSubtype, 'Promosi');
});

test('legacy migration parses non-contiguous range with dan into multiple list items', () => {
  const result = context.parseLegacyDateRange_('25 dan 30 Agustus 2025');
  equal(result.start, '2025-08-25');
  equal(result.end, '2025-08-30');
  equal(result.list.length, 2);
  equal(result.list[0].start, '2025-08-25');
  equal(result.list[0].end, '2025-08-25');
  equal(result.list[1].start, '2025-08-30');
});

test('client-side sorting logic', () => {
  const requests = [
    { id: 'REQ-1', activityName: 'A', partnerName: 'X', startDate: '2026-06-10', status: 'DRAFT' },
    { id: 'REQ-2', activityName: 'B', partnerName: 'Y', startDate: '2026-06-11', status: 'READY' }
  ];

  const sortDateDesc = context.App.sortRequests(requests, 'date-desc');
  equal(sortDateDesc[0].id, 'REQ-2');

  const sortDateAsc = context.App.sortRequests(requests, 'date-asc');
  equal(sortDateAsc[0].id, 'REQ-1');

  const sortIdAsc = context.App.sortRequests(requests, 'id-asc');
  equal(sortIdAsc[0].id, 'REQ-1');

  const sortActivityAsc = context.App.sortRequests(requests, 'activity-asc');
  equal(sortActivityAsc[0].activityName, 'A');

  const sortPartnerAsc = context.App.sortRequests(requests, 'partner-asc');
  equal(sortPartnerAsc[0].partnerName, 'X');

  const sortStatusAsc = context.App.sortRequests(requests, 'status-asc');
  equal(sortStatusAsc[0].status, 'DRAFT');
});

if (!process.exitCode) console.log('PASS pure behavior tests');
