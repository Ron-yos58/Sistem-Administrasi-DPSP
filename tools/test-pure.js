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
    formatDate: date => {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
  }
};
vm.createContext(context);

for (const file of ['Config.gs', 'Utils.gs', 'DataService.gs', 'Migration.gs']) {
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), context, { filename: file });
}

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
    '25 Juni 2026: 09.00 - 12.00 WIB\n30 Juni 2026: 13.00 - 15.00 WIB'
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
  equal(summary.timeDisplay, '09.00 - 15.00 WIB');
});

test('legacy time range parsing', () => {
  equal(
    context.parseTimeRange_('Pukul 09.00 - 12.30 WIB'),
    { startTime: '09:00', endTime: '12:30' }
  );
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

if (!process.exitCode) console.log('PASS pure behavior tests');
