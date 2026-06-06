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

test('email validation and deduplication', () => {
  equal(
    Array.from(context.emailList_('A@EXAMPLE.COM; invalid; a@example.com, b@example.com')),
    ['a@example.com', 'b@example.com']
  );
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

if (!process.exitCode) console.log('PASS pure behavior tests');
