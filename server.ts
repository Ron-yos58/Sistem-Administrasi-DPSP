import express from 'express';
import fs from 'fs';
import path from 'path';
import vm from 'vm';

const app = express();
const PORT = 3000;

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const DB_FILE = path.join(__dirname, 'db.json');

// Memory Database structure
interface DBStructure {
  spreadsheets: {
    [id: string]: {
      id: string;
      name: string;
      sheets: {
        [name: string]: {
          data: any[][];
          metadata: { key: string; value: string }[];
        };
      };
    };
  };
  userProperties: { [key: string]: string };
  scriptProperties: { [key: string]: string };
  files: {
    [id: string]: {
      id: string;
      name: string;
      content: string;
      mimeType: string;
      trashed: boolean;
    };
  };
  gmailDrafts: any[];
}

// Load or initialize Database
let db: DBStructure = {
  spreadsheets: {},
  userProperties: {},
  scriptProperties: {},
  files: {},
  gmailDrafts: []
};

if (fs.existsSync(DB_FILE)) {
  try {
    db = JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
  } catch (e) {
    console.error('Failed to parse database, resetting...', e);
  }
}

// Ensure the default spreadsheet is present in the database structure
const DEFAULT_SPREADSHEET_ID = '1jDY5XW86HmDdzYTcPVxpQBULHcP5qW-cxqCuEJZeyVI';
const DEFAULT_OUTPUT_FOLDER_ID = '1MsDCWr-xJRd2K91KZLyJago5uiV4L4Dx';

if (!db.spreadsheets[DEFAULT_SPREADSHEET_ID]) {
  db.spreadsheets[DEFAULT_SPREADSHEET_ID] = {
    id: DEFAULT_SPREADSHEET_ID,
    name: 'Sistem Surat DPSP',
    sheets: {}
  };
}

// Save database helper
function saveDB() {
  fs.writeFileSync(DB_FILE, JSON.stringify(db, null, 2), 'utf8');
}

// Seeding templates on Drive startup
const TEMPLATES_DRIVE_DATA: { [id: string]: { name: string; content: string } } = {
  '1GxHt4CYcsmKHjuMlLwzfpolhblG6GUtOOazaX70zuws': {
    name: 'Template Surat Tugas - Edu Fair',
    content: 'SURAT TUGAS\nNomor: {nomorSurat}\n\nSehubungan dengan kegiatan {namaKegiatan}...'
  },
  '1PcVlBo6Q81x9FjU6hpOECgI5Q5Pio0wD3-mwHsA4wRc': {
    name: 'Template Surat Tugas - Narasumber (Workshop)',
    content: 'SURAT TUGAS NARASUMBER\nNomor: {nomorSurat}\n\nDengan ini menugaskan:\n{narasumber}\n\nUntuk kegiatan {namaKegiatan}...'
  },
  '1uaEvHJxPqdhgcNTqpjG-LcssPo-7bAboaa_rsrglrQI': {
    name: 'Template Surat Tugas - Narasumber (Promosi)',
    content: 'SURAT TUGAS PROMOSI\nNomor: {nomorSurat}\n\nMenugaskan:\n{narasumber}...'
  },
  '1FB-0BchuGmPDEda7NyAdtGhxWv07VrMWz2YHbgR_L5o': {
    name: 'Template Surat Tugas - Campus Visit',
    content: 'SURAT TUGAS CAMPUS VISIT\nNomor: {nomorSurat}\n\nMenugaskan:\n{narasumber}...'
  },
  '1yjEBTJ4d-KODomgJqmFCeu6jeDhlEbF0E26Qx3mGoNQ': {
    name: 'Template Surat Izin Pimpinan - Campus Visit',
    content: 'SURAT IZIN PIMPINAN\nNomor: {nomorSurat}\n\nKepada: {kepadaYth}\n\nSehubungan dengan Campus Visit {namaMitra}...'
  },
  '1K_yPfDoq0ePXlJtYiLzpSBCxdOGDgubTsrtj4CeQGRA': {
    name: 'Template Surat Rekomendasi Campus Visit - SU',
    content: 'SURAT REKOMENDASI CAMPUS VISIT\nNomor: {nomorSurat}\n\nMenyarankan kegiatan Campus Visit {namaMitra}...'
  },
  '1kK5MORz9alFtAct2gAkygivqkpDTpN1MLXaGhkqVOMM': {
    name: 'Template Surat Balasan Campus Visit',
    content: 'SURAT BALASAN CAMPUS VISIT\nNomor: {nomorSurat}\n\nKepada {kepadaYth}\n\nMenanggapi surat Nomor {nomorSuratMasuk}...'
  },
  '1ZhLP76Ysse6IKMlO_mN-61IwMZc0iWcQqk6OMyFdSK8': {
    name: 'Template Surat Permohonan Narasumber (Dicarikan)',
    content: 'PERMOHONAN NARASUMBER\nNomor: {nomorSurat}\n\nKepada Dekan:\nMohon ditugaskan narasumber untuk {namaKegiatan}...'
  },
  '1K5aik94SZW20X0TPLqIXZEmJ0p-r5ZoxJy9FO4Ywd24': {
    name: 'Template Surat Permohonan Narasumber (Sudah Ada)',
    content: 'PERMOHONAN PENUGASAN NARASUMBER\nNomor: {nomorSurat}\n\nKepada Dekan:\nMohon ditugaskan:\n{narasumber}\n\nUntuk kegiatan {namaKegiatan}...'
  }
};

for (const [id, t] of Object.entries(TEMPLATES_DRIVE_DATA)) {
  if (!db.files[id]) {
    db.files[id] = {
      id,
      name: t.name,
      content: t.content,
      mimeType: 'application/vnd.google-apps.document',
      trashed: false
    };
  }
}
saveDB();

// Mock Classes
class MockRange {
  sheet: MockSheet;
  row: number;
  col: number;
  numRows: number;
  numCols: number;

  constructor(sheet: MockSheet, row: number, col: number, numRows: number, numCols: number) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }

  getValues() {
    const values: any[][] = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowIdx = this.row - 1 + r;
      const rowData: any[] = [];
      const sheetRow = this.sheet.data[rowIdx] || [];
      for (let c = 0; c < this.numCols; c++) {
        const colIdx = this.col - 1 + c;
        const cellVal = sheetRow[colIdx];
        rowData.push(cellVal === undefined || cellVal === null ? "" : cellVal);
      }
      values.push(rowData);
    }
    return values;
  }

  getDisplayValues() {
    const values = this.getValues();
    return values.map(row => row.map(v => {
      if (v instanceof Date) {
        const year = v.getFullYear();
        const month = String(v.getMonth() + 1).padStart(2, '0');
        const day = String(v.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
      }
      return v === null || v === undefined ? "" : String(v);
    }));
  }

  getValue() {
    return this.getValues()[0][0];
  }

  getDisplayValue() {
    return this.getDisplayValues()[0][0];
  }

  setValue(value: any) {
    this.setValues([[value]]);
  }

  setValues(values: any[][]) {
    for (let r = 0; r < values.length; r++) {
      const sheetRowIdx = this.row - 1 + r;
      while (this.sheet.data.length <= sheetRowIdx) {
        this.sheet.data.push([]);
      }
      const rowData = this.sheet.data[sheetRowIdx];
      for (let c = 0; c < values[r].length; c++) {
        const sheetColIdx = this.col - 1 + c;
        let cellVal = values[r][c];
        rowData[sheetColIdx] = cellVal;
      }
    }
    this.sheet.ss.save();
  }

  clearContent() {
    for (let r = 0; r < this.numRows; r++) {
      const rowIdx = this.row - 1 + r;
      const sheetRow = this.sheet.data[rowIdx];
      if (sheetRow) {
        for (let c = 0; c < this.numCols; c++) {
          const colIdx = this.col - 1 + c;
          sheetRow[colIdx] = "";
        }
      }
    }
    this.sheet.ss.save();
  }

  clear() {
    this.clearContent();
  }

  clearDataValidations() {
    return this;
  }

  setBackground(color: string) {
    return this;
  }

  setFontColor(color: string) {
    return this;
  }

  setFontWeight(weight: string) {
    return this;
  }

  setWrap(wrap: boolean) {
    return this;
  }

  setFormula(formula: string) {
    return this;
  }

  getFormulas() {
    const values = this.getValues();
    return values.map(row => row.map(() => ""));
  }

  getRow() {
    return this.row;
  }

  getNextDataCell(direction: string) {
    if (direction === 'UP') {
      let r = this.row - 1;
      while (r >= 0) {
        const val = this.sheet.data[r] ? this.sheet.data[r][this.col - 1] : undefined;
        if (val !== undefined && val !== null && String(val).trim() !== '') {
          return new MockRange(this.sheet, r + 1, this.col, 1, 1);
        }
        r--;
      }
      return new MockRange(this.sheet, 1, this.col, 1, 1);
    }
    return this;
  }
}

class MockSheet {
  ss: MockSpreadsheet;
  name: string;
  data: any[][];
  metadata: { key: string; value: string }[];
  id: number;

  constructor(ss: MockSpreadsheet, name: string, data: any[][] = [], metadata: { key: string; value: string }[] = []) {
    this.ss = ss;
    this.name = name;
    this.data = data;
    this.metadata = metadata;
    this.id = Math.floor(Math.random() * 1000000);
  }

  getName() {
    return this.name;
  }

  getSheetId() {
    return this.id;
  }

  getLastRow() {
    for (let r = this.data.length - 1; r >= 0; r--) {
      const row = this.data[r];
      if (row && row.some(cell => cell !== undefined && cell !== null && String(cell).trim() !== '')) {
        return r + 1;
      }
    }
    return 0;
  }

  getLastColumn() {
    let maxCol = 0;
    for (const row of this.data) {
      if (row) {
        for (let c = row.length - 1; c >= 0; c--) {
          if (row[c] !== undefined && row[c] !== null && String(row[c]).trim() !== '') {
            maxCol = Math.max(maxCol, c + 1);
          }
        }
      }
    }
    return maxCol;
  }

  getMaxColumns() {
    return Math.max(100, this.getLastColumn());
  }

  getMaxRows() {
    return 1000;
  }

  insertColumnsAfter(afterColumn: number, howMany: number) {
    return this;
  }

  insertRowsAfter(afterRow: number, howMany: number) {
    return this;
  }

  getRange(row: any, col?: number, numRows?: number, numCols?: number) {
    if (typeof row === 'string') {
      const match = row.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
      if (match) {
        const col1 = match[1].split('').reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0);
        const row1 = parseInt(match[2], 10);
        let col2 = col1;
        let row2 = row1;
        if (match[3]) {
          col2 = match[3].split('').reduce((acc, char) => acc * 26 + char.charCodeAt(0) - 64, 0);
          row2 = parseInt(match[4], 10);
        }
        return new MockRange(this, row1, col1, row2 - row1 + 1, col2 - col1 + 1);
      }
      return new MockRange(this, 1, 1, 1, 1);
    }
    return new MockRange(this, row, col!, numRows === undefined ? 1 : numRows, numCols === undefined ? 1 : numCols);
  }

  getDataRange() {
    return this.getRange(1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn()));
  }

  clear() {
    this.data = [];
    this.ss.save();
  }

  clearContent() {
    this.clear();
  }

  deleteColumns(start: number, count: number) {
    for (const row of this.data) {
      if (row) {
        row.splice(start - 1, count);
      }
    }
    this.ss.save();
  }

  deleteSheet() {
    this.ss.deleteSheet(this);
  }

  getDeveloperMetadata() {
    return this.metadata.map(item => ({
      getKey: () => item.key,
      getValue: () => item.value
    }));
  }

  addDeveloperMetadata(key: string, value: string) {
    const meta = { key, value };
    this.metadata.push(meta);
    this.ss.save();
    return {
      getKey: () => key,
      getValue: () => value
    };
  }

  copyTo(targetSpreadsheet: MockSpreadsheet) {
    const newSheet = targetSpreadsheet.insertSheet(this.name);
    newSheet.data = JSON.parse(JSON.stringify(this.data));
    newSheet.metadata = JSON.parse(JSON.stringify(this.metadata));
    targetSpreadsheet.save();
    return newSheet;
  }

  setFrozenRows(rows: number) {
    return this;
  }
}

class MockSpreadsheet {
  db: any;
  id: string;
  name: string;
  sheets: { [name: string]: MockSheet };

  constructor(db: any, id: string, name: string, sheetsData: any = {}) {
    this.db = db;
    this.id = id;
    this.name = name;
    this.sheets = {};
    for (const [sheetName, sheetConfig] of Object.entries(sheetsData)) {
      const config = sheetConfig as any;
      this.sheets[sheetName] = new MockSheet(this, sheetName, config.data || [], config.metadata || []);
    }
  }

  getId() {
    return this.id;
  }

  getName() {
    return this.name;
  }

  getSheetByName(name: string) {
    return this.sheets[name] || null;
  }

  getSheets() {
    return Object.values(this.sheets);
  }

  insertSheet(name: string) {
    if (!this.sheets[name]) {
      this.sheets[name] = new MockSheet(this, name, [], []);
      this.save();
    }
    return this.sheets[name];
  }

  deleteSheet(sheet: MockSheet) {
    delete this.sheets[sheet.name];
    this.save();
  }

  getUrl() {
    return `https://docs.google.com/spreadsheets/d/${this.id}/edit`;
  }

  setSpreadsheetTimeZone(tz: string) {
    return this;
  }

  save() {
    this.db.saveSpreadsheet(this);
  }
}

// Global VM Context Creator
function buildVMContext() {
  const context: any = {
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
    isNaN,
    parseInt,
    parseFloat,
    encodeURIComponent,
    decodeURIComponent,
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: (key: string) => db.scriptProperties[key] || ""
      }),
      getUserProperties: () => ({
        getProperty: (key: string) => db.userProperties[key] || "",
        setProperty: (key: string, val: string) => {
          db.userProperties[key] = val;
          saveDB();
        }
      })
    },
    CacheService: {
      getScriptCache: () => ({
        get: (key: string) => {
          // Dummy cache in-memory
          return null;
        },
        put: (key: string, val: string, duration: number) => {},
        remove: (key: string) => {},
        removeAll: (keys: string[]) => {}
      })
    },
    LockService: {
      getScriptLock: () => ({
        waitLock: () => {},
        releaseLock: () => {}
      })
    },
    Session: {
      getActiveUser: () => ({
        getEmail: () => 'ronaldsebastian@unpar.ac.id'
      })
    },
    Utilities: {
      getUuid: () => {
        return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
          var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
          return v.toString(16);
        });
      },
      formatDate: (date: Date, timeZone: string, pattern: string) => {
        const d = new Date(date);
        if (pattern === 'HH:mm') {
          return `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
        }
        if (pattern === 'yyyyMMdd') {
          const year = d.getFullYear();
          const month = String(d.getMonth() + 1).padStart(2, '0');
          const day = String(d.getDate()).padStart(2, '0');
          return `${year}${month}${day}`;
        }
        const year = d.getFullYear();
        const month = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
      },
      sleep: (ms: number) => {},
      base64Encode: (str: string) => Buffer.from(str).toString('base64'),
      newBlob: (content: string, mime: string, name: string) => {
        return {
          getName: () => name,
          getContent: () => content,
          getBlob: () => this,
          getAs: (type: string) => ({
            setName: (n: string) => {
              name = n;
            },
            getBlob: () => this,
            getAs: (t: string) => this
          })
        };
      },
      computeDigest: () => []
    },
    SpreadsheetApp: {
      Direction: { UP: 'UP', DOWN: 'DOWN' },
      openById: (id: string) => {
        if (!db.spreadsheets[id]) {
          db.spreadsheets[id] = { id, name: 'Sistem Surat DPSP', sheets: {} };
        }
        return new MockSpreadsheet({
          saveSpreadsheet: (ss: MockSpreadsheet) => {
            const sheetsData: any = {};
            for (const [sName, sObj] of Object.entries(ss.sheets)) {
              sheetsData[sName] = { data: sObj.data, metadata: sObj.metadata };
            }
            db.spreadsheets[ss.id] = { id: ss.id, name: ss.name, sheets: sheetsData };
            saveDB();
          }
        }, id, db.spreadsheets[id].name, db.spreadsheets[id].sheets);
      },
      getActiveSpreadsheet: () => {
        return context.SpreadsheetApp.openById(DEFAULT_SPREADSHEET_ID);
      },
      create: (name: string) => {
        const id = 'spread_' + Math.random().toString(36).substr(2, 9);
        db.spreadsheets[id] = { id, name, sheets: {} };
        saveDB();
        return context.SpreadsheetApp.openById(id);
      },
      flush: () => {
        saveDB();
      }
    },
    DriveApp: {
      getFileById: (id: string) => {
        const file = db.files[id];
        if (!file) throw new Error("File tidak ditemukan di Google Drive mock: " + id);
        return {
          getId: () => file.id,
          getName: () => file.name,
          getUrl: () => `/api/files/${file.id}`,
          getAs: (mimeType: string) => ({
            setName: (name: string) => {
              file.name = name;
              saveDB();
              return this;
            }
          }),
          getBlob: () => ({
            getAs: (mimeType: string) => ({
              setName: (name: string) => {
                file.name = name;
                saveDB();
                return this;
              }
            })
          }),
          makeCopy: (newName: string, folder: any) => {
            const copyId = 'file_' + Math.random().toString(36).substr(2, 9);
            db.files[copyId] = {
              id: copyId,
              name: newName,
              content: file.content || "",
              mimeType: file.mimeType,
              trashed: false
            };
            saveDB();
            return context.DriveApp.getFileById(copyId);
          },
          setTrashed: (val: boolean) => {
            file.trashed = val;
            saveDB();
          },
          moveTo: (folder: any) => {}
        };
      },
      getFolderById: (id: string) => {
        return {
          getId: () => id,
          getName: () => 'Mock Output Folder',
          createFile: (blob: any) => {
            const fileId = 'file_' + Math.random().toString(36).substr(2, 9);
            let name = "unnamed.pdf";
            if (blob && typeof blob.getName === 'function') {
              name = blob.getName();
            }
            db.files[fileId] = {
              id: fileId,
              name: name,
              content: `Merged Document Content for: ${name}`,
              mimeType: 'application/pdf',
              trashed: false
            };
            saveDB();
            return context.DriveApp.getFileById(fileId);
          }
        };
      },
      getFile: (id: string) => {
        return context.DriveApp.getFileById(id);
      }
    },
    DocumentApp: {
      openById: (id: string) => {
        const file = db.files[id];
        if (!file) throw new Error("Document tidak ditemukan: " + id);
        return {
          getId: () => file.id,
          getName: () => file.name,
          getBody: () => ({
            replaceText: (target: string, replacement: string) => {
              file.content = (file.content || "").replace(new RegExp(target, 'g'), replacement);
              saveDB();
            }
          }),
          saveAndClose: () => {
            saveDB();
          }
        };
      }
    },
    GmailApp: {
      getDraft: (id: string) => {
        const draft = db.gmailDrafts.find(d => d.id === id);
        if (!draft) return null;
        return {
          getId: () => draft.id,
          getMessage: () => ({
            getId: () => 'msg_' + draft.id,
            getBody: () => draft.body,
            getSubject: () => draft.subject,
            getTo: () => draft.to
          })
        };
      },
      getDrafts: () => {
        return db.gmailDrafts.map(draft => ({
          getId: () => draft.id,
          getMessage: () => ({
            getId: () => 'msg_' + draft.id,
            getBody: () => draft.body,
            getSubject: () => draft.subject,
            getTo: () => draft.to
          })
        }));
      },
      createDraft: (to: string, subject: string, body: string, options: any) => {
        const id = 'draft_' + Math.random().toString(36).substr(2, 9);
        const draft = { id, to, subject, body, options };
        db.gmailDrafts.push(draft);
        saveDB();
        return {
          getId: () => id,
          getMessage: () => ({
            getId: () => 'msg_' + id,
            getBody: () => body,
            getSubject: () => subject,
            getTo: () => to
          })
        };
      }
    },
    MailApp: {
      sendEmail: (to: string, subject: string, body: string, options: any) => {
        console.log(`[MailApp] Sending email to: ${to} | Subject: ${subject}`);
      }
    },
    UrlFetchApp: {
      fetch: (url: string) => ({
        getResponseCode: () => 200,
        getBlob: () => ({
          setName: (name: string) => ({
            getName: () => name
          })
        })
      })
    },
    ScriptApp: {
      getOAuthToken: () => 'mock-oauth-token'
    },
    HtmlService: {
      createTemplateFromFile: (name: string) => ({
        evaluate: () => ({
          setTitle: (title: string) => ({
            addMetaTag: (meta1: string, meta2: string) => ({
              getContent: () => getEvaluatedHtml()
            })
          })
        })
      }),
      createHtmlOutputFromFile: (name: string) => ({
        getContent: () => {
          const p = path.join(__dirname, `${name}.html`);
          return fs.existsSync(p) ? fs.readFileSync(p, 'utf8') : '';
        }
      })
    },
    MimeType: {
      PDF: 'application/pdf'
    }
  };

  vm.createContext(context);
  
  // Read and run all .gs backend files
  const gsFiles = [
    'Config.gs',
    'Utils.gs',
    'DataService.gs',
    'Migration.gs',
    'DocumentService.gs',
    'EmailService.gs',
    'ExportService.gs',
    'FinanceService.gs',
    'Setup.gs',
    'Code.gs'
  ];

  for (const file of gsFiles) {
    const code = fs.readFileSync(path.join(__dirname, file), 'utf8');
    vm.runInContext(code, context, { filename: file });
  }

  return context;
}

const vmContext = buildVMContext();

// Perform initial setup if MASTER sheet doesn't exist
const mainSpreadsheet = db.spreadsheets[DEFAULT_SPREADSHEET_ID];
if (!mainSpreadsheet || Object.keys(mainSpreadsheet.sheets).length === 0) {
  console.log('Main spreadsheet is empty. Initializing and running setupSystem()...');
  try {
    vmContext.setupSystem();
    console.log('setupSystem completed successfully!');
  } catch (err: any) {
    console.error('Error in setupSystem:', err.message);
  }
}

// Evaluates the full main Index.html file with its nested includes recursively
function getEvaluatedHtml(): string {
  let html = fs.readFileSync(path.join(__dirname, 'Index.html'), 'utf8');

  // Replace <?!= include('Filename'); ?> with corresponding .html or .gs file content
  const includeRegex = /<\?!=\s*include\(['"]([^'"]+)['"]\);\s*\?>/g;
  let match;
  while ((match = includeRegex.exec(html)) !== null) {
    const filename = match[1];
    let fileContent = '';
    const filepathHtml = path.join(__dirname, `${filename}.html`);
    const filepathGs = path.join(__dirname, `${filename}.gs`);
    if (fs.existsSync(filepathHtml)) {
      fileContent = fs.readFileSync(filepathHtml, 'utf8');
    } else if (fs.existsSync(filepathGs)) {
      fileContent = fs.readFileSync(filepathGs, 'utf8');
    }
    html = html.replace(match[0], fileContent);
    // Reset index to ensure we recursively replace tags
    includeRegex.lastIndex = 0;
  }

  // Inject a lightweight Google Apps Script client-side RPC proxy
  const shim = `
    <script>
      (function() {
        function makeRunner(successHandler, failureHandler) {
          return new Proxy({}, {
            get: function(target, prop) {
              if (prop === 'withSuccessHandler') {
                return function(sh) { return makeRunner(sh, failureHandler); };
              }
              if (prop === 'withFailureHandler') {
                return function(fh) { return makeRunner(successHandler, fh); };
              }
              return function(...args) {
                // Show blocking loader for mutating methods
                var mutatingMethods = [
                  'saveRequest', 'addDocument', 'saveDocumentDetails', 'generateDocument',
                  'generateGoogleDoc', 'generatePdf', 'createEmailDraft', 'deleteDocument',
                  'archiveRequest', 'processRequest', 'generateFinanceSheet', 'generateAndExportFinance',
                  'addReferenceCC', 'setupSystem', 'migrateLegacyData', 'cleanupMigratedAutocratColumns',
                  'repairMigratedMasterIds', 'repairMigratedEmployeeIds'
                ];
                var isMutating = mutatingMethods.indexOf(prop) !== -1;
                if (isMutating && window.App && typeof App.showOverlay === 'function') {
                  App.showOverlay('Sedang memproses...');
                }

                fetch('/api/rpc', {
                  method: 'POST',
                  headers: { 'Content-Type': 'application/json' },
                  body: JSON.stringify({ method: prop, args: args })
                })
                .then(res => {
                  if (!res.ok) {
                    return res.json().then(err => { throw new Error(err.error || 'Server error'); });
                  }
                  return res.json();
                })
                .then(data => {
                  if (isMutating && window.App && typeof App.hideOverlay === 'function') {
                    App.hideOverlay();
                  }
                  if (data.error) {
                    if (failureHandler) failureHandler(new Error(data.error));
                    else console.error('RPC Error:', data.error);
                  } else {
                    if (successHandler) successHandler(data.result);
                  }
                })
                .catch(err => {
                  if (isMutating && window.App && typeof App.hideOverlay === 'function') {
                    App.hideOverlay();
                  }
                  if (failureHandler) failureHandler(err);
                  else console.error('Fetch/Network error:', err);
                });
              };
            }
          });
        }
        window.google = {
          script: {
            run: makeRunner()
          }
        };
      })();
    </script>
  `;

  html = html.replace('</head>', `${shim}</head>`);
  return html;
}

// Server Endpoints
app.get('/', (req, res) => {
  try {
    const html = getEvaluatedHtml();
    res.send(html);
  } catch (err: any) {
    res.status(500).send(`Error rendering Index: ${err.message}`);
  }
});

// JSON API router to execute Google Apps Script server methods
app.post('/api/rpc', (req, res) => {
  const { method, args } = req.body;
  if (!method) {
    return res.status(400).json({ error: 'Method name is required' });
  }

  const fn = vmContext[method];
  if (typeof fn !== 'function') {
    return res.status(404).json({ error: `Function "${method}" not found or available on backend.` });
  }

  try {
    console.log(`[RPC] Invoking: ${method} with args:`, JSON.stringify(args));
    const result = fn(...(args || []));
    res.json({ result });
  } catch (err: any) {
    console.error(`[RPC] Error in: ${method}`, err);
    res.status(500).json({ error: err.message || String(err) });
  }
});

// Download/View Generated Virtual Drive Files
app.get('/api/files/:id', (req, res) => {
  const fileId = req.params.id;
  const file = db.files[fileId];
  if (!file) {
    return res.status(404).send('File virtual tidak ditemukan di Google Drive mock.');
  }

  res.send(`
    <!DOCTYPE html>
    <html lang="id">
    <head>
      <meta charset="UTF-8">
      <title>${file.name}</title>
      <style>
        body {
          font-family: 'Manrope', 'Inter', sans-serif;
          background-color: #f6f9f8;
          color: #002e2a;
          margin: 0;
          padding: 40px;
          display: flex;
          justify-content: center;
        }
        .document-container {
          background: white;
          width: 100%;
          max-width: 800px;
          min-height: 1000px;
          padding: 60px 80px;
          box-shadow: 0 4px 20px rgba(1, 88, 80, 0.08);
          border-radius: 8px;
          box-sizing: border-box;
          border: 1px solid #dcece9;
        }
        .header {
          text-align: center;
          margin-bottom: 40px;
          border-bottom: 2px solid #015850;
          padding-bottom: 20px;
        }
        .header h2 {
          margin: 0;
          color: #015850;
          font-size: 20px;
          font-weight: 700;
        }
        .content {
          line-height: 1.8;
          white-space: pre-wrap;
          font-size: 14px;
        }
        .meta {
          font-size: 11px;
          color: #8fa09d;
          margin-top: 60px;
          border-top: 1px solid #e5ebe9;
          padding-top: 10px;
          display: flex;
          justify-content: space-between;
        }
      </style>
    </head>
    <body>
      <div class="document-container">
        <div class="header">
          <h2>Virtual Google Drive Document Preview</h2>
          <small>Sistem Administrasi DPSP — Universitas Katolik Parahyangan</small>
        </div>
        <div class="content">${file.content}</div>
        <div class="meta">
          <span>File ID: ${file.id}</span>
          <span>Status: Virtual Generated</span>
        </div>
      </div>
    </body>
    </html>
  `);
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Server successfully running on http://0.0.0.0:${PORT}`);
});
