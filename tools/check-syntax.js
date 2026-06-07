const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const ignored = new Set([
  'Delete_Sheet.gs',
  'Draft_Email.gs',
  'Generate_Data_Pegawai_dan_Perjadin.gs',
  'Generate_Export_Sheet.gs',
  'Generate_Narsumber_Edu Fair_Campus Visit.gs',
  'Generate_Perjadin.gs'
]);
const gsFiles = fs.readdirSync(root)
  .filter(file => file.endsWith('.gs') && !ignored.has(file))
  .sort();

let failed = false;
const functionOwners = new Map();
const combinedSources = [];

for (const file of gsFiles) {
  const source = fs.readFileSync(path.join(root, file), 'utf8');
  combinedSources.push(`\n// ${file}\n${source}`);
  try {
    new vm.Script(source, { filename: file });
    console.log(`OK syntax ${file}`);
  } catch (error) {
    failed = true;
    console.error(`FAIL syntax ${file}: ${error.message}`);
  }

  const matches = source.matchAll(/function\s+([A-Za-z_$][\w$]*)\s*\(/g);
  for (const match of matches) {
    const name = match[1];
    if (functionOwners.has(name)) {
      failed = true;
      console.error(`FAIL duplicate function ${name}: ${functionOwners.get(name)} and ${file}`);
    } else {
      functionOwners.set(name, file);
    }
  }
}

try {
  new vm.Script(combinedSources.join('\n'), { filename: 'combined-apps-script.gs' });
  console.log('OK syntax combined Apps Script bundle');
} catch (error) {
  failed = true;
  console.error(`FAIL syntax combined Apps Script bundle: ${error.message}`);
}

const scriptsHtml = fs.readFileSync(path.join(root, 'Scripts.html'), 'utf8')
  .replace(/^\s*<script>\s*/, '')
  .replace(/\s*<\/script>\s*$/, '');
try {
  new vm.Script(scriptsHtml, { filename: 'Scripts.html' });
  console.log('OK syntax Scripts.html');
} catch (error) {
  failed = true;
  console.error(`FAIL syntax Scripts.html: ${error.message}`);
}

const indexHtml = fs.readFileSync(path.join(root, 'Index.html'), 'utf8');
const includeNames = new Set(Array.from(indexHtml.matchAll(/include\(['"]([^'"]+)['"]\)/g), match => match[1]));
for (const name of includeNames) {
  const file = `${name}.html`;
  if (!fs.existsSync(path.join(root, file))) {
    failed = true;
    console.error(`FAIL missing included HTML file: ${file}`);
  }
}

const htmlIds = new Set(Array.from(indexHtml.matchAll(/\bid="([^"]+)"/g), match => match[1]));
const referencedIds = new Set(Array.from(scriptsHtml.matchAll(/getElementById\(['"]([^'"]+)['"]\)/g), match => match[1]));
for (const id of referencedIds) {
  if (!htmlIds.has(id)) {
    failed = true;
    console.error(`FAIL missing HTML id referenced by Scripts.html: ${id}`);
  }
}

const serverMethods = new Set(Array.from(scriptsHtml.matchAll(/server\(['"]([A-Za-z0-9_]+)['"]/g), match => match[1]));
for (const method of serverMethods) {
  if (!functionOwners.has(method)) {
    failed = true;
    console.error(`FAIL missing server function called by frontend: ${method}`);
  }
}

const requiredFiles = [
  'appsscript.json', 'Code.gs', 'Config.gs', 'Utils.gs', 'Setup.gs',
  'DataService.gs', 'Migration.gs', 'DocumentService.gs', 'EmailService.gs',
  'FinanceService.gs', 'ExportService.gs', 'Index.html', 'Styles.html', 'Scripts.html'
];
for (const file of requiredFiles) {
  if (!fs.existsSync(path.join(root, file))) {
    failed = true;
    console.error(`FAIL missing ${file}`);
  }
}

if (failed) process.exit(1);
console.log(`PASS ${gsFiles.length} Apps Script files and frontend script`);
