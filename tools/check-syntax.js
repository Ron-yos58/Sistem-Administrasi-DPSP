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

const scriptFiles = [
  'ScriptsCore.html',
  'ScriptsApp.html',
  'ScriptsEvents.html',
  'ScriptsRender.html',
  'ScriptsForm.html',
  'ScriptsDetail.html',
  'ScriptsAdmin.html'
];
const scriptsHtml = scriptFiles.map(file => {
  return fs.readFileSync(path.join(root, file), 'utf8')
    .replace(/^\s*<script>\s*/, '')
    .replace(/\s*<\/script>\s*$/, '');
}).join('\n');
try {
  new vm.Script(scriptsHtml, { filename: 'CombinedScripts.js' });
  console.log('OK syntax combined frontend scripts');
} catch (error) {
  failed = true;
  console.error(`FAIL syntax combined frontend scripts: ${error.message}`);
}

const indexHtml = fs.readFileSync(path.join(root, 'Index.html'), 'utf8');
const stylesHtml = fs.readFileSync(path.join(root, 'Styles.html'), 'utf8');
if (/<script\b[^>]*\bsrc=["'][^"']*lucide/i.test(indexHtml)) {
  failed = true;
  console.error('FAIL app startup must not depend on an external Lucide script');
}
if (/data-lucide|\bwindow\.lucide\b|lucide-ready/.test(indexHtml + scriptsHtml + stylesHtml)) {
  failed = true;
  console.error('FAIL Lucide must be fully disabled');
}
if (scriptsHtml.includes('refreshIcons(')) {
  failed = true;
  console.error('FAIL removed icon runtime is still referenced');
}
if (scriptsHtml.includes('?.')) {
  failed = true;
  console.error('FAIL optional chaining is not allowed in the Apps Script web UI bundle');
}
if (scriptsHtml.includes("document.addEventListener('DOMContentLoaded', () => App.init()")) {
  failed = true;
  console.error('FAIL app startup must not depend solely on DOMContentLoaded');
}
const getIncludes = (content) => {
  return Array.from(content.matchAll(/include\(['"]([^'"]+)['"]\)/g), match => match[1]);
};
const includeNames = new Set();
getIncludes(indexHtml).forEach(name => includeNames.add(name));
if (fs.existsSync(path.join(root, 'Scripts.html'))) {
  const scriptsIndex = fs.readFileSync(path.join(root, 'Scripts.html'), 'utf8');
  getIncludes(scriptsIndex).forEach(name => includeNames.add(name));
}
for (const name of includeNames) {
  const file = `${name}.html`;
  if (!fs.existsSync(path.join(root, file))) {
    failed = true;
    console.error(`FAIL missing included HTML file: ${file}`);
  }
}

const htmlIdList = Array.from(indexHtml.matchAll(/\bid="([^"]+)"/g), match => match[1]);
const htmlIds = new Set(htmlIdList);
const duplicateHtmlIds = htmlIdList.filter((id, index) => htmlIdList.indexOf(id) !== index);
for (const id of new Set(duplicateHtmlIds)) {
  failed = true;
  console.error(`FAIL duplicate HTML id: ${id}`);
}
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

const keywords = new Set(['if', 'for', 'while', 'switch', 'catch', 'function']);
const appMethodNames = Array.from(
  scriptsHtml.matchAll(/^\s{4}([A-Za-z_$][\w$]*)\s*\([^)]*\)\s*\{/gm),
  match => match[1]
).filter(name => !keywords.has(name));
const duplicateAppMethods = appMethodNames.filter((name, index) => appMethodNames.indexOf(name) !== index);
for (const name of new Set(duplicateAppMethods)) {
  failed = true;
  console.error(`FAIL duplicate App method: ${name}`);
}

if (scriptsHtml.includes('lifecycleBadge(')) {
  failed = true;
  console.error('FAIL redundant lifecycleBadge renderer must not return');
}
if ((scriptsHtml.match(/data-action="process-all"/g) || []).length > 1) {
  failed = true;
  console.error('FAIL duplicate process-all action');
}
if (scriptsHtml.includes('<h3>Sesi kegiatan</h3>')) {
  failed = true;
  console.error('FAIL duplicate standalone Sesi kegiatan block');
}
if (/\.loader-spinner[^{]*\{[^}]*animation:\s*none/i.test(stylesHtml)) {
  failed = true;
  console.error('FAIL boot loader animation is disabled');
}

const topbarSource = indexHtml.match(/<header class="topbar">[\s\S]*?<\/header>/);
if (topbarSource && /data-view-target="form"/.test(topbarSource[0])) {
  failed = true;
  console.error('FAIL redundant create-request action in topbar');
}

const saveMethodStart = scriptsHtml.indexOf('async saveRequest(event)');
const resetMethodStart = scriptsHtml.indexOf('\n    resetForm() {', saveMethodStart);
const saveMethodSource = saveMethodStart !== -1 && resetMethodStart !== -1
  ? scriptsHtml.slice(saveMethodStart, resetMethodStart)
  : '';
if (saveMethodSource.includes('openRequestDetail(')) {
  failed = true;
  console.error('FAIL saveRequest must not open Detail automatically');
}

const requiredFiles = [
  'appsscript.json', 'Code.gs', 'Config.gs', 'Utils.gs', 'Setup.gs',
  'DataService.gs', 'DocumentService.gs', 'EmailService.gs',
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
