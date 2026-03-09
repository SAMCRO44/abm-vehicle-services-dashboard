/**
 * Analyze key sheets and formulas in Moves 3.5.26 (1).xlsx
 * Run from project root:
 *   node scripts/analyze-moves-formulas.js
 */
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const FILE_NAME = 'Moves 3.5.26 (1).xlsx';
const IMPORTANT_SHEETS = [
  'Names-IDs',
  'Field Ops Task',
  'Low Performers',
  'Copy of Scans per hour',
  'Summary'
];

const downloadsDir = path.join(process.env.USERPROFILE || process.env.HOME, 'Downloads');
const filePath = path.join(downloadsDir, FILE_NAME);

if (!fs.existsSync(filePath)) {
  console.error('File not found at', filePath);
  process.exit(1);
}

function collectFormulas(sheet) {
  const formulas = [];
  for (const addr of Object.keys(sheet)) {
    if (addr[0] === '!') continue;
    const cell = sheet[addr];
    if (cell && cell.f) {
      formulas.push({
        addr,
        f: cell.f,
        v: cell.v
      });
    }
  }
  return formulas;
}

const wb = XLSX.readFile(filePath);
const result = { file: FILE_NAME, sheets: [] };

for (const sheetName of wb.SheetNames) {
  if (!IMPORTANT_SHEETS.includes(sheetName)) continue;
  const sheet = wb.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  const headers = data[0] || [];
  const formulas = collectFormulas(sheet);

  // Take a small sample of formulas for readability
  const sampleFormulas = formulas.slice(0, 80);

  result.sheets.push({
    name: sheetName,
    headers,
    rowCount: data.length,
    totalFormulas: formulas.length,
    sampleFormulas
  });
}

console.log(JSON.stringify(result, null, 2));

