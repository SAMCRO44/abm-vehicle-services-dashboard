/**
 * One-time script to analyze Excel file structure.
 * Run: node scripts/analyze-xlsx.js
 */
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const files = [
  'Export - 2026-03-06T171825.716.xlsx',
  'Export - 2026-03-06T171818.487.xlsx',
  'Export - 2026-03-06T171814.704.xlsx',
  'Export - 2026-03-06T171816.546.xlsx',
  'Moves 3.5.26 (1).xlsx'
];

const downloads = path.join(process.env.USERPROFILE || process.env.HOME, 'Downloads');

function analyzeFile(filePath) {
  if (!fs.existsSync(filePath)) {
    return { error: 'File not found', path: filePath };
  }
  const wb = XLSX.readFile(filePath);
  const result = { name: path.basename(filePath), sheets: [] };
  for (const sheetName of wb.SheetNames) {
    const sheet = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
    const headers = (data[0] || []).filter(Boolean);
    const rowCount = data.length - (headers.length ? 1 : 0);
    const sample = data.slice(0, 4);
    result.sheets.push({
      name: sheetName,
      headers,
      rowCount,
      sample
    });
  }
  return result;
}

console.log(JSON.stringify({
  downloadsFolder: downloads,
  files: files.map(f => {
    const full = path.join(downloads, f);
    return analyzeFile(full);
  })
}, null, 2));
