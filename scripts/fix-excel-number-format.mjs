#!/usr/bin/env node
import path from 'node:path';
import process from 'node:process';
import XLSX from 'xlsx';

const numberHeaders = new Set([
  '申报次数',
  '模糊通过次数',
  '未通过次数',
  '举证未通过次数',
  '处理单量',
  '总审核量',
  '加权审核量',
  '加权处理量',
  '一审审核量',
  '一审通过量',
  '精准通过量',
  '未通过量',
  '举证拒绝量',
  '模糊通过量',
  '平均处理时长',
  '超时次数',
  '通过率',
  '精准通过率',
  '举证准确率',
]);

const usage = () => {
  console.log(`用法:
  node scripts/fix-excel-number-format.mjs <Excel文件路径>
  node scripts/fix-excel-number-format.mjs <Excel文件路径> --in-place

默认会生成「原文件名-数字修复.xlsx」，不会覆盖原文件。`);
};

const inputFile = process.argv.find((arg, index) => index > 1 && !arg.startsWith('--'));
const inPlace = process.argv.includes('--in-place');

if (!inputFile) {
  usage();
  process.exit(1);
}

const normalizeRef = (sheet) => {
  const cellAddresses = Object.keys(sheet).filter((key) => !key.startsWith('!'));
  if (!cellAddresses.length) return;

  let maxRow = 0;
  let maxCol = 0;
  for (const address of cellAddresses) {
    const cell = XLSX.utils.decode_cell(address);
    if (cell.r > maxRow) maxRow = cell.r;
    if (cell.c > maxCol) maxCol = cell.c;
  }

  sheet['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow, c: maxCol } });
};

const toNumericValue = (value) => {
  if (typeof value === 'number' && Number.isFinite(value)) return value;

  const text = String(value ?? '')
    .replace(/,/g, '')
    .replace(/%/g, '')
    .trim();

  if (!text) return null;

  const parsed = Number(text);
  return Number.isFinite(parsed) ? parsed : null;
};

const workbook = XLSX.readFile(inputFile, {
  cellDates: true,
  raw: true,
  sheetStubs: true,
});

let convertedCells = 0;

for (const sheetName of workbook.SheetNames) {
  const sheet = workbook.Sheets[sheetName];
  normalizeRef(sheet);

  const range = XLSX.utils.decode_range(sheet['!ref']);
  const columnHeaders = new Map();

  for (let col = range.s.c; col <= range.e.c; col += 1) {
    const address = XLSX.utils.encode_cell({ r: range.s.r, c: col });
    const header = String(sheet[address]?.v ?? '').trim();
    if (numberHeaders.has(header)) {
      columnHeaders.set(col, header);
    }
  }

  for (let row = range.s.r + 1; row <= range.e.r; row += 1) {
    for (const col of columnHeaders.keys()) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = sheet[address];
      if (!cell) continue;

      const numericValue = toNumericValue(cell.v);
      if (numericValue === null) continue;

      cell.t = 'n';
      cell.v = numericValue;
      delete cell.w;
      convertedCells += 1;
    }
  }
}

const parsedPath = path.parse(path.resolve(inputFile));
const outputFile = inPlace
  ? path.resolve(inputFile)
  : path.join(parsedPath.dir, `${parsedPath.name}-数字修复${parsedPath.ext || '.xlsx'}`);

XLSX.writeFile(workbook, outputFile, {
  bookType: 'xlsx',
  cellDates: true,
});

console.log(JSON.stringify({
  inputFile: path.resolve(inputFile),
  outputFile,
  inPlace,
  convertedCells,
}, null, 2));
