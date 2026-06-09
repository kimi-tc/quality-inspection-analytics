#!/usr/bin/env node
import fs from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import XLSX from 'xlsx';

const DEFAULT_DOWNLOAD_DIR = path.join(process.env.HOME || '.', 'Downloads', '预质检每日数据');
const downloadDir = process.env.QUALITY_DOWNLOAD_DIR || DEFAULT_DOWNLOAD_DIR;
const apiBaseUrl = (process.env.QUALITY_API_BASE_URL || 'http://127.0.0.1:3000').replace(/\/$/, '');
const explicitFile = process.env.QUALITY_IMPORT_FILE || process.argv[2] || '';
const dryRun = process.env.QUALITY_IMPORT_DRY_RUN === '1';

const requiredHeaders = [
  '场次',
  '批次',
  '申报次数',
  '模糊通过次数',
  '未通过次数',
  '举证未通过次数',
];

const dateHeaders = ['第一次线审完成时间', '日期', '第一次线审完成日期'];
const attributeHeaders = ['属性标签', '属性项'];
const auditorHeaders = ['第一次在线审核人', '审核人', '第一次线审审核人', '一审审核人'];
const auditorTeamHeaders = ['审核团队', '团队'];

const toNumber = (value) => {
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  const normalized = String(value ?? '')
    .replace(/,/g, '')
    .replace(/%/g, '')
    .trim();
  if (!normalized) return 0;
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : 0;
};

const normalizeDate = (value) => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return `${parsed.y}-${String(parsed.m).padStart(2, '0')}-${String(parsed.d).padStart(2, '0')}`;
    }
  }

  const text = String(value ?? '').trim();
  const match = text.match(/(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})/);
  if (match) {
    return `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
  }

  return text.slice(0, 10);
};

const hasHeaders = (headers) => {
  const headerSet = new Set(headers);
  return (
    requiredHeaders.every((header) => headerSet.has(header)) &&
    dateHeaders.some((header) => headerSet.has(header)) &&
    attributeHeaders.some((header) => headerSet.has(header))
  );
};

const sheetToRows = (sheet) => {
  const cellAddresses = Object.keys(sheet).filter((key) => !key.startsWith('!'));
  if (cellAddresses.length) {
    let maxRow = 0;
    let maxCol = 0;
    for (const address of cellAddresses) {
      const cell = XLSX.utils.decode_cell(address);
      if (cell.r > maxRow) maxRow = cell.r;
      if (cell.c > maxCol) maxCol = cell.c;
    }
    sheet['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow, c: maxCol } });
  }

  return XLSX.utils.sheet_to_json(sheet, {
    defval: '',
    raw: false,
  });
};

const parseWorkbook = async (filePath) => {
  const workbook = XLSX.readFile(filePath, {
    cellDates: true,
    raw: false,
  });

  for (const sheetName of workbook.SheetNames) {
    const worksheet = workbook.Sheets[sheetName];
    const rows = sheetToRows(worksheet);

    if (!rows.length) continue;

    const headers = Object.keys(rows[0] || {});
    if (!hasHeaders(headers)) continue;

    const dateHeader = dateHeaders.find((header) => headers.includes(header));
    const attributeHeader = attributeHeaders.find((header) => headers.includes(header));
    const auditorHeader = auditorHeaders.find((header) => headers.includes(header));
    const auditorTeamHeader = auditorTeamHeaders.find((header) => headers.includes(header));

    const parsedRows = rows
      .map((record) => ({
        date: normalizeDate(record[dateHeader]),
        auditor: auditorHeader ? String(record[auditorHeader] ?? '').trim() : '',
        auditorTeam: auditorTeamHeader ? String(record[auditorTeamHeader] ?? '').trim() : '',
        session: String(record['场次'] ?? '').trim(),
        batch: String(record['批次'] ?? '').trim(),
        category: String(record['属性项分类'] ?? '').trim(),
        attribute: String(record[attributeHeader] ?? '').trim(),
        declarations: toNumber(record['申报次数']),
        ambiguousPasses: toNumber(record['模糊通过次数']),
        rejects: toNumber(record['未通过次数']),
        proofRejects: toNumber(record['举证未通过次数']),
      }))
      .filter((row) => row.date && row.session && row.batch && row.attribute && Number.isFinite(row.declarations));

    return {
      rows: parsedRows,
      importedAt: new Date().toISOString(),
      sourceName: path.basename(filePath),
      sheetName,
    };
  }

  throw new Error(`未找到质量数据工作表：${filePath}`);
};

const listExcelFiles = async () => {
  const entries = await fs.readdir(downloadDir, { withFileTypes: true }).catch(() => []);
  const files = await Promise.all(
    entries
      .filter((entry) => entry.isFile())
      .filter((entry) => /\.(xlsx|xls)$/i.test(entry.name))
      .filter((entry) => !entry.name.startsWith('~$'))
      .map(async (entry) => {
        const filePath = path.join(downloadDir, entry.name);
        const stat = await fs.stat(filePath);
        return { filePath, mtimeMs: stat.mtimeMs };
      }),
  );

  return files.sort((a, b) => b.mtimeMs - a.mtimeMs);
};

const resolveFile = async () => {
  if (explicitFile) {
    return path.resolve(explicitFile);
  }

  const files = await listExcelFiles();
  if (!files.length) {
    throw new Error(`下载目录暂无 Excel 文件：${downloadDir}`);
  }

  return files[0].filePath;
};

const mergeDataset = async (payload) => {
  const response = await fetch(`${apiBaseUrl}/api/dataset/merge`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  const body = await response.json().catch(() => ({}));

  if (!response.ok) {
    throw new Error(body.message || `导入失败，HTTP ${response.status}`);
  }

  return body;
};

const main = async () => {
  const filePath = await resolveFile();
  const parsed = await parseWorkbook(filePath);

  if (!parsed.rows.length) {
    throw new Error(`文件没有解析到有效质量数据：${filePath}`);
  }

  if (dryRun) {
    console.log(JSON.stringify({
      dryRun: true,
      importedFile: parsed.sourceName,
      sheetName: parsed.sheetName,
      parsedRows: parsed.rows.length,
      firstRow: parsed.rows[0],
    }, null, 2));
    return;
  }

  const merged = await mergeDataset(parsed);
  console.log(JSON.stringify({
    importedFile: parsed.sourceName,
    sheetName: parsed.sheetName,
    importedRows: parsed.rows.length,
    totalRows: merged.rows?.length ?? 0,
    apiBaseUrl,
  }, null, 2));
};

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
