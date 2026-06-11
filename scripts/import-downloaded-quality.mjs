#!/usr/bin/env node
import fs from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import XLSX from 'xlsx';

const DEFAULT_DOWNLOAD_DIR = path.join(process.env.HOME || '.', 'Downloads', '预质检每日数据');
const downloadDir = process.env.QUALITY_DOWNLOAD_DIR || DEFAULT_DOWNLOAD_DIR;
const apiBaseUrls = (process.env.QUALITY_API_BASE_URLS || process.env.QUALITY_API_BASE_URL || 'http://127.0.0.1:3000')
  .split(',')
  .map((url) => url.trim().replace(/\/$/, ''))
  .filter(Boolean);
const explicitFile = process.env.QUALITY_IMPORT_FILE || process.argv[2] || '';
const dryRun = process.env.QUALITY_IMPORT_DRY_RUN === '1';

const requiredHeaders = [
  ['场次', 'sale_type'],
  ['批次', 'batch_flag'],
  ['申报次数', 'declare_cnt'],
  ['模糊通过次数', 'mix_pass_cnt'],
  ['未通过次数', 'failed_cnt'],
  ['举证未通过次数', 'proof_failed_cnt'],
];

const dateHeaders = ['第一次线审完成时间', '日期', '第一次线审完成日期', 'online1_complete_date'];
const attributeHeaders = ['属性标签', '属性项', 'property_tag'];
const auditorHeaders = ['第一次在线审核人', '审核人', '第一次线审审核人', '一审审核人', 'auditor_name'];
const auditorTeamHeaders = ['审核团队', '团队', 'auditor_team'];

const findHeader = (headers, aliases) => aliases.find((header) => headers.includes(header));

const getValue = (record, aliases) => {
  const header = aliases.find((candidate) => Object.prototype.hasOwnProperty.call(record, candidate));
  return header ? record[header] : undefined;
};

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
    requiredHeaders.every((aliases) => aliases.some((header) => headerSet.has(header))) &&
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

const inspectWorkbookHeaders = (workbook) =>
  workbook.SheetNames.map((sheetName) => {
    const rows = sheetToRows(workbook.Sheets[sheetName]);
    return { sheetName, headers: Object.keys(rows[0] || {}) };
  });

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

    const dateHeader = findHeader(headers, dateHeaders);
    const attributeHeader = findHeader(headers, attributeHeaders);
    const auditorHeader = findHeader(headers, auditorHeaders);
    const auditorTeamHeader = findHeader(headers, auditorTeamHeaders);

    const parsedRows = rows
      .map((record) => ({
        date: normalizeDate(record[dateHeader]),
        auditor: auditorHeader ? String(record[auditorHeader] ?? '').trim() : '',
        auditorTeam: auditorTeamHeader ? String(record[auditorTeamHeader] ?? '').trim() : '',
        session: String(getValue(record, ['场次', 'sale_type']) ?? '').trim(),
        batch: String(getValue(record, ['批次', 'batch_flag']) ?? '').trim(),
        category: String(record['属性项分类'] ?? '').trim(),
        attribute: String(record[attributeHeader] ?? '').trim(),
        declarations: toNumber(getValue(record, ['申报次数', 'declare_cnt'])),
        exactPasses: getValue(record, ['精准通过次数', 'precise_pass_cnt']) !== undefined
          ? toNumber(getValue(record, ['精准通过次数', 'precise_pass_cnt']))
          : undefined,
        ambiguousPasses: toNumber(getValue(record, ['模糊通过次数', 'mix_pass_cnt'])),
        rejects: toNumber(getValue(record, ['未通过次数', 'failed_cnt'])),
        proofRejects: toNumber(getValue(record, ['举证未通过次数', 'proof_failed_cnt'])),
      }))
      .filter((row) => row.date && row.session && row.batch && row.attribute && Number.isFinite(row.declarations));

    return {
      rows: parsedRows,
      importedAt: new Date().toISOString(),
      sourceName: path.basename(filePath),
      sheetName,
    };
  }

  const inspectedHeaders = inspectWorkbookHeaders(workbook)
    .map((item) => `${item.sheetName}: ${item.headers.join(', ') || '(空表头)'}`)
    .join(' | ');
  throw new Error(`未找到质量数据工作表：${filePath}。检测到的表头：${inspectedHeaders || '(无工作表)'}`);
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

const resolveParsedWorkbook = async () => {
  if (explicitFile) {
    const filePath = await resolveFile();
    return parseWorkbook(filePath);
  }

  const files = await listExcelFiles();
  if (!files.length) {
    throw new Error(`下载目录暂无 Excel 文件：${downloadDir}`);
  }

  const errors = [];
  for (const file of files) {
    try {
      return await parseWorkbook(file.filePath);
    } catch (error) {
      errors.push(error instanceof Error ? error.message : String(error));
    }
  }

  throw new Error(`下载目录内没有可解析的质量 Excel：${downloadDir}\n${errors.join('\n')}`);
};

const mergeDataset = async (payload, apiBaseUrl) => {
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
  const parsed = await resolveParsedWorkbook();

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

  const results = [];
  for (const apiBaseUrl of apiBaseUrls) {
    const merged = await mergeDataset(parsed, apiBaseUrl);
    results.push({
      apiBaseUrl,
      totalRows: merged.rows?.length ?? 0,
    });
  }

  console.log(JSON.stringify({
    importedFile: parsed.sourceName,
    sheetName: parsed.sheetName,
    importedRows: parsed.rows.length,
    targets: results,
  }, null, 2));
};

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
