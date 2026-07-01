import React, { ChangeEvent, useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { DayPicker, type DateRange } from 'react-day-picker';
import { addDays, addMonths, format, parseISO, startOfMonth, startOfWeek, subMonths } from 'date-fns';
import { zhCN } from 'date-fns/locale';
import {
  ResponsiveContainer,
  AreaChart,
  Area,
  LineChart,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  BarChart,
  Bar,
  PieChart,
  Pie,
  Cell,
} from 'recharts';
import {
  Upload,
  CalendarDays,
  Layers3,
  Tags,
  Boxes,
  ShieldCheck,
  Target,
  CircleSlash,
  Ban,
  Database,
  Trash2,
  ChevronDown,
  FileSpreadsheet,
  HardDriveDownload,
  X,
  ChevronLeft,
  ChevronRight,
  House,
  GitCompareArrows,
  Users,
  BrainCircuit,
  ClipboardList,
  ArrowUpRight,
  ArrowDownRight,
  Minus,
  Check,
} from 'lucide-react';
import { motion } from 'motion/react';
import {
  AuditorTeamDictionaryResponse,
  AuditorTeamEntry,
  EfficiencyDatasetResponse,
  EfficiencyRow,
  ImportRecord,
  ImportedRow,
  AiAnalysisResponse,
  MetricsCardData,
  ParsedEfficiencyWorkbook,
  ParsedWorkbook,
  PropertyCategoryDictionaryResponse,
  PropertyCategoryEntry,
  SharedDatasetResponse,
} from './types';

const REQUIRED_HEADERS = [
  ['第一次线审完成时间', 'online1_complete_date'],
  ['场次', 'sale_type'],
  ['批次', 'batch_flag'],
  ['属性标签', 'property_tag'],
  ['申报次数', 'declare_cnt'],
  ['模糊通过次数', 'mix_pass_cnt'],
  ['未通过次数', 'failed_cnt'],
  ['举证未通过次数', 'proof_failed_cnt'],
] as const;

const REQUIRED_EFFICIENCY_HEADERS = [
  '日期',
  '员工姓名',
  '团队',
  '处理单量',
  '平均处理时长',
  '超时次数',
] as const;

const REQUIRED_AUDIT_EFFICIENCY_HEADERS = [
  '日期',
  '员工姓名',
  '团队',
  '总审核量',
  '加权审核量',
  '一审审核量',
  '一审通过量',
  '精准通过量',
  '未通过量',
  '举证拒绝量',
  '模糊通过量',
] as const;

const REQUIRED_MAIL_EFFICIENCY_HEADERS = [
  'dt',
  'employee_name',
  'team',
  'total_audit_cnt',
  'weighted_audit_cnt',
  'first_audit_cnt',
  'first_audit_pass_cnt',
  'accurate_pass_cnt',
  'first_audit_reject_cnt',
  'proof_refusal_cnt',
  'ambiguous_pass_cnt',
] as const;

const ALL_OPTION = '全部';
const AMBIGUOUS_RATE_TARGET = 0.07;
type ViewKey = 'overview' | 'compare' | 'attribute' | 'efficiency' | 'ai' | 'import' | 'dictionary';
type CompareQualityMetric = 'proofAccuracy' | 'exactPassRate' | 'ambiguousRate' | 'rejectRate';
type CompareDimension = 'session' | 'attribute' | 'batch' | 'auditor' | 'auditorTeam';
const PROPERTY_CATEGORY_OPTIONS = ['维修项', '外观项', '功能项', 'SKU项', '其他', '售后补充项'] as const;
const CATEGORY_COLORS = ['#0f766e', '#1d4ed8', '#f59e0b', '#ef4444', '#8b5cf6', '#14b8a6'];
const SESSION_COLORS = ['#0ea5e9', '#f97316', '#10b981', '#6366f1', '#f43f5e', '#64748b', '#84cc16', '#a855f7'];
const COMPARE_DIMENSIONS: Array<{ key: CompareDimension; label: string; title: string; description: string }> = [
  {
    key: 'session',
    label: '场次',
    title: '场次 A / 场次 B 质量趋势',
    description: '在相同日期、批次和属性项条件下，按自然日对比两个场次的质量变化。',
  },
  {
    key: 'attribute',
    label: '属性项',
    title: '属性项 A / 属性项 B 质量趋势',
    description: '在相同日期、场次和批次条件下，按自然日对比两个属性项的质量变化。',
  },
  {
    key: 'batch',
    label: '批次',
    title: '批次 A / 批次 B 质量趋势',
    description: '在相同日期、场次和属性项条件下，按自然日对比两个批次的质量变化。',
  },
  {
    key: 'auditor',
    label: '审核人',
    title: '审核人 A / 审核人 B 质量趋势',
    description: '在相同日期、场次、批次和属性项条件下，按自然日对比两个审核人的质量变化。',
  },
  {
    key: 'auditorTeam',
    label: '团队',
    title: '团队 A / 团队 B 质量趋势',
    description: '在相同日期、场次、批次和属性项条件下，按自然日对比两个审核团队的质量变化。',
  },
];
const COMPARE_QUALITY_METRICS: Array<{
  key: CompareQualityMetric;
  label: string;
  dash?: string;
}> = [
  { key: 'proofAccuracy', label: '举证准确率' },
  { key: 'exactPassRate', label: '精准通过率', dash: '7 5' },
  { key: 'ambiguousRate', label: '模棱两可率', dash: '2 5' },
  { key: 'rejectRate', label: '拒绝率', dash: '12 5 2 5' },
];
const COMPARE_TREND_DATA_KEYS: Record<
  CompareQualityMetric,
  { left: string; right: string }
> = {
  proofAccuracy: { left: 'leftProofAccuracy', right: 'rightProofAccuracy' },
  exactPassRate: { left: 'leftExactPassRate', right: 'rightExactPassRate' },
  ambiguousRate: { left: 'leftAmbiguousRate', right: 'rightAmbiguousRate' },
  rejectRate: { left: 'leftRejectRate', right: 'rightRejectRate' },
};
const DEEPSEEK_MODEL_OPTIONS = [
  'deepseek-chat',
  'deepseek-reasoner',
  'deepseek-v4-flash',
  'deepseek-v4-pro',
] as const;
const CUSTOM_MODEL_OPTION = '自定义模型';

const normalizePropertyName = (value: string) => value.trim().replace(/与购买时不一致$/, '');

const normalizeDictionaryCategory = (value: string) => {
  const normalizedValue = value.trim();
  const legacyCategoryMap: Record<string, string> = {
    主观项: '外观项',
    零售附加项: '售后补充项',
    零售补充项: '售后补充项',
  };
  const mappedValue = legacyCategoryMap[normalizedValue] ?? normalizedValue;

  return PROPERTY_CATEGORY_OPTIONS.includes(mappedValue as (typeof PROPERTY_CATEGORY_OPTIONS)[number])
    ? mappedValue
    : '其他';
};

const buildDictionaryMap = (entries: PropertyCategoryEntry[]) =>
  new Map(entries.map((entry) => [normalizePropertyName(entry.propertyName), normalizeDictionaryCategory(entry.category)]));

const normalizeAuditorName = (value: string) => value.trim();

const buildAuditorTeamMap = (entries: AuditorTeamEntry[]) =>
  new Map(entries.map((entry) => [normalizeAuditorName(entry.auditorName), entry.team.trim()]));

const resolveAuditorTeam = (auditor: string, explicitTeam: string, dictionary: AuditorTeamEntry[]) => {
  const normalizedTeam = explicitTeam.trim();
  const dictionaryTeam = buildAuditorTeamMap(dictionary).get(normalizeAuditorName(auditor));

  return dictionaryTeam ?? normalizedTeam;
};

const resolveCategory = (category: string, attribute: string, dictionary: PropertyCategoryEntry[]) => {
  const normalizedAttribute = attribute.trim();
  const normalizedCategory = category.trim();

  if (normalizedAttribute.includes('售后补充项') || normalizedCategory.includes('售后补充项')) {
    return '售后补充项';
  }

  const dictionaryCategory = buildDictionaryMap(dictionary).get(normalizePropertyName(normalizedAttribute));

  return dictionaryCategory ?? normalizeDictionaryCategory(normalizedCategory || '其他');
};

const emptyWorkbook: ParsedWorkbook = {
  rows: [],
  importedAt: '',
  sourceName: '',
};

const emptyEfficiencyWorkbook: ParsedEfficiencyWorkbook = {
  rows: [],
  importedAt: '',
  sourceName: '',
};

const toNumber = (value: unknown): number => {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }

  const parsed = Number(
    String(value ?? '')
      .replace(/,/g, '')
      .replace(/%/g, '')
      .trim(),
  );
  return Number.isFinite(parsed) ? parsed : 0;
};

const safeRateFromCounts = (numerator: number, denominator: number) => (denominator ? numerator / denominator : 0);

const getRecordValue = (record: Record<string, unknown>, headers: string[]) => {
  for (const header of headers) {
    if (Object.prototype.hasOwnProperty.call(record, header)) {
      return record[header];
    }
  }

  return undefined;
};

const hasAnyHeader = (headerSet: Set<string>, headers: readonly string[]) =>
  headers.some((header) => headerSet.has(header));

const normalizeDate = (value: unknown): string => {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      const month = String(parsed.m).padStart(2, '0');
      const day = String(parsed.d).padStart(2, '0');
      return `${parsed.y}-${month}-${day}`;
    }
  }

  const raw = String(value ?? '').trim();
  if (!raw) {
    return '';
  }

  return raw.replace(/\//g, '-');
};

const sheetToRows = (sheet: XLSX.WorkSheet) => {
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

  return XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
    defval: '',
    raw: false,
  });
};

const pickDataSheet = (workbook: XLSX.WorkBook) => {
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const json = sheetToRows(sheet);

    if (!json.length) {
      continue;
    }

    const headerSet = new Set(Object.keys(json[0]));
    const matchedHeaders = REQUIRED_HEADERS.filter((headers) => hasAnyHeader(headerSet, headers));

    if (matchedHeaders.length === REQUIRED_HEADERS.length) {
      return { sheetName, json };
    }
  }

  return null;
};

const pickEfficiencySheet = (workbook: XLSX.WorkBook) => {
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const json = sheetToRows(sheet);

    if (!json.length) {
      continue;
    }

    const headerSet = new Set(Object.keys(json[0]));
    const matchedLegacyHeaders = REQUIRED_EFFICIENCY_HEADERS.filter((header) => headerSet.has(header));
    const matchedAuditHeaders = REQUIRED_AUDIT_EFFICIENCY_HEADERS.filter((header) => headerSet.has(header));
    const matchedMailHeaders = REQUIRED_MAIL_EFFICIENCY_HEADERS.filter((header) => headerSet.has(header));

    if (
      matchedLegacyHeaders.length === REQUIRED_EFFICIENCY_HEADERS.length ||
      matchedAuditHeaders.length === REQUIRED_AUDIT_EFFICIENCY_HEADERS.length ||
      matchedMailHeaders.length === REQUIRED_MAIL_EFFICIENCY_HEADERS.length
    ) {
      return { sheetName, json };
    }
  }

  return null;
};

const parseWorkbookFile = async (
  file: File,
  propertyCategoryDictionary: PropertyCategoryEntry[],
): Promise<ParsedWorkbook> => {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const matched = pickDataSheet(workbook);

  if (!matched) {
    throw new Error('未找到包含标准字段的工作表，请确认表头未被修改。');
  }

  const rows: ImportedRow[] = matched.json
    .map((record) => ({
      date: normalizeDate(getRecordValue(record, ['第一次线审完成时间', '日期', '第一次线审完成日期', 'online1_complete_date'])),
      auditor: String(
        getRecordValue(record, ['第一次在线审核人', '审核人', '第一次线审审核人', '一审审核人', 'auditor_name']) ?? '',
      ).trim(),
      auditorTeam: String(getRecordValue(record, ['审核团队', '团队', 'auditor_team']) ?? '').trim(),
      session: String(getRecordValue(record, ['场次', 'sale_type']) ?? '').trim(),
      batch: String(getRecordValue(record, ['批次', 'batch_flag']) ?? '').trim(),
      category: resolveCategory(
        String(record['属性项分类'] ?? ''),
        String(getRecordValue(record, ['属性标签', '属性项', 'property_tag']) ?? ''),
        propertyCategoryDictionary,
      ),
      attribute: String(getRecordValue(record, ['属性标签', '属性项', 'property_tag']) ?? '').trim(),
      declarations: toNumber(getRecordValue(record, ['申报次数', 'declare_cnt'])),
      exactPasses: getRecordValue(record, ['精准通过次数', 'precise_pass_cnt']) !== undefined
        ? toNumber(getRecordValue(record, ['精准通过次数', 'precise_pass_cnt']))
        : undefined,
      ambiguousPasses: toNumber(getRecordValue(record, ['模糊通过次数', 'mix_pass_cnt'])),
      rejects: toNumber(getRecordValue(record, ['未通过次数', 'failed_cnt'])),
      proofRejects: toNumber(getRecordValue(record, ['举证未通过次数', 'proof_failed_cnt'])),
    }))
    .filter(
      (row) =>
        row.date &&
        row.session &&
        row.batch &&
        row.attribute &&
        Number.isFinite(row.declarations),
    );

  return {
    rows,
    importedAt: new Date().toISOString(),
    sourceName: file.name,
  };
};

const parseEfficiencyWorkbookFile = async (file: File): Promise<ParsedEfficiencyWorkbook> => {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const matched = pickEfficiencySheet(workbook);

  if (!matched) {
    throw new Error('未找到包含人效标准字段的工作表，请确认表头包含新版审核人效字段，或旧版：日期、员工姓名、团队、处理单量、平均处理时长、超时次数。');
  }

  const rows: EfficiencyRow[] = matched.json
    .map((record) => {
      const firstAuditCount = toNumber(getRecordValue(record, ['一审审核量', 'first_audit_cnt']));
      const firstAuditPassCount = toNumber(getRecordValue(record, ['一审通过量', 'first_audit_pass_cnt']));
      const precisionPassCount = toNumber(getRecordValue(record, ['精准通过量', 'accurate_pass_cnt']));
      const proofRefusalCount = toNumber(getRecordValue(record, ['举证拒绝量', 'proof_refusal_cnt']));
      const ambiguousCount = toNumber(getRecordValue(record, ['模糊通过量', 'ambiguous_pass_cnt']));
      const handledCount = toNumber(getRecordValue(record, ['总审核量', '处理单量', 'total_audit_cnt']));

      return {
        date: normalizeDate(getRecordValue(record, ['日期', 'dt'])),
        employee: String(getRecordValue(record, ['员工姓名', 'employee_name']) ?? '').trim(),
        team: String(getRecordValue(record, ['团队', 'team']) ?? '').trim(),
        session: String(getRecordValue(record, ['场次', 'sale_type']) ?? '审核人效').trim() || '审核人效',
        batch: String(getRecordValue(record, ['批次', 'flag']) ?? '全部批次').trim() || '全部批次',
        handledCount,
        weightedHandledCount: toNumber(getRecordValue(record, ['加权审核量', '加权处理量', 'weighted_audit_cnt'])),
        firstAuditCount,
        firstAuditPassCount,
        precisionPassCount,
        auditNotPassCount: toNumber(getRecordValue(record, ['未通过量', 'first_audit_reject_cnt'])),
        proofRefusalCount,
        ambiguousCount,
        passRate: toNumber(getRecordValue(record, ['通过率', 'pass_rate'])) || safeRateFromCounts(firstAuditPassCount, firstAuditCount),
        precisionPassRate:
          toNumber(getRecordValue(record, ['精准通过率', 'accurate_pass_rate'])) ||
          safeRateFromCounts(precisionPassCount, firstAuditCount),
        proofAccuracy:
          toNumber(getRecordValue(record, ['举证准确率', 'proof_accuracy_rate'])) ||
          safeRateFromCounts(firstAuditCount - ambiguousCount - proofRefusalCount, firstAuditCount),
        avgHandleMinutes: toNumber(getRecordValue(record, ['平均处理时长', 'avg_handle_minutes'])),
        timeoutCount: toNumber(getRecordValue(record, ['超时次数', 'timeout_cnt'])),
      };
    })
    .filter((row) => row.date && row.employee && Number.isFinite(row.handledCount));

  return {
    rows,
    importedAt: new Date().toISOString(),
    sourceName: file.name,
  };
};

const parsePropertyCategoryDictionaryFile = async (file: File): Promise<PropertyCategoryEntry[]> => {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const json = sheetToRows(sheet);

    if (!json.length) {
      continue;
    }

    const headerSet = new Set(Object.keys(json[0]));
    if (!headerSet.has('属性项') || !headerSet.has('属性项分类')) {
      continue;
    }

    return json
      .map((record) => ({
        propertyName: String(record['属性项'] ?? '').trim(),
        category: normalizeDictionaryCategory(String(record['属性项分类'] ?? '')),
      }))
      .filter((entry) => entry.propertyName && entry.category);
  }

  throw new Error('未找到分类字典字段，请确认表头包含：属性项、属性项分类。');
};

const parseAuditorTeamDictionaryFile = async (file: File): Promise<AuditorTeamEntry[]> => {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });

  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const json = sheetToRows(sheet);

    if (!json.length) {
      continue;
    }

    const headerSet = new Set(Object.keys(json[0]));
    const auditorHeader = ['审核人', '第一次在线审核人', '员工姓名'].find((header) => headerSet.has(header));
    const teamHeader = ['团队', '审核团队'].find((header) => headerSet.has(header));
    if (!auditorHeader || !teamHeader) {
      continue;
    }

    return json
      .map((record) => ({
        auditorName: String(record[auditorHeader] ?? '').trim(),
        team: String(record[teamHeader] ?? '').trim(),
      }))
      .filter((entry) => entry.auditorName && entry.team);
  }

  throw new Error('未找到审核人团队字典字段，请确认表头包含：审核人、团队。');
};

const formatPercent = (value: number) => `${(value * 100).toFixed(2)}%`;
const formatInteger = (value: number) => value.toLocaleString('zh-CN');
const formatDateDisplay = (value: string) => (value === ALL_OPTION ? '' : value);
const formatMultiFilterDisplay = (values: string[], emptyLabel = ALL_OPTION) =>
  values.length === 0 ? emptyLabel : values.length <= 2 ? values.join('、') : `已选 ${values.length} 项`;
const parseDateValue = (value: string) => (value && value !== ALL_OPTION ? parseISO(value) : undefined);
const formatDateTime = (value: string) => (value ? new Date(value).toLocaleString('zh-CN') : '未导入');

const buildFallbackImportHistory = (
  dataset: Pick<ParsedWorkbook, 'sourceName' | 'importedAt'> | Pick<ParsedEfficiencyWorkbook, 'sourceName' | 'importedAt'>,
  dataType: ImportRecord['dataType'],
): ImportRecord[] =>
  dataset.sourceName
    ? dataset.sourceName.split(' + ').map((sourceName, index) => ({
        id: `${dataType}-fallback-${index}-${sourceName}`,
        sourceName,
        importedAt: dataset.importedAt,
        rowCount: 0,
        dataType,
      }))
    : [];

type FilterCriteria = {
  startDate: string;
  endDate: string;
  session: string | string[];
  batch: string | string[];
  attribute: string | string[];
  auditor?: string;
  auditorTeam?: string;
  auditorTeamDictionary?: AuditorTeamEntry[];
};

type WeekOption = {
  value: string;
  label: string;
  start: string;
  end: string;
};

const createOptions = (rows: ImportedRow[], key: keyof ImportedRow) =>
  [ALL_OPTION].concat(
    [...new Set(rows.map((row) => String(row[key] ?? '')).filter(Boolean))].sort((a, b) =>
      a.localeCompare(b, 'zh-CN'),
    ),
  );

const createDateOptions = (rows: Array<{ date: string }>) =>
  [ALL_OPTION].concat(
    [...new Set(rows.map((row) => row.date).filter(Boolean))].sort((a, b) => a.localeCompare(b)),
  );

const createStringOptions = <T,>(rows: T[], pickValue: (row: T) => string) =>
  [ALL_OPTION].concat(
    [...new Set(rows.map((row) => pickValue(row).trim()).filter(Boolean))].sort((a, b) =>
      a.localeCompare(b, 'zh-CN'),
    ),
  );

const createWeekOptions = (dates: string[]): WeekOption[] => {
  const weeks = new Map<string, WeekOption>();

  dates
    .filter((date) => date !== ALL_OPTION)
    .forEach((date) => {
      const weekStartDate = startOfWeek(parseISO(date), { weekStartsOn: 0 });
      const start = format(weekStartDate, 'yyyy-MM-dd');
      const end = format(addDays(weekStartDate, 6), 'yyyy-MM-dd');

      if (!weeks.has(start)) {
        weeks.set(start, {
          value: `${start}|${end}`,
          label: `${start} ~ ${end}`,
          start,
          end,
        });
      }
    });

  return [...weeks.values()].sort((a, b) => b.start.localeCompare(a.start));
};

const getDateCoverage = (rows: Array<{ date: string }>) => {
  const dates = rows
    .map((row) => row.date)
    .filter((date) => date && date !== ALL_OPTION)
    .sort((a, b) => a.localeCompare(b));

  return {
    start: dates[0] ?? '',
    end: dates[dates.length - 1] ?? '',
    count: dates.length,
  };
};

const getWeekValue = (startDate: string, endDate: string, weeks: WeekOption[]) =>
  weeks.find((week) => week.start === startDate && week.end === endDate)?.value ?? ALL_OPTION;

const filterRowsByCriteria = (rows: ImportedRow[], criteria: FilterCriteria) =>
  rows.filter((row) => {
    const startMatch = criteria.startDate === ALL_OPTION || row.date >= criteria.startDate;
    const endMatch = criteria.endDate === ALL_OPTION || row.date <= criteria.endDate;
    const sessionMatch = Array.isArray(criteria.session)
      ? criteria.session.length === 0 || criteria.session.includes(row.session)
      : criteria.session === ALL_OPTION || row.session === criteria.session;
    const batchMatch = Array.isArray(criteria.batch)
      ? criteria.batch.length === 0 || criteria.batch.includes(row.batch)
      : criteria.batch === ALL_OPTION || row.batch === criteria.batch;
    const attributeMatch = Array.isArray(criteria.attribute)
      ? criteria.attribute.length === 0 || criteria.attribute.includes(row.attribute)
      : criteria.attribute === ALL_OPTION || row.attribute === criteria.attribute;
    const auditorMatch = !criteria.auditor || criteria.auditor === ALL_OPTION || row.auditor === criteria.auditor;
    const rowAuditorTeam = resolveAuditorTeam(
      row.auditor ?? '',
      row.auditorTeam ?? '',
      criteria.auditorTeamDictionary ?? [],
    );
    const auditorTeamMatch =
      !criteria.auditorTeam || criteria.auditorTeam === ALL_OPTION || rowAuditorTeam === criteria.auditorTeam;

    return startMatch && endMatch && sessionMatch && batchMatch && attributeMatch && auditorMatch && auditorTeamMatch;
  });

const filterRowsByDateRange = <T extends { date: string }>(
  rows: T[],
  startDate: string,
  endDate: string,
) =>
  rows.filter((row) => {
    const startMatch = startDate === ALL_OPTION || row.date >= startDate;
    const endMatch = endDate === ALL_OPTION || row.date <= endDate;
    return startMatch && endMatch;
  });

const filterEfficiencyRows = (
  rows: EfficiencyRow[],
  startDate: string,
  endDate: string,
  team: string,
  sessions: string[],
) =>
  filterRowsByDateRange(rows, startDate, endDate).filter(
    (row) =>
      (team === ALL_OPTION || row.team === team) &&
      (sessions.length === 0 || sessions.includes(row.session)),
  );

const resolveWorkplace = (team: string) => {
  if (team.includes('常州')) return '常州';
  if (team.includes('上海')) return '上海';
  if (team.includes('老人') || team.includes('新人')) return '常州';
  if (team.includes('批')) return '上海';
  return '其他';
};

const resolveExactPasses = (row: Pick<ImportedRow, 'declarations' | 'ambiguousPasses' | 'rejects' | 'exactPasses'>) =>
  typeof row.exactPasses === 'number' && Number.isFinite(row.exactPasses)
    ? row.exactPasses
    : row.declarations - row.ambiguousPasses - row.rejects;

const aggregateMetrics = (rows: ImportedRow[]) => {
  const totals = rows.reduce(
    (acc, row) => {
      acc.declarations += row.declarations;
      acc.exactPasses += resolveExactPasses(row);
      acc.ambiguousPasses += row.ambiguousPasses;
      acc.rejects += row.rejects;
      acc.proofRejects += row.proofRejects;
      return acc;
    },
    { declarations: 0, exactPasses: 0, ambiguousPasses: 0, rejects: 0, proofRejects: 0 },
  );

  const safeDivide = (numerator: number, denominator: number) => (denominator ? numerator / denominator : 0);

  return {
    ...totals,
    proofAccuracy: safeDivide(
      totals.declarations - totals.ambiguousPasses - totals.proofRejects,
      totals.declarations,
    ),
    exactPassRate: safeDivide(totals.exactPasses, totals.declarations),
    ambiguousRate: safeDivide(totals.ambiguousPasses, totals.declarations),
    rejectRate: safeDivide(totals.rejects, totals.declarations),
  };
};

const aggregateTrend = (rows: ImportedRow[]) =>
  Object.values(
    rows.reduce<Record<string, { date: string; declarations: number; exactPasses: number; ambiguousPasses: number; rejects: number; proofRejects: number }>>(
      (acc, row) => {
        if (!acc[row.date]) {
          acc[row.date] = { date: row.date, declarations: 0, exactPasses: 0, ambiguousPasses: 0, rejects: 0, proofRejects: 0 };
        }

        acc[row.date].declarations += row.declarations;
        acc[row.date].exactPasses += resolveExactPasses(row);
        acc[row.date].ambiguousPasses += row.ambiguousPasses;
        acc[row.date].rejects += row.rejects;
        acc[row.date].proofRejects += row.proofRejects;
        return acc;
      },
      {},
    ),
  )
    .map((item) => ({
      ...item,
      proofAccuracy: item.declarations
        ? (item.declarations - item.ambiguousPasses - item.proofRejects) / item.declarations
        : 0,
      exactPassRate: item.declarations ? item.exactPasses / item.declarations : 0,
      ambiguousRate: item.declarations ? item.ambiguousPasses / item.declarations : 0,
      rejectRate: item.declarations ? item.rejects / item.declarations : 0,
    }))
    .sort((a, b) => a.date.localeCompare(b.date));

const aggregateAttributes = (rows: ImportedRow[]) =>
  Object.values(
    rows.reduce<Record<string, { attribute: string; declarations: number; rejects: number }>>((acc, row) => {
      if (!acc[row.attribute]) {
        acc[row.attribute] = { attribute: row.attribute, declarations: 0, rejects: 0 };
      }

      acc[row.attribute].declarations += row.declarations;
      acc[row.attribute].rejects += row.rejects;
      return acc;
    }, {}),
  )
    .sort((a, b) => b.declarations - a.declarations)
    .slice(0, 10);

const aggregateAmbiguousAttributes = (rows: ImportedRow[]) => {
  const totalAmbiguousPasses = rows.reduce((sum, row) => sum + row.ambiguousPasses, 0);

  return Object.values(
    rows.reduce<
      Record<
        string,
        {
          attribute: string;
          declarations: number;
          ambiguousPasses: number;
        }
      >
    >((acc, row) => {
      if (!acc[row.attribute]) {
        acc[row.attribute] = { attribute: row.attribute, declarations: 0, ambiguousPasses: 0 };
      }

      acc[row.attribute].declarations += row.declarations;
      acc[row.attribute].ambiguousPasses += row.ambiguousPasses;
      return acc;
    }, {}),
  )
    .map((item) => ({
      ...item,
      ambiguousRate: item.declarations ? item.ambiguousPasses / item.declarations : 0,
      contribution: totalAmbiguousPasses ? item.ambiguousPasses / totalAmbiguousPasses : 0,
    }))
    .filter((item) => item.ambiguousPasses > 0)
    .sort((a, b) => b.ambiguousPasses - a.ambiguousPasses)
    .slice(0, 10);
};

const aggregateCategories = (rows: ImportedRow[], propertyCategoryDictionary: PropertyCategoryEntry[]) =>
  Object.values(
    rows.reduce<
      Record<
        string,
        {
          name: string;
          declarations: number;
          exactPasses: number;
          ambiguousPasses: number;
          rejects: number;
          proofRejects: number;
        }
      >
    >((acc, row) => {
      const key = resolveCategory(row.category, row.attribute, propertyCategoryDictionary) || '未分类';
      if (!acc[key]) {
        acc[key] = { name: key, declarations: 0, exactPasses: 0, ambiguousPasses: 0, rejects: 0, proofRejects: 0 };
      }

      acc[key].declarations += row.declarations;
      acc[key].exactPasses += resolveExactPasses(row);
      acc[key].ambiguousPasses += row.ambiguousPasses;
      acc[key].rejects += row.rejects;
      acc[key].proofRejects += row.proofRejects;
      return acc;
    }, {}),
  )
    .map((item, _, allItems) => {
      const total = allItems.reduce((sum, current) => sum + current.declarations, 0);
      const safeDivide = (numerator: number, denominator: number) => (denominator ? numerator / denominator : 0);
      return {
        ...item,
        value: total ? item.declarations / total : 0,
        proofAccuracy: safeDivide(item.declarations - item.ambiguousPasses - item.proofRejects, item.declarations),
        exactPassRate: safeDivide(item.exactPasses, item.declarations),
      };
    })
    .sort((a, b) => b.value - a.value);

const aggregateSessionShares = (rows: ImportedRow[]) =>
  Object.values(
    rows.reduce<
      Record<
        string,
        {
          name: string;
          declarations: number;
          exactPasses: number;
          ambiguousPasses: number;
          rejects: number;
          proofRejects: number;
        }
      >
    >((acc, row) => {
      const key = row.session || '未识别场次';
      if (!acc[key]) {
        acc[key] = { name: key, declarations: 0, exactPasses: 0, ambiguousPasses: 0, rejects: 0, proofRejects: 0 };
      }

      acc[key].declarations += row.declarations;
      acc[key].exactPasses += resolveExactPasses(row);
      acc[key].ambiguousPasses += row.ambiguousPasses;
      acc[key].rejects += row.rejects;
      acc[key].proofRejects += row.proofRejects;
      return acc;
    }, {}),
  )
    .map((item, _, allItems) => {
      const total = allItems.reduce((sum, current) => sum + current.declarations, 0);
      const safeDivide = (numerator: number, denominator: number) => (denominator ? numerator / denominator : 0);
      return {
        ...item,
        value: total ? item.declarations / total : 0,
        proofAccuracy: safeDivide(item.declarations - item.ambiguousPasses - item.proofRejects, item.declarations),
        exactPassRate: safeDivide(item.exactPasses, item.declarations),
      };
    })
    .sort((a, b) => b.declarations - a.declarations);

const aggregateCompareTrend = (leftRows: ImportedRow[], rightRows: ImportedRow[], leftLabel: string, rightLabel: string) => {
  const safeRate = (numerator: number, denominator: number) => (denominator ? numerator / denominator : 0);

  const buildSeries = (rows: ImportedRow[]) => {
    const grouped = rows.reduce<
      Record<string, { declarations: number; exactPasses: number; ambiguous: number; rejects: number; proofRejects: number }>
    >((acc, row) => {
      if (!acc[row.date]) {
        acc[row.date] = { declarations: 0, exactPasses: 0, ambiguous: 0, rejects: 0, proofRejects: 0 };
      }
      acc[row.date].declarations += row.declarations;
      acc[row.date].exactPasses += resolveExactPasses(row);
      acc[row.date].ambiguous += row.ambiguousPasses;
      acc[row.date].rejects += row.rejects;
      acc[row.date].proofRejects += row.proofRejects;
      return acc;
    }, {});

    return Object.entries(grouped)
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([date, item], index) => ({
        dayIndex: index + 1,
        date,
        declarations: item.declarations,
        proofAccuracy: safeRate(item.declarations - item.ambiguous - item.proofRejects, item.declarations),
        exactPassRate: safeRate(item.exactPasses, item.declarations),
      }));
  };

  const leftSeries = buildSeries(leftRows);
  const rightSeries = buildSeries(rightRows);
  const maxLength = Math.max(leftSeries.length, rightSeries.length);
  const rows = Array.from({ length: maxLength }, (_, index) => {
    const left = leftSeries[index];
    const right = rightSeries[index];

    return {
      dayLabel: `第${index + 1}天`,
      leftProofAccuracy: left?.proofAccuracy ?? null,
      rightProofAccuracy: right?.proofAccuracy ?? null,
      leftExactPassRate: left?.exactPassRate ?? null,
      rightExactPassRate: right?.exactPassRate ?? null,
      leftDate: left?.date ?? '-',
      rightDate: right?.date ?? '-',
    };
  });

  return {
    leftLabel,
    rightLabel,
    rows,
  };
};

const aggregateSessionCompareTrend = (
  leftRows: ImportedRow[],
  rightRows: ImportedRow[],
  leftLabel: string,
  rightLabel: string,
) => {
  const leftTrend = new Map(aggregateTrend(leftRows).map((row) => [row.date, row]));
  const rightTrend = new Map(aggregateTrend(rightRows).map((row) => [row.date, row]));
  const dates = [...new Set([...leftTrend.keys(), ...rightTrend.keys()])].sort((a, b) => a.localeCompare(b));

  return {
    leftLabel,
    rightLabel,
    rows: dates.map((date) => ({
      date,
      leftProofAccuracy: leftTrend.get(date)?.proofAccuracy ?? null,
      rightProofAccuracy: rightTrend.get(date)?.proofAccuracy ?? null,
      leftExactPassRate: leftTrend.get(date)?.exactPassRate ?? null,
      rightExactPassRate: rightTrend.get(date)?.exactPassRate ?? null,
      leftAmbiguousRate: leftTrend.get(date)?.ambiguousRate ?? null,
      rightAmbiguousRate: rightTrend.get(date)?.ambiguousRate ?? null,
      leftRejectRate: leftTrend.get(date)?.rejectRate ?? null,
      rightRejectRate: rightTrend.get(date)?.rejectRate ?? null,
    })),
  };
};

const aggregateDimensionMetrics = (rows: ImportedRow[], key: 'session' | 'batch') => {
  const safeRate = (numerator: number, denominator: number) => (denominator ? numerator / denominator : 0);
  const getBatchOrder = (name: string) => {
    if (name === '其他') return 999;
    const match = name.match(/第(\d+)批/);
    return match ? Number(match[1]) : Number.MAX_SAFE_INTEGER;
  };

  return Object.values(
    rows.reduce<
      Record<string, { name: string; declarations: number; exactPasses: number; ambiguous: number; rejects: number; proofRejects: number }>
    >((acc, row) => {
      const name = row[key];
      if (!acc[name]) {
        acc[name] = { name, declarations: 0, exactPasses: 0, ambiguous: 0, rejects: 0, proofRejects: 0 };
      }

      acc[name].declarations += row.declarations;
      acc[name].exactPasses += resolveExactPasses(row);
      acc[name].ambiguous += row.ambiguousPasses;
      acc[name].rejects += row.rejects;
      acc[name].proofRejects += row.proofRejects;
      return acc;
    }, {}),
  )
    .map((item) => ({
      name: item.name,
      declarations: item.declarations,
      proofAccuracy: safeRate(item.declarations - item.ambiguous - item.proofRejects, item.declarations),
      exactPassRate: safeRate(item.exactPasses, item.declarations),
    }))
    .sort((a, b) => (key === 'batch' ? getBatchOrder(a.name) - getBatchOrder(b.name) : b.declarations - a.declarations))
    .slice(0, 10);
};

const aggregateSessionComparison = (leftRows: ImportedRow[], rightRows: ImportedRow[]) => {
  const sessions = new Map<
    string,
    {
      session: string;
      leftDeclarations: number;
      rightDeclarations: number;
      leftProofAccuracy: number;
      rightProofAccuracy: number;
    }
  >();

  const safeRate = (numerator: number, denominator: number) => (denominator ? numerator / denominator : 0);

  const consume = (rows: ImportedRow[], side: 'left' | 'right') => {
    const grouped = rows.reduce<
      Record<string, { declarations: number; ambiguous: number; proofRejects: number }>
    >((acc, row) => {
      if (!acc[row.session]) {
        acc[row.session] = { declarations: 0, ambiguous: 0, proofRejects: 0 };
      }
      acc[row.session].declarations += row.declarations;
      acc[row.session].ambiguous += row.ambiguousPasses;
      acc[row.session].proofRejects += row.proofRejects;
      return acc;
    }, {});

    Object.entries(grouped).forEach(([session, item]) => {
      if (!sessions.has(session)) {
        sessions.set(session, {
          session,
          leftDeclarations: 0,
          rightDeclarations: 0,
          leftProofAccuracy: 0,
          rightProofAccuracy: 0,
        });
      }

      const current = sessions.get(session)!;
      if (side === 'left') {
        current.leftDeclarations = item.declarations;
        current.leftProofAccuracy = safeRate(item.declarations - item.ambiguous - item.proofRejects, item.declarations);
      } else {
        current.rightDeclarations = item.declarations;
        current.rightProofAccuracy = safeRate(item.declarations - item.ambiguous - item.proofRejects, item.declarations);
      }
    });
  };

  consume(leftRows, 'left');
  consume(rightRows, 'right');

  return [...sessions.values()].sort(
    (a, b) =>
      Math.abs(b.leftProofAccuracy - b.rightProofAccuracy) - Math.abs(a.leftProofAccuracy - a.rightProofAccuracy),
  );
};

const aggregateBatchComparison = (leftRows: ImportedRow[], rightRows: ImportedRow[]) => {
  const batches = new Map<
    string,
    {
      batch: string;
      leftDeclarations: number;
      rightDeclarations: number;
      leftProofAccuracy: number;
      rightProofAccuracy: number;
    }
  >();

  const safeRate = (numerator: number, denominator: number) => (denominator ? numerator / denominator : 0);

  const consume = (rows: ImportedRow[], side: 'left' | 'right') => {
    const grouped = rows.reduce<
      Record<string, { declarations: number; ambiguous: number; proofRejects: number }>
    >((acc, row) => {
      if (!acc[row.batch]) {
        acc[row.batch] = { declarations: 0, ambiguous: 0, proofRejects: 0 };
      }
      acc[row.batch].declarations += row.declarations;
      acc[row.batch].ambiguous += row.ambiguousPasses;
      acc[row.batch].proofRejects += row.proofRejects;
      return acc;
    }, {});

    Object.entries(grouped).forEach(([batch, item]) => {
      if (!batches.has(batch)) {
        batches.set(batch, {
          batch,
          leftDeclarations: 0,
          rightDeclarations: 0,
          leftProofAccuracy: 0,
          rightProofAccuracy: 0,
        });
      }

      const current = batches.get(batch)!;
      if (side === 'left') {
        current.leftDeclarations = item.declarations;
        current.leftProofAccuracy = safeRate(item.declarations - item.ambiguous - item.proofRejects, item.declarations);
      } else {
        current.rightDeclarations = item.declarations;
        current.rightProofAccuracy = safeRate(item.declarations - item.ambiguous - item.proofRejects, item.declarations);
      }
    });
  };

  consume(leftRows, 'left');
  consume(rightRows, 'right');

  return [...batches.values()].sort(
    (a, b) =>
      Math.abs(b.leftProofAccuracy - b.rightProofAccuracy) - Math.abs(a.leftProofAccuracy - a.rightProofAccuracy),
  );
};

const aggregateEfficiency = (rows: EfficiencyRow[]) => {
  const totals = rows.reduce(
    (acc, row) => {
      acc.handledCount += row.handledCount;
      acc.weightedHandledCount += row.weightedHandledCount;
      acc.firstAuditCount += row.firstAuditCount;
      acc.firstAuditPassCount += row.firstAuditPassCount;
      acc.precisionPassCount += row.precisionPassCount;
      acc.auditNotPassCount += row.auditNotPassCount;
      acc.proofRefusalCount += row.proofRefusalCount;
      acc.ambiguousCount += row.ambiguousCount;
      acc.timeoutCount += row.timeoutCount;
      acc.weightedHandleMinutes += row.avgHandleMinutes * row.handledCount;
      return acc;
    },
    {
      handledCount: 0,
      weightedHandledCount: 0,
      firstAuditCount: 0,
      firstAuditPassCount: 0,
      precisionPassCount: 0,
      auditNotPassCount: 0,
      proofRefusalCount: 0,
      ambiguousCount: 0,
      timeoutCount: 0,
      weightedHandleMinutes: 0,
    },
  );

  const employees = new Set(rows.map((row) => row.employee).filter(Boolean));
  const teams = new Set(rows.map((row) => row.team).filter(Boolean));

  return {
    ...totals,
    employeeCount: employees.size,
    teamCount: teams.size,
    avgHandleMinutes: totals.handledCount ? totals.weightedHandleMinutes / totals.handledCount : 0,
    timeoutRate: totals.handledCount ? totals.timeoutCount / totals.handledCount : 0,
    passRate: safeRateFromCounts(totals.firstAuditPassCount, totals.firstAuditCount),
    precisionPassRate: safeRateFromCounts(totals.precisionPassCount, totals.firstAuditCount),
    proofAccuracy: safeRateFromCounts(
      totals.firstAuditCount - totals.ambiguousCount - totals.proofRefusalCount,
      totals.firstAuditCount,
    ),
    ambiguousRate: safeRateFromCounts(totals.ambiguousCount, totals.firstAuditCount),
    proofRefusalRate: safeRateFromCounts(totals.proofRefusalCount, totals.firstAuditCount),
  };
};

const aggregateEfficiencyRanking = (rows: EfficiencyRow[]) =>
  Object.values(
    rows.reduce<
      Record<
        string,
        {
          employee: string;
          team: string;
          handledCount: number;
          weightedHandledCount: number;
          firstAuditCount: number;
          precisionPassCount: number;
          proofRefusalCount: number;
          ambiguousCount: number;
          timeoutCount: number;
          weightedHandleMinutes: number;
        }
      >
    >(
      (acc, row) => {
        if (!acc[row.employee]) {
          acc[row.employee] = {
            employee: row.employee,
            team: row.team,
            handledCount: 0,
            weightedHandledCount: 0,
            firstAuditCount: 0,
            precisionPassCount: 0,
            proofRefusalCount: 0,
            ambiguousCount: 0,
            timeoutCount: 0,
            weightedHandleMinutes: 0,
          };
        }

        acc[row.employee].handledCount += row.handledCount;
        acc[row.employee].weightedHandledCount += row.weightedHandledCount;
        acc[row.employee].firstAuditCount += row.firstAuditCount;
        acc[row.employee].precisionPassCount += row.precisionPassCount;
        acc[row.employee].proofRefusalCount += row.proofRefusalCount;
        acc[row.employee].ambiguousCount += row.ambiguousCount;
        acc[row.employee].timeoutCount += row.timeoutCount;
        acc[row.employee].weightedHandleMinutes += row.avgHandleMinutes * row.handledCount;
        return acc;
      },
      {},
    ),
  )
    .map((item) => ({
      ...item,
      avgHandleMinutes: item.handledCount ? item.weightedHandleMinutes / item.handledCount : 0,
      timeoutRate: item.handledCount ? item.timeoutCount / item.handledCount : 0,
      precisionPassRate: safeRateFromCounts(item.precisionPassCount, item.firstAuditCount),
      proofAccuracy: safeRateFromCounts(
        item.firstAuditCount - item.ambiguousCount - item.proofRefusalCount,
        item.firstAuditCount,
      ),
    }))
    .sort((a, b) => b.weightedHandledCount - a.weightedHandledCount)
    .slice(0, 10);

const aggregateEfficiencyTrend = (rows: EfficiencyRow[]) =>
  Object.values(
    rows.reduce<
      Record<
        string,
        {
          date: string;
          handledCount: number;
          weightedHandledCount: number;
          firstAuditCount: number;
          precisionPassCount: number;
          proofRefusalCount: number;
          ambiguousCount: number;
          timeoutCount: number;
        }
      >
    >((acc, row) => {
      if (!acc[row.date]) {
        acc[row.date] = {
          date: row.date,
          handledCount: 0,
          weightedHandledCount: 0,
          firstAuditCount: 0,
          precisionPassCount: 0,
          proofRefusalCount: 0,
          ambiguousCount: 0,
          timeoutCount: 0,
        };
      }

      acc[row.date].handledCount += row.handledCount;
      acc[row.date].weightedHandledCount += row.weightedHandledCount;
      acc[row.date].firstAuditCount += row.firstAuditCount;
      acc[row.date].precisionPassCount += row.precisionPassCount;
      acc[row.date].proofRefusalCount += row.proofRefusalCount;
      acc[row.date].ambiguousCount += row.ambiguousCount;
      acc[row.date].timeoutCount += row.timeoutCount;
      return acc;
    }, {}),
  )
    .map((item) => ({
      ...item,
      precisionPassRate: safeRateFromCounts(item.precisionPassCount, item.firstAuditCount),
      proofAccuracy: safeRateFromCounts(
        item.firstAuditCount - item.ambiguousCount - item.proofRefusalCount,
        item.firstAuditCount,
      ),
    }))
    .sort((a, b) => a.date.localeCompare(b.date));

const aggregateWorkplaceEfficiency = (rows: EfficiencyRow[]) =>
  Object.values(
    rows.reduce<
      Record<
        string,
        {
          workplace: string;
          teams: Set<string>;
          employees: Set<string>;
          handledCount: number;
          weightedHandledCount: number;
          firstAuditCount: number;
          precisionPassCount: number;
          proofRefusalCount: number;
          ambiguousCount: number;
          dailyStats: Map<string, { weightedHandledCount: number; employees: Set<string> }>;
        }
      >
    >((acc, row) => {
      const workplace = resolveWorkplace(row.team);
      if (!acc[workplace]) {
        acc[workplace] = {
          workplace,
          teams: new Set<string>(),
          employees: new Set<string>(),
          handledCount: 0,
          weightedHandledCount: 0,
          firstAuditCount: 0,
          precisionPassCount: 0,
          proofRefusalCount: 0,
          ambiguousCount: 0,
          dailyStats: new Map<string, { weightedHandledCount: number; employees: Set<string> }>(),
        };
      }

      acc[workplace].teams.add(row.team);
      acc[workplace].employees.add(row.employee);
      acc[workplace].handledCount += row.handledCount;
      acc[workplace].weightedHandledCount += row.weightedHandledCount;
      acc[workplace].firstAuditCount += row.firstAuditCount;
      acc[workplace].precisionPassCount += row.precisionPassCount;
      acc[workplace].proofRefusalCount += row.proofRefusalCount;
      acc[workplace].ambiguousCount += row.ambiguousCount;
      const dailyItem = acc[workplace].dailyStats.get(row.date) ?? {
        weightedHandledCount: 0,
        employees: new Set<string>(),
      };
      dailyItem.weightedHandledCount += row.weightedHandledCount;
      if (row.employee) {
        dailyItem.employees.add(row.employee);
      }
      acc[workplace].dailyStats.set(row.date, dailyItem);
      return acc;
    }, {}),
  )
    .map((item) => {
      const dailyWeightedPerEmployee = [...item.dailyStats.values()]
        .map((dailyItem) => safeRateFromCounts(dailyItem.weightedHandledCount, dailyItem.employees.size))
        .filter((value) => value > 0);

      return {
        ...item,
        teamCount: item.teams.size,
        employeeCount: item.employees.size,
        activeDayCount: item.dailyStats.size,
        avgDailyWeightedHandledPerEmployee: safeRateFromCounts(
          dailyWeightedPerEmployee.reduce((sum, value) => sum + value, 0),
          dailyWeightedPerEmployee.length,
        ),
        precisionPassRate: safeRateFromCounts(item.precisionPassCount, item.firstAuditCount),
        proofAccuracy: safeRateFromCounts(
          item.firstAuditCount - item.ambiguousCount - item.proofRefusalCount,
          item.firstAuditCount,
        ),
        teams: [...item.teams].filter(Boolean).sort((a, b) => a.localeCompare(b, 'zh-CN')),
      };
    })
    .sort((a, b) => {
      const order = ['常州', '上海', '其他'];
      return order.indexOf(a.workplace) - order.indexOf(b.workplace);
    });

type EfficiencyTimeDimension = 'day' | 'week' | 'month';

const getEfficiencyPeriodLabel = (date: string, dimension: EfficiencyTimeDimension) => {
  if (dimension === 'month') {
    return date.slice(0, 7);
  }

  if (dimension === 'week') {
    const weekStartDate = startOfWeek(parseISO(date), { weekStartsOn: 0 });
    const start = format(weekStartDate, 'yyyy-MM-dd');
    const end = format(addDays(weekStartDate, 6), 'yyyy-MM-dd');
    return `${start} ~ ${end}`;
  }

  return date;
};

const aggregateEmployeeEfficiencyDetail = (
  rows: EfficiencyRow[],
  employee: string,
  dimension: EfficiencyTimeDimension,
) => {
  const employeeRows = rows.filter((row) => row.employee === employee);
  const activeDates = new Set(employeeRows.map((row) => row.date).filter(Boolean));
  const totals = aggregateEfficiency(employeeRows);
  const activeDayCount = activeDates.size;
  const trendRows = Object.values(
    employeeRows.reduce<
      Record<
        string,
        {
          period: string;
          handledCount: number;
          weightedHandledCount: number;
          firstAuditCount: number;
          precisionPassCount: number;
          proofRefusalCount: number;
          ambiguousCount: number;
          activeDates: Set<string>;
        }
      >
    >((acc, row) => {
      const period = getEfficiencyPeriodLabel(row.date, dimension);
      if (!acc[period]) {
        acc[period] = {
          period,
          handledCount: 0,
          weightedHandledCount: 0,
          firstAuditCount: 0,
          precisionPassCount: 0,
          proofRefusalCount: 0,
          ambiguousCount: 0,
          activeDates: new Set<string>(),
        };
      }

      acc[period].handledCount += row.handledCount;
      acc[period].weightedHandledCount += row.weightedHandledCount;
      acc[period].firstAuditCount += row.firstAuditCount;
      acc[period].precisionPassCount += row.precisionPassCount;
      acc[period].proofRefusalCount += row.proofRefusalCount;
      acc[period].ambiguousCount += row.ambiguousCount;
      acc[period].activeDates.add(row.date);
      return acc;
    }, {}),
  )
    .map((item) => ({
      ...item,
      activeDayCount: item.activeDates.size,
      dailyHandledAverage: item.activeDates.size ? item.handledCount / item.activeDates.size : 0,
      dailyWeightedAverage: item.activeDates.size ? item.weightedHandledCount / item.activeDates.size : 0,
      precisionPassRate: safeRateFromCounts(item.precisionPassCount, item.firstAuditCount),
      proofAccuracy: safeRateFromCounts(
        item.firstAuditCount - item.ambiguousCount - item.proofRefusalCount,
        item.firstAuditCount,
      ),
    }))
    .sort((a, b) => a.period.localeCompare(b.period));

  return {
    rows: employeeRows,
    trendRows,
    totals,
    activeDayCount,
    dailyHandledAverage: activeDayCount ? totals.handledCount / activeDayCount : 0,
    dailyWeightedAverage: activeDayCount ? totals.weightedHandledCount / activeDayCount : 0,
  };
};

const aggregateEmployeeEfficiencyAnomalies = (rows: EfficiencyRow[]) => {
  const dates = [...new Set(rows.map((row) => row.date).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  if (dates.length < 2) {
    return { efficiencyChanges: [], combinedAnomalies: [] };
  }

  const midpoint = Math.ceil(dates.length / 2);
  const earlyDates = new Set(dates.slice(0, midpoint));
  const lateDates = new Set(dates.slice(midpoint));

  if (!earlyDates.size || !lateDates.size) {
    return { efficiencyChanges: [], combinedAnomalies: [] };
  }

  const grouped = rows.reduce<Record<string, EfficiencyRow[]>>((acc, row) => {
    if (!acc[row.employee]) {
      acc[row.employee] = [];
    }

    acc[row.employee].push(row);
    return acc;
  }, {});

  const summarize = (employeeRows: EfficiencyRow[], targetDates: Set<string>) => {
    const periodRows = employeeRows.filter((row) => targetDates.has(row.date));
    const activeDays = new Set(periodRows.map((row) => row.date).filter(Boolean)).size;
    const metrics = aggregateEfficiency(periodRows);

    return {
      activeDays,
      handledCount: metrics.handledCount,
      weightedHandledCount: metrics.weightedHandledCount,
      dailyWeightedAverage: activeDays ? metrics.weightedHandledCount / activeDays : 0,
      precisionPassRate: metrics.precisionPassRate,
      proofAccuracy: metrics.proofAccuracy,
    };
  };

  const anomalies = Object.entries(grouped)
    .map(([employee, employeeRows]) => {
      const early = summarize(employeeRows, earlyDates);
      const late = summarize(employeeRows, lateDates);
      const totalHandledCount = early.handledCount + late.handledCount;
      const efficiencyDelta = late.dailyWeightedAverage - early.dailyWeightedAverage;
      const efficiencyDeltaRate = early.dailyWeightedAverage ? efficiencyDelta / early.dailyWeightedAverage : 0;
      const precisionPassRateDelta = late.precisionPassRate - early.precisionPassRate;
      const proofAccuracyDelta = late.proofAccuracy - early.proofAccuracy;

      return {
        employee,
        team: employeeRows[0]?.team ?? '',
        totalHandledCount,
        early,
        late,
        efficiencyDelta,
        efficiencyDeltaRate,
        precisionPassRateDelta,
        proofAccuracyDelta,
        combinedScore: Math.abs(efficiencyDeltaRate) + Math.abs(precisionPassRateDelta) * 2,
      };
    })
    .filter((item) => item.totalHandledCount >= 20 && item.early.activeDays > 0 && item.late.activeDays > 0);

  const efficiencyChanges = [...anomalies]
    .filter((item) => Math.abs(item.efficiencyDeltaRate) >= 0.2 || Math.abs(item.efficiencyDelta) >= 10)
    .sort((a, b) => Math.abs(b.efficiencyDeltaRate) - Math.abs(a.efficiencyDeltaRate))
    .slice(0, 5);

  const combinedAnomalies = [...anomalies]
    .filter((item) => Math.abs(item.precisionPassRateDelta) >= 0.05 && (Math.abs(item.efficiencyDeltaRate) >= 0.15 || Math.abs(item.efficiencyDelta) >= 8))
    .sort((a, b) => b.combinedScore - a.combinedScore)
    .slice(0, 5);

  return { efficiencyChanges, combinedAnomalies };
};

const aggregateQualityDimension = (
  rows: ImportedRow[],
  key: 'session' | 'batch' | 'category' | 'attribute',
  propertyCategoryDictionary: PropertyCategoryEntry[] = [],
) =>
  Object.values(
    rows.reduce<Record<string, { name: string; declarations: number; exactPasses: number; ambiguousPasses: number; rejects: number; proofRejects: number }>>(
      (acc, row) => {
        const name =
          key === 'category'
            ? resolveCategory(row.category, row.attribute, propertyCategoryDictionary)
            : String(row[key] || '未分类');
        if (!acc[name]) {
          acc[name] = { name, declarations: 0, exactPasses: 0, ambiguousPasses: 0, rejects: 0, proofRejects: 0 };
        }

        acc[name].declarations += row.declarations;
        acc[name].exactPasses += resolveExactPasses(row);
        acc[name].ambiguousPasses += row.ambiguousPasses;
        acc[name].rejects += row.rejects;
        acc[name].proofRejects += row.proofRejects;
        return acc;
      },
      {},
    ),
  )
    .map((item) => ({
      ...item,
      metrics: aggregateMetrics([
        {
          date: '',
          session: '',
          batch: '',
          category: '',
          attribute: '',
          declarations: item.declarations,
          exactPasses: item.exactPasses,
          ambiguousPasses: item.ambiguousPasses,
          rejects: item.rejects,
          proofRejects: item.proofRejects,
        },
      ]),
    }))
    .sort((a, b) => b.declarations - a.declarations);

const splitTrendMetrics = (rows: ImportedRow[]) => {
  const dates = [...new Set(rows.map((row) => row.date))].sort((a, b) => a.localeCompare(b));
  const midpoint = Math.ceil(dates.length / 2);
  const earlyDates = new Set(dates.slice(0, midpoint));
  const lateDates = new Set(dates.slice(midpoint));
  const earlyMetrics = aggregateMetrics(rows.filter((row) => earlyDates.has(row.date)));
  const lateMetrics = aggregateMetrics(rows.filter((row) => lateDates.has(row.date)));

  return {
    dates,
    earlyMetrics,
    lateMetrics,
    proofAccuracyDelta: lateMetrics.proofAccuracy - earlyMetrics.proofAccuracy,
    exactPassRateDelta: lateMetrics.exactPassRate - earlyMetrics.exactPassRate,
  };
};

const aggregateAttributePrecisionVolatility = (rows: ImportedRow[]) => {
  const grouped = rows.reduce<Record<string, ImportedRow[]>>((acc, row) => {
    if (!acc[row.attribute]) {
      acc[row.attribute] = [];
    }

    acc[row.attribute].push(row);
    return acc;
  }, {});

  return Object.entries(grouped)
    .map(([attribute, attributeRows]) => {
      const trendRows = aggregateTrend(attributeRows);
      const rates = trendRows.filter((item) => item.declarations > 0).map((item) => item.exactPassRate);
      const totalDeclarations = attributeRows.reduce((sum, row) => sum + row.declarations, 0);
      const minRate = rates.length ? Math.min(...rates) : 0;
      const maxRate = rates.length ? Math.max(...rates) : 0;

      return {
        attribute,
        declarations: totalDeclarations,
        activeDays: rates.length,
        minRate,
        maxRate,
        volatility: maxRate - minRate,
        overallRate: aggregateMetrics(attributeRows).exactPassRate,
      };
    })
    .filter((item) => item.declarations >= 20 && item.activeDays >= 2)
    .sort((a, b) => b.volatility - a.volatility)
    .slice(0, 5);
};

const aggregateSessionAttributePassChanges = (rows: ImportedRow[]) => {
  const totalDeclarations = rows.reduce((sum, row) => sum + row.declarations, 0);
  const sessionDeclarations = rows.reduce<Record<string, number>>((acc, row) => {
    const session = row.session || '未识别场次';
    acc[session] = (acc[session] ?? 0) + row.declarations;
    return acc;
  }, {});
  const majorSessions = new Set(
    Object.entries(sessionDeclarations)
      .filter(([, declarations]) => totalDeclarations > 0 && declarations / totalDeclarations > 0.01)
      .map(([session]) => session),
  );
  const grouped = rows
    .filter((row) => majorSessions.has(row.session || '未识别场次'))
    .reduce<Record<string, ImportedRow[]>>((acc, row) => {
      const key = `${row.session || '未识别场次'}|||${row.attribute}`;
      if (!acc[key]) {
        acc[key] = [];
      }

      acc[key].push(row);
      return acc;
    }, {});

  return Object.entries(grouped)
    .map(([key, groupRows]) => {
      const [session, attribute] = key.split('|||');
      const trendRows = aggregateTrend(groupRows);
      const rates = trendRows.filter((item) => item.declarations > 0).map((item) => item.exactPassRate);
      const declarations = groupRows.reduce((sum, row) => sum + row.declarations, 0);
      const minRate = rates.length ? Math.min(...rates) : 0;
      const maxRate = rates.length ? Math.max(...rates) : 0;

      return {
        session,
        attribute,
        declarations,
        activeDays: rates.length,
        sessionShare: totalDeclarations ? (sessionDeclarations[session] ?? 0) / totalDeclarations : 0,
        minRate,
        maxRate,
        volatility: maxRate - minRate,
        overallRate: aggregateMetrics(groupRows).exactPassRate,
      };
    })
    .filter((item) => item.declarations >= 20 && item.activeDays >= 2)
    .sort((a, b) => b.volatility - a.volatility)
    .slice(0, 8);
};

const calculateHealthScore = (
  qualityMetrics: ReturnType<typeof aggregateMetrics>,
  trendSplit: ReturnType<typeof splitTrendMetrics>,
) => {
  if (!qualityMetrics.declarations) {
    return {
      score: 0,
      level: '暂无数据',
      factors: ['当前筛选范围内没有申报数据，暂不计算健康度。'],
    };
  }

  const proofRejectRate = qualityMetrics.declarations
    ? qualityMetrics.proofRejects / qualityMetrics.declarations
    : 0;
  const exactMissRate = 1 - qualityMetrics.exactPassRate;
  const trendPenalty = Math.max(0, -trendSplit.proofAccuracyDelta) * 100;
  const samplePenalty = qualityMetrics.declarations < 100 ? 6 : qualityMetrics.declarations < 500 ? 3 : 0;
  const ambiguousOverTarget = Math.max(0, qualityMetrics.ambiguousRate - AMBIGUOUS_RATE_TARGET);
  const ambiguousPenalty = ambiguousOverTarget * 80;
  const rejectPenalty = qualityMetrics.rejectRate * 45;
  const proofRejectPenalty = proofRejectRate * 35;
  const exactMissPenalty = exactMissRate * 12;
  const totalPenalty =
    ambiguousPenalty + rejectPenalty + proofRejectPenalty + exactMissPenalty + trendPenalty + samplePenalty;
  const score = Math.max(0, Math.min(100, Math.round(100 - totalPenalty)));
  const level = score >= 85 ? '健康' : score >= 70 ? '需关注' : score >= 55 ? '偏弱' : '高风险';
  const factors = [
    `模棱两可率扣分 ${ambiguousPenalty.toFixed(1)}：目标 ${formatPercent(AMBIGUOUS_RATE_TARGET)}，当前 ${formatPercent(qualityMetrics.ambiguousRate)}。`,
    `拒绝率扣分 ${rejectPenalty.toFixed(1)}：当前 ${formatPercent(qualityMetrics.rejectRate)}。`,
    `举证未通过扣分 ${proofRejectPenalty.toFixed(1)}：当前 ${formatPercent(proofRejectRate)}。`,
    `精准未通过扣分 ${exactMissPenalty.toFixed(1)}：精准通过率 ${formatPercent(qualityMetrics.exactPassRate)}。`,
    trendPenalty
      ? `趋势扣分 ${trendPenalty.toFixed(1)}：后半段举证准确率较前半段下降 ${formatPercent(-trendSplit.proofAccuracyDelta)}。`
      : '趋势扣分 0.0：后半段举证准确率未低于前半段。',
    samplePenalty ? `样本扣分 ${samplePenalty.toFixed(1)}：当前申报样本量偏小。` : '样本扣分 0.0：当前样本量充足。',
  ];

  return { score, level, factors };
};

const createAiAnalysis = ({
  qualityMetrics,
  qualityRows,
  topAttributes,
  categoryData,
  efficiencyMetrics,
  efficiencyRanking,
  efficiencyRows,
  propertyCategoryDictionary,
}: {
  qualityMetrics: ReturnType<typeof aggregateMetrics>;
  qualityRows: ImportedRow[];
  topAttributes: ReturnType<typeof aggregateAttributes>;
  categoryData: ReturnType<typeof aggregateCategories>;
  efficiencyMetrics: ReturnType<typeof aggregateEfficiency>;
  efficiencyRanking: ReturnType<typeof aggregateEfficiencyRanking>;
  efficiencyRows: EfficiencyRow[];
  propertyCategoryDictionary: PropertyCategoryEntry[];
}) => {
  const topAttribute = topAttributes[0];
  const topCategory = categoryData[0];
  const hasQualityData = qualityMetrics.declarations > 0;
  const hasEfficiencyData = efficiencyMetrics.handledCount > 0;
  const sessionDrivers = aggregateQualityDimension(qualityRows, 'session');
  const batchDrivers = aggregateQualityDimension(qualityRows, 'batch');
  const categoryDrivers = aggregateQualityDimension(qualityRows, 'category', propertyCategoryDictionary);
  const attributeDrivers = aggregateQualityDimension(qualityRows, 'attribute');
  const weakSessions = sessionDrivers
    .filter((item) => item.declarations >= 20)
    .sort((a, b) => a.metrics.proofAccuracy - b.metrics.proofAccuracy)
    .slice(0, 3);
  const weakBatches = batchDrivers
    .filter((item) => item.declarations >= 20)
    .sort((a, b) => a.metrics.proofAccuracy - b.metrics.proofAccuracy)
    .slice(0, 3);
  const riskyAttributes = attributeDrivers
    .filter((item) => item.declarations >= 10)
    .sort((a, b) => b.metrics.rejectRate - a.metrics.rejectRate)
    .slice(0, 5);
  const highRejectAttributes = attributeDrivers
    .filter((item) => item.declarations >= 20 && item.metrics.rejectRate >= qualityMetrics.rejectRate)
    .sort((a, b) => b.metrics.rejectRate - a.metrics.rejectRate)
    .slice(0, 5);
  const highAmbiguousAttributes = attributeDrivers
    .filter((item) => item.declarations >= 20 && item.metrics.ambiguousRate >= Math.max(AMBIGUOUS_RATE_TARGET, qualityMetrics.ambiguousRate))
    .sort((a, b) => b.metrics.ambiguousRate - a.metrics.ambiguousRate)
    .slice(0, 5);
  const precisionVolatileAttributes = aggregateAttributePrecisionVolatility(qualityRows);
  const sessionAttributePassChanges = aggregateSessionAttributePassChanges(qualityRows);
  const employeeAnomalies = aggregateEmployeeEfficiencyAnomalies(efficiencyRows);
  const ambiguousCategories = categoryDrivers
    .filter((item) => item.declarations >= 10)
    .sort((a, b) => b.metrics.ambiguousRate - a.metrics.ambiguousRate)
    .slice(0, 3);
  const trendSplit = splitTrendMetrics(qualityRows);
  const health = calculateHealthScore(qualityMetrics, trendSplit);
  const healthScore = health.score;
  const healthLevel = health.level;
  const qualityDateRange = qualityRows.length
    ? `${qualityRows[0].date} ~ ${qualityRows[qualityRows.length - 1].date}`
    : '暂无质量数据';

  const summary = [
    hasQualityData
      ? `质量健康评分 ${healthScore}/100，状态为「${healthLevel}」。当前共有 ${formatInteger(qualityMetrics.declarations)} 次申报，举证准确率 ${formatPercent(qualityMetrics.proofAccuracy)}，精准通过率 ${formatPercent(qualityMetrics.exactPassRate)}。`
      : '当前还没有可分析的质量数据，请先在数据导入模块导入质量周数据。',
    hasQualityData && trendSplit.dates.length > 1
      ? `从前半段到后半段看，举证准确率变化 ${formatPercent(trendSplit.proofAccuracyDelta)}，精准通过率变化 ${formatPercent(trendSplit.exactPassRateDelta)}。`
      : '当前日期样本不足，暂不判断周内趋势变化。',
    weakSessions.length
      ? `场次对比显示，${weakSessions.map((item) => `「${item.name}」举证准确率${formatPercent(item.metrics.proofAccuracy)}`).join('、')} 更需要优先复盘。`
      : '场次之间暂未出现明显质量断层，建议继续保持常规监控。',
    sessionAttributePassChanges.length
      ? `在申报占比超过 1% 的场次内，精准通过率变化较大的场次-属性项组合包括：${sessionAttributePassChanges.slice(0, 3).map((item) => `「${item.session}-${item.attribute}」波动${formatPercent(item.volatility)}`).join('；')}。`
      : '申报占比超过 1% 的场次内，暂未识别到高样本属性项通过率明显波动。',
    hasEfficiencyData
      ? `人效侧已导入 ${formatInteger(efficiencyMetrics.handledCount)} 单、覆盖 ${formatInteger(efficiencyMetrics.employeeCount)} 人；人员维度将重点看审核效率变化，以及精准通过率与效率同步异动。`
      : '当前还没有可分析的人效数据；AI 分析会先基于质量数据输出结论。',
  ];

  const risks = [
    weakSessions.length
      ? `场次拖累项：${weakSessions.map((item) => `「${item.name}」${formatPercent(item.metrics.proofAccuracy)} / ${formatInteger(item.declarations)}次`).join('；')}。`
      : '场次维度暂未发现明显拖累项，或样本量不足。',
    weakBatches.length
      ? `批次拖累项：${weakBatches.map((item) => `「${item.name}」${formatPercent(item.metrics.proofAccuracy)} / ${formatInteger(item.declarations)}次`).join('；')}。`
      : '批次维度暂未发现明显拖累项，或样本量不足。',
    qualityMetrics.ambiguousRate > AMBIGUOUS_RATE_TARGET
      ? `模棱两可率达到 ${formatPercent(qualityMetrics.ambiguousRate)}，高于 ${formatPercent(AMBIGUOUS_RATE_TARGET)} 目标，建议优先复盘模糊通过较集中的属性项。`
      : `模棱两可率为 ${formatPercent(qualityMetrics.ambiguousRate)}，低于 ${formatPercent(AMBIGUOUS_RATE_TARGET)} 目标，当前未触发高模糊风险。`,
    qualityMetrics.rejectRate > 0.12
      ? `拒绝率达到 ${formatPercent(qualityMetrics.rejectRate)}，需要关注未通过集中场次和批次。`
      : `拒绝率为 ${formatPercent(qualityMetrics.rejectRate)}，整体拒绝压力可控。`,
    riskyAttributes.length
      ? `拒绝风险属性项：${riskyAttributes.map((item) => `「${item.name}」拒绝率${formatPercent(item.metrics.rejectRate)}`).join('；')}。`
      : '属性项维度暂未发现高拒绝风险，或样本量不足。',
    highAmbiguousAttributes.length
      ? `高模棱两可属性项：${highAmbiguousAttributes.map((item) => `「${item.name}」模棱两可率${formatPercent(item.metrics.ambiguousRate)}、申报${formatInteger(item.declarations)}次`).join('；')}。`
      : '属性项维度暂未发现高样本高模棱两可风险。',
    precisionVolatileAttributes.length
      ? `精准通过率波动属性项：${precisionVolatileAttributes.map((item) => `「${item.attribute}」波动${formatPercent(item.volatility)}（${formatPercent(item.minRate)}~${formatPercent(item.maxRate)}）`).join('；')}。`
      : '属性项精准通过率暂未发现明显波动，或日期样本不足。',
    employeeAnomalies.efficiencyChanges.length
      ? `人员效率变化明显：${employeeAnomalies.efficiencyChanges.map((item) => `「${item.employee}」日均加权审核量${item.efficiencyDelta >= 0 ? '上升' : '下降'}${Math.abs(item.efficiencyDelta).toFixed(1)}（${formatPercent(Math.abs(item.efficiencyDeltaRate))}）`).join('；')}。`
      : '人员维度暂未识别明显效率变化，或人效样本不足。',
    employeeAnomalies.combinedAnomalies.length
      ? `人员通过率与效率双异动：${employeeAnomalies.combinedAnomalies.map((item) => `「${item.employee}」效率${item.efficiencyDelta >= 0 ? '上升' : '下降'}${Math.abs(item.efficiencyDelta).toFixed(1)}、精准通过率变化${formatPercent(item.precisionPassRateDelta)}`).join('；')}。`
      : '暂未识别通过率与效率同时明显异动的人员。',
    hasEfficiencyData
      ? '人效底表已导入，但本轮 AI 风险判断不纳入平均处理时长和超时率。'
      : '人效底表未导入，本轮 AI 风险判断仅基于质量数据。',
  ];

  const actions = [
    topAttribute
      ? `优先复盘高频属性项「${topAttribute.attribute}」，当前申报 ${formatInteger(topAttribute.declarations)} 次。`
      : '先补充质量底表，形成属性项维度的稳定样本。',
    topCategory
      ? `关注属性项分类「${topCategory.name}」，其占当前分类分布 ${formatPercent(topCategory.value)}。`
      : '属性项分类样本不足，暂不做分类归因。',
    sessionAttributePassChanges.length
      ? `场次内通过率变化专项：优先抽看${sessionAttributePassChanges.slice(0, 3).map((item) => `「${item.session}-${item.attribute}」`).join('、')}，对比高低日期样本，确认是场次结构变化、口径漂移还是人员判断差异。`
      : '主要场次内未识别明显通过率波动组合，暂不建立场次-属性项专项。',
    highAmbiguousAttributes.length
      ? `模糊口径复盘建议优先看：${highAmbiguousAttributes.map((item) => `「${item.name}」${formatPercent(item.metrics.ambiguousRate)}`).join('、')}。`
      : ambiguousCategories.length
        ? `模糊口径复盘可先看分类：${ambiguousCategories.map((item) => `「${item.name}」${formatPercent(item.metrics.ambiguousRate)}`).join('、')}。`
        : '模糊通过暂无明显属性项或分类集中，可保持常规抽检。',
    highRejectAttributes.length
      ? `拒绝率专项建议聚焦：${highRejectAttributes.map((item) => `「${item.name}」拒绝率${formatPercent(item.metrics.rejectRate)}、申报${formatInteger(item.declarations)}次`).join('；')}，逐项核对拒绝原因是否来自拍摄证据、规则口径或商品描述。`
      : '拒绝率未呈现高样本属性项集中，建议保持按场次抽检。',
    precisionVolatileAttributes.length
      ? `精准通过率波动建议优先抽看：${precisionVolatileAttributes.slice(0, 3).map((item) => `「${item.attribute}」`).join('、')}，对比高低日期样本，确认是否存在口径漂移或人员判断差异。`
      : '精准通过率波动不明显，暂不需要单独建立波动专项。',
    employeeAnomalies.combinedAnomalies.length
      ? `人员异动复盘建议优先看：${employeeAnomalies.combinedAnomalies.slice(0, 3).map((item) => `「${item.employee}」`).join('、')}，逐人核对其审核场次结构、属性项结构和抽检样本，判断是否属于效率提升带来的质量波动。`
      : employeeAnomalies.efficiencyChanges.length
        ? `人员效率变化建议关注：${employeeAnomalies.efficiencyChanges.slice(0, 3).map((item) => `「${item.employee}」`).join('、')}，确认效率变化是否来自排班、场次结构或审核难度变化。`
        : '人员侧暂未触发效率/通过率异动专项。',
    weakSessions[0] || weakBatches[0]
      ? `建议本周复盘优先级：先看${weakSessions[0] ? `场次「${weakSessions[0].name}」` : ''}${weakSessions[0] && weakBatches[0] ? '，再看' : ''}${weakBatches[0] ? `批次「${weakBatches[0].name}」` : ''}。`
      : '建议本周以高频属性项和分类抽检为主，暂不需要大规模专项复盘。',
  ];

  const drivers = [
    topAttribute
      ? `最大样本属性项：「${topAttribute.attribute}」，申报 ${formatInteger(topAttribute.declarations)} 次，适合作为口径校准样本。`
      : '暂无最大样本属性项。',
    topCategory
      ? `最大分类占比：「${topCategory.name}」，占比 ${formatPercent(topCategory.value)}，对整体指标解释权重较高。`
      : '暂无分类占比信息。',
    weakSessions[0]
      ? `场次对比核心差异：「${weakSessions[0].name}」举证准确率 ${formatPercent(weakSessions[0].metrics.proofAccuracy)}，申报 ${formatInteger(weakSessions[0].declarations)} 次。`
      : '暂未识别明显弱势场次。',
    sessionAttributePassChanges[0]
      ? `主要场次内通过率变化最大：「${sessionAttributePassChanges[0].session}-${sessionAttributePassChanges[0].attribute}」，精准通过率波动 ${formatPercent(sessionAttributePassChanges[0].volatility)}，场次占比 ${formatPercent(sessionAttributePassChanges[0].sessionShare)}。`
      : '暂未识别主要场次内通过率明显波动的属性项。',
    highRejectAttributes[0]
      ? `拒绝率最高属性项：「${highRejectAttributes[0].name}」，拒绝率 ${formatPercent(highRejectAttributes[0].metrics.rejectRate)}，申报 ${formatInteger(highRejectAttributes[0].declarations)} 次。`
      : '暂未识别高样本高拒绝属性项。',
    highAmbiguousAttributes[0]
      ? `模棱两可率最高属性项：「${highAmbiguousAttributes[0].name}」，模棱两可率 ${formatPercent(highAmbiguousAttributes[0].metrics.ambiguousRate)}，申报 ${formatInteger(highAmbiguousAttributes[0].declarations)} 次。`
      : '暂未识别高样本高模棱两可属性项。',
    employeeAnomalies.combinedAnomalies[0]
      ? `人员双异动最高：「${employeeAnomalies.combinedAnomalies[0].employee}」，日均加权审核量变化 ${employeeAnomalies.combinedAnomalies[0].efficiencyDelta.toFixed(1)}，精准通过率变化 ${formatPercent(employeeAnomalies.combinedAnomalies[0].precisionPassRateDelta)}。`
      : employeeAnomalies.efficiencyChanges[0]
        ? `人员效率变化最大：「${employeeAnomalies.efficiencyChanges[0].employee}」，日均加权审核量变化 ${employeeAnomalies.efficiencyChanges[0].efficiencyDelta.toFixed(1)}。`
        : '暂未识别人员效率或通过率明显异动。',
    precisionVolatileAttributes[0]
      ? `精准通过率波动最大属性项：「${precisionVolatileAttributes[0].attribute}」，波动 ${formatPercent(precisionVolatileAttributes[0].volatility)}。`
      : '暂未识别精准通过率明显波动属性项。',
  ];

  const report = [
    `【预质检质量看板 AI 分析】`,
    `分析范围：${qualityDateRange}`,
    `健康评分：${healthScore}/100（${healthLevel}）`,
    '',
    '一、核心结论',
    ...summary.map((item, index) => `${index + 1}. ${item}`),
    '',
    '二、关键驱动',
    ...drivers.map((item, index) => `${index + 1}. ${item}`),
    '',
    '三、健康度扣分拆解',
    ...health.factors.map((item, index) => `${index + 1}. ${item}`),
    '',
    '四、风险提醒',
    ...risks.map((item, index) => `${index + 1}. ${item}`),
    '',
    '五、建议动作',
    ...actions.map((item, index) => `${index + 1}. ${item}`),
  ].join('\n');

  return { summary, risks, actions, drivers, report, healthScore, healthLevel, healthFactors: health.factors };
};

const downloadTemplate = () => {
  const sample = [
    {
      第一次线审完成时间: '2026-05-23',
      场次: '京东寄卖',
      批次: '第4批',
      属性标签: '屏幕外观',
      申报次数: 2,
      模糊通过次数: 0,
      未通过次数: 1,
      举证未通过次数: 0,
    },
  ];

  const worksheet = XLSX.utils.json_to_sheet(sample);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, '模板');
  XLSX.writeFile(workbook, '周数据导入模板.xlsx');
};

const downloadEfficiencyTemplate = () => {
  const sample = [
    {
      日期: '2026-05-31',
      员工姓名: '张三',
      团队: '预质检一组',
      总审核量: 120,
      加权审核量: 126.4,
      一审审核量: 100,
      一审通过量: 82,
      通过率: 0.82,
      精准通过量: 76,
      精准通过率: 0.76,
      未通过量: 18,
      举证拒绝量: 3,
      模糊通过量: 6,
      举证准确率: 0.91,
    },
  ];

  const worksheet = XLSX.utils.json_to_sheet(sample);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, '人效模板');
  XLSX.writeFile(workbook, '人效周数据导入模板.xlsx');
};

const fetchSharedDataset = async (): Promise<SharedDatasetResponse> => {
  const response = await fetch('/api/dataset');
  if (!response.ok) {
    throw new Error('读取共享数据失败');
  }
  return response.json();
};

const mergeSharedDataset = async (payload: ParsedWorkbook): Promise<SharedDatasetResponse> => {
  const response = await fetch('/api/dataset/merge', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    throw new Error('写入共享数据失败');
  }

  return response.json();
};

const clearSharedDataset = async (): Promise<SharedDatasetResponse> => {
  const response = await fetch('/api/dataset', { method: 'DELETE' });
  if (!response.ok) {
    throw new Error('清空共享数据失败');
  }
  return response.json();
};

const fetchEfficiencyDataset = async (): Promise<EfficiencyDatasetResponse> => {
  const response = await fetch('/api/efficiency-dataset');
  if (!response.ok) {
    throw new Error('读取人效数据失败');
  }
  return response.json();
};

const mergeEfficiencyDataset = async (payload: ParsedEfficiencyWorkbook): Promise<EfficiencyDatasetResponse> => {
  const response = await fetch('/api/efficiency-dataset/merge', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    throw new Error('写入人效数据失败');
  }

  return response.json();
};

const clearEfficiencyDataset = async (): Promise<EfficiencyDatasetResponse> => {
  const response = await fetch('/api/efficiency-dataset', { method: 'DELETE' });
  if (!response.ok) {
    throw new Error('清空人效数据失败');
  }
  return response.json();
};

const fetchPropertyCategoryDictionary = async (): Promise<PropertyCategoryDictionaryResponse> => {
  const response = await fetch('/api/property-category-dictionary');
  if (!response.ok) {
    throw new Error('读取分类字典失败');
  }
  return response.json();
};

const savePropertyCategoryDictionary = async (
  entries: PropertyCategoryEntry[],
): Promise<PropertyCategoryDictionaryResponse> => {
  const response = await fetch('/api/property-category-dictionary', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ entries }),
  });

  if (!response.ok) {
    throw new Error('保存分类字典失败');
  }

  return response.json();
};

const resetPropertyCategoryDictionary = async (): Promise<PropertyCategoryDictionaryResponse> => {
  const response = await fetch('/api/property-category-dictionary/reset', { method: 'POST' });
  if (!response.ok) {
    throw new Error('重置分类字典失败');
  }
  return response.json();
};

const fetchAuditorTeamDictionary = async (): Promise<AuditorTeamDictionaryResponse> => {
  const response = await fetch('/api/auditor-team-dictionary');
  if (!response.ok) {
    throw new Error('读取审核人团队字典失败');
  }
  return response.json();
};

const saveAuditorTeamDictionary = async (
  entries: AuditorTeamEntry[],
): Promise<AuditorTeamDictionaryResponse> => {
  const response = await fetch('/api/auditor-team-dictionary', {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ entries }),
  });

  if (!response.ok) {
    throw new Error('保存审核人团队字典失败');
  }

  return response.json();
};

const resetAuditorTeamDictionary = async (): Promise<AuditorTeamDictionaryResponse> => {
  const response = await fetch('/api/auditor-team-dictionary/reset', { method: 'POST' });
  if (!response.ok) {
    throw new Error('重置审核人团队字典失败');
  }
  return response.json();
};

const generateModelAnalysis = async (payload: {
  report: string;
  model: string;
  context: {
    qualityRows: number;
    efficiencyRows: number;
    dateRange: string;
    filters: string;
  };
}): Promise<AiAnalysisResponse> => {
  const controller = new AbortController();
  const timeoutMs = 180000;
  const timeout = window.setTimeout(() => controller.abort(), timeoutMs);
  let response: Response;

  try {
    response = await fetch('/api/ai-analysis', {
      method: 'POST',
      signal: controller.signal,
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
  } catch (error) {
    if (error instanceof Error && error.name === 'AbortError') {
      throw new Error(`大模型分析等待超过 ${Math.round(timeoutMs / 1000)} 秒，已自动停止。请稍后重试，或切换更快的模型。`);
    }

    throw new Error('大模型接口连接失败，请确认本地/服务器后端正在运行，且当前访问地址没有跨域或网络中断。');
  } finally {
    window.clearTimeout(timeout);
  }

  const body = await response.json().catch(() => ({ message: '大模型接口返回异常，请稍后重试。' }));

  if (!response.ok) {
    throw new Error(body.message || '大模型分析生成失败');
  }

  return body;
};

const generateDailyAlerts = async (payload: {
  date?: string;
  push?: boolean;
}): Promise<{
  message: string;
  meta?: {
    targetDate?: string;
    alertCount?: number;
    textPath?: string;
    feishu?: {
      skipped?: boolean;
      status?: number;
      reason?: string;
    };
  } | null;
  error?: string;
  pushed: boolean;
  generatedAt: string;
}> => {
  const response = await fetch('/api/daily-alerts', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  });
  const body = await response.json().catch(() => ({ message: '每日预警接口返回异常，请稍后重试。' }));

  if (!response.ok) {
    throw new Error(body.message || '每日预警生成失败');
  }

  return body;
};

function App() {
  const [dataset, setDataset] = useState<ParsedWorkbook>(emptyWorkbook);
  const [efficiencyDataset, setEfficiencyDataset] = useState<ParsedEfficiencyWorkbook>(emptyEfficiencyWorkbook);
  const [propertyCategoryDictionary, setPropertyCategoryDictionary] = useState<PropertyCategoryEntry[]>([]);
  const [auditorTeamDictionary, setAuditorTeamDictionary] = useState<AuditorTeamEntry[]>([]);
  const [dictionaryDraft, setDictionaryDraft] = useState<PropertyCategoryEntry>({ propertyName: '', category: '' });
  const [auditorTeamDraft, setAuditorTeamDraft] = useState<AuditorTeamEntry>({ auditorName: '', team: '' });
  const [editingDictionaryKey, setEditingDictionaryKey] = useState('');
  const [editingAuditorTeamKey, setEditingAuditorTeamKey] = useState('');
  const [activeView, setActiveView] = useState<ViewKey>('overview');
  const [overviewStartDateFilter, setOverviewStartDateFilter] = useState(ALL_OPTION);
  const [overviewEndDateFilter, setOverviewEndDateFilter] = useState(ALL_OPTION);
  const [overviewSessionFilter, setOverviewSessionFilter] = useState<string[]>([]);
  const [overviewBatchFilter, setOverviewBatchFilter] = useState<string[]>([]);
  const [attributeStartDateFilter, setAttributeStartDateFilter] = useState(ALL_OPTION);
  const [attributeEndDateFilter, setAttributeEndDateFilter] = useState(ALL_OPTION);
  const [attributeSessionFilter, setAttributeSessionFilter] = useState<string[]>([]);
  const [attributeBatchFilter, setAttributeBatchFilter] = useState<string[]>([]);
  const [attributeFilter, setAttributeFilter] = useState<string[]>([]);
  const [compareStartDateFilter, setCompareStartDateFilter] = useState(ALL_OPTION);
  const [compareEndDateFilter, setCompareEndDateFilter] = useState(ALL_OPTION);
  const [compareDimension, setCompareDimension] = useState<CompareDimension>('session');
  const [compareSessionAFilter, setCompareSessionAFilter] = useState('');
  const [compareSessionBFilter, setCompareSessionBFilter] = useState('');
  const [compareAttributeAFilter, setCompareAttributeAFilter] = useState('');
  const [compareAttributeBFilter, setCompareAttributeBFilter] = useState('');
  const [compareBatchAFilter, setCompareBatchAFilter] = useState('');
  const [compareBatchBFilter, setCompareBatchBFilter] = useState('');
  const [compareAuditorAFilter, setCompareAuditorAFilter] = useState('');
  const [compareAuditorBFilter, setCompareAuditorBFilter] = useState('');
  const [compareAuditorTeamAFilter, setCompareAuditorTeamAFilter] = useState('');
  const [compareAuditorTeamBFilter, setCompareAuditorTeamBFilter] = useState('');
  const [compareQualityMetrics, setCompareQualityMetrics] = useState<CompareQualityMetric[]>([
    'proofAccuracy',
    'exactPassRate',
  ]);
  const [compareSessionFilter, setCompareSessionFilter] = useState(ALL_OPTION);
  const [compareBatchFilter, setCompareBatchFilter] = useState(ALL_OPTION);
  const [compareAttributeFilter, setCompareAttributeFilter] = useState(ALL_OPTION);
  const [compareAuditorFilter, setCompareAuditorFilter] = useState(ALL_OPTION);
  const [compareAuditorTeamFilter, setCompareAuditorTeamFilter] = useState(ALL_OPTION);
  const [efficiencyStartDateFilter, setEfficiencyStartDateFilter] = useState(ALL_OPTION);
  const [efficiencyEndDateFilter, setEfficiencyEndDateFilter] = useState(ALL_OPTION);
  const [efficiencyTeamFilter, setEfficiencyTeamFilter] = useState(ALL_OPTION);
  const [efficiencySessionFilter, setEfficiencySessionFilter] = useState<string[]>([]);
  const [efficiencyEmployeeFilter, setEfficiencyEmployeeFilter] = useState(ALL_OPTION);
  const [efficiencyTimeDimension, setEfficiencyTimeDimension] = useState<EfficiencyTimeDimension>('day');
  const [error, setError] = useState('');
  const [isImporting, setIsImporting] = useState(false);
  const [isEfficiencyImporting, setIsEfficiencyImporting] = useState(false);
  const [isDictionarySaving, setIsDictionarySaving] = useState(false);
  const [isModelAnalyzing, setIsModelAnalyzing] = useState(false);
  const [modelAnalysis, setModelAnalysis] = useState<AiAnalysisResponse | null>(null);
  const [modelAnalysisError, setModelAnalysisError] = useState('');
  const [selectedDeepseekModel, setSelectedDeepseekModel] = useState('deepseek-v4-flash');
  const [customDeepseekModel, setCustomDeepseekModel] = useState('');
  const [aiStartDateFilter, setAiStartDateFilter] = useState(ALL_OPTION);
  const [aiEndDateFilter, setAiEndDateFilter] = useState(ALL_OPTION);
  const [dailyAlertDate, setDailyAlertDate] = useState('');
  const [dailyAlertOutput, setDailyAlertOutput] = useState('');
  const [dailyAlertMeta, setDailyAlertMeta] = useState<{
    targetDate?: string;
    alertCount?: number;
    textPath?: string;
    feishu?: {
      skipped?: boolean;
      status?: number;
      reason?: string;
    };
  } | null>(null);
  const [dailyAlertError, setDailyAlertError] = useState('');
  const [isDailyAlertRunning, setIsDailyAlertRunning] = useState(false);
  const [isLoading, setIsLoading] = useState(true);

  useEffect(() => {
    const loadDataset = async () => {
      try {
        const dictionary = await fetchPropertyCategoryDictionary();
        if (Array.isArray(dictionary.entries)) {
          setPropertyCategoryDictionary(dictionary.entries);
        }
      } catch {
        setPropertyCategoryDictionary([]);
      }

      try {
        const auditorDictionary = await fetchAuditorTeamDictionary();
        if (Array.isArray(auditorDictionary.entries)) {
          setAuditorTeamDictionary(auditorDictionary.entries);
        }
      } catch {
        setAuditorTeamDictionary([]);
      }

      try {
        const sharedDataset = await fetchSharedDataset();
        if (Array.isArray(sharedDataset.rows)) {
          setDataset(sharedDataset);
        }
      } catch {
        setError('共享数据服务暂时不可用，请确认后端已启动。');
      }

      try {
        const sharedEfficiencyDataset = await fetchEfficiencyDataset();
        if (Array.isArray(sharedEfficiencyDataset.rows)) {
          setEfficiencyDataset(sharedEfficiencyDataset);
        }
      } catch {
        setEfficiencyDataset(emptyEfficiencyWorkbook);
      } finally {
        setIsLoading(false);
      }
    };

    void loadDataset();
  }, []);

  const options = useMemo(
    () => ({
      dates: createOptions(dataset.rows, 'date'),
      sessions: createOptions(dataset.rows, 'session'),
      batches: createOptions(dataset.rows, 'batch'),
      attributes: createOptions(dataset.rows, 'attribute'),
      auditors: createOptions(dataset.rows.filter((row) => row.auditor), 'auditor'),
      auditorTeams: [ALL_OPTION].concat(
        [
          ...new Set<string>(
            dataset.rows
              .map((row) => resolveAuditorTeam(row.auditor ?? '', row.auditorTeam ?? '', auditorTeamDictionary))
              .filter((team): team is string => Boolean(team)),
          ),
        ].sort((a, b) => a.localeCompare(b, 'zh-CN')),
      ),
    }),
    [auditorTeamDictionary, dataset.rows],
  );
  const weekOptions = useMemo(() => createWeekOptions(options.dates), [options.dates]);
  const efficiencyDateOptions = useMemo(() => createDateOptions(efficiencyDataset.rows), [efficiencyDataset.rows]);
  const efficiencyWeekOptions = useMemo(
    () => createWeekOptions(efficiencyDateOptions),
    [efficiencyDateOptions],
  );
  const efficiencyTeamOptions = useMemo(
    () => createStringOptions<EfficiencyRow>(efficiencyDataset.rows, (row) => row.team),
    [efficiencyDataset.rows],
  );
  const efficiencySessionOptions = useMemo(
    () => createStringOptions<EfficiencyRow>(efficiencyDataset.rows, (row) => row.session),
    [efficiencyDataset.rows],
  );
  const qualityCoverage = useMemo(() => getDateCoverage(dataset.rows), [dataset.rows]);
  const efficiencyCoverage = useMemo(() => getDateCoverage(efficiencyDataset.rows), [efficiencyDataset.rows]);
  const latestDataDate = useMemo(
    () => [qualityCoverage.end, efficiencyCoverage.end].filter(Boolean).sort((a, b) => b.localeCompare(a))[0] ?? '',
    [efficiencyCoverage.end, qualityCoverage.end],
  );
  const overallCoverage = useMemo(
    () => getDateCoverage([...dataset.rows, ...efficiencyDataset.rows]),
    [dataset.rows, efficiencyDataset.rows],
  );
  const aiDateOptions = useMemo(
    () => createDateOptions([...dataset.rows, ...efficiencyDataset.rows]),
    [dataset.rows, efficiencyDataset.rows],
  );
  const aiWeekOptions = useMemo(() => createWeekOptions(aiDateOptions), [aiDateOptions]);
  const aiFilteredRows = useMemo(
    () => filterRowsByDateRange(dataset.rows, aiStartDateFilter, aiEndDateFilter),
    [aiEndDateFilter, aiStartDateFilter, dataset.rows],
  );
  const aiFilteredEfficiencyRows = useMemo(
    () => filterRowsByDateRange(efficiencyDataset.rows, aiStartDateFilter, aiEndDateFilter),
    [aiEndDateFilter, aiStartDateFilter, efficiencyDataset.rows],
  );
  const aiDateCoverage = useMemo(
    () => getDateCoverage([...aiFilteredRows, ...aiFilteredEfficiencyRows]),
    [aiFilteredEfficiencyRows, aiFilteredRows],
  );

  const overviewFilteredRows = useMemo(
    () =>
      filterRowsByCriteria(dataset.rows, {
        startDate: overviewStartDateFilter,
        endDate: overviewEndDateFilter,
        session: overviewSessionFilter,
        batch: overviewBatchFilter,
        attribute: ALL_OPTION,
      }),
    [dataset.rows, overviewBatchFilter, overviewEndDateFilter, overviewSessionFilter, overviewStartDateFilter],
  );
  const attributeFilteredRows = useMemo(
    () =>
      filterRowsByCriteria(dataset.rows, {
        startDate: attributeStartDateFilter,
        endDate: attributeEndDateFilter,
        session: attributeSessionFilter,
        batch: attributeBatchFilter,
        attribute: attributeFilter,
      }),
    [
      attributeBatchFilter,
      attributeEndDateFilter,
      attributeFilter,
      attributeSessionFilter,
      attributeStartDateFilter,
      dataset.rows,
    ],
  );
  const filteredRows = useMemo(
    () => (activeView === 'attribute' ? attributeFilteredRows : overviewFilteredRows),
    [activeView, attributeFilteredRows, overviewFilteredRows],
  );
  const filteredEfficiencyRows = useMemo(
    () =>
      filterEfficiencyRows(
        efficiencyDataset.rows,
        efficiencyStartDateFilter,
        efficiencyEndDateFilter,
        efficiencyTeamFilter,
        efficiencySessionFilter,
      ),
    [
      efficiencyDataset.rows,
      efficiencyEndDateFilter,
      efficiencySessionFilter,
      efficiencyStartDateFilter,
      efficiencyTeamFilter,
    ],
  );
  const efficiencyEmployeeOptions = useMemo(
    () => createStringOptions<EfficiencyRow>(filteredEfficiencyRows, (row) => row.employee),
    [filteredEfficiencyRows],
  );

  const metrics = useMemo(() => aggregateMetrics(filteredRows), [filteredRows]);
  const trendData = useMemo(() => aggregateTrend(filteredRows), [filteredRows]);
  const topAttributes = useMemo(() => aggregateAttributes(filteredRows), [filteredRows]);
  const ambiguousAttributeRanking = useMemo(() => aggregateAmbiguousAttributes(filteredRows), [filteredRows]);
  const categoryData = useMemo(
    () => aggregateCategories(filteredRows, propertyCategoryDictionary),
    [filteredRows, propertyCategoryDictionary],
  );
  const aiMetrics = useMemo(() => aggregateMetrics(aiFilteredRows), [aiFilteredRows]);
  const aiTopAttributes = useMemo(() => aggregateAttributes(aiFilteredRows), [aiFilteredRows]);
  const aiCategoryData = useMemo(
    () => aggregateCategories(aiFilteredRows, propertyCategoryDictionary),
    [aiFilteredRows, propertyCategoryDictionary],
  );
  const aiEfficiencyMetrics = useMemo(
    () => aggregateEfficiency(aiFilteredEfficiencyRows),
    [aiFilteredEfficiencyRows],
  );
  const aiEfficiencyRanking = useMemo(
    () => aggregateEfficiencyRanking(aiFilteredEfficiencyRows),
    [aiFilteredEfficiencyRows],
  );
  const sessionShareData = useMemo(() => aggregateSessionShares(filteredRows), [filteredRows]);
  const compareSessionOptions = useMemo(
    () => aggregateSessionShares(dataset.rows).map((item) => item.name),
    [dataset.rows],
  );
  const compareAttributeOptions = useMemo(
    () => options.attributes.filter((option) => option !== ALL_OPTION),
    [options.attributes],
  );
  const compareBatchOptions = useMemo(
    () => options.batches.filter((option) => option !== ALL_OPTION),
    [options.batches],
  );
  const compareAuditorOptions = useMemo(
    () => options.auditors.filter((option) => option !== ALL_OPTION),
    [options.auditors],
  );
  const compareAuditorTeamOptions = useMemo(
    () => options.auditorTeams.filter((option) => option !== ALL_OPTION),
    [options.auditorTeams],
  );
  const compareDimensionConfig = COMPARE_DIMENSIONS.find((item) => item.key === compareDimension) ?? COMPARE_DIMENSIONS[0];
  const compareDimensionOptions =
    compareDimension === 'session'
      ? compareSessionOptions
      : compareDimension === 'attribute'
        ? compareAttributeOptions
        : compareDimension === 'batch'
          ? compareBatchOptions
          : compareDimension === 'auditor'
            ? compareAuditorOptions
            : compareAuditorTeamOptions;
  const compareAFilter =
    compareDimension === 'session'
      ? compareSessionAFilter
      : compareDimension === 'attribute'
        ? compareAttributeAFilter
        : compareDimension === 'batch'
          ? compareBatchAFilter
          : compareDimension === 'auditor'
            ? compareAuditorAFilter
            : compareAuditorTeamAFilter;
  const compareBFilter =
    compareDimension === 'session'
      ? compareSessionBFilter
      : compareDimension === 'attribute'
        ? compareAttributeBFilter
        : compareDimension === 'batch'
          ? compareBatchBFilter
          : compareDimension === 'auditor'
            ? compareAuditorBFilter
            : compareAuditorTeamBFilter;
  const setCompareAFilter =
    compareDimension === 'session'
      ? setCompareSessionAFilter
      : compareDimension === 'attribute'
        ? setCompareAttributeAFilter
        : compareDimension === 'batch'
          ? setCompareBatchAFilter
          : compareDimension === 'auditor'
            ? setCompareAuditorAFilter
            : setCompareAuditorTeamAFilter;
  const setCompareBFilter =
    compareDimension === 'session'
      ? setCompareSessionBFilter
      : compareDimension === 'attribute'
        ? setCompareAttributeBFilter
        : compareDimension === 'batch'
          ? setCompareBatchBFilter
          : compareDimension === 'auditor'
            ? setCompareAuditorBFilter
            : setCompareAuditorTeamBFilter;
  const compareA = compareAFilter || compareDimensionOptions[0] || '';
  const compareB = compareBFilter || compareDimensionOptions.find((option) => option !== compareA) || compareA;
  const compareBaseRows = useMemo(
    () =>
      filterRowsByCriteria(dataset.rows, {
        startDate: compareStartDateFilter,
        endDate: compareEndDateFilter,
        session: compareDimension === 'session' ? ALL_OPTION : compareSessionFilter,
        batch: compareDimension === 'batch' ? ALL_OPTION : compareBatchFilter,
        attribute: compareDimension === 'attribute' ? ALL_OPTION : compareAttributeFilter,
        auditor: compareDimension === 'auditor' ? ALL_OPTION : compareAuditorFilter,
        auditorTeam: compareDimension === 'auditorTeam' ? ALL_OPTION : compareAuditorTeamFilter,
        auditorTeamDictionary,
      }),
    [
      compareDimension,
      compareAttributeFilter,
      compareBatchFilter,
      compareAuditorFilter,
      compareAuditorTeamFilter,
      auditorTeamDictionary,
      compareEndDateFilter,
      compareSessionFilter,
      compareStartDateFilter,
      dataset.rows,
    ],
  );
  const compareARows = useMemo(
    () =>
      compareBaseRows.filter((row) =>
        compareDimension === 'auditorTeam'
          ? resolveAuditorTeam(row.auditor ?? '', row.auditorTeam ?? '', auditorTeamDictionary) === compareA
          : String(row[compareDimension] ?? '') === compareA,
      ),
    [auditorTeamDictionary, compareA, compareBaseRows, compareDimension],
  );
  const compareBRows = useMemo(
    () =>
      compareBaseRows.filter((row) =>
        compareDimension === 'auditorTeam'
          ? resolveAuditorTeam(row.auditor ?? '', row.auditorTeam ?? '', auditorTeamDictionary) === compareB
          : String(row[compareDimension] ?? '') === compareB,
      ),
    [auditorTeamDictionary, compareB, compareBaseRows, compareDimension],
  );
  const compareAMetrics = useMemo(() => aggregateMetrics(compareARows), [compareARows]);
  const compareBMetrics = useMemo(() => aggregateMetrics(compareBRows), [compareBRows]);
  const compareTrend = useMemo(
    () => aggregateSessionCompareTrend(compareARows, compareBRows, compareA, compareB),
    [compareA, compareARows, compareB, compareBRows],
  );
  const attributeSessionRows = useMemo(
    () => aggregateDimensionMetrics(filteredRows, 'session'),
    [filteredRows],
  );
  const attributeBatchRows = useMemo(
    () => aggregateDimensionMetrics(filteredRows, 'batch'),
    [filteredRows],
  );
  const efficiencyMetrics = useMemo(() => aggregateEfficiency(filteredEfficiencyRows), [filteredEfficiencyRows]);
  const efficiencyRanking = useMemo(() => aggregateEfficiencyRanking(filteredEfficiencyRows), [filteredEfficiencyRows]);
  const efficiencyTrend = useMemo(() => aggregateEfficiencyTrend(filteredEfficiencyRows), [filteredEfficiencyRows]);
  const workplaceEfficiency = useMemo(
    () => aggregateWorkplaceEfficiency(filteredEfficiencyRows),
    [filteredEfficiencyRows],
  );
  const selectedEfficiencyEmployee = useMemo(
    () =>
      efficiencyEmployeeFilter !== ALL_OPTION
        ? efficiencyEmployeeFilter
        : efficiencyEmployeeOptions.find((option) => option !== ALL_OPTION) ?? '',
    [efficiencyEmployeeFilter, efficiencyEmployeeOptions],
  );
  const employeeEfficiencyDetail = useMemo(
    () =>
      selectedEfficiencyEmployee
        ? aggregateEmployeeEfficiencyDetail(
            filteredEfficiencyRows,
            selectedEfficiencyEmployee,
            efficiencyTimeDimension,
          )
        : null,
    [efficiencyTimeDimension, filteredEfficiencyRows, selectedEfficiencyEmployee],
  );
  const qualityImportHistory = useMemo(
    () => dataset.importHistory?.length ? dataset.importHistory : buildFallbackImportHistory(dataset, 'quality'),
    [dataset],
  );
  const efficiencyImportHistory = useMemo(
    () =>
      efficiencyDataset.importHistory?.length
        ? efficiencyDataset.importHistory
        : buildFallbackImportHistory(efficiencyDataset, 'efficiency'),
    [efficiencyDataset],
  );
  const allImportHistory = useMemo(
    () =>
      [...qualityImportHistory, ...efficiencyImportHistory].sort((a, b) =>
        b.importedAt.localeCompare(a.importedAt),
      ),
    [efficiencyImportHistory, qualityImportHistory],
  );
  const [showAllImportHistory, setShowAllImportHistory] = useState(false);
  const visibleImportHistory = showAllImportHistory ? allImportHistory : allImportHistory.slice(0, 6);
  const aiAnalysis = useMemo(
    () =>
      createAiAnalysis({
        qualityMetrics: aiMetrics,
        qualityRows: aiFilteredRows,
        topAttributes: aiTopAttributes,
        categoryData: aiCategoryData,
        efficiencyMetrics: aiEfficiencyMetrics,
        efficiencyRanking: aiEfficiencyRanking,
        efficiencyRows: aiFilteredEfficiencyRows,
        propertyCategoryDictionary,
      }),
    [
      aiCategoryData,
      aiEfficiencyMetrics,
      aiEfficiencyRanking,
      aiFilteredEfficiencyRows,
      aiFilteredRows,
      aiMetrics,
      aiTopAttributes,
      propertyCategoryDictionary,
    ],
  );
  const aiContext = useMemo(
    () => ({
      qualityRows: aiFilteredRows.length,
      efficiencyRows: aiFilteredEfficiencyRows.length,
      dateRange:
        aiDateCoverage.start && aiDateCoverage.end
          ? `${aiDateCoverage.start} ~ ${aiDateCoverage.end}`
          : '暂无质量数据',
      filters: [
        `AI日期=${formatDateDisplay(aiStartDateFilter) || '全部'}~${formatDateDisplay(aiEndDateFilter) || '全部'}`,
        '场次=全部',
        '批次=全部',
      ].join('；'),
    }),
    [
      aiDateCoverage.end,
      aiDateCoverage.start,
      aiEndDateFilter,
      aiFilteredEfficiencyRows.length,
      aiFilteredRows.length,
      aiStartDateFilter,
    ],
  );
  const activeDeepseekModel =
    selectedDeepseekModel === CUSTOM_MODEL_OPTION
      ? customDeepseekModel.trim()
      : selectedDeepseekModel;
  const isOverviewView = activeView === 'overview';
  const toggleCompareQualityMetric = (metric: CompareQualityMetric) => {
    setCompareQualityMetrics((current) => {
      if (current.includes(metric)) {
        return current.length === 1 ? current : current.filter((item) => item !== metric);
      }
      return [...current, metric];
    });
  };

  const cards: MetricsCardData[] = [
    {
      title: '申报次数',
      value: formatInteger(metrics.declarations),
      hint: `${filteredRows.length.toLocaleString('zh-CN')} 条聚合记录`,
      tone: 'slate',
      icon: <Database size={18} />,
    },
    {
      title: '举证准确率',
      value: formatPercent(metrics.proofAccuracy),
      hint: '按你的定义实时重算',
      tone: 'emerald',
      icon: <ShieldCheck size={18} />,
    },
    {
      title: '精准通过率',
      value: formatPercent(metrics.exactPassRate),
      hint: '排除模糊通过与未通过',
      tone: 'blue',
      icon: <Target size={18} />,
    },
    {
      title: '模棱两可率',
      value: formatPercent(metrics.ambiguousRate),
      hint: `模糊通过 ${formatInteger(metrics.ambiguousPasses)}`,
      tone: 'amber',
      icon: <CircleSlash size={18} />,
    },
    {
      title: '拒绝率',
      value: formatPercent(metrics.rejectRate),
      hint: `未通过 ${formatInteger(metrics.rejects)}`,
      tone: 'rose',
      icon: <Ban size={18} />,
    },
  ];

  const handleImport = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    setIsImporting(true);
    setError('');

    try {
      const parsed = await parseWorkbookFile(file, propertyCategoryDictionary);
      const nextDataset = await mergeSharedDataset(parsed);
      setDataset(nextDataset);
      setOverviewStartDateFilter(ALL_OPTION);
      setOverviewEndDateFilter(ALL_OPTION);
      setOverviewSessionFilter([]);
      setOverviewBatchFilter([]);
      setAttributeStartDateFilter(ALL_OPTION);
      setAttributeEndDateFilter(ALL_OPTION);
      setAttributeSessionFilter([]);
      setAttributeBatchFilter([]);
      setAttributeFilter([]);
      setCompareStartDateFilter(ALL_OPTION);
      setCompareEndDateFilter(ALL_OPTION);
      setCompareSessionAFilter('');
      setCompareSessionBFilter('');
      setCompareBatchFilter(ALL_OPTION);
      setCompareAttributeFilter(ALL_OPTION);
      setCompareAuditorFilter(ALL_OPTION);
      setCompareAuditorTeamFilter(ALL_OPTION);
    } catch (err) {
      setError(err instanceof Error ? err.message : '导入失败，请检查文件格式。');
    } finally {
      setIsImporting(false);
      event.target.value = '';
    }
  };

  const clearData = async () => {
    try {
      setError('');
      const nextDataset = await clearSharedDataset();
      setDataset(nextDataset);
      setOverviewStartDateFilter(ALL_OPTION);
      setOverviewEndDateFilter(ALL_OPTION);
      setOverviewSessionFilter([]);
      setOverviewBatchFilter([]);
      setAttributeStartDateFilter(ALL_OPTION);
      setAttributeEndDateFilter(ALL_OPTION);
      setAttributeSessionFilter([]);
      setAttributeBatchFilter([]);
      setAttributeFilter([]);
      setCompareStartDateFilter(ALL_OPTION);
      setCompareEndDateFilter(ALL_OPTION);
      setCompareSessionAFilter('');
      setCompareSessionBFilter('');
      setCompareAuditorAFilter('');
      setCompareAuditorBFilter('');
      setCompareAuditorTeamAFilter('');
      setCompareAuditorTeamBFilter('');
      setCompareBatchFilter(ALL_OPTION);
      setCompareAttributeFilter(ALL_OPTION);
      setCompareAuditorFilter(ALL_OPTION);
      setCompareAuditorTeamFilter(ALL_OPTION);
    } catch (err) {
      setError(err instanceof Error ? err.message : '清空共享数据失败。');
    }
  };

  const handleEfficiencyImport = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    setIsEfficiencyImporting(true);
    setError('');

    try {
      const parsed = await parseEfficiencyWorkbookFile(file);
      const nextDataset = await mergeEfficiencyDataset(parsed);
      setEfficiencyDataset(nextDataset);
      setEfficiencyStartDateFilter(ALL_OPTION);
      setEfficiencyEndDateFilter(ALL_OPTION);
      setEfficiencyTeamFilter(ALL_OPTION);
      setEfficiencySessionFilter([]);
      setEfficiencyEmployeeFilter(ALL_OPTION);
      setActiveView('efficiency');
    } catch (err) {
      setError(err instanceof Error ? err.message : '人效数据导入失败，请检查文件格式。');
    } finally {
      setIsEfficiencyImporting(false);
      event.target.value = '';
    }
  };

  const clearEfficiencyData = async () => {
    try {
      setError('');
      const nextDataset = await clearEfficiencyDataset();
      setEfficiencyDataset(nextDataset);
      setEfficiencyStartDateFilter(ALL_OPTION);
      setEfficiencyEndDateFilter(ALL_OPTION);
      setEfficiencyTeamFilter(ALL_OPTION);
      setEfficiencySessionFilter([]);
      setEfficiencyEmployeeFilter(ALL_OPTION);
    } catch (err) {
      setError(err instanceof Error ? err.message : '清空人效数据失败。');
    }
  };

  const runModelAnalysis = async () => {
    if (!activeDeepseekModel) {
      setModelAnalysisError('请先选择或填写 DeepSeek 模型名称。');
      return;
    }

    setIsModelAnalyzing(true);
    setModelAnalysisError('');

    try {
      const response = await generateModelAnalysis({
        report: aiAnalysis.report,
        model: activeDeepseekModel,
        context: aiContext,
      });
      setModelAnalysis(response);
  } catch (err) {
      setModelAnalysisError(
        err instanceof Error
            ? err.message
            : '大模型分析生成失败。',
      );
    } finally {
      setIsModelAnalyzing(false);
    }
  };

  const runDailyAlert = async (push: boolean) => {
    setIsDailyAlertRunning(true);
    setDailyAlertError('');

    try {
      const response = await generateDailyAlerts({
        date: dailyAlertDate || undefined,
        push,
      });
      setDailyAlertOutput(response.message);
      setDailyAlertMeta(response.meta ?? null);
    } catch (err) {
      setDailyAlertError(err instanceof Error ? err.message : '每日预警生成失败。');
    } finally {
      setIsDailyAlertRunning(false);
    }
  };

  const persistPropertyCategoryDictionary = async (entries: PropertyCategoryEntry[]) => {
    setIsDictionarySaving(true);
    setError('');

    try {
      const response = await savePropertyCategoryDictionary(entries);
      setPropertyCategoryDictionary(response.entries);
      setDictionaryDraft({ propertyName: '', category: '' });
      setEditingDictionaryKey('');
    } catch (err) {
      setError(err instanceof Error ? err.message : '保存分类字典失败。');
    } finally {
      setIsDictionarySaving(false);
    }
  };

  const upsertDictionaryEntry = async () => {
    const propertyName = dictionaryDraft.propertyName.trim();
    const category = normalizeDictionaryCategory(dictionaryDraft.category);

    if (!propertyName || !category) {
      setError('请填写属性项和属性项分类。');
      return;
    }

    const nextEntries = propertyCategoryDictionary.filter(
      (entry) => normalizePropertyName(entry.propertyName) !== normalizePropertyName(editingDictionaryKey || propertyName),
    );

    await persistPropertyCategoryDictionary([...nextEntries, { propertyName, category }]);
  };

  const editDictionaryEntry = (entry: PropertyCategoryEntry) => {
    setDictionaryDraft(entry);
    setEditingDictionaryKey(entry.propertyName);
    setActiveView('dictionary');
  };

  const deleteDictionaryEntry = async (propertyName: string) => {
    await persistPropertyCategoryDictionary(
      propertyCategoryDictionary.filter(
        (entry) => normalizePropertyName(entry.propertyName) !== normalizePropertyName(propertyName),
      ),
    );
  };

  const handleDictionaryImport = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    setIsDictionarySaving(true);
    setError('');

    try {
      const entries = await parsePropertyCategoryDictionaryFile(file);
      await persistPropertyCategoryDictionary(entries);
    } catch (err) {
      setError(err instanceof Error ? err.message : '导入分类字典失败，请检查文件格式。');
    } finally {
      setIsDictionarySaving(false);
      event.target.value = '';
    }
  };

  const resetDictionary = async () => {
    setIsDictionarySaving(true);
    setError('');

    try {
      const response = await resetPropertyCategoryDictionary();
      setPropertyCategoryDictionary(response.entries);
      setDictionaryDraft({ propertyName: '', category: '' });
      setEditingDictionaryKey('');
    } catch (err) {
      setError(err instanceof Error ? err.message : '重置分类字典失败。');
    } finally {
      setIsDictionarySaving(false);
    }
  };

  const persistAuditorTeamDictionary = async (entries: AuditorTeamEntry[]) => {
    setIsDictionarySaving(true);
    setError('');

    try {
      const response = await saveAuditorTeamDictionary(entries);
      setAuditorTeamDictionary(response.entries);
      setAuditorTeamDraft({ auditorName: '', team: '' });
      setEditingAuditorTeamKey('');
    } catch (err) {
      setError(err instanceof Error ? err.message : '保存审核人团队字典失败。');
    } finally {
      setIsDictionarySaving(false);
    }
  };

  const upsertAuditorTeamEntry = async () => {
    const auditorName = auditorTeamDraft.auditorName.trim();
    const team = auditorTeamDraft.team.trim();

    if (!auditorName || !team) {
      setError('请填写审核人和团队。');
      return;
    }

    const nextEntries = auditorTeamDictionary.filter(
      (entry) => normalizeAuditorName(entry.auditorName) !== normalizeAuditorName(editingAuditorTeamKey || auditorName),
    );

    await persistAuditorTeamDictionary([...nextEntries, { auditorName, team }]);
  };

  const editAuditorTeamEntry = (entry: AuditorTeamEntry) => {
    setAuditorTeamDraft(entry);
    setEditingAuditorTeamKey(entry.auditorName);
    setActiveView('dictionary');
  };

  const deleteAuditorTeamEntry = async (auditorName: string) => {
    await persistAuditorTeamDictionary(
      auditorTeamDictionary.filter(
        (entry) => normalizeAuditorName(entry.auditorName) !== normalizeAuditorName(auditorName),
      ),
    );
  };

  const handleAuditorTeamImport = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }

    setIsDictionarySaving(true);
    setError('');

    try {
      const entries = await parseAuditorTeamDictionaryFile(file);
      await persistAuditorTeamDictionary(entries);
    } catch (err) {
      setError(err instanceof Error ? err.message : '导入审核人团队字典失败，请检查文件格式。');
    } finally {
      setIsDictionarySaving(false);
      event.target.value = '';
    }
  };

  const resetAuditorDictionary = async () => {
    setIsDictionarySaving(true);
    setError('');

    try {
      const response = await resetAuditorTeamDictionary();
      setAuditorTeamDictionary(response.entries);
      setAuditorTeamDraft({ auditorName: '', team: '' });
      setEditingAuditorTeamKey('');
    } catch (err) {
      setError(err instanceof Error ? err.message : '重置审核人团队字典失败。');
    } finally {
      setIsDictionarySaving(false);
    }
  };

  return (
    <div className="min-h-screen bg-[radial-gradient(circle_at_top,_#f6efe4_0%,_#f3f8f6_40%,_#eef2ff_100%)] text-slate-900">
      <div className="mx-auto flex max-w-[1600px] gap-6 px-4 py-6 sm:px-6 lg:px-8">
        <aside className="hidden w-[240px] shrink-0 lg:block">
          <div className="sticky top-6 rounded-[34px] bg-[linear-gradient(180deg,_#0d1424_0%,_#121a2d_100%)] p-5 text-white shadow-[0_24px_64px_rgba(15,23,42,0.24)]">
            <div className="mb-8 flex items-center gap-3">
              <div className="flex h-14 w-14 items-center justify-center rounded-3xl bg-[linear-gradient(135deg,_#2ad8ff_0%,_#1493ff_100%)] text-slate-950">
                <Layers3 size={24} />
              </div>
              <div>
                <p className="font-display text-xl font-semibold">PQA</p>
                <p className="text-xs tracking-[0.22em] text-slate-400">Dashboard</p>
              </div>
            </div>

            <div className="space-y-3">
              <SidebarNavItem
                icon={<House size={20} />}
                label="首页"
                active={activeView === 'overview'}
                onClick={() => setActiveView('overview')}
              />
              <SidebarNavItem
                icon={<GitCompareArrows size={20} />}
                label="对比分析"
                active={activeView === 'compare'}
                onClick={() => setActiveView('compare')}
              />
              <SidebarNavItem
                icon={<Tags size={20} />}
                label="属性项分析"
                active={activeView === 'attribute'}
                onClick={() => setActiveView('attribute')}
              />
              <SidebarNavItem
                icon={<Users size={20} />}
                label="人效分析"
                active={activeView === 'efficiency'}
                onClick={() => setActiveView('efficiency')}
              />
              <SidebarNavItem
                icon={<BrainCircuit size={20} />}
                label="AI 分析"
                active={activeView === 'ai'}
                onClick={() => setActiveView('ai')}
              />
              <SidebarNavItem
                icon={<ClipboardList size={20} />}
                label="数据导入"
                active={activeView === 'import'}
                onClick={() => setActiveView('import')}
              />
              <SidebarNavItem
                icon={<Tags size={20} />}
                label="分类字典"
                active={activeView === 'dictionary'}
                onClick={() => setActiveView('dictionary')}
              />
            </div>
          </div>
        </aside>

        <div className="min-w-0 flex-1">
          <div className="mb-4 flex gap-3 lg:hidden">
            <MobileNavChip label="首页" active={activeView === 'overview'} onClick={() => setActiveView('overview')} />
            <MobileNavChip label="对比分析" active={activeView === 'compare'} onClick={() => setActiveView('compare')} />
            <MobileNavChip label="属性项分析" active={activeView === 'attribute'} onClick={() => setActiveView('attribute')} />
            <MobileNavChip label="人效分析" active={activeView === 'efficiency'} onClick={() => setActiveView('efficiency')} />
            <MobileNavChip label="AI 分析" active={activeView === 'ai'} onClick={() => setActiveView('ai')} />
            <MobileNavChip label="数据导入" active={activeView === 'import'} onClick={() => setActiveView('import')} />
            <MobileNavChip label="分类字典" active={activeView === 'dictionary'} onClick={() => setActiveView('dictionary')} />
          </div>

          <motion.section
            initial={{ opacity: 0, y: 18 }}
            animate={{ opacity: 1, y: 0 }}
            className="overflow-visible rounded-[32px] border border-white/70 bg-white/80 shadow-[0_30px_80px_rgba(15,23,42,0.08)] backdrop-blur"
          >
          <div className={isOverviewView ? 'p-5 lg:p-6' : 'p-3 lg:p-4'}>
            <div className={`${isOverviewView ? 'rounded-[28px] p-4' : 'rounded-[22px] p-3'} border border-slate-200 bg-white/90 shadow-[0_12px_35px_rgba(15,23,42,0.04)]`}>
              <div className={`${isOverviewView ? 'gap-4' : 'gap-3'} flex flex-col xl:flex-row xl:items-center xl:justify-between`}>
                <div className="flex min-w-0 items-center gap-3">
                  <div className={`${isOverviewView ? 'h-12 w-12 rounded-2xl' : 'h-10 w-10 rounded-xl'} flex shrink-0 items-center justify-center bg-slate-900 text-white`}>
                    <Layers3 size={isOverviewView ? 22 : 18} />
                  </div>
                  <div className="min-w-0">
                    <div className="flex flex-wrap items-center gap-x-3 gap-y-1">
                      <h1 className={`${isOverviewView ? 'text-2xl' : 'text-xl'} font-display font-semibold text-slate-900`}>预质检质量看板</h1>
                      <span className="rounded-full bg-emerald-50 px-3 py-1 text-xs font-medium text-emerald-700">
                        {isLoading ? '加载中' : '每周复用'}
                      </span>
                    </div>
                    <p className={`${isOverviewView ? 'block' : 'hidden xl:block'} mt-1 text-sm text-slate-500`}>
                      聚焦举证准确率、精准通过率、模棱两可率与拒绝率，按场次、批次、属性项动态拆解。
                    </p>
                  </div>
                </div>
                <div className={`grid gap-2 sm:grid-cols-2 ${isOverviewView ? 'xl:min-w-[420px]' : 'xl:min-w-[360px]'}`}>
                  <div className={`${isOverviewView ? 'rounded-2xl px-4 py-3' : 'rounded-xl px-3 py-2'} border border-emerald-100 bg-emerald-50/80`}>
                    <div className="flex items-center gap-2 text-xs font-semibold uppercase tracking-[0.16em] text-emerald-700">
                      <CalendarDays size={14} />
                      最新数据
                    </div>
                    <p className={`${isOverviewView ? 'text-xl' : 'text-base'} mt-1 font-display font-semibold text-slate-950`}>
                      {latestDataDate || '暂无数据'}
                    </p>
                  </div>
                  <div className={`${isOverviewView ? 'rounded-2xl px-4 py-3' : 'rounded-xl px-3 py-2'} border border-slate-200 bg-slate-50/90`}>
                    <p className="text-xs font-semibold uppercase tracking-[0.16em] text-slate-400">可查询区间</p>
                    <p className="mt-1 text-sm font-semibold text-slate-800">
                      {overallCoverage.start && overallCoverage.end
                        ? `${overallCoverage.start} ~ ${overallCoverage.end}`
                        : '暂无可查询数据'}
                    </p>
                    <p className={`${isOverviewView ? 'block' : 'hidden'} mt-1 text-xs text-slate-500`}>
                      质量 {qualityCoverage.end || '-'} / 人效 {efficiencyCoverage.end || '-'}
                    </p>
                  </div>
                </div>
              </div>

              {error ? (
                <div className="mt-3 rounded-2xl border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-700">
                  {error}
                </div>
              ) : null}
            </div>

            <div className={`${isOverviewView ? 'mt-4 rounded-[28px] p-4' : 'mt-3 rounded-[22px] p-3'} bg-[linear-gradient(135deg,_#12212d_0%,_#182b39_100%)] text-white`}>
              <div className={`${isOverviewView ? 'gap-4' : 'gap-3'} flex flex-col`}>
                {activeView === 'compare' ? (
                  <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-[minmax(340px,1.8fr)_minmax(180px,0.9fr)] xl:items-end">
                    <DateRangeFilter
                      label="共同日期区间"
                      startValue={compareStartDateFilter}
                      endValue={compareEndDateFilter}
                      options={options.dates}
                      onStartChange={setCompareStartDateFilter}
                      onEndChange={setCompareEndDateFilter}
                      onClear={() => {
                        setCompareStartDateFilter(ALL_OPTION);
                        setCompareEndDateFilter(ALL_OPTION);
                      }}
                      compact
                    />
                    <WeekQuickSelect
                      label="共同周区间"
                      value={getWeekValue(compareStartDateFilter, compareEndDateFilter, weekOptions)}
                      options={weekOptions}
                      onChange={(week) => {
                        setCompareStartDateFilter(week.start);
                        setCompareEndDateFilter(week.end);
                      }}
                      onClear={() => {
                        setCompareStartDateFilter(ALL_OPTION);
                        setCompareEndDateFilter(ALL_OPTION);
                      }}
                      compact={!isOverviewView}
                    />
                  </div>
                ) : activeView === 'efficiency' ? (
                  <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-[minmax(340px,1.7fr)_minmax(180px,0.8fr)_minmax(170px,0.75fr)_minmax(170px,0.75fr)_minmax(160px,0.7fr)] xl:items-end">
                    <DateRangeFilter
                      label="人效日期区间"
                      startValue={efficiencyStartDateFilter}
                      endValue={efficiencyEndDateFilter}
                      options={efficiencyDateOptions}
                      onStartChange={setEfficiencyStartDateFilter}
                      onEndChange={setEfficiencyEndDateFilter}
                      onClear={() => {
                        setEfficiencyStartDateFilter(ALL_OPTION);
                        setEfficiencyEndDateFilter(ALL_OPTION);
                      }}
                      compact
                    />
                    <WeekQuickSelect
                      label="人效周区间"
                      value={getWeekValue(efficiencyStartDateFilter, efficiencyEndDateFilter, efficiencyWeekOptions)}
                      options={efficiencyWeekOptions}
                      onChange={(week) => {
                        setEfficiencyStartDateFilter(week.start);
                        setEfficiencyEndDateFilter(week.end);
                      }}
                      onClear={() => {
                        setEfficiencyStartDateFilter(ALL_OPTION);
                        setEfficiencyEndDateFilter(ALL_OPTION);
                      }}
                      compact={!isOverviewView}
                    />
                    <FilterSelect
                      label="团队"
                      icon={<Users size={16} />}
                      value={efficiencyTeamFilter}
                      options={efficiencyTeamOptions}
                      onChange={setEfficiencyTeamFilter}
                      compact={!isOverviewView}
                    />
                    <MultiFilterSelect
                      label="场次"
                      icon={<Layers3 size={16} />}
                      value={efficiencySessionFilter}
                      options={efficiencySessionOptions}
                      onChange={setEfficiencySessionFilter}
                    />
                    <div className="rounded-2xl border border-white/10 bg-white/5 px-4 py-3 text-sm text-slate-200">
                      <p className="font-medium text-white">当前人效记录</p>
                      <p className="mt-1 text-xs text-slate-300">
                        {formatInteger(filteredEfficiencyRows.length)} / {formatInteger(efficiencyDataset.rows.length)} 条
                      </p>
                    </div>
                  </div>
                ) : activeView === 'ai' || activeView === 'import' || activeView === 'dictionary' ? (
                  <div className="rounded-xl border border-white/10 bg-white/5 px-3 py-2.5 text-sm text-slate-200">
                    <div className="flex flex-wrap items-center justify-between gap-3">
                      <div>
                        <p className="font-medium text-white">
                          {activeView === 'ai'
                            ? 'AI 分析准备区'
                            : activeView === 'import'
                              ? '数据导入与记录区'
                              : '属性项分类字典管理区'}
                        </p>
                        <p className="mt-0.5 text-xs leading-5 text-slate-300">
                          {activeView === 'ai'
                            ? 'AI 分析将读取质量数据与人效数据，先沉淀规则化洞察，后续可接入大模型生成周报。'
                            : activeView === 'import'
                              ? '质量周数据和人效周数据在这里统一导入，并保留每次导入记录。'
                              : '底表可不再提供分类，看板会优先依据本地字典为属性项自动归类。'}
                        </p>
                      </div>
                      <span className="rounded-full bg-white/10 px-3 py-1 text-xs text-slate-200">
                        {activeView === 'ai'
                          ? `${formatInteger(dataset.rows.length)} 条质量 / ${formatInteger(efficiencyDataset.rows.length)} 条人效`
                          : activeView === 'import'
                            ? `${formatInteger(allImportHistory.length)} 条导入记录`
                            : `${formatInteger(propertyCategoryDictionary.length)} 条字典`}
                      </span>
                    </div>
                  </div>
                ) : activeView === 'attribute' ? (
                  <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-[minmax(340px,1.8fr)_minmax(180px,0.9fr)_repeat(3,minmax(150px,1fr))] xl:items-end">
                    <DateRangeFilter
                      label="日期区间"
                      startValue={attributeStartDateFilter}
                      endValue={attributeEndDateFilter}
                      options={options.dates}
                      onStartChange={setAttributeStartDateFilter}
                      onEndChange={setAttributeEndDateFilter}
                      onClear={() => {
                        setAttributeStartDateFilter(ALL_OPTION);
                        setAttributeEndDateFilter(ALL_OPTION);
                      }}
                      compact
                    />
                    <WeekQuickSelect
                      label="周区间"
                      value={getWeekValue(attributeStartDateFilter, attributeEndDateFilter, weekOptions)}
                      options={weekOptions}
                      onChange={(week) => {
                        setAttributeStartDateFilter(week.start);
                        setAttributeEndDateFilter(week.end);
                      }}
                      onClear={() => {
                        setAttributeStartDateFilter(ALL_OPTION);
                        setAttributeEndDateFilter(ALL_OPTION);
                      }}
                      compact={!isOverviewView}
                    />
                    <MultiFilterSelect
                      label="场次"
                      icon={<Layers3 size={16} />}
                      value={attributeSessionFilter}
                      options={options.sessions}
                      onChange={setAttributeSessionFilter}
                    />
                    <MultiFilterSelect
                      label="批次"
                      icon={<Boxes size={16} />}
                      value={attributeBatchFilter}
                      options={options.batches}
                      onChange={setAttributeBatchFilter}
                    />
                    <MultiFilterSelect
                      label="属性项"
                      icon={<Tags size={16} />}
                      value={attributeFilter}
                      options={options.attributes}
                      onChange={setAttributeFilter}
                    />
                  </div>
                ) : (
                  <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-[minmax(340px,1.8fr)_minmax(180px,0.9fr)_repeat(2,minmax(150px,1fr))] xl:items-end">
                    <DateRangeFilter
                      label="日期区间"
                      startValue={overviewStartDateFilter}
                      endValue={overviewEndDateFilter}
                      options={options.dates}
                      onStartChange={setOverviewStartDateFilter}
                      onEndChange={setOverviewEndDateFilter}
                      onClear={() => {
                        setOverviewStartDateFilter(ALL_OPTION);
                        setOverviewEndDateFilter(ALL_OPTION);
                      }}
                      compact
                    />
                    <WeekQuickSelect
                      label="周区间"
                      value={getWeekValue(overviewStartDateFilter, overviewEndDateFilter, weekOptions)}
                      options={weekOptions}
                      onChange={(week) => {
                        setOverviewStartDateFilter(week.start);
                        setOverviewEndDateFilter(week.end);
                      }}
                      onClear={() => {
                        setOverviewStartDateFilter(ALL_OPTION);
                        setOverviewEndDateFilter(ALL_OPTION);
                      }}
                    />
                    <MultiFilterSelect
                      label="场次"
                      icon={<Layers3 size={16} />}
                      value={overviewSessionFilter}
                      options={options.sessions}
                      onChange={setOverviewSessionFilter}
                    />
                    <MultiFilterSelect
                      label="批次"
                      icon={<Boxes size={16} />}
                      value={overviewBatchFilter}
                      options={options.batches}
                      onChange={setOverviewBatchFilter}
                    />
                  </div>
                )}
                {activeView === 'compare' ? (
                  <div className="grid gap-3 md:grid-cols-3 xl:grid-cols-5">
                    {compareDimension !== 'session' ? (
                      <FilterSelect
                        label="场次"
                        icon={<Layers3 size={16} />}
                        value={compareSessionFilter}
                        options={options.sessions}
                        onChange={setCompareSessionFilter}
                        compact={!isOverviewView}
                      />
                    ) : null}
                    {compareDimension !== 'batch' ? (
                      <FilterSelect
                        label="批次"
                        icon={<Boxes size={16} />}
                        value={compareBatchFilter}
                        options={options.batches}
                        onChange={setCompareBatchFilter}
                        compact={!isOverviewView}
                      />
                    ) : null}
                    {compareDimension !== 'attribute' ? (
                      <FilterSelect
                        label="属性项"
                        icon={<Tags size={16} />}
                        value={compareAttributeFilter}
                        options={options.attributes}
                        onChange={setCompareAttributeFilter}
                        compact={!isOverviewView}
                      />
                    ) : null}
                    {compareDimension !== 'auditor' ? (
                      <FilterSelect
                        label="审核人"
                        icon={<Users size={16} />}
                        value={compareAuditorFilter}
                        options={options.auditors}
                        onChange={setCompareAuditorFilter}
                        compact={!isOverviewView}
                      />
                    ) : null}
                    {compareDimension !== 'auditorTeam' ? (
                      <FilterSelect
                        label="审核团队"
                        icon={<Users size={16} />}
                        value={compareAuditorTeamFilter}
                        options={options.auditorTeams}
                        onChange={setCompareAuditorTeamFilter}
                        compact={!isOverviewView}
                      />
                    ) : null}
                  </div>
                ) : null}
                <details className={`${isOverviewView ? 'rounded-2xl px-4 py-3' : 'rounded-xl px-3 py-2'} border border-white/10 bg-white/5 text-sm text-slate-200`}>
                  <summary className="flex cursor-pointer list-none items-center justify-between gap-3 font-medium marker:content-none">
                    当前口径
                    <ChevronDown size={16} className="shrink-0" />
                  </summary>
                  <ul className="mt-3 space-y-2 text-xs leading-6 text-slate-300">
                    <li>精准通过次数：第一次在线审核结果为通过，且第一次在线审核变动属性值不为空。</li>
                    <li>模糊通过次数：第一次在线审核结果为通过，且第一次在线审核变动属性值为空。</li>
                    <li>多属性申请按整单影响处理：同一单满足精准或模糊条件时，申报属性都会计入对应次数。</li>
                    <li>精准通过率 = 精准通过次数 / 申报次数</li>
                    <li>举证准确率 = (申报次数 - 模糊通过次数 - 举证未通过次数) / 申报次数</li>
                    <li>模棱两可率 = 模糊通过次数 / 申报次数</li>
                    <li>拒绝率 = 未通过次数 / 申报次数</li>
                  </ul>
                </details>
              </div>
            </div>
          </div>
          </motion.section>

          {activeView === 'overview' ? (
            <>
              <section className="mt-8">
                <motion.div initial={{ opacity: 0, y: 14 }} animate={{ opacity: 1, y: 0 }}>
                  <div className="rounded-[26px] border border-slate-200 bg-white p-5 shadow-[0_14px_35px_rgba(15,23,42,0.04)]">
                    <div className="flex items-start justify-between gap-4">
                      <div>
                        <p className="text-sm text-slate-500">申报次数</p>
                        <p className="mt-3 font-display text-4xl text-slate-900">{formatInteger(metrics.declarations)}</p>
                      </div>
                      <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-slate-900 text-white">
                        <Database size={20} />
                      </div>
                    </div>
                    <p className="mt-4 text-sm text-slate-500">{filteredRows.length.toLocaleString('zh-CN')} 条聚合记录</p>
                  </div>
                </motion.div>
              </section>

              <section className="mt-5 grid gap-5 md:grid-cols-2 xl:grid-cols-4">
                {cards.slice(1).map((card, index) => (
                  <motion.div
                    key={card.title}
                    initial={{ opacity: 0, y: 14 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ delay: index * 0.05 }}
                  >
                    <StatCard {...card} />
                  </motion.div>
                ))}
              </section>

              <section className="mt-8">
                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <div className="mb-6 flex items-center justify-between">
                    <div>
                      <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Trend</p>
                      <h2 className="mt-2 font-display text-2xl text-slate-900">举证准确率与精准通过率趋势</h2>
                    </div>
                    <div className="rounded-full bg-slate-100 px-4 py-2 text-sm text-slate-600">
                      {dataset.rows.length ? `共 ${trendData.length} 天` : '等待导入'}
                    </div>
                  </div>

                  <div className="h-[320px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <AreaChart data={trendData}>
                        <defs>
                          <linearGradient id="proofAccuracyFill" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="5%" stopColor="#0f766e" stopOpacity={0.24} />
                            <stop offset="95%" stopColor="#0f766e" stopOpacity={0} />
                          </linearGradient>
                        </defs>
                        <CartesianGrid vertical={false} stroke="#e2e8f0" strokeDasharray="3 3" />
                        <XAxis dataKey="date" tickLine={false} axisLine={false} fontSize={12} />
                        <YAxis tickLine={false} axisLine={false} fontSize={12} domain={[0, 1]} tickFormatter={(value) => `${Math.round(value * 100)}%`} />
                        <Tooltip content={<RateTrendTooltip />} />
                        <Area
                          type="monotone"
                          dataKey="proofAccuracy"
                          stroke="#0f766e"
                          fill="url(#proofAccuracyFill)"
                          strokeWidth={2.5}
                          name="举证准确率"
                        />
                        <Area
                          type="monotone"
                          dataKey="exactPassRate"
                          stroke="#1d4ed8"
                          fill="transparent"
                          strokeWidth={2.5}
                          name="精准通过率"
                        />
                      </AreaChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </section>

              <section className="mt-8 grid gap-6 xl:grid-cols-2">
                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <div className="mb-6">
                    <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Category Mix</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">属性项分类分布</h2>
                    <p className="mt-2 text-sm leading-6 text-slate-500">按当前首页筛选范围内的申报次数计算，同时展示分类质量表现。</p>
                  </div>
                  <div className="grid gap-4 lg:grid-cols-[minmax(180px,0.8fr)_minmax(0,1.2fr)] lg:items-center">
                    <div className="h-[220px] min-w-0">
                      <ResponsiveContainer width="100%" height="100%">
                        <PieChart margin={{ top: 8, right: 8, bottom: 8, left: 8 }}>
                          <Pie
                            data={categoryData}
                            dataKey="value"
                            nameKey="name"
                            cx="50%"
                            cy="50%"
                            innerRadius={46}
                            outerRadius={82}
                            paddingAngle={2}
                          >
                            {categoryData.map((item, index) => (
                              <Cell
                                key={item.name}
                                fill={CATEGORY_COLORS[index % CATEGORY_COLORS.length]}
                              />
                            ))}
                          </Pie>
                          <Tooltip formatter={(value: number) => formatPercent(value)} />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                    <div className="space-y-2">
                      {categoryData.map((item, index) => (
                        <div key={item.name} className="rounded-2xl bg-slate-50 px-3 py-2.5">
                          <div className="flex items-center justify-between gap-3">
                            <div className="flex min-w-0 items-center gap-2">
                              <span
                                className="h-2.5 w-2.5 shrink-0 rounded-full"
                                style={{
                                  backgroundColor: CATEGORY_COLORS[index % CATEGORY_COLORS.length],
                                }}
                              />
                              <span className="truncate text-sm font-medium text-slate-700">{item.name}</span>
                            </div>
                            <span className="shrink-0 font-display text-sm font-semibold text-slate-900">
                              {formatPercent(item.value)}
                            </span>
                          </div>
                          <div className="mt-2 grid grid-cols-2 gap-2 text-xs text-slate-500">
                            <span>举证 {formatPercent(item.proofAccuracy)}</span>
                            <span>通过 {formatPercent(item.exactPassRate)}</span>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
                <SessionShareCard
                  data={sessionShareData}
                  title="场次占比分布"
                  subtitle="按当前首页筛选范围内的申报次数计算，各场次占比合计为 100%。"
                  compact
                  metricsOnly
                />
              </section>
            </>
          ) : activeView === 'compare' ? (
            <>
              <section className="mt-8 grid gap-6 xl:grid-cols-2">
                <SessionCompareSummaryCard label={`${compareDimensionConfig.label} A`} session={compareA} metrics={compareAMetrics} tone="cyan" />
                <SessionCompareSummaryCard label={`${compareDimensionConfig.label} B`} session={compareB} metrics={compareBMetrics} tone="orange" />
              </section>

              <section className="mt-8 rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                <div className="mb-6 flex flex-wrap items-start justify-between gap-4">
                  <div>
                    <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Quality Trend Compare</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">{compareDimensionConfig.title}</h2>
                    <p className="mt-2 text-sm text-slate-500">{compareDimensionConfig.description}</p>
                  </div>
                  <span className="rounded-full bg-slate-100 px-3 py-2 text-xs font-medium text-slate-600">
                    {formatDateDisplay(compareStartDateFilter) || '全部日期'} ~ {formatDateDisplay(compareEndDateFilter) || '全部日期'}
                  </span>
                </div>
                <div className="mb-4 rounded-2xl border border-slate-100 bg-slate-50/80 p-3">
                  <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
                    <div>
                      <p className="text-xs font-semibold uppercase tracking-[0.2em] text-slate-500">对比对象</p>
                      <p className="mt-1 text-xs text-slate-400">选择本次要对比的维度，再选择 A / B 两个对象。</p>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      {COMPARE_DIMENSIONS.map((dimension) => {
                        const selected = compareDimension === dimension.key;
                        return (
                          <button
                            key={dimension.key}
                            type="button"
                            onClick={() => setCompareDimension(dimension.key)}
                            aria-pressed={selected}
                            className={`rounded-full border px-3 py-2 text-xs font-medium transition ${
                              selected
                                ? 'border-slate-900 bg-slate-900 text-white shadow-sm'
                                : 'border-slate-200 bg-white text-slate-500 hover:border-slate-300 hover:text-slate-800'
                            }`}
                          >
                            {dimension.label}对比
                          </button>
                        );
                      })}
                    </div>
                  </div>
                  <div className="grid gap-3 md:grid-cols-2">
                    <FilterSelect
                      label={`${compareDimensionConfig.label} A`}
                      icon={<span className="h-2.5 w-2.5 rounded-full bg-cyan-500" />}
                      value={compareA}
                      options={compareDimensionOptions}
                      onChange={setCompareAFilter}
                      tone="light"
                      compact
                    />
                    <FilterSelect
                      label={`${compareDimensionConfig.label} B`}
                      icon={<span className="h-2.5 w-2.5 rounded-full bg-orange-500" />}
                      value={compareB}
                      options={compareDimensionOptions}
                      onChange={setCompareBFilter}
                      tone="light"
                      compact
                    />
                  </div>
                </div>
                <div className="mb-4 rounded-2xl border border-slate-100 bg-slate-50/80 p-3">
                  <div className="flex flex-wrap items-center justify-between gap-3">
                    <div>
                      <p className="text-xs font-semibold uppercase tracking-[0.2em] text-slate-500">质量维度</p>
                      <p className="mt-1 text-xs text-slate-400">支持多选，至少保留一项指标。</p>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      {COMPARE_QUALITY_METRICS.map((metric) => {
                        const selected = compareQualityMetrics.includes(metric.key);
                        return (
                          <button
                            key={metric.key}
                            type="button"
                            onClick={() => toggleCompareQualityMetric(metric.key)}
                            aria-pressed={selected}
                            className={`inline-flex items-center gap-2 rounded-full border px-3 py-2 text-xs font-medium transition ${
                              selected
                                ? 'border-slate-900 bg-slate-900 text-white shadow-sm'
                                : 'border-slate-200 bg-white text-slate-500 hover:border-slate-300 hover:text-slate-800'
                            }`}
                          >
                            <span
                              className={`h-0.5 w-5 rounded-full ${selected ? 'bg-white' : 'bg-slate-400'}`}
                              style={metric.dash ? { backgroundImage: 'linear-gradient(90deg, currentColor 55%, transparent 55%)', backgroundSize: '6px 2px' } : undefined}
                            />
                            {metric.label}
                          </button>
                        );
                      })}
                    </div>
                  </div>
                </div>
                {compareTrend.rows.length ? (
                  <div className="h-[430px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={compareTrend.rows} margin={{ top: 8, right: 12, left: 4, bottom: 0 }}>
                        <CartesianGrid vertical={false} stroke="#e2e8f0" strokeDasharray="3 3" />
                        <XAxis dataKey="date" tickLine={false} axisLine={false} fontSize={12} />
                        <YAxis tickLine={false} axisLine={false} fontSize={12} domain={[0, 1]} ticks={[0, 0.25, 0.5, 0.75, 1]} allowDataOverflow tickFormatter={(value) => `${Math.round(Number(value) * 100)}%`} />
                        <Tooltip content={<SessionCompareTrendTooltip leftLabel={compareA} rightLabel={compareB} />} />
                        {COMPARE_QUALITY_METRICS.filter((metric) => compareQualityMetrics.includes(metric.key)).flatMap((metric) => [
                          <Line
                            key={`left-${metric.key}`}
                            type="monotone"
                            dataKey={COMPARE_TREND_DATA_KEYS[metric.key].left}
                            stroke="#0891b2"
                            strokeWidth={metric.key === 'proofAccuracy' ? 3 : 2.25}
                            strokeDasharray={metric.dash}
                            dot={false}
                            connectNulls
                            name={`A · ${metric.label}`}
                          />,
                          <Line
                            key={`right-${metric.key}`}
                            type="monotone"
                            dataKey={COMPARE_TREND_DATA_KEYS[metric.key].right}
                            stroke="#f97316"
                            strokeWidth={metric.key === 'proofAccuracy' ? 3 : 2.25}
                            strokeDasharray={metric.dash}
                            dot={false}
                            connectNulls
                            name={`B · ${metric.label}`}
                          />,
                        ])}
                      </LineChart>
                    </ResponsiveContainer>
                  </div>
                ) : (
                  <div className="rounded-2xl border border-dashed border-slate-200 px-4 py-16 text-center text-sm text-slate-500">
                    当前筛选范围内暂无可对比的场次趋势。
                  </div>
                )}
              </section>
            </>
          ) : activeView === 'attribute' ? (
            <>
              <section className="mt-8 grid gap-5 xl:grid-cols-4">
                <motion.div initial={{ opacity: 0, y: 14 }} animate={{ opacity: 1, y: 0 }}>
                  <div className="flex h-full flex-col rounded-[26px] border border-slate-200 bg-white p-5 shadow-[0_14px_35px_rgba(15,23,42,0.04)]">
                    <div>
                      <p className="text-sm text-slate-500">当前属性项</p>
                      <p className="mt-3 font-display text-3xl text-slate-900">
                        {formatMultiFilterDisplay(attributeFilter, '全部属性项')}
                      </p>
                    </div>
                    <p className="mt-5 text-sm text-slate-500">可在顶部筛选区选择一个或多个属性项，再查看下方的时间、场次和批次表现。</p>
                  </div>
                </motion.div>
                <motion.div initial={{ opacity: 0, y: 14 }} animate={{ opacity: 1, y: 0 }}>
                  <StatCard title="申报次数" value={formatInteger(metrics.declarations)} hint={`${filteredRows.length.toLocaleString('zh-CN')} 条聚合记录`} tone="slate" icon={<Database size={18} />} />
                </motion.div>
                <motion.div initial={{ opacity: 0, y: 14 }} animate={{ opacity: 1, y: 0 }}>
                  <StatCard title="举证准确率" value={formatPercent(metrics.proofAccuracy)} hint="当前属性项口径" tone="emerald" icon={<ShieldCheck size={18} />} />
                </motion.div>
                <motion.div initial={{ opacity: 0, y: 14 }} animate={{ opacity: 1, y: 0 }}>
                  <StatCard title="精准通过率" value={formatPercent(metrics.exactPassRate)} hint="当前属性项口径" tone="blue" icon={<Target size={18} />} />
                </motion.div>
              </section>

              <section className="mt-8">
                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <div className="mb-6 flex items-center justify-between">
                    <div>
                      <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Attribute Trend</p>
                      <h2 className="mt-2 font-display text-2xl text-slate-900">属性项时间趋势</h2>
                      <p className="mt-2 text-sm text-slate-500">查看当前属性项在筛选区间内的举证准确率与精准通过率变化。</p>
                    </div>
                    <div className="rounded-full bg-slate-100 px-4 py-2 text-sm text-slate-600">共 {trendData.length} 天</div>
                  </div>
                  <div className="h-[320px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <AreaChart data={trendData}>
                        <defs>
                          <linearGradient id="attributeProofFill" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="5%" stopColor="#0f766e" stopOpacity={0.24} />
                            <stop offset="95%" stopColor="#0f766e" stopOpacity={0} />
                          </linearGradient>
                        </defs>
                        <CartesianGrid vertical={false} stroke="#e2e8f0" strokeDasharray="3 3" />
                        <XAxis dataKey="date" tickLine={false} axisLine={false} fontSize={12} />
                        <YAxis tickLine={false} axisLine={false} fontSize={12} domain={[0, 1]} tickFormatter={(value) => `${Math.round(value * 100)}%`} />
                        <Tooltip content={<RateTrendTooltip />} />
                        <Area type="monotone" dataKey="proofAccuracy" stroke="#0f766e" fill="url(#attributeProofFill)" strokeWidth={2.5} name="举证准确率" />
                        <Area type="monotone" dataKey="exactPassRate" stroke="#1d4ed8" fill="transparent" strokeWidth={2.5} name="精准通过率" />
                      </AreaChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </section>

              <section className="mt-8">
                <SessionShareCard
                  data={sessionShareData}
                  title="场次占比分布"
                  subtitle="按当前属性项分析筛选范围内的申报次数计算，用于观察该属性项主要集中在哪些场次。"
                  compact
                  metricsOnly
                />
              </section>

              <section className="mt-8 grid gap-8 xl:grid-cols-[minmax(0,1.35fr)_minmax(0,0.9fr)]">
                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <div className="mb-6">
                    <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Attribute Ranking</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">属性项排名参考</h2>
                    <p className="mt-2 text-sm text-slate-500">左侧看申报高频属性，右侧看模棱两可勾选贡献，减少页面纵向占用。</p>
                  </div>
                  <div className="grid gap-5 lg:grid-cols-2">
                    <div>
                      <div className="mb-3 flex items-center justify-between">
                        <span className="text-sm font-semibold text-slate-700">高频属性项</span>
                        <span className="text-xs text-slate-400">申报次数</span>
                      </div>
                      <div className="space-y-2">
                        {topAttributes.slice(0, 6).map((item, index) => {
                          const maxDeclarations = topAttributes[0]?.declarations || 1;
                          return (
                            <div key={item.attribute} className="rounded-2xl bg-slate-50 px-3 py-2.5">
                              <div className="flex items-center justify-between gap-3 text-sm">
                                <span className="min-w-0 truncate font-medium text-slate-700">{index + 1}. {item.attribute}</span>
                                <span className="shrink-0 font-semibold text-slate-900">{formatInteger(item.declarations)}</span>
                              </div>
                              <div className="mt-2 h-1.5 rounded-full bg-white">
                                <div className="h-1.5 rounded-full bg-blue-500" style={{ width: `${Math.min(100, (item.declarations / maxDeclarations) * 100)}%` }} />
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                    <div>
                      <div className="mb-3 flex items-center justify-between">
                        <span className="text-sm font-semibold text-slate-700">模棱两可勾选</span>
                        <span className="text-xs text-slate-400">次数 / 贡献</span>
                      </div>
                      {ambiguousAttributeRanking.length ? (
                        <div className="space-y-2">
                          {ambiguousAttributeRanking.slice(0, 6).map((item, index) => (
                            <div key={item.attribute} className="rounded-2xl bg-cyan-50/70 px-3 py-2.5">
                              <div className="flex items-center justify-between gap-3 text-sm">
                                <span className="min-w-0 truncate font-medium text-slate-700">{index + 1}. {item.attribute}</span>
                                <span className="shrink-0 font-semibold text-slate-900">
                                  {formatInteger(item.ambiguousPasses)} · {formatPercent(item.contribution)}
                                </span>
                              </div>
                              <div className="mt-2 flex items-center justify-between text-xs text-slate-500">
                                <span>模棱两可率 {formatPercent(item.ambiguousRate)}</span>
                                <span>申报 {formatInteger(item.declarations)}</span>
                              </div>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <div className="rounded-2xl border border-dashed border-slate-200 px-4 py-12 text-center text-sm text-slate-500">
                          当前暂无模棱两可勾选记录。
                        </div>
                      )}
                    </div>
                  </div>
                </div>

                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <div className="mb-6">
                    <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Batch Breakdown</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">批次表现</h2>
                    <p className="mt-2 text-sm text-slate-500">当前属性项在不同批次下的举证准确率。</p>
                  </div>
                  <div className="h-[320px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={attributeBatchRows} layout="vertical" margin={{ top: 0, right: 10, left: 10, bottom: 0 }}>
                        <CartesianGrid horizontal={false} stroke="#eef2f7" />
                        <XAxis type="number" domain={[0, 1]} tickFormatter={(value) => `${Math.round(value * 100)}%`} />
                        <YAxis type="category" dataKey="name" width={96} tickLine={false} axisLine={false} fontSize={12} />
                        <Tooltip content={<RateTrendTooltip />} />
                        <Bar dataKey="proofAccuracy" radius={[0, 10, 10, 0]} fill="#1d4ed8" name="举证准确率" />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </section>
            </>
          ) : activeView === 'efficiency' ? (
            <>
              <section className="mt-8 grid gap-5 md:grid-cols-2 xl:grid-cols-4">
                <StatCard title="总审核量" value={formatInteger(efficiencyMetrics.handledCount)} hint={`${formatInteger(filteredEfficiencyRows.length)} 条筛选后记录`} tone="slate" icon={<Database size={18} />} />
                <StatCard title="覆盖人员" value={formatInteger(efficiencyMetrics.employeeCount)} hint={`覆盖 ${formatInteger(efficiencyMetrics.teamCount)} 个团队`} tone="blue" icon={<Users size={18} />} />
                <StatCard title="加权审核量" value={efficiencyMetrics.weightedHandledCount.toFixed(1)} hint="按属性复杂度加权" tone="emerald" icon={<Target size={18} />} />
                <StatCard title="举证准确率" value={formatPercent(efficiencyMetrics.proofAccuracy)} hint={`模糊 ${formatInteger(efficiencyMetrics.ambiguousCount)} / 举证拒绝 ${formatInteger(efficiencyMetrics.proofRefusalCount)}`} tone="amber" icon={<ShieldCheck size={18} />} />
              </section>

              <section className="mt-8 rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                <div className="mb-6 flex flex-wrap items-end justify-between gap-3">
                  <div>
                    <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Workplace Quality</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">职场团队质量对比</h2>
                    <p className="mt-2 text-sm text-slate-500">常州包含老人、新人；上海包含所有批次。新增团队日均人均加权审核量：当日加权审核量 / 当日人数。</p>
                  </div>
                  <span className="rounded-full bg-slate-100 px-3 py-1.5 text-xs text-slate-600">
                    {formatInteger(workplaceEfficiency.length)} 个职场
                  </span>
                </div>
                {workplaceEfficiency.length ? (
                  <div className="grid gap-4 md:grid-cols-2">
                    {workplaceEfficiency.map((item) => (
                      <div key={item.workplace} className="rounded-3xl border border-slate-100 bg-slate-50/80 p-5">
                        <div className="flex flex-wrap items-start justify-between gap-3">
                          <div>
                            <p className="font-display text-2xl text-slate-900">{item.workplace}</p>
                            <p className="mt-1 text-xs leading-5 text-slate-500">
                              {formatInteger(item.employeeCount)} 人 · {formatInteger(item.teamCount)} 个团队 · {formatInteger(item.handledCount)} 总审核
                            </p>
                          </div>
                          <div className="grid gap-2 text-right">
                            <span className="rounded-full bg-white px-3 py-1.5 text-xs font-medium text-slate-500">
                              日均人均加权 {item.avgDailyWeightedHandledPerEmployee.toFixed(1)}
                            </span>
                            <span className="text-xs text-slate-400">
                              {formatInteger(item.activeDayCount)} 天 · 总加权 {item.weightedHandledCount.toFixed(1)}
                            </span>
                          </div>
                        </div>
                        <div className="mt-5 grid gap-3 sm:grid-cols-2">
                          <div className="rounded-2xl bg-white p-4">
                            <p className="text-xs uppercase tracking-[0.18em] text-slate-400">Proof Accuracy</p>
                            <p className="mt-2 font-display text-3xl text-emerald-700">{formatPercent(item.proofAccuracy)}</p>
                            <div className="mt-3 h-2 rounded-full bg-emerald-100">
                              <div className="h-2 rounded-full bg-emerald-500" style={{ width: `${Math.min(item.proofAccuracy * 100, 100)}%` }} />
                            </div>
                            <p className="mt-2 text-xs text-slate-500">举证准确率</p>
                          </div>
                          <div className="rounded-2xl bg-white p-4">
                            <p className="text-xs uppercase tracking-[0.18em] text-slate-400">Precision Pass</p>
                            <p className="mt-2 font-display text-3xl text-blue-700">{formatPercent(item.precisionPassRate)}</p>
                            <div className="mt-3 h-2 rounded-full bg-blue-100">
                              <div className="h-2 rounded-full bg-blue-500" style={{ width: `${Math.min(item.precisionPassRate * 100, 100)}%` }} />
                            </div>
                            <p className="mt-2 text-xs text-slate-500">精准通过率</p>
                          </div>
                          <div className="rounded-2xl bg-white p-4 sm:col-span-2">
                            <p className="text-xs uppercase tracking-[0.18em] text-slate-400">Daily Weighted Volume</p>
                            <p className="mt-2 font-display text-3xl text-cyan-700">{item.avgDailyWeightedHandledPerEmployee.toFixed(1)}</p>
                            <p className="mt-2 text-xs text-slate-500">团队日平均加权审核量：按天计算加权审核量 / 当日人数后取平均</p>
                          </div>
                        </div>
                        <p className="mt-4 line-clamp-2 text-xs leading-5 text-slate-400">
                          包含团队：{item.teams.join('、') || '未分组'}
                        </p>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="rounded-2xl border border-dashed border-slate-200 px-4 py-10 text-center text-sm text-slate-500">
                    暂无可展示的职场团队数据。
                  </div>
                )}
              </section>

              <section className="mt-8 grid gap-8 xl:grid-cols-[1.15fr_0.85fr]">
                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <div className="mb-6 flex flex-wrap items-end justify-between gap-4">
                    <div>
                      <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Employee Efficiency</p>
                      <h2 className="mt-2 font-display text-2xl text-slate-900">个人审核效率</h2>
                      <p className="mt-2 text-sm text-slate-500">选择某位员工，查看当前时间范围内的日均值与趋势变化。</p>
                    </div>
                    <div className="grid w-full gap-3 md:w-auto md:grid-cols-[220px_160px]">
                      <label className="block">
                        <span className="mb-2 block text-xs font-medium uppercase tracking-[0.18em] text-slate-400">员工</span>
                        <select
                          value={efficiencyEmployeeFilter}
                          onChange={(event) => setEfficiencyEmployeeFilter(event.target.value)}
                          className="dashboard-select h-12 w-full rounded-2xl border border-slate-200 bg-white px-4 text-sm font-medium text-slate-700 outline-none transition focus:border-blue-300 focus:ring-4 focus:ring-blue-100"
                        >
                          {efficiencyEmployeeOptions.map((employee) => (
                            <option key={employee} value={employee}>
                              {employee === ALL_OPTION ? '自动选择员工' : employee}
                            </option>
                          ))}
                        </select>
                      </label>
                      <label className="block">
                        <span className="mb-2 block text-xs font-medium uppercase tracking-[0.18em] text-slate-400">时间维度</span>
                        <select
                          value={efficiencyTimeDimension}
                          onChange={(event) => setEfficiencyTimeDimension(event.target.value as EfficiencyTimeDimension)}
                          className="dashboard-select h-12 w-full rounded-2xl border border-slate-200 bg-white px-4 text-sm font-medium text-slate-700 outline-none transition focus:border-blue-300 focus:ring-4 focus:ring-blue-100"
                        >
                          <option value="day">按日</option>
                          <option value="week">按周</option>
                          <option value="month">按月</option>
                        </select>
                      </label>
                    </div>
                  </div>
                  {employeeEfficiencyDetail && selectedEfficiencyEmployee ? (
                    <>
                      <div className="grid gap-4 md:grid-cols-2">
                        <div className="rounded-3xl bg-slate-50 p-5">
                          <p className="text-sm text-slate-500">当前员工</p>
                          <p className="mt-2 font-display text-3xl text-slate-900">{selectedEfficiencyEmployee}</p>
                          <p className="mt-2 text-xs text-slate-400">{formatInteger(employeeEfficiencyDetail.activeDayCount)} 个有审核日期</p>
                        </div>
                        <div className="rounded-3xl bg-emerald-50 p-5">
                          <p className="text-sm text-emerald-700">个人审核效率</p>
                          <p className="mt-2 font-display text-3xl text-emerald-900">{formatPercent(employeeEfficiencyDetail.totals.precisionPassRate)}</p>
                          <p className="mt-2 text-xs text-emerald-600">举证准确率 {formatPercent(employeeEfficiencyDetail.totals.proofAccuracy)}</p>
                        </div>
                        <div className="rounded-3xl bg-blue-50 p-5">
                          <p className="text-sm text-blue-700">日均总审核量</p>
                          <p className="mt-2 font-display text-3xl text-blue-900">{employeeEfficiencyDetail.dailyHandledAverage.toFixed(1)}</p>
                          <p className="mt-2 text-xs text-blue-600">总计 {formatInteger(employeeEfficiencyDetail.totals.handledCount)}</p>
                        </div>
                        <div className="rounded-3xl bg-cyan-50 p-5">
                          <p className="text-sm text-cyan-700">日均加权审核量</p>
                          <p className="mt-2 font-display text-3xl text-cyan-900">{employeeEfficiencyDetail.dailyWeightedAverage.toFixed(1)}</p>
                          <p className="mt-2 text-xs text-cyan-600">总计 {employeeEfficiencyDetail.totals.weightedHandledCount.toFixed(1)}</p>
                        </div>
                      </div>
                      <div className="mt-6 grid gap-6 2xl:grid-cols-2">
                        <div className="rounded-3xl border border-slate-100 bg-slate-50/70 p-5">
                          <p className="mb-4 text-sm font-medium text-slate-700">审核量趋势</p>
                          <div className="h-[280px]">
                            <ResponsiveContainer width="100%" height="100%">
                              <BarChart data={employeeEfficiencyDetail.trendRows}>
                                <CartesianGrid vertical={false} stroke="#e2e8f0" strokeDasharray="3 3" />
                                <XAxis dataKey="period" tickLine={false} axisLine={false} fontSize={12} />
                                <YAxis tickLine={false} axisLine={false} fontSize={12} />
                                <Tooltip formatter={(value, name) => [Number(value).toLocaleString('zh-CN'), name]} />
                                <Bar dataKey="dailyHandledAverage" fill="#1d4ed8" radius={[10, 10, 0, 0]} name="日均总审核量" />
                                <Bar dataKey="dailyWeightedAverage" fill="#06b6d4" radius={[10, 10, 0, 0]} name="日均加权审核量" />
                              </BarChart>
                            </ResponsiveContainer>
                          </div>
                        </div>
                        <div className="rounded-3xl border border-slate-100 bg-slate-50/70 p-5">
                          <p className="mb-4 text-sm font-medium text-slate-700">质量趋势</p>
                          <div className="h-[280px]">
                            <ResponsiveContainer width="100%" height="100%">
                              <LineChart data={employeeEfficiencyDetail.trendRows}>
                                <CartesianGrid vertical={false} stroke="#e2e8f0" strokeDasharray="3 3" />
                                <XAxis dataKey="period" tickLine={false} axisLine={false} fontSize={12} />
                                <YAxis tickLine={false} axisLine={false} fontSize={12} domain={[0, 1]} tickFormatter={(value) => `${Math.round(Number(value) * 100)}%`} />
                                <Tooltip formatter={(value, name) => [formatPercent(Number(value)), name]} />
                                <Line type="monotone" dataKey="precisionPassRate" stroke="#1d4ed8" strokeWidth={2.5} dot={false} name="精准通过率" />
                                <Line type="monotone" dataKey="proofAccuracy" stroke="#0f766e" strokeWidth={2.5} dot={false} name="举证准确率" />
                              </LineChart>
                            </ResponsiveContainer>
                          </div>
                        </div>
                      </div>
                    </>
                  ) : (
                    <div className="rounded-2xl border border-dashed border-slate-200 px-4 py-10 text-center text-sm text-slate-500">
                      当前筛选范围内没有可分析的个人数据。
                    </div>
                  )}
                </div>

                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <div className="mb-6">
                    <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Employee Ranking</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">人员加权审核排行</h2>
                    <p className="mt-2 text-sm text-slate-500">按加权审核量排序，并展示个人精准通过率与举证准确率。</p>
                  </div>
                  <div className="space-y-3">
                    {efficiencyRanking.length ? (
                      efficiencyRanking.map((item, index) => (
                        <div key={item.employee} className="rounded-2xl border border-slate-100 bg-slate-50/80 px-4 py-3">
                          <div className="flex items-center justify-between gap-4">
                            <div className="min-w-0">
                              <p className="truncate font-medium text-slate-900">
                                {index + 1}. {item.employee}
                              </p>
                              <p className="mt-1 text-xs text-slate-500">
                                {item.team || '未分组'} · 精准 {formatPercent(item.precisionPassRate)} · 举证 {formatPercent(item.proofAccuracy)}
                              </p>
                            </div>
                            <div className="text-right">
                              <p className="font-display text-xl font-semibold text-slate-900">{item.weightedHandledCount.toFixed(1)}</p>
                              <p className="text-xs text-slate-500">总审核 {formatInteger(item.handledCount)}</p>
                            </div>
                          </div>
                        </div>
                      ))
                    ) : (
                      <div className="rounded-2xl border border-dashed border-slate-200 px-4 py-10 text-center text-sm text-slate-500">
                        还没有导入人效底表。请使用顶部“导入人效周数据”。
                      </div>
                    )}
                  </div>
                </div>
              </section>

            </>
          ) : activeView === 'import' ? (
            <>
              <section className="mt-8 grid gap-6 xl:grid-cols-2">
                <ImportDatasetCard
                  title="质量周数据"
                  description="用于首页、对比分析和属性项分析。字段包含日期、场次、批次、属性项、申报次数等。"
                  icon={<FileSpreadsheet size={22} />}
                  tone="slate"
                  isImporting={isImporting}
                  importLabel="追加导入质量周数据"
                  templateLabel="下载质量模板"
                  clearLabel="清空质量数据"
                  sourceName={dataset.sourceName || '尚未导入文件'}
                  importedAt={dataset.importedAt}
                  rowCount={dataset.rows.length}
                  importHistory={qualityImportHistory}
                  onImport={handleImport}
                  onDownloadTemplate={downloadTemplate}
                  onClear={clearData}
                />
                <ImportDatasetCard
                  title="人效周数据"
                  description="独立用于人效分析。支持新版审核人效字段：总审核量、加权审核量、一审审核量、精准通过量、举证拒绝量、模糊通过量等。"
                  icon={<Users size={22} />}
                  tone="blue"
                  isImporting={isEfficiencyImporting}
                  importLabel="导入人效周数据"
                  templateLabel="下载人效模板"
                  clearLabel="清空人效数据"
                  sourceName={efficiencyDataset.sourceName || '尚未导入文件'}
                  importedAt={efficiencyDataset.importedAt}
                  rowCount={efficiencyDataset.rows.length}
                  importHistory={efficiencyImportHistory}
                  onImport={handleEfficiencyImport}
                  onDownloadTemplate={downloadEfficiencyTemplate}
                  onClear={clearEfficiencyData}
                />
              </section>

              <section className="mt-8 rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                <div className="mb-5 flex flex-wrap items-end justify-between gap-3">
                  <div>
                    <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Import History</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">导入记录</h2>
                    <p className="mt-2 text-sm text-slate-500">记录质量数据与人效数据每次导入的文件、时间和解析行数。</p>
                  </div>
                  <span className="rounded-full bg-slate-100 px-3 py-1.5 text-xs text-slate-600">
                    共 {formatInteger(allImportHistory.length)} 次
                  </span>
                </div>

                <div className="space-y-3">
                  {allImportHistory.length ? (
                    visibleImportHistory.map((record) => (
                      <div key={record.id} className="grid gap-3 rounded-2xl border border-slate-100 bg-slate-50/80 px-4 py-3 md:grid-cols-[120px_minmax(0,1fr)_180px_120px] md:items-center">
                        <span
                          className={`w-fit rounded-full px-3 py-1 text-xs font-medium ${
                            record.dataType === 'quality'
                              ? 'bg-slate-900 text-white'
                              : 'bg-blue-100 text-blue-700'
                          }`}
                        >
                          {record.dataType === 'quality' ? '质量数据' : '人效数据'}
                        </span>
                        <div className="min-w-0">
                          <p className="truncate text-sm font-medium text-slate-800" title={record.sourceName}>
                            {record.sourceName}
                          </p>
                          <p className="mt-1 text-xs text-slate-400">
                            {record.dataType === 'quality' ? 'Quality import' : 'Efficiency import'}
                          </p>
                        </div>
                        <span className="text-sm text-slate-500">{formatDateTime(record.importedAt)}</span>
                        <span className="text-sm text-slate-500">
                          {record.rowCount ? `${formatInteger(record.rowCount)} 行` : '历史记录'}
                        </span>
                      </div>
                    ))
                  ) : (
                    <div className="rounded-2xl border border-dashed border-slate-200 px-4 py-10 text-center text-sm text-slate-500">
                      暂无导入记录。请先导入质量周数据或人效周数据。
                    </div>
                  )}
                </div>
                {allImportHistory.length > 6 ? (
                  <div className="mt-5 flex justify-center">
                    <button
                      type="button"
                      onClick={() => setShowAllImportHistory((current) => !current)}
                      className="inline-flex items-center gap-2 rounded-full border border-slate-200 bg-white px-4 py-2 text-sm font-medium text-slate-600 transition hover:border-slate-300 hover:bg-slate-50"
                    >
                      {showAllImportHistory
                        ? '收起历史记录'
                        : `展开全部 ${formatInteger(allImportHistory.length)} 条记录`}
                      <ChevronDown
                        size={16}
                        className={`transition ${showAllImportHistory ? 'rotate-180' : ''}`}
                      />
                    </button>
                  </div>
                ) : null}
              </section>
            </>
          ) : activeView === 'dictionary' ? (
            <>
              <section className="mt-8 grid gap-6 xl:grid-cols-[0.8fr_1.2fr]">
                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Dictionary Form</p>
                  <h2 className="mt-2 font-display text-2xl text-slate-900">
                    {editingDictionaryKey ? '修改分类规则' : '新增分类规则'}
                  </h2>
                  <p className="mt-2 text-sm leading-6 text-slate-500">
                    字典字段以你的模板为准：属性项、属性项分类。底表分类会优先被这里的字典覆盖。
                  </p>

                  <div className="mt-6 grid gap-4">
                    <label className="block">
                      <span className="mb-2 block text-xs uppercase tracking-[0.2em] text-slate-400">属性项</span>
                      <input
                        value={dictionaryDraft.propertyName}
                        onChange={(event) =>
                          setDictionaryDraft((current) => ({ ...current, propertyName: event.target.value }))
                        }
                        className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm text-slate-900 outline-none transition focus:border-slate-400"
                        placeholder="例如：屏幕外观"
                      />
                    </label>
                    <label className="block">
                      <span className="mb-2 block text-xs uppercase tracking-[0.2em] text-slate-400">属性项分类</span>
                      <select
                        value={dictionaryDraft.category}
                        onChange={(event) =>
                          setDictionaryDraft((current) => ({ ...current, category: event.target.value }))
                        }
                        className="dashboard-select w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm text-slate-900 outline-none transition focus:border-slate-400"
                      >
                        <option value="">请选择分类</option>
                        {PROPERTY_CATEGORY_OPTIONS.map((category) => (
                          <option key={category} value={category}>
                            {category}
                          </option>
                        ))}
                      </select>
                    </label>
                  </div>

                  <div className="mt-6 flex flex-wrap gap-2">
                    <button
                      type="button"
                      onClick={() => void upsertDictionaryEntry()}
                      disabled={isDictionarySaving}
                      className="inline-flex items-center gap-2 rounded-2xl bg-slate-900 px-4 py-2.5 text-sm font-medium text-white transition hover:bg-slate-800 disabled:cursor-not-allowed disabled:opacity-60"
                    >
                      <Tags size={16} />
                      {isDictionarySaving ? '保存中...' : editingDictionaryKey ? '保存修改' : '新增规则'}
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setDictionaryDraft({ propertyName: '', category: '' });
                        setEditingDictionaryKey('');
                      }}
                      className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 bg-white px-4 py-2.5 text-sm font-medium text-slate-700 transition hover:bg-slate-50"
                    >
                      取消编辑
                    </button>
                  </div>
                </div>

                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <div className="flex flex-wrap items-start justify-between gap-4">
                    <div>
                      <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Dictionary Actions</p>
                      <h2 className="mt-2 font-display text-2xl text-slate-900">字典导入与维护</h2>
                      <p className="mt-2 text-sm text-slate-500">当前共有 {formatInteger(propertyCategoryDictionary.length)} 条分类规则。</p>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      <label className="inline-flex cursor-pointer items-center gap-2 rounded-2xl bg-blue-700 px-4 py-2.5 text-sm font-medium text-white transition hover:bg-blue-800">
                        <Upload size={16} />
                        导入字典 Excel
                        <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleDictionaryImport} />
                      </label>
                      <button
                        type="button"
                        onClick={() => void resetDictionary()}
                        className="inline-flex items-center gap-2 rounded-2xl border border-amber-200 bg-amber-50 px-4 py-2.5 text-sm font-medium text-amber-700 transition hover:bg-amber-100"
                      >
                        恢复模板
                      </button>
                    </div>
                  </div>

                  <div className="mt-5 rounded-2xl border border-blue-100 bg-blue-50 px-4 py-3 text-sm leading-6 text-blue-700">
                    新导入质量底表时会按字典写入分类；历史数据展示时也会按最新字典重新归类，所以修改字典后首页分类分布会立即变化。
                  </div>
                </div>
              </section>

              <section className="mt-8 rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                <div className="mb-5 flex flex-wrap items-end justify-between gap-3">
                  <div>
                    <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Dictionary Table</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">分类字典明细</h2>
                  </div>
                  <span className="rounded-full bg-slate-100 px-3 py-1.5 text-xs text-slate-600">
                    {formatInteger(propertyCategoryDictionary.length)} 条
                  </span>
                </div>

                <div className="max-h-[520px] overflow-auto rounded-2xl border border-slate-100">
                  <table className="w-full min-w-[720px] border-collapse text-left text-sm">
                    <thead className="sticky top-0 bg-slate-50 text-xs uppercase tracking-[0.18em] text-slate-400">
                      <tr>
                        <th className="px-4 py-3 font-medium">属性项</th>
                        <th className="px-4 py-3 font-medium">属性项分类</th>
                        <th className="px-4 py-3 text-right font-medium">操作</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {propertyCategoryDictionary.map((entry) => (
                        <tr key={`${entry.propertyName}-${entry.category}`} className="bg-white">
                          <td className="px-4 py-3 font-medium text-slate-800">{entry.propertyName}</td>
                          <td className="px-4 py-3 text-slate-600">{entry.category}</td>
                          <td className="px-4 py-3">
                            <div className="flex justify-end gap-2">
                              <button
                                type="button"
                                onClick={() => editDictionaryEntry(entry)}
                                className="rounded-full bg-slate-100 px-3 py-1.5 text-xs font-medium text-slate-700 transition hover:bg-slate-200"
                              >
                                修改
                              </button>
                              <button
                                type="button"
                                onClick={() => void deleteDictionaryEntry(entry.propertyName)}
                                className="rounded-full bg-rose-50 px-3 py-1.5 text-xs font-medium text-rose-700 transition hover:bg-rose-100"
                              >
                                删除
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </section>

              <section className="mt-8 grid gap-6 xl:grid-cols-[0.8fr_1.2fr]">
                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Auditor Team Form</p>
                  <h2 className="mt-2 font-display text-2xl text-slate-900">
                    {editingAuditorTeamKey ? '修改团队规则' : '新增团队规则'}
                  </h2>
                  <p className="mt-2 text-sm leading-6 text-slate-500">
                    用于将质量底表中的审核人归属到团队，对比分析中的“团队对比”会优先使用这里的维护结果。
                  </p>

                  <div className="mt-6 grid gap-4">
                    <label className="block">
                      <span className="mb-2 block text-xs uppercase tracking-[0.2em] text-slate-400">审核人</span>
                      <input
                        value={auditorTeamDraft.auditorName}
                        onChange={(event) =>
                          setAuditorTeamDraft((current) => ({ ...current, auditorName: event.target.value }))
                        }
                        className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm text-slate-900 outline-none transition focus:border-slate-400"
                        placeholder="例如：候伟强"
                      />
                    </label>
                    <label className="block">
                      <span className="mb-2 block text-xs uppercase tracking-[0.2em] text-slate-400">团队</span>
                      <input
                        value={auditorTeamDraft.team}
                        onChange={(event) =>
                          setAuditorTeamDraft((current) => ({ ...current, team: event.target.value }))
                        }
                        className="w-full rounded-2xl border border-slate-200 bg-white px-4 py-3 text-sm text-slate-900 outline-none transition focus:border-slate-400"
                        placeholder="例如：常州_老人"
                      />
                    </label>
                  </div>

                  <div className="mt-6 flex flex-wrap gap-2">
                    <button
                      type="button"
                      onClick={() => void upsertAuditorTeamEntry()}
                      disabled={isDictionarySaving}
                      className="inline-flex items-center gap-2 rounded-2xl bg-slate-900 px-4 py-2.5 text-sm font-medium text-white transition hover:bg-slate-800 disabled:cursor-not-allowed disabled:opacity-60"
                    >
                      <Users size={16} />
                      {isDictionarySaving ? '保存中...' : editingAuditorTeamKey ? '保存修改' : '新增规则'}
                    </button>
                    <button
                      type="button"
                      onClick={() => {
                        setAuditorTeamDraft({ auditorName: '', team: '' });
                        setEditingAuditorTeamKey('');
                      }}
                      className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 bg-white px-4 py-2.5 text-sm font-medium text-slate-700 transition hover:bg-slate-50"
                    >
                      取消编辑
                    </button>
                  </div>
                </div>

                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <div className="flex flex-wrap items-start justify-between gap-4">
                    <div>
                      <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Auditor Team Actions</p>
                      <h2 className="mt-2 font-display text-2xl text-slate-900">审核人团队字典</h2>
                      <p className="mt-2 text-sm text-slate-500">当前共有 {formatInteger(auditorTeamDictionary.length)} 条团队规则。</p>
                    </div>
                    <div className="flex flex-wrap gap-2">
                      <label className="inline-flex cursor-pointer items-center gap-2 rounded-2xl bg-blue-700 px-4 py-2.5 text-sm font-medium text-white transition hover:bg-blue-800">
                        <Upload size={16} />
                        导入团队字典 Excel
                        <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleAuditorTeamImport} />
                      </label>
                      <button
                        type="button"
                        onClick={() => void resetAuditorDictionary()}
                        className="inline-flex items-center gap-2 rounded-2xl border border-amber-200 bg-amber-50 px-4 py-2.5 text-sm font-medium text-amber-700 transition hover:bg-amber-100"
                      >
                        恢复模板
                      </button>
                    </div>
                  </div>

                  <div className="mt-5 rounded-2xl border border-blue-100 bg-blue-50 px-4 py-3 text-sm leading-6 text-blue-700">
                    修改团队字典后，对比分析中的团队选项会按最新字典重新生成；新导入质量底表也会自动写入审核团队。
                  </div>
                </div>
              </section>

              <section className="mt-8 rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                <div className="mb-5 flex flex-wrap items-end justify-between gap-3">
                  <div>
                    <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Auditor Team Table</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">审核人团队明细</h2>
                  </div>
                  <span className="rounded-full bg-slate-100 px-3 py-1.5 text-xs text-slate-600">
                    {formatInteger(auditorTeamDictionary.length)} 条
                  </span>
                </div>

                <div className="max-h-[520px] overflow-auto rounded-2xl border border-slate-100">
                  <table className="w-full min-w-[720px] border-collapse text-left text-sm">
                    <thead className="sticky top-0 bg-slate-50 text-xs uppercase tracking-[0.18em] text-slate-400">
                      <tr>
                        <th className="px-4 py-3 font-medium">审核人</th>
                        <th className="px-4 py-3 font-medium">团队</th>
                        <th className="px-4 py-3 text-right font-medium">操作</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {auditorTeamDictionary.map((entry) => (
                        <tr key={`${entry.auditorName}-${entry.team}`} className="bg-white">
                          <td className="px-4 py-3 font-medium text-slate-800">{entry.auditorName}</td>
                          <td className="px-4 py-3 text-slate-600">{entry.team}</td>
                          <td className="px-4 py-3">
                            <div className="flex justify-end gap-2">
                              <button
                                type="button"
                                onClick={() => editAuditorTeamEntry(entry)}
                                className="rounded-full bg-slate-100 px-3 py-1.5 text-xs font-medium text-slate-700 transition hover:bg-slate-200"
                              >
                                修改
                              </button>
                              <button
                                type="button"
                                onClick={() => void deleteAuditorTeamEntry(entry.auditorName)}
                                className="rounded-full bg-rose-50 px-3 py-1.5 text-xs font-medium text-rose-700 transition hover:bg-rose-100"
                              >
                                删除
                              </button>
                            </div>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </section>
            </>
          ) : (
            <>
              <section className="mt-8 rounded-[28px] border border-slate-200 bg-white p-5 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                <div className="flex flex-wrap items-start justify-between gap-4">
                  <div>
                    <p className="text-sm uppercase tracking-[0.26em] text-slate-400">AI Scope</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">AI 分析范围</h2>
                    <p className="mt-2 text-sm leading-6 text-slate-500">
                      这里的日期只影响 AI 分析、规则草稿和飞书预警，不会改变首页、属性项或人效模块的筛选。
                    </p>
                  </div>
                  <span className="rounded-full bg-cyan-50 px-3 py-1.5 text-xs font-medium text-cyan-700">
                    当前范围：{aiContext.dateRange}
                  </span>
                </div>
                <div className="mt-5 grid gap-3 md:grid-cols-2 xl:grid-cols-[minmax(340px,1.8fr)_minmax(180px,0.9fr)] xl:items-end">
                  <DateRangeFilter
                    label="AI 日期区间"
                    startValue={aiStartDateFilter}
                    endValue={aiEndDateFilter}
                    options={aiDateOptions}
                    onStartChange={(value) => {
                      setAiStartDateFilter(value);
                      setModelAnalysis(null);
                    }}
                    onEndChange={(value) => {
                      setAiEndDateFilter(value);
                      setModelAnalysis(null);
                    }}
                    onClear={() => {
                      setAiStartDateFilter(ALL_OPTION);
                      setAiEndDateFilter(ALL_OPTION);
                      setModelAnalysis(null);
                    }}
                    compact
                    tone="light"
                  />
                  <WeekQuickSelect
                    label="AI 周区间"
                    value={getWeekValue(aiStartDateFilter, aiEndDateFilter, aiWeekOptions)}
                    options={aiWeekOptions}
                    onChange={(week) => {
                      setAiStartDateFilter(week.start);
                      setAiEndDateFilter(week.end);
                      setModelAnalysis(null);
                    }}
                    onClear={() => {
                      setAiStartDateFilter(ALL_OPTION);
                      setAiEndDateFilter(ALL_OPTION);
                      setModelAnalysis(null);
                    }}
                    compact
                    tone="light"
                  />
                </div>
              </section>

              <section className="mt-8 grid gap-6 xl:grid-cols-[0.9fr_1.1fr]">
                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <p className="text-sm uppercase tracking-[0.26em] text-slate-400">AI Analysis</p>
                  <h2 className="mt-2 font-display text-2xl text-slate-900">AI 分析已生成</h2>
                  <p className="mt-3 text-sm leading-7 text-slate-500">
                    默认先生成规则版分析；服务端配置 DeepSeek Key 后，可在右侧一键生成更深入的周报归因。
                  </p>
                  <div className="mt-6 rounded-3xl bg-[linear-gradient(135deg,_#0f172a_0%,_#1e3a5f_100%)] p-5 text-white">
                    <p className="text-sm text-slate-300">质量健康评分</p>
                    <div className="mt-3 flex items-end gap-3">
                      <p className="font-display text-5xl">{aiAnalysis.healthScore}</p>
                      <p className="pb-2 text-sm text-slate-300">/ 100 · {aiAnalysis.healthLevel}</p>
                    </div>
                    <div className="mt-4 h-2 rounded-full bg-white/15">
                      <div
                        className="h-2 rounded-full bg-cyan-300"
                        style={{ width: `${aiAnalysis.healthScore}%` }}
                      />
                    </div>
                  </div>
                  <div className="mt-4 grid gap-3 sm:grid-cols-2">
                    <StatCard title="质量记录" value={formatInteger(aiFilteredRows.length)} hint="当前 AI 范围内" tone="blue" icon={<Database size={18} />} />
                    <StatCard title="人效记录" value={formatInteger(aiFilteredEfficiencyRows.length)} hint="当前 AI 范围内" tone="emerald" icon={<Users size={18} />} />
                  </div>
                </div>

                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <div className="flex flex-wrap items-start justify-between gap-4">
                    <div>
                      <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Feishu Draft</p>
                      <h2 className="mt-2 font-display text-2xl text-slate-900">飞书周报草稿</h2>
                      <p className="mt-2 text-sm text-slate-500">可直接复制到飞书，再人工微调措辞。</p>
                    </div>
                    <button
                      type="button"
                      onClick={() => void navigator.clipboard?.writeText(aiAnalysis.report)}
                      className="inline-flex items-center gap-2 rounded-2xl bg-slate-900 px-4 py-2.5 text-sm font-medium text-white transition hover:bg-slate-800"
                    >
                      <ClipboardList size={16} />
                      复制草稿
                    </button>
                  </div>
                  <pre className="mt-5 max-h-[360px] overflow-auto whitespace-pre-wrap rounded-2xl bg-slate-950 p-4 text-sm leading-7 text-slate-100">
                    {aiAnalysis.report}
                  </pre>
                </div>
              </section>

              <section className="mt-8 rounded-[28px] border border-cyan-100 bg-[linear-gradient(135deg,_#f8fdff_0%,_#ffffff_55%,_#eef8ff_100%)] p-6 shadow-[0_18px_45px_rgba(14,116,144,0.08)]">
                <div className="flex flex-wrap items-start justify-between gap-4">
                  <div>
                    <p className="text-sm uppercase tracking-[0.26em] text-cyan-500">DeepSeek Analysis</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">大模型深度分析</h2>
                    <p className="mt-2 max-w-3xl text-sm leading-7 text-slate-500">
                      仅发送当前筛选后的聚合报告和基础上下文，不上传原始明细。适合生成更像“人写的”飞书周报、归因和下周动作。
                    </p>
                  </div>
                  <div className="flex flex-col gap-3 sm:min-w-[320px]">
                    <label className="text-xs uppercase tracking-[0.2em] text-slate-400">模型选择</label>
                    <div className="grid gap-3 sm:grid-cols-[1fr_auto]">
                      <select
                        value={selectedDeepseekModel}
                        onChange={(event) => {
                          setSelectedDeepseekModel(event.target.value);
                          setModelAnalysisError('');
                        }}
                        disabled={isModelAnalyzing}
                        className="dashboard-select h-12 rounded-2xl border border-cyan-100 bg-white px-4 text-sm font-medium text-slate-700 outline-none transition focus:border-cyan-300 focus:ring-4 focus:ring-cyan-100"
                      >
                        {DEEPSEEK_MODEL_OPTIONS.map((model) => (
                          <option key={model} value={model}>
                            {model}
                          </option>
                        ))}
                        <option value={CUSTOM_MODEL_OPTION}>{CUSTOM_MODEL_OPTION}</option>
                      </select>
                      <button
                        type="button"
                        onClick={() => void runModelAnalysis()}
                        disabled={isModelAnalyzing || !activeDeepseekModel}
                        className="inline-flex h-12 items-center justify-center gap-2 rounded-2xl bg-cyan-500 px-4 text-sm font-medium text-white shadow-[0_12px_28px_rgba(6,182,212,0.24)] transition hover:bg-cyan-600 disabled:cursor-not-allowed disabled:bg-slate-300 disabled:shadow-none"
                      >
                        <BrainCircuit size={16} />
                        {isModelAnalyzing ? '等待返回...' : '调用分析'}
                      </button>
                    </div>
                    {selectedDeepseekModel === CUSTOM_MODEL_OPTION ? (
                      <input
                        value={customDeepseekModel}
                        onChange={(event) => {
                          setCustomDeepseekModel(event.target.value);
                          setModelAnalysisError('');
                        }}
                        placeholder="输入模型名，例如 deepseek-v4-flash"
                        className="h-12 rounded-2xl border border-cyan-100 bg-white px-4 text-sm font-medium text-slate-700 outline-none transition placeholder:text-slate-300 focus:border-cyan-300 focus:ring-4 focus:ring-cyan-100"
                      />
                    ) : null}
                    <p className="text-xs text-slate-400">
                      {isModelAnalyzing
                        ? '正在等待模型返回，最多 75 秒后自动恢复。'
                        : `本次将使用：${activeDeepseekModel || '未选择'}`}
                    </p>
                  </div>
                </div>
                {modelAnalysisError ? (
                  <div className="mt-5 rounded-2xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm leading-6 text-amber-700">
                    {modelAnalysisError}
                  </div>
                ) : null}
                {modelAnalysis ? (
                  <div className="mt-5">
                    <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
                      <div>
                        <p className="text-xs uppercase tracking-[0.22em] text-cyan-500">Generated Result</p>
                        <h3 className="mt-1 font-display text-xl text-slate-900">
                          DeepSeek 深度分析 · {modelAnalysis.model}
                        </h3>
                      </div>
                      <span className="rounded-full bg-white px-3 py-1.5 text-xs text-slate-500 shadow-sm">
                        生成时间：{formatDateTime(modelAnalysis.generatedAt)}
                      </span>
                    </div>
                    <div className="flex justify-end">
                      <button
                        type="button"
                        onClick={() =>
                          void navigator.clipboard?.writeText(
                            [
                              `模型：${modelAnalysis.model}`,
                              `生成时间：${formatDateTime(modelAnalysis.generatedAt)}`,
                              '',
                              modelAnalysis.analysis,
                            ].join('\n'),
                          )
                        }
                        className="mb-3 inline-flex items-center gap-2 rounded-2xl border border-cyan-100 bg-white px-4 py-2 text-sm font-medium text-cyan-700 transition hover:bg-cyan-50"
                      >
                        <ClipboardList size={16} />
                        复制深度分析
                      </button>
                    </div>
                    <pre className="max-h-[460px] overflow-auto whitespace-pre-wrap rounded-2xl border border-cyan-100 bg-white/80 p-4 text-sm leading-7 text-slate-700">
                      {modelAnalysis.analysis}
                    </pre>
                  </div>
                ) : (
                  <div className="mt-5 grid gap-3 md:grid-cols-3">
                    <div className="rounded-2xl bg-white/75 px-4 py-3 text-sm text-slate-600">输入：规则版 AI 草稿</div>
                    <div className="rounded-2xl bg-white/75 px-4 py-3 text-sm text-slate-600">范围：{aiContext.dateRange}</div>
                    <div className="rounded-2xl bg-white/75 px-4 py-3 text-sm text-slate-600">过滤：{aiContext.filters}</div>
                  </div>
                )}
              </section>

              <section className="mt-8 rounded-[28px] border border-sky-100 bg-[linear-gradient(135deg,_#f8fbff_0%,_#ffffff_55%,_#eff6ff_100%)] p-6 shadow-[0_18px_45px_rgba(37,99,235,0.07)]">
                <div className="flex flex-wrap items-start justify-between gap-4">
                  <div>
                    <p className="text-sm uppercase tracking-[0.26em] text-sky-500">Feishu Alert Push</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">飞书每日预警推送</h2>
                    <p className="mt-2 max-w-3xl text-sm leading-7 text-slate-500">
                      嵌入现有每日预警脚本，可先生成预览；确认后再推送到已配置的飞书机器人。
                    </p>
                  </div>
                  <div className="flex flex-col gap-2 sm:min-w-[320px]">
                    <label className="text-xs uppercase tracking-[0.2em] text-slate-400">预警日期</label>
                    <input
                      type="date"
                      value={dailyAlertDate}
                      max={latestDataDate || undefined}
                      onChange={(event) => {
                        setDailyAlertDate(event.target.value);
                        setDailyAlertError('');
                      }}
                      className="h-12 rounded-2xl border border-sky-100 bg-white px-4 text-sm font-medium text-slate-700 outline-none transition focus:border-sky-300 focus:ring-4 focus:ring-sky-100"
                    />
                    <p className="text-xs text-slate-400">
                      留空则使用最新有质量数据日期；当前最新数据：{latestDataDate || '暂无'}。
                    </p>
                  </div>
                </div>

                <div className="mt-5 flex flex-wrap gap-3">
                  <button
                    type="button"
                    onClick={() => void runDailyAlert(false)}
                    disabled={isDailyAlertRunning}
                    className="inline-flex h-12 items-center justify-center gap-2 rounded-2xl border border-sky-100 bg-white px-4 text-sm font-medium text-sky-700 transition hover:bg-sky-50 disabled:cursor-not-allowed disabled:text-slate-300"
                  >
                    <ClipboardList size={16} />
                    {isDailyAlertRunning ? '生成中...' : '生成预览'}
                  </button>
                  <button
                    type="button"
                    onClick={() => void runDailyAlert(true)}
                    disabled={isDailyAlertRunning}
                    className="inline-flex h-12 items-center justify-center gap-2 rounded-2xl bg-slate-900 px-4 text-sm font-medium text-white shadow-[0_12px_28px_rgba(15,23,42,0.18)] transition hover:bg-slate-800 disabled:cursor-not-allowed disabled:bg-slate-300 disabled:shadow-none"
                  >
                    <ArrowUpRight size={16} />
                    {isDailyAlertRunning ? '推送中...' : '推送飞书'}
                  </button>
                  {dailyAlertOutput ? (
                    <button
                      type="button"
                      onClick={() => void navigator.clipboard?.writeText(dailyAlertOutput)}
                      className="inline-flex h-12 items-center justify-center gap-2 rounded-2xl border border-slate-200 bg-white px-4 text-sm font-medium text-slate-700 transition hover:bg-slate-50"
                    >
                      <ClipboardList size={16} />
                      复制结果
                    </button>
                  ) : null}
                </div>

                {dailyAlertError ? (
                  <div className="mt-5 rounded-2xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm leading-6 text-amber-700">
                    {dailyAlertError}
                  </div>
                ) : null}

                {dailyAlertMeta ? (
                  <div className="mt-5 flex flex-wrap items-center gap-2 rounded-2xl border border-sky-100 bg-white/75 px-4 py-3 text-xs text-slate-500">
                    <span className="rounded-full bg-sky-50 px-3 py-1 font-medium text-sky-700">
                      日期：{dailyAlertMeta.targetDate || dailyAlertDate || '最新有数日'}
                    </span>
                    <span className="rounded-full bg-slate-100 px-3 py-1">
                      预警数：{formatInteger(dailyAlertMeta.alertCount ?? 0)}
                    </span>
                    <span
                      className={`rounded-full px-3 py-1 font-medium ${
                        dailyAlertMeta.feishu?.skipped === false
                          ? 'bg-emerald-50 text-emerald-700'
                          : 'bg-amber-50 text-amber-700'
                      }`}
                    >
                      {dailyAlertMeta.feishu?.skipped === false
                        ? `飞书已推送${dailyAlertMeta.feishu.status ? ` · HTTP ${dailyAlertMeta.feishu.status}` : ''}`
                        : `未推送飞书${dailyAlertMeta.feishu?.reason ? ` · ${dailyAlertMeta.feishu.reason}` : ''}`}
                    </span>
                  </div>
                ) : null}

                <pre className="mt-5 max-h-[420px] overflow-auto whitespace-pre-wrap rounded-2xl border border-sky-100 bg-white/85 p-4 text-sm leading-7 text-slate-700">
                  {dailyAlertOutput || '点击“生成预览”后，这里会展示即将发送到飞书的每日预警文本。'}
                </pre>
              </section>

              <section className="mt-8 grid gap-6 xl:grid-cols-4">
                <AiInsightColumn title="核心结论" items={aiAnalysis.summary} tone="blue" />
                <AiInsightColumn title="关键驱动" items={aiAnalysis.drivers} tone="slate" />
                <AiInsightColumn title="风险提醒" items={aiAnalysis.risks} tone="amber" />
                <AiInsightColumn title="建议动作" items={aiAnalysis.actions} tone="emerald" />
              </section>

              <section className="mt-8 rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Score Breakdown</p>
                <h2 className="mt-2 font-display text-2xl text-slate-900">健康度扣分明细</h2>
                <div className="mt-5 grid gap-3 md:grid-cols-2">
                  {aiAnalysis.healthFactors.map((factor) => (
                    <div key={factor} className="rounded-2xl border border-slate-100 bg-slate-50 px-4 py-3 text-sm leading-6 text-slate-600">
                      {factor}
                    </div>
                  ))}
                </div>
              </section>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

function SidebarNavItem({
  icon,
  label,
  active,
  onClick,
}: {
  icon: React.ReactNode;
  label: string;
  active: boolean;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={`flex w-full items-center gap-3 rounded-2xl px-4 py-4 text-left transition ${
        active
          ? 'bg-[linear-gradient(90deg,_rgba(0,216,255,0.22)_0%,_rgba(0,216,255,0.08)_100%)] text-cyan-300 shadow-[inset_4px_0_0_#22d3ee]'
          : 'text-slate-400 hover:bg-white/5 hover:text-white'
      }`}
    >
      <span>{icon}</span>
      <span className="text-lg font-medium">{label}</span>
    </button>
  );
}

function MobileNavChip({
  label,
  active,
  onClick,
}: {
  label: string;
  active: boolean;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={`rounded-full px-4 py-2 text-sm font-medium transition ${
        active ? 'bg-slate-900 text-white' : 'bg-white text-slate-600'
      }`}
    >
      {label}
    </button>
  );
}

function ImportDatasetCard({
  title,
  description,
  icon,
  tone,
  isImporting,
  importLabel,
  templateLabel,
  clearLabel,
  sourceName,
  importedAt,
  rowCount,
  importHistory,
  onImport,
  onDownloadTemplate,
  onClear,
}: {
  title: string;
  description: string;
  icon: React.ReactNode;
  tone: 'slate' | 'blue';
  isImporting: boolean;
  importLabel: string;
  templateLabel: string;
  clearLabel: string;
  sourceName: string;
  importedAt: string;
  rowCount: number;
  importHistory: ImportRecord[];
  onImport: (event: ChangeEvent<HTMLInputElement>) => void;
  onDownloadTemplate: () => void;
  onClear: () => void;
}) {
  const accentClass =
    tone === 'blue'
      ? 'bg-blue-700 text-white hover:bg-blue-800'
      : 'bg-slate-900 text-white hover:bg-slate-800';
  const latestImport = [...importHistory].sort((a, b) => b.importedAt.localeCompare(a.importedAt))[0];
  const latestSourceName = latestImport?.sourceName || sourceName.split(' + ').at(-1) || sourceName;
  const latestImportedAt = latestImport?.importedAt || importedAt;
  const latestRowCount = latestImport?.rowCount || 0;
  const totalImportedRows = importHistory.reduce((sum, record) => sum + (record.rowCount || 0), 0);

  return (
    <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
      <div className="flex items-start gap-4">
        <div
          className={`flex h-12 w-12 shrink-0 items-center justify-center rounded-2xl ${
            tone === 'blue' ? 'bg-blue-50 text-blue-700' : 'bg-slate-100 text-slate-900'
          }`}
        >
          {icon}
        </div>
        <div className="min-w-0">
          <h2 className="font-display text-2xl text-slate-900">{title}</h2>
          <p className="mt-2 text-sm leading-6 text-slate-500">{description}</p>
        </div>
      </div>

      <div className="mt-6 rounded-3xl border border-slate-200 bg-slate-50/70 p-4">
        <div className="flex items-start gap-3">
          <div className="mt-0.5 rounded-2xl bg-white p-2 text-slate-500 shadow-sm">
            <FileSpreadsheet size={16} />
          </div>
          <div className="min-w-0 flex-1">
            <p className="text-xs font-medium uppercase tracking-[0.18em] text-slate-400">Latest file</p>
            <p className="mt-1 truncate text-sm font-semibold text-slate-800" title={latestSourceName}>
              {latestSourceName}
            </p>
          </div>
        </div>
        <div className="mt-4 grid gap-2 sm:grid-cols-3">
          <MiniImportMetric label="导入批次" value={`${formatInteger(importHistory.length)} 次`} />
          <MiniImportMetric label="最新时间" value={formatDateTime(latestImportedAt)} />
          <MiniImportMetric label="当前记录" value={`${formatInteger(rowCount)} 条`} />
        </div>
        {latestRowCount ? (
          <p className="mt-3 text-xs text-slate-400">
            最近解析 {formatInteger(latestRowCount)} 行
            {totalImportedRows ? ` · 累计解析 ${formatInteger(totalImportedRows)} 行` : ''}
          </p>
        ) : null}
      </div>

      <div className="mt-6 flex flex-wrap gap-2">
        <label className={`inline-flex cursor-pointer items-center gap-2 rounded-2xl px-4 py-2.5 text-sm font-medium transition ${accentClass}`}>
          <Upload size={16} />
          {isImporting ? '导入中...' : importLabel}
          <input type="file" accept=".xlsx,.xls" className="hidden" onChange={onImport} />
        </label>
        <button
          type="button"
          onClick={onDownloadTemplate}
          className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 bg-white px-4 py-2.5 text-sm font-medium text-slate-700 transition hover:border-slate-300 hover:bg-slate-50"
        >
          <HardDriveDownload size={16} />
          {templateLabel}
        </button>
        <button
          type="button"
          onClick={onClear}
          className="inline-flex items-center gap-2 rounded-2xl border border-rose-200 bg-rose-50 px-4 py-2.5 text-sm font-medium text-rose-700 transition hover:bg-rose-100"
        >
          <Trash2 size={16} />
          {clearLabel}
        </button>
      </div>
    </div>
  );
}

function AiInsightColumn({
  title,
  items,
  tone,
}: {
  title: string;
  items: string[];
  tone: 'blue' | 'amber' | 'emerald' | 'slate';
}) {
  const toneClass =
    tone === 'blue'
      ? 'border-blue-100 bg-blue-50/70 text-blue-700'
      : tone === 'amber'
        ? 'border-amber-100 bg-amber-50/70 text-amber-700'
        : tone === 'emerald'
          ? 'border-emerald-100 bg-emerald-50/70 text-emerald-700'
          : 'border-slate-100 bg-slate-50/80 text-slate-700';

  return (
    <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
      <h3 className="font-display text-xl text-slate-900">{title}</h3>
      <div className="mt-5 space-y-3">
        {items.map((item, index) => (
          <div key={`${title}-${item}`} className={`rounded-2xl border px-4 py-3 text-sm leading-6 ${toneClass}`}>
            <span className="mr-2 font-semibold">{index + 1}.</span>
            {item}
          </div>
        ))}
      </div>
    </div>
  );
}

function SessionShareCard({
  data,
  title,
  subtitle,
  compact = false,
  metricsOnly = false,
}: {
  data: ReturnType<typeof aggregateSessionShares>;
  title: string;
  subtitle: string;
  compact?: boolean;
  metricsOnly?: boolean;
}) {
  const totalDeclarations = data.reduce((sum, item) => sum + item.declarations, 0);
  const topSession = data[0];
  const visibleData = metricsOnly ? data.filter((item) => item.value > 0.01) : data;

  return (
    <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
      <div className={`${compact ? 'mb-4' : 'mb-6'} flex flex-wrap items-start justify-between gap-4`}>
        <div>
          <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Session Mix</p>
          <h2 className="mt-2 font-display text-2xl text-slate-900">{title}</h2>
          <p className="mt-2 text-sm leading-6 text-slate-500">{subtitle}</p>
        </div>
        {!metricsOnly && <div className="rounded-2xl bg-slate-50 px-4 py-3 text-right">
          <p className="text-xs uppercase tracking-[0.18em] text-slate-400">申报合计</p>
          <p className="mt-1 font-display text-2xl text-slate-900">{formatInteger(totalDeclarations)}</p>
          <p className="mt-1 text-xs text-slate-500">
            Top：{topSession ? `${topSession.name} ${formatPercent(topSession.value)}` : '暂无'}
          </p>
        </div>}
      </div>

      {visibleData.length ? (
        <div className={metricsOnly ? 'grid grid-cols-2 gap-2' : `${compact ? 'grid gap-4 lg:grid-cols-[minmax(180px,0.8fr)_minmax(0,1.2fr)]' : 'grid gap-6 lg:grid-cols-[minmax(260px,0.95fr)_minmax(0,1.05fr)]'} lg:items-center`}>
          {!metricsOnly && <div className={`${compact ? 'h-[220px]' : 'h-[300px]'} min-w-0`}>
            <ResponsiveContainer width="100%" height="100%">
              <PieChart margin={{ top: 8, right: 8, bottom: 8, left: 8 }}>
                <Pie
                  data={data}
                  dataKey="value"
                  nameKey="name"
                  cx="50%"
                  cy="50%"
                  innerRadius={compact ? 46 : 66}
                  outerRadius={compact ? 82 : 112}
                  paddingAngle={2}
                >
                  {data.map((item, index) => (
                    <Cell key={item.name} fill={SESSION_COLORS[index % SESSION_COLORS.length]} />
                  ))}
                </Pie>
                <Tooltip formatter={(value: number, _name, item) => [formatPercent(value), item.payload.name]} />
              </PieChart>
            </ResponsiveContainer>
          </div>}

          <div className={metricsOnly ? 'contents' : compact ? 'space-y-2' : 'grid gap-3 md:grid-cols-2'}>
            {visibleData.map((item, index) => (
              <div key={item.name} className={`${compact ? 'px-3 py-2.5' : 'border border-slate-100 px-4 py-3'} rounded-2xl bg-slate-50/80`}>
                <div className="flex items-center justify-between gap-3">
                  <div className="flex min-w-0 items-center gap-2">
                    <span
                      className="h-2.5 w-2.5 shrink-0 rounded-full"
                      style={{ backgroundColor: SESSION_COLORS[index % SESSION_COLORS.length] }}
                    />
                    <span className="truncate text-sm font-medium text-slate-700">{item.name}</span>
                  </div>
                  <span className="shrink-0 font-display text-base font-semibold text-slate-900">
                    {formatPercent(item.value)}
                  </span>
                </div>
                {!metricsOnly && <div className={`${compact ? 'mt-2' : 'mt-3'} h-2 rounded-full bg-white`}>
                  <div
                    className="h-2 rounded-full"
                    style={{
                      width: `${Math.min(item.value * 100, 100)}%`,
                      backgroundColor: SESSION_COLORS[index % SESSION_COLORS.length],
                    }}
                  />
                </div>}
                {metricsOnly ? (
                  <div className="mt-2 space-y-1 text-[11px] leading-4 text-slate-500">
                    <div className="flex items-center justify-between gap-2">
                      <span>举证</span>
                      <span className="font-medium text-slate-700">{formatPercent(item.proofAccuracy)}</span>
                    </div>
                    <div className="flex items-center justify-between gap-2">
                      <span>通过</span>
                      <span className="font-medium text-slate-700">{formatPercent(item.exactPassRate)}</span>
                    </div>
                  </div>
                ) : (
                  <p className="mt-2 text-xs text-slate-500">申报 {formatInteger(item.declarations)}</p>
                )}
              </div>
            ))}
          </div>
        </div>
      ) : (
        <div className="rounded-2xl border border-dashed border-slate-200 px-4 py-10 text-center text-sm text-slate-500">
          当前筛选范围内暂无场次数据。
        </div>
      )}
    </div>
  );
}

function FilterSelect({
  label,
  value,
  options,
  onChange,
  icon,
  tone = 'dark',
  compact = false,
}: {
  label: string;
  value: string;
  options: string[];
  onChange: (value: string) => void;
  icon: React.ReactNode;
  tone?: 'dark' | 'light';
  compact?: boolean;
}) {
  return (
    <label className="block">
      <span
        className={`${compact ? 'mb-1.5' : 'mb-2'} flex items-center gap-2 text-xs uppercase tracking-[0.2em] ${
          tone === 'dark' ? 'text-slate-400' : 'text-slate-500'
        }`}
      >
        {icon}
        {label}
      </span>
      <select
        value={value}
        onChange={(event) => onChange(event.target.value)}
        className={`dashboard-select w-full text-sm outline-none transition ${compact ? 'rounded-xl px-3 py-2' : 'rounded-2xl px-4 py-2.5'} ${
          tone === 'dark'
            ? 'border border-white/10 bg-white/8 text-white focus:border-white/40'
            : 'border border-slate-200 bg-white text-slate-900 focus:border-slate-400'
        }`}
      >
        {options.map((option) => (
          <option key={option} value={option}>
            {option}
          </option>
        ))}
      </select>
    </label>
  );
}

function MultiFilterSelect({
  label,
  value,
  options,
  onChange,
  icon,
}: {
  label: string;
  value: string[];
  options: string[];
  onChange: (value: string[]) => void;
  icon: React.ReactNode;
}) {
  const containerRef = useRef<HTMLDivElement | null>(null);
  const [isOpen, setIsOpen] = useState(false);
  const selectableOptions = options.filter((option) => option !== ALL_OPTION);
  const summary =
    value.length === 0
      ? ALL_OPTION
      : value.length <= 2
        ? value.join('、')
        : `已选 ${value.length} 项`;

  useEffect(() => {
    if (!isOpen) return;

    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [isOpen]);

  const toggleOption = (option: string) => {
    onChange(value.includes(option) ? value.filter((item) => item !== option) : [...value, option]);
  };

  return (
    <div ref={containerRef} className="relative">
      <span className="mb-2 flex items-center gap-2 text-xs uppercase tracking-[0.2em] text-slate-400">
        {icon}
        {label}
        {value.length ? (
          <span className="rounded-full bg-cyan-400/15 px-2 py-0.5 text-[10px] font-semibold text-cyan-200">
            {value.length}
          </span>
        ) : null}
      </span>
      <button
        type="button"
        aria-expanded={isOpen}
        onClick={() => setIsOpen((current) => !current)}
        className="flex w-full items-center justify-between gap-3 rounded-2xl border border-white/10 bg-white/8 px-4 py-2.5 text-left text-sm text-white outline-none transition hover:border-white/25 focus:border-white/40"
      >
        <span className="truncate">{summary}</span>
        <ChevronDown size={16} className={`shrink-0 transition ${isOpen ? 'rotate-180' : ''}`} />
      </button>
      {isOpen ? (
        <div className="absolute left-0 top-[calc(100%+8px)] z-[90] w-full min-w-[220px] overflow-hidden rounded-2xl border border-slate-200 bg-white p-2 text-slate-700 shadow-[0_20px_50px_rgba(15,23,42,0.22)]">
          <div className="flex items-center justify-between gap-3 border-b border-slate-100 px-2 pb-2">
            <span className="text-xs font-medium text-slate-500">{value.length ? `已选择 ${value.length} 项` : '当前为全部'}</span>
            {value.length ? (
              <button type="button" onClick={() => onChange([])} className="text-xs font-medium text-cyan-700 hover:text-cyan-900">
                清空
              </button>
            ) : null}
          </div>
          <div className="mt-1 max-h-64 overflow-y-auto">
            {selectableOptions.map((option) => {
              const selected = value.includes(option);
              return (
                <button
                  key={option}
                  type="button"
                  onClick={() => toggleOption(option)}
                  className={`flex w-full items-center gap-3 rounded-xl px-3 py-2.5 text-left text-sm transition ${
                    selected ? 'bg-cyan-50 text-cyan-900' : 'hover:bg-slate-50'
                  }`}
                >
                  <span
                    className={`flex h-5 w-5 shrink-0 items-center justify-center rounded-md border ${
                      selected ? 'border-cyan-500 bg-cyan-500 text-white' : 'border-slate-300 bg-white'
                    }`}
                  >
                    {selected ? <Check size={13} strokeWidth={3} /> : null}
                  </span>
                  <span className="truncate">{option}</span>
                </button>
              );
            })}
          </div>
        </div>
      ) : null}
    </div>
  );
}

function WeekQuickSelect({
  label,
  value,
  options,
  onChange,
  onClear,
  compact = false,
  tone = 'dark',
}: {
  label: string;
  value: string;
  options: WeekOption[];
  onChange: (week: WeekOption) => void;
  onClear: () => void;
  compact?: boolean;
  tone?: 'dark' | 'light';
}) {
  const isLight = tone === 'light';

  return (
    <label className="block">
      <span className={`${compact ? 'mb-1.5' : 'mb-2'} flex items-center gap-2 text-xs uppercase tracking-[0.2em] ${isLight ? 'text-slate-500' : 'text-slate-400'}`}>
        <CalendarDays size={compact ? 14 : 16} />
        {label}
      </span>
      <select
        value={value}
        onChange={(event) => {
          const nextValue = event.target.value;
          if (nextValue === ALL_OPTION) {
            onClear();
            return;
          }

          const selectedWeek = options.find((week) => week.value === nextValue);
          if (selectedWeek) {
            onChange(selectedWeek);
          }
        }}
        className={`dashboard-select w-full text-sm outline-none transition ${compact ? 'rounded-xl px-3 py-2' : 'rounded-2xl px-4 py-2.5'} ${
          isLight
            ? 'border border-slate-200 bg-white text-slate-900 focus:border-sky-300 focus:ring-4 focus:ring-sky-100'
            : 'border border-white/10 bg-white/8 text-white focus:border-white/40'
        }`}
      >
        <option value={ALL_OPTION}>
          全部周
        </option>
        {options.map((week) => (
          <option key={week.value} value={week.value}>
            {week.label}
          </option>
        ))}
      </select>
    </label>
  );
}

function ComparePeriodPanel({
  title,
  tone,
  startValue,
  endValue,
  dateOptions,
  weekValue,
  weekOptions,
  onStartChange,
  onEndChange,
  onWeekChange,
  onClear,
}: {
  title: string;
  tone: 'emerald' | 'blue';
  startValue: string;
  endValue: string;
  dateOptions: string[];
  weekValue: string;
  weekOptions: WeekOption[];
  onStartChange: (value: string) => void;
  onEndChange: (value: string) => void;
  onWeekChange: (week: WeekOption) => void;
  onClear: () => void;
}) {
  const toneClass =
    tone === 'emerald'
      ? 'border-emerald-300/40 bg-emerald-400/10 text-emerald-200'
      : 'border-blue-300/40 bg-blue-400/10 text-blue-200';

  return (
    <div className={`rounded-[22px] border p-4 ${toneClass}`}>
      <div className="mb-3 flex items-center justify-between gap-3">
        <div className="flex items-center gap-2 text-sm font-semibold">
          <span className={`h-2.5 w-2.5 rounded-full ${tone === 'emerald' ? 'bg-emerald-300' : 'bg-blue-300'}`} />
          {title}
        </div>
        <button
          type="button"
          onClick={onClear}
          className="rounded-full border border-white/10 px-3 py-1 text-xs text-slate-200 transition hover:bg-white/10"
        >
          清空
        </button>
      </div>
      <div className="grid gap-3 lg:grid-cols-[minmax(280px,1.4fr)_minmax(180px,0.9fr)]">
        <DateRangeFilter
          label="日期区间"
          startValue={startValue}
          endValue={endValue}
          options={dateOptions}
          onStartChange={onStartChange}
          onEndChange={onEndChange}
          onClear={onClear}
          compact
        />
        <WeekQuickSelect
          label="周区间"
          value={weekValue}
          options={weekOptions}
          onChange={onWeekChange}
          onClear={onClear}
        />
      </div>
    </div>
  );
}

function DateRangeFilter({
  label = '日期区间',
  startValue,
  endValue,
  options,
  onStartChange,
  onEndChange,
  onClear,
  compact = false,
  tone = 'dark',
}: {
  label?: string;
  startValue: string;
  endValue: string;
  options: string[];
  onStartChange: (value: string) => void;
  onEndChange: (value: string) => void;
  onClear: () => void;
  compact?: boolean;
  tone?: 'dark' | 'light';
}) {
  const hasValue = startValue !== ALL_OPTION || endValue !== ALL_OPTION;
  const isLight = tone === 'light';
  const containerRef = useRef<HTMLDivElement | null>(null);
  const [isOpen, setIsOpen] = useState(false);
  const [draftRange, setDraftRange] = useState<DateRange | undefined>(undefined);
  const [visibleMonth, setVisibleMonth] = useState<Date | undefined>(undefined);
  const availableDates = options.filter((option) => option !== ALL_OPTION);
  const minDate = availableDates[0] ? parseISO(availableDates[0]) : undefined;
  const maxDate = availableDates.length ? parseISO(availableDates[availableDates.length - 1]) : undefined;
  const visibleStartMonth = minDate ? startOfMonth(subMonths(minDate, 1)) : undefined;
  const visibleEndMonth = maxDate ? startOfMonth(addMonths(maxDate, 1)) : undefined;
  const selectedRange: DateRange | undefined =
    startValue !== ALL_OPTION || endValue !== ALL_OPTION
      ? {
          from: parseDateValue(startValue),
          to: endValue !== ALL_OPTION ? parseDateValue(endValue) : undefined,
        }
      : undefined;

  useEffect(() => {
    setDraftRange(selectedRange);
  }, [startValue, endValue]);

  useEffect(() => {
    if (!isOpen) {
      setVisibleMonth(selectedRange?.from ?? minDate);
    }
  }, [isOpen, minDate, selectedRange]);

  useEffect(() => {
    if (!isOpen) {
      return;
    }

    const handleClickOutside = (event: MouseEvent) => {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setIsOpen(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, [isOpen]);

  const handleSelect = (range: DateRange | undefined, triggerDate: Date) => {
    if (!range?.from) {
      onClear();
      return;
    }

    if (!draftRange?.from || draftRange.to) {
      const nextRange = { from: triggerDate, to: undefined };
      setDraftRange(nextRange);
      setVisibleMonth(triggerDate);
      onStartChange(format(triggerDate, 'yyyy-MM-dd'));
      onEndChange(ALL_OPTION);
      return;
    }

    const draftStart = draftRange.from;
    const triggerValue = format(triggerDate, 'yyyy-MM-dd');
    const draftStartValue = format(draftStart, 'yyyy-MM-dd');
    const normalizedRange =
      triggerValue === draftStartValue
        ? { from: draftStart, to: draftStart }
        : triggerDate < draftStart
          ? { from: triggerDate, to: draftStart }
          : { from: draftStart, to: triggerDate };

    setDraftRange(normalizedRange);
    setVisibleMonth(normalizedRange.to ?? normalizedRange.from);

    const nextStart = format(normalizedRange.from, 'yyyy-MM-dd');
    const nextEnd = normalizedRange.to ? format(normalizedRange.to, 'yyyy-MM-dd') : ALL_OPTION;

    onStartChange(nextStart);
    onEndChange(nextEnd);

    if (normalizedRange.from && normalizedRange.to) {
      setIsOpen(false);
    }
  };

  return (
    <div className={`relative ${compact ? '' : 'md:col-span-2 xl:col-span-2'}`} ref={containerRef}>
      <span className={`${compact ? 'mb-1.5' : 'mb-2'} flex items-center gap-2 text-xs uppercase tracking-[0.2em] ${isLight ? 'text-slate-500' : 'text-slate-400'}`}>
        <CalendarDays size={compact ? 14 : 16} />
        {label}
      </span>
      <div
        role="button"
        tabIndex={0}
        onClick={() => {
          setVisibleMonth(selectedRange?.from ?? minDate);
          setIsOpen((value) => !value);
        }}
        onKeyDown={(event) => {
          if (event.key === 'Enter' || event.key === ' ') {
            event.preventDefault();
            setVisibleMonth(selectedRange?.from ?? minDate);
            setIsOpen((value) => !value);
          }
        }}
        className={`flex w-full items-center text-left outline-none transition ${compact ? 'rounded-xl px-3 py-2' : 'rounded-2xl px-5 py-3'} ${
          isLight
            ? 'border border-slate-200 bg-white text-slate-900 shadow-[0_8px_24px_rgba(15,23,42,0.05)] hover:border-sky-300 focus:border-sky-300 focus:ring-4 focus:ring-sky-100'
            : compact
              ? 'border border-white/10 bg-white/[0.08] text-white shadow-inner shadow-white/[0.03] hover:border-white/30'
              : 'border border-[#4B8DFF] bg-white text-slate-800 shadow-[0_8px_24px_rgba(37,99,235,0.12)] hover:border-[#2f7cff]'
        }`}
      >
        <span className={`min-w-0 flex-1 ${compact ? 'text-sm' : 'text-[15px]'} font-medium ${isLight || !compact ? 'text-slate-800' : 'text-white'}`}>
          {formatDateDisplay(startValue) || '开始日期'}
        </span>
        <span className={`${compact ? 'mx-2' : 'mx-4'} shrink-0 text-slate-400`}>→</span>
        <span className={`min-w-0 flex-1 ${compact ? 'text-sm' : 'text-[15px]'} font-medium ${isLight || !compact ? 'text-slate-800' : 'text-white'}`}>
          {formatDateDisplay(endValue) || '结束日期'}
        </span>
        <button
          type="button"
          onClick={(event) => {
            event.stopPropagation();
            onClear();
          }}
          className={`${compact ? 'ml-2 h-6 w-6' : 'ml-3 h-7 w-7'} flex items-center justify-center rounded-full transition ${
            isLight
              ? hasValue
                ? 'bg-slate-200 text-slate-500 hover:bg-slate-300'
                : 'bg-slate-100 text-slate-300'
              : compact
              ? hasValue
                ? 'bg-white/20 text-white hover:bg-white/30'
                : 'bg-white/10 text-slate-400'
              : hasValue
                ? 'bg-slate-300 text-white hover:bg-slate-400'
                : 'bg-slate-100 text-slate-300'
          }`}
          aria-label="清空日期区间"
        >
          <X size={14} />
        </button>
      </div>

      {isOpen ? (
        <div className="absolute left-0 top-[calc(100%+12px)] z-[80] w-max max-w-[calc(100vw-3rem)] overflow-visible rounded-[28px] border border-slate-200 bg-white p-4 shadow-[0_24px_64px_rgba(15,23,42,0.18)]">
          <DayPicker
            locale={zhCN}
            mode="range"
            numberOfMonths={2}
            pagedNavigation
            weekStartsOn={0}
            month={visibleMonth}
            onMonthChange={setVisibleMonth}
            selected={draftRange}
            onSelect={handleSelect}
            startMonth={visibleStartMonth}
            endMonth={visibleEndMonth}
            showOutsideDays
            disabled={[
              ...(minDate ? [{ before: minDate }] : []),
              ...(maxDate ? [{ after: maxDate }] : []),
            ]}
            className="date-range-picker"
            formatters={{
              formatCaption: (date) => format(date, 'yyyy年 M月', { locale: zhCN }),
              formatWeekdayName: (date) => format(date, 'EEEEE', { locale: zhCN }),
            }}
            classNames={{
              months: 'rdp-months',
              month: 'rdp-month',
              month_caption: 'rdp-month-caption',
              caption_label: 'rdp-caption-label',
              nav: 'rdp-nav',
              button_previous: 'rdp-nav-button',
              button_next: 'rdp-nav-button',
              weekdays: 'rdp-weekdays',
              weekday: 'rdp-weekday',
              week: 'rdp-week',
              day: 'rdp-day',
              day_button: 'rdp-day-button',
              selected: 'rdp-selected',
              range_start: 'rdp-range-start',
              range_middle: 'rdp-range-middle',
              range_end: 'rdp-range-end',
              outside: 'rdp-outside',
              today: 'rdp-today',
              disabled: 'rdp-disabled',
            }}
            components={{
              Chevron: ({ orientation }) =>
                orientation === 'left' ? <ChevronLeft size={18} /> : <ChevronRight size={18} />,
            }}
          />
        </div>
      ) : null}
    </div>
  );
}

function MiniImportMetric({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-2xl bg-white px-3 py-2 shadow-sm">
      <p className="text-[11px] font-medium text-slate-400">{label}</p>
      <p className="mt-1 truncate text-sm font-semibold text-slate-800" title={value}>
        {value}
      </p>
    </div>
  );
}

function ToolbarChip({ icon, label, wide = false }: { icon: React.ReactNode; label: string; wide?: boolean }) {
  return (
    <div
      className={`inline-flex items-start gap-2 rounded-[18px] border border-slate-200 bg-slate-50 px-3 py-1.5 ${
        wide ? 'max-w-full sm:max-w-[520px]' : ''
      }`}
      title={label}
    >
      <span className="mt-0.5 shrink-0 text-slate-400">{icon}</span>
      <span className={wide ? 'whitespace-normal break-all leading-5' : 'max-w-[240px] truncate'}>{label}</span>
    </div>
  );
}

function CompareCard({
  title,
  leftValue,
  rightValue,
  deltaValue,
  formatter,
}: {
  title: string;
  leftValue: string;
  rightValue: string;
  deltaValue: number;
  formatter: (value: number) => string;
}) {
  return (
    <div className="flex h-full flex-col rounded-[26px] border border-slate-200 bg-white p-5 shadow-[0_14px_35px_rgba(15,23,42,0.04)]">
      <p className="text-sm text-slate-500">{title}</p>
      <div className="mt-4 grid gap-3">
        <div className="rounded-2xl bg-slate-50 px-4 py-3">
          <p className="text-xs uppercase tracking-[0.2em] text-slate-400">区间 A</p>
          <p className="mt-2 font-display text-2xl text-slate-900">{leftValue}</p>
        </div>
        <div className="rounded-2xl bg-slate-50 px-4 py-3">
          <p className="text-xs uppercase tracking-[0.2em] text-slate-400">区间 B</p>
          <p className="mt-2 font-display text-2xl text-slate-900">{rightValue}</p>
        </div>
      </div>
      <div className="mt-4">
        <DeltaBadge delta={deltaValue} formatter={formatter} />
      </div>
    </div>
  );
}

function CompareSummaryBlock({
  title,
  dateText,
  metrics,
  rowCount,
}: {
  title: string;
  dateText: string;
  metrics: ReturnType<typeof aggregateMetrics>;
  rowCount: number;
}) {
  return (
    <div className="flex h-full flex-col rounded-2xl border border-slate-100 bg-slate-50/90 p-4">
      <p className="text-sm font-medium text-slate-900">{title}</p>
      <p className="mt-1 text-xs text-slate-500">{dateText}</p>
      <div className="mt-4 space-y-2 text-sm text-slate-600">
        <div className="flex items-center justify-between">
          <span>申报次数</span>
          <span className="font-medium text-slate-900">{formatInteger(metrics.declarations)}</span>
        </div>
        <div className="flex items-center justify-between">
          <span>举证准确率</span>
          <span className="font-medium text-slate-900">{formatPercent(metrics.proofAccuracy)}</span>
        </div>
        <div className="flex items-center justify-between">
          <span>精准通过率</span>
          <span className="font-medium text-slate-900">{formatPercent(metrics.exactPassRate)}</span>
        </div>
        <div className="flex items-center justify-between">
          <span>记录数</span>
          <span className="font-medium text-slate-900">{formatInteger(rowCount)}</span>
        </div>
      </div>
    </div>
  );
}

function DeltaBadge({
  delta,
  formatter,
}: {
  delta: number;
  formatter: (value: number) => string;
}) {
  const isPositive = delta > 0;
  const isNegative = delta < 0;

  return (
    <div
      className={`inline-flex items-center gap-2 rounded-full px-3 py-1.5 text-sm font-medium ${
        isPositive
          ? 'bg-emerald-50 text-emerald-700'
          : isNegative
            ? 'bg-rose-50 text-rose-700'
            : 'bg-slate-100 text-slate-600'
      }`}
    >
      {isPositive ? <ArrowUpRight size={16} /> : isNegative ? <ArrowDownRight size={16} /> : <Minus size={16} />}
      差值 {formatter(Math.abs(delta))}
    </div>
  );
}

function RateTrendTooltip({
  active,
  payload,
  label,
}: {
  active?: boolean;
  payload?: Array<{ color?: string; name?: string; value?: number }>;
  label?: string;
}) {
  if (!active || !payload?.length) {
    return null;
  }

  return (
    <div className="rounded-2xl border border-slate-200 bg-white/95 px-4 py-3 shadow-xl shadow-slate-900/10 backdrop-blur">
      <p className="text-sm font-medium text-slate-900">{label}</p>
      <div className="mt-2 space-y-2">
        {payload.map((entry) => (
          <div key={entry.name} className="flex items-center justify-between gap-4 text-sm">
            <div className="flex items-center gap-2 text-slate-600">
              <span
                className="h-2.5 w-2.5 rounded-full"
                style={{ backgroundColor: entry.color ?? '#94a3b8' }}
              />
              <span>{entry.name}</span>
            </div>
            <span className="font-medium text-slate-900">{formatPercent(entry.value ?? 0)}</span>
          </div>
        ))}
      </div>
    </div>
  );
}

function SessionCompareSummaryCard({
  label,
  session,
  metrics,
  tone,
}: {
  label: string;
  session: string;
  metrics: ReturnType<typeof aggregateMetrics>;
  tone: 'cyan' | 'orange';
}) {
  const styles =
    tone === 'cyan'
      ? {
          panel: 'border-cyan-100 bg-cyan-50/70',
          badge: 'bg-cyan-500 text-white',
          accent: 'text-cyan-800',
        }
      : {
          panel: 'border-orange-100 bg-orange-50/70',
          badge: 'bg-orange-500 text-white',
          accent: 'text-orange-800',
        };

  return (
    <div className={`rounded-[28px] border p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)] ${styles.panel}`}>
      <div className="flex flex-wrap items-start justify-between gap-3">
        <div>
          <span className={`inline-flex rounded-full px-3 py-1 text-xs font-semibold ${styles.badge}`}>{label}</span>
          <h2 className="mt-3 font-display text-2xl text-slate-900">{session || '暂无可选对象'}</h2>
        </div>
        <div className="text-right">
          <p className="text-xs uppercase tracking-[0.18em] text-slate-400">申报次数</p>
          <p className={`mt-1 font-display text-3xl font-semibold ${styles.accent}`}>{formatInteger(metrics.declarations)}</p>
        </div>
      </div>
      <div className="mt-6 grid grid-cols-2 gap-3 sm:grid-cols-4">
        {[
          ['举证准确率', formatPercent(metrics.proofAccuracy)],
          ['精准通过率', formatPercent(metrics.exactPassRate)],
          ['模棱两可率', formatPercent(metrics.ambiguousRate)],
          ['拒绝率', formatPercent(metrics.rejectRate)],
        ].map(([metricLabel, value]) => (
          <div key={metricLabel} className="rounded-2xl bg-white/90 px-3 py-3">
            <p className="text-xs text-slate-500">{metricLabel}</p>
            <p className="mt-1 font-display text-xl font-semibold text-slate-900">{value}</p>
          </div>
        ))}
      </div>
    </div>
  );
}

function SessionCompareTrendTooltip({
  active,
  payload,
  label,
  leftLabel,
  rightLabel,
}: {
  active?: boolean;
  payload?: Array<{ color?: string; name?: string; value?: number | null; dataKey?: string }>;
  label?: string;
  leftLabel: string;
  rightLabel: string;
}) {
  if (!active || !payload?.length) {
    return null;
  }

  return (
    <div className="rounded-2xl border border-slate-200 bg-white/95 px-4 py-3 shadow-xl shadow-slate-900/10 backdrop-blur">
      <p className="text-sm font-medium text-slate-900">{label}</p>
      <div className="mt-2 space-y-2">
        {payload.map((entry) => {
          const session = entry.dataKey?.startsWith('left') ? leftLabel : rightLabel;
          const metric =
            COMPARE_QUALITY_METRICS.find(({ key }) =>
              Object.values(COMPARE_TREND_DATA_KEYS[key]).includes(entry.dataKey ?? ''),
            )?.label ?? entry.name ?? '质量指标';
          return (
            <div key={entry.dataKey} className="flex items-center justify-between gap-5 text-sm">
              <div className="flex items-center gap-2 text-slate-600">
                <span className="h-2.5 w-2.5 rounded-full" style={{ backgroundColor: entry.color ?? '#94a3b8' }} />
                <span>{session} · {metric}</span>
              </div>
              <span className="font-medium text-slate-900">
                {entry.value == null ? '-' : formatPercent(entry.value)}
              </span>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function CompareTrendTooltip({
  active,
  payload,
  label,
}: {
  active?: boolean;
  payload?: Array<{ color?: string; name?: string; value?: number | null; payload?: { leftDate?: string; rightDate?: string } }>;
  label?: string;
}) {
  if (!active || !payload?.length) {
    return null;
  }

  const row = payload[0]?.payload;

  return (
    <div className="rounded-2xl border border-slate-200 bg-white/95 px-4 py-3 shadow-xl shadow-slate-900/10 backdrop-blur">
      <p className="text-sm font-medium text-slate-900">{label}</p>
      <div className="mt-1 space-y-1 text-xs text-slate-500">
        <div>区间 A 日期：{row?.leftDate ?? '-'}</div>
        <div>区间 B 日期：{row?.rightDate ?? '-'}</div>
      </div>
      <div className="mt-2 space-y-2">
        {payload.map((entry) => (
          <div key={entry.name} className="flex items-center justify-between gap-4 text-sm">
            <div className="flex items-center gap-2 text-slate-600">
              <span
                className="h-2.5 w-2.5 rounded-full"
                style={{ backgroundColor: entry.color ?? '#94a3b8' }}
              />
              <span>{entry.name}</span>
            </div>
            <span className="font-medium text-slate-900">
              {entry.value == null ? '-' : formatPercent(entry.value)}
            </span>
          </div>
        ))}
      </div>
    </div>
  );
}

function InfoPill({ label, value, compact = false }: { label: string; value: string; compact?: boolean }) {
  return (
    <div className={`rounded-2xl border border-slate-200 bg-white px-4 ${compact ? 'py-3' : 'py-4'}`}>
      <p className="text-xs uppercase tracking-[0.24em] text-slate-400">{label}</p>
      <p className={`mt-2 font-medium text-slate-700 ${compact ? 'text-xs' : 'text-sm'}`}>{value}</p>
    </div>
  );
}

function StatCard({ title, value, hint, icon, tone }: MetricsCardData) {
  const toneClass = {
    slate: 'bg-slate-900 text-white',
    emerald: 'bg-emerald-600 text-white',
    blue: 'bg-blue-600 text-white',
    amber: 'bg-amber-500 text-slate-950',
    rose: 'bg-rose-500 text-white',
  }[tone];

  return (
    <div className="flex h-full flex-col rounded-[26px] border border-slate-200 bg-white p-5 shadow-[0_14px_35px_rgba(15,23,42,0.04)]">
      <div className="flex items-start justify-between gap-4">
        <div>
          <p className="text-sm text-slate-500">{title}</p>
          <p className="mt-3 font-display text-3xl text-slate-900">{value}</p>
        </div>
        <div className={`flex h-11 w-11 items-center justify-center rounded-2xl ${toneClass}`}>{icon}</div>
      </div>
      <p className="mt-5 text-sm text-slate-500">{hint}</p>
    </div>
  );
}

export default App;
