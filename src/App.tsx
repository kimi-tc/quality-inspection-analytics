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
  '第一次线审完成时间',
  '场次',
  '批次',
  '属性标签',
  '申报次数',
  '模糊通过次数',
  '未通过次数',
  '举证未通过次数',
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

const ALL_OPTION = '全部';
const AMBIGUOUS_RATE_TARGET = 0.07;
type ViewKey = 'overview' | 'compare' | 'attribute' | 'efficiency' | 'ai' | 'import' | 'dictionary';
type CompareQualityMetric = 'proofAccuracy' | 'exactPassRate' | 'ambiguousRate' | 'rejectRate';
const PROPERTY_CATEGORY_OPTIONS = ['维修项', '外观项', '功能项', 'SKU项', '其他', '售后补充项'] as const;
const CATEGORY_COLORS = ['#0f766e', '#1d4ed8', '#f59e0b', '#ef4444', '#8b5cf6', '#14b8a6'];
const SESSION_COLORS = ['#0ea5e9', '#f97316', '#10b981', '#6366f1', '#f43f5e', '#64748b', '#84cc16', '#a855f7'];
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

  const parsed = Number(String(value ?? '').replace(/,/g, '').trim());
  return Number.isFinite(parsed) ? parsed : 0;
};

const safeRateFromCounts = (numerator: number, denominator: number) => (denominator ? numerator / denominator : 0);

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
    const ranges = cellAddresses.map((address) => XLSX.utils.decode_cell(address));
    const maxRow = Math.max(...ranges.map((cell) => cell.r));
    const maxCol = Math.max(...ranges.map((cell) => cell.c));
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
    const matchedHeaders = REQUIRED_HEADERS.filter((header) => headerSet.has(header));

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

    if (
      matchedLegacyHeaders.length === REQUIRED_EFFICIENCY_HEADERS.length ||
      matchedAuditHeaders.length === REQUIRED_AUDIT_EFFICIENCY_HEADERS.length
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
      date: normalizeDate(record['第一次线审完成时间']),
      session: String(record['场次'] ?? '').trim(),
      batch: String(record['批次'] ?? '').trim(),
      category: resolveCategory(
        String(record['属性项分类'] ?? ''),
        String(record['属性标签'] ?? ''),
        propertyCategoryDictionary,
      ),
      attribute: String(record['属性标签'] ?? '').trim(),
      declarations: toNumber(record['申报次数']),
      ambiguousPasses: toNumber(record['模糊通过次数']),
      rejects: toNumber(record['未通过次数']),
      proofRejects: toNumber(record['举证未通过次数']),
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
      const firstAuditCount = toNumber(record['一审审核量']);
      const firstAuditPassCount = toNumber(record['一审通过量']);
      const precisionPassCount = toNumber(record['精准通过量']);
      const proofRefusalCount = toNumber(record['举证拒绝量']);
      const ambiguousCount = toNumber(record['模糊通过量']);
      const handledCount = toNumber(record['总审核量'] ?? record['处理单量']);

      return {
        date: normalizeDate(record['日期']),
        employee: String(record['员工姓名'] ?? '').trim(),
        team: String(record['团队'] ?? '').trim(),
        session: String(record['场次'] ?? '审核人效').trim() || '审核人效',
        batch: String(record['批次'] ?? '全部批次').trim() || '全部批次',
        handledCount,
        weightedHandledCount: toNumber(record['加权审核量'] ?? record['加权处理量']),
        firstAuditCount,
        firstAuditPassCount,
        precisionPassCount,
        auditNotPassCount: toNumber(record['未通过量']),
        proofRefusalCount,
        ambiguousCount,
        passRate: toNumber(record['通过率']) || safeRateFromCounts(firstAuditPassCount, firstAuditCount),
        precisionPassRate: toNumber(record['精准通过率']) || safeRateFromCounts(precisionPassCount, firstAuditCount),
        proofAccuracy:
          toNumber(record['举证准确率']) ||
          safeRateFromCounts(firstAuditCount - ambiguousCount - proofRefusalCount, firstAuditCount),
        avgHandleMinutes: toNumber(record['平均处理时长']),
        timeoutCount: toNumber(record['超时次数']),
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

const formatPercent = (value: number) => `${(value * 100).toFixed(2)}%`;
const formatInteger = (value: number) => value.toLocaleString('zh-CN');
const formatDateDisplay = (value: string) => (value === ALL_OPTION ? '' : value);
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
  attribute: string;
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
    const attributeMatch = criteria.attribute === ALL_OPTION || row.attribute === criteria.attribute;

    return startMatch && endMatch && sessionMatch && batchMatch && attributeMatch;
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
) =>
  filterRowsByDateRange(rows, startDate, endDate).filter(
    (row) => team === ALL_OPTION || row.team === team,
  );

const resolveWorkplace = (team: string) => {
  if (team.includes('常州')) return '常州';
  if (team.includes('上海')) return '上海';
  if (team.includes('老人') || team.includes('新人')) return '常州';
  if (team.includes('批')) return '上海';
  return '其他';
};

const aggregateMetrics = (rows: ImportedRow[]) => {
  const totals = rows.reduce(
    (acc, row) => {
      acc.declarations += row.declarations;
      acc.ambiguousPasses += row.ambiguousPasses;
      acc.rejects += row.rejects;
      acc.proofRejects += row.proofRejects;
      return acc;
    },
    { declarations: 0, ambiguousPasses: 0, rejects: 0, proofRejects: 0 },
  );

  const safeDivide = (numerator: number, denominator: number) => (denominator ? numerator / denominator : 0);

  return {
    ...totals,
    proofAccuracy: safeDivide(
      totals.declarations - totals.ambiguousPasses - totals.proofRejects,
      totals.declarations,
    ),
    exactPassRate: safeDivide(
      totals.declarations - totals.ambiguousPasses - totals.rejects,
      totals.declarations,
    ),
    ambiguousRate: safeDivide(totals.ambiguousPasses, totals.declarations),
    rejectRate: safeDivide(totals.rejects, totals.declarations),
  };
};

const aggregateTrend = (rows: ImportedRow[]) =>
  Object.values(
    rows.reduce<Record<string, { date: string; declarations: number; ambiguousPasses: number; rejects: number; proofRejects: number }>>(
      (acc, row) => {
        if (!acc[row.date]) {
          acc[row.date] = { date: row.date, declarations: 0, ambiguousPasses: 0, rejects: 0, proofRejects: 0 };
        }

        acc[row.date].declarations += row.declarations;
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
      exactPassRate: item.declarations
        ? (item.declarations - item.ambiguousPasses - item.rejects) / item.declarations
        : 0,
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

const aggregateCategories = (rows: ImportedRow[], propertyCategoryDictionary: PropertyCategoryEntry[]) =>
  Object.values(
    rows.reduce<Record<string, { name: string; value: number }>>((acc, row) => {
      const key = resolveCategory(row.category, row.attribute, propertyCategoryDictionary) || '未分类';
      if (!acc[key]) {
        acc[key] = { name: key, value: 0 };
      }

      acc[key].value += Math.max(row.declarations, row.ambiguousPasses, row.rejects, row.proofRejects);
      return acc;
    }, {}),
  )
    .map((item, _, allItems) => {
      const total = allItems.reduce((sum, current) => sum + current.value, 0);
      return {
        ...item,
        value: total ? item.value / total : 0,
      };
    })
    .sort((a, b) => b.value - a.value);

const aggregateSessionShares = (rows: ImportedRow[]) =>
  Object.values(
    rows.reduce<Record<string, { name: string; declarations: number }>>((acc, row) => {
      const key = row.session || '未识别场次';
      if (!acc[key]) {
        acc[key] = { name: key, declarations: 0 };
      }

      acc[key].declarations += row.declarations;
      return acc;
    }, {}),
  )
    .map((item, _, allItems) => {
      const total = allItems.reduce((sum, current) => sum + current.declarations, 0);
      return {
        ...item,
        value: total ? item.declarations / total : 0,
      };
    })
    .sort((a, b) => b.declarations - a.declarations);

const aggregateCompareTrend = (leftRows: ImportedRow[], rightRows: ImportedRow[], leftLabel: string, rightLabel: string) => {
  const safeRate = (numerator: number, denominator: number) => (denominator ? numerator / denominator : 0);

  const buildSeries = (rows: ImportedRow[]) => {
    const grouped = rows.reduce<
      Record<string, { declarations: number; ambiguous: number; rejects: number; proofRejects: number }>
    >((acc, row) => {
      if (!acc[row.date]) {
        acc[row.date] = { declarations: 0, ambiguous: 0, rejects: 0, proofRejects: 0 };
      }
      acc[row.date].declarations += row.declarations;
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
        exactPassRate: safeRate(item.declarations - item.ambiguous - item.rejects, item.declarations),
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

  return Object.values(
    rows.reduce<
      Record<string, { name: string; declarations: number; ambiguous: number; rejects: number; proofRejects: number }>
    >((acc, row) => {
      const name = row[key];
      if (!acc[name]) {
        acc[name] = { name, declarations: 0, ambiguous: 0, rejects: 0, proofRejects: 0 };
      }

      acc[name].declarations += row.declarations;
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
      exactPassRate: safeRate(item.declarations - item.ambiguous - item.rejects, item.declarations),
    }))
    .sort((a, b) => b.declarations - a.declarations)
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

const aggregateQualityDimension = (
  rows: ImportedRow[],
  key: 'session' | 'batch' | 'category' | 'attribute',
  propertyCategoryDictionary: PropertyCategoryEntry[] = [],
) =>
  Object.values(
    rows.reduce<Record<string, { name: string; declarations: number; ambiguousPasses: number; rejects: number; proofRejects: number }>>(
      (acc, row) => {
        const name =
          key === 'category'
            ? resolveCategory(row.category, row.attribute, propertyCategoryDictionary)
            : String(row[key] || '未分类');
        if (!acc[name]) {
          acc[name] = { name, declarations: 0, ambiguousPasses: 0, rejects: 0, proofRejects: 0 };
        }

        acc[name].declarations += row.declarations;
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
  propertyCategoryDictionary,
}: {
  qualityMetrics: ReturnType<typeof aggregateMetrics>;
  qualityRows: ImportedRow[];
  topAttributes: ReturnType<typeof aggregateAttributes>;
  categoryData: ReturnType<typeof aggregateCategories>;
  efficiencyMetrics: ReturnType<typeof aggregateEfficiency>;
  efficiencyRanking: ReturnType<typeof aggregateEfficiencyRanking>;
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
  const precisionVolatileAttributes = aggregateAttributePrecisionVolatility(qualityRows);
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
    hasEfficiencyData
      ? `人效侧已导入 ${formatInteger(efficiencyMetrics.handledCount)} 单、覆盖 ${formatInteger(efficiencyMetrics.employeeCount)} 人；本次 AI 分析仍以质量指标和属性项口径为主。`
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
    precisionVolatileAttributes.length
      ? `精准通过率波动属性项：${precisionVolatileAttributes.map((item) => `「${item.attribute}」波动${formatPercent(item.volatility)}（${formatPercent(item.minRate)}~${formatPercent(item.maxRate)}）`).join('；')}。`
      : '属性项精准通过率暂未发现明显波动，或日期样本不足。',
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
    ambiguousCategories.length
      ? `模糊口径复盘建议优先看：${ambiguousCategories.map((item) => `「${item.name}」${formatPercent(item.metrics.ambiguousRate)}`).join('、')}。`
      : '模糊通过暂无明显分类集中，可保持常规抽检。',
    highRejectAttributes.length
      ? `拒绝率专项建议聚焦：${highRejectAttributes.map((item) => `「${item.name}」拒绝率${formatPercent(item.metrics.rejectRate)}、申报${formatInteger(item.declarations)}次`).join('；')}，逐项核对拒绝原因是否来自拍摄证据、规则口径或商品描述。`
      : '拒绝率未呈现高样本属性项集中，建议保持按场次抽检。',
    precisionVolatileAttributes.length
      ? `精准通过率波动建议优先抽看：${precisionVolatileAttributes.slice(0, 3).map((item) => `「${item.attribute}」`).join('、')}，对比高低日期样本，确认是否存在口径漂移或人员判断差异。`
      : '精准通过率波动不明显，暂不需要单独建立波动专项。',
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
    highRejectAttributes[0]
      ? `拒绝率最高属性项：「${highRejectAttributes[0].name}」，拒绝率 ${formatPercent(highRejectAttributes[0].metrics.rejectRate)}，申报 ${formatInteger(highRejectAttributes[0].declarations)} 次。`
      : '暂未识别高样本高拒绝属性项。',
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
  const timeout = window.setTimeout(() => controller.abort(), 75000);
  const response = await fetch('/api/ai-analysis', {
    method: 'POST',
    signal: controller.signal,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload),
  }).finally(() => window.clearTimeout(timeout));
  const body = await response.json().catch(() => ({ message: '大模型接口返回异常，请稍后重试。' }));

  if (!response.ok) {
    throw new Error(body.message || '大模型分析生成失败');
  }

  return body;
};

function App() {
  const [dataset, setDataset] = useState<ParsedWorkbook>(emptyWorkbook);
  const [efficiencyDataset, setEfficiencyDataset] = useState<ParsedEfficiencyWorkbook>(emptyEfficiencyWorkbook);
  const [propertyCategoryDictionary, setPropertyCategoryDictionary] = useState<PropertyCategoryEntry[]>([]);
  const [dictionaryDraft, setDictionaryDraft] = useState<PropertyCategoryEntry>({ propertyName: '', category: '' });
  const [editingDictionaryKey, setEditingDictionaryKey] = useState('');
  const [activeView, setActiveView] = useState<ViewKey>('overview');
  const [overviewStartDateFilter, setOverviewStartDateFilter] = useState(ALL_OPTION);
  const [overviewEndDateFilter, setOverviewEndDateFilter] = useState(ALL_OPTION);
  const [overviewSessionFilter, setOverviewSessionFilter] = useState<string[]>([]);
  const [overviewBatchFilter, setOverviewBatchFilter] = useState<string[]>([]);
  const [attributeStartDateFilter, setAttributeStartDateFilter] = useState(ALL_OPTION);
  const [attributeEndDateFilter, setAttributeEndDateFilter] = useState(ALL_OPTION);
  const [attributeSessionFilter, setAttributeSessionFilter] = useState(ALL_OPTION);
  const [attributeBatchFilter, setAttributeBatchFilter] = useState(ALL_OPTION);
  const [attributeFilter, setAttributeFilter] = useState(ALL_OPTION);
  const [compareStartDateFilter, setCompareStartDateFilter] = useState(ALL_OPTION);
  const [compareEndDateFilter, setCompareEndDateFilter] = useState(ALL_OPTION);
  const [compareSessionAFilter, setCompareSessionAFilter] = useState('');
  const [compareSessionBFilter, setCompareSessionBFilter] = useState('');
  const [compareQualityMetrics, setCompareQualityMetrics] = useState<CompareQualityMetric[]>([
    'proofAccuracy',
    'exactPassRate',
  ]);
  const [compareBatchFilter, setCompareBatchFilter] = useState(ALL_OPTION);
  const [compareAttributeFilter, setCompareAttributeFilter] = useState(ALL_OPTION);
  const [efficiencyStartDateFilter, setEfficiencyStartDateFilter] = useState(ALL_OPTION);
  const [efficiencyEndDateFilter, setEfficiencyEndDateFilter] = useState(ALL_OPTION);
  const [efficiencyTeamFilter, setEfficiencyTeamFilter] = useState(ALL_OPTION);
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
    }),
    [dataset.rows],
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
      ),
    [efficiencyDataset.rows, efficiencyEndDateFilter, efficiencyStartDateFilter, efficiencyTeamFilter],
  );
  const efficiencyEmployeeOptions = useMemo(
    () => createStringOptions<EfficiencyRow>(filteredEfficiencyRows, (row) => row.employee),
    [filteredEfficiencyRows],
  );

  const metrics = useMemo(() => aggregateMetrics(filteredRows), [filteredRows]);
  const trendData = useMemo(() => aggregateTrend(filteredRows), [filteredRows]);
  const topAttributes = useMemo(() => aggregateAttributes(filteredRows), [filteredRows]);
  const categoryData = useMemo(
    () => aggregateCategories(filteredRows, propertyCategoryDictionary),
    [filteredRows, propertyCategoryDictionary],
  );
  const sessionShareData = useMemo(() => aggregateSessionShares(filteredRows), [filteredRows]);
  const compareSessionOptions = useMemo(
    () => aggregateSessionShares(dataset.rows).map((item) => item.name),
    [dataset.rows],
  );
  const compareSessionA = compareSessionAFilter || compareSessionOptions[0] || '';
  const compareSessionB =
    compareSessionBFilter || compareSessionOptions.find((option) => option !== compareSessionA) || compareSessionA;
  const compareBaseRows = useMemo(
    () =>
      filterRowsByCriteria(dataset.rows, {
        startDate: compareStartDateFilter,
        endDate: compareEndDateFilter,
        session: ALL_OPTION,
        batch: compareBatchFilter,
        attribute: compareAttributeFilter,
      }),
    [
      compareAttributeFilter,
      compareBatchFilter,
      compareEndDateFilter,
      compareStartDateFilter,
      dataset.rows,
    ],
  );
  const compareSessionARows = useMemo(
    () => compareBaseRows.filter((row) => row.session === compareSessionA),
    [compareBaseRows, compareSessionA],
  );
  const compareSessionBRows = useMemo(
    () => compareBaseRows.filter((row) => row.session === compareSessionB),
    [compareBaseRows, compareSessionB],
  );
  const compareSessionAMetrics = useMemo(() => aggregateMetrics(compareSessionARows), [compareSessionARows]);
  const compareSessionBMetrics = useMemo(() => aggregateMetrics(compareSessionBRows), [compareSessionBRows]);
  const compareSessionTrend = useMemo(
    () => aggregateSessionCompareTrend(compareSessionARows, compareSessionBRows, compareSessionA, compareSessionB),
    [compareSessionA, compareSessionARows, compareSessionB, compareSessionBRows],
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
  const aiAnalysis = useMemo(
    () =>
      createAiAnalysis({
        qualityMetrics: metrics,
        qualityRows: filteredRows,
        topAttributes,
        categoryData,
        efficiencyMetrics,
        efficiencyRanking,
        propertyCategoryDictionary,
      }),
    [categoryData, efficiencyMetrics, efficiencyRanking, filteredRows, metrics, propertyCategoryDictionary, topAttributes],
  );
  const aiContext = useMemo(
    () => ({
      qualityRows: filteredRows.length,
      efficiencyRows: filteredEfficiencyRows.length,
      dateRange:
        filteredRows.length > 0
          ? `${filteredRows[0].date} ~ ${filteredRows[filteredRows.length - 1].date}`
          : '暂无质量数据',
      filters: [
        `日期=${formatDateDisplay(overviewStartDateFilter) || '全部'}~${formatDateDisplay(overviewEndDateFilter) || '全部'}`,
        `场次=${overviewSessionFilter.length ? overviewSessionFilter.join('、') : ALL_OPTION}`,
        `批次=${overviewBatchFilter.length ? overviewBatchFilter.join('、') : ALL_OPTION}`,
      ].join('；'),
    }),
    [
      filteredEfficiencyRows.length,
      filteredRows,
      overviewBatchFilter,
      overviewEndDateFilter,
      overviewSessionFilter,
      overviewStartDateFilter,
    ],
  );
  const activeDeepseekModel =
    selectedDeepseekModel === CUSTOM_MODEL_OPTION
      ? customDeepseekModel.trim()
      : selectedDeepseekModel;
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
      setAttributeSessionFilter(ALL_OPTION);
      setAttributeBatchFilter(ALL_OPTION);
      setAttributeFilter(ALL_OPTION);
      setCompareStartDateFilter(ALL_OPTION);
      setCompareEndDateFilter(ALL_OPTION);
      setCompareSessionAFilter('');
      setCompareSessionBFilter('');
      setCompareBatchFilter(ALL_OPTION);
      setCompareAttributeFilter(ALL_OPTION);
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
      setAttributeSessionFilter(ALL_OPTION);
      setAttributeBatchFilter(ALL_OPTION);
      setAttributeFilter(ALL_OPTION);
      setCompareStartDateFilter(ALL_OPTION);
      setCompareEndDateFilter(ALL_OPTION);
      setCompareSessionAFilter('');
      setCompareSessionBFilter('');
      setCompareBatchFilter(ALL_OPTION);
      setCompareAttributeFilter(ALL_OPTION);
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
        err instanceof Error && err.name === 'AbortError'
          ? '大模型分析等待超过 75 秒，已自动停止。请稍后重试，或切换更快的模型。'
          : err instanceof Error
            ? err.message
            : '大模型分析生成失败。',
      );
    } finally {
      setIsModelAnalyzing(false);
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
          <div className="p-5 lg:p-6">
            <div className="rounded-[28px] border border-slate-200 bg-white/90 p-4 shadow-[0_12px_35px_rgba(15,23,42,0.04)]">
              <div className="flex flex-col gap-4 xl:flex-row xl:items-center xl:justify-between">
                <div className="flex min-w-0 items-center gap-3">
                  <div className="flex h-12 w-12 shrink-0 items-center justify-center rounded-2xl bg-slate-900 text-white">
                    <Layers3 size={22} />
                  </div>
                  <div className="min-w-0">
                    <div className="flex flex-wrap items-center gap-x-3 gap-y-1">
                      <h1 className="font-display text-2xl font-semibold text-slate-900">预质检质量看板</h1>
                      <span className="rounded-full bg-emerald-50 px-3 py-1 text-xs font-medium text-emerald-700">
                        {isLoading ? '加载中' : '每周复用'}
                      </span>
                    </div>
                    <p className="mt-1 text-sm text-slate-500">
                      聚焦举证准确率、精准通过率、模棱两可率与拒绝率，按场次、批次、属性项动态拆解。
                    </p>
                  </div>
                </div>
                <div className="grid gap-2 sm:grid-cols-2 xl:min-w-[420px]">
                  <div className="rounded-2xl border border-emerald-100 bg-emerald-50/80 px-4 py-3">
                    <div className="flex items-center gap-2 text-xs font-semibold uppercase tracking-[0.16em] text-emerald-700">
                      <CalendarDays size={14} />
                      最新数据
                    </div>
                    <p className="mt-1 font-display text-xl font-semibold text-slate-950">
                      {latestDataDate || '暂无数据'}
                    </p>
                  </div>
                  <div className="rounded-2xl border border-slate-200 bg-slate-50/90 px-4 py-3">
                    <p className="text-xs font-semibold uppercase tracking-[0.16em] text-slate-400">可查询区间</p>
                    <p className="mt-1 text-sm font-semibold text-slate-800">
                      {overallCoverage.start && overallCoverage.end
                        ? `${overallCoverage.start} ~ ${overallCoverage.end}`
                        : '暂无可查询数据'}
                    </p>
                    <p className="mt-1 text-xs text-slate-500">
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

            <div className="mt-4 rounded-[28px] bg-[linear-gradient(135deg,_#12212d_0%,_#182b39_100%)] p-4 text-white">
              <div className="flex flex-col gap-4">
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
                    />
                  </div>
                ) : activeView === 'efficiency' ? (
                  <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-[minmax(340px,1.8fr)_minmax(180px,0.9fr)_minmax(170px,0.8fr)_minmax(180px,0.8fr)] xl:items-end">
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
                    />
                    <FilterSelect
                      label="团队"
                      icon={<Users size={16} />}
                      value={efficiencyTeamFilter}
                      options={efficiencyTeamOptions}
                      onChange={setEfficiencyTeamFilter}
                    />
                    <div className="rounded-2xl border border-white/10 bg-white/5 px-4 py-3 text-sm text-slate-200">
                      <p className="font-medium text-white">当前人效记录</p>
                      <p className="mt-1 text-xs text-slate-300">
                        {formatInteger(filteredEfficiencyRows.length)} / {formatInteger(efficiencyDataset.rows.length)} 条
                      </p>
                    </div>
                  </div>
                ) : activeView === 'ai' || activeView === 'import' || activeView === 'dictionary' ? (
                  <div className="rounded-2xl border border-white/10 bg-white/5 px-4 py-4 text-sm text-slate-200">
                    <div className="flex flex-wrap items-center justify-between gap-3">
                      <div>
                        <p className="font-medium text-white">
                          {activeView === 'ai'
                            ? 'AI 分析准备区'
                            : activeView === 'import'
                              ? '数据导入与记录区'
                              : '属性项分类字典管理区'}
                        </p>
                        <p className="mt-1 text-xs leading-6 text-slate-300">
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
                    />
                    <FilterSelect
                      label="场次"
                      icon={<Layers3 size={16} />}
                      value={attributeSessionFilter}
                      options={options.sessions}
                      onChange={setAttributeSessionFilter}
                    />
                    <FilterSelect
                      label="批次"
                      icon={<Boxes size={16} />}
                      value={attributeBatchFilter}
                      options={options.batches}
                      onChange={setAttributeBatchFilter}
                    />
                    <FilterSelect
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
                  <div className="grid gap-3 md:grid-cols-2">
                    <FilterSelect
                      label="批次"
                      icon={<Boxes size={16} />}
                      value={compareBatchFilter}
                      options={options.batches}
                      onChange={setCompareBatchFilter}
                    />
                    <FilterSelect
                      label="属性项"
                      icon={<Tags size={16} />}
                      value={compareAttributeFilter}
                      options={options.attributes}
                      onChange={setCompareAttributeFilter}
                    />
                  </div>
                ) : null}
                <details className="rounded-2xl border border-white/10 bg-white/5 px-4 py-3 text-sm text-slate-200">
                  <summary className="flex cursor-pointer list-none items-center justify-between gap-3 font-medium marker:content-none">
                    当前口径
                    <ChevronDown size={16} className="shrink-0" />
                  </summary>
                  <ul className="mt-3 space-y-2 text-xs leading-6 text-slate-300">
                    <li>举证准确率 = (申报次数 - 模糊通过次数 - 举证未通过次数) / 申报次数</li>
                    <li>精准通过率 = (申报次数 - 模糊通过次数 - 未通过次数) / 申报次数</li>
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

              <section className="mt-8 grid gap-8 xl:grid-cols-[1.25fr_0.75fr]">
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
                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <div className="mb-6">
                    <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Category Mix</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">属性项分类分布</h2>
                  </div>
                  <div className="grid min-h-[320px] gap-4 lg:grid-cols-[minmax(0,1.1fr)_minmax(180px,0.9fr)] lg:items-center">
                    <div className="h-[280px] min-w-0">
                      <ResponsiveContainer width="100%" height="100%">
                        <PieChart margin={{ top: 8, right: 8, bottom: 8, left: 8 }}>
                          <Pie
                            data={categoryData}
                            dataKey="value"
                            nameKey="name"
                            cx="50%"
                            cy="50%"
                            innerRadius={58}
                            outerRadius={102}
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
                    <div className="space-y-3">
                      {categoryData.map((item, index) => (
                        <div key={item.name} className="flex items-center justify-between gap-3 rounded-2xl bg-slate-50 px-3 py-2">
                          <div className="flex min-w-0 items-center gap-2">
                            <span
                              className="h-2.5 w-2.5 shrink-0 rounded-full"
                              style={{
                                backgroundColor: CATEGORY_COLORS[index % CATEGORY_COLORS.length],
                              }}
                            />
                            <span className="truncate text-sm text-slate-600">{item.name}</span>
                          </div>
                          <span className="shrink-0 font-display text-sm font-semibold text-slate-900">
                            {formatPercent(item.value)}
                          </span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </section>

              <section className="mt-8">
                <SessionShareCard
                  data={sessionShareData}
                  title="场次占比分布"
                  subtitle="按当前首页筛选范围内的申报次数计算，各场次占比合计为 100%。"
                />
              </section>
            </>
          ) : activeView === 'compare' ? (
            <>
              <section className="mt-8 grid gap-6 xl:grid-cols-2">
                <SessionCompareSummaryCard label="场次 A" session={compareSessionA} metrics={compareSessionAMetrics} tone="cyan" />
                <SessionCompareSummaryCard label="场次 B" session={compareSessionB} metrics={compareSessionBMetrics} tone="orange" />
              </section>

              <section className="mt-8 rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                <div className="mb-6 flex flex-wrap items-start justify-between gap-4">
                  <div>
                    <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Session Trend Compare</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">场次 A / 场次 B 质量趋势</h2>
                    <p className="mt-2 text-sm text-slate-500">在相同日期、批次和属性项条件下，按自然日对比两个场次的质量变化。</p>
                  </div>
                  <span className="rounded-full bg-slate-100 px-3 py-2 text-xs font-medium text-slate-600">
                    {formatDateDisplay(compareStartDateFilter) || '全部日期'} ~ {formatDateDisplay(compareEndDateFilter) || '全部日期'}
                  </span>
                </div>
                <div className="mb-6 grid gap-4 rounded-3xl border border-slate-100 bg-slate-50/80 p-4 md:grid-cols-2">
                  <FilterSelect
                    label="场次 A"
                    icon={<span className="h-2.5 w-2.5 rounded-full bg-cyan-500" />}
                    value={compareSessionA}
                    options={compareSessionOptions}
                    onChange={setCompareSessionAFilter}
                    tone="light"
                  />
                  <FilterSelect
                    label="场次 B"
                    icon={<span className="h-2.5 w-2.5 rounded-full bg-orange-500" />}
                    value={compareSessionB}
                    options={compareSessionOptions}
                    onChange={setCompareSessionBFilter}
                    tone="light"
                  />
                </div>
                <div className="mb-6 rounded-3xl border border-slate-100 bg-slate-50/80 p-4">
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
                {compareSessionTrend.rows.length ? (
                  <div className="h-[430px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <LineChart data={compareSessionTrend.rows} margin={{ top: 8, right: 12, left: 4, bottom: 0 }}>
                        <CartesianGrid vertical={false} stroke="#e2e8f0" strokeDasharray="3 3" />
                        <XAxis dataKey="date" tickLine={false} axisLine={false} fontSize={12} />
                        <YAxis tickLine={false} axisLine={false} fontSize={12} domain={[0, 1]} ticks={[0, 0.25, 0.5, 0.75, 1]} allowDataOverflow tickFormatter={(value) => `${Math.round(Number(value) * 100)}%`} />
                        <Tooltip content={<SessionCompareTrendTooltip leftLabel={compareSessionA} rightLabel={compareSessionB} />} />
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
                      <p className="mt-3 font-display text-3xl text-slate-900">{attributeFilter === ALL_OPTION ? '全部属性项' : attributeFilter}</p>
                    </div>
                    <p className="mt-5 text-sm text-slate-500">建议先在顶部筛选区选择一个具体属性项，再查看下方的时间、场次和批次表现。</p>
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
                <SessionShareCard
                  data={sessionShareData}
                  title="场次占比分布"
                  subtitle="按当前属性项分析筛选范围内的申报次数计算，用于观察该属性项主要集中在哪些场次。"
                />
              </section>

              <section className="mt-8 grid gap-8 xl:grid-cols-[1.15fr_0.85fr]">
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

                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <div className="mb-6">
                    <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Attribute Ranking</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">高频属性项参考</h2>
                    <p className="mt-2 text-sm text-slate-500">如果还没锁定要分析的属性项，可以先看当前筛选下申报次数靠前的属性项。</p>
                  </div>
                  <div className="h-[320px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={topAttributes} layout="vertical" margin={{ top: 0, right: 10, left: 10, bottom: 0 }}>
                        <CartesianGrid horizontal={false} stroke="#eef2f7" />
                        <XAxis type="number" hide />
                        <YAxis type="category" dataKey="attribute" width={96} tickLine={false} axisLine={false} fontSize={12} />
                        <Tooltip />
                        <Bar dataKey="declarations" radius={[0, 10, 10, 0]} fill="#1d4ed8" name="申报次数" />
                      </BarChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </section>

              <section className="mt-8 grid gap-8 xl:grid-cols-2">
                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <div className="mb-6">
                    <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Session Breakdown</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">场次表现</h2>
                    <p className="mt-2 text-sm text-slate-500">当前属性项在不同场次下的举证准确率。</p>
                  </div>
                  <div className="h-[320px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={attributeSessionRows} layout="vertical" margin={{ top: 0, right: 10, left: 10, bottom: 0 }}>
                        <CartesianGrid horizontal={false} stroke="#eef2f7" />
                        <XAxis type="number" domain={[0, 1]} tickFormatter={(value) => `${Math.round(value * 100)}%`} />
                        <YAxis type="category" dataKey="name" width={96} tickLine={false} axisLine={false} fontSize={12} />
                        <Tooltip content={<RateTrendTooltip />} />
                        <Bar dataKey="proofAccuracy" radius={[0, 10, 10, 0]} fill="#0f766e" name="举证准确率" />
                      </BarChart>
                    </ResponsiveContainer>
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
                    allImportHistory.map((record) => (
                      <div key={record.id} className="grid gap-3 rounded-2xl border border-slate-100 bg-slate-50/80 px-4 py-3 md:grid-cols-[120px_1fr_180px_120px] md:items-center">
                        <span
                          className={`w-fit rounded-full px-3 py-1 text-xs font-medium ${
                            record.dataType === 'quality'
                              ? 'bg-slate-900 text-white'
                              : 'bg-blue-100 text-blue-700'
                          }`}
                        >
                          {record.dataType === 'quality' ? '质量数据' : '人效数据'}
                        </span>
                        <span className="break-all text-sm font-medium text-slate-800">{record.sourceName}</span>
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
            </>
          ) : (
            <>
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
                    <StatCard title="质量记录" value={formatInteger(dataset.rows.length)} hint="来自质量周数据" tone="blue" icon={<Database size={18} />} />
                    <StatCard title="人效记录" value={formatInteger(efficiencyDataset.rows.length)} hint="来自人效周数据" tone="emerald" icon={<Users size={18} />} />
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
  onImport: (event: ChangeEvent<HTMLInputElement>) => void;
  onDownloadTemplate: () => void;
  onClear: () => void;
}) {
  const accentClass =
    tone === 'blue'
      ? 'bg-blue-700 text-white hover:bg-blue-800'
      : 'bg-slate-900 text-white hover:bg-slate-800';

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

      <div className="mt-5 grid gap-3 text-xs text-slate-500">
        <ToolbarChip icon={<FileSpreadsheet size={14} />} label={sourceName} wide />
        <div className="flex flex-wrap gap-2">
          <ToolbarChip icon={<CalendarDays size={14} />} label={formatDateTime(importedAt)} />
          <ToolbarChip icon={<Database size={14} />} label={`${formatInteger(rowCount)} 条记录`} />
        </div>
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
}: {
  data: ReturnType<typeof aggregateSessionShares>;
  title: string;
  subtitle: string;
}) {
  const totalDeclarations = data.reduce((sum, item) => sum + item.declarations, 0);
  const topSession = data[0];

  return (
    <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
      <div className="mb-6 flex flex-wrap items-start justify-between gap-4">
        <div>
          <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Session Mix</p>
          <h2 className="mt-2 font-display text-2xl text-slate-900">{title}</h2>
          <p className="mt-2 text-sm leading-6 text-slate-500">{subtitle}</p>
        </div>
        <div className="rounded-2xl bg-slate-50 px-4 py-3 text-right">
          <p className="text-xs uppercase tracking-[0.18em] text-slate-400">申报合计</p>
          <p className="mt-1 font-display text-2xl text-slate-900">{formatInteger(totalDeclarations)}</p>
          <p className="mt-1 text-xs text-slate-500">
            Top：{topSession ? `${topSession.name} ${formatPercent(topSession.value)}` : '暂无'}
          </p>
        </div>
      </div>

      {data.length ? (
        <div className="grid gap-6 lg:grid-cols-[minmax(260px,0.95fr)_minmax(0,1.05fr)] lg:items-center">
          <div className="h-[300px] min-w-0">
            <ResponsiveContainer width="100%" height="100%">
              <PieChart margin={{ top: 8, right: 8, bottom: 8, left: 8 }}>
                <Pie
                  data={data}
                  dataKey="value"
                  nameKey="name"
                  cx="50%"
                  cy="50%"
                  innerRadius={66}
                  outerRadius={112}
                  paddingAngle={2}
                >
                  {data.map((item, index) => (
                    <Cell key={item.name} fill={SESSION_COLORS[index % SESSION_COLORS.length]} />
                  ))}
                </Pie>
                <Tooltip formatter={(value: number, _name, item) => [formatPercent(value), item.payload.name]} />
              </PieChart>
            </ResponsiveContainer>
          </div>

          <div className="grid gap-3 md:grid-cols-2">
            {data.map((item, index) => (
              <div key={item.name} className="rounded-2xl border border-slate-100 bg-slate-50/80 px-4 py-3">
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
                <div className="mt-3 h-2 rounded-full bg-white">
                  <div
                    className="h-2 rounded-full"
                    style={{
                      width: `${Math.min(item.value * 100, 100)}%`,
                      backgroundColor: SESSION_COLORS[index % SESSION_COLORS.length],
                    }}
                  />
                </div>
                <p className="mt-2 text-xs text-slate-500">申报 {formatInteger(item.declarations)}</p>
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
}: {
  label: string;
  value: string;
  options: string[];
  onChange: (value: string) => void;
  icon: React.ReactNode;
  tone?: 'dark' | 'light';
}) {
  return (
    <label className="block">
      <span
        className={`mb-2 flex items-center gap-2 text-xs uppercase tracking-[0.2em] ${
          tone === 'dark' ? 'text-slate-400' : 'text-slate-500'
        }`}
      >
        {icon}
        {label}
      </span>
      <select
        value={value}
        onChange={(event) => onChange(event.target.value)}
        className={`dashboard-select w-full rounded-2xl px-4 py-2.5 text-sm outline-none transition ${
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
}: {
  label: string;
  value: string;
  options: WeekOption[];
  onChange: (week: WeekOption) => void;
  onClear: () => void;
}) {
  return (
    <label className="block">
      <span className="mb-2 flex items-center gap-2 text-xs uppercase tracking-[0.2em] text-slate-400">
        <CalendarDays size={16} />
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
        className="dashboard-select w-full rounded-2xl border border-white/10 bg-white/8 px-4 py-2.5 text-sm text-white outline-none transition focus:border-white/40"
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
}: {
  label?: string;
  startValue: string;
  endValue: string;
  options: string[];
  onStartChange: (value: string) => void;
  onEndChange: (value: string) => void;
  onClear: () => void;
  compact?: boolean;
}) {
  const hasValue = startValue !== ALL_OPTION || endValue !== ALL_OPTION;
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
      <span className="mb-2 flex items-center gap-2 text-xs uppercase tracking-[0.2em] text-slate-400">
        <CalendarDays size={16} />
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
        className={`flex w-full items-center rounded-2xl px-5 py-3 text-left outline-none transition ${
          compact
            ? 'border border-white/10 bg-white/[0.08] text-white shadow-inner shadow-white/[0.03] hover:border-white/30'
            : 'border border-[#4B8DFF] bg-white text-slate-800 shadow-[0_8px_24px_rgba(37,99,235,0.12)] hover:border-[#2f7cff]'
        }`}
      >
        <span className={`min-w-0 flex-1 text-[15px] font-medium ${compact ? 'text-white' : 'text-slate-800'}`}>
          {formatDateDisplay(startValue) || '开始日期'}
        </span>
        <span className="mx-4 shrink-0 text-slate-400">→</span>
        <span className={`min-w-0 flex-1 text-[15px] font-medium ${compact ? 'text-white' : 'text-slate-800'}`}>
          {formatDateDisplay(endValue) || '结束日期'}
        </span>
        <button
          type="button"
          onClick={(event) => {
            event.stopPropagation();
            onClear();
          }}
          className={`ml-3 flex h-7 w-7 items-center justify-center rounded-full transition ${
            compact
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
          <h2 className="mt-3 font-display text-2xl text-slate-900">{session || '暂无可选场次'}</h2>
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
