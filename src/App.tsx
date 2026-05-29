import React, { ChangeEvent, useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { DayPicker, type DateRange } from 'react-day-picker';
import { addDays, addMonths, format, parseISO, startOfMonth, startOfWeek, subMonths } from 'date-fns';
import { zhCN } from 'date-fns/locale';
import {
  ResponsiveContainer,
  AreaChart,
  Area,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  BarChart,
  Bar,
  PieChart,
  Pie,
  Cell,
  Legend,
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
  ArrowUpRight,
  ArrowDownRight,
  Minus,
} from 'lucide-react';
import { motion } from 'motion/react';
import { ImportedRow, MetricsCardData, ParsedWorkbook, SharedDatasetResponse } from './types';

const REQUIRED_HEADERS = [
  '第一次线审完成时间',
  '场次',
  '批次',
  '属性项分类',
  '属性标签',
  '申报次数',
  '模糊通过次数',
  '未通过次数',
  '举证未通过次数',
] as const;

const ALL_OPTION = '全部';

const emptyWorkbook: ParsedWorkbook = {
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

const pickDataSheet = (workbook: XLSX.WorkBook) => {
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
      defval: '',
      raw: false,
    });

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

const parseWorkbookFile = async (file: File): Promise<ParsedWorkbook> => {
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
      category: String(record['属性项分类'] ?? '').trim(),
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

const formatPercent = (value: number) => `${(value * 100).toFixed(2)}%`;
const formatInteger = (value: number) => value.toLocaleString('zh-CN');
const formatDateDisplay = (value: string) => (value === ALL_OPTION ? '' : value);
const parseDateValue = (value: string) => (value && value !== ALL_OPTION ? parseISO(value) : undefined);

type FilterCriteria = {
  startDate: string;
  endDate: string;
  session: string;
  batch: string;
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

const getWeekValue = (startDate: string, endDate: string, weeks: WeekOption[]) =>
  weeks.find((week) => week.start === startDate && week.end === endDate)?.value ?? ALL_OPTION;

const filterRowsByCriteria = (rows: ImportedRow[], criteria: FilterCriteria) =>
  rows.filter((row) => {
    const startMatch = criteria.startDate === ALL_OPTION || row.date >= criteria.startDate;
    const endMatch = criteria.endDate === ALL_OPTION || row.date <= criteria.endDate;
    const sessionMatch = criteria.session === ALL_OPTION || row.session === criteria.session;
    const batchMatch = criteria.batch === ALL_OPTION || row.batch === criteria.batch;
    const attributeMatch = criteria.attribute === ALL_OPTION || row.attribute === criteria.attribute;

    return startMatch && endMatch && sessionMatch && batchMatch && attributeMatch;
  });

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

const aggregateCategories = (rows: ImportedRow[]) =>
  Object.values(
    rows.reduce<Record<string, { name: string; value: number }>>((acc, row) => {
      const key = row.category || '未分类';
      if (!acc[key]) {
        acc[key] = { name: key, value: 0 };
      }

      acc[key].value += row.declarations;
      return acc;
    }, {}),
  ).sort((a, b) => b.value - a.value);

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

const downloadTemplate = () => {
  const sample = [
    {
      第一次线审完成时间: '2026-05-23',
      场次: '京东寄卖',
      批次: '第4批',
      属性项分类: '主观项',
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

function App() {
  const [dataset, setDataset] = useState<ParsedWorkbook>(emptyWorkbook);
  const [activeView, setActiveView] = useState<'overview' | 'compare' | 'attribute'>('overview');
  const [startDateFilter, setStartDateFilter] = useState(ALL_OPTION);
  const [endDateFilter, setEndDateFilter] = useState(ALL_OPTION);
  const [sessionFilter, setSessionFilter] = useState(ALL_OPTION);
  const [batchFilter, setBatchFilter] = useState(ALL_OPTION);
  const [attributeFilter, setAttributeFilter] = useState(ALL_OPTION);
  const [compareLeftStart, setCompareLeftStart] = useState(ALL_OPTION);
  const [compareLeftEnd, setCompareLeftEnd] = useState(ALL_OPTION);
  const [compareRightStart, setCompareRightStart] = useState(ALL_OPTION);
  const [compareRightEnd, setCompareRightEnd] = useState(ALL_OPTION);
  const [compareSessionSelection, setCompareSessionSelection] = useState(ALL_OPTION);
  const [compareBatchFirstSelection, setCompareBatchFirstSelection] = useState(ALL_OPTION);
  const [compareBatchSecondSelection, setCompareBatchSecondSelection] = useState(ALL_OPTION);
  const [error, setError] = useState('');
  const [isImporting, setIsImporting] = useState(false);
  const [isLoading, setIsLoading] = useState(true);

  useEffect(() => {
    const loadDataset = async () => {
      try {
        const sharedDataset = await fetchSharedDataset();
        if (Array.isArray(sharedDataset.rows)) {
          setDataset(sharedDataset);
        }
      } catch {
        setError('共享数据服务暂时不可用，请确认后端已启动。');
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

  const filteredRows = useMemo(
    () =>
      filterRowsByCriteria(dataset.rows, {
        startDate: startDateFilter,
        endDate: endDateFilter,
        session: sessionFilter,
        batch: batchFilter,
        attribute: attributeFilter,
      }),
    [attributeFilter, batchFilter, dataset.rows, endDateFilter, sessionFilter, startDateFilter],
  );

  const metrics = useMemo(() => aggregateMetrics(filteredRows), [filteredRows]);
  const trendData = useMemo(() => aggregateTrend(filteredRows), [filteredRows]);
  const topAttributes = useMemo(() => aggregateAttributes(filteredRows), [filteredRows]);
  const categoryData = useMemo(() => aggregateCategories(filteredRows), [filteredRows]);
  const compareLeftRows = useMemo(
    () =>
      filterRowsByCriteria(dataset.rows, {
        startDate: compareLeftStart,
        endDate: compareLeftEnd,
        session: sessionFilter,
        batch: batchFilter,
        attribute: attributeFilter,
      }),
    [attributeFilter, batchFilter, compareLeftEnd, compareLeftStart, dataset.rows, sessionFilter],
  );
  const compareRightRows = useMemo(
    () =>
      filterRowsByCriteria(dataset.rows, {
        startDate: compareRightStart,
        endDate: compareRightEnd,
        session: sessionFilter,
        batch: batchFilter,
        attribute: attributeFilter,
      }),
    [attributeFilter, batchFilter, compareRightEnd, compareRightStart, dataset.rows, sessionFilter],
  );
  const compareLeftMetrics = useMemo(() => aggregateMetrics(compareLeftRows), [compareLeftRows]);
  const compareRightMetrics = useMemo(() => aggregateMetrics(compareRightRows), [compareRightRows]);
  const compareTrend = useMemo(
    () =>
      aggregateCompareTrend(
        compareLeftRows,
        compareRightRows,
        `${formatDateDisplay(compareLeftStart) || '区间A'} → ${formatDateDisplay(compareLeftEnd) || '未选'}`,
        `${formatDateDisplay(compareRightStart) || '区间B'} → ${formatDateDisplay(compareRightEnd) || '未选'}`,
      ),
    [compareLeftEnd, compareLeftRows, compareLeftStart, compareRightEnd, compareRightRows, compareRightStart],
  );
  const compareSessionRows = useMemo(
    () => aggregateSessionComparison(compareLeftRows, compareRightRows),
    [compareLeftRows, compareRightRows],
  );
  const compareBatchRows = useMemo(
    () => aggregateBatchComparison(compareLeftRows, compareRightRows),
    [compareLeftRows, compareRightRows],
  );
  const compareSessionOptions = useMemo(
    () => [ALL_OPTION, ...compareSessionRows.map((item) => item.session)],
    [compareSessionRows],
  );
  const compareBatchOptions = useMemo(
    () => [ALL_OPTION, ...compareBatchRows.map((item) => item.batch)],
    [compareBatchRows],
  );
  const compareSelectedSession = useMemo(
    () =>
      compareSessionRows.find((item) => item.session === compareSessionSelection) ??
      compareSessionRows[0] ??
      null,
    [compareSessionRows, compareSessionSelection],
  );
  const compareSelectedBatchFirst = useMemo(
    () =>
      compareBatchRows.find((item) => item.batch === compareBatchFirstSelection) ??
      compareBatchRows[0] ??
      null,
    [compareBatchFirstSelection, compareBatchRows],
  );
  const compareSelectedBatchSecond = useMemo(
    () =>
      compareBatchRows.find((item) => item.batch === compareBatchSecondSelection) ??
      compareBatchRows.find((item) => item.batch !== (compareBatchFirstSelection === ALL_OPTION ? compareBatchRows[0]?.batch : compareBatchFirstSelection)) ??
      compareBatchRows[1] ??
      null,
    [compareBatchFirstSelection, compareBatchRows, compareBatchSecondSelection],
  );
  const compareBatchChartData = useMemo(() => {
    if (!compareSelectedBatchFirst || !compareSelectedBatchSecond) {
      return [];
    }

    return [
      {
        metric: '区间A',
        [compareSelectedBatchFirst.batch]: compareSelectedBatchFirst.leftProofAccuracy,
        [compareSelectedBatchSecond.batch]: compareSelectedBatchSecond.leftProofAccuracy,
      },
      {
        metric: '区间B',
        [compareSelectedBatchFirst.batch]: compareSelectedBatchFirst.rightProofAccuracy,
        [compareSelectedBatchSecond.batch]: compareSelectedBatchSecond.rightProofAccuracy,
      },
    ];
  }, [compareSelectedBatchFirst, compareSelectedBatchSecond]);
  const attributeSessionRows = useMemo(
    () => aggregateDimensionMetrics(filteredRows, 'session'),
    [filteredRows],
  );
  const attributeBatchRows = useMemo(
    () => aggregateDimensionMetrics(filteredRows, 'batch'),
    [filteredRows],
  );

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
      const parsed = await parseWorkbookFile(file);
      const nextDataset = await mergeSharedDataset(parsed);
      setDataset(nextDataset);
      setStartDateFilter(ALL_OPTION);
      setEndDateFilter(ALL_OPTION);
      setCompareLeftStart(ALL_OPTION);
      setCompareLeftEnd(ALL_OPTION);
      setCompareRightStart(ALL_OPTION);
      setCompareRightEnd(ALL_OPTION);
      setSessionFilter(ALL_OPTION);
      setBatchFilter(ALL_OPTION);
      setAttributeFilter(ALL_OPTION);
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
      setStartDateFilter(ALL_OPTION);
      setEndDateFilter(ALL_OPTION);
      setCompareLeftStart(ALL_OPTION);
      setCompareLeftEnd(ALL_OPTION);
      setCompareRightStart(ALL_OPTION);
      setCompareRightEnd(ALL_OPTION);
      setSessionFilter(ALL_OPTION);
      setBatchFilter(ALL_OPTION);
      setAttributeFilter(ALL_OPTION);
    } catch (err) {
      setError(err instanceof Error ? err.message : '清空共享数据失败。');
    }
  };

  const compareCards = [
    {
      title: '申报次数',
      leftValue: compareLeftMetrics.declarations,
      rightValue: compareRightMetrics.declarations,
      formatter: formatInteger,
    },
    {
      title: '举证准确率',
      leftValue: compareLeftMetrics.proofAccuracy,
      rightValue: compareRightMetrics.proofAccuracy,
      formatter: formatPercent,
    },
    {
      title: '精准通过率',
      leftValue: compareLeftMetrics.exactPassRate,
      rightValue: compareRightMetrics.exactPassRate,
      formatter: formatPercent,
    },
    {
      title: '模棱两可率',
      leftValue: compareLeftMetrics.ambiguousRate,
      rightValue: compareRightMetrics.ambiguousRate,
      formatter: formatPercent,
    },
    {
      title: '拒绝率',
      leftValue: compareLeftMetrics.rejectRate,
      rightValue: compareRightMetrics.rejectRate,
      formatter: formatPercent,
    },
  ];

  const compareModeReady =
    compareLeftStart !== ALL_OPTION &&
    compareLeftEnd !== ALL_OPTION &&
    compareRightStart !== ALL_OPTION &&
    compareRightEnd !== ALL_OPTION;

  useEffect(() => {
    if (!compareSessionRows.length) {
      setCompareSessionSelection(ALL_OPTION);
      return;
    }

    if (
      compareSessionSelection === ALL_OPTION ||
      !compareSessionRows.some((item) => item.session === compareSessionSelection)
    ) {
      setCompareSessionSelection(compareSessionRows[0].session);
    }
  }, [compareSessionRows, compareSessionSelection]);

  useEffect(() => {
    if (!compareBatchRows.length) {
      setCompareBatchFirstSelection(ALL_OPTION);
      setCompareBatchSecondSelection(ALL_OPTION);
      return;
    }

    const batchNames = compareBatchRows.map((item) => item.batch);

    if (
      compareBatchFirstSelection === ALL_OPTION ||
      !batchNames.includes(compareBatchFirstSelection)
    ) {
      setCompareBatchFirstSelection(batchNames[0]);
    }

    if (
      compareBatchSecondSelection === ALL_OPTION ||
      !batchNames.includes(compareBatchSecondSelection) ||
      compareBatchSecondSelection === compareBatchFirstSelection
    ) {
      setCompareBatchSecondSelection(batchNames.find((item) => item !== batchNames[0]) ?? batchNames[0]);
    }
  }, [compareBatchFirstSelection, compareBatchRows, compareBatchSecondSelection]);

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
            </div>
          </div>
        </aside>

        <div className="min-w-0 flex-1">
          <div className="mb-4 flex gap-3 lg:hidden">
            <MobileNavChip label="首页" active={activeView === 'overview'} onClick={() => setActiveView('overview')} />
            <MobileNavChip label="对比分析" active={activeView === 'compare'} onClick={() => setActiveView('compare')} />
            <MobileNavChip label="属性项分析" active={activeView === 'attribute'} onClick={() => setActiveView('attribute')} />
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
                      导入新的数仓底表后，会自动追加到历史数据池并按统一口径重算核心指标。
                    </p>
                  </div>
                </div>

                <div className="flex flex-col gap-3 xl:items-end">
                  <div className="flex flex-wrap gap-2">
                    <label className="inline-flex cursor-pointer items-center gap-2 rounded-2xl bg-slate-900 px-4 py-2.5 text-sm font-medium text-white transition hover:bg-slate-800">
                      <Upload size={16} />
                      {isImporting ? '导入中...' : '追加导入周数据'}
                      <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleImport} />
                    </label>
                    <button
                      type="button"
                      onClick={downloadTemplate}
                      className="inline-flex items-center gap-2 rounded-2xl border border-slate-200 bg-white px-4 py-2.5 text-sm font-medium text-slate-700 transition hover:border-slate-300 hover:bg-slate-50"
                    >
                      <HardDriveDownload size={16} />
                      模板
                    </button>
                    <button
                      type="button"
                      onClick={clearData}
                      className="inline-flex items-center gap-2 rounded-2xl border border-rose-200 bg-rose-50 px-4 py-2.5 text-sm font-medium text-rose-700 transition hover:bg-rose-100"
                    >
                      <Trash2 size={16} />
                      清空
                    </button>
                  </div>

                  <div className="flex flex-wrap items-center gap-2 text-xs text-slate-500">
                    <ToolbarChip icon={<FileSpreadsheet size={14} />} label={dataset.sourceName || '尚未导入文件'} />
                    <ToolbarChip
                      icon={<CalendarDays size={14} />}
                      label={dataset.importedAt ? new Date(dataset.importedAt).toLocaleString('zh-CN') : '未导入'}
                    />
                    <ToolbarChip icon={<Database size={14} />} label={`${formatInteger(dataset.rows.length)} 条记录`} />
                  </div>
                  <p className="text-xs text-slate-400">重复记录会自动去重，后续导入不会覆盖历史周数据。</p>
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
                  <div className="grid gap-4 xl:grid-cols-2">
                    <ComparePeriodPanel
                      title="区间 A"
                      tone="emerald"
                      startValue={compareLeftStart}
                      endValue={compareLeftEnd}
                      dateOptions={options.dates}
                      weekValue={getWeekValue(compareLeftStart, compareLeftEnd, weekOptions)}
                      weekOptions={weekOptions}
                      onStartChange={setCompareLeftStart}
                      onEndChange={setCompareLeftEnd}
                      onWeekChange={(week) => {
                        setCompareLeftStart(week.start);
                        setCompareLeftEnd(week.end);
                      }}
                      onClear={() => {
                        setCompareLeftStart(ALL_OPTION);
                        setCompareLeftEnd(ALL_OPTION);
                      }}
                    />
                    <ComparePeriodPanel
                      title="区间 B"
                      tone="blue"
                      startValue={compareRightStart}
                      endValue={compareRightEnd}
                      dateOptions={options.dates}
                      weekValue={getWeekValue(compareRightStart, compareRightEnd, weekOptions)}
                      weekOptions={weekOptions}
                      onStartChange={setCompareRightStart}
                      onEndChange={setCompareRightEnd}
                      onWeekChange={(week) => {
                        setCompareRightStart(week.start);
                        setCompareRightEnd(week.end);
                      }}
                      onClear={() => {
                        setCompareRightStart(ALL_OPTION);
                        setCompareRightEnd(ALL_OPTION);
                      }}
                    />
                  </div>
                ) : (
                  <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-[minmax(340px,1.8fr)_minmax(180px,0.9fr)_repeat(3,minmax(150px,1fr))] xl:items-end">
                    <DateRangeFilter
                      label="日期区间"
                      startValue={startDateFilter}
                      endValue={endDateFilter}
                      options={options.dates}
                      onStartChange={setStartDateFilter}
                      onEndChange={setEndDateFilter}
                      onClear={() => {
                        setStartDateFilter(ALL_OPTION);
                        setEndDateFilter(ALL_OPTION);
                      }}
                      compact
                    />
                    <WeekQuickSelect
                      label="周区间"
                      value={getWeekValue(startDateFilter, endDateFilter, weekOptions)}
                      options={weekOptions}
                      onChange={(week) => {
                        setStartDateFilter(week.start);
                        setEndDateFilter(week.end);
                      }}
                      onClear={() => {
                        setStartDateFilter(ALL_OPTION);
                        setEndDateFilter(ALL_OPTION);
                      }}
                    />
                    <FilterSelect
                      label="场次"
                      icon={<Layers3 size={16} />}
                      value={sessionFilter}
                      options={options.sessions}
                      onChange={setSessionFilter}
                    />
                    <FilterSelect
                      label="批次"
                      icon={<Boxes size={16} />}
                      value={batchFilter}
                      options={options.batches}
                      onChange={setBatchFilter}
                    />
                    <FilterSelect
                      label="属性项"
                      icon={<Tags size={16} />}
                      value={attributeFilter}
                      options={options.attributes}
                      onChange={setAttributeFilter}
                    />
                  </div>
                )}
                {activeView === 'compare' ? (
                  <div className="grid gap-3 md:grid-cols-3">
                    <FilterSelect
                      label="场次"
                      icon={<Layers3 size={16} />}
                      value={sessionFilter}
                      options={options.sessions}
                      onChange={setSessionFilter}
                    />
                    <FilterSelect
                      label="批次"
                      icon={<Boxes size={16} />}
                      value={batchFilter}
                      options={options.batches}
                      onChange={setBatchFilter}
                    />
                    <FilterSelect
                      label="属性项"
                      icon={<Tags size={16} />}
                      value={attributeFilter}
                      options={options.attributes}
                      onChange={setAttributeFilter}
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
                  <div className="h-[320px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <PieChart>
                        <Pie
                          data={categoryData}
                          dataKey="value"
                          nameKey="name"
                          cx="45%"
                          cy="50%"
                          innerRadius={58}
                          outerRadius={102}
                          paddingAngle={2}
                        >
                          {categoryData.map((item, index) => (
                            <Cell
                              key={item.name}
                              fill={['#0f766e', '#1d4ed8', '#f59e0b', '#ef4444', '#8b5cf6', '#14b8a6'][index % 6]}
                            />
                          ))}
                        </Pie>
                        <Tooltip formatter={(value: number) => formatInteger(value)} />
                        <Legend verticalAlign="middle" align="right" layout="vertical" iconType="circle" />
                      </PieChart>
                    </ResponsiveContainer>
                  </div>
                </div>
              </section>
            </>
          ) : activeView === 'compare' ? (
            <>
              <section className="mt-8 grid gap-5 xl:grid-cols-5">
                {compareCards.map((card, index) => (
                  <motion.div
                    key={card.title}
                    initial={{ opacity: 0, y: 14 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ delay: index * 0.05 }}
                  >
                    <CompareCard
                      title={card.title}
                      leftValue={card.formatter(card.leftValue)}
                      rightValue={card.formatter(card.rightValue)}
                      deltaValue={card.leftValue - card.rightValue}
                      formatter={card.formatter}
                    />
                  </motion.div>
                ))}
              </section>

              <section className="mt-8">
                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <div className="mb-6 flex items-center justify-between">
                    <div>
                      <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Comparison</p>
                      <h2 className="mt-2 font-display text-2xl text-slate-900">区间概览</h2>
                    </div>
                    <span className="rounded-full bg-slate-100 px-4 py-2 text-sm text-slate-600">
                      {compareModeReady ? '已选择双区间' : '请先完成 A/B 区间选择'}
                    </span>
                  </div>
                  <div className="grid gap-4 md:grid-cols-2">
                    <CompareSummaryBlock
                      title="区间 A"
                      dateText={`${formatDateDisplay(compareLeftStart) || '未选'} → ${formatDateDisplay(compareLeftEnd) || '未选'}`}
                      metrics={compareLeftMetrics}
                      rowCount={compareLeftRows.length}
                    />
                    <CompareSummaryBlock
                      title="区间 B"
                      dateText={`${formatDateDisplay(compareRightStart) || '未选'} → ${formatDateDisplay(compareRightEnd) || '未选'}`}
                      metrics={compareRightMetrics}
                      rowCount={compareRightRows.length}
                    />
                  </div>
                </div>
              </section>

              <section className="mt-8 grid gap-8 xl:grid-cols-[1.1fr_0.9fr]">
                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <div className="mb-6 flex items-center justify-between">
                    <div>
                      <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Trend Compare</p>
                      <h2 className="mt-2 font-display text-2xl text-slate-900">双区间趋势对比</h2>
                      <p className="mt-2 text-sm text-slate-500">按区间内第 N 天对齐比较举证准确率，不直接按自然日对齐。</p>
                    </div>
                    <div className="text-xs text-slate-500">举证准确率</div>
                  </div>
                  <div className="h-[340px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <AreaChart data={compareTrend.rows}>
                        <defs>
                          <linearGradient id="compareLeftFill" x1="0" y1="0" x2="0" y2="1">
                            <stop offset="5%" stopColor="#0f766e" stopOpacity={0.22} />
                            <stop offset="95%" stopColor="#0f766e" stopOpacity={0} />
                          </linearGradient>
                        </defs>
                        <CartesianGrid vertical={false} stroke="#e2e8f0" strokeDasharray="3 3" />
                        <XAxis dataKey="dayLabel" tickLine={false} axisLine={false} fontSize={12} />
                        <YAxis tickLine={false} axisLine={false} fontSize={12} domain={[0, 1]} tickFormatter={(value) => `${Math.round(value * 100)}%`} />
                        <Tooltip content={<CompareTrendTooltip />} />
                        <Area
                          type="monotone"
                          dataKey="leftProofAccuracy"
                          stroke="#0f766e"
                          fill="url(#compareLeftFill)"
                          strokeWidth={2.5}
                          name={`区间A ${compareTrend.leftLabel}`}
                        />
                        <Area
                          type="monotone"
                          dataKey="rightProofAccuracy"
                          stroke="#1d4ed8"
                          fill="transparent"
                          strokeWidth={2.5}
                          name={`区间B ${compareTrend.rightLabel}`}
                        />
                      </AreaChart>
                    </ResponsiveContainer>
                  </div>
                </div>

                <div className="rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                  <div className="mb-6 flex items-end justify-between gap-4">
                    <div>
                    <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Session Compare</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">场次对比</h2>
                      <p className="mt-2 text-sm text-slate-500">选择一个场次后查看 A/B 区间下的举证准确率与申报次数。</p>
                    </div>
                    <div className="w-full max-w-[240px]">
                      <FilterSelect
                        label="选择场次"
                        icon={<Layers3 size={16} />}
                        value={compareSessionSelection}
                        options={compareSessionOptions}
                        onChange={setCompareSessionSelection}
                        tone="light"
                      />
                    </div>
                  </div>
                  <div className="space-y-3">
                    {compareSelectedSession ? (
                        <div
                          key={compareSelectedSession.session}
                          className="rounded-2xl border border-slate-100 bg-slate-50/80 px-4 py-3"
                        >
                          <div className="flex items-center justify-between gap-4">
                            <div>
                              <p className="font-medium text-slate-900">{compareSelectedSession.session}</p>
                              <p className="mt-1 text-xs text-slate-500">
                                A申报 {formatInteger(compareSelectedSession.leftDeclarations)} / B申报 {formatInteger(compareSelectedSession.rightDeclarations)}
                              </p>
                            </div>
                            <DeltaBadge
                              delta={compareSelectedSession.leftProofAccuracy - compareSelectedSession.rightProofAccuracy}
                              formatter={formatPercent}
                            />
                          </div>
                          <div className="mt-3 grid grid-cols-2 gap-3 text-sm">
                            <div className="rounded-xl bg-white px-3 py-2">
                              <p className="text-xs uppercase tracking-[0.16em] text-slate-400">区间 A</p>
                              <p className="mt-1 font-medium text-slate-900">{formatPercent(compareSelectedSession.leftProofAccuracy)}</p>
                            </div>
                            <div className="rounded-xl bg-white px-3 py-2">
                              <p className="text-xs uppercase tracking-[0.16em] text-slate-400">区间 B</p>
                              <p className="mt-1 font-medium text-slate-900">{formatPercent(compareSelectedSession.rightProofAccuracy)}</p>
                            </div>
                          </div>
                        </div>
                    ) : (
                      <div className="rounded-2xl border border-dashed border-slate-200 px-4 py-8 text-center text-sm text-slate-500">
                        先选择两个完整区间，再查看场次对比。
                      </div>
                    )}
                  </div>
                </div>
              </section>

              <section className="mt-8 rounded-[28px] border border-slate-200 bg-white p-6 shadow-[0_18px_45px_rgba(15,23,42,0.05)]">
                <div className="mb-5 flex items-end justify-between gap-4">
                  <div>
                    <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Batch Compare</p>
                    <h2 className="mt-2 font-display text-2xl text-slate-900">批次对比</h2>
                    <p className="mt-2 text-sm text-slate-500">选择两个批次，图示比较它们在 A/B 区间下的举证准确率。</p>
                  </div>
                  <div className="rounded-full bg-slate-100 px-3 py-1.5 text-xs text-slate-600">
                    当前对比指标：举证准确率
                  </div>
                </div>
                <div className="grid gap-4 lg:grid-cols-[260px_260px_1fr]">
                  {compareBatchRows.length ? (
                    <>
                      <FilterSelect
                        label="批次一"
                        icon={<Boxes size={16} />}
                        value={compareBatchFirstSelection}
                        options={compareBatchOptions}
                        onChange={setCompareBatchFirstSelection}
                        tone="light"
                      />
                      <FilterSelect
                        label="批次二"
                        icon={<Boxes size={16} />}
                        value={compareBatchSecondSelection}
                        options={compareBatchOptions}
                        onChange={setCompareBatchSecondSelection}
                        tone="light"
                      />
                      <div className="rounded-2xl border border-slate-100 bg-slate-50/70 p-4">
                        <div className="mb-4 flex flex-wrap items-center gap-3 text-xs text-slate-500">
                          <span className="rounded-full bg-white px-3 py-1.5">
                            {compareSelectedBatchFirst?.batch ?? '批次一'}：A申报 {formatInteger(compareSelectedBatchFirst?.leftDeclarations ?? 0)} / B申报 {formatInteger(compareSelectedBatchFirst?.rightDeclarations ?? 0)}
                          </span>
                          <span className="rounded-full bg-white px-3 py-1.5">
                            {compareSelectedBatchSecond?.batch ?? '批次二'}：A申报 {formatInteger(compareSelectedBatchSecond?.leftDeclarations ?? 0)} / B申报 {formatInteger(compareSelectedBatchSecond?.rightDeclarations ?? 0)}
                          </span>
                        </div>
                        <div className="h-[280px] w-full">
                          <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={compareBatchChartData} barGap={18}>
                              <CartesianGrid vertical={false} stroke="#e2e8f0" strokeDasharray="3 3" />
                              <XAxis dataKey="metric" tickLine={false} axisLine={false} fontSize={12} />
                              <YAxis
                                tickLine={false}
                                axisLine={false}
                                fontSize={12}
                                domain={[0, 1]}
                                tickFormatter={(value) => `${Math.round(value * 100)}%`}
                              />
                              <Tooltip content={<RateTrendTooltip />} />
                              <Bar
                                dataKey={compareSelectedBatchFirst?.batch ?? '批次一'}
                                radius={[10, 10, 0, 0]}
                                fill="#0f766e"
                              />
                              <Bar
                                dataKey={compareSelectedBatchSecond?.batch ?? '批次二'}
                                radius={[10, 10, 0, 0]}
                                fill="#1d4ed8"
                              />
                            </BarChart>
                          </ResponsiveContainer>
                        </div>
                      </div>
                    </>
                  ) : (
                    <div className="rounded-2xl border border-dashed border-slate-200 px-4 py-8 text-center text-sm text-slate-500 lg:col-span-3">
                      先选择两个完整区间，再查看批次对比。
                    </div>
                  )}
                </div>
              </section>
            </>
          ) : (
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
        className={`w-full rounded-2xl px-4 py-2.5 text-sm outline-none transition ${
          tone === 'dark'
            ? 'border border-white/10 bg-white/8 text-white focus:border-white/40'
            : 'border border-slate-200 bg-white text-slate-900 focus:border-slate-400'
        }`}
      >
        {options.map((option) => (
          <option key={option} value={option} className="text-slate-900">
            {option}
          </option>
        ))}
      </select>
    </label>
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
        className="w-full rounded-2xl border border-white/10 bg-white/8 px-4 py-2.5 text-sm text-white outline-none transition focus:border-white/40"
      >
        <option value={ALL_OPTION} className="text-slate-900">
          全部周
        </option>
        {options.map((week) => (
          <option key={week.value} value={week.value} className="text-slate-900">
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

function ToolbarChip({ icon, label }: { icon: React.ReactNode; label: string }) {
  return (
    <div className="inline-flex items-center gap-2 rounded-full border border-slate-200 bg-slate-50 px-3 py-1.5">
      <span className="text-slate-400">{icon}</span>
      <span className="max-w-[240px] truncate">{label}</span>
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
