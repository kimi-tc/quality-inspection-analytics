import React, { ChangeEvent, useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { DayPicker, type DateRange } from 'react-day-picker';
import { addMonths, format, parseISO, startOfMonth, subMonths } from 'date-fns';
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

const createOptions = (rows: ImportedRow[], key: keyof ImportedRow) =>
  [ALL_OPTION].concat(
    [...new Set(rows.map((row) => String(row[key] ?? '')).filter(Boolean))].sort((a, b) =>
      a.localeCompare(b, 'zh-CN'),
    ),
  );

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
  const [startDateFilter, setStartDateFilter] = useState(ALL_OPTION);
  const [endDateFilter, setEndDateFilter] = useState(ALL_OPTION);
  const [sessionFilter, setSessionFilter] = useState(ALL_OPTION);
  const [batchFilter, setBatchFilter] = useState(ALL_OPTION);
  const [attributeFilter, setAttributeFilter] = useState(ALL_OPTION);
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

  const filteredRows = useMemo(
    () =>
      dataset.rows.filter((row) => {
        const startMatch = startDateFilter === ALL_OPTION || row.date >= startDateFilter;
        const endMatch = endDateFilter === ALL_OPTION || row.date <= endDateFilter;
        const sessionMatch = sessionFilter === ALL_OPTION || row.session === sessionFilter;
        const batchMatch = batchFilter === ALL_OPTION || row.batch === batchFilter;
        const attributeMatch = attributeFilter === ALL_OPTION || row.attribute === attributeFilter;

        return startMatch && endMatch && sessionMatch && batchMatch && attributeMatch;
      }),
    [attributeFilter, batchFilter, dataset.rows, endDateFilter, sessionFilter, startDateFilter],
  );

  const metrics = useMemo(() => aggregateMetrics(filteredRows), [filteredRows]);
  const trendData = useMemo(() => aggregateTrend(filteredRows), [filteredRows]);
  const topAttributes = useMemo(() => aggregateAttributes(filteredRows), [filteredRows]);

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
      setSessionFilter(ALL_OPTION);
      setBatchFilter(ALL_OPTION);
      setAttributeFilter(ALL_OPTION);
    } catch (err) {
      setError(err instanceof Error ? err.message : '清空共享数据失败。');
    }
  };

  return (
    <div className="min-h-screen bg-[radial-gradient(circle_at_top,_#f6efe4_0%,_#f3f8f6_40%,_#eef2ff_100%)] text-slate-900">
      <div className="mx-auto max-w-7xl px-4 py-6 sm:px-6 lg:px-8">
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
                <div className="grid gap-3 md:grid-cols-2 xl:grid-cols-[repeat(5,minmax(0,1fr))_auto] xl:items-end">
                  <DateRangeFilter
                    startValue={startDateFilter}
                    endValue={endDateFilter}
                    options={options.dates}
                    onStartChange={setStartDateFilter}
                    onEndChange={setEndDateFilter}
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
          </div>
        </motion.section>

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
              <p className="text-sm uppercase tracking-[0.26em] text-slate-400">Ranking</p>
              <h2 className="mt-2 font-display text-2xl text-slate-900">高频属性项</h2>
            </div>
            <div className="h-[320px] w-full">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={topAttributes} layout="vertical" margin={{ top: 0, right: 10, left: 10, bottom: 0 }}>
                  <CartesianGrid horizontal={false} stroke="#eef2f7" />
                  <XAxis type="number" hide />
                  <YAxis
                    type="category"
                    dataKey="attribute"
                    width={88}
                    tickLine={false}
                    axisLine={false}
                    fontSize={12}
                  />
                  <Tooltip />
                  <Bar dataKey="declarations" radius={[0, 10, 10, 0]} fill="#1d4ed8" name="申报次数" />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        </section>
      </div>
    </div>
  );
}

function FilterSelect({
  label,
  value,
  options,
  onChange,
  icon,
}: {
  label: string;
  value: string;
  options: string[];
  onChange: (value: string) => void;
  icon: React.ReactNode;
}) {
  return (
    <label className="block">
      <span className="mb-2 flex items-center gap-2 text-xs uppercase tracking-[0.2em] text-slate-400">
        {icon}
        {label}
      </span>
      <select
        value={value}
        onChange={(event) => onChange(event.target.value)}
        className="w-full rounded-2xl border border-white/10 bg-white/8 px-4 py-2.5 text-sm text-white outline-none transition focus:border-white/40"
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

function DateRangeFilter({
  startValue,
  endValue,
  options,
  onStartChange,
  onEndChange,
  onClear,
}: {
  startValue: string;
  endValue: string;
  options: string[];
  onStartChange: (value: string) => void;
  onEndChange: (value: string) => void;
  onClear: () => void;
}) {
  const hasValue = startValue !== ALL_OPTION || endValue !== ALL_OPTION;
  const containerRef = useRef<HTMLDivElement | null>(null);
  const [isOpen, setIsOpen] = useState(false);
  const availableDates = options.filter((option) => option !== ALL_OPTION);
  const availableDateSet = new Set(availableDates);
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

  const handleSelect = (range: DateRange | undefined) => {
    if (!range?.from) {
      onClear();
      return;
    }

    const nextStart = format(range.from, 'yyyy-MM-dd');
    const nextEnd = range.to ? format(range.to, 'yyyy-MM-dd') : ALL_OPTION;

    onStartChange(nextStart);
    onEndChange(nextEnd);

    if (range.from && range.to) {
      setIsOpen(false);
    }
  };

  return (
    <div className="relative md:col-span-2 xl:col-span-2" ref={containerRef}>
      <span className="mb-2 flex items-center gap-2 text-xs uppercase tracking-[0.2em] text-slate-400">
        <CalendarDays size={16} />
        日期区间
      </span>
      <div
        role="button"
        tabIndex={0}
        onClick={() => setIsOpen((value) => !value)}
        onKeyDown={(event) => {
          if (event.key === 'Enter' || event.key === ' ') {
            event.preventDefault();
            setIsOpen((value) => !value);
          }
        }}
        className="flex w-full items-center rounded-2xl border border-[#4B8DFF] bg-white px-5 py-3 text-left shadow-[0_8px_24px_rgba(37,99,235,0.12)] outline-none transition hover:border-[#2f7cff]"
      >
        <span className="min-w-0 flex-1 text-[15px] font-medium text-slate-800">
          {formatDateDisplay(startValue) || '开始日期'}
        </span>
        <span className="mx-4 shrink-0 text-slate-400">→</span>
        <span className="min-w-0 flex-1 text-[15px] font-medium text-slate-800">
          {formatDateDisplay(endValue) || '结束日期'}
        </span>
        <button
          type="button"
          onClick={(event) => {
            event.stopPropagation();
            onClear();
          }}
          className={`ml-3 flex h-7 w-7 items-center justify-center rounded-full transition ${
            hasValue ? 'bg-slate-300 text-white hover:bg-slate-400' : 'bg-slate-100 text-slate-300'
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
            selected={selectedRange}
            onSelect={handleSelect}
            defaultMonth={selectedRange?.from ?? minDate}
            startMonth={visibleStartMonth}
            endMonth={visibleEndMonth}
            showOutsideDays
            disabled={(date) => !availableDateSet.has(format(date, 'yyyy-MM-dd'))}
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
    <div className="rounded-[26px] border border-slate-200 bg-white p-5 shadow-[0_14px_35px_rgba(15,23,42,0.04)]">
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
