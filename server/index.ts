import express from 'express';
import fs from 'node:fs/promises';
import fsSync from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import dotenv from 'dotenv';

dotenv.config();

type ImportedRow = {
  date: string;
  auditor?: string;
  auditorTeam?: string;
  session: string;
  batch: string;
  category: string;
  attribute: string;
  declarations: number;
  ambiguousPasses: number;
  rejects: number;
  proofRejects: number;
};

type SharedDataset = {
  rows: ImportedRow[];
  importedAt: string;
  sourceName: string;
  importHistory?: ImportRecord[];
};

type EfficiencyRow = {
  date: string;
  employee: string;
  team: string;
  session: string;
  batch: string;
  handledCount: number;
  weightedHandledCount?: number;
  firstAuditCount?: number;
  firstAuditPassCount?: number;
  precisionPassCount?: number;
  auditNotPassCount?: number;
  proofRefusalCount?: number;
  ambiguousCount?: number;
  passRate?: number;
  precisionPassRate?: number;
  proofAccuracy?: number;
  avgHandleMinutes: number;
  timeoutCount: number;
};

type EfficiencyDataset = {
  rows: EfficiencyRow[];
  importedAt: string;
  sourceName: string;
  importHistory?: ImportRecord[];
};

type ImportRecord = {
  id: string;
  sourceName: string;
  importedAt: string;
  rowCount: number;
  dataType: 'quality' | 'efficiency';
};

type PropertyCategoryEntry = {
  propertyName: string;
  category: string;
};

type AuditorTeamEntry = {
  auditorName: string;
  team: string;
};

type AiAnalysisRequest = {
  report?: string;
  model?: string;
  context?: {
    qualityRows?: number;
    efficiencyRows?: number;
    dateRange?: string;
    filters?: string;
  };
};

const propertyCategoryOptions = ['维修项', '外观项', '功能项', 'SKU项', '其他', '售后补充项'];
const normalizeDictionaryCategory = (value: string) => {
  const normalizedValue = value.trim();
  const legacyCategoryMap: Record<string, string> = {
    主观项: '外观项',
    零售附加项: '售后补充项',
    零售补充项: '售后补充项',
  };
  const mappedValue = legacyCategoryMap[normalizedValue] ?? normalizedValue;

  return propertyCategoryOptions.includes(mappedValue) ? mappedValue : '其他';
};

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');
const dataDir = process.env.DATA_DIR
  ? path.resolve(process.env.DATA_DIR)
  : path.join(rootDir, 'data');
const dataFile = path.join(dataDir, 'shared-dataset.json');
const efficiencyDataFile = path.join(dataDir, 'efficiency-dataset.json');
const propertyCategoryDictionaryFile = path.join(dataDir, 'property-category-dictionary.json');
const propertyCategorySeedFile = path.join(rootDir, 'config', 'property-category-dictionary.seed.json');
const auditorTeamDictionaryFile = path.join(dataDir, 'auditor-team-dictionary.json');
const auditorTeamSeedFile = path.join(rootDir, 'config', 'auditor-team-dictionary.seed.json');
const distDir = path.join(rootDir, 'dist');

const emptyDataset: SharedDataset = {
  rows: [],
  importedAt: '',
  sourceName: '',
  importHistory: [],
};

const emptyEfficiencyDataset: EfficiencyDataset = {
  rows: [],
  importedAt: '',
  sourceName: '',
  importHistory: [],
};

const readSeedPropertyCategoryDictionary = async (): Promise<PropertyCategoryEntry[]> => {
  try {
    const content = await fs.readFile(propertyCategorySeedFile, 'utf8');
    const parsed = JSON.parse(content) as PropertyCategoryEntry[];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
};

const readSeedAuditorTeamDictionary = async (): Promise<AuditorTeamEntry[]> => {
  try {
    const content = await fs.readFile(auditorTeamSeedFile, 'utf8');
    const parsed = JSON.parse(content) as AuditorTeamEntry[];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
};

const app = express();
const port = Number(process.env.PORT || 8787);
const sharedDataApiBaseUrl = process.env.SHARED_DATA_API_BASE_URL?.trim().replace(/\/$/, '') || '';
const sharedDataApiTimeoutMs = Number(process.env.SHARED_DATA_API_TIMEOUT_MS || 120000);
const sharedDataApiPrefixes = [
  '/api/dataset',
  '/api/efficiency-dataset',
  '/api/property-category-dictionary',
];

app.use(express.json({ limit: '100mb' }));

app.get('/api/data-source', (_req, res) => {
  res.json({
    mode: sharedDataApiBaseUrl ? 'remote' : 'local',
    baseUrl: sharedDataApiBaseUrl || `http://127.0.0.1:${port}`,
  });
});

app.use(async (req, res, next) => {
  const shouldProxy =
    sharedDataApiBaseUrl &&
    sharedDataApiPrefixes.some(
      (prefix) => req.path === prefix || req.path.startsWith(`${prefix}/`),
    );

  if (!shouldProxy) {
    next();
    return;
  }

  try {
    const upstreamResponse = await fetch(`${sharedDataApiBaseUrl}${req.originalUrl}`, {
      method: req.method,
      headers: {
        accept: 'application/json',
        ...(req.method === 'GET' || req.method === 'HEAD'
          ? {}
          : { 'content-type': 'application/json' }),
      },
      body:
        req.method === 'GET' || req.method === 'HEAD'
          ? undefined
          : JSON.stringify(req.body ?? {}),
      signal: AbortSignal.timeout(sharedDataApiTimeoutMs),
    });
    const responseBody = await upstreamResponse.text();
    const contentType = upstreamResponse.headers.get('content-type');

    if (contentType) {
      res.setHeader('content-type', contentType);
    }
    res.status(upstreamResponse.status).send(responseBody);
  } catch (error) {
    res.status(502).json({
      message: `统一数据服务连接失败：${error instanceof Error ? error.message : '未知错误'}`,
      baseUrl: sharedDataApiBaseUrl,
    });
  }
});

const buildRowKey = (row: ImportedRow) =>
  [
    row.date,
    row.auditor ?? '',
    row.auditorTeam ?? '',
    row.session,
    row.batch,
    row.category,
    row.attribute,
    row.declarations,
    row.ambiguousPasses,
    row.rejects,
    row.proofRejects,
  ].join('::');

const buildEfficiencyRowKey = (row: EfficiencyRow) =>
  [
    row.date,
    row.employee,
    row.team,
    row.session,
    row.batch,
  ].join('::');

const mergeRows = (existingRows: ImportedRow[], incomingRows: ImportedRow[]) => {
  const rowMap = new Map<string, ImportedRow>();

  for (const row of existingRows) {
    rowMap.set(buildRowKey(row), row);
  }

  for (const row of incomingRows) {
    rowMap.set(buildRowKey(row), row);
  }

  return [...rowMap.values()].sort((a, b) => {
    if (a.date !== b.date) return a.date.localeCompare(b.date);
    if ((a.auditor ?? '') !== (b.auditor ?? '')) {
      return (a.auditor ?? '').localeCompare(b.auditor ?? '', 'zh-CN');
    }
    if ((a.auditorTeam ?? '') !== (b.auditorTeam ?? '')) {
      return (a.auditorTeam ?? '').localeCompare(b.auditorTeam ?? '', 'zh-CN');
    }
    if (a.session !== b.session) return a.session.localeCompare(b.session, 'zh-CN');
    if (a.batch !== b.batch) return a.batch.localeCompare(b.batch, 'zh-CN');
    return a.attribute.localeCompare(b.attribute, 'zh-CN');
  });
};

const mergeEfficiencyRows = (existingRows: EfficiencyRow[], incomingRows: EfficiencyRow[]) => {
  const rowMap = new Map<string, EfficiencyRow>();

  for (const row of existingRows) {
    rowMap.set(buildEfficiencyRowKey(row), row);
  }

  for (const row of incomingRows) {
    rowMap.set(buildEfficiencyRowKey(row), row);
  }

  return [...rowMap.values()].sort((a, b) => {
    if (a.date !== b.date) return a.date.localeCompare(b.date);
    if (a.team !== b.team) return a.team.localeCompare(b.team, 'zh-CN');
    return a.employee.localeCompare(b.employee, 'zh-CN');
  });
};

const ensureDataFile = async () => {
  await fs.mkdir(dataDir, { recursive: true });

  try {
    await fs.access(dataFile);
  } catch {
    await fs.writeFile(dataFile, JSON.stringify(emptyDataset, null, 2), 'utf8');
  }
};

const ensureEfficiencyDataFile = async () => {
  await fs.mkdir(dataDir, { recursive: true });

  try {
    await fs.access(efficiencyDataFile);
  } catch {
    await fs.writeFile(efficiencyDataFile, JSON.stringify(emptyEfficiencyDataset, null, 2), 'utf8');
  }
};

const ensurePropertyCategoryDictionaryFile = async () => {
  await fs.mkdir(dataDir, { recursive: true });

  try {
    await fs.access(propertyCategoryDictionaryFile);
  } catch {
    const seed = await readSeedPropertyCategoryDictionary();
    await fs.writeFile(propertyCategoryDictionaryFile, JSON.stringify(seed, null, 2), 'utf8');
  }
};

const ensureAuditorTeamDictionaryFile = async () => {
  await fs.mkdir(dataDir, { recursive: true });

  try {
    await fs.access(auditorTeamDictionaryFile);
  } catch {
    const seed = await readSeedAuditorTeamDictionary();
    await fs.writeFile(auditorTeamDictionaryFile, JSON.stringify(seed, null, 2), 'utf8');
  }
};

const readDataset = async (): Promise<SharedDataset> => {
  await ensureDataFile();
  const content = await fs.readFile(dataFile, 'utf8');

  try {
    const parsed = JSON.parse(content) as SharedDataset;
    if (!Array.isArray(parsed.rows)) {
      return emptyDataset;
    }
    return parsed;
  } catch {
    return emptyDataset;
  }
};

const writeDataset = async (dataset: SharedDataset) => {
  await ensureDataFile();
  await fs.writeFile(dataFile, JSON.stringify(dataset, null, 2), 'utf8');
};

const readEfficiencyDataset = async (): Promise<EfficiencyDataset> => {
  await ensureEfficiencyDataFile();
  const content = await fs.readFile(efficiencyDataFile, 'utf8');

  try {
    const parsed = JSON.parse(content) as EfficiencyDataset;
    if (!Array.isArray(parsed.rows)) {
      return emptyEfficiencyDataset;
    }
    return {
      ...parsed,
      rows: mergeEfficiencyRows([], parsed.rows),
    };
  } catch {
    return emptyEfficiencyDataset;
  }
};

const writeEfficiencyDataset = async (dataset: EfficiencyDataset) => {
  await ensureEfficiencyDataFile();
  await fs.writeFile(efficiencyDataFile, JSON.stringify(dataset, null, 2), 'utf8');
};

const readPropertyCategoryDictionary = async (): Promise<PropertyCategoryEntry[]> => {
  await ensurePropertyCategoryDictionaryFile();
  const content = await fs.readFile(propertyCategoryDictionaryFile, 'utf8');

  try {
    const parsed = JSON.parse(content) as PropertyCategoryEntry[];
    if (!Array.isArray(parsed)) {
      return [];
    }

    return parsed
      .map((entry) => ({
        propertyName: String(entry.propertyName ?? '').trim(),
        category: normalizeDictionaryCategory(String(entry.category ?? '')),
      }))
      .filter((entry) => entry.propertyName && entry.category);
  } catch {
    return [];
  }
};

const writePropertyCategoryDictionary = async (entries: PropertyCategoryEntry[]) => {
  await ensurePropertyCategoryDictionaryFile();
  const normalizedEntries = entries
    .map((entry) => ({
      propertyName: String(entry.propertyName ?? '').trim(),
      category: normalizeDictionaryCategory(String(entry.category ?? '')),
    }))
    .filter((entry) => entry.propertyName && entry.category);

  await fs.writeFile(propertyCategoryDictionaryFile, JSON.stringify(normalizedEntries, null, 2), 'utf8');
};

const readAuditorTeamDictionary = async (): Promise<AuditorTeamEntry[]> => {
  await ensureAuditorTeamDictionaryFile();
  const content = await fs.readFile(auditorTeamDictionaryFile, 'utf8');

  try {
    const parsed = JSON.parse(content) as AuditorTeamEntry[];
    if (!Array.isArray(parsed)) {
      return [];
    }

    return parsed
      .map((entry) => ({
        auditorName: String(entry.auditorName ?? '').trim(),
        team: String(entry.team ?? '').trim(),
      }))
      .filter((entry) => entry.auditorName && entry.team);
  } catch {
    return [];
  }
};

const writeAuditorTeamDictionary = async (entries: AuditorTeamEntry[]) => {
  await ensureAuditorTeamDictionaryFile();
  const normalizedEntries = entries
    .map((entry) => ({
      auditorName: String(entry.auditorName ?? '').trim(),
      team: String(entry.team ?? '').trim(),
    }))
    .filter((entry) => entry.auditorName && entry.team);

  await fs.writeFile(auditorTeamDictionaryFile, JSON.stringify(normalizedEntries, null, 2), 'utf8');
};

const createImportRecord = (
  dataType: ImportRecord['dataType'],
  sourceName: string | undefined,
  rowCount: number,
): ImportRecord => {
  const importedAt = new Date().toISOString();
  return {
    id: `${dataType}-${importedAt}-${Math.random().toString(36).slice(2, 8)}`,
    sourceName: sourceName || (dataType === 'quality' ? '新导入数据' : '新人效数据'),
    importedAt,
    rowCount,
    dataType,
  };
};

app.get('/api/dataset', async (_req, res) => {
  const dataset = await readDataset();
  res.json(dataset);
});

app.post('/api/dataset/merge', async (req, res) => {
  const body = req.body as Partial<SharedDataset>;

  if (!Array.isArray(body.rows)) {
    return res.status(400).json({ message: 'rows is required' });
  }

  const current = await readDataset();
  const mergedRows = mergeRows(current.rows, body.rows as ImportedRow[]);
  const importRecord = createImportRecord('quality', body.sourceName, body.rows.length);
  const nextDataset: SharedDataset = {
    rows: mergedRows,
    importedAt: importRecord.importedAt,
    sourceName: current.sourceName
      ? `${current.sourceName} + ${body.sourceName ?? '新导入数据'}`
      : body.sourceName ?? '新导入数据',
    importHistory: [...(current.importHistory ?? []), importRecord],
  };

  await writeDataset(nextDataset);
  res.json(nextDataset);
});

app.delete('/api/dataset', async (_req, res) => {
  await writeDataset(emptyDataset);
  res.json(emptyDataset);
});

app.get('/api/efficiency-dataset', async (_req, res) => {
  const dataset = await readEfficiencyDataset();
  res.json(dataset);
});

app.post('/api/efficiency-dataset/merge', async (req, res) => {
  const body = req.body as Partial<EfficiencyDataset>;

  if (!Array.isArray(body.rows)) {
    return res.status(400).json({ message: 'rows is required' });
  }

  const current = await readEfficiencyDataset();
  const mergedRows = mergeEfficiencyRows(current.rows, body.rows as EfficiencyRow[]);
  const importRecord = createImportRecord('efficiency', body.sourceName, body.rows.length);
  const nextDataset: EfficiencyDataset = {
    rows: mergedRows,
    importedAt: importRecord.importedAt,
    sourceName: current.sourceName
      ? `${current.sourceName} + ${body.sourceName ?? '新人效数据'}`
      : body.sourceName ?? '新人效数据',
    importHistory: [...(current.importHistory ?? []), importRecord],
  };

  await writeEfficiencyDataset(nextDataset);
  res.json(nextDataset);
});

app.delete('/api/efficiency-dataset', async (_req, res) => {
  await writeEfficiencyDataset(emptyEfficiencyDataset);
  res.json(emptyEfficiencyDataset);
});

app.get('/api/property-category-dictionary', async (_req, res) => {
  const dictionary = await readPropertyCategoryDictionary();
  res.json({ entries: dictionary });
});

app.put('/api/property-category-dictionary', async (req, res) => {
  const body = req.body as { entries?: PropertyCategoryEntry[] };

  if (!Array.isArray(body.entries)) {
    return res.status(400).json({ message: 'entries is required' });
  }

  await writePropertyCategoryDictionary(body.entries);
  const dictionary = await readPropertyCategoryDictionary();
  res.json({ entries: dictionary });
});

app.post('/api/property-category-dictionary/reset', async (_req, res) => {
  const seed = await readSeedPropertyCategoryDictionary();
  await writePropertyCategoryDictionary(seed);
  res.json({ entries: seed });
});

app.get('/api/auditor-team-dictionary', async (_req, res) => {
  const dictionary = await readAuditorTeamDictionary();
  res.json({ entries: dictionary });
});

app.put('/api/auditor-team-dictionary', async (req, res) => {
  const body = req.body as { entries?: AuditorTeamEntry[] };

  if (!Array.isArray(body.entries)) {
    return res.status(400).json({ message: 'entries is required' });
  }

  await writeAuditorTeamDictionary(body.entries);
  const dictionary = await readAuditorTeamDictionary();
  res.json({ entries: dictionary });
});

app.post('/api/auditor-team-dictionary/reset', async (_req, res) => {
  const seed = await readSeedAuditorTeamDictionary();
  await writeAuditorTeamDictionary(seed);
  res.json({ entries: seed });
});

app.post('/api/ai-analysis', async (req, res) => {
  const apiKey = process.env.DEEPSEEK_API_KEY?.trim();
  const baseUrl = process.env.DEEPSEEK_BASE_URL?.trim() || 'https://api.deepseek.com';
  const timeoutMs = Number(process.env.DEEPSEEK_TIMEOUT_MS || 180000);

  if (!apiKey) {
    return res.status(503).json({
      message: '服务端未配置 DEEPSEEK_API_KEY，当前只能使用规则版 AI 分析。',
    });
  }

  const body = req.body as AiAnalysisRequest;
  const report = body.report?.trim();
  const requestedModel = body.model?.trim();
  const model = requestedModel || process.env.DEEPSEEK_MODEL?.trim() || 'deepseek-chat';

  if (!report) {
    return res.status(400).json({ message: 'report is required' });
  }

  try {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), timeoutMs);
    const upstreamResponse = await fetch(`${baseUrl.replace(/\/$/, '')}/chat/completions`, {
      method: 'POST',
      signal: controller.signal,
      headers: {
        Authorization: `Bearer ${apiKey}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        model,
        temperature: 0.3,
        messages: [
          {
            role: 'system',
            content: [
              '你是一名资深预质检质量运营分析师。',
              '请基于用户提供的聚合看板结果做深入分析，不要编造未提供的数据。',
              '输出中文，适合直接放入飞书周报。',
              '请包含：1段总览、3-5条关键洞察、2-4条风险归因、3-5条下周动作建议。',
              '分析重点优先级：1. 仅关注申报占比超过 1% 的场次，并在这些场次内定位通过率变化较大的属性项；2. 高拒绝率属性项；3. 高模棱两可率属性项；4. 精准通过率波动明显的属性项；5. 人员维度中审核效率明显变化、以及精准通过率和效率同时异动的人员；6. 批次差异作为辅助解释。',
              '模棱两可率目标为 7%，高于 7% 时需要明确指出超标幅度和优先复盘方向。',
              '人员效率只看总审核量、加权审核量、日均加权审核量等审核产出口径；不要分析平均处理时长、超时率等人效效率指标。',
              '请尽量把建议落到具体场次、具体属性项、具体人员和可执行复盘动作上；如果草稿中给出了“场次-属性项”组合或人员异动，优先围绕这些对象诊断。',
              '避免空泛话术，每条建议尽量说明优先级和落地方式。',
            ].join('\n'),
          },
          {
            role: 'user',
            content: [
              body.context ? `补充上下文：${JSON.stringify(body.context)}` : '',
              `规则版分析草稿如下：\n${report}`,
            ].filter(Boolean).join('\n\n'),
          },
        ],
      }),
    }).finally(() => clearTimeout(timeout));

    const rawPayload = await upstreamResponse.text();
    let payload: {
      choices?: Array<{ message?: { content?: string } }>;
      error?: { message?: string };
    } = {};

    try {
      payload = rawPayload ? JSON.parse(rawPayload) : {};
    } catch {
      payload = {
        error: {
          message: rawPayload.slice(0, 240) || 'DeepSeek 返回了非 JSON 响应。',
        },
      };
    }

    if (!upstreamResponse.ok) {
      return res.status(upstreamResponse.status).json({
        message: payload.error?.message || `DeepSeek 分析生成失败，HTTP ${upstreamResponse.status}`,
      });
    }

    res.json({
      model,
      analysis: payload.choices?.[0]?.message?.content?.trim() || 'DeepSeek 已响应，但未返回可展示文本。',
      generatedAt: new Date().toISOString(),
    });
  } catch (error) {
    const message =
      error instanceof Error && error.name === 'AbortError'
        ? `DeepSeek 请求超过 ${Math.round(timeoutMs / 1000)} 秒未返回，请稍后重试或切换更快的模型。`
        : error instanceof Error
          ? error.message
          : '大模型分析生成失败';
    res.status(500).json({ message });
  }
});

if (fsSync.existsSync(distDir)) {
  app.use(express.static(distDir));

  app.get('*', (req, res, next) => {
    if (req.path.startsWith('/api/')) {
      next();
      return;
    }

    res.sendFile(path.join(distDir, 'index.html'));
  });
}

app.listen(port, async () => {
  await ensureDataFile();
  await ensureEfficiencyDataFile();
  await ensurePropertyCategoryDictionaryFile();
  await ensureAuditorTeamDictionaryFile();
  console.log(`Shared dataset API running on http://127.0.0.1:${port}`);
});
