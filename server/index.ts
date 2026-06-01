import express from 'express';
import fs from 'node:fs/promises';
import fsSync from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

type ImportedRow = {
  date: string;
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

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');
const dataDir = process.env.DATA_DIR
  ? path.resolve(process.env.DATA_DIR)
  : path.join(rootDir, 'data');
const dataFile = path.join(dataDir, 'shared-dataset.json');
const efficiencyDataFile = path.join(dataDir, 'efficiency-dataset.json');
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

const app = express();
const port = Number(process.env.PORT || 8787);

app.use(express.json({ limit: '20mb' }));

const buildRowKey = (row: ImportedRow) =>
  [
    row.date,
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
    row.handledCount,
    row.avgHandleMinutes,
    row.timeoutCount,
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
    return parsed;
  } catch {
    return emptyEfficiencyDataset;
  }
};

const writeEfficiencyDataset = async (dataset: EfficiencyDataset) => {
  await ensureEfficiencyDataFile();
  await fs.writeFile(efficiencyDataFile, JSON.stringify(dataset, null, 2), 'utf8');
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
  console.log(`Shared dataset API running on http://127.0.0.1:${port}`);
});
