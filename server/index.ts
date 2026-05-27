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
};

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');
const dataDir = process.env.DATA_DIR
  ? path.resolve(process.env.DATA_DIR)
  : path.join(rootDir, 'data');
const dataFile = path.join(dataDir, 'shared-dataset.json');
const distDir = path.join(rootDir, 'dist');

const emptyDataset: SharedDataset = {
  rows: [],
  importedAt: '',
  sourceName: '',
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

const ensureDataFile = async () => {
  await fs.mkdir(dataDir, { recursive: true });

  try {
    await fs.access(dataFile);
  } catch {
    await fs.writeFile(dataFile, JSON.stringify(emptyDataset, null, 2), 'utf8');
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
  const nextDataset: SharedDataset = {
    rows: mergedRows,
    importedAt: new Date().toISOString(),
    sourceName: current.sourceName
      ? `${current.sourceName} + ${body.sourceName ?? '新导入数据'}`
      : body.sourceName ?? '新导入数据',
  };

  await writeDataset(nextDataset);
  res.json(nextDataset);
});

app.delete('/api/dataset', async (_req, res) => {
  await writeDataset(emptyDataset);
  res.json(emptyDataset);
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
  console.log(`Shared dataset API running on http://127.0.0.1:${port}`);
});
