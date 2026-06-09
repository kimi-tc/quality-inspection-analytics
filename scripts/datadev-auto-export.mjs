#!/usr/bin/env node
import fs from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import { spawn } from 'node:child_process';
import { chromium } from 'playwright';

const projectDir = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
const datadevUrl = process.env.DATADEV_URL || 'http://datadev.aihuishou.com/workbench';
const downloadDir = process.env.QUALITY_DOWNLOAD_DIR || path.join(process.env.HOME || '.', 'Downloads', '预质检每日数据');
const profileDir = process.env.DATADEV_BROWSER_PROFILE || path.join(process.env.HOME || '.', '.quality-inspection-datadev-profile');
const browserChannel = process.env.DATADEV_BROWSER_CHANNEL || 'chrome';
const sqlFile = process.env.DATADEV_SQL_FILE || '';
const sqlText = process.env.DATADEV_SQL_TEXT || '';
const timeoutMs = Number(process.env.DATADEV_WAIT_TIMEOUT_MS || 20 * 60 * 1000);
const importAfterDownload = process.env.DATADEV_IMPORT_AFTER_DOWNLOAD !== '0';
const generatedSqlDir = process.env.DATADEV_SQL_OUTPUT_DIR || path.join(projectDir, 'generated-sql');

const selectors = {
  editor: process.env.DATADEV_SQL_EDITOR_SELECTOR || '',
  runButton: process.env.DATADEV_RUN_BUTTON_SELECTOR || '',
  exportButton: process.env.DATADEV_EXPORT_BUTTON_SELECTOR || '',
};

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const ensureDownloadDir = async () => {
  await fs.mkdir(downloadDir, { recursive: true });
};

const loadSql = async () => {
  if (sqlText.trim()) return sqlText.trim();
  if (sqlFile.trim()) return fs.readFile(path.resolve(sqlFile), 'utf8');
  return '';
};

const listGeneratedSql = async () => {
  const entries = await fs.readdir(generatedSqlDir, { withFileTypes: true }).catch(() => []);
  const files = await Promise.all(
    entries
      .filter((entry) => entry.isFile())
      .filter((entry) => entry.name.endsWith('.sql'))
      .map(async (entry) => {
        const filePath = path.join(generatedSqlDir, entry.name);
        const stat = await fs.stat(filePath);
        return { filePath, mtimeMs: stat.mtimeMs };
      }),
  );

  return files.sort((a, b) => b.mtimeMs - a.mtimeMs).slice(0, 4);
};

const newestExcelFile = async (afterTimeMs) => {
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

  return files
    .filter((file) => file.mtimeMs >= afterTimeMs)
    .sort((a, b) => b.mtimeMs - a.mtimeMs)[0]?.filePath || '';
};

const importQualityFile = (filePath) =>
  new Promise((resolve, reject) => {
    const child = spawn(
      process.execPath,
      [path.join(projectDir, 'scripts', 'import-downloaded-quality.mjs'), filePath],
      {
        cwd: projectDir,
        env: {
          ...process.env,
          QUALITY_DOWNLOAD_DIR: downloadDir,
        },
        stdio: 'inherit',
      },
    );

    child.on('exit', (code) => {
      if (code === 0) resolve();
      else reject(new Error(`导入脚本退出，code=${code}`));
    });
  });

const waitForManualDownload = async (page, startedAt) => {
  console.log('等待导出下载。你可以在网页中手动执行 SQL 并点击导出；脚本会自动捕获下载文件。');

  const downloadPromise = page.waitForEvent('download', { timeout: timeoutMs }).catch(() => null);
  const pollingPromise = (async () => {
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      const filePath = await newestExcelFile(startedAt);
      if (filePath) return filePath;
      await sleep(3000);
    }
    return '';
  })();

  const download = await Promise.race([downloadPromise, pollingPromise]);
  if (typeof download === 'string') return download;
  if (!download) return '';

  const suggestedName = download.suggestedFilename();
  const targetPath = path.join(downloadDir, suggestedName);
  await download.saveAs(targetPath);
  return targetPath;
};

const tryConfiguredAutomation = async (page, sql) => {
  if (!selectors.editor || !selectors.runButton || !selectors.exportButton || !sql) {
    return '';
  }

  console.log('检测到页面选择器配置，尝试自动填 SQL、执行并导出。');
  await page.locator(selectors.editor).click({ timeout: 15000 });
  await page.keyboard.press('Meta+A').catch(async () => page.keyboard.press('Control+A'));
  await page.keyboard.insertText(sql);
  await page.locator(selectors.runButton).click({ timeout: 15000 });
  const downloadPromise = page.waitForEvent('download', { timeout: timeoutMs });
  await page.locator(selectors.exportButton).click({ timeout: timeoutMs });
  const download = await downloadPromise;
  const targetPath = path.join(downloadDir, download.suggestedFilename());
  await download.saveAs(targetPath);
  return targetPath;
};

const main = async () => {
  await ensureDownloadDir();
  const sql = await loadSql();
  const startedAt = Date.now() - 1000;

  const context = await chromium.launchPersistentContext(profileDir, {
    headless: false,
    channel: browserChannel || undefined,
    acceptDownloads: true,
    downloadsPath: downloadDir,
    viewport: { width: 1440, height: 920 },
  });

  const page = context.pages()[0] || await context.newPage();
  console.log(`打开数仓页面：${datadevUrl}`);
  const generatedSqlFiles = await listGeneratedSql();
  if (generatedSqlFiles.length) {
    console.log('最近生成的 SQL 文件：');
    generatedSqlFiles.forEach((file) => console.log(`- ${file.filePath}`));
  } else {
    console.log('尚未发现 generated-sql 文件。可先执行：npm run sql:datadev');
  }
  await page.goto(datadevUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });

  console.log('如果页面要求登录，请在打开的浏览器中手动完成登录。登录完成后脚本会继续等待导出。');
  await page.bringToFront();

  const configuredDownloadedPath = await tryConfiguredAutomation(page, sql).catch((error) => {
    console.warn(`自动执行选择器流程失败，将切换为手动辅助模式：${error.message}`);
    return '';
  });

  const downloadedPath = configuredDownloadedPath || await waitForManualDownload(page, startedAt);

  if (!downloadedPath) {
    await context.close();
    throw new Error('等待下载超时。请确认已导出 Excel，或调大 DATADEV_WAIT_TIMEOUT_MS。');
  }

  console.log(`已检测到下载文件：${downloadedPath}`);
  await context.close();

  if (importAfterDownload) {
    await importQualityFile(downloadedPath);
  }
};

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
