#!/usr/bin/env node
import path from 'node:path';
import process from 'node:process';
import { spawnSync } from 'node:child_process';
import dotenv from 'dotenv';

const projectDir = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
dotenv.config({ path: path.join(projectDir, '.mail-reports.env'), quiet: true });
dotenv.config({ quiet: true });

const baseDir = path.resolve(
  process.env.MAIL_REPORT_DOWNLOAD_DIR ||
    path.join(process.env.HOME || '.', 'Downloads', '预质检每日数据'),
);
const syncToEcs = process.env.MAIL_REPORT_SYNC_TO_ECS === '1';
const skipFetch = process.env.MAIL_REPORT_SKIP_FETCH === '1';
const dryRun = process.env.MAIL_REPORT_DRY_RUN === '1';

const runNodeScript = (scriptName, env = {}) => {
  const result = spawnSync(
    process.execPath,
    [path.join(projectDir, 'scripts', scriptName)],
    {
      cwd: projectDir,
      env: {
        ...process.env,
        ...env,
      },
      stdio: 'inherit',
    },
  );

  if (result.status !== 0) {
    throw new Error(`${scriptName} 执行失败，退出码 ${result.status}。可单独执行：node scripts/${scriptName} 查看更完整日志。`);
  }
};

if (!skipFetch) {
  runNodeScript('fetch-mail-reports.mjs');
}

const apiTargets = syncToEcs
  ? 'http://127.0.0.1:3000,http://39.107.221.251:3000'
  : 'http://127.0.0.1:3000';

runNodeScript('import-downloaded-quality.mjs', {
  QUALITY_DOWNLOAD_DIR: path.join(baseDir, 'quality'),
  QUALITY_API_BASE_URLS: apiTargets,
  QUALITY_IMPORT_DRY_RUN: dryRun ? '1' : '',
});

runNodeScript('import-downloaded-efficiency.mjs', {
  EFFICIENCY_DOWNLOAD_DIR: path.join(baseDir, 'efficiency'),
  EFFICIENCY_API_BASE_URLS: apiTargets,
  EFFICIENCY_IMPORT_DRY_RUN: dryRun ? '1' : '',
});
