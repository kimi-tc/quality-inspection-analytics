#!/usr/bin/env node
import fs from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import { spawnSync } from 'node:child_process';

const projectDir = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
const templateDir = path.join(projectDir, 'sql', 'templates');
const outputDir = process.env.DATADEV_SQL_OUTPUT_DIR || path.join(projectDir, 'generated-sql');

const pad = (value) => String(value).padStart(2, '0');
const formatDate = (date) => `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
const parseDate = (value) => {
  const match = String(value || '').match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!match) {
    throw new Error(`日期格式错误：${value}，请使用 yyyy-MM-dd`);
  }

  const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  if (Number.isNaN(date.getTime())) {
    throw new Error(`日期无效：${value}`);
  }

  return date;
};
const addDays = (date, days) => new Date(date.getFullYear(), date.getMonth(), date.getDate() + days);
const yesterday = () => {
  const now = new Date();
  return addDays(new Date(now.getFullYear(), now.getMonth(), now.getDate()), -1);
};

const args = process.argv.slice(2);
const readArg = (name) => {
  const index = args.indexOf(name);
  return index >= 0 ? args[index + 1] : '';
};
const hasArg = (name) => args.includes(name);

const type = readArg('--type') || 'both';
const dateArg = readArg('--date');
const startArg = readArg('--start');
const endArg = readArg('--end');
const copyType = readArg('--copy');
const shouldPrint = hasArg('--print');

if (!['quality', 'efficiency', 'both'].includes(type)) {
  throw new Error('--type 仅支持 quality、efficiency、both');
}

const startDate = dateArg ? parseDate(dateArg) : startArg ? parseDate(startArg) : yesterday();
const endDate = dateArg ? parseDate(dateArg) : endArg ? parseDate(endArg) : startDate;

if (endDate < startDate) {
  throw new Error('结束日期不能早于开始日期');
}

const variables = {
  START_DATE: formatDate(startDate),
  END_DATE: formatDate(endDate),
  END_EXCLUSIVE: formatDate(addDays(endDate, 1)),
};

const renderTemplate = async (templateName) => {
  const templatePath = path.join(templateDir, `${templateName}.sql.template`);
  const template = await fs.readFile(templatePath, 'utf8');
  return template.replace(/\{\{(START_DATE|END_DATE|END_EXCLUSIVE)\}\}/g, (_, key) => variables[key]);
};

const writeSql = async (templateName) => {
  const sql = await renderTemplate(templateName);
  await fs.mkdir(outputDir, { recursive: true });
  const fileName = `${templateName}_${variables.START_DATE.replace(/-/g, '')}_${variables.END_DATE.replace(/-/g, '')}.sql`;
  const outputPath = path.join(outputDir, fileName);
  await fs.writeFile(outputPath, sql, 'utf8');
  return { templateName, outputPath, sql };
};

const copyToClipboard = (text) => {
  const result = spawnSync('pbcopy', { input: text });
  if (result.status !== 0) {
    throw new Error('复制到剪贴板失败，请确认当前系统支持 pbcopy。');
  }
};

const main = async () => {
  const targets = type === 'both' ? ['quality', 'efficiency'] : [type];
  const outputs = [];

  for (const target of targets) {
    outputs.push(await writeSql(target));
  }

  if (copyType) {
    const matched = outputs.find((item) => item.templateName === copyType);
    if (!matched) {
      throw new Error(`--copy ${copyType} 不在本次生成范围内。`);
    }

    copyToClipboard(matched.sql);
  }

  if (shouldPrint && outputs.length === 1) {
    console.log(outputs[0].sql);
    return;
  }

  console.log(JSON.stringify({
    dateRange: `${variables.START_DATE} ~ ${variables.END_DATE}`,
    endExclusive: variables.END_EXCLUSIVE,
    copied: copyType || '',
    files: outputs.map((item) => item.outputPath),
  }, null, 2));
};

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
