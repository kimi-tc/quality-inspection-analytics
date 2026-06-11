#!/usr/bin/env node
import fs from 'node:fs/promises';
import path from 'node:path';
import process from 'node:process';
import dotenv from 'dotenv';
import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';

const projectDir = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
dotenv.config({ path: path.join(projectDir, '.mail-reports.env'), quiet: true });
dotenv.config({ quiet: true });

const requiredEnv = ['FEISHU_IMAP_HOST', 'FEISHU_EMAIL', 'FEISHU_EMAIL_AUTH_CODE'];
const missingEnv = requiredEnv.filter((key) => !process.env[key]);
if (missingEnv.length) {
  console.error(`缺少邮箱配置：${missingEnv.join(', ')}。请在 .mail-reports.env 中配置。`);
  process.exit(1);
}

const baseDir = path.resolve(
  process.env.MAIL_REPORT_DOWNLOAD_DIR ||
    path.join(process.env.HOME || '.', 'Downloads', '预质检每日数据'),
);
const inboxDir = path.join(baseDir, 'inbox');
const qualityDir = path.join(baseDir, 'quality');
const efficiencyDir = path.join(baseDir, 'efficiency');
const stateFile = path.join(baseDir, '.mail-report-state.json');
const mailbox = process.env.FEISHU_IMAP_MAILBOX || 'INBOX';
const lookbackDays = Number(process.env.MAIL_REPORT_LOOKBACK_DAYS || 3);
const dryRun = process.env.MAIL_REPORT_DRY_RUN === '1';

const reportRules = [
  {
    type: 'quality',
    subjectKeyword: process.env.MAIL_REPORT_QUALITY_SUBJECT || '看板_基础数据',
    targetDir: qualityDir,
  },
  {
    type: 'efficiency',
    subjectKeyword: process.env.MAIL_REPORT_EFFICIENCY_SUBJECT || '看板_人效',
    targetDir: efficiencyDir,
  },
];

const sanitizeFileName = (value) =>
  String(value || 'attachment.xlsx')
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, ' ')
    .trim();

const readState = async () => {
  try {
    const content = await fs.readFile(stateFile, 'utf8');
    const parsed = JSON.parse(content);
    return {
      downloaded: Array.isArray(parsed.downloaded) ? parsed.downloaded : [],
    };
  } catch {
    return { downloaded: [] };
  }
};

const writeState = async (state) => {
  await fs.mkdir(baseDir, { recursive: true });
  await fs.writeFile(stateFile, JSON.stringify(state, null, 2), 'utf8');
};

const ensureDirs = async () => {
  await fs.mkdir(inboxDir, { recursive: true });
  await fs.mkdir(qualityDir, { recursive: true });
  await fs.mkdir(efficiencyDir, { recursive: true });
};

const getSinceDate = () => {
  const date = new Date();
  date.setDate(date.getDate() - lookbackDays);
  date.setHours(0, 0, 0, 0);
  return date;
};

const classifyMessage = (subject) =>
  reportRules.find((rule) => String(subject || '').includes(rule.subjectKeyword));

const isExcelAttachment = (attachment) =>
  /\.(xlsx|xls)$/i.test(attachment.filename || '') ||
  [
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'application/vnd.ms-excel',
  ].includes(String(attachment.contentType || '').toLowerCase());

const formatDateToken = (value) => {
  const date = value instanceof Date && !Number.isNaN(value.getTime()) ? value : new Date();
  return date.toISOString().slice(0, 10);
};

const saveAttachment = async ({ messageUid, subject, messageDate, attachment, rule, state }) => {
  const originalName = sanitizeFileName(attachment.filename || `${rule.type}.xlsx`);
  const stateKey = `${messageUid}:${rule.type}:${originalName}:${attachment.size || attachment.content?.length || 0}`;
  if (state.downloaded.includes(stateKey)) {
    return { skipped: true, reason: 'already-downloaded', filePath: '' };
  }

  const fileName = `${rule.type}_${formatDateToken(messageDate)}_${messageUid}_${originalName}`;
  const targetPath = path.join(rule.targetDir, fileName);
  const inboxPath = path.join(inboxDir, fileName);

  if (!dryRun) {
    await fs.writeFile(targetPath, attachment.content);
    await fs.writeFile(inboxPath, attachment.content);
    state.downloaded.push(stateKey);
  }

  return {
    skipped: false,
    type: rule.type,
    subject,
    fileName,
    filePath: targetPath,
  };
};

const main = async () => {
  await ensureDirs();
  const state = await readState();
  const client = new ImapFlow({
    host: process.env.FEISHU_IMAP_HOST,
    port: Number(process.env.FEISHU_IMAP_PORT || 993),
    secure: process.env.FEISHU_IMAP_SECURE !== '0',
    auth: {
      user: process.env.FEISHU_EMAIL,
      pass: process.env.FEISHU_EMAIL_AUTH_CODE,
    },
    logger: false,
  });

  const saved = [];

  await client.connect();
  try {
    await client.mailboxOpen(mailbox);
    const since = getSinceDate();

    for await (const message of client.fetch({ since }, { uid: true, envelope: true, source: true })) {
      const subject = message.envelope?.subject || '';
      const rule = classifyMessage(subject);
      if (!rule) {
        continue;
      }

      const parsed = await simpleParser(message.source);
      for (const attachment of parsed.attachments || []) {
        if (!isExcelAttachment(attachment)) {
          continue;
        }

        const result = await saveAttachment({
          messageUid: message.uid,
          subject,
          messageDate: message.envelope?.date,
          attachment,
          rule,
          state,
        });

        if (!result.skipped) {
          saved.push(result);
        }
      }
    }
  } finally {
    await client.logout().catch(() => {});
  }

  if (!dryRun) {
    state.downloaded = state.downloaded.slice(-500);
    await writeState(state);
  }

  console.log(JSON.stringify({
    dryRun,
    baseDir,
    savedCount: saved.length,
    saved,
  }, null, 2));
};

main().catch((error) => {
  if (error instanceof Error) {
    console.error(error.stack || error.message);
    if ('code' in error && error.code) {
      console.error(`error.code=${error.code}`);
    }
    if ('response' in error && error.response) {
      console.error(`error.response=${error.response}`);
    }
  } else {
    console.error(String(error));
  }
  process.exit(1);
});
