#!/usr/bin/env node
import fs from 'node:fs/promises';
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
const inboxDir = path.join(baseDir, 'inbox');
const qualityDir = path.join(baseDir, 'quality');
const efficiencyDir = path.join(baseDir, 'efficiency');
const lookbackDays = Number(process.env.MAIL_REPORT_LOOKBACK_DAYS || 3);
const mailboxName = process.env.APPLE_MAILBOX_NAME || '';
const accountName = process.env.APPLE_MAIL_ACCOUNT_NAME || '';
const qualitySubject = process.env.MAIL_REPORT_QUALITY_SUBJECT || '看板_基础数据';
const efficiencySubject = process.env.MAIL_REPORT_EFFICIENCY_SUBJECT || '看板_人效';
const dryRun = process.env.MAIL_REPORT_DRY_RUN === '1';

const ensureDirs = async () => {
  await fs.mkdir(inboxDir, { recursive: true });
  await fs.mkdir(qualityDir, { recursive: true });
  await fs.mkdir(efficiencyDir, { recursive: true });
};

const q = (value) => JSON.stringify(String(value));

const buildAppleScript = () => `
set baseDir to ${q(baseDir)}
set inboxDir to ${q(inboxDir)}
set qualityDir to ${q(qualityDir)}
set efficiencyDir to ${q(efficiencyDir)}
set lookbackDays to ${lookbackDays}
set mailboxName to ${q(mailboxName)}
set accountName to ${q(accountName)}
set qualitySubject to ${q(qualitySubject)}
set efficiencySubject to ${q(efficiencySubject)}
set dryRun to ${dryRun ? 'true' : 'false'}
set sinceDate to (current date) - (lookbackDays * days)
set savedLines to {}

on sanitizeFileName(fileName)
  set AppleScript's text item delimiters to ":"
  set parts to text items of fileName
  set AppleScript's text item delimiters to "_"
  set fileName to parts as text
  set AppleScript's text item delimiters to "/"
  set parts to text items of fileName
  set AppleScript's text item delimiters to "_"
  set fileName to parts as text
  set AppleScript's text item delimiters to ""
  return fileName
end sanitizeFileName

on isExcelFile(fileName)
  set lowerName to do shell script "printf %s " & quoted form of fileName & " | tr '[:upper:]' '[:lower:]'"
  return lowerName ends with ".xlsx" or lowerName ends with ".xls"
end isExcelFile

on saveOneAttachment(theAttachment, targetDir, inboxDir, reportType, messageId, dryRun)
  set originalName to my sanitizeFileName(name of theAttachment)
  if not my isExcelFile(originalName) then return ""
  set targetPath to targetDir & "/" & reportType & "_" & messageId & "_" & originalName
  set inboxPath to inboxDir & "/" & reportType & "_" & messageId & "_" & originalName
  if dryRun is false then
    try
      do shell script "test -f " & quoted form of targetPath
      return ""
    on error
      save theAttachment in POSIX file targetPath
      do shell script "cp " & quoted form of targetPath & " " & quoted form of inboxPath
    end try
  end if
  return reportType & tab & targetPath
end saveOneAttachment

tell application "Mail"
  if accountName is not "" then
    if mailboxName is not "" then
      set targetMailbox to mailbox mailboxName of account accountName
    else
      set targetMailbox to inbox of account accountName
    end if
  else
    if mailboxName is not "" then
      set targetMailbox to mailbox mailboxName
    else
      set targetMailbox to inbox
    end if
  end if

  set reportMessages to messages of targetMailbox whose date received ≥ sinceDate
  repeat with theMessage in reportMessages
    set theSubject to subject of theMessage
    set reportType to ""
    set targetDir to ""
    if theSubject contains qualitySubject then
      set reportType to "quality"
      set targetDir to qualityDir
    else if theSubject contains efficiencySubject then
      set reportType to "efficiency"
      set targetDir to efficiencyDir
    end if

    if reportType is not "" then
      set messageId to id of theMessage as text
      repeat with theAttachment in mail attachments of theMessage
        set savedLine to my saveOneAttachment(theAttachment, targetDir, inboxDir, reportType, messageId, dryRun)
        if savedLine is not "" then set end of savedLines to savedLine
      end repeat
    end if
  end repeat
end tell

set AppleScript's text item delimiters to linefeed
return savedLines as text
`;

const main = async () => {
  await ensureDirs();
  const result = spawnSync('osascript', ['-e', buildAppleScript()], {
    encoding: 'utf8',
    maxBuffer: 1024 * 1024 * 5,
  });

  if (result.status !== 0) {
    throw new Error(result.stderr || result.stdout || `osascript 退出码 ${result.status}`);
  }

  const saved = result.stdout
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const [type, filePath] = line.split('\t');
      return { type, filePath };
    });

  console.log(JSON.stringify({
    dryRun,
    source: 'Apple Mail',
    mailboxName: mailboxName || '(inbox)',
    accountName: accountName || '(default)',
    lookbackDays,
    baseDir,
    savedCount: saved.length,
    saved,
  }, null, 2));
};

main().catch((error) => {
  console.error(error instanceof Error ? error.stack || error.message : String(error));
  process.exit(1);
});
