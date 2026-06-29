#!/usr/bin/env node
import fs from 'node:fs';
import crypto from 'node:crypto';
import path from 'node:path';
import process from 'node:process';
import dotenv from 'dotenv';

const projectDir = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
dotenv.config({ path: path.join(projectDir, '.mail-reports.env'), quiet: true });
dotenv.config({ path: path.join(projectDir, '.env'), quiet: true });

const dataDir = path.resolve(process.env.DATA_DIR || path.join(projectDir, 'data'));
const outputDir = path.resolve(process.env.ALERT_OUTPUT_DIR || path.join(projectDir, 'reports', 'alerts'));
const apiBaseUrl = (process.env.ALERT_API_BASE_URL || '').trim().replace(/\/$/, '');
const dashboardPublicUrl = (process.env.DASHBOARD_PUBLIC_URL || '').trim();
const requestedDate = process.env.ALERT_DATE || process.argv[2] || '';
const shouldPushToFeishu = process.env.ALERT_PUSH_TO_FEISHU === '1';
const feishuWebhookUrl = (process.env.FEISHU_WEBHOOK_URL || '').trim();
const feishuWebhookSecret = (process.env.FEISHU_WEBHOOK_SECRET || '').trim();

const thresholds = {
  ambiguousTarget: Number(process.env.ALERT_AMBIGUOUS_TARGET || 0.07),
  sessionMinShare: Number(process.env.ALERT_SESSION_MIN_SHARE || 0.01),
  sessionPrecisionDropPp: Number(process.env.ALERT_SESSION_PRECISION_DROP_PP || 0.05),
  attributeMinDeclarations: Number(process.env.ALERT_ATTRIBUTE_MIN_DECLARATIONS || 20),
  attributeRejectRate: Number(process.env.ALERT_ATTRIBUTE_REJECT_RATE || 0.3),
  attributeAmbiguousRate: Number(process.env.ALERT_ATTRIBUTE_AMBIGUOUS_RATE || 0.15),
  attributePrecisionVolatilityPp: Number(process.env.ALERT_ATTRIBUTE_PRECISION_VOLATILITY_PP || 0.12),
  teamPrecisionDropPp: Number(process.env.ALERT_TEAM_PRECISION_DROP_PP || 0.05),
  auditorMinDeclarations: Number(process.env.ALERT_AUDITOR_MIN_DECLARATIONS || 20),
  auditorPrecisionDropPp: Number(process.env.ALERT_AUDITOR_PRECISION_DROP_PP || 0.08),
  auditorAmbiguousRate: Number(process.env.ALERT_AUDITOR_AMBIGUOUS_RATE || 0.15),
  auditorRejectRate: Number(process.env.ALERT_AUDITOR_REJECT_RATE || 0.3),
  employeeEfficiencyDropRate: Number(process.env.ALERT_EMPLOYEE_EFFICIENCY_DROP_RATE || 0.2),
  employeePrecisionMovePp: Number(process.env.ALERT_EMPLOYEE_PRECISION_MOVE_PP || 0.08),
  employeeVolumeMinWeighted: Number(process.env.ALERT_EMPLOYEE_VOLUME_MIN_WEIGHTED || 50),
  employeeVolumeBaselineIncreaseRate: Number(process.env.ALERT_EMPLOYEE_VOLUME_BASELINE_INCREASE_RATE || 0.35),
  employeeVolumePeerRatio: Number(process.env.ALERT_EMPLOYEE_VOLUME_PEER_RATIO || 1.5),
};

const readJson = (filePath, fallback) => {
  try {
    return JSON.parse(fs.readFileSync(filePath, 'utf8'));
  } catch {
    return fallback;
  }
};

const fetchJson = async (endpoint, fallbackPath, fallback) => {
  if (apiBaseUrl) {
    const response = await fetch(`${apiBaseUrl}${endpoint}`, {
      signal: AbortSignal.timeout(Number(process.env.ALERT_API_TIMEOUT_MS || 120000)),
    });
    if (!response.ok) {
      throw new Error(`读取 ${apiBaseUrl}${endpoint} 失败，HTTP ${response.status}`);
    }
    return response.json();
  }

  return readJson(fallbackPath, fallback);
};

const propertyDictionary = readJson(path.join(dataDir, 'property-category-dictionary.json'), []);
const auditorTeamDictionary = readJson(path.join(dataDir, 'auditor-team-dictionary.json'), []);

const propertyCategoryByName = new Map(
  (Array.isArray(propertyDictionary) ? propertyDictionary : [])
    .map((entry) => [String(entry.propertyName ?? '').trim(), String(entry.category ?? '').trim()])
    .filter(([propertyName, category]) => propertyName && category),
);

const auditorTeamByName = new Map(
  (Array.isArray(auditorTeamDictionary) ? auditorTeamDictionary : [])
    .map((entry) => [String(entry.auditorName ?? '').trim(), String(entry.team ?? '').trim()])
    .filter(([auditorName, team]) => auditorName && team),
);

const normalizeDate = (value) => String(value ?? '').slice(0, 10);
const formatInt = (value) => Math.round(Number(value || 0)).toLocaleString('zh-CN');
const formatPct = (value, digits = 1) => `${(Number(value || 0) * 100).toFixed(digits)}%`;
const formatPp = (value, digits = 1) => `${value >= 0 ? '+' : ''}${(Number(value || 0) * 100).toFixed(digits)}pp`;

const metricEmpty = () => ({
  declarations: 0,
  precisePasses: 0,
  ambiguousPasses: 0,
  rejects: 0,
  proofRejects: 0,
});

const addMetric = (metric, row) => {
  metric.declarations += row.declarations;
  metric.precisePasses += row.precisePasses;
  metric.ambiguousPasses += row.ambiguousPasses;
  metric.rejects += row.rejects;
  metric.proofRejects += row.proofRejects;
  return metric;
};

const withRates = (metric) => ({
  ...metric,
  preciseRate: metric.declarations ? metric.precisePasses / metric.declarations : 0,
  ambiguousRate: metric.declarations ? metric.ambiguousPasses / metric.declarations : 0,
  rejectRate: metric.declarations ? metric.rejects / metric.declarations : 0,
  proofAccuracyRate: metric.declarations
    ? (metric.declarations - metric.ambiguousPasses - metric.proofRejects) / metric.declarations
    : 0,
});

const normalizeQualityRow = (row) => {
  const declarations = Number(row.declarations || 0);
  let precisePasses = Number(row.exactPasses);
  if (!Number.isFinite(precisePasses)) {
    precisePasses = Math.max(0, declarations - Number(row.ambiguousPasses || 0) - Number(row.rejects || 0));
  }

  const auditor = String(row.auditor ?? row.auditorName ?? '').trim();
  const attribute = String(row.attribute ?? row.propertyTag ?? row.tag ?? '').trim();

  return {
    date: normalizeDate(row.date),
    auditor,
    auditorTeam: String(row.auditorTeam ?? row.team ?? '').trim() || auditorTeamByName.get(auditor) || '未标记',
    session: String(row.session ?? row.saleType ?? '').trim() || '未标记',
    batch: String(row.batch ?? row.batchFlag ?? '').trim() || '未标记',
    category: String(row.category ?? row.attributeCategory ?? '').trim() || propertyCategoryByName.get(attribute) || '未标记',
    attribute: attribute || '未标记',
    declarations,
    precisePasses,
    ambiguousPasses: Number(row.ambiguousPasses || 0),
    rejects: Number(row.rejects || 0),
    proofRejects: Number(row.proofRejects || 0),
  };
};

const normalizeEfficiencyRow = (row) => ({
  date: normalizeDate(row.date),
  employee: String(row.employee ?? row.auditor ?? '').trim() || '未标记',
  team: String(row.team ?? '').trim() || '未标记',
  handledCount: Number(row.handledCount || row.totalAuditCount || 0),
  weightedHandledCount: Number(row.weightedHandledCount || 0),
  firstAuditCount: Number(row.firstAuditCount || 0),
  precisionPassCount: Number(row.precisionPassCount || row.precisePassCount || 0),
});

const groupBy = (rows, keySelector) => {
  const groups = new Map();
  for (const row of rows) {
    const key = keySelector(row) || '未标记';
    if (!groups.has(key)) groups.set(key, metricEmpty());
    addMetric(groups.get(key), row);
  }
  return [...groups.entries()].map(([key, metric]) => ({ key, ...withRates(metric) }));
};

const isConfiguredTeam = (team) => {
  const normalized = String(team ?? '').trim();
  return Boolean(normalized) && normalized !== '未配置团队' && normalized !== '未标记';
};

const auditorKey = (row) => `${row.auditor}\u0000${row.auditorTeam}`;

const groupAuditorQuality = (rows) => {
  const groups = new Map();
  for (const row of rows) {
    if (!row.auditor || !isConfiguredTeam(row.auditorTeam)) continue;
    const key = auditorKey(row);
    if (!groups.has(key)) {
      groups.set(key, {
        key,
        auditor: row.auditor,
        team: row.auditorTeam,
        ...metricEmpty(),
      });
    }
    addMetric(groups.get(key), row);
  }

  return [...groups.values()].map((group) => ({
    ...group,
    ...withRates(group),
  }));
};

const summarize = (rows) => withRates(rows.reduce(addMetric, metricEmpty()));

const averageMetricByDate = (rows, dates, keySelector) => {
  const targetDates = new Set(dates);
  const metricsByKeyDate = new Map();

  for (const row of rows) {
    if (!targetDates.has(row.date)) continue;
    const key = keySelector(row) || '未标记';
    const mapKey = `${key}\u0000${row.date}`;
    if (!metricsByKeyDate.has(mapKey)) metricsByKeyDate.set(mapKey, metricEmpty());
    addMetric(metricsByKeyDate.get(mapKey), row);
  }

  const valuesByKey = new Map();
  for (const [mapKey, metric] of metricsByKeyDate.entries()) {
    const [key] = mapKey.split('\u0000');
    if (!valuesByKey.has(key)) valuesByKey.set(key, []);
    valuesByKey.get(key).push(withRates(metric));
  }

  const result = new Map();
  for (const [key, values] of valuesByKey.entries()) {
    const sum = values.reduce((acc, item) => {
      acc.declarations += item.declarations;
      acc.preciseRate += item.preciseRate;
      acc.ambiguousRate += item.ambiguousRate;
      acc.rejectRate += item.rejectRate;
      acc.proofAccuracyRate += item.proofAccuracyRate;
      return acc;
    }, { declarations: 0, preciseRate: 0, ambiguousRate: 0, rejectRate: 0, proofAccuracyRate: 0 });
    result.set(key, {
      declarations: sum.declarations,
      days: values.length,
      preciseRate: sum.preciseRate / values.length,
      ambiguousRate: sum.ambiguousRate / values.length,
      rejectRate: sum.rejectRate / values.length,
      proofAccuracyRate: sum.proofAccuracyRate / values.length,
    });
  }

  return result;
};

const averageEfficiencyByEmployee = (rows, dates) => {
  const targetDates = new Set(dates);
  const groups = new Map();
  for (const row of rows) {
    if (!targetDates.has(row.date)) continue;
    const key = row.employee;
    if (!groups.has(key)) {
      groups.set(key, {
        employee: key,
        team: row.team,
        days: new Set(),
        weightedHandledCount: 0,
        firstAuditCount: 0,
        precisionPassCount: 0,
      });
    }
    const group = groups.get(key);
    group.days.add(row.date);
    group.weightedHandledCount += row.weightedHandledCount;
    group.firstAuditCount += row.firstAuditCount;
    group.precisionPassCount += row.precisionPassCount;
  }

  const result = new Map();
  for (const [key, group] of groups.entries()) {
    result.set(key, {
      employee: group.employee,
      team: group.team,
      days: group.days.size,
      averageWeightedHandledCount: group.days.size ? group.weightedHandledCount / group.days.size : 0,
      precisionPassRate: group.firstAuditCount ? group.precisionPassCount / group.firstAuditCount : 0,
    });
  }
  return result;
};

const median = (values) => {
  const sorted = values
    .map((value) => Number(value || 0))
    .filter((value) => Number.isFinite(value) && value > 0)
    .sort((a, b) => a - b);

  if (!sorted.length) return 0;
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2
    ? sorted[mid]
    : (sorted[mid - 1] + sorted[mid]) / 2;
};

const alert = (level, type, title, detail, payload = {}) => ({
  level,
  type,
  title,
  detail,
  ...payload,
});

const levelRank = { high: 3, medium: 2, low: 1, info: 0 };

const buildMessage = ({ targetDate, baselineDates, total, previousTotal, alerts, dataSource }) => {
  const highAlerts = alerts.filter((item) => item.level === 'high');
  const mediumAlerts = alerts.filter((item) => item.level === 'medium');
  const levelLabel = (level) => (level === 'high' ? '高' : level === 'medium' ? '中' : '低');
  const formatAlertLine = (item, index) => `${index + 1}. 【${levelLabel(item.level)}】${item.title}：${item.detail}`;
  const alertSections = [
    {
      title: '一、整体预警',
      matcher: (item) => item.type.startsWith('overall_'),
      limit: 3,
    },
    {
      title: '二、场次/团队预警',
      matcher: (item) => item.type.startsWith('session_') || item.type.startsWith('team_'),
      limit: 5,
    },
    {
      title: '三、属性项预警',
      matcher: (item) => item.type.startsWith('attribute_'),
      limit: 6,
    },
    {
      title: '四、审核人预警',
      matcher: (item) => item.type.startsWith('auditor_'),
      limit: 5,
    },
    {
      title: '五、人效异动',
      matcher: (item) => item.type.startsWith('employee_'),
      limit: 3,
    },
  ];

  const lines = [
    `【预质检每日预警】${targetDate}`,
    `数据源：${dataSource}`,
    `申报 ${formatInt(total.declarations)} 次｜精准通过率 ${formatPct(total.preciseRate)}｜举证准确率 ${formatPct(total.proofAccuracyRate)}｜模棱两可率 ${formatPct(total.ambiguousRate)}｜拒绝率 ${formatPct(total.rejectRate)}`,
    `较上一有数日：精准通过率 ${formatPp(total.preciseRate - previousTotal.preciseRate)}，模棱两可率 ${formatPp(total.ambiguousRate - previousTotal.ambiguousRate)}，拒绝率 ${formatPp(total.rejectRate - previousTotal.rejectRate)}`,
    `预警数：高 ${highAlerts.length} / 中 ${mediumAlerts.length} / 总 ${alerts.length}`,
  ];

  if (baselineDates.length) {
    lines.push(`基准窗口：${baselineDates[0]} ~ ${baselineDates.at(-1)}`);
  }

  if (!alerts.length) {
    lines.push('');
    lines.push('今日未触发核心预警，建议保持常规抽检。');
  } else {
    for (const section of alertSections) {
      const sectionAlerts = alerts.filter(section.matcher).slice(0, section.limit);
      if (!sectionAlerts.length) continue;

      lines.push('');
      lines.push(section.title);
      sectionAlerts.forEach((item, index) => {
        lines.push(formatAlertLine(item, index));
      });
    }
  }

  lines.push('');
  lines.push('建议动作：优先抽看高拒绝/高模糊属性项样本，再按场次和团队确认是否集中在某个审核口径。');
  return lines.join('\n');
};

const buildFeishuPayload = (text) => {
  const isFlowWebhook = feishuWebhookUrl.includes('/flow/api/trigger-webhook/');
  if (isFlowWebhook) {
    return {
      text,
      title: `预质检每日预警`,
      content: {
        text,
      },
    };
  }

  const payload = {
    msg_type: 'text',
    content: {
      text,
    },
  };

  if (!feishuWebhookSecret) {
    return payload;
  }

  const timestamp = Math.floor(Date.now() / 1000).toString();
  const stringToSign = `${timestamp}\n${feishuWebhookSecret}`;
  const sign = crypto
    .createHmac('sha256', stringToSign)
    .update('')
    .digest('base64');

  return {
    ...payload,
    timestamp,
    sign,
  };
};

const pushToFeishu = async (text) => {
  if (!shouldPushToFeishu) {
    return {
      skipped: true,
      reason: 'ALERT_PUSH_TO_FEISHU is not 1',
    };
  }

  if (!feishuWebhookUrl) {
    if (process.env.ALERT_PUSH_STRICT === '1') {
      throw new Error('已启用飞书推送，但未配置 FEISHU_WEBHOOK_URL');
    }

    return {
      skipped: true,
      reason: 'FEISHU_WEBHOOK_URL is not configured',
    };
  }

  const response = await fetch(feishuWebhookUrl, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(buildFeishuPayload(text)),
    signal: AbortSignal.timeout(Number(process.env.FEISHU_WEBHOOK_TIMEOUT_MS || 30000)),
  });

  const bodyText = await response.text();
  let body = {};
  try {
    body = JSON.parse(bodyText);
  } catch {
    body = { raw: bodyText };
  }

  if (!response.ok || (body && typeof body === 'object' && 'code' in body && body.code !== 0)) {
    throw new Error(`飞书推送失败，HTTP ${response.status}：${bodyText}`);
  }

  return {
    skipped: false,
    status: response.status,
    body,
  };
};

const main = async () => {
  const qualityDataset = await fetchJson('/api/dataset', path.join(dataDir, 'shared-dataset.json'), { rows: [] });
  const efficiencyDataset = await fetchJson('/api/efficiency-dataset', path.join(dataDir, 'efficiency-dataset.json'), { rows: [] });

  const qualityRows = (qualityDataset.rows || [])
    .map(normalizeQualityRow)
    .filter((row) => row.date && row.declarations > 0);
  const efficiencyRows = (efficiencyDataset.rows || [])
    .map(normalizeEfficiencyRow)
    .filter((row) => row.date);

  const dates = [...new Set(qualityRows.map((row) => row.date))].sort();
  const targetDate = requestedDate || dates.at(-1);
  if (!targetDate) throw new Error('没有可分析的质量数据');

  const targetRows = qualityRows.filter((row) => row.date === targetDate);
  if (!targetRows.length) throw new Error(`指定日期没有质量数据：${targetDate}`);

  const previousDate = dates.filter((date) => date < targetDate).at(-1) || '';
  const previousRows = previousDate ? qualityRows.filter((row) => row.date === previousDate) : [];
  const baselineDates = dates.filter((date) => date < targetDate).slice(-7);

  const total = summarize(targetRows);
  const previousTotal = summarize(previousRows);
  const baselineBySession = averageMetricByDate(qualityRows, baselineDates, (row) => row.session);
  const baselineByTeam = averageMetricByDate(qualityRows, baselineDates, (row) => row.auditorTeam);

  const alerts = [];

  if (total.ambiguousRate > thresholds.ambiguousTarget) {
    alerts.push(alert(
      'high',
      'overall_ambiguous_rate',
      '整体模棱两可率高于目标',
      `当前 ${formatPct(total.ambiguousRate)}，目标 ${formatPct(thresholds.ambiguousTarget)}，超出 ${formatPp(total.ambiguousRate - thresholds.ambiguousTarget)}`,
      { currentRate: total.ambiguousRate, targetRate: thresholds.ambiguousTarget, declarations: total.declarations },
    ));
  }

  const sessionRows = groupBy(targetRows, (row) => row.session)
    .map((row) => ({ ...row, share: total.declarations ? row.declarations / total.declarations : 0 }))
    .filter((row) => row.share >= thresholds.sessionMinShare);

  for (const row of sessionRows) {
    const baseline = baselineBySession.get(row.key);
    if (baseline && baseline.days >= 2 && baseline.declarations >= thresholds.attributeMinDeclarations) {
      const drop = baseline.preciseRate - row.preciseRate;
      if (drop >= thresholds.sessionPrecisionDropPp) {
        alerts.push(alert(
          'high',
          'session_precision_drop',
          `场次精准通过率下降：${row.key}`,
          `当前 ${formatPct(row.preciseRate)}，近 ${baseline.days} 日均值 ${formatPct(baseline.preciseRate)}，下降 ${formatPp(-drop)}，申报占比 ${formatPct(row.share)}`,
          { object: row.key, currentRate: row.preciseRate, baselineRate: baseline.preciseRate, declarations: row.declarations, share: row.share },
        ));
      }
    }
  }

  const attributeRows = groupBy(targetRows, (row) => row.attribute)
    .filter((row) => row.declarations >= thresholds.attributeMinDeclarations);

  for (const row of attributeRows) {
    if (row.rejectRate >= thresholds.attributeRejectRate) {
      alerts.push(alert(
        'high',
        'attribute_high_reject',
        `属性项拒绝率高：${row.key}`,
        `拒绝率 ${formatPct(row.rejectRate)}，申报 ${formatInt(row.declarations)} 次`,
        { object: row.key, currentRate: row.rejectRate, declarations: row.declarations },
      ));
    }

    if (row.ambiguousRate >= thresholds.attributeAmbiguousRate) {
      alerts.push(alert(
        row.ambiguousRate >= thresholds.attributeAmbiguousRate * 1.5 ? 'high' : 'medium',
        'attribute_high_ambiguous',
        `属性项模棱两可率高：${row.key}`,
        `模棱两可率 ${formatPct(row.ambiguousRate)}，申报 ${formatInt(row.declarations)} 次`,
        { object: row.key, currentRate: row.ambiguousRate, declarations: row.declarations },
      ));
    }
  }

  const attributeByPrevious = new Map(groupBy(previousRows, (row) => row.attribute).map((row) => [row.key, row]));
  for (const row of attributeRows) {
    const previous = attributeByPrevious.get(row.key);
    if (!previous || previous.declarations < thresholds.attributeMinDeclarations) continue;
    const movement = row.preciseRate - previous.preciseRate;
    if (Math.abs(movement) >= thresholds.attributePrecisionVolatilityPp) {
      alerts.push(alert(
        'medium',
        'attribute_precision_volatility',
        `属性项精准通过率波动：${row.key}`,
        `当前 ${formatPct(row.preciseRate)}，上一有数日 ${formatPct(previous.preciseRate)}，变化 ${formatPp(movement)}`,
        { object: row.key, currentRate: row.preciseRate, previousRate: previous.preciseRate, movement, declarations: row.declarations },
      ));
    }
  }

  const teamRows = groupBy(targetRows, (row) => row.auditorTeam);
  for (const row of teamRows) {
    if (!isConfiguredTeam(row.key)) continue;

    const baseline = baselineByTeam.get(row.key);
    if (baseline && baseline.days >= 2 && baseline.declarations >= thresholds.attributeMinDeclarations) {
      const drop = baseline.preciseRate - row.preciseRate;
      if (drop >= thresholds.teamPrecisionDropPp) {
        alerts.push(alert(
          'high',
          'team_precision_drop',
          `团队精准通过率下降：${row.key}`,
          `当前 ${formatPct(row.preciseRate)}，近 ${baseline.days} 日均值 ${formatPct(baseline.preciseRate)}，下降 ${formatPp(-drop)}`,
          { object: row.key, currentRate: row.preciseRate, baselineRate: baseline.preciseRate, declarations: row.declarations },
        ));
      }
    }
  }

  const currentAuditorRows = groupAuditorQuality(targetRows)
    .filter((row) => row.declarations >= thresholds.auditorMinDeclarations);
  const previousAuditorMap = new Map(
    groupAuditorQuality(previousRows)
      .filter((row) => row.declarations >= thresholds.auditorMinDeclarations)
      .map((row) => [row.key, row]),
  );

  for (const row of currentAuditorRows) {
    const previous = previousAuditorMap.get(row.key);
    if (previous) {
      const drop = previous.preciseRate - row.preciseRate;
      if (drop >= thresholds.auditorPrecisionDropPp) {
        alerts.push(alert(
          'medium',
          'auditor_precision_drop',
          `审核人精准通过率下降：${row.auditor}`,
          `${row.team}｜当前 ${formatPct(row.preciseRate)}，上一有数日 ${formatPct(previous.preciseRate)}，下降 ${formatPp(-drop)}，申报 ${formatInt(row.declarations)} 次`,
          {
            object: row.auditor,
            team: row.team,
            currentRate: row.preciseRate,
            previousRate: previous.preciseRate,
            movement: row.preciseRate - previous.preciseRate,
            declarations: row.declarations,
          },
        ));
      }
    }

    if (row.ambiguousRate >= thresholds.auditorAmbiguousRate) {
      alerts.push(alert(
        row.ambiguousRate >= thresholds.auditorAmbiguousRate * 1.5 ? 'high' : 'medium',
        'auditor_high_ambiguous',
        `审核人模棱两可率高：${row.auditor}`,
        `${row.team}｜模棱两可率 ${formatPct(row.ambiguousRate)}，申报 ${formatInt(row.declarations)} 次`,
        {
          object: row.auditor,
          team: row.team,
          currentRate: row.ambiguousRate,
          declarations: row.declarations,
        },
      ));
    }

    if (row.rejectRate >= thresholds.auditorRejectRate) {
      alerts.push(alert(
        'medium',
        'auditor_high_reject',
        `审核人拒绝率高：${row.auditor}`,
        `${row.team}｜拒绝率 ${formatPct(row.rejectRate)}，申报 ${formatInt(row.declarations)} 次`,
        {
          object: row.auditor,
          team: row.team,
          currentRate: row.rejectRate,
          declarations: row.declarations,
        },
      ));
    }
  }

  const efficiencyDates = [...new Set(efficiencyRows.map((row) => row.date))].sort();
  const efficiencyDate = efficiencyDates.includes(targetDate) ? targetDate : efficiencyDates.at(-1) || '';
  const previousEfficiencyDates = efficiencyDates.filter((date) => date < efficiencyDate).slice(-7);
  const efficiencyBaseline = averageEfficiencyByEmployee(efficiencyRows, previousEfficiencyDates);
  const targetEfficiency = averageEfficiencyByEmployee(efficiencyRows.filter((row) => row.date === efficiencyDate), [efficiencyDate]);
  const currentEfficiencyByTeam = new Map();
  for (const current of targetEfficiency.values()) {
    if (!isConfiguredTeam(current.team)) continue;
    if (!currentEfficiencyByTeam.has(current.team)) currentEfficiencyByTeam.set(current.team, []);
    currentEfficiencyByTeam.get(current.team).push(current);
  }

  for (const [employee, current] of targetEfficiency.entries()) {
    if (!isConfiguredTeam(current.team)) continue;

    const baseline = efficiencyBaseline.get(employee);
    const currentWeighted = current.averageWeightedHandledCount;
    const teamPeers = currentEfficiencyByTeam.get(current.team) || [];
    const peerMedian = median(
      teamPeers
        .filter((item) => item.employee !== employee)
        .map((item) => item.averageWeightedHandledCount),
    );

    const volumeIncreaseRate = baseline?.averageWeightedHandledCount
      ? (currentWeighted - baseline.averageWeightedHandledCount) / baseline.averageWeightedHandledCount
      : 0;
    const peerRatio = peerMedian ? currentWeighted / peerMedian : 0;
    const volumeAboveBaseline = Boolean(
      baseline
        && baseline.days >= 2
        && baseline.averageWeightedHandledCount > 0
        && currentWeighted >= thresholds.employeeVolumeMinWeighted
        && volumeIncreaseRate >= thresholds.employeeVolumeBaselineIncreaseRate,
    );
    const volumeAbovePeers = Boolean(
      peerMedian > 0
        && teamPeers.length >= 3
        && currentWeighted >= thresholds.employeeVolumeMinWeighted
        && peerRatio >= thresholds.employeeVolumePeerRatio,
    );

    if (volumeAboveBaseline || volumeAbovePeers) {
      const details = [];
      if (volumeAboveBaseline) {
        details.push(`较本人近 ${baseline.days} 日均值 ${formatInt(baseline.averageWeightedHandledCount)} 提升 ${formatPct(volumeIncreaseRate)}`);
      }
      if (volumeAbovePeers) {
        details.push(`为同团队当日中位数 ${formatInt(peerMedian)} 的 ${peerRatio.toFixed(1)} 倍`);
      }

      alerts.push(alert(
        volumeAboveBaseline && volumeAbovePeers ? 'high' : 'medium',
        'auditor_high_volume',
        `审核人完成量异常偏高：${employee}`,
        `${current.team}｜当日加权审核量 ${formatInt(currentWeighted)}，${details.join('；')}`,
        {
          object: employee,
          team: current.team,
          currentVolume: currentWeighted,
          baselineVolume: baseline?.averageWeightedHandledCount || 0,
          peerMedian,
          volumeIncreaseRate,
          peerRatio,
          declarations: currentWeighted,
        },
      ));
    }

    if (!baseline || baseline.days < 2 || baseline.averageWeightedHandledCount <= 0) continue;

    const efficiencyMove = (current.averageWeightedHandledCount - baseline.averageWeightedHandledCount) / baseline.averageWeightedHandledCount;
    const precisionMove = current.precisionPassRate - baseline.precisionPassRate;
    if (efficiencyMove <= -thresholds.employeeEfficiencyDropRate && Math.abs(precisionMove) >= thresholds.employeePrecisionMovePp) {
      alerts.push(alert(
        'medium',
        'employee_efficiency_quality_move',
        `人员效率与质量双异动：${employee}`,
        `人均加权审核量较近 ${baseline.days} 日下降 ${formatPct(Math.abs(efficiencyMove))}，精准通过率变化 ${formatPp(precisionMove)}`,
        { object: employee, team: current.team, efficiencyMove, precisionMove },
      ));
    }
  }

  alerts.sort((a, b) => (
    levelRank[b.level] - levelRank[a.level] ||
    Number(b.declarations || 0) - Number(a.declarations || 0) ||
    Math.abs(Number(b.movement || 0)) - Math.abs(Number(a.movement || 0))
  ));

  const dataSource = dashboardPublicUrl || apiBaseUrl || dataDir;
  const message = buildMessage({
    targetDate,
    baselineDates,
    total,
    previousTotal,
    alerts,
    dataSource,
  });

  const result = {
    generatedAt: new Date().toISOString(),
    dataSource,
    targetDate,
    previousDate,
    baselineDates,
    thresholds,
    summary: {
      declarations: total.declarations,
      preciseRate: total.preciseRate,
      proofAccuracyRate: total.proofAccuracyRate,
      ambiguousRate: total.ambiguousRate,
      rejectRate: total.rejectRate,
      previousPreciseRate: previousTotal.preciseRate,
      previousAmbiguousRate: previousTotal.ambiguousRate,
      previousRejectRate: previousTotal.rejectRate,
    },
    alerts,
    message,
  };

  const outputDate = targetDate.replaceAll('-', '');
  const dateOutputDir = path.join(outputDir, outputDate);
  fs.mkdirSync(dateOutputDir, { recursive: true });
  const jsonPath = path.join(dateOutputDir, `daily_alerts_${outputDate}.json`);
  const textPath = path.join(dateOutputDir, `daily_alerts_${outputDate}.txt`);
  fs.writeFileSync(jsonPath, `${JSON.stringify(result, null, 2)}\n`, 'utf8');
  fs.writeFileSync(textPath, `${message}\n`, 'utf8');

  const feishuResult = await pushToFeishu(message);

  console.log(message);
  console.log('');
  console.log(JSON.stringify({
    targetDate,
    alertCount: alerts.length,
    jsonPath,
    textPath,
    feishu: feishuResult,
  }, null, 2));
};

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
