#!/usr/bin/env node
import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';

const projectDir = path.resolve(path.dirname(new URL(import.meta.url).pathname), '..');
const dataDir = path.resolve(process.env.DATA_DIR || path.join(projectDir, 'data'));
const reportsDir = path.resolve(process.env.REPORTS_DIR || path.join(projectDir, 'reports'));
const requestedDate = process.env.REPORT_DATE || process.argv[2] || '';

const qualityPath = path.join(dataDir, 'shared-dataset.json');
const efficiencyPath = path.join(dataDir, 'efficiency-dataset.json');
const propertyDictionaryPath = path.join(dataDir, 'property-category-dictionary.json');
const auditorTeamDictionaryPath = path.join(dataDir, 'auditor-team-dictionary.json');

const readJson = (filePath, fallback) => {
  try {
    return JSON.parse(fs.readFileSync(filePath, 'utf8'));
  } catch {
    return fallback;
  }
};

const qualityDataset = readJson(qualityPath, { rows: [] });
const efficiencyDataset = readJson(efficiencyPath, { rows: [] });
const propertyDictionary = readJson(propertyDictionaryPath, []);
const auditorTeamDictionary = readJson(auditorTeamDictionaryPath, []);

const propertyCategoryByName = new Map(
  (Array.isArray(propertyDictionary) ? propertyDictionary : [])
    .map((entry) => [
      String(entry.propertyName ?? '').trim(),
      String(entry.category ?? '').trim(),
    ])
    .filter(([propertyName, category]) => propertyName && category),
);

const auditorTeamByName = new Map(
  (Array.isArray(auditorTeamDictionary) ? auditorTeamDictionary : [])
    .map((entry) => [
      String(entry.auditorName ?? '').trim(),
      String(entry.team ?? '').trim(),
    ])
    .filter(([auditorName, team]) => auditorName && team),
);

const escapeHtml = (value) =>
  String(value ?? '').replace(/[&<>'"]/g, (char) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    "'": '&#39;',
    '"': '&quot;',
  })[char]);

const formatInt = (value) => Math.round(Number(value || 0)).toLocaleString('zh-CN');
const formatRate = (numerator, denominator, digits = 1) =>
  denominator ? `${((Number(numerator || 0) / Number(denominator || 0)) * 100).toFixed(digits)}%` : '0.0%';
const formatPp = (value, digits = 1) => `${value >= 0 ? '+' : ''}${(value * 100).toFixed(digits)}pp`;
const formatNumber = (value, digits = 1) => Number(value || 0).toFixed(digits);

const metricEmpty = () => ({
  declarations: 0,
  exactPasses: 0,
  ambiguousPasses: 0,
  rejects: 0,
  proofRejects: 0,
});

const normalizeQualityRow = (row) => {
  const declarations = Number(row.declarations || 0);
  let exactPasses = Number(row.exactPasses);
  if (!Number.isFinite(exactPasses)) {
    exactPasses = Math.max(0, declarations - Number(row.ambiguousPasses || 0) - Number(row.rejects || 0));
  }

  const attribute = String(row.attribute ?? row.propertyTag ?? row.tag ?? '').trim();
  const auditor = String(row.auditor ?? row.auditorName ?? '').trim();
  const rawTeam = String(row.auditorTeam ?? row.team ?? '').trim();
  const rawCategory = String(row.category ?? row.attributeCategory ?? '').trim();

  return {
    date: String(row.date ?? '').slice(0, 10),
    auditor,
    auditorTeam: rawTeam || auditorTeamByName.get(auditor) || '未标记',
    session: String(row.session ?? row.saleType ?? '').trim() || '未标记',
    batch: String(row.batch ?? row.batchFlag ?? '').trim() || '未标记',
    category: rawCategory || propertyCategoryByName.get(attribute) || '未标记',
    attribute: attribute || '未标记',
    declarations,
    exactPasses,
    ambiguousPasses: Number(row.ambiguousPasses || 0),
    rejects: Number(row.rejects || 0),
    proofRejects: Number(row.proofRejects || 0),
  };
};

const qualityRows = (qualityDataset.rows || [])
  .map(normalizeQualityRow)
  .filter((row) => row.date && row.declarations > 0);

const availableDates = [...new Set(qualityRows.map((row) => row.date))].sort();
const targetDate = requestedDate || availableDates.at(-1);

if (!targetDate) {
  throw new Error(`没有可用于生成日报的质量数据：${qualityPath}`);
}

const targetRows = qualityRows.filter((row) => row.date === targetDate);
if (!targetRows.length) {
  throw new Error(`指定日期没有质量数据：${targetDate}`);
}

const previousDate = [...availableDates].filter((date) => date < targetDate).at(-1) || '';
const previousRows = previousDate ? qualityRows.filter((row) => row.date === previousDate) : [];

const addMetric = (metric, row) => {
  metric.declarations += row.declarations;
  metric.exactPasses += row.exactPasses;
  metric.ambiguousPasses += row.ambiguousPasses;
  metric.rejects += row.rejects;
  metric.proofRejects += row.proofRejects;
  return metric;
};

const withRates = (metric) => ({
  ...metric,
  preciseRate: metric.declarations ? metric.exactPasses / metric.declarations : 0,
  proofAccuracyRate: metric.declarations
    ? (metric.declarations - metric.ambiguousPasses - metric.proofRejects) / metric.declarations
    : 0,
  ambiguousRate: metric.declarations ? metric.ambiguousPasses / metric.declarations : 0,
  rejectRate: metric.declarations ? metric.rejects / metric.declarations : 0,
});

const summarizeRows = (rows) => withRates(rows.reduce(addMetric, metricEmpty()));

const groupBy = (rows, keySelector) => {
  const groups = new Map();
  for (const row of rows) {
    const key = keySelector(row) || '未标记';
    if (!groups.has(key)) groups.set(key, metricEmpty());
    addMetric(groups.get(key), row);
  }

  return [...groups.entries()].map(([key, metric]) => ({
    key,
    ...withRates(metric),
  }));
};

const total = summarizeRows(targetRows);
const previousTotal = summarizeRows(previousRows);

const sessionRows = groupBy(targetRows, (row) => row.session)
  .map((row) => ({ ...row, share: total.declarations ? row.declarations / total.declarations : 0 }))
  .filter((row) => row.share >= 0.01)
  .sort((a, b) => b.declarations - a.declarations);

const teamRows = groupBy(targetRows, (row) => row.auditorTeam).sort((a, b) => b.declarations - a.declarations);
const categoryRows = groupBy(targetRows, (row) => row.category)
  .map((row) => ({ ...row, share: total.declarations ? row.declarations / total.declarations : 0 }))
  .sort((a, b) => b.declarations - a.declarations);

const attributeRows = groupBy(targetRows, (row) => row.attribute).filter((row) => row.declarations >= 10);
const previousAttributeMap = new Map(
  groupBy(previousRows, (row) => row.attribute).map((row) => [row.key, row]),
);

const highAmbiguousAttributes = [...attributeRows]
  .sort((a, b) => b.ambiguousRate - a.ambiguousRate || b.ambiguousPasses - a.ambiguousPasses)
  .slice(0, 10);
const highRejectAttributes = [...attributeRows]
  .sort((a, b) => b.rejectRate - a.rejectRate || b.rejects - a.rejects)
  .slice(0, 10);
const movingAttributes = attributeRows
  .map((row) => {
    const previous = previousAttributeMap.get(row.key);
    return {
      ...row,
      previous,
      preciseDelta: previous ? row.preciseRate - previous.preciseRate : 0,
    };
  })
  .filter((row) => row.previous && row.previous.declarations >= 10)
  .sort((a, b) => Math.abs(b.preciseDelta) - Math.abs(a.preciseDelta))
  .slice(0, 12);

const efficiencyRows = (efficiencyDataset.rows || [])
  .map((row) => ({
    date: String(row.date ?? '').slice(0, 10),
    employee: String(row.employee ?? row.auditor ?? '').trim() || '未标记',
    team: String(row.team ?? '').trim() || '未标记',
    handledCount: Number(row.handledCount || row.totalAuditCount || 0),
    weightedHandledCount: Number(row.weightedHandledCount || 0),
    firstAuditCount: Number(row.firstAuditCount || 0),
    precisionPassCount: Number(row.precisionPassCount || row.precisePassCount || 0),
    ambiguousCount: Number(row.ambiguousCount || 0),
    proofRefusalCount: Number(row.proofRefusalCount || 0),
  }))
  .filter((row) => row.date);

const efficiencyDates = [...new Set(efficiencyRows.map((row) => row.date))].sort();
const efficiencyDate = efficiencyDates.includes(targetDate) ? targetDate : efficiencyDates.at(-1) || '';
const targetEfficiencyRows = efficiencyDate ? efficiencyRows.filter((row) => row.date === efficiencyDate) : [];

const efficiencyByTeam = (() => {
  const groups = new Map();
  for (const row of targetEfficiencyRows) {
    if (!groups.has(row.team)) {
      groups.set(row.team, {
        key: row.team,
        people: new Set(),
        handledCount: 0,
        weightedHandledCount: 0,
        firstAuditCount: 0,
        precisionPassCount: 0,
      });
    }
    const group = groups.get(row.team);
    group.people.add(row.employee);
    group.handledCount += row.handledCount;
    group.weightedHandledCount += row.weightedHandledCount;
    group.firstAuditCount += row.firstAuditCount;
    group.precisionPassCount += row.precisionPassCount;
  }

  return [...groups.values()]
    .map((group) => ({
      ...group,
      peopleCount: group.people.size,
      averageWeightedCount: group.people.size ? group.weightedHandledCount / group.people.size : 0,
      preciseRate: group.firstAuditCount ? group.precisionPassCount / group.firstAuditCount : 0,
    }))
    .sort((a, b) => b.weightedHandledCount - a.weightedHandledCount);
})();

const metricCard = (label, value, subText) => `
  <div class="metric">
    <div class="metric-label">${escapeHtml(label)}</div>
    <div class="metric-value">${escapeHtml(value)}</div>
    <div class="metric-sub">${escapeHtml(subText)}</div>
  </div>
`;

const table = (headers, rows, renderRow) => `
  <table>
    <thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join('')}</tr></thead>
    <tbody>${rows.length ? rows.map(renderRow).join('') : `<tr><td colspan="${headers.length}" class="muted">暂无数据</td></tr>`}</tbody>
  </table>
`;

const barCell = (share, label) => {
  const width = Math.max(2, Math.min(100, share * 100));
  return `<td><div class="bar"><span style="width:${width}%"></span><b>${escapeHtml(label)}</b></div></td>`;
};

const deltaCell = (value) => `<span class="${value >= 0 ? 'up' : 'down'}">${escapeHtml(formatPp(value))}</span>`;

const summaryItems = [
  `最新可查询日期为 ${targetDate}，当日申报 ${formatInt(total.declarations)} 次，精准通过率 ${formatRate(total.exactPasses, total.declarations)}，较上一有数日 ${previousDate || '无'} ${formatPp(total.preciseRate - previousTotal.preciseRate)}。`,
  `模棱两可率 ${formatRate(total.ambiguousPasses, total.declarations)}，目标为 7%，较上一有数日 ${formatPp(total.ambiguousRate - previousTotal.ambiguousRate)}；拒绝率 ${formatRate(total.rejects, total.declarations)}。`,
  sessionRows[0]
    ? `场次侧，${sessionRows[0].key} 是最大申报来源，占 ${formatRate(sessionRows[0].declarations, total.declarations)}，精准通过率 ${formatRate(sessionRows[0].exactPasses, sessionRows[0].declarations)}。`
    : '场次侧暂无可用分布数据。',
  highAmbiguousAttributes[0]
    ? `属性侧，${highAmbiguousAttributes[0].key} 的模棱两可率最高，为 ${formatRate(highAmbiguousAttributes[0].ambiguousPasses, highAmbiguousAttributes[0].declarations)}，建议优先复核。`
    : '属性侧暂无满足样本量阈值的异常项。',
];

const outputDateKey = targetDate.replaceAll('-', '');
const outputDir = path.join(reportsDir, `daily-${outputDateKey}`);
fs.mkdirSync(outputDir, { recursive: true });
const outputPath = path.join(outputDir, `daily_report_${outputDateKey}.html`);

const html = `<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>预质检日报 - ${escapeHtml(targetDate)}</title>
<style>
:root { --ink:#10192f; --muted:#667797; --line:#e3ebf6; --card:#fff; --blue:#1f6feb; --cyan:#11c5dc; --red:#ef4444; --bg:#f5f9fc; }
* { box-sizing:border-box; }
body { margin:0; font-family:-apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei",sans-serif; color:var(--ink); background:radial-gradient(circle at 18% 0%, #ddfbff 0, transparent 28%), linear-gradient(135deg,#f8fbfd,#eef4f7); }
.report { width:min(1180px, calc(100vw - 44px)); margin:32px auto 56px; }
.hero { padding:34px 38px; border-radius:30px; color:#fff; background:linear-gradient(135deg,#091325 0%,#113b55 52%,#0787a4 100%); box-shadow:0 24px 70px rgba(12,37,64,.22); overflow:hidden; position:relative; }
.hero:after { content:""; position:absolute; right:-120px; top:-150px; width:380px; height:380px; border-radius:50%; background:rgba(255,255,255,.13); }
.kicker { letter-spacing:.36em; color:#63e5ff; font-weight:800; font-size:13px; text-transform:uppercase; }
h1 { margin:12px 0 8px; font-size:38px; line-height:1.14; }
.hero p { margin:0; color:#cfe2ef; font-size:16px; }
.meta { display:flex; gap:10px; flex-wrap:wrap; margin-top:22px; }
.pill { padding:9px 13px; border-radius:999px; background:rgba(255,255,255,.12); color:#eaf7ff; border:1px solid rgba(255,255,255,.18); }
section { margin-top:22px; padding:26px; border:1px solid var(--line); background:rgba(255,255,255,.88); border-radius:24px; box-shadow:0 16px 42px rgba(50,69,99,.08); }
h2 { margin:0 0 14px; font-size:24px; }
h3 { margin:0 0 12px; font-size:18px; }
.summary { display:grid; grid-template-columns:repeat(2,minmax(0,1fr)); gap:14px; }
.summary div { padding:16px 18px; background:#f7fbfe; border-radius:18px; color:#23334e; line-height:1.7; border:1px solid #e7eef8; }
.metrics { display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:13px; margin-top:14px; }
.metric { padding:18px; border-radius:20px; background:linear-gradient(180deg,#fff,#f8fbff); border:1px solid #e0e9f5; }
.metric-label { color:var(--muted); font-size:13px; }
.metric-value { font-size:28px; font-weight:850; margin:8px 0 3px; }
.metric-sub { color:#7f8da6; font-size:12px; }
.grid-2 { display:grid; grid-template-columns:1fr 1fr; gap:18px; }
table { width:100%; border-collapse:separate; border-spacing:0; overflow:hidden; border-radius:16px; font-size:14px; }
th,td { padding:12px 13px; border-bottom:1px solid #edf2f8; text-align:left; vertical-align:middle; }
th { background:#f5f8fc; color:#60718f; font-weight:750; }
td.num, th.num { text-align:right; font-variant-numeric:tabular-nums; }
tr:last-child td { border-bottom:0; }
.bar { height:28px; min-width:110px; border-radius:999px; background:#eef4fb; position:relative; overflow:hidden; }
.bar span { display:block; height:100%; border-radius:999px; background:linear-gradient(90deg,var(--cyan),var(--blue)); opacity:.85; }
.bar b { position:absolute; inset:0; display:flex; align-items:center; justify-content:center; font-size:12px; color:#12213a; }
.note { color:var(--muted); line-height:1.75; margin:8px 0 0; }
.up { color:#087f5b; font-weight:850; }
.down { color:#d6336c; font-weight:850; }
.muted { color:#8a98af; }
.footer { margin-top:18px; color:#7b8aa5; font-size:13px; text-align:center; }
@media (max-width: 900px) { .summary,.grid-2 { grid-template-columns:1fr; } .metrics { grid-template-columns:repeat(2,minmax(0,1fr)); } .report { width:calc(100vw - 24px); margin-top:16px; } .hero { padding:26px; } h1 { font-size:30px; } }
</style>
</head>
<body>
<main class="report">
  <header class="hero">
    <div class="kicker">Daily Quality Review</div>
    <h1>预质检每日质量复盘</h1>
    <p>跟随当前机器的数据源生成，适合放在 Mac mini/ECS 的每日自动任务后执行。</p>
    <div class="meta">
      <span class="pill">分析日期：${escapeHtml(targetDate)}</span>
      <span class="pill">对比基准：${escapeHtml(previousDate || '无上一日期')}</span>
      <span class="pill">质量数据：${formatInt(qualityRows.length)} 行</span>
      <span class="pill">数据源：${escapeHtml(qualityDataset.sourceName || qualityDataset.sourceNames || path.basename(qualityPath))}</span>
    </div>
  </header>

  <section>
    <h2>Executive Summary</h2>
    <div class="summary">${summaryItems.map((item) => `<div>${escapeHtml(item)}</div>`).join('')}</div>
    <div class="metrics">
      ${metricCard('申报次数', formatInt(total.declarations), `较上一日期 ${formatInt(total.declarations - previousTotal.declarations)}`)}
      ${metricCard('精准通过率', formatRate(total.exactPasses, total.declarations), `较上一日期 ${formatPp(total.preciseRate - previousTotal.preciseRate)}`)}
      ${metricCard('举证准确率', formatRate(total.declarations - total.ambiguousPasses - total.proofRejects, total.declarations), `较上一日期 ${formatPp(total.proofAccuracyRate - previousTotal.proofAccuracyRate)}`)}
      ${metricCard('模棱两可率', formatRate(total.ambiguousPasses, total.declarations), `目标 7%，较上一日期 ${formatPp(total.ambiguousRate - previousTotal.ambiguousRate)}`)}
      ${metricCard('拒绝率', formatRate(total.rejects, total.declarations), `较上一日期 ${formatPp(total.rejectRate - previousTotal.rejectRate)}`)}
    </div>
  </section>

  <section>
    <h2>场次结构与质量</h2>
    <p class="note">仅展示申报占比超过 1% 的场次。日报优先关注“大盘占比高 + 精准通过率下滑/模棱两可率高”的场次。</p>
    ${table(['场次', '申报次数', '占比', '精准通过率', '举证准确率', '模棱两可率', '拒绝率'], sessionRows, (row) => `
      <tr>
        <td>${escapeHtml(row.key)}</td>
        <td class="num">${formatInt(row.declarations)}</td>
        ${barCell(row.share, formatRate(row.declarations, total.declarations))}
        <td class="num">${formatRate(row.exactPasses, row.declarations)}</td>
        <td class="num">${formatRate(row.declarations - row.ambiguousPasses - row.proofRejects, row.declarations)}</td>
        <td class="num">${formatRate(row.ambiguousPasses, row.declarations)}</td>
        <td class="num">${formatRate(row.rejects, row.declarations)}</td>
      </tr>
    `)}
  </section>

  <section>
    <h2>属性项风险雷达</h2>
    <div class="grid-2">
      <div>
        <h3>高模棱两可率属性项</h3>
        ${table(['属性项', '申报', '模棱两可率', '精准通过率'], highAmbiguousAttributes, (row) => `
          <tr>
            <td>${escapeHtml(row.key)}</td>
            <td class="num">${formatInt(row.declarations)}</td>
            <td class="num">${formatRate(row.ambiguousPasses, row.declarations)}</td>
            <td class="num">${formatRate(row.exactPasses, row.declarations)}</td>
          </tr>
        `)}
      </div>
      <div>
        <h3>高拒绝率属性项</h3>
        ${table(['属性项', '申报', '拒绝率', '举证准确率'], highRejectAttributes, (row) => `
          <tr>
            <td>${escapeHtml(row.key)}</td>
            <td class="num">${formatInt(row.declarations)}</td>
            <td class="num">${formatRate(row.rejects, row.declarations)}</td>
            <td class="num">${formatRate(row.declarations - row.ambiguousPasses - row.proofRejects, row.declarations)}</td>
          </tr>
        `)}
      </div>
    </div>
  </section>

  <section>
    <h2>精准通过率波动属性项</h2>
    <p class="note">与上一有数日相比，筛选两天申报均不少于 10 次的属性项，按精准通过率变动幅度排序。</p>
    ${table(['属性项', '当日申报', '当日精准通过率', '上一日期精准通过率', '变化'], movingAttributes, (row) => `
      <tr>
        <td>${escapeHtml(row.key)}</td>
        <td class="num">${formatInt(row.declarations)}</td>
        <td class="num">${formatRate(row.exactPasses, row.declarations)}</td>
        <td class="num">${formatRate(row.previous.exactPasses, row.previous.declarations)}</td>
        <td class="num">${deltaCell(row.preciseDelta)}</td>
      </tr>
    `)}
  </section>

  <section>
    <h2>团队质量对比</h2>
    <p class="note">团队优先读取质量底表字段；为空时使用本地审核人团队字典补齐。</p>
    ${table(['团队', '申报次数', '精准通过率', '举证准确率', '模棱两可率', '拒绝率'], teamRows, (row) => `
      <tr>
        <td>${escapeHtml(row.key)}</td>
        <td class="num">${formatInt(row.declarations)}</td>
        <td class="num">${formatRate(row.exactPasses, row.declarations)}</td>
        <td class="num">${formatRate(row.declarations - row.ambiguousPasses - row.proofRejects, row.declarations)}</td>
        <td class="num">${formatRate(row.ambiguousPasses, row.declarations)}</td>
        <td class="num">${formatRate(row.rejects, row.declarations)}</td>
      </tr>
    `)}
  </section>

  <section>
    <h2>属性分类表现</h2>
    <p class="note">分类优先读取底表字段；为空时使用本地属性项分类字典补齐。</p>
    ${table(['属性分类', '申报次数', '占比', '精准通过率', '举证准确率', '模棱两可率'], categoryRows, (row) => `
      <tr>
        <td>${escapeHtml(row.key)}</td>
        <td class="num">${formatInt(row.declarations)}</td>
        ${barCell(row.share, formatRate(row.declarations, total.declarations))}
        <td class="num">${formatRate(row.exactPasses, row.declarations)}</td>
        <td class="num">${formatRate(row.declarations - row.ambiguousPasses - row.proofRejects, row.declarations)}</td>
        <td class="num">${formatRate(row.ambiguousPasses, row.declarations)}</td>
      </tr>
    `)}
  </section>

  <section>
    <h2>人效补充观察</h2>
    <p class="note">人效数据最新匹配日期为 ${escapeHtml(efficiencyDate || '暂无')}；若与质量日期不一致，仅作为补充观察。</p>
    ${table(['团队', '人数', '加权审核量', '人均加权审核量', '精准通过率'], efficiencyByTeam, (row) => `
      <tr>
        <td>${escapeHtml(row.key)}</td>
        <td class="num">${formatInt(row.peopleCount)}</td>
        <td class="num">${formatInt(row.weightedHandledCount)}</td>
        <td class="num">${formatNumber(row.averageWeightedCount)}</td>
        <td class="num">${formatRate(row.precisionPassCount, row.firstAuditCount)}</td>
      </tr>
    `)}
  </section>

  <section>
    <h2>建议动作</h2>
    <p class="note"><strong>每日看数顺序：</strong>先确认最新数据日期，再看场次占比超过 1% 的场次是否出现精准通过率下滑；随后锁定高拒绝率、高模棱两可率、精准通过率波动最大的属性项；最后按团队拆解确认是否集中在某个审核分组。</p>
  </section>

  <div class="footer">Generated at ${escapeHtml(new Date().toLocaleString('zh-CN', { hour12: false }))} · DATA_DIR=${escapeHtml(dataDir)}</div>
</main>
</body>
</html>`;

fs.writeFileSync(outputPath, html, 'utf8');

console.log(JSON.stringify({
  outputPath,
  targetDate,
  previousDate,
  qualityRows: qualityRows.length,
  efficiencyDate,
  totalDeclarations: total.declarations,
}, null, 2));
