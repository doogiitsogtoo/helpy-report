// ard-mpers.js — Lotto (last 4 months from "now") + Osticket1 (prev/curr week) → HTML → PDF → Email
// --------------------------------------------------------------------------------
import "dotenv/config";
import fs from "fs";
import path from "path";
import xlsx from "xlsx";
import puppeteer from "puppeteer";
import nodemailer from "nodemailer";
import cron from "node-cron";
import dayjs from "dayjs";
import utc from "dayjs/plugin/utc.js";
import tz from "dayjs/plugin/timezone.js";

dayjs.extend(utc);
dayjs.extend(tz);

// ────────────────────────────────────────────────────────────────
// CONFIG
// ────────────────────────────────────────────────────────────────
const CONFIG = {
  TIMEZONE: process.env.TIMEZONE || "Asia/Ulaanbaatar",

  PREV_FILE: process.env.CURR_FILE || "./ARD 09.22-09.28.xlsx",
  CURR_FILE: process.env.PREV_FILE || "./ARD 09.29-10.05.xlsx",

  APP_SHEET: process.env.APP_SHEET || "Osticket1", // түүхий мөрүүд (7 хоног)
  MPERS_SHEET: process.env.MPERS_SHEET || "Lotto", // саруудын хүснэгт (social reach)

  OUT_DIR: process.env.OUT_DIR || "./out",
  REPORT_TITLE: process.env.REPORT_TITLE || "МПЕРС ХХК — 7 хоногийн тайлан",
  SUBJECT_PREFIX: process.env.SUBJECT_PREFIX || "[Ard MPERS Weekly]",

  EMAIL_ENABLED: String(process.env.EMAIL_ENABLED ?? "true") === "true",
  SCHED_ENABLED: String(process.env.SCHED_ENABLED ?? "false") === "true",

  CSS_FILE: process.env.CSS_FILE || "./css/template.css",
  COMPANY_FILTER: process.env.COMPANY_FILTER || "Ард МПЕРС",

  JS_LIBS: ["https://cdn.jsdelivr.net/npm/apexcharts"],
};

// ────────────────────────────────────────────────────────────────
// Helpers
// ────────────────────────────────────────────────────────────────
const nnum = (v) => Number(String(v ?? "").replace(/[^\d.-]/g, "")) || 0;
const pad2 = (n) => String(n).padStart(2, "0");
const norm = (s) =>
  String(s || "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();

function parseWeekFromFilename(p) {
  const base = path.basename(p);
  const m = base.match(
    /(\d{1,2})[./-](\d{1,2})\s*[-–]\s*(\d{1,2})[./-](\d{1,2})/
  );
  if (!m) return null;
  return {
    m1: +m[1],
    d1: +m[2],
    m2: +m[3],
    d2: +m[4],
    raw: `${pad2(m[1])}.${pad2(m[2])}-${pad2(m[3])}.${pad2(m[4])}`,
  };
}
const makeYmd = (y, M, D) => `${y}-${pad2(M)}-${pad2(D)}`;

function parseExcelDate(v) {
  if (v == null || v === "") return null;
  const asNum = Number(v);
  if (Number.isFinite(asNum) && asNum > 20000) {
    const ms = (asNum - 25569) * 86400 * 1000;
    const d = dayjs(new Date(ms));
    return d.isValid() ? d : null;
  }
  const d = dayjs(v);
  return d.isValid() ? d : null;
}
function inferYearFromSheet(ws, dateColIndexes = []) {
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  const sample = [];
  for (let r = 1; r < Math.min(rows.length, 120); r++) {
    const row = rows[r] || [];
    for (const c of dateColIndexes) {
      if (c >= 0 && row[c] != null) {
        const d = parseExcelDate(row[c]);
        if (d) sample.push(d.toDate());
      }
    }
  }
  const years = sample.map((d) => dayjs(d).year());
  return years.length ? years[0] : dayjs().year();
}
function getColIdx(headers, patterns) {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "");
    if (patterns.some((re) => re.test(h))) return i;
  }
  return -1;
}
function inRangeInclusive(d, start, end) {
  if (!d || !start || !end) return false;
  return (
    (d.isAfter(start) || d.isSame(start)) && (d.isBefore(end) || d.isSame(end))
  );
}

// ────────────────────────────────────────────────────────────────
// EXTRACTORS (Osticket1) — prev/curr 7 хоног
// ────────────────────────────────────────────────────────────────
function countWeekFromOsticketByCategoryAndChannel(file, sheetName, company) {
  const wb = xlsx.readFile(file, { cellDates: true });
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet not found: ${sheetName} (${file})`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  const hdr = rows[0].map((x) => String(x || ""));

  const idx = {
    company: getColIdx(hdr, [/Компани/i]),
    category: getColIdx(hdr, [/Ангилал/i]),
    subcat: getColIdx(hdr, [/Туслах\s*ангилал/i]),
    created: getColIdx(hdr, [/Үүссэн\s*огноо/i, /Нээсэн\s*огноо/i, /Created/i]),
    closed: getColIdx(hdr, [/Хаагдсан\s*огноо/i, /Closed/i]),
    channel: getColIdx(hdr, [/Шугам/i]), // Phone / Social
  };
  const dateCol = idx.created >= 0 ? idx.created : idx.closed;

  const wk = parseWeekFromFilename(file);
  const year = inferYearFromSheet(ws, [dateCol]);
  const start = wk ? dayjs(makeYmd(year, wk.m1, wk.d1)).startOf("day") : null;
  const end = wk ? dayjs(makeYmd(year, wk.m2, wk.d2)).endOf("day") : null;

  // Тоолно (Үйлчилгээг тоолж хадгална, UI-д хэрэглэхгүй)
  const cats = { Лавлагаа: 0, Гомдол: 0, Үйлчилгээ: 0 };
  const chan = { Phone: 0, Social: 0 };
  const top = { Лавлагаа: new Map(), Гомдол: new Map() };

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r] || [];
    if (company) {
      const comp = String(row[idx.company] || "").trim();
      if (comp !== company) continue;
    }
    const when = dateCol >= 0 ? parseExcelDate(row[dateCol]) : null;
    if (start && end && !inRangeInclusive(when, start, end)) continue;

    const cat = String(row[idx.category] || "").trim();
    const ch = String(row[idx.channel] || "").trim();
    if (cats[cat] != null) cats[cat] += 1;
    if (chan[ch] != null) chan[ch] += 1;

    if (cat === "Лавлагаа" || cat === "Гомдол") {
      const sub = String(row[idx.subcat] || "").trim();
      if (sub) {
        const bag = top[cat];
        bag.set(sub, (bag.get(sub) || 0) + 1);
      }
    }
  }
  return { cats, chan, top, weekLabel: wk?.raw || "7 хоног" };
}

function prevCurrCompare(filePrev, fileCurr, sheetName, company) {
  const prev = countWeekFromOsticketByCategoryAndChannel(
    filePrev,
    sheetName,
    company
  );
  const curr = countWeekFromOsticketByCategoryAndChannel(
    fileCurr,
    sheetName,
    company
  );

  const labels = [
    parseWeekFromFilename(filePrev)?.raw || "Prev",
    parseWeekFromFilename(fileCurr)?.raw || "Curr",
  ];

  const mergeTop = (mapPrev, mapCurr, limit = 10) => {
    const keys = new Set([...mapPrev.keys(), ...mapCurr.keys()]);
    const rows = [...keys].map((k) => {
      const a = mapPrev.get(k) || 0;
      const b = mapCurr.get(k) || 0;
      const base = a > 0 ? a : b || 1;
      return { name: k, prev: a, curr: b, delta: (b - a) / base };
    });
    rows.sort((x, y) => y.curr - x.curr || y.prev - x.prev);
    return rows.slice(0, limit);
  };

  return {
    labels,
    catsPrev: prev.cats,
    catsCurr: curr.cats,
    chanPrev: prev.chan,
    chanCurr: curr.chan,
    topLav: mergeTop(prev.top["Лавлагаа"], curr.top["Лавлагаа"]),
    topGom: mergeTop(prev.top["Гомдол"], curr.top["Гомдол"]),
  };
}

// ────────────────────────────────────────────────────────────────
// Lotto — “Нийт хандалт / Сүүлийн 4 сар” (rolling 4)
// ────────────────────────────────────────────────────────────────
function month4FromLotto(file, sheetName) {
  const wb = xlsx.readFile(file, { cellDates: true });
  const ws = wb.Sheets[sheetName];
  if (!ws) return { labels: [], data: [] };
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  if (!rows.length) return { labels: [], data: [] };

  const digits = (v) => String(v ?? "").replace(/[^\d]/g, "");
  const looksLikeYear = (v) => {
    const d = digits(v);
    const y = Number(d);
    return d.length === 4 && y >= 2019 && y <= 2100;
  };
  const monthFromCell = (v) => {
    const s = String(v ?? "")
      .trim()
      .toLowerCase();
    const m = (s.match(/^(\d{1,2})\s*сар$/) ||
      s.match(/^0?(\d{1,2})$/) ||
      [])[1];
    return m ? +m : null;
  };

  // Толгойн мөрөөс жилүүдийг цуглуулна
  let headerRow = -1;
  const yearCols = new Map(); // year -> col
  for (let r = 0; r < Math.min(rows.length, 80); r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      if (looksLikeYear(row[c])) {
        if (headerRow < 0) headerRow = r;
        yearCols.set(Number(digits(row[c])), c);
      }
    }
    if (yearCols.size && headerRow >= 0) break;
  }
  if (headerRow < 0 || yearCols.size === 0) return { labels: [], data: [] };

  // Сарын нэр байрлах баганыг олно (ихэвчлэн A/B)
  let monthCol = 0;
  outer: for (const cand of [0, 1]) {
    for (let r = headerRow + 1; r < rows.length; r++) {
      if (monthFromCell((rows[r] || [])[cand])) {
        monthCol = cand;
        break outer;
      }
      const s = String((rows[r] || [])[cand] ?? "")
        .toLowerCase()
        .trim();
      if (s === "нийт" || s === "niit") break;
    }
  }

  // Жил*сар → утга
  const mm = new Map(); // year -> Map(month->value)
  for (const [yy] of yearCols.entries()) mm.set(yy, new Map());
  for (let r = headerRow + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const label = row[monthCol];
    const name = String(label ?? "")
      .trim()
      .toLowerCase();
    if (name === "нийт" || name === "niit") break;
    const m = monthFromCell(label);
    if (m == null) continue;
    for (const [yy, col] of yearCols.entries()) {
      const v = nnum(row[col]);
      mm.get(yy).set(m, v);
    }
  }

  // Одоогийн огнооноос 4 сар (M-3..M)
  const now = dayjs().tz(CONFIG.TIMEZONE);
  const seq = [];
  for (let i = 3; i >= 0; i--) {
    const t = now.subtract(i, "month");
    seq.push({ y: t.year(), m: t.month() + 1 });
  }

  const labels = [];
  const data = [];
  for (const { y, m } of seq) {
    labels.push(`${m}сар`);
    data.push(Number(mm.get(y)?.get(m) ?? 0));
  }
  return { labels, data };
}

// ────────────────────────────────────────────────────────────────
// HTML
// ────────────────────────────────────────────────────────────────
const pct100 = (x, d = 0) => `${(x * 100).toFixed(d)}%`;
const num = (n) => Number(n || 0).toLocaleString();

function tableBlock(title, labels, rows) {
  return `
<div class="card">
  <div class="table-wrap">
    <table class="cmp">
      <thead>
        <tr>
          <th width="50%">${title}</th>
          <th>${labels[0]}</th>
          <th>${labels[1]}</th>
          <th>%</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map((r) => {
            const base = r.prev > 0 ? r.prev : r.curr || 1;
            const d = (r.curr - r.prev) / base;
            const up = d >= 0;
            return `<tr>
              <td>${r.name}</td>
              <td class="num">${num(r.prev)}</td>
              <td class="num">${num(r.curr)}</td>
              <td class="num ${up ? "up" : "down"}">${up ? "▲" : "▼"} ${pct100(
              Math.abs(d),
              0
            )}</td>
            </tr>`;
          })
          .join("")}
      </tbody>
    </table>
  </div>
</div>`;
}

function buildHtml(payload, cssText) {
  const libs = CONFIG.JS_LIBS.map((s) => `<script src="${s}"></script>`).join(
    "\n"
  );

  // Нийт бүртгэл = Лавлагаа + Гомдол
  const lavPrev = payload.compare.catsPrev["Лавлагаа"] || 0;
  const lavCurr = payload.compare.catsCurr["Лавлагаа"] || 0;
  const gomPrev = payload.compare.catsPrev["Гомдол"] || 0;
  const gomCurr = payload.compare.catsCurr["Гомдол"] || 0;

  const currTotal = lavCurr + gomCurr;
  const prevTotal = lavPrev + gomPrev;
  const totDelta = prevTotal ? (currTotal - prevTotal) / prevTotal : 0;

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<style>
${cssText}
/* ==== Gold Theme (charts untouched) ==== */
:root{
  --gold-50:#fffbeb;
  --gold-100:#fef3c7;
  --gold-150:#fbe9b6;
  --gold-200:#f1d79a;
  --gold-300:#e2c074;
  --gold-400:#d4a941;
  --gold-500:#c69026;
  --gold-600:#a87a1a;
  --gray-200:#e5e7eb; --gray-500:#6b7280; --gray-700:#374151;
}

body{font-family:Arial,Helvetica,sans-serif}
.sheet{margin-top:16px}
.grid{display:grid;gap:16px}
.grid-2{grid-template-columns:1.2fr .8fr}
.grid-1-1{grid-template-columns:1fr 1fr}
.grid-1-1-1{grid-template-columns:1fr 1fr 1fr}

.card{
  background:#fff;border:1px solid var(--gray-200);
  border-radius:12px;padding:12px;break-inside:avoid;
  box-shadow:0 4px 12px rgba(0,0,0,.04);
}

/* Том гарчиг/тендер */
.brand{
  background:linear-gradient(135deg,#d8b76a 0%, #f0d9a6 50%, #d1b076 100%);
  color:#fff;padding:14px 16px;border:none;border-radius:12px;
  font-weight:700; letter-spacing:.2px;
}

/* Жижиг хэсгийн толгой—илүү зөөлөн алтлаг */
.section-head{
  background:linear-gradient(135deg,#f8e7b8 0%, #f2d89a 100%);
  color:#7a5b21; padding:10px 12px; border-radius:10px; font-weight:700;
}

/* Хүснэгтүүд */
table.cmp{width:100%;border-collapse:collapse;font-size:12.5px}
table.cmp th, table.cmp td{border:1px solid var(--gold-200);padding:6px 8px;vertical-align:top}
table.cmp thead th{
  background:linear-gradient(180deg,#fff7e1 0%, #fde9b8 100%);
  color:#7a5b21
}
td.num{text-align:right;white-space:nowrap}
.up{color:#16a34a}.down{color:#b91c1c}

.footer{margin-top:24px;color:#6b7280;font-size:11px}
.bullet{margin:6px 0 0 18px;padding:0}
.bullet li{margin:6px 0}
</style>
${libs}
<title>Ard MPERS Report</title>
</head>
<body>

<!-- Том баннер маягийн толгой -->
<div class="brand" style="height:auto">
  МПЕРС ХХК
</div>

<section class="sheet">
  <div class="grid grid-2">
    <div class="card">
      <div class="section-head">НИЙТ ХАНДАЛТ /Сүүлийн 4 сараар/</div>
      <div id="mpers-handalt" style="height:230px;margin-top:8px"></div>
    </div>
    <div class="card">
      <div class="section-head">НИЙТ БҮРТГЭЛ /Сүүлийн 2 долоо хоногоор/</div>
      <div id="mpers-week" style="height:230px;margin-top:8px"></div>
      <ul class="bullet">
        <li>Тайлант 7 хоногт <b>${num(
          currTotal
        )}</b> бүртгэл хийгдсэн. Өмнөх 7 хоногоос
          <span class="${totDelta >= 0 ? "up" : "down"}">${(
    totDelta * 100
  ).toFixed(0)}%</span>
          ${totDelta >= 0 ? "өссөн" : "буурсан"}.</li>
      </ul>
    </div>
  </div>
</section>

<section class="sheet">
  <div class="grid grid-1-1-1">
    <div class="card">
      <h3 style="margin:0;color:#7a5b21">СУВГААР (Phone/Social)</h3>
      <div id="mpers-chan" style="height:150px"></div>
    </div>
    <div class="card">
      <h3 style="margin:0;color:#7a5b21">ЛАВЛАГАА</h3>
      <div id="mpers-mini-lav" style="height:150px"></div>
    </div>
    <div class="card">
      <h3 style="margin:0;color:#7a5b21">ГОМДОЛ</h3>
      <div id="mpers-mini-gom" style="height:150px"></div>
    </div>
  </div>
</section>

<section class="sheet">
  <div class="grid grid-1-1">
    ${tableBlock(
      "ТОП Лавлагаа",
      payload.compare.labels,
      payload.compare.topLav
    )}
    ${tableBlock("ТОП Гомдол", payload.compare.labels, payload.compare.topGom)}
  </div>
</section>

<div class="footer">Автоматаар бэлтгэсэн тайлан (MPERS / Lotto)</div>

<script>
(function(){
  // months (Lotto – last 4 months from now) — CHART COLORS NOT OVERRIDDEN
  new ApexCharts(document.querySelector("#mpers-handalt"), {
    series: [{ name:"Нийт", data: ${JSON.stringify(payload.month4.data)} }],
    chart: { height: 220, type: "line", toolbar: { show:false } },
    dataLabels: { enabled: true },
    stroke: { curve: "straight", width: 3 },
    markers: { size: 4 },
    grid: { row:{ colors:["#f7f7f7","transparent"], opacity: .5 } },
    xaxis: { categories: ${JSON.stringify(payload.month4.labels)} },
    yaxis: { min: 0 }
  }).render();

  // weeks (prev vs curr) — Лавлагаа + Гомдол (NO custom colors)
  new ApexCharts(document.querySelector("#mpers-week"), {
    chart: { type:"bar", height:220, stacked:false, toolbar:{show:false} },
    plotOptions: { bar:{ horizontal:false, columnWidth:"55%", endingShape:"rounded" } },
    dataLabels: {
      enabled:true,
      style:{colors:["#fff"]},
      background:{ enabled:true, foreColor:"#000", padding:4, borderRadius:4, borderWidth:1, borderColor:"#c8a968", opacity:.9 }
    },
    stroke: { show:true, width:2, colors:["transparent"] },
    series: [
      { name:"Гомдол",   data: [${gomPrev}, ${gomCurr}] },
      { name:"Лавлагаа", data: [${lavPrev}, ${lavCurr}] }
    ],
    xaxis: { categories: ${JSON.stringify(payload.compare.labels)} },
    fill: { opacity:1 },
    legend: { position:"bottom" }
  }).render();

  // channel chart (Phone/Social) — prev vs curr (NO custom colors)
  new ApexCharts(document.querySelector("#mpers-chan"), {
    chart: { type:"bar", height:150, stacked:false, toolbar:{show:false} },
    plotOptions: { bar:{ horizontal:false, columnWidth:"55%", borderRadius:6 } },
    dataLabels: { enabled:true },
    series: [
      { name:${JSON.stringify(payload.compare.labels[0])},
        data: ${JSON.stringify([
          payload.compare.chanPrev.Phone || 0,
          payload.compare.chanPrev.Social || 0,
        ])} },
      { name:${JSON.stringify(payload.compare.labels[1])},
        data: ${JSON.stringify([
          payload.compare.chanCurr.Phone || 0,
          payload.compare.chanCurr.Social || 0,
        ])} }
    ],
    xaxis: { categories: ${JSON.stringify(["Phone", "Social"])} },
    legend: { position:"bottom" }
  }).render();

  // mini bars (Lavlagaa/Gomdol) (NO custom colors)
  const mkMini = (el, vals) => new ApexCharts(document.querySelector(el), {
    chart:{ type:"bar", height:150, toolbar:{show:false} },
    plotOptions:{ bar:{ horizontal:false, columnWidth:"60%", borderRadius:8 } },
    dataLabels:{ enabled:true },
    series:[{ name:"Prev/Curr", data: vals }],
    xaxis:{ categories: ${JSON.stringify(payload.compare.labels)} },
    yaxis:{ min:0 }
  }).render();
  mkMini("#mpers-mini-lav", [${lavPrev}, ${lavCurr}]);
  mkMini("#mpers-mini-gom", [${gomPrev}, ${gomCurr}]);
})();
</script>

</body>
</html>`;
}

// ────────────────────────────────────────────────────────────────
// PDF + EMAIL
// ────────────────────────────────────────────────────────────────
async function htmlToPdf(html, outPath) {
  const browser = await puppeteer.launch({
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
    defaultViewport: { width: 1600, height: 1000, deviceScaleFactor: 2 },
  });
  try {
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });
    await page.emulateMediaType("screen");
    // ApexCharts render хүлээе
    await page
      .waitForSelector(".apexcharts-svg", { timeout: 15000 })
      .catch(() => {});
    await new Promise((res) => setTimeout(res, 600));
    await page.pdf({
      path: outPath,
      format: "A4",
      landscape: true,
      printBackground: true,
      preferCSSPageSize: true,
      margin: { top: "12mm", right: "12mm", bottom: "12mm", left: "12mm" },
    });
  } finally {
    await browser.close();
  }
}

async function sendEmailWithPdf(pdfPath, subject) {
  if (!CONFIG.EMAIL_ENABLED) {
    console.log("[EMAIL] Disabled → skipping send.");
    return;
  }
  const port = Number(process.env.SMTP_PORT || 587);
  const secure =
    port === 465 ? true : String(process.env.SMTP_SECURE || "false") === "true";

  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port,
    secure,
    auth:
      process.env.SMTP_USER && process.env.SMTP_PASS
        ? { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS }
        : undefined,
    pool: true,
    maxConnections: 1,
    connectionTimeout: 20000,
    greetingTimeout: 15000,
    socketTimeout: 30000,
    requireTLS: port === 587,
    tls: { minVersion: "TLSv1.2" },
  });

  await transporter.verify();
  await transporter.sendMail({
    from: process.env.FROM_EMAIL,
    to: process.env.RECIPIENTS,
    subject,
    html: `<p>Сайн байна уу,</p><p>МПЕРС (Lotto) 7 хоногийн тайланг хавсаргав.</p><p style="color:#666;font-size:12px">Автоматаар илгээв.</p>`,
    attachments: [{ filename: path.basename(pdfPath), path: pdfPath }],
  });
}

// ────────────────────────────────────────────────────────────────
// MAIN
// ────────────────────────────────────────────────────────────────
async function runOnce() {
  if (!fs.existsSync(CONFIG.CURR_FILE))
    throw new Error(`Missing file: ${CONFIG.CURR_FILE}`);
  if (!fs.existsSync(CONFIG.PREV_FILE))
    throw new Error(`Missing file: ${CONFIG.PREV_FILE}`);
  if (!fs.existsSync(CONFIG.CSS_FILE))
    throw new Error(`Missing CSS: ${CONFIG.CSS_FILE}`);
  if (!fs.existsSync(CONFIG.OUT_DIR))
    fs.mkdirSync(CONFIG.OUT_DIR, { recursive: true });

  const compare = prevCurrCompare(
    CONFIG.PREV_FILE,
    CONFIG.CURR_FILE,
    CONFIG.APP_SHEET,
    CONFIG.COMPANY_FILTER
  );

  // Lotto — last 4 calendar months from "now"
  let month4 = month4FromLotto(CONFIG.CURR_FILE, CONFIG.MPERS_SHEET);
  if (!month4) month4 = { labels: [], data: [] };

  const cssText = fs.readFileSync(CONFIG.CSS_FILE, "utf-8");
  const html = buildHtml({ compare, month4 }, cssText);

  const monday = dayjs().tz(CONFIG.TIMEZONE).startOf("week").add(1, "day");
  const pdfName = `ard-mpers-weekly-${monday.format("YYYYMMDD")}.pdf`;
  const pdfPath = path.join(CONFIG.OUT_DIR, pdfName);

  await htmlToPdf(html, pdfPath);

  const subject = `${CONFIG.SUBJECT_PREFIX} ${
    CONFIG.REPORT_TITLE
  } — ${monday.format("YYYY-MM-DD")}`;
  await sendEmailWithPdf(pdfPath, subject);

  console.log(
    `[OK] Sent ${pdfName} → ${process.env.RECIPIENTS || "(no recipients set)"}`
  );
}

// Scheduler (optional)
function startScheduler() {
  if (!CONFIG.SCHED_ENABLED) {
    console.log("Scheduler disabled.");
    return;
  }
  cron.schedule(
    "0 9 * * 1",
    async () => {
      try {
        await runOnce();
      } catch (e) {
        console.error(e);
      }
    },
    { timezone: CONFIG.TIMEZONE }
  );
  console.log(`Scheduler started → Every Monday 09:00 (${CONFIG.TIMEZONE})`);
}

// Entry
const runNow = process.argv.includes("--once");
if (runNow) {
  runOnce().catch((e) => {
    console.error(e);
    process.exit(1);
  });
} else {
  startScheduler();
}
