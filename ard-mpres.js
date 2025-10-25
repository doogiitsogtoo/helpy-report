// ard-mpers.js — Lotto (last 4 months from "now") + Osticket1 (prev/curr week) → HTML → PDF → Email
// ────────────────────────────────────────────────────────────────
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

  PREV_FILE: process.env.PREV_FILE || "./ARD 09.22-09.28.xlsx",
  CURR_FILE: process.env.CURR_FILE || "./ARD 09.29-10.05.xlsx",
  GOMDOL_FILE: "./gomdol-weekly.xlsx",

  // Sheets
  APP_SHEET: process.env.APP_SHEET || "Osticket1", // raw rows (weekly)
  MPERS_SHEET: process.env.MPERS_SHEET || "Lotto", // months table (social reach)
  COMPANY_FILTER: process.env.COMPANY_FILTER || "Ард МПЕРС",

  // PDF / Email
  OUT_DIR: process.env.OUT_DIR || "./out",
  CSS_FILE: process.env.CSS_FILE || "./css/template.css",
  REPORT_TITLE: process.env.REPORT_TITLE || "МПЕРС ХХК — 7 хоногийн тайлан",
  SUBJECT_PREFIX: process.env.SUBJECT_PREFIX || "[Ard MPERS Weekly]",
  EMAIL_ENABLED: String(process.env.EMAIL_ENABLED ?? "true") === "true",
  SCHED_ENABLED: String(process.env.SCHED_ENABLED ?? "false") === "true",
};

// ────────────────────────────────────────────────────────────────
// Helpers
// ────────────────────────────────────────────────────────────────
const esc = (s) =>
  String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
const nnum = (v) => Number(String(v ?? "").replace(/[^\d.-]/g, "")) || 0;
const pad2 = (n) => String(n).padStart(2, "0");

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
  for (let r = 1; r < Math.min(rows.length, 200); r++) {
    const row = rows[r] || [];
    for (const c of dateColIndexes) {
      if (c >= 0 && row[c] != null) {
        const d = parseExcelDate(row[c]);
        if (d) return d.year();
      }
    }
  }
  return dayjs().year();
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
/** Lotto — “Нийт хандалт / Сүүлийн 4 сар” (rolling 4 from now) */
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

  // Толгойн мөр дэх жилүүд
  let headerRow = -1;
  const yearCols = new Map();
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

  // Сарын шошго байрлах багана
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

  // Одоогоос 4 сар (M-3..M)
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
// HTML (Chart.js, plugin-гүй — тоог canvas дээр бичнэ)
// ────────────────────────────────────────────────────────────────
const num = (n) => Number(n || 0).toLocaleString();
const pct = (x) => `${(x * 100).toFixed(0)}%`;

function renderCover({ company, periodText }) {
  return `
  <section class="hero" style="margin-bottom:16px">
    <div style="background:linear-gradient(135deg,#d8b76a,#f0d9a6);
                border-radius:12px;padding:24px;display:flex;justify-content:space-between;align-items:center;min-height:200px;">
      <div style="background:#fff;border-radius:16px;padding:18px 22px;display:inline-block">
        <div style="font-weight:800;font-size:26px;letter-spacing:.5px;color:#b38b3c">ARD</div>
        <div style="color:#666;margin-top:4px">Хүчтэй. Хамтдаа.</div>
      </div>
      <div style="color:#fff;text-align:right;padding:8px 16px">
        <div style="font-size:32px;font-weight:800;line-height:1.1">${esc(
          company
        )}</div>
        <div style="opacity:.9;margin-top:8px">${esc(periodText || "")}</div>
      </div>
    </div>
  </section>`;
}

function tableBlock(title, labels, rows) {
  return `
  <div class="card">
    <div class="card-title">${esc(title)}</div>
    <table class="cmp">
      <thead><tr><th>Туслах ангилал</th><th>${labels[0]}</th><th>${
    labels[1]
  }</th><th>%</th></tr></thead>
      <tbody>
        ${rows
          .map((r) => {
            const base = r.prev > 0 ? r.prev : r.curr || 1;
            const d = (r.curr - r.prev) / base;
            const up = d >= 0;
            return `<tr>
            <td>${esc(r.name)}</td>
            <td class="num">${num(r.prev)}</td>
            <td class="num">${num(r.curr)}</td>
            <td class="num ${up ? "up" : "down"}">${up ? "▲" : "▼"} ${pct(
              Math.abs(d)
            )}</td>
          </tr>`;
          })
          .join("")}
      </tbody>
    </table>
  </div>`;
}

function renderLayout({ month4, compare }) {
  const monthsLabels = month4.labels || [];
  const monthsData = month4.data || [];

  const lavPrev = compare.catsPrev["Лавлагаа"] || 0;
  const lavCurr = compare.catsCurr["Лавлагаа"] || 0;
  const gomPrev = compare.catsPrev["Гомдол"] || 0;
  const gomCurr = compare.catsCurr["Гомдол"] || 0;

  const prevTotal = lavPrev + gomPrev;
  const currTotal = lavCurr + gomCurr;
  const delta = prevTotal ? (currTotal - prevTotal) / prevTotal : 0;

  // Canvas дээр тоо бичих жижиг туслах функц (plugin биш)
  const drawValuesFn = `
    function drawValues(chart){
      const {ctx} = chart;
      ctx.save();
      ctx.font = '12px system-ui,-apple-system,Segoe UI,Roboto,Arial';
      ctx.fillStyle = '#111';
      ctx.textAlign = 'center';
      chart.data.datasets.forEach((ds, di) => {
        const meta = chart.getDatasetMeta(di);
        (ds.data || []).forEach((v, i) => {
          if (v == null) return;
          const el = meta.data[i];
          if (!el) return;
          const pos = el.tooltipPosition ? el.tooltipPosition() : {x: el.x, y: el.y};
          ctx.fillText(String(v), pos.x, pos.y - 6);
        });
      });
      ctx.restore();
    }`;

  const lineCard = `
  <div class="card" style="height: 500px; margin-bottom: 4rem;">
    <div class="card-title">НИЙТ ХАНДАЛТ /Сүүлийн ${
      monthsLabels.length
    } сараар/</div>
    <canvas id="lottoLine"></canvas>
  </div>
  <script>(function(){
    ${drawValuesFn}
    const ctx = document.getElementById('lottoLine').getContext('2d');
    const ch  = new Chart(ctx,{
      type:'line',
      data:{ labels:${JSON.stringify(monthsLabels)},
        datasets:[{ label:'', data:${JSON.stringify(
          monthsData
        )}, tension:.3, pointRadius:4 }]},
      options:{ animation:false, plugins:{legend:{display:false}}, scales:{ y:{ beginAtZero:true } } }
    });
    setTimeout(()=>drawValues(ch),0);
  })();</script>`;

  const weeklyCard = `
  <div class="card">
    <div class="card-title">НИЙТ БҮРТГЭЛ /Сүүлийн 2 долоо хоног/</div>
    <div class="grid">
      <canvas id="weekBar"></canvas>
      <ul style="margin:10px 0 0 18px;line-height:1.6">
        <li>Тайлант 7 хоногт нийт <b>${num(currTotal)}</b>.</li>
        <li>Өмнөх 7 хоногоос <b>${
          delta >= 0 ? "өссөн" : "буурсан"
        }</b>: <b>${pct(Math.abs(delta))}</b>.</li>
      </ul>
    </div>
  </div>
  <script>(function(){
    ${drawValuesFn}
    const ctx = document.getElementById('weekBar').getContext('2d');
    const ch  = new Chart(ctx,{
      type:'bar',
      data:{
        labels:${JSON.stringify(compare.labels)},
        datasets:[
          {label:'Лавлагаа', data:[${lavPrev}, ${lavCurr}]},
          {label:'Гомдол',   data:[${gomPrev}, ${gomCurr}]}
        ]},
      options:{ animation:false, plugins:{legend:{position:'bottom'}}, scales:{ y:{ beginAtZero:true } } }
    });
    setTimeout(()=>drawValues(ch),0);
  })();</script>`;

  const chanCard = `
  <div class="card">
    <div class="card-title">СУВГААР (Phone / Social)</div>
    <canvas id="chanBar"></canvas>
  </div>
  <script>(function(){
    ${drawValuesFn}
    const ctx = document.getElementById('chanBar').getContext('2d');
    const ch  = new Chart(ctx,{
      type:'bar',
      data:{
        labels:${JSON.stringify(["Phone", "Social"])},
        datasets:[
          {label:${JSON.stringify(compare.labels[0])}, data:${JSON.stringify([
    compare.chanPrev.Phone || 0,
    compare.chanPrev.Social || 0,
  ])}},
          {label:${JSON.stringify(compare.labels[1])}, data:${JSON.stringify([
    compare.chanCurr.Phone || 0,
    compare.chanCurr.Social || 0,
  ])}}
        ]},
      options:{ animation:false, plugins:{legend:{position:'bottom'}}, scales:{ y:{ beginAtZero:true } } }
    });
    setTimeout(()=>drawValues(ch),0);
  })();</script>`;

  const topLavTable = tableBlock(
    "ТОП Лавлагаа",
    compare.labels,
    compare.topLav
  );
  const topGomTable = tableBlock("ТОП Гомдол", compare.labels, compare.topGom);

  return `
  <section>
    <div class="grid">
      ${lineCard}
      ${weeklyCard}
    </div>
    <div class="grid grid-1-1" style="margin-top:8px">
      ${chanCard}
      <div></div>
    </div>
    <div class="grid grid-1" style="margin-top:8px">
      ${topLavTable}
      ${topGomTable}
    </div>
  </section>`;
}

function wrapHtml(bodyHtml) {
  const css = fs.readFileSync(CONFIG.CSS_FILE, "utf-8");
  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
  <style>${css}
    .card{background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:12px;margin-bottom:12px;page-break-inside:avoid}
    .card-title{font-weight:700;margin-bottom:6px}
    table.cmp{table-layout:fixed;width:100%;border-collapse:collapse;font-size:12.5px}
    table.cmp th,table.cmp td{border:1px solid #e5e7eb;padding:6px 8px;vertical-align:top}
    table.cmp thead th{background:#fff8e1}
    td.num{text-align:right;white-space:nowrap}
    .up{color:#16a34a}.down{color:#ef4444}
  </style>
</head>
<body>
 <div class="container py-3">
    <div class="row g-3">
      ${bodyHtml}
      <div class="footer">Автоматаар бэлтгэсэн тайлан (Node.js)</div>
    </div>
  </div>
</body>
</html>`;
}

// ────────────────────────────────────────────────────────────────
// PDF + EMAIL
// ────────────────────────────────────────────────────────────────
async function htmlToPdf(html, outPath) {
  const browser = await puppeteer.launch({
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });
  try {
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });
    // Chart.js болон дор хаяж 2 canvas бий эсэхийг шалгана
    await page.waitForFunction(
      () => window.Chart && document.querySelectorAll("canvas").length >= 2,
      { timeout: 8000 }
    );
    await page.emulateMediaType("screen");
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
    html: `<p>Сайн байна уу,</p><p>МПЕРС 7 хоногийн тайланг хавсаргав.</p><p style="color:#666;font-size:12px">Автоматаар илгээв.</p>`,
    attachments: [
      {
        filename: path.basename(pdfPath),
        path: pdfPath,
        contentType: "application/pdf",
      },
    ],
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
  let month4 = month4FromLotto(CONFIG.CURR_FILE, CONFIG.MPERS_SHEET) || {
    labels: [],
    data: [],
  };

  const cover = renderCover({
    company: "МПЕРС ХХК",
    periodText: `${compare.labels[0]} – ${compare.labels[1]}`,
  });
  const body = cover + renderLayout({ month4, compare });
  const html = wrapHtml(body);

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
if (process.argv.includes("--once")) {
  runOnce().catch((e) => {
    console.error(e);
    process.exit(1);
  });
} else {
  startScheduler();
}
