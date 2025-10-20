// ard-leasing.js — ALS (months+weeks) + Osticket1 (prev/curr) → HTML → PDF → Email
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

  APP_SHEET: process.env.APP_SHEET || "Osticket1",
  ALS_SHEET: process.env.ALS_SHEET || "ALS",

  OUT_DIR: process.env.OUT_DIR || "./out",
  REPORT_TITLE: process.env.REPORT_TITLE || "Ард Лизинг — 7 хоногийн тайлан",
  SUBJECT_PREFIX: process.env.SUBJECT_PREFIX || "[ArdLeasing Weekly]",

  EMAIL_ENABLED: String(process.env.EMAIL_ENABLED ?? "true") === "true",
  SCHED_ENABLED: String(process.env.SCHED_ENABLED ?? "false") === "true",

  CSS_FILE: process.env.CSS_FILE || "./css/template.css",
  COMPANY_FILTER: process.env.COMPANY_FILTER || "Ард Лизинг",

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
function makeYmd(y, M, D) {
  return `${y}-${pad2(M)}-${pad2(D)}`;
}

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

function inferYearFromDates(series, fallbackYear = dayjs().year()) {
  const years = series.filter(Boolean).map((d) => dayjs(d).year());
  return years.length ? years[0] : fallbackYear;
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
  return inferYearFromDates(sample, dayjs().year());
}
function getColIdx(headers, patterns) {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "");
    if (patterns.some((re) => re.test(h))) return i;
  }
  return -1;
}
function findRowByNameAnywhere(rows, name) {
  const want = norm(name);
  for (const r of rows) {
    if (!r) continue;
    for (let c = 0; c < r.length; c++) if (norm(r[c]) === want) return r;
  }
  return null;
}
function parseWeekLabelCell(s) {
  const m = String(s || "").match(
    /(\d{1,2})[./-](\d{1,2}).*?[-–].*?(\d{1,2})[./-](\d{1,2})/
  );
  if (!m) return null;
  return { m1: +m[1], d1: +m[2], m2: +m[3], d2: +m[4], raw: m[0] };
}
function findWeekColumnsFuzzy(rows) {
  const cols = [];
  for (let c = 0; c < (rows[0] || []).length; c++) {
    for (let r = 0; r < rows.length; r++) {
      const p = parseWeekLabelCell((rows[r] || [])[c]);
      if (p) {
        cols.push({ col: c, label: p.raw, parsed: p });
        break;
      }
    }
  }
  cols.sort((a, b) => a.col - b.col);
  return cols;
}
function inRangeInclusive(d, start, end) {
  if (!d || !start || !end) return false;
  return (
    (d.isAfter(start) || d.isSame(start)) && (d.isBefore(end) || d.isSame(end))
  );
}

// ────────────────────────────────────────────────────────────────
// EXTRACTORS
// ────────────────────────────────────────────────────────────────

// Osticket1 → ангиллын тоонууд (файлын нэрийн долоо хоногоор шүүнэ)
function countByCategoryWithinFile(file, sheetName, companyFilter) {
  const wb = xlsx.readFile(file, { cellDates: true });
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet not found: ${sheetName} (${file})`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  if (!rows.length) return { Лавлагаа: 0, Үйлчилгээ: 0, Гомдол: 0 };

  const hdr = rows[0].map((x) => String(x || ""));
  const idx = {
    company: getColIdx(hdr, [/Компани/i]),
    category: getColIdx(hdr, [/Ангилал/i]),
    created: getColIdx(hdr, [/Үүссэн\s*огноо/i, /Нээсэн\s*огноо/i, /Created/i]),
    closed: getColIdx(hdr, [/Хаагдсан\s*огноо/i, /Closed/i]),
  };
  const dateCol = idx.created >= 0 ? idx.created : idx.closed;

  const wk = parseWeekFromFilename(file);
  const year = inferYearFromSheet(ws, [dateCol]);
  const start = wk ? dayjs(makeYmd(year, wk.m1, wk.d1)).startOf("day") : null;
  const end = wk ? dayjs(makeYmd(year, wk.m2, wk.d2)).endOf("day") : null;

  const acc = { Лавлагаа: 0, Үйлчилгээ: 0, Гомдол: 0 };
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r] || [];
    if (companyFilter) {
      const comp = String(row[idx.company] || "").trim();
      if (comp !== companyFilter) continue;
    }
    const when = dateCol >= 0 ? parseExcelDate(row[dateCol]) : null;
    if (start && end && !inRangeInclusive(when, start, end)) continue;

    const cat = String(row[idx.category] || "").trim();
    if (acc[cat] != null) acc[cat] += 1;
  }
  return acc;
}

// ALS → "Нийт хандалт /Сүүлийн 4 сараар/" (Outbound блок; fallback дээд хүснэгт)
// ALS → Нийт хандалт /Сүүлийн 4 сараар/
function month4FromALS(file, sheetName) {
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
      .toLowerCase()
      .replace(/\s+/g, "");
    // "1 сар" / "1сар" / "01"
    const m = (s.match(/^(\d{1,2})сар$/) || s.match(/^0?(\d{1,2})$/) || [])[1];
    const n = m ? Number(m) : NaN;
    return n >= 1 && n <= 12 ? n : null;
  };

  // ── 1) Outbound блок дотроос унших
  let obRow = -1;
  for (let r = 0; r < rows.length; r++) {
    const a0 = String((rows[r] || [])[0] ?? "")
      .trim()
      .toLowerCase();
    if (a0 === "outbound") {
      obRow = r;
      break;
    }
  }

  const buildFromMatrix = (startRow, headerRow) => {
    const header = rows[headerRow] || [];
    // одоогийн жил байвал тэрийг, үгүй бол хамгийн баруун "жил"-ийг авна
    let yearCol = -1;
    const prefer = String(dayjs().year());
    for (let c = 0; c < header.length; c++) {
      if (looksLikeYear(header[c]) && digits(header[c]) === prefer) {
        yearCol = c;
        break;
      }
    }
    if (yearCol < 0) {
      for (let c = header.length - 1; c >= 0; c--) {
        if (looksLikeYear(header[c])) {
          yearCol = c;
          break;
        }
      }
    }
    if (yearCol < 0) return null;

    const mm = [];
    for (let r = headerRow + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const a0 = String(row[0] ?? "")
        .trim()
        .toLowerCase();
      if (/^нийт$|^niit$/.test(a0)) break; // "Нийт" дээр зогсоно
      const m = monthFromCell(row[0]) ?? monthFromCell(row[1]); // зарим файлд сар 2-р баганад байдаг
      if (!m) continue;
      mm[m] = nnum(row[yearCol]);
    }
    const present = [];
    for (let m = 1; m <= 12; m++) if (mm[m] != null) present.push([m, mm[m]]);
    if (!present.length) return null;

    const last4 = present.slice(-4);
    return {
      labels: last4.map(([m]) => `${m}сар`),
      data: last4.map(([, v]) => Number(v) || 0),
    };
  };

  // Outbound байгаа бол тэндээс
  if (obRow >= 0) {
    // Outbound-ын дараах мөрүүдээс "жил" агуулагдсан толгой мөрийг олно
    let yyRow = obRow + 1;
    while (yyRow < rows.length && !(rows[yyRow] || []).some(looksLikeYear))
      yyRow++;
    if (yyRow < rows.length) {
      const got = buildFromMatrix(obRow + 1, yyRow);
      if (got) return got;
    }
  }

  // ── 2) Fallback — дээд үндсэн хүснэгтээс
  // (жилийн толгой ба сарууд 1-р баганад байрлах стандарт хуваарийн дагуу)
  let headerRow = -1,
    yearCol = -1;
  for (let r = 0; r < Math.min(rows.length, 30); r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      if (looksLikeYear(row[c])) {
        headerRow = r;
        yearCol = c;
        break;
      }
    }
    if (headerRow >= 0) break;
  }
  if (headerRow >= 0) {
    const got = buildFromMatrix(headerRow + 1, headerRow);
    if (got) return got;
  }

  // юу ч олдохгүй бол хоосон буцаана
  return { labels: [], data: [] };
}

// ALS → 4 долоо хоног (Лавлагаа/Гомдол)
function weeks4FromALS(file, sheetName) {
  const wb = xlsx.readFile(file, { cellDates: true });
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  if (!rows.length) return null;

  const weekCols = findWeekColumnsFuzzy(rows);
  if (!weekCols.length) return null;
  const last4 = weekCols.slice(-4);

  const wantRows = {
    Лавлагаа: findRowByNameAnywhere(rows, "Лавлагаа") || [],
    Гомдол: findRowByNameAnywhere(rows, "Гомдол") || [],
  };
  return {
    labels: last4.map((x) => String(x.label)),
    series: [
      {
        name: "Гомдол",
        data: last4.map((x) => nnum(wantRows["Гомдол"][x.col])),
      },
      {
        name: "Лавлагаа",
        data: last4.map((x) => nnum(wantRows["Лавлагаа"][x.col])),
      },
    ],
  };
}

// Fallback: prev/curr хоёр долоо хоног (Osticket1)
function lastWeeksFromPrevCurrFallback(
  prevFile,
  currFile,
  sheetName,
  companyFilter
) {
  const prev = countByCategoryWithinFile(prevFile, sheetName, companyFilter);
  const curr = countByCategoryWithinFile(currFile, sheetName, companyFilter);
  const wPrev = parseWeekFromFilename(prevFile)?.raw || "Өмнөх 7 хоног";
  const wCurr = parseWeekFromFilename(currFile)?.raw || "Одоогийн 7 хоног";
  return {
    labels: [wPrev, wCurr],
    series: [
      { name: "Гомдол", data: [prev["Гомдол"] || 0, curr["Гомдол"] || 0] },
      {
        name: "Лавлагаа",
        data: [prev["Лавлагаа"] || 0, curr["Лавлагаа"] || 0],
      },
    ],
  };
}

// ТОП хүснэгтүүд (prev vs curr) — туслах ангиллаар
function buildTopFromTwoFiles(
  prevFile,
  currFile,
  sheetName,
  companyFilter,
  limitPerGroup = 10
) {
  const readOne = (file) => {
    const wb = xlsx.readFile(file, { cellDates: true });
    const ws = wb.Sheets[sheetName];
    if (!ws) throw new Error(`Sheet not found: ${sheetName} (${file})`);
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
    const hdr = rows[0].map((x) => String(x || ""));

    const idx = {
      company: hdr.findIndex((h) => /Компани/i.test(h)),
      category: hdr.findIndex((h) => /Ангилал/i.test(h)),
      subcat: hdr.findIndex((h) => /Туслах\s*ангилал/i.test(h)),
      created: hdr.findIndex(
        (h) =>
          /Үүссэн\s*огноо/i.test(h) ||
          /Нээсэн\s*огноо/i.test(h) ||
          /Created/i.test(h)
      ),
      closed: hdr.findIndex(
        (h) => /Хаагдсан\s*огноо/i.test(h) || /Closed/i.test(h)
      ),
    };
    const dateCol = idx.created >= 0 ? idx.created : idx.closed;

    const wk = parseWeekFromFilename(file);
    const year = inferYearFromSheet(ws, [dateCol]);
    const start = wk ? dayjs(makeYmd(year, wk.m1, wk.d1)).startOf("day") : null;
    const end = wk ? dayjs(makeYmd(year, wk.m2, wk.d2)).endOf("day") : null;

    const bag = { Лавлагаа: new Map(), Гомдол: new Map() };

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r] || [];
      if (companyFilter) {
        const comp = String(row[idx.company] || "").trim();
        if (comp !== companyFilter) continue;
      }
      const when = dateCol >= 0 ? parseExcelDate(row[dateCol]) : null;
      if (start && end && !inRangeInclusive(when, start, end)) continue;

      const cat = String(row[idx.category] || "").trim();
      const sub = String(row[idx.subcat] || "").trim();
      if (!sub || !bag[cat]) continue;

      bag[cat].set(sub, (bag[cat].get(sub) || 0) + 1);
    }
    return { bag, label: wk?.raw || "7 хоног" };
  };

  const prev = readOne(prevFile);
  const curr = readOne(currFile);

  const joinTop = (cat) => {
    const mapPrev = prev.bag[cat] || new Map();
    const mapCurr = curr.bag[cat] || new Map();
    const names = new Set([...mapPrev.keys(), ...mapCurr.keys()]);
    const arr = [...names].map((nm) => {
      const a = mapPrev.get(nm) || 0;
      const b = mapCurr.get(nm) || 0;
      const base = a > 0 ? a : b > 0 ? b : 1;
      const delta = (b - a) / base;
      return { name: nm, prev: a, curr: b, delta };
    });
    arr.sort((x, y) => y.curr - x.curr || y.prev - x.prev);
    return arr.slice(0, limitPerGroup);
  };

  return {
    labels: [prev.label, curr.label],
    groups: {
      Лавлагаа: joinTop("Лавлагаа"),
      Гомдол: joinTop("Гомдол"),
    },
  };
}

// ────────────────────────────────────────────────────────────────
// HTML
// ────────────────────────────────────────────────────────────────
function pct100(x, d = 0) {
  return `${(x * 100).toFixed(d)}%`;
}
function updown(v) {
  return v >= 0 ? "▲" : "▼";
}
function updownCls(v) {
  return v >= 0 ? "up" : "down";
}
function num(n) {
  return Number(n || 0).toLocaleString();
}

function renderTopTableBlock(title, labels, rows) {
  const head = `
    <thead>
      <tr>
        <th width="50%">${title}</th>
        <th>${labels[0]}</th>
        <th>${labels[1]}</th>
        <th>%</th>
      </tr>
    </thead>`;
  const body = rows
    .map(
      (r) => `
      <tr>
        <td>${r.name ? String(r.name) : ""}</td>
        <td class="num">${num(r.prev)}</td>
        <td class="num">${num(r.curr)}</td>
        <td class="num ${updownCls(r.delta)}">${updown(r.delta)} ${pct100(
        Math.abs(r.delta),
        0
      )}</td>
      </tr>`
    )
    .join("");

  return `
    <div class="card">
      <div class="table-wrap">
        <table class="cmp">
          ${head}
          <tbody>${body}</tbody>
        </table>
      </div>
    </div>`;
}

function renderAllTopTables(top) {
  return `
  <section class="sheet">
    <div class="card head red">ТОП Лавлагаа</div>
    ${renderTopTableBlock("ТОП Лавлагаа", top.labels, top.groups["Лавлагаа"])}
    <div class="spacer16"></div>

    <div class="card head red">ТОП Гомдол</div>
    ${renderTopTableBlock("ТОП Гомдол", top.labels, top.groups["Гомдол"])}
  </section>`;
}

function buildHtml(payload, cssText) {
  const libs = CONFIG.JS_LIBS.map((s) => `<script src="${s}"></script>`).join(
    "\n"
  );

  const weeks = payload.weeks; // {labels, series[ Gomdol, Lavlagaa ]}
  const currTotal =
    (weeks.series?.[0]?.data?.slice(-1)[0] || 0) +
    (weeks.series?.[1]?.data?.slice(-1)[0] || 0);
  const prevTotal =
    (weeks.series?.[0]?.data?.slice(-2)[0] || 0) +
    (weeks.series?.[1]?.data?.slice(-2)[0] || 0);
  const totDelta = prevTotal ? (currTotal - prevTotal) / prevTotal : 0;

  const extraCSS = `
    body{font-family:Arial,Helvetica,sans-serif}
    .sheet{margin-top:16px}
    .grid{display:grid;gap:16px}
    .grid-2{grid-template-columns:1.2fr .8fr}
    .grid-1-1{grid-template-columns:1fr 1fr}
    .card{background:#fff;border:1px solid #e5e7eb;border-radius:8px;padding:12px;break-inside:avoid}
    .head{font-weight:700}
    .red{background:#0ea5e9;color:#fff;padding:10px 12px;border:none}
    .bullet{margin:6px 0 0 18px;padding:0}
    .bullet li{margin:6px 0}
    table.cmp{width:100%;border-collapse:collapse;font-size:12.5px}
    table.cmp th, table.cmp td{border:1px solid #e5e7eb;padding:6px 8px;vertical-align:top}
    table.cmp thead th{background:#e0f2fe}
    td.num{text-align:right;white-space:nowrap}
    .kpi{font-weight:700}
    .spacer16{height:16px}
    .footer{margin-top:24px;color:#6b7280;font-size:11px}
    .badge{display:inline-block;padding:2px 6px;border-radius:6px;color:#fff;font-size:11px}
    .up{background:#16a34a}.down{background:#ef4444}
  `;

  const deltaBadge = (prev, curr) => {
    const base = prev > 0 ? prev : curr || 1;
    const d = (curr - prev) / base;
    const cls = d >= 0 ? "up" : "down";
    return `<span class="badge ${cls}">${pct100(Math.abs(d), 0)}</span>`;
  };

  // mini (лавлагаа/гомдол) – prev vs curr
  const lavPrev = weeks.series?.[1]?.data?.slice(-2)[0] || 0;
  const lavCurr = weeks.series?.[1]?.data?.slice(-1)[0] || 0;
  const gomPrev = weeks.series?.[0]?.data?.slice(-2)[0] || 0;
  const gomCurr = weeks.series?.[0]?.data?.slice(-1)[0] || 0;

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <style>${cssText}\n${extraCSS}</style>
  ${libs}
  <title>Ard Leasing Report</title>
</head>
<body>

  <div class="card head red">АРД ЛИЗИНГ</div>

  <!-- Top row -->
  <section class="sheet">
    <div class="grid grid-2">
      <div class="card">
        <div class="card head red">НИЙТ ХАНДАЛТ /Сүүлийн 4 сараар/</div>
        <div id="leasing-handalt" style="height:230px;margin-top:8px"></div>
      </div>
      <div class="card">
        <div class="card head red">НИЙТ БҮРТГЭЛ /Сүүлийн 4 долоо хоногоор/</div>
        <div id="leasing-week" style="height:230px;margin-top:8px"></div>
        <ul class="bullet">
          <li>Тайлант 7 хоногт <span class="kpi">${num(
            currTotal
          )}</span> бүртгэл хийгдсэн.
            Өмнөх 7 хоногоос <span class="${
              totDelta >= 0 ? "up" : "down"
            }">${pct100(totDelta, 0)}</span> ${
    totDelta >= 0 ? "өссөн" : "буурсан"
  }.</li>
          <li>Лавлагаа: ${deltaBadge(lavPrev, lavCurr)} • Гомдол: ${deltaBadge(
    gomPrev,
    gomCurr
  )}</li>
        </ul>
      </div>
    </div>
  </section>

  <!-- Mini blocks -->
  <section class="sheet">
    <div class="grid grid-1-1">
      <div class="card">
        <h3 style="margin:0">ЛАВЛАГАА – ${num(lavCurr)}</h3>
        <div id="leasing-mini-lav" style="height:120px"></div>
      </div>
      <div class="card">
        <h3 style="margin:0">ГОМДОЛ – ${num(gomCurr)}</h3>
        <div id="leasing-mini-gom" style="height:120px"></div>
      </div>
    </div>
  </section>

  ${renderAllTopTables(payload.top)}

  <div class="footer">Автоматаар бэлтгэсэн тайлан (Ard Leasing)</div>

  <script>
    (function(){
      // months (line)
      new ApexCharts(document.querySelector("#leasing-handalt"), {
        series: [{ name:"Нийт", data: ${JSON.stringify(payload.month4.data)} }],
        chart: { height: 220, type: "line", toolbar: { show:false } },
        dataLabels: { enabled: true },
        stroke: { curve: "straight", width: 3 },
        markers: { size: 4 },
        grid: { row:{ colors:["#f3f3f3","transparent"], opacity: .5 } },
        xaxis: { categories: ${JSON.stringify(payload.month4.labels)} },
        yaxis: { min: 0 }
      }).render();

      // weeks (bar)
      new ApexCharts(document.querySelector("#leasing-week"), {
        chart: { type:"bar", height:220, stacked:false, toolbar:{show:false} },
        plotOptions: { bar:{ horizontal:false, columnWidth:"55%", endingShape:"rounded" } },
        dataLabels: { enabled:true, style:{colors:["#fff"]},
          background:{ enabled:true, foreColor:"#000", padding:4, borderRadius:4, borderWidth:1, borderColor:"#1E90FF", opacity:.9 } },
        stroke: { show:true, width:2, colors:["transparent"] },
        series: ${JSON.stringify(payload.weeks.series)},
        xaxis: { categories: ${JSON.stringify(payload.weeks.labels)} },
        fill: { opacity:1 },
        legend: { position:"bottom" }
      }).render();

      const mkMini = (el, cats, vals)=> new ApexCharts(document.querySelector(el), {
        series: [{ name:"Value", data: vals }],
        chart: { type:"bar", height:110, toolbar:{show:false} },
        plotOptions: { bar:{ horizontal:false, columnWidth:"70%", borderRadius:6 } },
        dataLabels: { enabled:true },
        xaxis: { categories: cats },
        yaxis: { min:0 },
        colors: ["#546E7A"]
      }).render();

      const miniCats = ${JSON.stringify(payload.weeks.labels.slice(-2))};
      mkMini("#leasing-mini-lav", miniCats, [${lavPrev}, ${lavCurr}]);
      mkMini("#leasing-mini-gom", miniCats, [${gomPrev}, ${gomCurr}]);
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
    console.log("[EMAIL] Disabled");
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
    html: `<p>Сайн байна уу,</p><p>Ард Лизинг 7 хоногийн тайланг хавсаргав.</p><p style="color:#666;font-size:12px">Автоматаар илгээв.</p>`,
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

  const top = buildTopFromTwoFiles(
    CONFIG.PREV_FILE,
    CONFIG.CURR_FILE,
    CONFIG.APP_SHEET,
    CONFIG.COMPANY_FILTER,
    10
  );

  let weeks = weeks4FromALS(CONFIG.CURR_FILE, CONFIG.ALS_SHEET);
  if (!weeks)
    weeks = lastWeeksFromPrevCurrFallback(
      CONFIG.PREV_FILE,
      CONFIG.CURR_FILE,
      CONFIG.APP_SHEET,
      CONFIG.COMPANY_FILTER
    );

  let month4 = month4FromALS(
    CONFIG.CURR_FILE,
    CONFIG.ALS_SHEET,
    dayjs().year()
  );
  if (!month4) month4 = { labels: [], data: [] };

  const cssText = fs.readFileSync(CONFIG.CSS_FILE, "utf-8");
  const html = buildHtml({ weeks, month4, top }, cssText);

  const monday = dayjs().tz(CONFIG.TIMEZONE).startOf("week").add(1, "day");
  const pdfName = `ardleasing-weekly-${monday.format("YYYYMMDD")}.pdf`;
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
