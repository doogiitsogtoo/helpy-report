// ard-security.js — Osticket1 (prev & curr) + ASC (4 months/weeks) → PDF + Email
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

  CURR_FILE: process.env.CURR_FILE || "./ARD 8.04-8.10.xlsx",
  PREV_FILE: process.env.PREV_FILE || "./ARD 7.28-8.03.xlsx",

  APP_SHEET: process.env.APP_SHEET || "Osticket1", // raw tickets
  ASC_SHEET: process.env.ASC_SHEET || "ASC", // aggregated (months + weeks)

  OUT_DIR: process.env.OUT_DIR || "./out",
  REPORT_TITLE: process.env.REPORT_TITLE || "Ард Секюритиз — 7 хоногийн тайлан",
  SUBJECT_PREFIX: process.env.SUBJECT_PREFIX || "[ArdSecurity Weekly]",

  EMAIL_ENABLED: String(process.env.EMAIL_ENABLED ?? "true") === "true",
  SCHED_ENABLED: String(process.env.SCHED_ENABLED ?? "true") === "true",

  CSS_FILE: process.env.CSS_FILE || "./css/template.css",
  COMPANY_FILTER: process.env.COMPANY_FILTER || "Ард Секюритиз",

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

// Osticket1 → ангиллын тоонууд (filename week)
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

// ASC → 4 months (current year column preferred; ENV ASC_YEAR_COL for 0-based index)
function month4FromASC(file, sheetName, preferYear = dayjs().year()) {
  const wb = xlsx.readFile(file, { cellDates: true });
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  if (!rows.length) return null;

  // 1) year column
  let yearCol = Number(process.env.ASC_YEAR_COL);
  let headerRow = -1;
  const looksLikeYear = (v) =>
    /^\d{4}$/.test(String(v || "").trim()) &&
    +String(v).trim() >= 2019 &&
    +String(v).trim() <= 2100;

  if (!Number.isInteger(yearCol)) {
    for (let r = 0; r < Math.min(rows.length, 25); r++) {
      const row = rows[r] || [];
      for (let c = 0; c < row.length; c++) {
        if (
          looksLikeYear(row[c]) &&
          String(row[c]).trim() === String(preferYear)
        ) {
          yearCol = c;
          headerRow = r;
          break;
        }
      }
      if (yearCol >= 0) break;
    }
  }
  if (!Number.isInteger(yearCol)) {
    // fallback: last year-looking cell
    for (let r = 0; r < Math.min(rows.length, 25); r++) {
      const row = rows[r] || [];
      for (let c = row.length - 1; c >= 0; c--) {
        if (looksLikeYear(row[c])) {
          yearCol = c;
          headerRow = r;
          break;
        }
      }
      if (Number.isInteger(yearCol)) break;
    }
  }
  if (!Number.isInteger(yearCol)) {
    // H (0-based 7) эсвэл F (5) зэрэг таны файлын стандарт байршилд тааруулж болно.
    yearCol = 5;
    headerRow = 1;
  }

  // 2) month rows
  const monthFromCell = (v) => {
    const s = String(v ?? "")
      .trim()
      .toLowerCase();
    const m1 = s.match(/^(\d{1,2})\s*сар$/);
    const m2 = s.match(/^(\d{1,2})сар$/);
    const m3 = s.match(/^0?(\d{1,2})$/);
    const m = m1 ? +m1[1] : m2 ? +m2[1] : m3 ? +m3[1] : null;
    return m && m >= 1 && m <= 12 ? m : null;
  };

  const mmMap = new Map();
  for (let r = headerRow + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const first = String(row[0] || "")
      .trim()
      .toLowerCase();
    if (first === "нийт" || first === "niit") break;
    const m = monthFromCell(row[0]) || monthFromCell(row[1]);
    if (!m) continue;
    mmMap.set(m, nnum(row[yearCol]));
  }
  if (!mmMap.size) return { labels: [], data: [] };

  const currM = dayjs().tz(CONFIG.TIMEZONE).month() + 1;
  const track = [];
  for (let m = 1; m <= currM; m++)
    if (mmMap.has(m)) track.push([m, mmMap.get(m)]);
  const last4 = track.slice(-4);
  return {
    labels: last4.map(([m]) => `${m}сар`),
    data: last4.map(([, v]) => Number(v) || 0),
  };
}

// ASC → 4 weeks stacked (Lavlagaa/Uilchilgee/Gomdol)
function last4WeeksByCategoryFromASC(file, sheetName) {
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
    Үйлчилгээ: findRowByNameAnywhere(rows, "Үйлчилгээ") || [],
    Гомдол: findRowByNameAnywhere(rows, "Гомдол") || [],
  };

  return {
    labels: last4.map((x) => String(x.label)),
    series: [
      {
        name: "Үйлчилгээ",
        data: last4.map((x) => nnum(wantRows["Үйлчилгээ"][x.col])),
      },
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

// Fallback: 2 weeks only from prev/curr Osticket1
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
      {
        name: "Үйлчилгээ",
        data: [prev["Үйлчилгээ"] || 0, curr["Үйлчилгээ"] || 0],
      },
      { name: "Гомдол", data: [prev["Гомдол"] || 0, curr["Гомдол"] || 0] },
      {
        name: "Лавлагаа",
        data: [prev["Лавлагаа"] || 0, curr["Лавлагаа"] || 0],
      },
    ],
  };
}

// TOP хүснэгтүүд (prev vs curr) — туслах ангиллаар
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

    const bag = {
      Лавлагаа: new Map(),
      Үйлчилгээ: new Map(),
      Гомдол: new Map(),
    };

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
      Үйлчилгээ: joinTop("Үйлчилгээ"),
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

    <div class="card head red">ТОП Үйлчилгээ</div>
    ${renderTopTableBlock("ТОП Үйлчилгээ", top.labels, top.groups["Үйлчилгээ"])}
    <div class="spacer16"></div>

    <div class="card head red">ТОП Гомдол</div>
    ${renderTopTableBlock("ТОП Гомдол", top.labels, top.groups["Гомдол"])}
  </section>`;
}

function buildHtml(payload, cssText) {
  const libs = CONFIG.JS_LIBS.map((s) => `<script src="${s}"></script>`).join(
    "\n"
  );

  const totPrev =
    payload.prevCat["Лавлагаа"] +
    payload.prevCat["Үйлчилгээ"] +
    payload.prevCat["Гомдол"];
  const totCurr =
    payload.currCat["Лавлагаа"] +
    payload.currCat["Үйлчилгээ"] +
    payload.currCat["Гомдол"];
  const totDelta = totPrev > 0 ? (totCurr - totPrev) / totPrev : 0;

  const cG = payload.currCat["Гомдол"];
  const pG = payload.prevCat["Гомдол"];
  const gDelta = pG > 0 ? (cG - pG) / pG : 0;

  const mini = payload.mini;

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
    .mini-row{display:flex;align-items:center;gap:8px}
    .badge{display:inline-block;padding:2px 6px;border-radius:6px;color:#fff;font-size:11px}
    .badge.up{background:#16a34a}.badge.down{background:#ef4444}
    .footer{margin-top:24px;color:#6b7280;font-size:11px}
  `;

  const miniBadge = (prev, curr) => {
    const base = prev > 0 ? prev : curr || 1;
    const d = (curr - prev) / base;
    const cls = d >= 0 ? "up" : "down";
    return `<span class="badge ${cls}">${pct100(Math.abs(d), 0)}</span>`;
  };

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <style>${cssText}\n${extraCSS}</style>
  ${libs}
  <title>Ard Security Report</title>
</head>
<body>

  <div class="card head red">АРД СЕКЮРИТИЗ</div>

  <!-- Top row -->
  <section class="sheet">
    <div class="grid grid-2">
      <div class="card">
        <div class="card head red">НИЙТ ХАНДАЛТ /Сүүлийн 4 сараар/</div>
        <div id="sec-handalt" style="height:230px;margin-top:8px"></div>
      </div>
      <div class="card">
        <div class="card head red">НИЙТ БҮРТГЭЛ /Сүүлийн 4 долоо хоногоор/</div>
        <div id="sec-week" style="height:230px;margin-top:8px"></div>
        <ul class="bullet">
          <li>Тайлант хугацаанд <span class="kpi">${num(
            totCurr
          )}</span> харилцагчид үйлчилсэн бөгөөд өмнөх 7 хоногоос <span class="${
    totDelta >= 0 ? "up" : "down"
  }">${pct100(totDelta, 1)}</span> ${totDelta >= 0 ? "өссөн" : "буурсан"}.</li>
          <li>Нийт <span class="kpi">${num(
            cG
          )}</span> гомдол бүртгэгдсэн бөгөөд өмнөх 7 хоногтой харьцуулахад <span class="${
    gDelta >= 0 ? "up" : "down"
  }">${pct100(gDelta, 0)}</span> ${gDelta >= 0 ? "өссөн" : "буурсан"}.</li>
        </ul>
      </div>
    </div>
  </section>

  <!-- Mini bars -->
  <section class="sheet">
    <div class="grid grid-1-1">
      <div class="card">
        <div class="mini-row">
          <h3 style="margin:0">ЛАВЛАГАА – ${num(
            payload.currCat["Лавлагаа"]
          )} (${pct100(payload.currCat["Лавлагаа"] / (totCurr || 1), 0)})</h3>
          ${miniBadge(mini.lavlagaa.prev, mini.lavlagaa.curr)}
        </div>
        <div id="sec-lavlagaa" style="height:120px"></div>
      </div>
      <div class="card">
        <div class="mini-row">
          <h3 style="margin:0">ҮЙЛЧИЛГЭЭ – ${num(
            payload.currCat["Үйлчилгээ"]
          )} (${pct100(payload.currCat["Үйлчилгээ"] / (totCurr || 1), 0)})</h3>
          ${miniBadge(mini.uilchilgee.prev, mini.uilchilgee.curr)}
        </div>
        <div id="sec-uilchilgee" style="height:120px"></div>
      </div>
    </div>
    <div class="card" style="margin-top:12px">
      <div class="mini-row">
        <h3 style="margin:0">ГОМДОЛ – ${num(
          payload.currCat["Гомдол"]
        )} (${pct100(payload.currCat["Гомдол"] / (totCurr || 1), 0)})</h3>
        ${miniBadge(mini.gomdol.prev, mini.gomdol.curr)}
      </div>
      <div id="sec-gomdol" style="height:120px"></div>
    </div>
  </section>

  ${renderAllTopTables(payload.top)}

  <div class="footer">Автоматаар бэлтгэсэн тайлан (Ard Securities)</div>

  <script>
    (function(){
      // 4 months
      new ApexCharts(document.querySelector("#sec-handalt"), {
        series: [{ name:"Нийт", data: ${JSON.stringify(payload.month4.data)} }],
        chart: { height: 220, type: "line", toolbar: { show:false } },
        dataLabels: { enabled: true },
        stroke: { curve: "straight", width: 3 },
        markers: { size: 4 },
        grid: { row:{ colors:["#f3f3f3","transparent"], opacity: .5 } },
        xaxis: { categories: ${JSON.stringify(payload.month4.labels)} },
        yaxis: { min: 0 }
      }).render();

      // 4 weeks stacked
      new ApexCharts(document.querySelector("#sec-week"), {
        chart: { type:"bar", height:220, stacked:false, toolbar:{show:false} },
        plotOptions: { bar:{ horizontal:false, columnWidth:"55%", endingShape:"rounded" } },
        dataLabels: { enabled:true, style:{colors:["#fff"]}, background:{ enabled:true, foreColor:"#000", padding:4, borderRadius:4, borderWidth:1, borderColor:"#1E90FF", opacity:.9 } },
        stroke: { show:true, width:2, colors:["transparent"] },
        series: ${JSON.stringify(payload.weeks4.series)},
        xaxis: { categories: ${JSON.stringify(payload.weeks4.labels)} },
        fill: { opacity:1 },
        legend: { position:"bottom" }
      }).render();

      // mini bars
      const mkMini = (el, cats, vals)=> new ApexCharts(document.querySelector(el), {
        series: [{ name:"Value", data: vals }],
        chart: { type:"bar", height:110, toolbar:{show:false} },
        plotOptions: { bar:{ horizontal:false, columnWidth:"70%", borderRadius:6 } },
        dataLabels: { enabled:true },
        xaxis: { categories: cats },
        yaxis: { min:0 },
        colors: ["#546E7A"]
      }).render();

      const mini = ${JSON.stringify(payload.mini)};
      mkMini("#sec-lavlagaa", mini.labels, [mini.lavlagaa.prev, mini.lavlagaa.curr]);
      mkMini("#sec-uilchilgee", mini.labels, [mini.uilchilgee.prev, mini.uilchilgee.curr]);
      mkMini("#sec-gomdol", mini.labels, [mini.gomdol.prev, mini.gomdol.curr]);
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
    html: `<p>Сайн байна уу,</p><p>Ард Секюритиз 7 хоногийн тайланг хавсаргав.</p><p style="color:#666;font-size:12px">Автоматаар илгээв.</p>`,
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

  // TOP хүснэгтүүд
  const top = buildTopFromTwoFiles(
    CONFIG.PREV_FILE,
    CONFIG.CURR_FILE,
    CONFIG.APP_SHEET,
    CONFIG.COMPANY_FILTER,
    10
  );

  // prev/curr ангиллын нийлбэр
  const prevCat = countByCategoryWithinFile(
    CONFIG.PREV_FILE,
    CONFIG.APP_SHEET,
    CONFIG.COMPANY_FILTER
  );
  const currCat = countByCategoryWithinFile(
    CONFIG.CURR_FILE,
    CONFIG.APP_SHEET,
    CONFIG.COMPANY_FILTER
  );

  const mini = {
    labels: [
      parseWeekFromFilename(CONFIG.PREV_FILE)?.raw || "Prev",
      parseWeekFromFilename(CONFIG.CURR_FILE)?.raw || "Curr",
    ],
    lavlagaa: {
      prev: prevCat["Лавлагаа"] || 0,
      curr: currCat["Лавлагаа"] || 0,
    },
    uilchilgee: {
      prev: prevCat["Үйлчилгээ"] || 0,
      curr: currCat["Үйлчилгээ"] || 0,
    },
    gomdol: { prev: prevCat["Гомдол"] || 0, curr: currCat["Гомдол"] || 0 },
  };

  // 4 months — ASC > fallback (Osticket monthly from app)
  let month4 = month4FromASC(
    CONFIG.CURR_FILE,
    CONFIG.ASC_SHEET,
    dayjs().year()
  );
  if (!month4) month4 = { labels: [], data: [] }; // ASC заавал байх ёстой — хоосон бол график цулгуй

  // 4 weeks — ASC > fallback (prev/curr)
  let weeks4 = last4WeeksByCategoryFromASC(CONFIG.CURR_FILE, CONFIG.ASC_SHEET);
  if (!weeks4)
    weeks4 = lastWeeksFromPrevCurrFallback(
      CONFIG.PREV_FILE,
      CONFIG.CURR_FILE,
      CONFIG.APP_SHEET,
      CONFIG.COMPANY_FILTER
    );

  const wkCurr =
    parseWeekFromFilename(CONFIG.CURR_FILE)?.raw || "Одоогийн 7 хоног";
  const wkPrev =
    parseWeekFromFilename(CONFIG.PREV_FILE)?.raw || "Өмнөх 7 хоног";

  const cssText = fs.readFileSync(CONFIG.CSS_FILE, "utf-8");
  const html = buildHtml(
    {
      weekPrev: wkPrev,
      weekCurr: wkCurr,
      month4,
      weeks4,
      mini,
      top,
      prevCat,
      currCat,
    },
    cssText
  );

  const monday = dayjs().tz(CONFIG.TIMEZONE).startOf("week").add(1, "day");
  const pdfName = `ardsecurity-weekly-${monday.format("YYYYMMDD")}.pdf`;
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

// Scheduler
function startScheduler() {
  if (!CONFIG.SCHED_ENABLED) {
    console.log("Scheduler disabled (SCHED_ENABLED=false).");
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
