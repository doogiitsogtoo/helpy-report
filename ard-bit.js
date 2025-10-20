// ard-bit.dynamic.js — АРД БИТ (ARB + Osticket1) → HTML → PDF (+optional email)
// ---------------------------------------------------------------------------------
// ✔ Таны ard-app.dynamic.js-тэй ижил санаануудыг АРД БИТ-д нутагшуулав:
//   - PREV_FILE хоосон бол → ижил хавтаснаас өмнөх 7 хоногийн Excel-ийг нэрээс нь автоматаар олно
//   - ASS_YEAR = "auto" → ARB шитийн баруун талын идэвхтэй жилээр автоматаар сонгоно
//   - ARB: сүүлийн N сар (default 4) шугаман график
//   - ARB: баруунаас 4 week-like багана олдвол (Лавлагаа/Үйлчилгээ/Гомдол) стэктэй багана; олдохгүй бол Osticket1 prev/curr fallback
//   - Osticket1: ТОП (туслах/дэд ангилал) — өмнөх vs одоогийн долоо хоног, ASS_COMPANY-гаар шүүнэ
//   - HTML→PDF (Puppeteer), и-мэйл илгээх сонголт, Даваа 09:00 scheduler
// ---------------------------------------------------------------------------------

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

  // Excel files (хоосон PREV_FILE → автоматаар олоно)
  PREV_FILE: process.env.PREV_FILE || "./ARD 09.22-09.28.xlsx",
  CURR_FILE: process.env.CURR_FILE || "./ARD 09.29-10.05.xlsx",
  GOMDOL_FILE: process.env.GOMDOL_FILE || "", // reserved

  // Sheets
  ASS_SHEET: process.env.ASS_SHEET || "ARB",
  ASS_COMPANY: process.env.ASS_COMPANY || "Ард Бит",
  ASS_YEAR: process.env.ASS_YEAR || "auto", // "auto" → баруун талын идэвхтэй жил
  ASS_TAKE_LAST_N_MONTHS: Number(process.env.ASS_TAKE_LAST_N_MONTHS || 4),
  OST_SHEET: process.env.OST_SHEET || "Osticket1",

  // PDF / Email
  OUT_DIR: process.env.OUT_DIR || "./out",
  CSS_FILE: process.env.CSS_FILE || "./css/template.css",
  SAVE_HTML: String(process.env.SAVE_HTML ?? "true") === "true",
  HTML_NAME_PREFIX: process.env.HTML_NAME_PREFIX || "report",
  REPORT_TITLE: process.env.REPORT_TITLE || "АРД БИТ — Харилцагчийн үйлчилгээ",
  SUBJECT_PREFIX: process.env.SUBJECT_PREFIX || "[ArdBit Weekly]",
  EMAIL_ENABLED: String(process.env.EMAIL_ENABLED ?? "false") === "true",
  SCHED_ENABLED: String(process.env.SCHED_ENABLED ?? "true") === "true",
};

// ────────────────────────────────────────────────────────────────
// Helpers
// ────────────────────────────────────────────────────────────────
const asNum = (v) =>
  Number(
    String(v ?? "")
      .replace(/[\s,\u00A0]/g, "")
      .replace(/[^\d.-]/g, "")
  ) || 0;
const norm = (s) =>
  String(s || "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
const pad2 = (n) => String(n).padStart(2, "0");

function escapeHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function parseExcelDate(v) {
  if (v == null || v === "") return null;
  if (v instanceof Date) {
    const d = dayjs(v);
    return d.isValid() ? d : null;
  }
  const n = Number(v);
  if (Number.isFinite(n) && n > 20000) {
    const ms = (n - 25569) * 86400 * 1000; // Excel serial
    const d = dayjs(new Date(ms));
    return d.isValid() ? d : null;
  }
  const d = dayjs(v);
  return d.isValid() ? d : null;
}

function parseWeekFromFilename(p) {
  if (!p) return null;
  const b = path.basename(p);
  const m = b.match(/(\d{1,2})[./-](\d{1,2})\s*[-–]\s*(\d{1,2})[./-](\d{1,2})/);
  if (!m) return null;
  return {
    m1: +m[1],
    d1: +m[2],
    m2: +m[3],
    d2: +m[4],
    raw: `${pad2(m[1])}.${pad2(m[2])}-${pad2(m[3])}.${pad2(m[4])}`,
  };
}
function labelFromRange(start, end) {
  if (!start || !end) return "";
  return `${pad2(start.month() + 1)}.${pad2(start.date())}-${pad2(
    end.month() + 1
  )}.${pad2(end.date())}`;
}

function inferYearFromSheet(ws, dateColIndexes = []) {
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  for (let r = 1; r < Math.min(rows.length, 100); r++) {
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

function getLastWeekLikeFromHeader(header) {
  const isWeek = (s) =>
    /(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)\s*[-–]\s*(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)/.test(
      String(s || "")
    );
  for (let i = header.length - 1; i >= 0; i--)
    if (isWeek(header[i])) return String(header[i]).trim();
  return null;
}

function autoFindPrevFile(currFile) {
  if (!currFile) return "";
  const dir = path.dirname(currFile);
  const base = path.basename(currFile);
  const files = fs.readdirSync(dir).filter((f) => /\.xlsx$/i.test(f));
  const wkCurr = parseWeekFromFilename(base);
  if (!wkCurr) return "";
  // infer year from the current file’s first sheet
  let year = dayjs().year();
  try {
    const wb = xlsx.readFile(currFile, { cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
    outer: for (let r = 1; r < Math.min(60, rows.length); r++) {
      for (const v of rows[r] || []) {
        const d = parseExcelDate(v);
        if (d) {
          year = d.year();
          break outer;
        }
      }
    }
  } catch {}

  const endCurr = dayjs(`${year}-${pad2(wkCurr.m2)}-${pad2(wkCurr.d2)}`);
  let best = null;
  for (const f of files) {
    if (f === base) continue;
    const wk = parseWeekFromFilename(f);
    if (!wk) continue;
    const end = dayjs(`${year}-${pad2(wk.m2)}-${pad2(wk.d2)}`);
    if (end.isBefore(endCurr)) {
      const gap = endCurr.diff(end, "day");
      if (!best || gap < best.gap) best = { f: path.join(dir, f), gap };
    }
  }
  return best ? best.f : "";
}

function detectRangeFromCurrFile(currFile) {
  const wk = parseWeekFromFilename(currFile);
  if (!wk) {
    const start = dayjs().tz(CONFIG.TIMEZONE).startOf("week").add(1, "day");
    const end = start.add(6, "day");
    return { start, end, label: labelFromRange(start, end) };
  }
  const wb = xlsx.readFile(currFile, { cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  const headers = rows[0] || [];
  const dateCol = headers.findIndex((h) =>
    /(Үүссэн\s*огноо|Нээсэн\s*огноо|Created|Open|Хаагдсан\s*огноо|Closed)/i.test(
      String(h || "")
    )
  );
  const year = inferYearFromSheet(ws, [dateCol]);
  const start = dayjs(`${year}-${pad2(wk.m1)}-${pad2(wk.d1)}`).startOf("day");
  const end = dayjs(`${year}-${pad2(wk.m2)}-${pad2(wk.d2)}`).endOf("day");
  return { start, end, label: labelFromRange(start, end) };
}

function parseWeekLabelCell(s) {
  const m = String(s || "").match(
    /(\d{1,2})[./-](\d{1,2}).*?[-–].*?(\d{1,2})[./-](\d{1,2})/
  );
  if (!m) return null;
  return { m1: +m[1], d1: +m[2], m2: +m[3], d2: +m[4], raw: m[0] };
}

function indexOfHeader(headers, regexes) {
  const hdrs = headers.map((x) => String(x || ""));
  for (let i = 0; i < hdrs.length; i++)
    if (regexes.some((re) => re.test(hdrs[i]))) return i;
  return -1;
}

function findRowByKeywords(rows, keywords) {
  const want = keywords.map(norm);
  for (const r of rows) {
    const cells = (r || []).slice(0, 5);
    for (const v of cells) {
      const x = norm(v);
      if (!x) continue;
      if (want.some((w) => x === w || x.startsWith(w) || x.includes(w)))
        return r;
    }
  }
  return null;
}

// ────────────────────────────────────────────────────────────────
// Extractors — ARB sheet (months & weeks)
// ────────────────────────────────────────────────────────────────
function extractARB_MonthsLatestN(
  file,
  sheetName,
  yearLabel = "auto",
  takeLast = 4
) {
  const wb = xlsx.readFile(file, { cellDates: true });
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`[ARB] Sheet not found: ${sheetName}`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  if (!rows.length)
    return { year: String(yearLabel), points: [], allMonths: [] };

  // Row with year headers (… 2023 | 2024 | 2025 …)
  const yearHeadRowIdx = rows.findIndex(
    (r) => r && r.filter(Boolean).some((v) => /^\d{4}$/.test(String(v)))
  );
  if (yearHeadRowIdx < 0) throw new Error("[ARB] Year header row not found");
  const header = (rows[yearHeadRowIdx] || []).map((v) =>
    String(v || "").trim()
  );

  let year = String(yearLabel);
  if (year === "auto") {
    const yCols = header
      .map((h, i) => (/^\d{4}$/.test(h) ? i : -1))
      .filter((i) => i >= 0);
    let pick = null;
    for (const i of yCols) {
      const has = rows.slice(yearHeadRowIdx + 1).some((r) => asNum(r?.[i]) > 0);
      if (has) pick = i; // keep the rightmost with data
    }
    if (pick == null) pick = yCols.at(-1) ?? -1;
    if (pick < 0) throw new Error("[ARB] No usable year column");
    year = header[pick];
  }
  const yearCol = header.findIndex((v) => v === String(year));
  if (yearCol < 0) throw new Error(`[ARB] Year col not found: ${year}`);

  const monthLike = (s) =>
    /^\s*\d+\s*(сар|cap)?\s*$/i.test(String(s || "").trim());
  const pickLabel = (row) => {
    for (const c of [0, 1, 2])
      if (monthLike(row[c])) return String(row[c]).replace(/cap/i, "сар");
    return null;
  };

  const monthRows = [];
  for (let r = yearHeadRowIdx + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const label = pickLabel(row);
    if (!label) {
      if (monthRows.length) break;
      else continue;
    }
    const value = asNum(row[yearCol]);
    monthRows.push({ label, value });
  }

  const active = monthRows.filter((m) => m.value > 0);
  const points = active.slice(-takeLast);
  return { year, points, allMonths: monthRows };
}

function extractARB_Last4WeeksByCategory(file, sheetName) {
  const wb = xlsx.readFile(file, { cellDates: true });
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  if (!rows.length) return null;

  const weekCols = [];
  for (let c = 0; c < (rows[0] || []).length; c++) {
    let parsed = null;
    for (let r = 0; r < rows.length; r++) {
      const p = parseWeekLabelCell((rows[r] || [])[c]);
      if (p) {
        parsed = p;
        break;
      }
    }
    if (parsed) weekCols.push({ col: c, label: parsed.raw, parsed });
  }
  if (!weekCols.length) return null;
  weekCols.sort((a, b) => a.col - b.col);
  const last4 = weekCols.slice(-4);

  const rowLav = findRowByKeywords(rows, ["лавлагаа", "lavlagaa", "lavlaga"]);
  const rowUil = findRowByKeywords(rows, [
    "үйлчилгээ",
    "uilchilgee",
    "uilchilge",
  ]);
  const rowGom = findRowByKeywords(rows, ["гомдол", "gomdol"]);
  if (!rowLav && !rowUil && !rowGom) return null;

  const pick = (row) => last4.map((w) => asNum((row || [])[w.col]));
  return {
    labels: last4.map((w) => String(w.label)),
    lav: pick(rowLav || []),
    uil: pick(rowUil || []),
    gom: pick(rowGom || []),
  };
}

// ────────────────────────────────────────────────────────────────
// Osticket1 helpers (prev/curr fallback + TOP tables)
// ────────────────────────────────────────────────────────────────
function countByCategoryWithinFile(file, sheetName, companyFilter) {
  const compMatch = (cell, filter) => {
    if (!filter) return true;
    const a = norm(cell);
    const b = norm(filter);
    return !!a && (a === b || a.includes(b) || b.includes(a));
  };

  const wb = xlsx.readFile(file, { cellDates: true });
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet not found: ${sheetName} (${file})`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  if (!rows.length) return { Лавлагаа: 0, Үйлчилгээ: 0, Гомдол: 0 };

  const hdr = rows[0].map((x) => String(x || ""));
  const idx = {
    company: indexOfHeader(hdr, [/Компани/i, /Company/i]),
    category: indexOfHeader(hdr, [/Ангилал/i, /Category/i]),
    created: indexOfHeader(hdr, [
      /Үүссэн\s*огноо/i,
      /Нээсэн\s*огноо/i,
      /Created/i,
    ]),
    closed: indexOfHeader(hdr, [/Хаагдсан\s*огноо/i, /Closed/i]),
  };
  const dateCol = idx.created >= 0 ? idx.created : idx.closed;

  const wk = parseWeekFromFilename(file);
  const year = inferYearFromSheet(ws, [dateCol]);
  const start = wk
    ? dayjs(`${year}-${pad2(wk.m1)}-${pad2(wk.d1)}`).startOf("day")
    : null;
  const end = wk
    ? dayjs(`${year}-${pad2(wk.m2)}-${pad2(wk.d2)}`).endOf("day")
    : null;

  const acc = { Лавлагаа: 0, Үйлчилгээ: 0, Гомдол: 0 };
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r] || [];
    if (!compMatch(row[idx.company], companyFilter)) continue;
    const when = dateCol >= 0 ? parseExcelDate(row[dateCol]) : null;
    if (start && end && (!when || when.isBefore(start) || when.isAfter(end)))
      continue;
    const cat = String(row[idx.category] || "").trim();
    if (acc[cat] != null) acc[cat] += 1;
  }
  return acc;
}

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
    lav: [prev["Лавлагаа"] || 0, curr["Лавлагаа"] || 0],
    uil: [prev["Үйлчилгээ"] || 0, curr["Үйлчилгээ"] || 0],
    gom: [prev["Гомдол"] || 0, curr["Гомдол"] || 0],
  };
}

function buildTopFromTwoFiles(
  prevFile,
  currFile,
  sheetName,
  companyFilter,
  limitPerGroup = 10
) {
  const compMatch = (cell, filter) => {
    if (!filter) return true;
    const a = norm(cell);
    const b = norm(filter);
    return !!a && (a === b || a.includes(b) || b.includes(a));
  };

  const readOne = (file) => {
    const wb = xlsx.readFile(file, { cellDates: true });
    const ws = wb.Sheets[sheetName];
    if (!ws) throw new Error(`Sheet not found: ${sheetName} (${file})`);
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
    const hdr = rows[0] || [];
    const idx = {
      company: indexOfHeader(hdr, [/Компани/i, /Company/i]),
      category: indexOfHeader(hdr, [/Ангилал/i, /Category/i]),
      subcat: indexOfHeader(hdr, [
        /Туслах\s*ангилал/i,
        /Дэд\s*ангилал/i,
        /Sub.?category/i,
      ]),
      created: indexOfHeader(hdr, [
        /Үүссэн\s*огноо/i,
        /Нээсэн\s*огноо/i,
        /Created/i,
      ]),
      closed: indexOfHeader(hdr, [/Хаагдсан\s*огноо/i, /Closed/i]),
    };
    const dateCol = idx.created >= 0 ? idx.created : idx.closed;

    const wk = parseWeekFromFilename(file);
    const year = inferYearFromSheet(ws, [dateCol]);
    const start = wk
      ? dayjs(`${year}-${pad2(wk.m1)}-${pad2(wk.d1)}`).startOf("day")
      : null;
    const end = wk
      ? dayjs(`${year}-${pad2(wk.m2)}-${pad2(wk.d2)}`).endOf("day")
      : null;

    const bag = {
      Лавлагаа: new Map(),
      Үйлчилгээ: new Map(),
      Гомдол: new Map(),
    };
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r] || [];
      if (!compMatch(row[idx.company], companyFilter)) continue;
      const when = dateCol >= 0 ? parseExcelDate(row[dateCol]) : null;
      if (start && end && (!when || when.isBefore(start) || when.isAfter(end)))
        continue;

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
    const mP = prev.bag[cat] || new Map(),
      mC = curr.bag[cat] || new Map();
    const names = new Set([...mP.keys(), ...mC.keys()]);
    const arr = [...names].map((nm) => {
      const a = mP.get(nm) || 0,
        b = mC.get(nm) || 0;
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
// HTML (ard-app.js-тэй ижил хэв маяг: Bootstrap + Chart.js)
// ────────────────────────────────────────────────────────────────
function wrapHtml(bodyHtml) {
  const css = fs.readFileSync(CONFIG.CSS_FILE, "utf-8");
  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
  <style>${css}</style>
</head>
<body>
 <div class="container py-3">
    <div class="row g-3">
      ${bodyHtml}
      <div class="footer">Автоматаар бэлтгэсэн тайлан (Ard Bit)</div>
    </div>
  </div>
</body>
</html>`;
}
function renderAssCover({ company, periodText }) {
  return `
  <section class="hero" style="margin-bottom:16px">
    <div style="background:linear-gradient(135deg,#ef4444,#f97316);border-radius:12px;padding:28px;display:flex;justify-content:space-between;align-items:center;min-height:220px;">
      <div style="background:#fff;border-radius:16px;padding:20px 24px;display:inline-block">
        <div style="font-weight:700;font-size:28px;letter-spacing:.5px;color:#ef4444">ARD</div>
        <div style="color:#666;margin-top:4px">Хүчтэй. Хамтдаа.</div>
      </div>
      <div style="color:#fff;text-align:right;padding:8px 16px">
        <div style="font-size:36px;font-weight:800;line-height:1.1">${escapeHtml(
          company
        )}</div>
        <div style="opacity:.9;margin-top:8px">${escapeHtml(
          periodText || ""
        )}</div>
      </div>
    </div>
  </section>`;
}

function renderBitLayout({ monthN, weeks, top }) {
  // Chart.js config helpers
  const dataLbl = `
    const dataLabel={id:'dataLabel',afterDatasetsDraw(ch){
      const {ctx, data:{datasets},getDatasetMeta}=ch; ctx.save();
      ctx.font='12px system-ui,-apple-system,Segoe UI,Roboto,Arial'; ctx.textAlign='center';
      datasets.forEach((ds,di)=>{ const meta=getDatasetMeta(di);
        (ds.data||[]).forEach((v,i)=>{ if(v==null) return; const pt=meta.data[i];
          ctx.fillStyle='#111'; ctx.fillText(String(v), pt.x, pt.y-6);
        });
      }); ctx.restore();
    }};`;

  const lineCard = `
    <div class="card" style="height: 500px; margin-bottom: 4rem;">
      <div class="card-title">НИЙТ ХАНДАЛТ /Сүүлийн ${
        monthN.labels.length
      } сараар/</div>
      <canvas id="arbLine"></canvas>
    </div>
    <script>(function(){
      const ctx=document.getElementById('arbLine').getContext('2d');
      new Chart(ctx,{ type:'line', data:{ labels:${JSON.stringify(
        monthN.labels
      )}, datasets:[{label:'', data:${JSON.stringify(
    monthN.data
  )}, tension:.3, pointRadius:4 }] }, options:{ animation:false, plugins:{legend:{display:false}}, scales:{y:{beginAtZero:true}} } });
    })();</script>`;

  const weeklyCard = `
    <div class="card">
      <div class="card-title">НИЙТ БҮРТГЭЛ /Сүүлийн ${
        weeks.labels.length
      } долоо хоногоор/</div>
      <div class="grid">
        <canvas id="arbWeeks"></canvas>
      </div>
    </div>
    <script>(function(){
      ${dataLbl}
      const ctx=document.getElementById('arbWeeks').getContext('2d');
      new Chart(ctx,{type:'bar', data:{labels:${JSON.stringify(
        weeks.labels
      )}, datasets:[
        {label:'Лавлагаа', data:${JSON.stringify(weeks.lav)}},
        {label:'Үйлчилгээ', data:${JSON.stringify(weeks.uil)}},
        {label:'Гомдол', data:${JSON.stringify(weeks.gom)}}
      ]}, options:{animation:false,plugins:{legend:{position:'bottom'}},scales:{y:{beginAtZero:true}}}, plugins:[dataLabel]});
    })();</script>`;

  const topTable = (title, rows) => `
    <div class="card">
      <div class="card-title">${title}</div>
      <table class="cmp">
        <thead><tr><th></th><th>${top.labels[0]}</th><th>${
    top.labels[1]
  }</th><th>%</th></tr></thead>
        <tbody>
          ${(rows || [])
            .map(
              (r) =>
                `<tr><td>${escapeHtml(r.name)}</td><td class="num">${
                  r.prev || 0
                }</td><td class="num">${r.curr || 0}</td><td class="num ${
                  r.delta >= 0 ? "up" : "down"
                }">${r.delta >= 0 ? "▲" : "▼"} ${(
                  Math.abs(r.delta) * 100
                ).toFixed(0)}%</td></tr>`
            )
            .join("")}
        </tbody>
      </table>
    </div>`;

  return `
  <section>
    <div class="grid">
      ${lineCard}
      ${weeklyCard}
    </div>
    <div class="grid grid-1" style="margin-top:8px">
      ${topTable("ТОП Лавлагаа", top.groups["Лавлагаа"])}
      ${topTable("ТОП Үйчилгээ", top.groups["Үйлчилгээ"])}
      ${topTable("ТОП Гомдол", top.groups["Гомдол"])}
    </div>
  </section>`;
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
      margin: { top: "16mm", right: "14mm", bottom: "16mm", left: "14mm" },
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
  const t = nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port,
    secure,
    auth:
      process.env.SMTP_USER && process.env.SMTP_PASS
        ? { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS }
        : undefined,
    pool: true,
    maxConnections: 1,
    requireTLS: port === 587,
    tls: { minVersion: "TLSv1.2" },
  });
  await t.verify();
  await t.sendMail({
    from: process.env.FROM_EMAIL,
    to: process.env.RECIPIENTS,
    subject,
    html: `<p>Сайн байна уу,</p><p>АРД БИТ 7 хоногийн тайлан хавсаргав.</p><p style="color:#666;font-size:12px">Автоматаар илгээв.</p>`,
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
  if (!CONFIG.CURR_FILE)
    throw new Error("Missing CURR_FILE (path to current Excel)");
  [CONFIG.CURR_FILE, CONFIG.CSS_FILE].forEach((p) => {
    if (!fs.existsSync(p)) throw new Error(`Missing file: ${p}`);
  });
  if (!fs.existsSync(CONFIG.OUT_DIR))
    fs.mkdirSync(CONFIG.OUT_DIR, { recursive: true });

  // PREV_FILE автоматаар
  let prevPath = CONFIG.PREV_FILE;
  if (!prevPath) {
    prevPath = autoFindPrevFile(CONFIG.CURR_FILE);
    if (prevPath) console.log(`[auto] PREV_FILE = ${prevPath}`);
  }
  if (!prevPath)
    throw new Error(
      "Prev файл олдсонгүй: PREV_FILE тохируулах эсвэл файлын нэрсээ 09.29-10.05 маягаар байлгаарай."
    );

  // Range (not shown but useful for future)
  const { start: currStart, end: currEnd } = detectRangeFromCurrFile(
    CONFIG.CURR_FILE
  );
  const { start: prevStart, end: prevEnd } = detectRangeFromCurrFile(prevPath);
  void currStart;
  void currEnd;
  void prevStart;
  void prevEnd;

  // ARB months (ASS_YEAR auto)
  const months = extractARB_MonthsLatestN(
    CONFIG.CURR_FILE,
    CONFIG.ASS_SHEET,
    CONFIG.ASS_YEAR,
    CONFIG.ASS_TAKE_LAST_N_MONTHS
  );
  const monthN = {
    labels: months.points.map((p) => p.label + ` (${months.year})`),
    data: months.points.map((p) => Number(p.value) || 0),
  };

  // ARB weeks or fallback from Osticket1 prev/curr
  let weeks = extractARB_Last4WeeksByCategory(
    CONFIG.CURR_FILE,
    CONFIG.ASS_SHEET
  );
  if (!weeks || !weeks.labels?.length) {
    weeks = lastWeeksFromPrevCurrFallback(
      prevPath,
      CONFIG.CURR_FILE,
      CONFIG.OST_SHEET,
      CONFIG.ASS_COMPANY || ""
    );
  }

  // Osticket1 TOP tables (prev vs curr) for selected company
  const top = buildTopFromTwoFiles(
    prevPath,
    CONFIG.CURR_FILE,
    CONFIG.OST_SHEET,
    CONFIG.ASS_COMPANY,
    10
  );

  // Totals & shares (current)
  const currCat = countByCategoryWithinFile(
    CONFIG.CURR_FILE,
    CONFIG.OST_SHEET,
    CONFIG.ASS_COMPANY
  );
  const prevCat = countByCategoryWithinFile(
    prevPath,
    CONFIG.OST_SHEET,
    CONFIG.ASS_COMPANY
  );
  const totalPrev =
    (prevCat["Лавлагаа"] || 0) +
    (prevCat["Үйлчилгээ"] || 0) +
    (prevCat["Гомдол"] || 0);
  const totalCurr =
    (currCat["Лавлагаа"] || 0) +
    (currCat["Үйлчилгээ"] || 0) +
    (currCat["Гомдол"] || 0);
  const deltaTot = totalPrev > 0 ? (totalCurr - totalPrev) / totalPrev : 0;
  const pctShare = {
    lav: totalCurr ? (currCat["Лавлагаа"] || 0) / totalCurr : 0,
    uil: totalCurr ? (currCat["Үйлчилгээ"] || 0) / totalCurr : 0,
    gom: totalCurr ? (currCat["Гомдол"] || 0) / totalCurr : 0,
  };

  // Caption from filenames (prev → curr)
  const caption = `${parseWeekFromFilename(prevPath)?.raw || ""} → ${
    parseWeekFromFilename(CONFIG.CURR_FILE)?.raw || ""
  }`;

  // Build HTML (ard-app style)
  const cover = renderAssCover({
    company: CONFIG.ASS_COMPANY || "АРД БИТ",
    periodText: caption,
  });
  const body = cover + renderBitLayout({ monthN, weeks, top });
  const html = wrapHtml(body);

  // PDF
  const monday = dayjs().tz(CONFIG.TIMEZONE).startOf("week").add(1, "day");
  const stamp = monday.format("YYYYMMDD");
  const pdfPath = path.join(CONFIG.OUT_DIR, `ard-bit-${stamp}.pdf`);
  await htmlToPdf(html, pdfPath);

  if (CONFIG.SAVE_HTML) {
    const htmlPath = path.join(
      CONFIG.OUT_DIR,
      `${CONFIG.HTML_NAME_PREFIX}-${stamp}.html`
    );
    fs.writeFileSync(htmlPath, html, "utf8");
    console.log(`[OK] HTML saved → ${htmlPath}`);
  }

  const subject = `${CONFIG.SUBJECT_PREFIX} ${
    CONFIG.REPORT_TITLE
  } — ${monday.format("YYYY-MM-DD")}`;
  await sendEmailWithPdf(pdfPath, subject);
  console.log(
    `[OK] Sent ${path.basename(pdfPath)} → ${
      process.env.RECIPIENTS || "(no recipients set)"
    }`
  );
}

// ────────────────────────────────────────────────────────────────
// Scheduler (Даваа 09:00)
// ────────────────────────────────────────────────────────────────
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
if (process.argv.includes("--once")) {
  runOnce().catch((e) => {
    console.error(e);
    process.exit(1);
  });
} else {
  startScheduler();
}
