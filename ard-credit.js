// ard-credit.js — АРД КРЕДИТ (Osticket1 + ADB/totaly) → HTML → PDF → Email
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
import { createRequire } from "module";

dayjs.extend(utc);
dayjs.extend(tz);
const require = createRequire(import.meta.url);

// ────────────────────────────────────────────────────────────────
// CONFIG
// ────────────────────────────────────────────────────────────────
const CONFIG = {
  TIMEZONE: process.env.TIMEZONE || "Asia/Ulaanbaatar",

  // Excel files
  PREV_FILE: process.env.PREV_FILE || "./ARD 10.06-10.12.xlsx",
  CURR_FILE: process.env.CURR_FILE || "./ARD 10.13-10.19.xlsx",
  GOMDOL_FILE: "./gomdol-weekly.xlsx",

  // Sheets
  OST_SHEET: process.env.OST_SHEET || "Osticket1",
  ADB_SHEET: process.env.ADB_SHEET || "ADB", // totaly эсвэл ADB
  TOTALY_SHEET: process.env.TOTALY_SHEET || "totaly",

  // Filters
  COMPANY: process.env.COMPANY || "Ард Кредит",

  // PDF / Email
  OUT_DIR: process.env.OUT_DIR || "./out",
  CSS_FILE: process.env.CSS_FILE || "./css/template.css",
  SAVE_HTML: String(process.env.SAVE_HTML ?? "true") === "true",
  HTML_NAME_PREFIX: process.env.HTML_NAME_PREFIX || "report",
  REPORT_TITLE: process.env.REPORT_TITLE || "Ард Кредит — 7 хоногийн тайлан",
  SUBJECT_PREFIX: process.env.SUBJECT_PREFIX || "[ArdCredit Weekly]",
  EMAIL_ENABLED: String(process.env.EMAIL_ENABLED ?? "false") === "true",
  SCHED_ENABLED: String(process.env.SCHED_ENABLED ?? "false") === "true",

  // Chart.js inline fallback (ENV зам эсвэл node_modules → inline)
  CHART_PATH: process.env.CHART_PATH || "",
};

// ────────────────────────────────────────────────────────────────
// Helpers
// ────────────────────────────────────────────────────────────────
const nnum = (v) => Number(String(v ?? "").replace(/[^\d.-]/g, "")) || 0;
const pad2 = (n) => String(n).padStart(2, "0");
const num = (n) => Number(n || 0).toLocaleString();
const pct = (v) => `${(v * 100).toFixed(0)}%`;
const norm = (s) =>
  String(s || "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();

function compMatch(cell, filter) {
  if (!filter) return true;
  const a = norm(cell),
    b = norm(filter);
  if (!a || !b) return false;
  return a === b || a.includes(b) || b.includes(a);
}

function escapeHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function parseWeekFromFilename(p) {
  const base = path.basename(p || "");
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
  const n = Number(v);
  if (Number.isFinite(n) && n > 20000) {
    const ms = (n - 25569) * 86400 * 1000;
    const d = dayjs(new Date(ms));
    return d.isValid() ? d : null;
  }
  const d = dayjs(v);
  return d.isValid() ? d : null;
}
function inferYearFromSheet(ws, idx = []) {
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  for (let r = 1; r < Math.min(rows.length, 120); r++) {
    for (const c of idx) {
      const d = parseExcelDate((rows[r] || [])[c]);
      if (d) return d.year();
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
// CHART.JS INLINE (локал → CDN fallback)
// ────────────────────────────────────────────────────────────────
function inlineChartJs() {
  try {
    if (CONFIG.CHART_PATH && fs.existsSync(CONFIG.CHART_PATH)) {
      return fs.readFileSync(CONFIG.CHART_PATH, "utf-8");
    }
  } catch {}
  try {
    const p = require.resolve("chart.js/dist/chart.umd.js");
    return fs.readFileSync(p, "utf-8");
  } catch {}
  return ""; // сүүлчийн арга — CDN дээр найдна
}

// ────────────────────────────────────────────────────────────────
// ADB/totaly extractor-ууд
// ────────────────────────────────────────────────────────────────
function findWeekColumnsFuzzy(rows) {
  const cols = [];
  const weekLike = (s) =>
    /(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)\s*[-–]\s*(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)/.test(
      String(s || "")
    );
  for (let c = 0; c < (rows[0] || []).length; c++) {
    for (let r = 0; r < rows.length; r++) {
      const v = (rows[r] || [])[c];
      if (v && weekLike(v)) {
        cols.push({ col: c, label: String(v).trim() });
        break;
      }
    }
  }
  cols.sort((a, b) => a.col - b.col);
  return cols;
}
function findRowByNameAnywhere(rows, name) {
  const want = norm(name);
  for (const r of rows) {
    const cells = (r || []).slice(0, 8);
    for (const v of cells) {
      const x = norm(v);
      if (!x) continue;
      if (
        x === want ||
        x.includes(want) ||
        x.startsWith(want) ||
        (want === "гомдол" && /(gom|gomdol)/i.test(x)) ||
        (want === "лавлагаа" && /(lavl|lavlagaa)/i.test(x)) ||
        (want === "үйлчилгээ" && /(uilch|uilchilgee)/i.test(x))
      )
        return r;
    }
  }
  return null;
}
function month4FromADB(file, sheetName, preferYear = dayjs().year()) {
  const wb = xlsx.readFile(file, { cellDates: true });
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;

  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  if (!rows.length) return null;

  const monthRegex = /^\s*(\d{1,2})\s*сар\s*$/i;
  let monthRowStart = -1,
    monthCol = -1;

  outer: for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const m = String(row[c] ?? "").match(monthRegex);
      if (m) {
        monthRowStart = r;
        monthCol = c;
        break outer;
      }
    }
  }
  if (monthRowStart < 0) return null;

  // жил нь толгой дээрх нэг баганад байна
  let yearCol = -1;
  for (let c = monthCol + 1; c < (rows[0] || []).length; c++) {
    for (let up = 1; up <= 4; up++) {
      const rr = monthRowStart - up;
      if (rr < 0) break;
      const cell = String((rows[rr] || [])[c] ?? "").trim();
      if (cell === String(preferYear)) {
        yearCol = c;
        break;
      }
    }
    if (yearCol >= 0) break;
  }
  if (yearCol < 0) yearCol = monthCol + 2;

  const byMonth = new Map();
  for (let r = monthRowStart; r < rows.length; r++) {
    const row = rows[r] || [];
    const label = String(row[monthCol] ?? "");
    const m = label.match(monthRegex);
    if (!m) break;
    byMonth.set(Number(m[1]), nnum(row[yearCol]));
  }
  if (!byMonth.size) return null;

  const nowM = dayjs().tz(CONFIG.TIMEZONE).month() + 1;
  const arr = [];
  for (let m = 1; m <= nowM; m++)
    if (byMonth.has(m)) arr.push([m, byMonth.get(m)]);
  const last4 = arr.slice(-4);
  return {
    labels: last4.map(([m]) => `${m}сар`),
    data: last4.map(([, v]) => v),
  };
}

function weeklyFromTotaly(file, sheetName = CONFIG.TOTALY_SHEET) {
  const wb = xlsx.readFile(file, { cellDates: true });
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  if (!rows.length) return null;

  const head = rows[0] || [];
  const weekCols = [];
  const weekLike = (s) =>
    /(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)\s*[-–]\s*(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)/.test(
      String(s || "")
    );
  for (let i = head.length - 1; i >= 0 && weekCols.length < 4; i--) {
    if (head[i] && weekLike(head[i]))
      weekCols.push({ col: i, label: String(head[i]).trim() });
  }
  weekCols.reverse();
  if (!weekCols.length) return null;

  const findRow = (re) => rows.find((r) => r && r[0] && re.test(String(r[0])));
  const toNum = (r, i) =>
    Number(String(r?.[i] ?? "").replace(/[^\d.-]/g, "")) || 0;

  const rLav = findRow(/^Лавлагаа$/i);
  const rUil = findRow(/^Үйлчилгээ$/i);
  const rGom = findRow(/^Гомдол$/i);
  if (!rLav || !rUil || !rGom) return null;

  return {
    labels: weekCols.map((w) => w.label),
    series: [
      { name: "Гомдол", data: weekCols.map((w) => toNum(rGom, w.col)) },
      { name: "Үйлчилгээ", data: weekCols.map((w) => toNum(rUil, w.col)) },
      { name: "Лавлагаа", data: weekCols.map((w) => toNum(rLav, w.col)) },
    ],
  };
}

// ────────────────────────────────────────────────────────────────
// Osticket1 counters + TOP
// ────────────────────────────────────────────────────────────────
function countByCategoryWithinFile(file, sheetName, companyFilter) {
  const wb = xlsx.readFile(file, { cellDates: true });
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`[Ost] Sheet not found: ${sheetName} (${file})`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  if (!rows.length) return { Лавлагаа: 0, Үйлчилгээ: 0, Гомдол: 0 };

  const hdr = rows[0].map((x) => String(x || ""));
  const idx = {
    company: getColIdx(hdr, [/Компани/i, /Company/i]),
    category: getColIdx(hdr, [/Ангилал/i, /Category/i]),
    created: getColIdx(hdr, [/Үүссэн\s*огноо/i, /Нээсэн\s*огноо/i, /Created/i]),
    closed: getColIdx(hdr, [/Хаагдсан\s*огноо/i, /Closed/i]),
  };
  const hasDate = idx.created >= 0 || idx.closed >= 0;
  const dateCol = idx.created >= 0 ? idx.created : idx.closed;

  const wk = parseWeekFromFilename(file);
  const year = inferYearFromSheet(ws, hasDate ? [dateCol] : []);
  const start = wk ? dayjs(makeYmd(year, wk.m1, wk.d1)).startOf("day") : null;
  const end = wk ? dayjs(makeYmd(year, wk.m2, wk.d2)).endOf("day") : null;

  const keep = new Set(["лавлагаа", "үйлчилгээ", "гомдол"]);
  const acc = { Лавлагаа: 0, Үйлчилгээ: 0, Гомдол: 0 };

  for (let r = 1; r < rows.length; r++) {
    const row = rows[r] || [];
    if (!compMatch(row[idx.company], companyFilter)) continue;

    if (hasDate && start && end) {
      const d = parseExcelDate(row[dateCol]);
      if (!inRangeInclusive(d, start, end)) continue;
    }

    const cat = norm(row[idx.category] || "");
    if (!keep.has(cat)) continue;
    if (cat === "лавлагаа") acc["Лавлагаа"]++;
    else if (cat === "үйлчилгээ") acc["Үйлчилгээ"]++;
    else if (cat === "гомдол") acc["Гомдол"]++;
  }
  return acc;
}

function buildTopFromTwoFiles(
  prevFile,
  currFile,
  sheetName,
  companyFilter,
  limit = 10
) {
  const read = (file) => {
    const wb = xlsx.readFile(file, { cellDates: true });
    const ws = wb.Sheets[sheetName];
    if (!ws) throw new Error(`[Ost] Sheet not found: ${sheetName} (${file})`);
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
    const hdr = rows[0].map((x) => String(x || ""));

    const idx = {
      company: getColIdx(hdr, [/Компани/i, /Company/i]),
      category: getColIdx(hdr, [/Ангилал/i, /Category/i]),
      subcat: getColIdx(hdr, [/Туслах\s*ангилал|Дэд\s*ангилал/i]),
      created: getColIdx(hdr, [
        /Үүссэн\s*огноо/i,
        /Нээсэн\s*огноо/i,
        /Created/i,
      ]),
      closed: getColIdx(hdr, [/Хаагдсан\s*огноо/i, /Closed/i]),
    };
    const hasDate = idx.created >= 0 || idx.closed >= 0;
    const dateCol = idx.created >= 0 ? idx.created : idx.closed;

    const wk = parseWeekFromFilename(file);
    const year = inferYearFromSheet(ws, hasDate ? [dateCol] : []);
    const start = wk ? dayjs(makeYmd(year, wk.m1, wk.d1)).startOf("day") : null;
    const end = wk ? dayjs(makeYmd(year, wk.m2, wk.d2)).endOf("day") : null;

    const bag = {
      Лавлагаа: new Map(),
      Үйлчилгээ: new Map(),
      Гомдол: new Map(),
    };

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r] || [];
      if (!compMatch(row[idx.company], companyFilter)) continue;

      if (hasDate && start && end) {
        const d = parseExcelDate(row[dateCol]);
        if (!inRangeInclusive(d, start, end)) continue;
      }

      const cat = String(row[idx.category] || "").trim();
      const sub = String(row[idx.subcat] || "").trim();
      if (!sub || !bag[cat]) continue;
      bag[cat].set(sub, (bag[cat].get(sub) || 0) + 1);
    }
    return { bag, label: wk?.raw || "7 хоног" };
  };

  const prev = read(prevFile);
  const curr = read(currFile);

  const merge = (cat) => {
    const a = prev.bag[cat] || new Map();
    const b = curr.bag[cat] || new Map();
    const names = new Set([...a.keys(), ...b.keys()]);
    const rows = [...names].map((n) => {
      const p = a.get(n) || 0,
        c = b.get(n) || 0;
      const base = p > 0 ? p : c > 0 ? c : 1;
      return { name: n, prev: p, curr: c, delta: (c - p) / base };
    });
    rows.sort((x, y) => y.curr - x.curr || y.prev - x.prev);
    return rows.slice(0, limit);
  };

  return {
    labels: [prev.label, curr.label],
    lav: merge("Лавлагаа"),
    uil: merge("Үйлчилгээ"),
    gom: merge("Гомдол"),
  };
}

// ────────────────────────────────────────────────────────────────
// RENDER (жишээтэй ижил layout)
// ────────────────────────────────────────────────────────────────
function renderCover({ company, periodText }) {
  return `
  <section class="hero" style="margin-bottom:16px">
    <div style="background:linear-gradient(135deg,#ef4444,#f97316);
                border-radius:12px;padding:28px;display:flex;
                justify-content:space-between;align-items:center;min-height:220px;">
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

function renderLayout({ month4, weeks4, top10, prevCat, currCat, weekLabels }) {
  const totalCurr =
    (currCat["Лавлагаа"] || 0) +
    (currCat["Үйлчилгээ"] || 0) +
    (currCat["Гомдол"] || 0);
  const totalPrev =
    (prevCat["Лавлагаа"] || 0) +
    (prevCat["Үйлчилгээ"] || 0) +
    (prevCat["Гомдол"] || 0);
  const delta = totalPrev ? (totalCurr - totalPrev) / totalPrev : 0;

  const dataLbl = `
    // тоог давхцахгүй зурах + stacked bar дээр НИЙТ дүнг гаргах
    const dataLabelPlus = {
      id:'dataLabelPlus',
      afterDatasetsDraw(chart){
        const {ctx, data, scales, config} = chart;
        const y = scales?.y, sets = data?.datasets||[], labels = data?.labels||[];
        const isBar = config.type==='bar';
        const isStacked = !!(chart.options?.scales?.x?.stacked && chart.options?.scales?.y?.stacked);

        ctx.save();
        ctx.font='bold 12px system-ui,-apple-system,Segoe UI,Roboto,Arial';
        ctx.textAlign='center'; ctx.lineWidth=3;
        ctx.strokeStyle='rgba(255,255,255,.96)'; ctx.fillStyle='#111';

        // dataset-н өөрийн тоо (намхан сегментүүдийг нууж давхцахаас сэргийлнэ)
        sets.forEach((ds,di)=>{
          const meta = chart.getDatasetMeta(di);
          (ds.data||[]).forEach((v,i)=>{
            if(v==null) return;
            const el = meta.data?.[i]; if(!el) return;
            if(isBar){ const h=Math.abs((el.base??el.y)-el.y); if(h<18) return; }
            const pos = el.tooltipPosition ? el.tooltipPosition() : el;
            const tx = Number(v).toLocaleString('mn-MN');
            const x = pos.x;
            const yText = isBar ? (el.y + (el.base??el.y))/2 : (pos.y-6);
            ctx.strokeText(tx,x,yText); ctx.fillText(tx,x,yText);
          });
        });

        // stacked bar → НИЙТ дүнг дээд талд нь зурах (бусад шошготой мөргөлдөхгүй)
        if(isBar && isStacked && y && typeof y.getPixelForValue==='function'){
          const totals = labels.map((_,i)=>sets.reduce((s,ds)=>s+(+ds.data?.[i]||0),0));
          const metas = sets.map((_,di)=>chart.getDatasetMeta(di));
          totals.forEach((tot,i)=>{
            const base = metas[0]?.data?.[i]; if(!base) return;
            const x = (base.tooltipPosition?base.tooltipPosition():base).x;
            const yTop = y.getPixelForValue(tot);
            let minSegLabelY = Infinity;
            metas.forEach(m=>{
              const el = m.data?.[i]; if(!el) return;
              const h = Math.abs((el.base??el.y)-el.y);
              if(h>=18){ const yc=(el.y+(el.base??el.y))/2; minSegLabelY=Math.min(minSegLabelY,yc); }
            });
            let yText = yTop - 10;
            if (yText > (minSegLabelY - 12)) yText = minSegLabelY - 12;
            const tx = Number(tot).toLocaleString('mn-MN');
            ctx.strokeText(tx, x, yText); ctx.fillText(tx, x, yText);
          });
        }
        ctx.restore();
      }
    };
  `;
  const chartBundle = inlineChartJs();

  const weeklyCard = `
  <div class="card">
    <div class="card-title">НИЙТ БҮРТГЭЛ /Сүүлийн 4 долоо хоногоор/</div>
    <div class="grid">
      <canvas id="weeklyCat"></canvas>
      <ul style="margin:10px 0 0 18px;line-height:1.6">
        <li>Тайлант 7 хоногт нийт <b>${num(totalCurr)}</b>
        (${pct((currCat["Лавлагаа"] || 0) / Math.max(1, totalCurr))} лавлагаа,
         ${pct((currCat["Үйлчилгээ"] || 0) / Math.max(1, totalCurr))} үйлчилгээ,
         ${pct((currCat["Гомдол"] || 0) / Math.max(1, totalCurr))} гомдол).</li>
        <li>Өмнөх 7 хоногоос <b>${
          delta >= 0 ? "өссөн" : "буурсан"
        }</b>: <b>${pct(Math.abs(delta))}</b>.</li>
      </ul>
    </div>
  </div>`;

  const mini = (title, prev, curr) => {
    const d = prev ? (curr - prev) / prev : 0;
    const w = Math.min(100, Math.round((curr / Math.max(curr, prev, 1)) * 100));
    return `
      <div class="card soft" style="padding:10px 14px">
        <div style="font-weight:600;margin-bottom:4px">${title}</div>
        <div style="display:flex;align-items:center;gap:10px">
          <div style="flex:1;height:10px;background:#eee;border-radius:999px;overflow:hidden">
            <div style="width:${w}%;height:100%;background:#3b82f6"></div>
          </div>
          <div style="min-width:120px">${num(prev)} → <b>${num(curr)}</b> (${
      d >= 0 ? "+" : ""
    }${(d * 100).toFixed(0)}%)</div>
        </div>
      </div>`;
  };

  const topTable = (title, rows, labels) => `
    <div class="card">
      <div class="card-title">${title}</div>
      <table class="cmp">
        <thead><tr><th></th><th>${labels[0]}</th><th>${
    labels[1]
  }</th><th>%</th></tr></thead>
        <tbody>
          ${(rows || [])
            .map(
              (r) => `
            <tr>
              <td>${escapeHtml(r.name)}</td>
              <td class="num">${num(r.prev || 0)}</td>
              <td class="num">${num(r.curr || 0)}</td>
              <td class="num ${r.delta >= 0 ? "up" : "down"}">${
                r.delta >= 0 ? "▲" : "▼"
              } ${(Math.abs(r.delta) * 100).toFixed(0)}%</td>
            </tr>`
            )
            .join("")}
        </tbody>
      </table>
    </div>`;

  return `
  <section>
    <div class="grid">
      <div class="card" style="height: 500px; margin-bottom: 4rem;">
        <div class="card-title">НИЙТ ХАНДАЛТ /Сүүлийн ${
          month4.labels.length
        } сараар/</div>
        <canvas id="apsLine"></canvas>
      </div>
      ${weeklyCard}
    </div>

    <div class="grid grid-3" style="margin-top:8px">
      ${mini("ЛАВЛАГАА", prevCat["Лавлагаа"] || 0, currCat["Лавлагаа"] || 0)}
      ${mini("ҮЙЛЧИЛГЭЭ", prevCat["Үйлчилгээ"] || 0, currCat["Үйлчилгээ"] || 0)}
      ${mini("ГОМДОЛ", prevCat["Гомдол"] || 0, currCat["Гомдол"] || 0)}
    </div>

    <div class="grid grid-1" style="margin-top:8px">
      ${topTable("ТОП Лавлагаа", top10.lav, weekLabels)}
      ${topTable("ТОП Үйлчилгээ", top10.uil, weekLabels)}
      ${topTable("ТОП Гомдол", top10.gom, weekLabels)}
    </div>
  </section>

  ${
    chartBundle
      ? `<script>${chartBundle}</script>`
      : `<script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>`
  }
  <script>(function(){
    ${dataLbl}
    // Line (months)
    (function(){
      const ctx=document.getElementById('apsLine').getContext('2d');
     new Chart(ctx,{
  type:'line',
  data:{ labels: ${JSON.stringify(month4.labels)},
         datasets:[{label:'', data:${JSON.stringify(
           month4.data
         )}, tension:.3, pointRadius:4 }]},
  options:{ animation:false, plugins:{legend:{display:false}}, scales:{ y:{ beginAtZero:true } } },
  plugins:[dataLabelPlus]   // ← энд солино
});

    })();

    // Weekly stacked bar
    (function(){
      const ctx=document.getElementById('weeklyCat').getContext('2d');
     new Chart(ctx,{
  type:'bar',
  data:{ labels:${JSON.stringify(weeks4.labels)},
         datasets:${JSON.stringify(weeks4.series)} },
  options:{
    animation:false,
    plugins:{ legend:{ position:'bottom' } },
    scales:{
      x:{ stacked:true, categoryPercentage:0.6, barPercentage:0.8 },
      y:{ stacked:true, beginAtZero:true, ticks:{ maxTicksLimit:6 } }
    },
    responsive:true, maintainAspectRatio:false
  },
  plugins:[dataLabelPlus]   // ← энд солино
});

    })();
  })();</script>`;
}

function wrapHtml(bodyHtml) {
  const css = fs.readFileSync(CONFIG.CSS_FILE, "utf-8");
  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css"
        rel="stylesheet">
  <style>${css}
  /* …таны template.css-ийн дараа… */
  #apsLine   { width:100% !important; height:340px !important; }
  #weeklyCat { width:100% !important; height:300px !important; }
  </style>
</head>
<body>
 <div class="container py-3">
    <div class="row g-3">
      ${bodyHtml}
      <div class="footer">Автоматаар бэлтгэсэн тайлан (Ard Credit)</div>
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
  const transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port: Number(process.env.SMTP_PORT || 587),
    secure: String(process.env.SMTP_SECURE || "false") === "true",
    auth:
      process.env.SMTP_USER && process.env.SMTP_PASS
        ? { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS }
        : undefined,
    pool: true,
    maxConnections: 1,
    connectionTimeout: 20000,
    greetingTimeout: 15000,
    socketTimeout: 30000,
    requireTLS: String(process.env.SMTP_PORT) === "587",
    tls: { minVersion: "TLSv1.2" },
  });

  await transporter.verify();
  await transporter.sendMail({
    from: process.env.FROM_EMAIL,
    to: process.env.RECIPIENTS,
    subject,
    html: `<p>Сайн байна уу,</p><p>Ард Кредит 7 хоногийн тайланг хавсаргав.</p><p style="color:#666;font-size:12px">Автоматаар илгээв.</p>`,
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
  [CONFIG.CURR_FILE, CONFIG.PREV_FILE, CONFIG.CSS_FILE].forEach((p) => {
    if (!fs.existsSync(p)) throw new Error(`Missing file: ${p}`);
  });
  if (!fs.existsSync(CONFIG.OUT_DIR))
    fs.mkdirSync(CONFIG.OUT_DIR, { recursive: true });

  const month4 = month4FromADB(CONFIG.CURR_FILE, CONFIG.ADB_SHEET) || {
    labels: [],
    data: [],
  };

  const weeks4 = weeklyFromTotaly(CONFIG.CURR_FILE, CONFIG.TOTALY_SHEET) || {
    labels: [],
    series: [],
  };

  const prevCat = countByCategoryWithinFile(
    CONFIG.PREV_FILE,
    CONFIG.OST_SHEET,
    CONFIG.COMPANY
  );
  const currCat = countByCategoryWithinFile(
    CONFIG.CURR_FILE,
    CONFIG.OST_SHEET,
    CONFIG.COMPANY
  );

  const top10 = buildTopFromTwoFiles(
    CONFIG.PREV_FILE,
    CONFIG.CURR_FILE,
    CONFIG.OST_SHEET,
    CONFIG.COMPANY,
    10
  );

  const weekLabels = [
    parseWeekFromFilename(CONFIG.PREV_FILE)?.raw || "Өмнөх 7 хоног",
    parseWeekFromFilename(CONFIG.CURR_FILE)?.raw || "Одоогийн 7 хоног",
  ];

  const cover = renderCover({
    company: "АРД КРЕДИТ",
    periodText: `${weekLabels[0]} → ${weekLabels[1]}`,
  });

  const body =
    cover +
    renderLayout({ month4, weeks4, top10, prevCat, currCat, weekLabels });

  const html = wrapHtml(body);

  const monday = dayjs().tz(CONFIG.TIMEZONE).startOf("week").add(1, "day");
  const stamp = monday.format("YYYYMMDD");
  const pdfPath = path.join(CONFIG.OUT_DIR, `ard-credit-${stamp}.pdf`);

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

// Scheduler (Даваа 09:00)
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
