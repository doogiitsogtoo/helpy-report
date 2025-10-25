// ard-security.js — Osticket1 (prev & curr) + ASC (4 months/weeks) + GOMDOL(Sheet1) → HTML → PDF → Email
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

  PREV_FILE: process.env.PREV_FILE || "./ARD 10.06-10.12.xlsx",
  CURR_FILE: process.env.CURR_FILE || "./ARD 10.13-10.19.xlsx",

  // Тусдаа гомдлын файл (Sheet1)
  GOMDOL_FILE: process.env.GOMDOL_FILE || "./gomdol-weekly.xlsx",
  GOMDOL_SHEET: process.env.GOMDOL_SHEET || "Sheet1",

  // Sheets
  APP_SHEET: process.env.APP_SHEET || "Osticket1", // raw tickets
  ASC_SHEET: process.env.ASC_SHEET || "ASC", // aggregated (months + weeks)
  COMPANY_FILTER: process.env.COMPANY_FILTER || "Ард Секюритиз",

  // PDF / Email
  OUT_DIR: process.env.OUT_DIR || "./out",
  CSS_FILE: process.env.CSS_FILE || "./css/template.css",
  REPORT_TITLE: process.env.REPORT_TITLE || "Ард Секюритиз — 7 хоногийн тайлан",
  SUBJECT_PREFIX: process.env.SUBJECT_PREFIX || "[ArdSecurity Weekly]",
  EMAIL_ENABLED: String(process.env.EMAIL_ENABLED ?? "true") === "true",
  SCHED_ENABLED: String(process.env.SCHED_ENABLED ?? "true") === "true",
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
const norm = (s) =>
  String(s || "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
const makeYmd = (y, M, D) => `${y}-${pad2(M)}-${pad2(D)}`;
const num = (n) => Number(n || 0).toLocaleString();
const pct = (x, d = 0) => `${((x || 0) * 100).toFixed(d)}%`;

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

/** ASC → “Нийт хандалт / Сүүлийн 4 сар”
 * Зөвхөн “Ард Секюритиз” үндсэн блок (Outbound-оос өмнө, “Нийт” мөрт хүрэхэд зогсоно)
 * Одоогийн сараас −3..одоогийн сар; бүгд 0 бол хүснэгт дээрх бодит сүүлийн 4 сар руу fallback.
 */
function month4FromASC_SMART(file, sheetName) {
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
    return m ? Math.max(1, Math.min(12, +m)) : null;
  };
  const isStopRow = (cell) => {
    const s = String(cell ?? "")
      .trim()
      .toLowerCase();
    return s === "нийт" || s === "niit" || s === "outbound";
  };

  // 1) Жилийн баганууд
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

  // 2) Сарын шошго байрлах багана (ихэвчлэн A)
  let monthCol = 0,
    best = -1;
  for (const cand of [0, 1, 2]) {
    let score = 0;
    for (let r = headerRow + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      if (isStopRow(row[0])) break; // зөвхөн дээд блок
      if (monthFromCell(row[cand])) score++;
    }
    if (score > best) {
      best = score;
      monthCol = cand;
    }
  }
  if (best <= 0) return { labels: [], data: [] };

  // 3) year→month→value (Outbound/Нийт-ээс өмнөх мөрүүд)
  const mm = new Map();
  for (const [y] of yearCols) mm.set(y, new Map());
  for (let r = headerRow + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    if (isStopRow(row[0])) break;
    const m = monthFromCell(row[monthCol]);
    if (!m) continue;
    for (const [y, col] of yearCols) {
      const v = nnum(row[col]);
      mm.get(y).set(m, v);
    }
  }

  // 4) Одоогоос 4 сар
  const now = dayjs().tz(CONFIG.TIMEZONE);
  const want = [];
  for (let i = 3; i >= 0; i--) {
    const t = now.subtract(i, "month");
    want.push({ y: t.year(), m: t.month() + 1 });
  }
  let labels = [],
    data = [];
  for (const { y, m } of want) {
    labels.push(`${m}сар`);
    data.push(Number(mm.get(y)?.get(m) ?? 0));
  }

  // 5) Бүгд 0 бол — sheet дэх бодит сүүлийн 4 сар
  if (data.every((v) => v === 0)) {
    const all = [];
    for (const [y, map] of mm)
      for (const [m, v] of map) all.push({ y, m, v: Number(v || 0) });
    all.sort((a, b) => a.y - b.y || a.m - b.m);
    const tail = all.slice(-4);
    if (tail.length) {
      labels = tail.map((x) => `${x.m}сар`);
      data = tail.map((x) => x.v);
    }
  }
  return { labels, data };
}

// ASC → 4 weeks by category (Lav/Uilch/Gomdol) if available
function last4WeeksByCategoryFromASC(file, sheetName) {
  const wb = xlsx.readFile(file, { cellDates: true });
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;

  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  if (!rows.length) return null;

  const weekCols = findWeekColumnsFuzzy(rows);
  if (!weekCols.length) return null;
  const last4 = weekCols.slice(-4);

  const findRow = (name) => {
    const want = norm(name);
    for (const r of rows)
      for (const cell of r || []) if (norm(cell) === want) return r;
    return [];
  };

  const wantRows = {
    Лавлагаа: findRow("Лавлагаа") || [],
    Үйлчилгээ: findRow("Үйлчилгээ") || [],
    Гомдол: findRow("Гомдол") || [],
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

// Fallback: 2 weeks only (prev/curr)
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
    _rawPrev: prev,
    _rawCurr: curr,
  };
}

// Гомдлын override: GOMDOL_FILE → Sheet1
function overrideGomdolFromSheet1(weeksPayload) {
  try {
    if (!fs.existsSync(CONFIG.GOMDOL_FILE)) return weeksPayload;
    const wb = xlsx.readFile(CONFIG.GOMDOL_FILE, { cellDates: true });
    const ws = wb.Sheets[CONFIG.GOMDOL_SHEET];
    if (!ws) return weeksPayload;

    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
    if (!rows.length) return weeksPayload;

    const weekCols = findWeekColumnsFuzzy(rows);
    if (!weekCols.length) return weeksPayload;

    const gRow =
      rows.find((r) => (r || []).some((c) => norm(c) === "гомдол")) || [];
    if (!gRow.length) return weeksPayload;

    const gomData = weeksPayload.labels.map((lbl) => {
      const match = weekCols.find((w) => String(w.label) === String(lbl));
      return match ? nnum(gRow[match.col]) : 0;
    });

    const newSeries = (weeksPayload.series || []).map((s) =>
      s.name === "Гомдол" ? { ...s, data: gomData } : s
    );
    return { ...weeksPayload, series: newSeries };
  } catch {
    return weeksPayload;
  }
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
      created: hdr.findIndex((h) =>
        /Үүссэн\s*огноо|Нээсэн\s*огноо|Created/i.test(h)
      ),
      closed: hdr.findIndex((h) => /Хаагдсан\s*огноо|Closed/i.test(h)),
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
// HTML (Chart.js + value-label plugin)
// ────────────────────────────────────────────────────────────────
function renderCover({ company, periodText }) {
  return `
  <section class="sheet">
    <div style="background:linear-gradient(135deg,#ff6a6a,#ffa46a); border-radius:12px; padding:20px 24px; display:flex; justify-content:space-between; align-items:center;">
      <div style="background:#fff;border-radius:12px;padding:14px 18px;display:inline-block">
        <div style="font-weight:800;font-size:24px;letter-spacing:.5px;color:#b91c1c">ARD</div>
        <div style="color:#777;margin-top:4px">Хүчтэй. Хамтдаа.</div>
      </div>
      <div style="color:#fff;text-align:right">
        <div style="font-size:28px;font-weight:800;line-height:1.1">${esc(
          company
        )}</div>
        <div style="opacity:.95;margin-top:6px">${esc(periodText || "")}</div>
      </div>
    </div>
  </section>`;
}

function renderTopTableBlock(title, labels, rows) {
  const l0 = esc(labels?.[0] ?? "");
  const l1 = esc(labels?.[1] ?? "");
  return `
  <div class="sheet">
    <div class="card-title">${esc(title)}</div>
    <div class="table-wrap">
      <table class="cmp">
        <thead>
          <tr>
            <th width="50%">${esc(title)}</th>
            <th>${l0}</th>
            <th>${l1}</th>
            <th>%</th>
          </tr>
        </thead>
        <tbody>
          ${(rows || [])
            .map((r) => {
              const up = (r.delta || 0) >= 0;
              return `<tr>
                <td>${esc(r.name || "")}</td>
                <td class="num">${num(r.prev)}</td>
                <td class="num">${num(r.curr)}</td>
                <td class="${up ? "up" : "down"}">${up ? "▲" : "▼"} ${pct(
                Math.abs(r.delta || 0),
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

function renderAllTopTables(top) {
  return `
  <section class="sheet">
    ${renderTopTableBlock("ТОП Лавлагаа", top.labels, top.groups["Лавлагаа"])}
    <div class="spacer16"></div>
    ${renderTopTableBlock("ТОП Үйлчилгээ", top.labels, top.groups["Үйлчилгээ"])}
    <div class="spacer16"></div>
    ${renderTopTableBlock("ТОП Гомдол", top.labels, top.groups["Гомдол"])}
  </section>`;
}

function renderLayout({
  month4,
  weeks4,
  mini: miniData,
  prevCat,
  currCat,
  top,
}) {
  // Value labels plugin (хоёр чартад ашиглана)
  const valueLabelPluginCode = `
    const ValueLabelPlugin = {
      id: 'value-labels',
      afterDatasetsDraw(chart){
        const { ctx, data } = chart;
        ctx.save();
        ctx.font = '11px system-ui,-apple-system,Segoe UI,Roboto,Arial';
        ctx.fillStyle = '#111';
        ctx.textAlign = 'center';
        (data.datasets||[]).forEach((ds,di)=>{
          const meta = chart.getDatasetMeta(di);
          (ds.data||[]).forEach((v,i)=>{
            const el = meta.data?.[i];
            if(!el || v==null) return;
            const p = el.tooltipPosition ? el.tooltipPosition() : {x:el.x,y:el.y};
            const y = chart.config.type==='bar' ? p.y-6 : p.y-8;
            ctx.fillText(String(v), p.x, y);
          });
        });
        ctx.restore();
      }
    };
    window.ValueLabelPlugin = ValueLabelPlugin;
  `;

  const lavPrev = prevCat["Лавлагаа"] || 0,
    lavCurr = currCat["Лавлагаа"] || 0;
  const gomPrev = prevCat["Гомдол"] || 0,
    gomCurr = currCat["Гомдол"] || 0;
  const uilPrev = prevCat["Үйлчилгээ"] || 0,
    uilCurr = currCat["Үйлчилгээ"] || 0;
  const totPrev = lavPrev + gomPrev + uilPrev;
  const totCurr = lavCurr + gomCurr + uilCurr;
  const totDelta = totPrev ? (totCurr - totPrev) / totPrev : 0;

  // 4 сарын Line
  const lineCard = `
  <div class="sheet">
    <div class="card-title">НИЙТ ХАНДАЛТ /Сүүлийн ${
      month4.labels.length
    } сараар/</div>
    <div class="chart-wrap" style="height:200px"><canvas id="secMonths"></canvas></div>
  </div>
  <script>(function(){
    ${valueLabelPluginCode}
    new Chart(document.getElementById('secMonths').getContext('2d'),{
      type:'line',
      data:{ labels:${JSON.stringify(
        month4.labels
      )}, datasets:[{ label:'', data:${JSON.stringify(
    month4.data
  )}, tension:.3, pointRadius:3 }]},
      options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{display:false}, 'value-labels':{} }, scales:{ y:{ beginAtZero:true } } },
      plugins:[window.ValueLabelPlugin]
    });
  })();</script>`;

  // 4 долоо хоногийн Bar
  const dataU =
    (weeks4.series || []).find((s) => s.name === "Үйлчилгээ")?.data || [];
  const dataG =
    (weeks4.series || []).find((s) => s.name === "Гомдол")?.data || [];
  const dataL =
    (weeks4.series || []).find((s) => s.name === "Лавлагаа")?.data || [];
  const weeksCard = `
  <div class="sheet">
    <div class="card-title">НИЙТ БҮРТГЭЛ /Сүүлийн ${
      weeks4.labels.length
    } долоо хоногоор/</div>
    <div class="chart-wrap" style="height:200px"><canvas id="secWeeks"></canvas></div>
    <div class="kpi-note">
      • Тайлант хугацаанд нийт <b>${num(totCurr)}</b>.
      • Өмнөхтэй харьцуулахад <b class="${totDelta >= 0 ? "up" : "down"}">${pct(
    Math.abs(totDelta),
    0
  )}</b> ${totDelta >= 0 ? "өссөн" : "буурсан"}.
    </div>
  </div>
  <script>(function(){
    new Chart(document.getElementById('secWeeks').getContext('2d'),{
      type:'bar',
      data:{ labels:${JSON.stringify(weeks4.labels)}, datasets:[
        {label:'Үйлчилгээ', data:${JSON.stringify(dataU)}},
        {label:'Гомдол',   data:${JSON.stringify(dataG)}},
        {label:'Лавлагаа', data:${JSON.stringify(dataL)}}
      ]},
      options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{position:'bottom'}, 'value-labels':{} }, scales:{ y:{ beginAtZero:true } } },
      plugins:[window.ValueLabelPlugin]
    });
  })();</script>`;

  // Mini charts
  const miniSection = `
  <div class="sheet">
    <div class="grid grid-2">
      <div>
        <div class="card-title">ЛАВЛАГАА – ${num(lavCurr)} (${pct(
    lavCurr / (totCurr || 1),
    0
  )})</div>
        <div class="chart-wrap" style="height:140px"><canvas id="miniLav"></canvas></div>
      </div>
      <div>
        <div class="card-title">ҮЙЛЧИЛГЭЭ – ${num(uilCurr)} (${pct(
    uilCurr / (totCurr || 1),
    0
  )})</div>
        <div class="chart-wrap" style="height:140px"><canvas id="miniUil"></canvas></div>
      </div>
    </div>
    <div class="sheet">
      <div class="card-title">ГОМДОЛ – ${num(gomCurr)} (${pct(
    gomCurr / (totCurr || 1),
    0
  )})</div>
      <div class="chart-wrap" style="height:140px"><canvas id="miniGom"></canvas></div>
    </div>
  </div>
  <script>(function(){
    const labels = ${JSON.stringify(miniData.labels || [])};
    const mk = (id, arr) => new Chart(document.getElementById(id).getContext('2d'),{
      type:'bar',
      data:{ labels, datasets:[{ label:'Prev/Curr', data:arr }] },
      options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{display:false}, 'value-labels':{} }, scales:{ y:{ beginAtZero:true } } },
      plugins:[window.ValueLabelPlugin]
    });
    mk('miniLav', ${JSON.stringify([
      miniData.lavlagaa?.prev || 0,
      miniData.lavlagaa?.curr || 0,
    ])});
    mk('miniUil', ${JSON.stringify([
      miniData.uilchilgee?.prev || 0,
      miniData.uilchilgee?.curr || 0,
    ])});
    mk('miniGom', ${JSON.stringify([
      miniData.gomdol?.prev || 0,
      miniData.gomdol?.curr || 0,
    ])});
  })();</script>`;

  return `${lineCard}${weeksCard}${miniSection}${renderAllTopTables(top)}`;
}

function wrapHtml(bodyHtml, coverHtml) {
  const css = fs.readFileSync(CONFIG.CSS_FILE, "utf-8");
  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
  <style>${css}</style>
  <title>Ard Security Report</title>
</head>
<body>
  <div class="header"></div>
  ${coverHtml}
  ${bodyHtml}
  <div class="footer">Автоматаар бэлтгэсэн тайлан (Ard Securities)</div>
</body>
</html>`;
}

// ────────────────────────────────────────────────────────────────
// PDF + EMAIL
// ────────────────────────────────────────────────────────────────
async function htmlToPdf(html, outPath) {
  const browser = await puppeteer.launch({
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
    defaultViewport: { width: 1400, height: 900, deviceScaleFactor: 2 },
  });
  try {
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });
    await page.waitForFunction(
      () => window.Chart && document.querySelectorAll("canvas").length >= 2,
      { timeout: 10000 }
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

  const top = buildTopFromTwoFiles(
    CONFIG.PREV_FILE,
    CONFIG.CURR_FILE,
    CONFIG.APP_SHEET,
    CONFIG.COMPANY_FILTER,
    10
  );

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

  // 4 months — SMART (зөв функц!!)
  let month4 = month4FromASC_SMART(CONFIG.CURR_FILE, CONFIG.ASC_SHEET) || {
    labels: [],
    data: [],
  };

  // 4 weeks — ASC > fallback
  let weeks4 =
    last4WeeksByCategoryFromASC(CONFIG.CURR_FILE, CONFIG.ASC_SHEET) ||
    lastWeeksFromPrevCurrFallback(
      CONFIG.PREV_FILE,
      CONFIG.CURR_FILE,
      CONFIG.APP_SHEET,
      CONFIG.COMPANY_FILTER
    );

  // Гомдлын дата override
  weeks4 = overrideGomdolFromSheet1(weeks4);

  // mini
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

  const cover = renderCover({
    company: "АРД СЕКЮРИТИЗ",
    periodText: `${parseWeekFromFilename(CONFIG.PREV_FILE)?.raw || ""} – ${
      parseWeekFromFilename(CONFIG.CURR_FILE)?.raw || ""
    }`,
  });

  const html = wrapHtml(
    renderLayout({ month4, weeks4, mini, prevCat, currCat, top }),
    cover
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
if (process.argv.includes("--once")) {
  runOnce().catch((e) => {
    console.error(e);
    process.exit(1);
  });
} else {
  startScheduler();
}
