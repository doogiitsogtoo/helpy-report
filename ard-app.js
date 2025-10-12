// ard-app.js — Osticket1 (prev & curr) + Social (week columns) → PDF + Email
// ------------------------------------------------------------------------
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
  SOCIAL_SHEET: process.env.SOCIAL_SHEET || "Social",

  OUT_DIR: process.env.OUT_DIR || "./out",
  REPORT_TITLE: process.env.REPORT_TITLE || "Ард Апп — 7 хоногийн тайлан",
  SUBJECT_PREFIX: process.env.SUBJECT_PREFIX || "[ArdApp Weekly]",

  EMAIL_ENABLED: String(process.env.EMAIL_ENABLED ?? "true") === "true",
  SCHED_ENABLED: String(process.env.SCHED_ENABLED ?? "true") === "true",

  CSS_FILE: process.env.CSS_FILE || "./css/template.css",
  COMPANY_FILTER: process.env.COMPANY_FILTER || "Ард Апп", // Osticket1: зөвхөн энэ компанийн мөр

  JS_LIBS: ["https://cdn.jsdelivr.net/npm/apexcharts"], // chart
};

// ────────────────────────────────────────────────────────────────
/* Helpers */
// ────────────────────────────────────────────────────────────────
const nnum = (v) => Number(String(v ?? "").replace(/[^\d.-]/g, "")) || 0;
const norm = (s) =>
  String(s || "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
const pad2 = (n) => String(n).padStart(2, "0");

function parseWeekFromFilename(p) {
  // ".../ARD 8.04-8.10.xlsx" → {m1,d1,m2,d2, raw:"8.04-8.10"}
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
    raw: `${m[1]}.${m[2]}-${m[3]}.${m[4]}`,
  };
}
function makeYmd(y, M, D) {
  return `${y}-${pad2(M)}-${pad2(D)}`;
}
function formatWeekLabelFromFile(filePath) {
  const wk = parseWeekFromFilename(filePath);
  if (!wk) return "";
  return `${pad2(wk.m1)}.${pad2(wk.d1)}-${pad2(wk.m2)}.${pad2(wk.d2)}`;
}
function parseExcelDate(v) {
  if (v == null || v === "") return null;
  const asNum = Number(v);
  if (Number.isFinite(asNum) && asNum > 20000) {
    const ms = (asNum - 25569) * 86400 * 1000; // Excel serial → JS
    const d = dayjs(new Date(ms));
    return d.isValid() ? d : null;
  }
  const d = dayjs(v);
  return d.isValid() ? d : null;
}
function getColIdx(headers, patterns) {
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || "");
    if (patterns.some((re) => re.test(h))) return i;
  }
  return -1;
}
function inferYearFromDates(series, fallbackYear = dayjs().year()) {
  const years = series.filter(Boolean).map((d) => dayjs(d).year());
  return years.length ? years[0] : fallbackYear;
}
function inferYearFromSheet(ws, dateColIndexes = []) {
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  const sample = [];
  for (let r = 1; r < Math.min(rows.length, 80); r++) {
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

// Social week-like labels
function parseWeekLabelCell(s) {
  const m = String(s || "").match(
    /(\d{1,2})[./-](\d{1,2}).*?[-–].*?(\d{1,2})[./-](\d{1,2})/
  );
  if (!m) return null;
  return { m1: +m[1], d1: +m[2], m2: +m[3], d2: +m[4], raw: m[0] };
}
function findWeekColumnsFuzzy(rows, target) {
  const cands = new Map(); // col → parsed
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const p = parseWeekLabelCell(row[c]);
      if (p) cands.set(c, p);
    }
  }
  if (!cands.size) return null;

  const tEnd = dayjs(makeYmd(target.year, target.m2, target.d2));
  const scored = [...cands.entries()].map(([col, p]) => {
    const end = dayjs(makeYmd(target.year, p.m2, p.d2));
    const start = dayjs(makeYmd(target.year, p.m1, p.d1));
    const score =
      Math.abs(end.diff(tEnd, "day")) * 2 +
      Math.abs(
        start.diff(dayjs(makeYmd(target.year, target.m1, target.d1)), "day")
      );
    return {
      col,
      label: p.raw,
      score,
      deltaEnd: Math.abs(end.diff(tEnd, "day")),
    };
  });

  const close = scored
    .filter((s) => s.deltaEnd <= 1)
    .sort((a, b) => a.score - b.score);
  if (close.length) return { currCol: close[0].col, currLabel: close[0].label };
  const rightMost = Math.max(...[...cands.keys()]);
  return { currCol: rightMost, currLabel: cands.get(rightMost).raw };
}
function findRowByNameAnywhere(rows, name) {
  const want = norm(name);
  for (const r of rows) {
    if (!r) continue;
    for (let c = 0; c < r.length; c++) if (norm(r[c]) === want) return r;
  }
  return null;
}

// ────────────────────────────────────────────────────────────────
/* EXTRACTORS */
// ────────────────────────────────────────────────────────────────

// Phone vs Social — Osticket1 (company filter + week from filename)
function extractPhoneSocialFromOsticket(
  currFile,
  sheetName,
  companyFilter = CONFIG.COMPANY_FILTER
) {
  const wb = xlsx.readFile(currFile, { cellDates: true });
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet not found: ${sheetName}`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  if (!rows.length)
    return { phone: 0, social: 0, year: dayjs().year(), weekLabelFromFile: "" };

  const headers = rows[0].map((x) => String(x || ""));
  const idx = {
    company: getColIdx(headers, [/Компани/i]),
    line: getColIdx(headers, [/Шугам/i]),
    created: getColIdx(headers, [
      /Үүссэн\s*огноо/i,
      /Нээсэн\s*огноо/i,
      /Created/i,
      /Open(ed)?\s*date/i,
    ]),
    closed: getColIdx(headers, [/Хаагдсан\s*огноо/i, /Closed/i]),
  };
  if (idx.company < 0 || idx.line < 0) {
    throw new Error(`[Osticket1] "Компани" эсвэл "Шугам" багана олдсонгүй.`);
  }
  const useDateCol = idx.created >= 0 ? idx.created : idx.closed;
  if (useDateCol < 0) {
    throw new Error(`[Osticket1] "Үүссэн/Хаагдсан огноо" багана олдсонгүй.`);
  }

  const wk = parseWeekFromFilename(currFile);
  const year = inferYearFromSheet(ws, [useDateCol]);
  const start = wk ? dayjs(makeYmd(year, wk.m1, wk.d1)).startOf("day") : null;
  const end = wk ? dayjs(makeYmd(year, wk.m2, wk.d2)).endOf("day") : null;

  const inRange = (d) => {
    if (!d || !d.isValid()) return false;
    if (!start || !end) return true;
    return (
      (d.isAfter(start) || d.isSame(start)) &&
      (d.isBefore(end) || d.isSame(end))
    );
  };

  let phone = 0,
    social = 0;
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const comp = String(row[idx.company] || "").trim();
    if (companyFilter && comp !== companyFilter) continue;

    const when = parseExcelDate(row[useDateCol]);
    if (!inRange(when)) continue;

    const rawLine = String(row[idx.line] || "").trim();
    if (!rawLine || rawLine === ".") continue;
    const L = rawLine.toLowerCase();
    if (L === "phone" || /утас/i.test(rawLine)) phone++;
    else if (L === "social" || /social/i.test(L)) social++;
  }

  return {
    phone,
    social,
    year,
    weekLabelFromFile: wk
      ? `${pad2(wk.m1)}.${pad2(wk.d1)}-${pad2(wk.m2)}.${pad2(wk.d2)}`
      : "",
  };
}

// Social → суваг чинь (fuzzy week col)
function extractSocialBasicFuzzy(currFile, sheetName) {
  const wb = xlsx.readFile(currFile, { cellDates: true });
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;

  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  const wk = parseWeekFromFilename(currFile);
  const year = dayjs().year();

  let currCol = null,
    prevCol = null,
    labelCurr = "",
    labelPrev = "";
  if (wk) {
    const fz = findWeekColumnsFuzzy(rows, { ...wk, year });
    if (fz) {
      currCol = fz.currCol;
      labelCurr = fz.currLabel;
      // өмнөх week-like багана (баруун→зүүн)
      for (let c = currCol - 1; c >= 0; c--) {
        if (
          parseWeekLabelCell(rows[0]?.[c]) ||
          rows.some((r) => parseWeekLabelCell((r || [])[c]))
        ) {
          prevCol = c;
          const anyRow = rows.find((r) => r && r[prevCol]);
          labelPrev = String(anyRow?.[prevCol] || "").trim() || "Өмнөх 7 хоног";
          break;
        }
      }
    }
  }
  if (currCol == null) return null;

  const KEYS = [
    "Chat",
    "Comment",
    "Telegram",
    "Instagram",
    "Email",
    "Other",
    "Total",
  ];
  const out = {
    labels: [labelPrev || "Өмнөх 7 хоног", labelCurr || "Одоогийн 7 хоног"],
    rows: {},
    totalPrev: 0,
    totalCurr: 0,
  };

  for (const k of KEYS) {
    const row = findRowByNameAnywhere(rows, k) || [];
    const p = prevCol != null ? nnum(row[prevCol]) : 0;
    const c = nnum(row[currCol]);
    out.rows[k] = { prev: p, curr: c };
  }
  out.totalPrev = out.rows.Total?.prev || 0;
  out.totalCurr = out.rows.Total?.curr || 0;
  return out;
}

// Чатбот лавлагаа/үйлчилгээ — Social дээрх нэрлэсэн мөрүүд (2 week col)
const LAVLAGAA_ROWS = [
  "Нууц код сэргээх",
  "Гүйлгээ пин код",
  "Данс цэнэглэх",
  "ТАН код",
  "И-мэйл солих заавар",
  "Баталгаажуулалт хийх заавар",
];
const UILCHILGEE_ROWS = ["Дугаар солих", "Идэвхгүй төлөв", "ҮЦ данс нээх"];

function pickNamedRowsFromSocialFuzzy(currFile, names, sheetName) {
  const wb = xlsx.readFile(currFile, { cellDates: true });
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;

  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  const wk = parseWeekFromFilename(currFile);
  const year = dayjs().year();

  const fz = wk ? findWeekColumnsFuzzy(rows, { ...wk, year }) : null;
  if (!fz) return null;

  const currCol = fz.currCol;
  let prevCol = null,
    labelPrev = "Өмнөх 7 хоног";
  for (let c = currCol - 1; c >= 0; c--) {
    if (
      parseWeekLabelCell(rows[0]?.[c]) ||
      rows.some((r) => parseWeekLabelCell((r || [])[c]))
    ) {
      prevCol = c;
      const anyRow = rows.find((r) => r && r[prevCol]);
      labelPrev = String(anyRow?.[prevCol] || "").trim() || labelPrev;
      break;
    }
  }
  const labelCurr = fz.currLabel;

  const items = names.map((nm) => {
    const r = findRowByNameAnywhere(rows, nm) || [];
    return {
      name: nm,
      prev: prevCol != null ? nnum(r[prevCol]) : 0,
      curr: nnum(r[currCol]),
    };
  });

  return { labels: [labelPrev, labelCurr], items };
}

// ────────────────────────────────────────────────────────────────
// TOP tables from two files (prev & curr) — dynamic by subcategory
// ────────────────────────────────────────────────────────────────
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
    if (idx.company < 0 || idx.category < 0 || idx.subcat < 0) {
      throw new Error(
        `[Osticket1] "Компани/Ангилал/Туслах ангилал" багануудаа шалгаарай (${file}).`
      );
    }
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
      if (start && end && (!when || when.isBefore(start) || when.isAfter(end)))
        continue;

      const cat = String(row[idx.category] || "").trim();
      const sub = String(row[idx.subcat] || "").trim();
      if (!sub) continue;
      if (!bag[cat]) continue; // keep only 3 major groups

      bag[cat].set(sub, (bag[cat].get(sub) || 0) + 1);
    }

    const label = wk
      ? `${pad2(wk.m1)}.${pad2(wk.d1)}-${pad2(wk.m2)}.${pad2(wk.d2)}`
      : "7 хоног";
    return { bag, label };
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
/* HTML (ApexCharts + template.css) */
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
        <th>${title}</th>
        <th>${labels[0]}</th>
        <th>${labels[1]}</th>
        <th>▲</th>
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
  const titleLav = `ТОП ЛАВЛАГАА`;
  const titleUil = `ТОП ҮЙЛЧИЛГЭЭ`;
  const titleGom = `ТОП ГОМДОЛ`;

  return `
  <section class="sheet">
    <div class="grid grid-2">
      ${renderTopTableBlock(titleLav, top.labels, top.groups["Лавлагаа"])}
      ${renderTopTableBlock(titleUil, top.labels, top.groups["Үйлчилгээ"])}
    </div>
    <div class="spacer16"></div>
    ${renderTopTableBlock(titleGom, top.labels, top.groups["Гомдол"])}
  </section>`;
}

function buildHtml(payload, cssText) {
  const libs = CONFIG.JS_LIBS.map((s) => `<script src="${s}"></script>`).join(
    "\n"
  );
  const total = (payload.donut.phone + payload.donut.social).toLocaleString();

  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1"/>
  <style>${cssText}</style>
  ${libs}
  <title>Ard App Report</title>
</head>
<body>
  <div class="header">
    <h1>Ард Апп — Харилцагчийн үйлчилгээний тайлан</h1>
    <div class="caption">${payload.weekCurr}</div>
  </div>

  <section class="sheet">
    <div class="grid grid-2">
      <div class="card">
        <div class="card-title">Бүртгэл /Сувгаар/</div>
        <div id="chart-burtgel-social" class="chart-wrap" style="height:300px"></div>
        <p class="kpi-note">2 сувгаар нийт <b>${total}</b> удаа хандсан.</p>
      </div>

      <div class="card">
        <div class="card-title">Чатбот лавлагаа</div>
        <div id="chart-chatbot1" class="chart-wrap" style="height:320px"></div>
      </div>
    </div>
  </section>

  <section class="sheet">
    <div class="card">
      <div class="card-title">Чатбот үйлчилгээ</div>
      <div id="chart-chatbot2" class="chart-wrap" style="height:340px"></div>
    </div>
  </section>

  ${renderAllTopTables(payload.top)}

  <div class="footer">Автоматаар бэлтгэсэн тайлан (Ard App)</div>

  <script>
    (function(){
      // Week labels for series legend (from filenames)
      const legendPrev = ${JSON.stringify(
        payload.weekSeries?.[0] || "Өмнөх 7 хоног"
      )};
      const legendCurr = ${JSON.stringify(
        payload.weekSeries?.[1] || "Одоогийн 7 хоног"
      )};

      // Donut (Phone vs Social)
      const phone = ${Number(payload.donut.phone)};
      const social = ${Number(payload.donut.social)};
      const donutTotal = phone + social;

      new ApexCharts(document.querySelector("#chart-burtgel-social"), {
        series: [phone, social],
        chart: { type: "donut", height: 300 },
        labels: ["Phone", "Social"],
        colors: ["#546E7A", "#7986CB"],
        dataLabels: { enabled: true, style: { fontSize: "14px", colors: ["#fff"] } },
        legend: { position: "top" },
        plotOptions: {
          pie: { donut: { size: "65%", labels: {
            show: true, name: { show:false }, value: { show:false },
            total: { show:true, label:"Нийт", formatter: ()=> donutTotal.toLocaleString() }
          }}}
        },
        tooltip: { y: { formatter: (v)=> v.toLocaleString() } }
      }).render();

      // Bar #1 — лавлагаа
      new ApexCharts(document.querySelector("#chart-chatbot1"), {
        series: [
          { name: legendPrev, data: ${JSON.stringify(payload.chart1.prev)} },
          { name: legendCurr, data: ${JSON.stringify(payload.chart1.curr)} }
        ],
        chart: { type:"bar", height: 320, toolbar:{show:false} },
        plotOptions: { bar: { columnWidth:"55%", borderRadius:5, borderRadiusApplication:"end" } },
        dataLabels: {
          enabled:true,
          style:{ colors:["#fff"] },
          background:{ enabled:true, foreColor:"#000", padding:4, borderRadius:4, borderWidth:1, borderColor:"#1E90FF", opacity:.9 }
        },
        xaxis: { categories: ${JSON.stringify(payload.chart1.categories)} },
        legend: { position: "top" },
        fill: { opacity: 1 }
      }).render();

      // Bar #2 — үйлчилгээ
      new ApexCharts(document.querySelector("#chart-chatbot2"), {
        series: [
          { name: legendPrev, data: ${JSON.stringify(payload.chart2.prev)} },
          { name: legendCurr, data: ${JSON.stringify(payload.chart2.curr)} }
        ],
        chart: { type:"bar", height: 340, toolbar:{show:false} },
        plotOptions: { bar: { columnWidth:"55%", borderRadius:5, borderRadiusApplication:"end" } },
        dataLabels: {
          enabled:true,
          style:{ colors:["#fff"] },
          background:{ enabled:true, foreColor:"#000", padding:4, borderRadius:4, borderWidth:1, borderColor:"#1E90FF", opacity:.9 }
        },
        xaxis: { categories: ${JSON.stringify(payload.chart2.categories)} },
        legend: { position: "top" },
        fill: { opacity: 1 }
      }).render();
    })();
  </script>
</body>
</html>`;
}

// ────────────────────────────────────────────────────────────────
/* PDF + EMAIL */
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
    logger: true,
    debug: true,
  });

  await transporter.verify();
  await transporter.sendMail({
    from: process.env.FROM_EMAIL,
    to: process.env.RECIPIENTS,
    subject,
    html: `<p>Сайн байна уу,</p><p>Ард Апп 7 хоногийн тайланг хавсаргав.</p><p style="color:#666;font-size:12px">Автоматаар илгээв.</p>`,
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
/* MAIN */
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

  // 1) Phone/Social (curr)
  const phoneSoc = extractPhoneSocialFromOsticket(
    CONFIG.CURR_FILE,
    CONFIG.APP_SHEET,
    CONFIG.COMPANY_FILTER
  );

  // 2) Social (curr) — fuzzy week labels
  const social = extractSocialBasicFuzzy(CONFIG.CURR_FILE, CONFIG.SOCIAL_SHEET);
  const weekPrev = social?.labels?.[0] || "Өмнөх 7 хоног";
  const weekCurr =
    social?.labels?.[1] || phoneSoc.weekLabelFromFile || "Одоогийн 7 хоног";

  // LEGEND labels (from filenames) → fixes the “7” label
  const legendPrev = formatWeekLabelFromFile(CONFIG.PREV_FILE) || weekPrev;
  const legendCurr = formatWeekLabelFromFile(CONFIG.CURR_FILE) || weekCurr;

  // 3) Чатбот лавлагаа/үйлчилгээ charts (curr vs prev col)
  const lav = pickNamedRowsFromSocialFuzzy(
    CONFIG.CURR_FILE,
    LAVLAGAA_ROWS,
    CONFIG.SOCIAL_SHEET
  );
  const uil = pickNamedRowsFromSocialFuzzy(
    CONFIG.CURR_FILE,
    UILCHILGEE_ROWS,
    CONFIG.SOCIAL_SHEET
  );

  const chart1 = {
    categories: (lav?.items || []).map((x) => x.name),
    prev: (lav?.items || []).map((x) => x.prev),
    curr: (lav?.items || []).map((x) => x.curr),
  };
  const chart2 = {
    categories: (uil?.items || []).map((x) => x.name),
    prev: (uil?.items || []).map((x) => x.prev),
    curr: (uil?.items || []).map((x) => x.curr),
  };

  // 4) TOP хүснэгтүүд — Osticket1 (prev vs curr) динамик
  const top = buildTopFromTwoFiles(
    CONFIG.PREV_FILE,
    CONFIG.CURR_FILE,
    CONFIG.APP_SHEET,
    CONFIG.COMPANY_FILTER,
    10
  );

  // 5) HTML → PDF
  const cssText = fs.readFileSync(CONFIG.CSS_FILE, "utf-8");
  const html = buildHtml(
    {
      weekPrev,
      weekCurr,
      weekSeries: [legendPrev, legendCurr], // <— legend names in charts
      donut: {
        phone: phoneSoc.phone,
        social: social?.totalCurr || phoneSoc.social,
      },
      chart1,
      chart2,
      top,
    },
    cssText
  );

  const monday = dayjs().tz(CONFIG.TIMEZONE).startOf("week").add(1, "day");
  const pdfName = `ardapp-weekly-${monday.format("YYYYMMDD")}.pdf`;
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
