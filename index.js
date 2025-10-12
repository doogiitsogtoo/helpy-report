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

const CONFIG = {
  TOTALY_SHEET: "totaly",
  OST_SHEET: "Osticket1",

  EXCEL_FILE_CURR: "./ARD 8.04-8.10.xlsx",
  EXCEL_FILE_PREV: "./ARD 7.28-8.03.xlsx",

  COMPANIES: [
    "Ард Актив",
    "Ард Апп",
    "Ард Бит",
    "Ард Кредит",
    "Ард Лизинг",
    "Ард МПЕРС",
    "Ард Санхүүгийн Нэгдэл",
    "Ард Секюритиз",
    "Ард Экс",
    "Ардын Тэтгэврийн Данс",
  ],

  SHOW_TOP_CHARTS: false,

  GOMDOL_FILE: "./gomdol-weekly.xlsx",
  GOMDOL_SHEET: "Comp",

  GOMDOL_FORCE_WEEK_LABEL: "03.17-03.23",
  GOMDOL_FORCE_WEEK_COL: null,

  COMPANY_WEEK_LABEL: "07.28-08.03",

  COMPANY_WEEK_PICK_LOWER: true,
  // Сонголттой: яг 2 баганын нэрийг зааж өгч болно (баруун талд байгаа хэлбэрээр)
  // WEEK_LABELS: ['09/08 - 09/14', '09/15 - 09/21'],

  // Эсвэл индексээр (жиш: E=4, D=3) баруун→зүүн
  // WEEK_COL_INDEXES: [4, 3],
};

const {
  EXCEL_FILE = "./ARD 8.04-8.10.xlsx",
  REPORT_TITLE = "7 хоногийн тайлан",
  TIMEZONE = "Asia/Ulaanbaatar",
  RECIPIENTS = "",
  FROM_EMAIL = "doogiitsogtoo08@gmail.com",
  SUBJECT_PREFIX = "[Weekly]",
  SMTP_HOST,
  SMTP_PORT,
  SMTP_SECURE,
  SMTP_USER,
  SMTP_PASS,
} = process.env;

const KPI_ROWS = {
  total: /^Нийт\s*дуудлага/i,
  answered: /^Амжилттай\s*хариулсан\s*дуудлага/i,
  ivr: /(Автомат\s*хариулагч|ivr)/i,
  successRate:
    /(Амжилттай\s*холбогдсон\s*хув|Амжилттай\s*хариулсан\s*үзүүлэлт)/i,
  avgTalk: /(Ярьсан\s*дунд|АНГ)/i, // цагийн форматтай мөр
};

const OUTPUT_DIR = "./out";
const CSS_FILE = "./css/template.css";

const ROWS_MAP = {
  // Лавлах шугам → "Нийт дуудлага" (тоо)
  lavlahTotal: /^Нийт\s*дуудлага/i,
  // Цахим суваг → Social (тоо)
  social: /^Social/i,
  // Гадагшаа залгалт → Outbound (тоо)  | түүний амжилтын хувь → success outbound
  outbound: /^Outbound/i,
  outboundSR: /^success\s*outbound/i,
  // Салбар → Branch (тоо)
  branch: /^Branch/i,
  // Автомат хариулагч → ivr (тоо)
  ivr: /(Автомат\s*хариулагч|ivr)/i,
  // Гомдол → Гомдол (тоо)  (дээд хэсгийн мөрийг ашиглана)
  gomdol: /^Гомдол$/i,
  // Нийт амжилттай хувь (*AR)
  answeredRate:
    /(Амжилттай\s*холбогдсон\s*хув|Амжилттай\s*хариулсан\s*үзүүлэлт)/i,
};

// ────────────────────────────────────────────────────────────────
// Utils
// ────────────────────────────────────────────────────────────────
function escapeHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getSingleWeekLabelFromTotaly(wb, sheetName = CONFIG.TOTALY_SHEET) {
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  const header = rows[0] || [];
  // баруун талаас хоосон биш week шиг харагдах хамгийн сүүлийнхийг авна
  const weekLike = (s) =>
    /(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)\s*[-–]\s*(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)/.test(
      String(s || "")
    );
  for (let i = header.length - 1; i >= 0; i--) {
    const v = String(header[i] || "").trim();
    if (!v) continue;
    if (/өрчлөлт|эзлэх\s*хув/i.test(v)) continue;
    if (weekLike(v)) return v;
  }
  return null;
}

function extractOsticketTotalsFromTwoBooks(
  prevWb,
  currWb,
  sheetName = CONFIG.OST_SHEET
) {
  const norm = (s) => String(s || "").trim();
  const wantCompany = new Set(CONFIG.COMPANIES.map(norm));

  const tally = (wb) => {
    const ws = wb.Sheets[sheetName];
    if (!ws) throw new Error(`Sheet not found: ${sheetName}`);
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
    const hdr = rows[0].map((v) => norm(v));
    const idx = {
      company: hdr.findIndex((h) => /Компани/i.test(h)),
      category: hdr.findIndex((h) => /Ангилал/i.test(h)),
    };
    for (const k of Object.keys(idx))
      if (idx[k] < 0) throw new Error(`[Ost-totals] column not found: ${k}`);

    const total = { Лавлагаа: 0, Үйлчилгээ: 0 };
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;
      const company = norm(row[idx.company]);
      if (!wantCompany.has(company)) continue;
      const cat = norm(row[idx.category]);
      if (cat === "Лавлагаа" || cat === "Үйлчилгээ") total[cat]++;
    }
    return total;
  };

  const prev = tally(prevWb);
  const curr = tally(currWb);

  return {
    labels: [
      getSingleWeekLabelFromTotaly(prevWb) || "Өмнөх 7 хоног",
      getSingleWeekLabelFromTotaly(currWb) || "Одоогийн 7 хоног",
    ],
    lavlagaa: { prev: prev["Лавлагаа"], curr: curr["Лавлагаа"] },
    uilchilgee: { prev: prev["Үйлчилгээ"], curr: curr["Үйлчилгээ"] },
  };
}

function renderTotalsCompareBlock(title, pair, labels, canvasId) {
  const delta = pair.prev
    ? (pair.curr - pair.prev) / pair.prev
    : pair.curr
    ? 1
    : 0;
  const badge = `
    <div style="position:absolute;left:50%;top:50%;transform:translate(-50%,-50%);
                background:#ef4444;color:#fff;border-radius:8px;padding:2px 6px;font-weight:700">
      ${delta * 100 > 0 ? "+" : ""}${(delta * 100).toFixed(0)}%
    </div>`;

  return `
  <div class="card">
    <div class="card-title">${title} — ${pair.curr.toLocaleString()}</div>
    <div style="position:relative">
      ${badge}
      <canvas id="${canvasId}" height="180"></canvas>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
    <script>
      (function(){
        const ctx = document.getElementById('${canvasId}').getContext('2d');
        new Chart(ctx, {
          type:'bar',
          data:{
            labels: ${JSON.stringify(labels)},
            datasets:[{ label:'', data:[${pair.prev}, ${pair.curr}] }]
          },
          options:{
            animation:false,
            plugins:{ legend:{display:false}, tooltip:{enabled:false} },
            scales:{ y:{ beginAtZero:true } }
          }
        });
      })();
    </script>
  </div>`;
}

function extractOsticketTopBySubcatFromTwoBooks(
  prevWb,
  currWb,
  sheetName = CONFIG.OST_SHEET
) {
  const norm = (s) => String(s || "").trim();
  const wantCompany = new Set(CONFIG.COMPANIES.map(norm));

  const read = (wb) => {
    const ws = wb.Sheets[sheetName];
    if (!ws) throw new Error(`Sheet not found: ${sheetName}`);
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
    const hdr = rows[0].map((v) => norm(v));
    const idx = {
      company: hdr.findIndex((h) => /Компани/i.test(h)),
      category: hdr.findIndex((h) => /Ангилал/i.test(h)),
      subcat: hdr.findIndex((h) => /Туслах\s*ангилал/i.test(h)),
    };
    for (const k of Object.keys(idx))
      if (idx[k] < 0) throw new Error(`[Ost-2wb] column not found: ${k}`);

    const bag = { Лавлагаа: new Map(), Үйлчилгээ: new Map() };
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;
      const company = norm(row[idx.company]);
      if (!wantCompany.has(company)) continue;
      const cat = norm(row[idx.category]);
      if (cat !== "Лавлагаа" && cat !== "Үйлчилгээ") continue;
      const sub = norm(row[idx.subcat]);
      if (!sub) continue;
      bag[cat].set(sub, (bag[cat].get(sub) || 0) + 1);
    }
    return bag;
  };

  const prevBag = read(prevWb);
  const currBag = read(currWb);

  const toTop10 = (cat) => {
    const names = new Set([...prevBag[cat].keys(), ...currBag[cat].keys()]);
    const arr = [...names].map((name) => {
      const a = prevBag[cat].get(name) || 0;
      const b = currBag[cat].get(name) || 0;
      const deltaPct = a ? (b - a) / a : b ? 1 : 0;
      return { name, prev: a, curr: b, deltaPct };
    });
    arr.sort((x, y) => y.curr - x.curr);
    return arr.slice(0, 10);
  };

  return {
    labels: [
      getSingleWeekLabelFromTotaly(prevWb) || "Өмнөх 7 хоног",
      getSingleWeekLabelFromTotaly(currWb) || "Одоогийн 7 хоног",
    ],
    lavlagaaTop: toTop10("Лавлагаа"),
    uilchilgeeTop: toTop10("Үйлчилгээ"),
  };
}

// totaly sheet-ээс 2 долоо хоногийн KPI гаргаж авах
function extractTotalyKPI(wb, sheetName = CONFIG.TOTALY_SHEET) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet not found: ${sheetName}`);

  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  const header = rows[0] || [];
  const pos = getLastTwoDataCols(header);
  if (!pos) {
    console.warn(
      "[KPI] Could not resolve last 2 week columns. Header =",
      header
    );
    return null; // render талдаа guard байгаа
  }
  const { currCol, prevCol } = pos;

  const pickRow = (re) => rows.find((r) => r && r[0] && re.test(String(r[0])));
  const num = (x) => Number(String(x ?? "").replace(/[^\d.-]/g, "")) || 0;
  const pct = (x) => {
    const s = String(x ?? "").trim();
    return /^-?\d+(\.\d+)?%$/.test(s)
      ? Number(s.replace("%", "")) / 100
      : Number(s) || 0;
  };
  const fmtPct = (x) => `${(x * 100).toFixed(0)}%`;
  const growth = (a, b) => (a ? (b - a) / a : 0);

  const rTotal = pickRow(/^Нийт\s*дуудлага/i);
  const rAns = pickRow(/^Амжилттай\s*хариулсан\s*дуудлага/i);
  const rIVR = pickRow(/(Автомат\s*хариулагч|ivr)/i);
  const rSucc = pickRow(
    /(Амжилттай\s*холбогдсон\s*хув|Амжилттай\s*хариулсан\s*үзүүлэлт)/i
  );
  const rAvg = pickRow(/(Ярьсан\s*дунд|АНТ)/i);

  const labelPrev = String(header[prevCol] || "").trim();
  const labelCurr = String(header[currCol] || "").trim();

  const data = {
    labels: [labelPrev, labelCurr],
    total: [num(rTotal?.[prevCol]), num(rTotal?.[currCol])],
    answered: [num(rAns?.[prevCol]), num(rAns?.[currCol])],
    ivr: [num(rIVR?.[prevCol]), num(rIVR?.[currCol])],
    success: [pct(rSucc?.[prevCol]), pct(rSucc?.[currCol])],
    avgTalk: [String(rAvg?.[prevCol] ?? ""), String(rAvg?.[currCol] ?? "")],
  };

  return {
    data,
    deltas: {
      total: fmtPct(growth(data.total[0], data.total[1])),
      answered: fmtPct(growth(data.answered[0], data.answered[1])),
      ivr: fmtPct(growth(data.ivr[0], data.ivr[1])),
      success: fmtPct(growth(data.success[0], data.success[1])),
    },
  };
}
const fmt = (n) => Number(n || 0).toLocaleString("en-US");
const pct = (x, d = 0) => `${(x * 100).toFixed(d)}%`;

function makeSocialNarrative(soc, options = {}) {
  const safe = (n) => Math.max(0, Number(n) || 0);

  const totalCurr = safe(soc.totalCurr);
  const totalPrev = safe(soc.totalPrev);
  const chatCurr = safe(soc.rows?.Chat?.curr);

  // Сувгуудын нийлбэр (Total алдаатай үед ашиглана)
  const sumChannels = [
    "Chat",
    "Comment",
    "Telegram",
    "Instagram",
    "Email",
    "Other",
  ].reduce((s, k) => s + safe(soc.rows?.[k]?.curr), 0);

  // Нийт тоог зохистой болгох (fallback)
  let effTotal = totalCurr;
  if (effTotal <= 0 && sumChannels > 0) effTotal = sumChannels;
  if (effTotal > 0 && chatCurr > effTotal) effTotal = chatCurr; // clamp

  const otherCurr = Math.max(effTotal - chatCurr, 0);
  const chatShare = effTotal ? chatCurr / effTotal : 0;
  const otherShare = effTotal ? otherCurr / effTotal : 0;

  // Өөрчлөлт
  let deltaTxt = "";
  let deltaVal = 0;
  if (totalPrev > 0 && effTotal > 0) {
    deltaVal = (effTotal - totalPrev) / totalPrev;
    deltaTxt = ` (Өмнөх 7 хоногтой харьцуулахад ${
      deltaVal >= 0 ? "өссөн" : "буурсан"
    }, ${pct(Math.abs(deltaVal), 0)}.)`;
  }

  // 1-р өгүүлбэр
  let line1 = "";
  if (effTotal > 0) {
    line1 =
      `Цахим сувгаар ${fmt(effTotal)} харилцагчид үйлчилсэн. ` +
      `ЧАТБОТ (автомат хариулагч) нийт хандалтын ${fmt(chatCurr)} (${pct(
        chatShare,
        0
      )})-д хариулсан бол ` +
      `бусад сувгаар ${fmt(otherCurr)} (${pct(
        otherShare,
        0
      )}) харилцагчид лавлагаа/үйлчилгээ авсан.` +
      deltaTxt;
  } else {
    line1 = `Энэ долоо хоногт цахим сувгаар үйлчилгээ үзүүлсэн бүртгэл алга.`;
  }

  // 2-р өгүүлбэр (сонголттой)
  let line2 = "";
  if (options.chatDup != null && options.chatSucc != null) {
    const dup = safe(options.chatDup);
    const succ = safe(options.chatSucc);
    const succRate = dup ? succ / dup : 0;
    line2 = ` Чатботоор давхардсан тоогоор ${fmt(dup)} хүсэлт ирснээс ${fmt(
      succ
    )} үйлчилгээ (${pct(
      succRate,
      0
    )}) нь шаардлага хангаж амжилттай үйлчилгээ үзүүлсэн.`;
  }

  return `<p class="kpi-note">${line1}${line2}</p>`;
}
// HTML хэсэг (зурагтай адил блок)
function renderTotalyKPISection(kpi) {
  if (
    !kpi ||
    !kpi.data ||
    !Array.isArray(kpi.data.labels) ||
    kpi.data.labels.length < 2
  ) {
    console.warn("[KPI] missing data for totaly sheet");
    return `
      <section class="kpi">
        <div class="card">
          <div class="card-title">Дуудлагын үзүүлэлт</div>
          <p style="color:#a33">KPI өгөгдөл олдсонгүй (totaly sheet). Header-ийн сүүлийн 2 долоо хоног, 
          мөрийн нэрс (Нийт дуудлага, Амжилттай хариулсан дуудлага, ivr, ...)-ээ шалгана уу.</p>
        </div>
      </section>`;
  }

  const { data, deltas } = kpi;
  // Chart.js-ийг CDN-ээс татаж ашиглана (Puppeteer PDF рүү render хийнэ)
  return `
  <section class="kpi">
    <div class="grid grid-2">
      <div class="card">
        <div class="card-title">Дуудлагын үзүүлэлт</div>
        <canvas id="kpiChart" width="900" height="420"></canvas>
        <div class="legend">
          <span class="dot dot-total"></span> Нийт дуудлага
          <span class="dot dot-ans"></span> Амжилттай холбогдсон тоо
          <span class="dot dot-ivr"></span> IVR-т хандсан
          <span class="dot dot-line"></span> Амжилттай холбогдсон хувь
        </div>
      </div>
      <div class="card">
        <div class="card-title">Өмнөх 7 хоногтой харьцуулахад</div>
        <table class="cmp">
          <thead>
            <tr><th></th><th>${data.labels[0]}</th><th>${
    data.labels[1]
  }</th><th>Хувь</th></tr>
          </thead>
          <tbody>
            <tr><td>Нийт дуудлага</td>
                <td>${data.total[0].toLocaleString()}</td>
                <td>${data.total[1].toLocaleString()}</td>
                <td class="${deltas.total.startsWith("-") ? "down" : "up"}">${
    deltas.total
  }</td></tr>
            <tr><td>Амжилттай хариулсан дуудлага</td>
                <td>${data.answered[0].toLocaleString()}</td>
                <td>${data.answered[1].toLocaleString()}</td>
                <td class="${
                  deltas.answered.startsWith("-") ? "down" : "up"
                }">${deltas.answered}</td></tr>
            <tr><td>IVR-т хандсан дуудлага</td>
                <td>${data.ivr[0].toLocaleString()}</td>
                <td>${data.ivr[1].toLocaleString()}</td>
                <td class="${deltas.ivr.startsWith("-") ? "down" : "up"}">${
    deltas.ivr
  }</td></tr>
            <tr><td>Ярьсан дундаж хугацаа</td>
                <td>${data.avgTalk[0]}</td>
                <td>${data.avgTalk[1]}</td>
                <td></td></tr>
          </tbody>
        </table>
      </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
    <script>
      (function(){
        const ctx = document.getElementById('kpiChart').getContext('2d');
        const chart = new Chart(ctx, {
          type: 'bar',
          data: {
            labels: ${JSON.stringify(data.labels)},
            datasets: [
              { label:'Нийт дуудлага', data:${JSON.stringify(
                data.total
              )}, yAxisID:'yBar' },
              { label:'Амжилттай хариулсан тоо', data:${JSON.stringify(
                data.answered
              )}, yAxisID:'yBar' },
              { label:'IVR-т хандсан', data:${JSON.stringify(
                data.ivr
              )}, yAxisID:'yBar' },
              { label:'Амжилттай холбогдсон хувь', type:'line', data:${JSON.stringify(
                data.success.map((v) => +(v * 100).toFixed(1))
              )}, yAxisID:'yLine' }
            ]
          },
          options: {
            animation:false,
            scales: {
              yBar: { position:'left', beginAtZero:true, title:{ display:true, text:'Дуудлагын тоо' } },
              yLine:{ position:'right', beginAtZero:true, max:100, ticks:{ callback:(v)=>v+'%' }, grid:{ drawOnChartArea:false }, title:{ display:true, text:'Хувь' } }
            },
            plugins:{
              legend:{ display:false },
              tooltip:{ callbacks:{ label:(ctx)=>{
                const lab = ctx.dataset.label || '';
                const v = ctx.parsed.y;
                return lab.includes('хувь') ? lab+': '+v+'%' : lab+': '+v.toLocaleString();
              }}}
            }
          }
        });
      })();
    </script>
  </section>
  `;
}

function parseNumber(x) {
  const n = Number(String(x ?? "").replace(/[^\d.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
}
function parsePercentTo01(x) {
  const s = String(x ?? "").trim();
  if (/^[-+]?\d+(\.\d+)?%$/.test(s)) return Number(s.replace("%", "")) / 100;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}

function findRow(rows, regex) {
  return rows.find((r) => r && r[0] && regex.test(String(r[0])));
}

function getLastTwoDataCols(header) {
  // 1) Хэрэв хэрэглэгч WEEK_COL_INDEXES өгсөн бол шууд хэрэглэ
  if (
    Array.isArray(CONFIG.WEEK_COL_INDEXES) &&
    CONFIG.WEEK_COL_INDEXES.length === 2
  ) {
    return {
      currCol: CONFIG.WEEK_COL_INDEXES[0],
      prevCol: CONFIG.WEEK_COL_INDEXES[1],
    };
  }

  // 2) Хэрэв WEEK_LABELS өгсөн бол нэрээр тааруулж ол
  if (Array.isArray(CONFIG.WEEK_LABELS) && CONFIG.WEEK_LABELS.length === 2) {
    const idx1 = header.findIndex(
      (h) => String(h || "").trim() === CONFIG.WEEK_LABELS[1]
    );
    const idx0 = header.findIndex(
      (h) => String(h || "").trim() === CONFIG.WEEK_LABELS[0]
    );
    if (idx1 >= 0 && idx0 >= 0) return { currCol: idx1, prevCol: idx0 };
  }

  // 3) Толгойгоос “07.14-07.20”, “09/08 - 09/14”, “09.08 – 09.14” гэх мэт range-тай багануудыг барина
  const weekLike = (s) => {
    const t = String(s || "")
      .replace(/\s+/g, " ")
      .trim();
    // 07.14-07.20 | 07.14 - 07.20 | 09/08 - 09/14 | 2025.07.14 - 2025.07.20
    return /(\d{1,4}[./-]\d{1,2}[./-]\d{1,2}|\d{1,2}[./-]\d{1,2})\s*[-–]\s*(\d{1,4}[./-]\d{1,2}[./-]\d{1,2}|\d{1,2}[./-]\d{1,2})/.test(
      t
    );
  };

  // “Өөрчлөлт”, “Эзлэх хув(ь)” зэрэг туслах багануудыг алгас
  const ban = (s) => /өрчлөлт|эзлэх\s*хув/i.test(String(s || ""));

  const candidates = [];
  for (let i = header.length - 1; i >= 0; i--) {
    const h = header[i];
    if (!h) continue;
    if (ban(h)) continue;
    if (weekLike(h)) candidates.push(i);
    if (candidates.length === 2) break;
  }
  if (candidates.length === 2) {
    return { currCol: candidates[0], prevCol: candidates[1] }; // баруун→зүүн
  }

  // 4) Ядаж баруун талын 2 хоосон биш баганыг авъя (fallback)
  const nonEmpty = [];
  for (let i = header.length - 1; i >= 0; i--) {
    const v = String(header[i] || "").trim();
    if (v && !ban(v)) nonEmpty.push(i);
    if (nonEmpty.length === 2) break;
  }
  if (nonEmpty.length === 2) {
    return { currCol: nonEmpty[0], prevCol: nonEmpty[1] };
  }

  return null;
}

function extractChannelsFromTotaly(wb, sheetName = CONFIG.TOTALY_SHEET) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet not found: ${sheetName}`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  const header = rows[0] || [];
  const { currCol, prevCol } = getLastTwoDataCols(header);

  const r = {
    lavlahTotal: findRow(rows, ROWS_MAP.lavlahTotal),
    social: findRow(rows, ROWS_MAP.social),
    outbound: findRow(rows, ROWS_MAP.outbound),
    outboundSR: findRow(rows, ROWS_MAP.outboundSR),
    branch: findRow(rows, ROWS_MAP.branch),
    ivr: findRow(rows, ROWS_MAP.ivr),
    gomdol: findRow(rows, ROWS_MAP.gomdol),
    answeredRate: findRow(rows, ROWS_MAP.answeredRate),
  };

  const labelPrev = String(header[prevCol] || "").trim();
  const labelCurr = String(header[currCol] || "").trim();

  const val = (row, col, kind = "num") => {
    if (!row) return 0;
    return kind === "pct" ? parsePercentTo01(row[col]) : parseNumber(row[col]);
  };

  const data = {
    labels: [labelPrev, labelCurr],
    lavlah: {
      title: "Лавлах шугам",
      curr: val(r.lavlahTotal, currCol, "num"),
      prev: val(r.lavlahTotal, prevCol, "num"),
      // нийт амжилтын хувь (*AR)
      tagLabel: "*AR",
      tagCurr: val(r.answeredRate, currCol, "pct"), // 0..1
    },
    social: {
      title: "Цахим суваг",
      curr: val(r.social, currCol, "num"),
      prev: val(r.social, prevCol, "num"),
    },
    outbound: {
      title: "Гадагшаа залгалт",
      curr: val(r.outbound, currCol, "num"),
      prev: val(r.outbound, prevCol, "num"),
      tagLabel: "*SR",
      tagCurr: val(r.outboundSR, currCol, "pct"), // 0..1
    },
    branch: {
      title: "Салбар",
      curr: val(r.branch, currCol, "num"),
      prev: val(r.branch, prevCol, "num"),
    },
    ivr: {
      title: "Автомат хариулагч",
      curr: val(r.ivr, currCol, "num"),
      prev: val(r.ivr, prevCol, "num"),
    },
    gomdol: {
      title: "Гомдол",
      curr: val(r.gomdol, currCol, "num"),
      prev: val(r.gomdol, prevCol, "num"),
    },
  };

  // өсөлт/бууралтын % (prev→curr)
  for (const k of Object.keys(data)) {
    if (k === "labels") continue;
    const d = data[k];
    const delta = d.prev ? (d.curr - d.prev) / d.prev : 0;
    d.deltaPct = delta; // 0..1 (+/-)
  }

  return data;
}

// ───────── сувгийн картуудыг HTML-р буулгах ─────────
function renderChannelCards(ch) {
  const fmtInt = (n) => Number(n || 0).toLocaleString();
  const pct = (x, digits = 0) => `${(x * 100).toFixed(digits)}%`;
  const arrow = (v) => (v >= 0 ? "▲" : "▼");
  const updownCls = (v) => (v >= 0 ? "up" : "down");

  // Товч тайлбар (нэгдсэн)
  const totalVisitors =
    (ch.lavlah?.curr || 0) +
    (ch.social?.curr || 0) +
    (ch.outbound?.curr || 0) +
    (ch.branch?.curr || 0) +
    (ch.ivr?.curr || 0) +
    (ch.gomdol?.curr || 0);
  const totalDelta =
    (ch.lavlah.deltaPct +
      ch.social.deltaPct +
      ch.outbound.deltaPct +
      ch.branch.deltaPct +
      ch.ivr.deltaPct +
      ch.gomdol.deltaPct) /
    6;

  const topNote = `
    <div class="card-title">Нэгдсэн үзүүлэлт</div>
    <p class="kpi-note">
      Нийт ${fmtInt(totalVisitors)} харилцагч 6 төрлийн сувгаар хандсан.
      Нийт хандалт өмнөх 7 хоногтой харьцуулахад
      <span class="${updownCls(totalDelta)}">${arrow(totalDelta)} ${pct(
    Math.abs(totalDelta)
  )}</span>.
      ${
        ch.lavlah?.tagCurr != null
          ? `Амжилттай холболт ${pct(ch.lavlah.tagCurr)}.`
          : ""
      }
    </p>
  `;

  const card = (o, color) => `
    <div class="mini-card" style="--grad:${color}">
      <a class="mini-title">${o.title}</a>
      <div class="mini-value">${fmtInt(o.curr)}</div>
      ${
        o.tagLabel
          ? `<div class="mini-sub">*${o.tagLabel}: ${pct(o.tagCurr)}</div>`
          : ""
      }
      <div class="mini-delta ${updownCls(o.deltaPct)}">${arrow(
    o.deltaPct
  )} ${pct(Math.abs(o.deltaPct))}</div>
    </div>
  `;

  return `
  <section class="channel-cards">
    <div class="card">${topNote}</div>
    <div class="mini-grid">
      ${card(ch.lavlah, "linear-gradient(135deg,#3b82f6,#60a5fa)")}
      ${card(ch.social, "linear-gradient(135deg,#10b981,#34d399)")}
      ${card(ch.outbound, "linear-gradient(135deg,#f59e0b,#fbbf24)")}
      ${card(ch.branch, "linear-gradient(135deg,#ef4444,#fb7185)")}
      ${card(ch.ivr, "linear-gradient(135deg,#7c3aed,#8b5cf6)")}
      ${card(ch.gomdol, "linear-gradient(135deg,#ef4444,#fca5a5)")}
    </div>
  </section>
  `;
}

function renderOsticketTopBySubcatSection(top) {
  if (!top) return "";

  const arrow = (v) => (v >= 0 ? "▲" : "▼");
  const cls = (v) => (v >= 0 ? "up" : "down");
  const pct = (v) => `${Math.abs(v * 100).toFixed(0)}%`;

  const makeTable = (rows, labels) => `
    <table class="cmp">
      <thead>
        <tr>
          <th>Туслах ангилал</th>
          <th>${labels[0]}</th>
          <th>${labels[1]}</th>
          <th>Хувь</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map(
            (r) => `
          <tr>
            <td>${escapeHtml(r.name)}</td>
            <td class="num">${r.prev.toLocaleString()}</td>
            <td class="num">${r.curr.toLocaleString()}</td>
            <td class="num ${cls(r.deltaPct)}">${arrow(r.deltaPct)} ${pct(
              r.deltaPct
            )}</td>
          </tr>
        `
          )
          .join("")}
      </tbody>
    </table>`;

  const block = (title, rows, labels) => `
    <div class="card">
      <div class="card-title">${title}</div>
      ${makeTable(rows, labels)}
    </div>`;

  return `
  <section class="ost-top">
    ${block("ТОП ЛАВЛАГАА / туслах ангилалаар /", top.lavlagaaTop, top.labels)}
    ${block(
      "ТОП ҮЙЛЧИЛГЭЭ / туслах ангилалаар /",
      top.uilchilgeeTop,
      top.labels
    )}
  </section>`;
}

// totals + ostTop-ийг нэг grid дотор зэрэгцүүлж харуулна
function renderTotalsAndTopSection(totals, top) {
  if (!totals || !top) return "";

  const arrow = (v) => (v >= 0 ? "▲" : "▼");
  const cls = (v) => (v >= 0 ? "up" : "down");
  const pct = (v) => `${Math.abs(v * 100).toFixed(0)}%`;

  const makeTopTable = (rows, labels) => `
    <table class="cmp">
      <thead>
        <tr>
          <th>Туслах ангилал</th>
          <th>${labels[0]}</th>
          <th>${labels[1]}</th>
          <th>Хувь</th>
        </tr>
      </thead>
      <tbody>
        ${rows
          .map(
            (r) => `
          <tr>
            <td>${escapeHtml(r.name)}</td>
            <td class="num">${r.prev.toLocaleString()}</td>
            <td class="num">${r.curr.toLocaleString()}</td>
            <td class="num ${cls(r.deltaPct)}">${arrow(r.deltaPct)} ${pct(
              r.deltaPct
            )}</td>
          </tr>
        `
          )
          .join("")}
      </tbody>
    </table>`;

  const totalsBlock = `
    <div class="stack">
      ${renderTotalsCompareBlock(
        "Лавлагаа",
        totals.lavlagaa,
        totals.labels,
        "cmpLav"
      )}
      ${renderTotalsCompareBlock(
        "Үйлчилгээ",
        totals.uilchilgee,
        totals.labels,
        "cmpUil"
      )}
    </div>`;

  const topBlock = `
    <div class="stack">
      <div class="card">
        <div class="card-title">ТОП ЛАВЛАГАА / туслах ангилалаар /</div>
        ${makeTopTable(top.lavlagaaTop, top.labels)}
      </div>
      <div class="card">
        <div class="card-title">ТОП ҮЙЛЧИЛГЭЭ / туслах ангилалаар /</div>
        ${makeTopTable(top.uilchilgeeTop, top.labels)}
      </div>
    </div>`;

  return `
  <section class="totals-and-top">
    <div class="grid grid-2">
      ${totalsBlock}
      ${topBlock}
    </div>
  </section>`;
}

// Social heseg
function renderSocialSection(soc) {
  if (!soc) return "";

  const fmt = (n) => Number(n || 0).toLocaleString();
  const pct = (x, d = 0) => `${(x * 100).toFixed(d)}%`;
  const arrow = (v) => (v >= 0 ? "▲" : "▼");
  const cls = (v) => (v >= 0 ? "up" : "down");

  const order = [
    "Chat",
    "Comment",
    "Telegram",
    "Instagram",
    "Email",
    "Other",
    "Total",
  ];

  const trs = order
    .map((k) => {
      const r = soc.rows[k];
      return `
      <tr>
        <td>${k}</td>
        <td class="num">${fmt(r.prev)}</td>
        <td class="num">${fmt(r.curr)}</td>
        <td class="num ${cls(r.delta)}">${arrow(r.delta)} ${pct(
        Math.abs(r.delta),
        0
      )}</td>
      </tr>`;
    })
    .join("");

  return `
  <section class="social">
    <div class="card">
      <div class="card-title">Сошиал сувагийн үзүүлэлт</div>
      <table class="cmp">
        <thead>
          <tr>
            <th></th>
            <th>${soc.labels[0]}</th>
            <th>${soc.labels[1]}</th>
            <th>Хувь</th>
          </tr>
        </thead>
        <tbody>
          ${trs}
        </tbody>
      </table>
    </div>
  </section>`;
}
function parseWeekRange(s) {
  const m = String(s || "").match(
    /(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?\s*[-–]\s*(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?/
  );
  if (!m) return null;
  const y1 = m[3] ? Number(m[3]) : undefined;
  const y2 = m[6] ? Number(m[6]) : y1;
  const pad = (n) => String(n).padStart(2, "0");
  const mk = (y, M, D) => `${y ?? 2000}-${pad(M)}-${pad(D)}`;
  return {
    start: mk(y1, Number(m[1]), Number(m[2])),
    end: mk(y2, Number(m[4]), Number(m[5])),
    raw: m[0],
  };
}

function daysDiff(a, b) {
  const A = new Date(a),
    B = new Date(b);
  return Math.round((B - A) / 86400000);
}

// Social sheet дотор totaly-ийн шошготой хамгийн ойр (яг/±1 өдөр) баганыг олох
function findColIndexByWeekLabelFuzzy(rows, totalyLabel) {
  const target = parseWeekRange(totalyLabel);
  if (!target) return -1;

  // sheet-ийн бүх “week-like” нүдийг цуглуулна
  const cands = new Map(); // col -> bestLabel
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      const p = parseWeekRange(row[c]);
      if (p) cands.set(c, p.raw);
    }
  }
  if (!cands.size) return -1;

  // яг таарах эсвэл төгсгөлийн өдөр ±1 хооронд хамгийн ойрыг сонгоно
  let bestCol = -1,
    bestScore = 1e9;
  for (const [col, raw] of cands.entries()) {
    const w = parseWeekRange(raw);
    const diffEnd = Math.abs(daysDiff(w.end, target.end));
    const diffStart = Math.abs(daysDiff(w.start, target.start));
    const score = diffEnd * 2 + diffStart; // end-д илүү жин өгье
    if (diffEnd <= 1 && score < bestScore) {
      bestScore = score;
      bestCol = col;
    }
  }
  // хэрэв ±1-д багтсан олдсонгүй бол хамгийн баруун талын week-like баганыг авна
  if (bestCol < 0) {
    bestCol = Math.max(...Array.from(cands.keys()));
  }
  return bestCol;
}
function findRowByNameAnywhere(rows, name) {
  const nm = String(name).trim().toLowerCase();
  for (const r of rows) {
    if (!r) continue;
    for (let c = 0; c < r.length; c++) {
      const v = r[c];
      if (v != null && String(v).trim().toLowerCase() === nm) return r;
    }
  }
  return null;
}
function parseCellNumber(v) {
  const n = Number(String(v ?? "").replace(/[^\d.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
}
function extractSocialStatsFromTwoBooks(prevWb, currWb, sheetName = "Social") {
  const keys = [
    "Chat",
    "Comment",
    "Telegram",
    "Instagram",
    "Email",
    "Other",
    "Total",
  ];

  const readOne = (wb) => {
    const labelFromTotaly = getSingleWeekLabelFromTotaly(wb) || "";
    const ws = wb.Sheets[sheetName];
    if (!ws) throw new Error(`Sheet not found: ${sheetName}`);
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });

    const col = findColIndexByWeekLabelFuzzy(rows, labelFromTotaly);
    if (col < 0)
      throw new Error(`[Social] week column not found for ${labelFromTotaly}`);

    const data = {};
    for (const k of keys) {
      const row = findRowByNameAnywhere(rows, k);
      data[k] = parseCellNumber(row ? row[col] : 0);
    }
    return { label: labelFromTotaly, data };
  };

  const prev = readOne(prevWb);
  const curr = readOne(currWb);

  const out = {
    labels: [prev.label || "Өмнөх 7 хоног", curr.label || "Одоогийн 7 хоног"],
    rows: {},
    totalPrev: 0,
    totalCurr: 0,
  };
  for (const k of keys) {
    const a = prev.data[k] || 0,
      b = curr.data[k] || 0;
    const d = a ? (b - a) / a : b ? 1 : 0;
    out.rows[k] = { prev: a, curr: b, delta: d, share: 0 };
  }
  out.totalPrev = out.rows.Total.prev;
  out.totalCurr = out.rows.Total.curr;
  const T = out.totalCurr || 0;
  for (const k of keys) out.rows[k].share = T ? out.rows[k].curr / T : 0;
  return out;
}

// ── Бот лавлагаа / Бот үйлчилгээ (Social sheet) → 2 файлын харьцуулалт ──
function extractBotBlocksFromTwoBooks(prevWb, currWb, sheetName = "Social") {
  const LAVLAGAA_ROWS = [
    "Нууц код сэргээх",
    "Гүйлгээ пин код",
    "Данс цэнэглэх",
    "ТАН код",
    "И-мэйл солих заавар",
    "Баталгаажуулалт хийх заавар",
    "Нийт",
  ];
  const UILCHILGEE_ROWS = [
    "Дугаар солих",
    "Идэвхгүй төлөв",
    "ҮЦ данс нээх",
    "Нийт",
    "Амжилттай үзүүлсэн",
  ];

  const readOne = (wb) => {
    const label = getSingleWeekLabelFromTotaly(wb) || "";
    const ws = wb.Sheets[sheetName];
    if (!ws) throw new Error(`Sheet not found: ${sheetName}`);
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });

    const col = findColIndexByWeekLabelFuzzy(rows, label);
    if (col < 0) throw new Error(`[Bot] week column not found for ${label}`);

    const pick = (nameList) =>
      nameList.map((name) => {
        const r = findRowByNameAnywhere(rows, name);
        return { name, val: parseCellNumber(r ? r[col] : 0) };
      });

    return {
      label,
      lavlagaa: pick(LAVLAGAA_ROWS),
      uilchilgee: pick(UILCHILGEE_ROWS),
    };
  };

  const prev = readOne(prevWb);
  const curr = readOne(currWb);

  const join = (arrPrev, arrCurr) =>
    arrCurr.map((c, i) => {
      const a = arrPrev[i]?.val ?? 0;
      const b = c.val ?? 0;
      const d = a ? (b - a) / a : b ? 1 : 0;
      return { name: c.name, prev: a, curr: b, delta: d };
    });

  return {
    labels: [prev.label || "Өмнөх 7 хоног", curr.label || "Одоогийн 7 хоног"],
    lavlagaa: join(prev.lavlagaa, curr.lavlagaa),
    uilchilgee: join(prev.uilchilgee, curr.uilchilgee),
  };
}

function renderBotSections(bot) {
  const arrow = (v) => (v >= 0 ? "▲" : "▼");
  const cls = (v) => (v >= 0 ? "up" : "down");
  const pct = (v) => `${Math.abs(v * 100).toFixed(0)}%`;

  const makeTable = (rows, labels) => `
    <table class="cmp">
      <thead>
        <tr><th></th><th>${labels[0]}</th><th>${
    labels[1]
  }</th><th>Өөрчлөлт</th></tr>
      </thead>
      <tbody>
        ${rows
          .map(
            (r) => `
          <tr>
            <td>${escapeHtml(r.name)}</td>
            <td class="num">${(r.prev || 0).toLocaleString()}</td>
            <td class="num">${(r.curr || 0).toLocaleString()}</td>
            <td class="num ${cls(r.delta)}">${arrow(r.delta)} ${pct(
              r.delta
            )}</td>
          </tr>`
          )
          .join("")}
      </tbody>
    </table>`;

  const makeChart = (rows, labels, id) => {
    const cats = rows.map((r) => r.name);
    const prev = rows.map((r) => r.prev || 0);
    const curr = rows.map((r) => r.curr || 0);
    const h = Math.max(220, cats.length * 28 + 40);
    return `
      <canvas id="${id}" height="${h}"></canvas>
      <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
      <script>(function(){
        const ctx = document.getElementById('${id}').getContext('2d');
        new Chart(ctx, {
          type:'bar',
          data:{
            labels:${JSON.stringify(cats)},
            datasets:[
              { label:'${labels[0]}', data:${JSON.stringify(prev)} },
              { label:'${labels[1]}', data:${JSON.stringify(curr)} }
            ]
          },
          options:{
            indexAxis:'y',
            animation:false,
            scales:{ x:{ beginAtZero:true } },
            plugins:{ legend:{ position:'bottom' } }
          }
        });
      })();</script>`;
  };

  const block = (title, rows, chartId) => `
    <div class="grid grid-2">
      <div class="card soft">${makeChart(rows, bot.labels, chartId)}</div>
      <div class="card soft">${makeTable(rows, bot.labels)}</div>
    </div>`;

  return `
  <section class="bot-sections">
    <div class="card">
      <div class="card-title">БОТ Лавлагаа</div>
      ${block("Бот лавлагаа", bot.lavlagaa, "botLavChart")}
    </div>
    <div class="spacer16"></div>
    <div class="card">
      <div class="card-title">БОТ Үйлчилгээ</div>
      ${block("Бот үйлчилгээ", bot.uilchilgee, "botUilChart")}
    </div>
  </section>`;
}

// gomdol page
// ───────── common helpers ─────────
function weekLikeLabel(s) {
  return /(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?\s*[-–]\s*(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?/.test(
    String(s || "")
  );
}
function findWeekCols(rows) {
  const cols = new Set();
  for (const r of rows) {
    if (!r) continue;
    for (let c = 0; c < r.length; c++) {
      if (weekLikeLabel(r[c])) cols.add(c);
    }
  }
  return [...cols].sort((a, b) => a - b);
}
function norm(s) {
  return String(s || "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}
function findRowByName(rows, name, scopeRows = null) {
  const nm = norm(name);
  const range = scopeRows ?? rows;
  for (const r of range) {
    if (!r) continue;
    for (let c = 0; c < Math.min(r.length, 50); c++) {
      if (norm(r[c]) === nm) return r;
    }
  }
  return null;
}
function parseNum(v) {
  const n = Number(String(v ?? "").replace(/[^\d.-]/g, ""));
  return Number.isFinite(n) ? n : 0;
}

// ───────── GOMDOL: Comp sheet extractors ─────────
function extractGomdolFromComp(wb, sheetName = CONFIG.GOMDOL_SHEET) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet not found: ${sheetName}`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });

  const weekCols = findWeekCols(rows); // бүх 7 хоногийн баганууд
  const last4 = weekCols.slice(-4); // баруун талаас 4 долоо хоног

  // (a) Нийт гомдол (сүүлийн 4 сар) — зүүн дээд талын "6сар, 7сар..." мөрүүдээс
  // сарууд A баганад бичигддэг тул "сар" тексттэй мөрүүдийг баруун талын тоотой нь уншина
  const monthRows = rows.filter(
    (r) => r && /^(\d+)\s*сар$/i.test(String(r[0] || "").trim())
  );
  const last4Months = monthRows.slice(-4).map((r) => ({
    label: String(r[0]).trim(),
    value: parseNum(
      r[2] ?? r[1] ?? r.find((x, i) => i > 0 && String(x || "").trim() !== "")
    ), // аль байгаа тоог авна
  }));

  // (b) Шийдвэрлэлтийн 7 хоног — "Шийдэгдсэн", "Шийдэгдээгүй" мөрүүдээс
  const rSolved = findRowByName(rows, "Шийдэгдсэн") || [];
  const rUnsolved = findRowByName(rows, "Шийдэгдээгүй") || [];
  const byWeek = last4.map((c) => ({
    label: String(rows.find((r) => r && r[c])?.[c] || ""),
    solved: parseNum(rSolved[c]),
    unsolved: parseNum(rUnsolved[c]),
  }));

  // (c) ТОП 5 гомдол — доод талын “ТОП гомдол” хэсгийн хамгийн сүүлийн 7 хоногийн багана
  // Anchor: "ТОП гомдол" гэсэн нүдийг олж түүнээс доошхи мөрүүдийг уншина
  const anchorRowIdx = rows.findIndex((r) => r && norm(r[0]) === "топ гомдол");
  let top5 = [];
  if (anchorRowIdx >= 0) {
    const topArea = rows.slice(anchorRowIdx + 1, anchorRowIdx + 50); // дараагийн 50 мөр
    const lastWeekCol = last4.slice(-1)[0] ?? weekCols.slice(-1)[0];
    const pairs = [];
    for (const r of topArea) {
      const name = String(r?.[0] || "").trim();
      if (!name) break;
      const val = parseNum(r?.[lastWeekCol]);
      if (val > 0) pairs.push({ name, value: val });
    }
    pairs.sort((a, b) => b.value - a.value);
    top5 = pairs.slice(0, 5);
  }

  // (d) "Хугацаа хэтэрсэн гомдол" — комп-д баруун талд байдаг жагсаалтыг week колоноос уншина
  // Anchor-ууд: "системээс шалгалтсан", "харилцагчаас шалгалтсан", "зээл татгалзсан" гэх мэт
  // Тэдгээр нь нэг хүснэгтэд дарааллаа хадгалдаг → хоосон мөр хүртэл уншина
  const overdueAnchors = [
    "Системээс шалгалтсан",
    "Харилцагчаас шалгалтсан",
    "Play Store-оос татагдахгүй",
    "Зээлийн төлөв нээгдэхгүй",
    "Зээл хаасан, АрдКоины блок гарсан уу",
    "Иргэний үнэмлэхний баталгаажуулалт хийхгүй",
    "Нэвтэрч болохгүй",
    "Зээл татгалзсан",
    "Мөнгө шилжээгүй", // жишээ
  ];
  // Overdue хүснэгтийн дээд гарчгийг барихын тулд эдгээрийн аль нэг байрлаж буй хэсгийг хайна
  let overdueStart = rows.findIndex(
    (r) => r && r.some((v) => overdueAnchors.some((a) => norm(v) === norm(a)))
  );
  let overdue = [];
  if (overdueStart >= 0) {
    const lastWeekCol = last4.slice(-1)[0] ?? weekCols.slice(-1)[0];
    for (let i = overdueStart; i < overdueStart + 40; i++) {
      const r = rows[i];
      if (!r) break;
      const nameCell = r.find((v) => !!v);
      const name = String(nameCell || "").trim();
      if (!name) break;
      const val = parseNum(r[lastWeekCol]);
      if (!val && !weekLikeLabel(rows[0]?.[lastWeekCol])) continue;
      overdue.push({ name, value: val });
    }
    overdue = overdue.filter((x) => x.value > 0).slice(0, 10);
  }

  return { last4Months, byWeek, top5, overdue };
}

function renderGomdolSection(data) {
  const num = (n) => Number(n || 0).toLocaleString();
  const pct = (x) => `${(x * 100).toFixed(0)}%`;

  // (a) Сарын шугам график
  const months = data.last4Months.map((x) => x.label);
  const mVals = data.last4Months.map((x) => x.value);

  // (b) 7 хоног шийдвэрлэлт stacked bar
  const wLabels = data.byWeek.map((x) => x.label);
  const solved = data.byWeek.map((x) => x.solved);
  const unsolved = data.byWeek.map((x) => x.unsolved);

  // (c) ТОП 5
  const topVals = data.top5.map((x) => x.value);

  // (d) Хугацаа хэтэрсэн
  const ovVals = data.overdue.map((x) => x.value);
  const ovSum = ovVals.reduce((a, b) => a + b, 0);

  return `
  <section class="gomdol">
    <div class="grid grid-2">
      <div class="card">
        <div class="card-title">НИЙТ ГОМДОЛ /Сүүлийн 4 сараар/</div>
        <canvas id="gmMonth" height="180"></canvas>
      </div>
      <div class="card">
        <div class="card-title">ГОМДОЛ ШИЙДВЭРЛЭЛТ /7 хоногоор/</div>
        <canvas id="gmWeekly" height="200"></canvas>
      </div>
    </div>

    <div class="grid grid-2">
      ${renderTop5Card(data.top5)}
      <div class="card">
        <div class="card-title">ХУГАЦАА ХЭТРЭСЭН ГОМДОЛ /туслах ангилалаар/</div>
        <table class="cmp">
          <thead><tr><th>Туслах ангилал</th><th>Нийт</th><th>Хувь</th></tr></thead>
          <tbody>
            ${data.overdue
              .map(
                (r) => `
                <tr>
                  <td>${escapeHtml(r.name)}</td>
                  <td class="num">${num(r.value)}</td>
                  <td class="num">${pct(ovSum ? r.value / ovSum : 0)}</td>
                </tr>
              `
              )
              .join("")}
          </tbody>
        </table>
      </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
    <script>
    (function(){
      // Months line
      new Chart(document.getElementById('gmMonth').getContext('2d'),{
        type:'line',
        data:{ labels:${JSON.stringify(
          months
        )}, datasets:[{ label:'Нийт гомдол', data:${JSON.stringify(mVals)} }]},
        options:{ animation:false, plugins:{legend:{display:false}}, scales:{ y:{ beginAtZero:true } } }
      });

      // Weekly stacked
      new Chart(document.getElementById('gmWeekly').getContext('2d'),{
        type:'bar',
        data:{ labels:${JSON.stringify(wLabels)}, datasets:[
          { label:'Шийдэгдсэн', data:${JSON.stringify(solved)} },
          { label:'Шийдэгдээгүй', data:${JSON.stringify(unsolved)} }
        ]},
        options:{ animation:false, scales:{ x:{ stacked:true }, y:{ stacked:true, beginAtZero:true } }, plugins:{ legend:{ position:'bottom' } } }
      });
    })();
    </script>
  </section>`;
}

function parsePct01(x) {
  const s = String(x ?? "").trim();
  if (/^-?\d+(\.\d+)?\s*%$/.test(s)) return Number(s.replace("%", "")) / 100;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}
function parseWeekRangeForSort(s) {
  const m = String(s || "").match(
    /(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?\s*[-–]\s*(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?/
  );
  if (!m) return null;
  const y = m[6] ? Number(m[6]) : m[3] ? Number(m[3]) : 2000;
  const pad = (n) => String(n).padStart(2, "0");
  // sort хийхэд ашиглах YYYY-MM-DD (range-ийн ТӨГСГӨЛ) -> илүү сүүлчийн долоо хоногийг сонгоно
  return `${y}-${pad(m[5])}-${pad(m[4])}`;
}
// Sheet2 → зүүн талын пивот (A:B) : Row Labels | Count of ...
function extractTop5FromComp(wb, sheetName = CONFIG.GOMDOL_SHEET) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet not found: ${sheetName}`);

  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });

  // Helpers
  const isPct = (v) =>
    typeof v === "string" && /^\s*-?\d+(?:\.\d+)?\s*%$/.test(v.trim());
  const toPct01 = (v) =>
    isPct(v) ? Number(String(v).replace("%", "")) / 100 : 0;
  const toInt = (v) => {
    const n = Number(String(v ?? "").replace(/[^\d-]/g, ""));
    return Number.isFinite(n) ? n : 0;
  };
  const weekLike = (s) =>
    /(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?\s*[-–]\s*(\d{1,2})[./-](\d{1,2})(?:[./-](\d{2,4}))?/.test(
      String(s || "")
    );

  // 1) "ТОП гомдол" anchor мөр
  const anchorR = rows.findIndex(
    (r) => r && r.some((v) => /ТОП\s*гомдол/i.test(String(v || "")))
  );
  if (anchorR < 0) {
    console.warn("[TOP5] Anchor not found");
    return [];
  }
  const anchorRow = rows[anchorR] || [];

  // 2) Сонгох 7 хоногийн COUNT баганын (currCol) индексийг олно
  let currCol = null;

  if (typeof CONFIG.GOMDOL_FORCE_WEEK_COL === "number") {
    currCol = CONFIG.GOMDOL_FORCE_WEEK_COL;
    console.log(
      "[TOP5] Using forced week COL:",
      currCol,
      "label=",
      anchorRow[currCol]
    );
  } else if (CONFIG.GOMDOL_FORCE_WEEK_LABEL) {
    const want = String(CONFIG.GOMDOL_FORCE_WEEK_LABEL).trim();
    const idx = anchorRow.findIndex((v) => String(v || "").trim() === want);
    if (idx >= 0) {
      currCol = idx;
      console.log("[TOP5] Using forced week LABEL:", want, "→ col:", currCol);
    }
  }

  if (currCol == null) {
    // автоматаар: anchor мөрөн дэх бүх week-like толгой
    const cand = [];
    for (let c = 0; c < anchorRow.length; c++)
      if (weekLike(anchorRow[c])) cand.push(c);
    if (!cand.length) {
      console.warn("[TOP5] No week-like columns in anchor row");
      return [];
    }
    currCol = cand[cand.length - 1]; // хамгийн баруунд буйг сүүлийн 7 хоног гэж үзнэ
    console.log(
      "[TOP5] Auto week col =",
      currCol,
      "label=",
      anchorRow[currCol]
    );
  }

  // 3) currCol орчимд (зүүн тийш 3, баруун тийш 4) “Нэр|Тоо|%” гурвалыг мөр бүрээс түүж авна
  const W_LEFT = 3; // name ихэвчлэн currCol-1 эсвэл -2 талд байдаг
  const W_RIGHT = 4; // тоо=currCol, % баруун талд 1–2 багана
  const items = [];

  for (let r = anchorR + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    // бүхэл мөр хоосон бол блок дууслаа
    if (!row.some((v) => String(v ?? "").trim())) break;

    // 3.1 нэрийг currCol-1 .. currCol-3 дотроос хамгийн ойрыг сонгоно
    let name = "";
    for (let c = currCol - 1; c >= Math.max(0, currCol - W_LEFT); c--) {
      const s = String(row[c] ?? "").trim();
      if (!s) continue;
      // тоо/хувь/долоо хоногийн шошго биш, “цэвэр текст” байвал нэр гэж үзнэ
      if (!/^-?\d+(\.\d+)?$/.test(s) && !isPct(s) && !weekLike(s)) {
        name = s;
        break;
      }
    }

    // 3.2 тоо = currCol эсвэл баруун тийш ойрын integer
    let count = toInt(row[currCol]);
    if (!count) {
      for (let c = currCol + 1; c <= currCol + W_RIGHT; c++) {
        const v = row[c];
        const n = toInt(v);
        if (n) {
          count = n;
          break;
        }
      }
    }

    // 3.3 хувь = currCol+1 .. currCol+W_RIGHT дотроос эхний %
    let pct = 0;
    for (let c = currCol + 1; c <= currCol + W_RIGHT; c++) {
      const v = row[c];
      if (isPct(v)) {
        pct = toPct01(v);
        break;
      }
    }

    // 3.4 Хүчинтэй мөрийг хадгална
    if (name && count > 0) {
      items.push({ name, value: count, pct });
    }
  }

  // 4) TOP-5 буцаана
  items.sort((a, b) => b.value - a.value);
  return items.slice(0, 5);
}
function renderTop5Chart(top5) {
  const labels = top5.map((x) => x.name);
  const counts = top5.map((x) => x.value);
  const pcts = top5.map((x) => x.pct ?? 0);

  // Өндрийг мөрийн тооноос автоматаар
  const h = Math.max(220, labels.length * 30 + 40);

  // % текстийг барын ард бичих жижиг plugin
  const plugin = `
    const pctLabel = {
      id: 'pctLabel',
      afterDatasetsDraw(chart){
        const {ctx, scales:{x,y}} = chart;
        ctx.save();
        ctx.font = '12px system-ui,-apple-system,Segoe UI,Roboto,Arial';
        ctx.fillStyle = '#333';
        ${JSON.stringify(counts)}.forEach((v,i)=>{
          const xPos = x.getPixelForValue(v) + 8;
          const yPos = y.getPixelForValue(i);
          const pct  = (${JSON.stringify(pcts)}[i] || 0) * 100;
          ctx.fillText(pct.toFixed(0) + '%', xPos, yPos + 4);
        });
        ctx.restore();
      }
    };
  `;

  return `
    <canvas id="gmTop5" height="${h}"></canvas>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
    <script>
      (function () {
        ${plugin}
        const ctx = document.getElementById('gmTop5').getContext('2d');
        new Chart(ctx, {
          type: 'bar',
          data: {
            labels: ${JSON.stringify(labels)},
            datasets: [{ label: '', data: ${JSON.stringify(counts)} }]
          },
          options: {
            indexAxis: 'y',
            animation: false,
            plugins: { legend: { display: false } },
            scales: { x: { beginAtZero: true } }
          },
          plugins: [pctLabel]
        });
      })();
    </script>
  `;
}
function renderTop5Card(top5) {
  if (!top5 || !top5.length) {
    return `<div class="card"><div class="card-title">ТОП 5 гомдол</div>
            <p style="color:#a33">Мэдээлэл олдсонгүй.</p></div>`;
  }
  return `
    <div class="card">
      <div class="card-title">ТОП 5 гомдол /хамгийн сүүлийн 7 хоног/</div>
      ${renderTop5Chart(top5)}
    </div>
  `;
}
function extractOverdueFromSheet2(wb, sheetName = "Sheet2") {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet not found: ${sheetName}`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });

  const isPct = (v) =>
    typeof v === "string" && /\d+(\.\d+)?\s*%$/.test(v.trim());
  const isNum = (v) =>
    Number.isFinite(Number(String(v ?? "").replace(/[^\d.-]/g, "")));
  const num = (v) => Number(String(v ?? "").replace(/[^\d.-]/g, ""));

  let best = { col: -1, start: -1, items: [] };

  // бүх багануудаар гүйлгээд "Нэр | Тоо | %" 3-тын хэв загварыг хайна
  const maxCols = rows.reduce((m, r) => Math.max(m, r?.length || 0), 0);
  for (let c = 0; c <= maxCols - 3; c++) {
    let items = [];
    let started = false;

    for (let r = 0; r < rows.length; r++) {
      const name = String(rows[r]?.[c] ?? "").trim();
      const cnt = rows[r]?.[c + 1];
      const pct = rows[r]?.[c + 2];

      const nameOk = !!name && !/(Grand\s*Total|\(blank\))/i.test(name);
      const cntOk = isNum(cnt);
      const pctOk = isPct(pct);

      // эхлэх нөхцөл: гурвуулаа хэвийн
      if (!started) {
        if (nameOk && cntOk && pctOk) {
          started = true;
          items.push({ name, value: num(cnt) });
        }
        continue;
      }

      // үргэлжлэх: нэр хоосон бол энэ блок дууссан гэж үзнэ
      if (!name) break;

      if (nameOk && cntOk) {
        items.push({ name, value: num(cnt) });
      } else {
        // эвдэрсэн мөр таарвал блок дууслаа
        break;
      }
    }

    if (items.length > best.items.length) {
      best = { col: c, start: 0, items };
    }
  }

  // 0-с биш тоонуудыг үлдээгээд TOP-10
  return (best.items || []).filter((x) => x.value > 0).slice(0, 10);
}
// gomdol end

// gomdol shiidverlelt
function extractOsticketResolutionBlock(wb, sheetName = CONFIG.OST_SHEET) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet not found: ${sheetName}`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });

  // дэмжих нэршлүүд (үл ялиг бичгийн алдааг бас барина)
  const KEY = {
    total: /нийт\s*хаагдсан/i,
    most: /мост\s*кол+|мост\s*кал+/i,
    uz: /\bүз\b|анхны\s*хандалт/i,
    gomdolAj: /гомдлын\s*аж(илтан)?/i,
    busad: /бусад/i,
  };

  const toInt = (v) => {
    const n = Number(String(v ?? "").replace(/[^\d-]/g, ""));
    return Number.isFinite(n) ? n : 0;
  };
  const isPct = (v) =>
    typeof v === "string" && /\d+(\.\d+)?\s*%$/.test(v.trim());
  const toPct01 = (v) =>
    isPct(v) ? Number(String(v).replace("%", "")) / 100 : 0;

  // мөр бүрээс: нэршил таарвал → баруун талд орших эхний тоо (count) ба эхний % (pct)-ийг авна
  const pickRow = (re) => {
    for (const r of rows) {
      if (!r) continue;
      // мөрт байгаа бүх нүднүүдээр нэршил хайна
      if (!r.some((v) => re.test(String(v || "")))) continue;

      let count = 0,
        pct = 0;
      // баруун тийшхи эхний тоо ба эхний %-ийг цуглуулъя
      for (let i = 0; i < r.length; i++) {
        const v = r[i];
        if (!count && toInt(v)) count = toInt(v);
        if (!pct && isPct(v)) pct = toPct01(v);
        if (count && pct) break;
      }
      // тоо байхгүй ч өөр мөр дээр (дараагийн 1-2 мөр) багана шилжилт байж магадгүй — нэмж хайна
      if (!count || !pct) {
        const idx = rows.indexOf(r);
        for (let k = 1; k <= 2 && idx + k < rows.length; k++) {
          const rr = rows[idx + k] || [];
          for (let i = 0; i < rr.length; i++) {
            const v = rr[i];
            if (!count && toInt(v)) count = toInt(v);
            if (!pct && isPct(v)) pct = toPct01(v);
            if (count && pct) break;
          }
          if (count && pct) break;
        }
      }
      return { count, pct };
    }
    return { count: 0, pct: 0 };
  };

  const total = pickRow(KEY.total).count;
  const uz = pickRow(KEY.uz).count;
  const gomdolAj = pickRow(KEY.gomdolAj).count;
  let busad = pickRow(KEY.busad).count;

  // Хэрэв "Бусад" байхгүй бол нийтээс тооцоолж гаргая
  if (!busad && total) {
    busad = Math.max(total - (uz + gomdolAj), 0);
  }

  const safeTot = total || uz + gomdolAj + busad;
  const share = (n) => (safeTot ? n / safeTot : 0);

  return {
    total: safeTot,
    most: uz + gomdolAj,
    uz,
    gomdolAj,
    busad,
    shares: {
      most: share(uz + gomdolAj),
      uz: share(uz),
      gomdolAj: share(gomdolAj),
      busad: share(busad),
    },
  };
}

function renderOsticketResolutionCard(data) {
  const num = (n) => Number(n || 0).toLocaleString();
  const pct = (x) => `${(x * 100).toFixed(0)}%`;

  return `
  <section class="ost-resolution">
    <div class="card">
      <div class="card-title">Гомдол шийдвэрлэлт (Osticket)</div>
      <table class="cmp">
        <tbody>
          <tr><td>Нийт хаагдсан</td><td class="num">${num(
            data.total
          )}</td><td></td></tr>
          <tr><td>Мост колл</td><td class="num">${num(data.most)}</td><td>${pct(
    data.shares.most
  )}</td></tr>
          <tr><td>&nbsp;&nbsp;• ҮЗ (анхны хандалт)</td><td class="num">${num(
            data.uz
          )}</td><td>${pct(data.shares.uz)}</td></tr>
          <tr><td>&nbsp;&nbsp;• Гомдлын ажилтан</td><td class="num">${num(
            data.gomdolAj
          )}</td><td>${pct(data.shares.gomdolAj)}</td></tr>
          <tr><td>Шилжүүлсэн (бусад компани)</td><td class="num">${num(
            data.busad
          )}</td><td>${pct(data.shares.busad)}</td></tr>
        </tbody>
      </table>
    </div>
  </section>`;
}
// gomdol shiidverlelt duusna

//gomdol company
// --- Comp sheet: ЯГ заасан 7 хоногийн баганаас утгыг шууд унших ---
function extractCompanyTotalsLatestWeek(wb, sheetName = "Comp") {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet not found: ${sheetName}`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });

  const WEEK = String(CONFIG.COMPANY_WEEK_LABEL || "").trim();
  const NAME_COL = Number.isFinite(CONFIG.COMPANY_NAME_COL)
    ? CONFIG.COMPANY_NAME_COL
    : 0;

  const WANT = (CONFIG.COMPANIES || []).map((s) => String(s).trim());
  const order = new Map(WANT.map((n, i) => [n, i]));
  const norm = (s) =>
    String(s || "")
      .replace(/\s+/g, " ")
      .trim()
      .toLowerCase();
  const toInt = (v) => {
    const n = Number(String(v ?? "").replace(/[^\d-]/g, ""));
    return Number.isFinite(n) ? n : 0;
  };

  // 1) WEEK шошготой баганыг олно (нэгээс олон бол хамгийн доорх блоктой дүйхээр нь → хамгийн доод мөр дээрхийг сонгоё)
  const candidates = [];
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r] || [];
    for (let c = 0; c < row.length; c++) {
      if (String(row[c] || "").trim() === WEEK) candidates.push({ r, c });
    }
  }
  if (!candidates.length) {
    console.warn("[Company] Week label not found:", WEEK);
    return WANT.map((n) => ({ name: n, value: 0, pct: 0 }));
  }
  // ихэнх файлуудад дээш/доош 2 удаа давтагддаг → доод талынх нь компанийн жагсаалттай ойр
  const { r: weekHdrRow, c: weekCol } = candidates[candidates.length - 1];

  // 2) Компанийн нэртэй мөрийг олж weekCol дээрх утгыг авна
  const bag = new Map(); // name -> value
  const wantSet = new Map(WANT.map((n) => [norm(n), n]));

  for (let r = weekHdrRow + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    // нэрийг NAME_COL-оос; хэрэв хоосон бол зүүн 0..3 багана дундаас хайна
    let rawName = row[NAME_COL];
    if (!rawName) {
      for (let k = 0; k < Math.min(4, row.length); k++) {
        if (row[k]) {
          rawName = row[k];
          break;
        }
      }
    }
    const canon = wantSet.get(norm(rawName));
    if (!canon) continue;

    // яг тэр баганын тоо (заримдаа 1-2 нүд баруун тийш гажсан байж болзошгүй → ойрынд нь fallback)
    let val = toInt(row[weekCol]);
    if (!val && row.length > weekCol + 1) val = toInt(row[weekCol + 1]);
    if (!val && row.length > weekCol + 2) val = toInt(row[weekCol + 2]);

    bag.set(canon, val);
  }

  // 3) Дарааллаар буцаах (байхгүй нь 0)
  const out = WANT.map((n) => ({ name: n, value: bag.get(n) ?? 0, pct: 0 }));
  out.sort((a, b) => (order.get(a.name) ?? 999) - (order.get(b.name) ?? 999));
  console.log(
    `[Company] Using week "${WEEK}" at col=${weekCol}, headerRow=${weekHdrRow}`
  );
  return out;
}

function renderComplaintsByCompanySection(items) {
  if (!items || !items.length) {
    return `<section class="card"><div class="card-title">Гомдол /Компанийн/</div>
            <p style="color:#a33">Мэдээлэл олдсонгүй.</p></section>`;
  }
  const labels = items.map((x) => x.name);
  const counts = items.map((x) => x.value);
  const pcts = items.map((x) => Math.round((x.pct || 0) * 100));
  const hasPct = pcts.some((v) => v > 0);

  const plugin = `
    const badge = {
      id:'badge',
      afterDatasetsDraw(chart){
        const {ctx, scales:{x,y,y1}} = chart;
        ctx.save();
        ctx.font = '12px system-ui,-apple-system,Segoe UI,Roboto,Arial';
        ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
        // count badge
        chart.data.datasets[0].data.forEach((v,i)=>{
          const xPos = x.getPixelForValue(i);
          const yPos = y.getPixelForValue(v) - 10;
          const text = String(v);
          const w = ctx.measureText(text).width + 10, h = 16;
          ctx.fillStyle = 'rgba(59,130,246,.95)';
          ctx.fillRect(xPos - w/2, yPos - h, w, h);
          ctx.fillStyle = '#fff'; ctx.fillText(text, xPos, yPos - h/2);
        });
        // pct badge (байвал)
        if (${hasPct}) {
          chart.data.datasets[1].data.forEach((v,i)=>{
            const xPos = x.getPixelForValue(i);
            const yPos = y1.getPixelForValue(v);
            const text = v + '%';
            const w = ctx.measureText(text).width + 10, h = 16;
            ctx.fillStyle = 'rgba(239,68,68,.95)';
            ctx.fillRect(xPos - w/2, yPos - h - 4, w, h);
            ctx.fillStyle = '#fff'; ctx.fillText(text, xPos, yPos - h/2 - 4);
          });
        }
        ctx.restore();
      }
    };
  `;

  return `
  <section class="gomdol-company">
    <div class="card">
      <div class="card-title">ГОМДОЛ /Компанийн/ — ${
        CONFIG.COMPANY_WEEK_LABEL || "Сүүлийн 7 хоног"
      }</div>
      <canvas id="cmpCompany" height="360"></canvas>
      <div class="legend" style="margin-top:8px">
        <span class="dot dot-total"></span> Бүртгэсэн гомдол
        ${
          hasPct
            ? `<span class="dot dot-line" style="margin-left:16px"></span> Хугацаа хэтэрсний хувь`
            : ""
        }
      </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
    <script>
      (function(){
        ${plugin}
        const ctx = document.getElementById('cmpCompany').getContext('2d');
        new Chart(ctx, {
          data: {
            labels: ${JSON.stringify(labels)},
            datasets: [
              { type:'bar',  label:'Бүртгэсэн гомдол', data:${JSON.stringify(
                counts
              )}, yAxisID:'y' }
              ${
                hasPct
                  ? `, { type:'line', label:'Хугацаа хэтэрсэн хувь', data:${JSON.stringify(
                      pcts
                    )}, yAxisID:'y1', borderDash:[6,6], pointRadius:4 }`
                  : ``
              }
            ]
          },
          options: {
            animation:false,
            scales:{
              y:  { beginAtZero:true, title:{display:true, text:'Тоо'} },
              ${
                hasPct
                  ? `y1: { position:'right', beginAtZero:true, max:100, title:{display:true, text:'%'}, ticks:{ callback:(v)=>v+'%' }, grid:{ drawOnChartArea:false } },`
                  : ``
              }
            },
            plugins:{ legend:{ display:false } }
          },
          plugins:[badge]
        });
      })();
    </script>
  </section>`;
}

//gomdol company end

async function htmlToPdf(html, outPath) {
  const browser = await puppeteer.launch({
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });
  try {
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });
    await page.emulateMediaType("screen"); // сонголттой
    await page.pdf({
      path: outPath,
      format: "A4", // ← size биш format
      landscape: true, // ← хөндлөн
      printBackground: true,
      preferCSSPageSize: true,
      margin: { top: "16mm", right: "14mm", bottom: "16mm", left: "14mm" },
    });
  } finally {
    await browser.close();
  }
}

const EMAIL_ENABLED = String(process.env.EMAIL_ENABLED ?? "true") === "true";
const SCHED_ENABLED = String(process.env.SCHED_ENABLED ?? "true") === "true";

async function sendEmailWithPdf(pdfPath, subject) {
  // 465 бол secure=true; бусад нь STARTTLS (secure=false)
  if (!EMAIL_ENABLED) {
    console.log("[EMAIL] Disabled by EMAIL_ENABLED=false. Skipping send.");
    return; // ← мэйл илгээхгүй
  }

  const port = Number(process.env.SMTP_PORT || 587);
  const secure =
    port === 465 ? true : String(process.env.SMTP_SECURE || "false") === "true";

  const transporter = nodemailer.createTransport({
    host: SMTP_HOST,
    port: Number(SMTP_PORT || 587),
    secure: String(SMTP_SECURE || "false") === "true", // 465:true, 587:false
    auth:
      SMTP_USER && SMTP_PASS ? { user: SMTP_USER, pass: SMTP_PASS } : undefined,
    // ↓ Сүлжээ талын асуудлыг оношлох/давтахад туслана
    pool: true,
    maxConnections: 1,
    connectionTimeout: 20_000, // 20s
    greetingTimeout: 15_000,
    socketTimeout: 30_000,
    requireTLS: SMTP_PORT === "587", // STARTTLS шаардана (шаардлагатай бол)
    tls: {
      // Зарим enterprise SMTP зөвхөн TLS1.2+ шаарддаг
      minVersion: "TLSv1.2",
      // Хуучин/self-signed сертификаттай бол түр хугацаанд:
      // rejectUnauthorized: false
    },
    logger: true, // консолд протоколын лог гаргана
    debug: true,
  });

  // Эхлээд конфиг зөв эсэхийг шалгаж үзнэ
  try {
    await transporter.verify(); // handshake test
    console.log("[SMTP] verify OK");
  } catch (e) {
    console.error("[SMTP] verify FAILED:", e);
    throw e;
  }

  const htmlIntro = `
    <div style="font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;line-height:1.45">
      <p>Сайн байна уу,</p>
      <p>Хавсралтад долоо хоногийн тайланг (PDF) илгээлээ.</p>
      <p style="color:#666;font-size:12px">Автоматаар илгээв.</p>
    </div>
  `;

  await transporter.sendMail({
    from: process.env.FROM_EMAIL,
    to: process.env.RECIPIENTS,
    subject,
    html: htmlIntro,
    attachments: [
      {
        filename: path.basename(pdfPath),
        path: pdfPath,
        contentType: "application/pdf",
      },
    ],
  });
}

function wrapHtml(bodyHtml) {
  const css = fs.readFileSync(CSS_FILE, "utf-8");
  return `<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <style>${css}</style>
</head>
<body>
  ${bodyHtml}
  <div class="footer">Автоматаар бэлтгэсэн тайлан (Node.js)</div>
</body>
</html>`;
}

// ────────────────────────────────────────────────────────────────
// Гол ажил: Excel → HTML → PDF → Mail
// ────────────────────────────────────────────────────────────────
async function runOnce() {
  if (!fs.existsSync(CONFIG.EXCEL_FILE_CURR))
    throw new Error(`EXCEL_FILE_CURR not found: ${CONFIG.EXCEL_FILE_CURR}`);
  if (!fs.existsSync(CONFIG.EXCEL_FILE_PREV))
    throw new Error(`EXCEL_FILE_PREV not found: ${CONFIG.EXCEL_FILE_PREV}`);
  if (!fs.existsSync(CSS_FILE))
    throw new Error(`CSS template not found: ${CSS_FILE}`);
  if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

  const wbPrev = xlsx.readFile(CONFIG.EXCEL_FILE_PREV, { cellDates: true });
  const wbCurr = xlsx.readFile(CONFIG.EXCEL_FILE_CURR, { cellDates: true });

  const channels = extractChannelsFromTotaly(wbCurr);
  const kpi = extractTotalyKPI(wbCurr);

  const totals = extractOsticketTotalsFromTwoBooks(wbPrev, wbCurr);
  const ostTop = extractOsticketTopBySubcatFromTwoBooks(wbPrev, wbCurr);
  const social = extractSocialStatsFromTwoBooks(wbPrev, wbCurr);

  const bot = extractBotBlocksFromTwoBooks(wbPrev, wbCurr);

  const wbG = xlsx.readFile(CONFIG.GOMDOL_FILE, { cellDates: true });
  const core = extractGomdolFromComp(wbG); // сар + 7 хоног шийдвэрлэлт
  const top5 = extractTop5FromComp(wbG, "Comp");
  const overdue = extractOverdueFromSheet2(wbG);
  const gomdol = {
    last4Months: core.last4Months,
    byWeek: core.byWeek,
    top5,
    overdue,
  };

  const ostRes = extractOsticketResolutionBlock(wbCurr, CONFIG.OST_SHEET);
  const ostHtml = renderOsticketResolutionCard(ostRes);

  const compByCompany = extractCompanyTotalsLatestWeek(wbG, "Comp");
  const compCompanyHtml = renderComplaintsByCompanySection(compByCompany);

  let socialHtml = `
  <section class="social">
    <div class="card">
      <div class="card-title">Сошиал сувагийн үзүүлэлт</div>
      ${makeSocialNarrative(social, { chatDup: 126, chatSucc: 63 })}
      ${
        /* энд Social хүснэгтээ renderSocialSection(social)-ийн table хэсгийг тавина */ ""
      }
    </div>
  </section>
`;

  // const social = extractSocialStats(wbCurr);

  // 3) Нүүр хуудасны KPI блок
  let body = "";
  body += renderTotalyKPISection(kpi);
  body += renderChannelCards(channels);
  body += renderTotalsAndTopSection(totals, ostTop);
  body += renderSocialSection(social);
  body += socialHtml;
  body += ostHtml;
  body += renderBotSections(bot);
  body += renderGomdolSection(gomdol);
  body += compCompanyHtml;

  // (хүсвэл бусад sheet-үүдийн хүснэгтийг доор нэмж болно)
  // body += buildHtmlFromWorkbook(wb); // хүсвэл бүх sheet-үүдийг дагаад нэм

  // 4) Финал HTML
  const html = wrapHtml(body);

  // 5) PDF болгож илгээх
  const monday = dayjs().tz(TIMEZONE).startOf("week").add(1, "day");
  const stamp = monday.format("YYYYMMDD");
  const pdfName = `weekly-report-${stamp}.pdf`;
  const pdfPath = path.join(OUTPUT_DIR, pdfName);

  await htmlToPdf(html, pdfPath);

  const subject = `${SUBJECT_PREFIX} ${REPORT_TITLE} — ${monday.format(
    "YYYY-MM-DD"
  )}`;
  await sendEmailWithPdf(pdfPath, subject);

  console.log(`[OK] Sent ${pdfName} to ${RECIPIENTS}`);
}
// ────────────────────────────────────────────────────────────────
// Scheduler: Даваа бүр 09:00 (Asia/Ulaanbaatar)
// ────────────────────────────────────────────────────────────────
function startScheduler() {
  // “0 9 * * 1” → 09:00 every Monday
  cron.schedule(
    "0 9 * * 1",
    async () => {
      try {
        await runOnce();
      } catch (err) {
        console.error("[ERROR] runOnce:", err);
      }
    },
    {
      timezone: TIMEZONE,
    }
  );

  console.log(`Scheduler started → Every Monday 09:00 (${TIMEZONE})`);
}

// ────────────────────────────────────────────────────────────────
const runNow = process.argv.includes("--once");
if (runNow) {
  runOnce().catch((err) => {
    console.error(err);
    process.exit(1);
  });
} else {
  if (SCHED_ENABLED) {
    startScheduler();
  } else {
    console.log(`Scheduler disabled (SCHED_ENABLED=false). No auto emails.`);
  }
}
