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

  // Excel files
  CURR_FILE: "./ARD 8.04-8.10.xlsx",
  PREV_FILE: "./ARD 7.28-8.03.xlsx",
  GOMDOL_FILE: "./gomdol-weekly.xlsx",
  // Sheets
  ASS_SHEET: "APS",
  ASS_COMPANY: "Ардын Тэтгэврийн Данс",
  ASS_YEAR: "2025",
  ASS_TAKE_LAST_N_MONTHS: 4,
  OST_SHEET: "Osticket1",

  OUTBOUND_FILE: "./Outbound.xlsx",
  OUTBOUND_SHEET: "AARD-HVL",

  // PDF / Email
  OUT_DIR: "./out",
  CSS_FILE: "./css/template.css",
  SAVE_HTML: true,
  HTML_NAME_PREFIX: "report",
  REPORT_TITLE: process.env.REPORT_TITLE || "Ард Апп — 7 хоногийн тайлан",
  SUBJECT_PREFIX: process.env.SUBJECT_PREFIX || "[ArdApp Weekly]",
  EMAIL_ENABLED: String(process.env.EMAIL_ENABLED ?? "true") === "true",
  SCHED_ENABLED: String(process.env.SCHED_ENABLED ?? "true") === "true",
};

function _normTxt(s) {
  return String(s || "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}
function findRowIndex(rows, pred) {
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || [];
    if (pred(r, i)) return i;
  }
  return -1;
}

const num = (v) => Number(String(v ?? "").replace(/[^\d.-]/g, "")) || 0;
const pct01 = (v) => {
  const s = String(v ?? "").trim();
  return /^\d+(\.\d+)?%$/.test(s)
    ? Number(s.replace("%", "")) / 100
    : Number(s) || 0;
};

function extractOutbound_AARD_HVL(xlsPath, sheetName = CONFIG.OUTBOUND_SHEET) {
  const wb = xlsx.readFile(xlsPath, { cellDates: true });
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`[Outbound] Sheet not found: ${sheetName}`);

  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  // Толгой мөрийг олоход уян хатан – 'Огноо/Залгалт/Амжилттай/SR'
  const headRowIdx = rows.findIndex(
    (r) =>
      r?.some((v) => /огноо/i.test(v)) &&
      r?.some((v) => /залгалт/i.test(v)) &&
      r?.some((v) => /амжилттай/i.test(v)) &&
      r?.some((v) => /^sr$/i.test(String(v).trim()))
  );
  if (headRowIdx < 0) throw new Error("[Outbound] Header row not found");

  const head = rows[headRowIdx].map((v) =>
    String(v || "")
      .trim()
      .toLowerCase()
  );
  const idx = {
    date: head.findIndex((h) => /огноо/i.test(h)),
    total: head.findIndex((h) => /залгалт/i.test(h)),
    success: head.findIndex((h) => /амжилттай/i.test(h)),
    sr: head.findIndex((h) => /^sr$/.test(h)),
  };
  if (Object.values(idx).some((i) => i < 0))
    throw new Error("[Outbound] Missing columns");

  const items = [];
  for (let r = headRowIdx + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const name = String(row[idx.date] ?? "").trim();
    if (!name) continue;
    if (/нийт/i.test(name)) break; // “Нийт” мөрөөс цааш уншихгүй

    const total = num(row[idx.total]);
    const success = num(row[idx.success]);
    // SR багана байвал авна, байхгүй бол тооцоолно
    const sr = (pct01(row[idx.sr]) || (total ? success / total : 0)) * 100;

    items.push({
      date: name,
      total,
      success,
      sr: Math.round(sr),
    });
  }

  const sums = items.reduce(
    (a, b) => ({
      total: a.total + b.total,
      success: a.success + b.success,
    }),
    { total: 0, success: 0 }
  );
  const srTotal = Math.round((sums.success * 100) / Math.max(1, sums.total));

  return { items, sums: { ...sums, sr: srTotal } };
}

function renderOutboundTable(outbound) {
  return `
  <div class="row g-3">
    <div class="col-lg-6">
      <div class="card shadow-sm">
        <div class="card-header text-center fw-semibold" 
             style="background:#f6c342;color:#4a3d00;border-top-left-radius:.5rem;border-top-right-radius:.5rem">
          OUTBOUND
        </div>
        <div class="card-body p-0">
          <table class="table table-bordered table-striped mb-0 align-middle">
            <thead class="table-light text-center">
              <tr>
                <th style="width:28%">Огноо</th>
                <th style="width:24%">Залгалт</th>
                <th style="width:24%">Амжилттай</th>
                <th style="width:24%">SR</th>
              </tr>
            </thead>
            <tbody>
              ${outbound.items
                .map(
                  (r) => `
                <tr>
                  <td>${escapeHtml(r.date)}</td>
                  <td class="text-end">${r.total.toLocaleString()}</td>
                  <td class="text-end">${r.success.toLocaleString()}</td>
                  <td class="text-end">${r.sr}%</td>
                </tr>`
                )
                .join("")}
              <tr class="fw-bold">
                <td>Нийт</td>
                <td class="text-end">${outbound.sums.total.toLocaleString()}</td>
                <td class="text-end">${outbound.sums.success.toLocaleString()}</td>
                <td class="text-end">${outbound.sums.sr}%</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- Хүсвэл баруун талд (col-lg-6) өсөлт/бууралтын график эсвэл өөр блок тавина -->
    <div class="col-lg-6">
      <div class="card h-100">
        <div class="card-header fw-semibold">Өсөлт / бууралт</div>
        <div class="card-body">
          <canvas id="outboundTrend" height="180"></canvas>
        </div>
      </div>
    </div>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
  <script>
    (function(){
      const ctx = document.getElementById('outboundTrend').getContext('2d');
      new Chart(ctx,{
        type:'bar',
        data:{
          labels:${JSON.stringify(outbound.items.map((i) => i.date))},
          datasets:[
            {label:'Залгалт', data:${JSON.stringify(
              outbound.items.map((i) => i.total)
            )}},
            {label:'Амжилттай', data:${JSON.stringify(
              outbound.items.map((i) => i.success)
            )}}
          ]
        },
        options:{animation:false, plugins:{legend:{position:'bottom'}}, scales:{y:{beginAtZero:true}}}
      });
    })();
  </script>`;
}

function escapeHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

// APS sheet → зүүн блокийн жилүүдтэй хүснэгтээс (2025 багана) сүүлийн N сар
function extractAPSLeftBlockLastMonths(
  wb,
  { sheetName = "APS", yearLabel = "2025", takeLast = 4 } = {}
) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`[APS] Sheet not found: ${sheetName}`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });

  // 1) Жилүүдийн толгойг олох (… 2021 2022 2023 2024 2025 …)
  const headRowIdx = rows.findIndex((r) =>
    (r || []).filter(Boolean).some((v) => /^\d{4}$/.test(String(v)))
  );
  if (headRowIdx < 0) throw new Error("[APS] Year header row not found");

  const header = rows[headRowIdx].map((v) => String(v || "").trim());
  const yearCol = header.findIndex((v) => v === String(yearLabel));
  if (yearCol < 0) throw new Error(`[APS] Year column not found: ${yearLabel}`);

  // 2) Доош сараар явах: тоон утгуудыг цуглуулна, "Нийт" мөр хүртэл эсвэл хоосон блок хүртэл
  const values = [];
  for (let r = headRowIdx + 1, m = 1; r < rows.length; r++, m++) {
    const row = rows[r] || [];
    const leftCell = String(row[0] ?? row[1] ?? "").trim();
    if (/^нийт$/i.test(leftCell)) break; // блокын төгсгөл
    const v = Number(String(row[yearCol] ?? "").replace(/[^\d.-]/g, "")) || 0;
    values.push({ month: m, value: v }); // m → 1..12 гэж үзнэ
  }

  // 3) 0 биш сүүлийн N сар + шошго "… сар"
  const active = values.filter((x) => x.value > 0);
  const lastN = active.slice(-takeLast);
  return {
    year: yearLabel,
    points: lastN.map((x) => ({ label: `${x.month} сар`, value: x.value })),
    allMonths: values,
  };
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

function renderAssMonthlyLineCard(ass) {
  const labels = ass.points.map((p) => p.label);
  const data = ass.points.map((p) => p.value);

  // энгийн dataLabel plugin (ойн дээр тоонуудыг харуулна)
  const plugin = `
    const dataLabel = { id:'dataLabel', afterDatasetsDraw(chart){
      const {ctx, data:{datasets}, scales:{x,y}} = chart;
      ctx.save(); ctx.font='12px system-ui,-apple-system,Segoe UI,Roboto,Arial'; ctx.fillStyle='#111';
      (datasets[0].data||[]).forEach((v,i)=>{ 
        const xp = x.getPixelForValue(i), yp = y.getPixelForValue(v);
        ctx.fillText(String(v), xp+6, yp-6);
      });
      ctx.restore();
    }};
  `;

  return `
  <section class="card">
    <div class="card-title">НИЙТ ХАНДАЛТ /Сүүлийн ${labels.length} сараар/</div>
    <canvas id="assLine"></canvas>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
    <script>(function(){
      ${plugin}
      const ctx = document.getElementById('assLine').getContext('2d');
      new Chart(ctx, {
        type:'line',
        data:{ labels:${JSON.stringify(
          labels
        )}, datasets:[{ label:'', data:${JSON.stringify(
    data
  )}, tension:.3, pointRadius:4 }] },
        options:{ animation:false, plugins:{ legend:{display:false} }, scales:{ y:{beginAtZero:true} } },
        plugins:[dataLabel]
      });
    })();</script>
  </section>`;
}

const isWeekLike = (s) =>
  /(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)\s*[-–]\s*(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)/.test(
    String(s || "")
  );

function getLastWeekLabelFromASS(wb, sheetName = "APS") {
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });

  // week-like байгаа бүх баганын индексуудыг цуглуулаад хамгийн сүүлийг нь сонгоно
  const weekCols = new Set();
  rows.forEach((r) =>
    (r || []).forEach((v, i) => {
      if (isWeekLike(v)) weekCols.add(i);
    })
  );
  const indices = [...weekCols].sort((a, b) => a - b);
  if (!indices.length) return null;

  const lastCol = indices.at(-1);
  // тэр баганын аль нэг мөрөн дэх week-like текстийг авч буцаана
  for (let r = 0; r < rows.length; r++) {
    const v = rows[r]?.[lastCol];
    if (isWeekLike(v)) return String(v).trim();
  }
  return null;
}

function getSingleWeekLabelFromTotaly(wb, sheetName = "totaly") {
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  const header = rows[0] || [];
  for (let i = header.length - 1; i >= 0; i--) {
    const v = String(header[i] || "").trim();
    if (!v) continue;
    if (/өрчлөлт|эзлэх\s*хув/i.test(v)) continue;
    if (isWeekLike(v)) return v;
  }
  return null;
}

function extractOstTop10ForCompany(
  prevWb,
  currWb,
  {
    sheetName = CONFIG?.ASS_SHEET || "APS",
    company = "Ардын Тэтгэврийн Данс",
  } = {}
) {
  const norm = (s) => String(s || "").trim();

  const readOne = (wb) => {
    const ws = wb.Sheets[sheetName];
    if (!ws)
      throw new Error(
        `Sheet not found: ${sheetName}. Available: ${wb.SheetNames.join(", ")}`
      );
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
    const hdr = rows[0].map(norm);

    const idx = {
      company: hdr.findIndex((h) => /Компани/i.test(h)),
      category: hdr.findIndex((h) => /Ангилал/i.test(h)),
      subcat: hdr.findIndex((h) => /Туслах\s*ангилал/i.test(h)),
    };
    for (const k of Object.keys(idx))
      if (idx[k] < 0) {
        throw new Error(`[Ost-CompanyTop] column not found: ${k}`);
      }

    // тооллогыг гурван багцад хадгална
    const bag = {
      Лавлагаа: new Map(),
      Үйлчилгээ: new Map(),
      Гомдол: new Map(),
    };

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;
      if (norm(row[idx.company]) !== norm(company)) continue;

      const cat = norm(row[idx.category]);
      if (!(cat in bag)) continue;

      const sub = norm(row[idx.subcat]);
      if (!sub) continue;

      bag[cat].set(sub, (bag[cat].get(sub) || 0) + 1);
    }
    return bag;
  };

  const prev = readOne(prevWb);
  const curr = readOne(currWb);

  const labelPrev = getSingleWeekLabelFromTotaly(prevWb) || "Өмнөх 7 хоног";
  const labelCurr = getSingleWeekLabelFromTotaly(currWb) || "Одоогийн 7 хоног";

  const toTop10 = (cat) => {
    const names = new Set([...prev[cat].keys(), ...curr[cat].keys()]);
    const arr = [...names].map((name) => {
      const a = prev[cat].get(name) || 0;
      const b = curr[cat].get(name) || 0;
      const delta = a ? (b - a) / a : b ? 1 : 0;
      return { name, prev: a, curr: b, deltaPct: delta };
    });
    arr.sort((x, y) => y.curr - x.curr);
    return arr.slice(0, 10);
  };

  return {
    labels: [labelPrev, labelCurr],
    lavlagaa: toTop10("Лавлагаа"),
    uilchilgee: toTop10("Үйлчилгээ"),
    gomdol: toTop10("Гомдол"),
    company,
  };
}

// weekly = { labels:[...], lav:[...], uil:[...], gom:[...] }
function renderWeeklyByCategory(weekly) {
  const sumArr = (a, b, c) =>
    a.map((_, i) => (a[i] || 0) + (b[i] || 0) + (c[i] || 0));
  const totals = sumArr(weekly.lav || [], weekly.uil || [], weekly.gom || []);
  const lastIdx = weekly.labels.length - 1;
  const last = {
    lav: weekly.lav[lastIdx] || 0,
    uil: weekly.uil[lastIdx] || 0,
    gom: weekly.gom[lastIdx] || 0,
    total: totals[lastIdx] || 0,
  };
  const prevTotal = totals.length > 1 ? totals[lastIdx - 1] : 0;
  const delta = prevTotal ? (last.total - prevTotal) / prevTotal : 0;
  const pct = (x) => `${(x * 100).toFixed(0)}%`;

  // энгийн data label plugin (бүх dataset-д тоог харуулна)
  const plugin = `
    const dataLabel = { id:'dataLabel', afterDatasetsDraw(chart){
      const {ctx, data:{datasets}, scales:{x,y}} = chart;
      ctx.save();
      ctx.font = '12px system-ui,-apple-system,Segoe UI,Roboto,Arial';
      ctx.textAlign = 'center'; ctx.textBaseline = 'bottom';
      datasets.forEach((ds,di)=>{
        (ds.data||[]).forEach((v,i)=>{
          if(v==null) return;
          const meta = chart.getDatasetMeta(di);
          const pt = meta.data[i];
          const xPos = pt.x, yPos = pt.y - 4;
          ctx.fillStyle = '#111';
          ctx.fillText(String(v), xPos, yPos);
        });
      });
      ctx.restore();
    }};`;

  return `
  <section class="card">
    <div class="grid">
      <div>
        <div class="card-title">БҮРТГЭЛ /Ангиллаар/</div>
        <canvas id="byCategory" style="height: 500px;"></canvas>
      </div>
      <div>
        <ul style="margin-top:34px;line-height:1.6">
          <li>Сүүлчийн 7 хоногт нийт <b>${last.total.toLocaleString()}</b> (${pct(
    last.lav / Math.max(1, last.total)
  )} лавлагаа, ${pct(last.uil / Math.max(1, last.total))} үйлчилгээ, ${pct(
    last.gom / Math.max(1, last.total)
  )} гомдол).</li>
          <li>Өмнөх 7 хоногоос ${delta >= 0 ? "өссөн" : "буурсан"}: <b>${pct(
    Math.abs(delta)
  )}</b>.</li>
        </ul>
      </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
    <script>(function(){
      ${plugin}
      const ctx = document.getElementById('byCategory').getContext('2d');
      new Chart(ctx, {
        type:'bar',
        data:{
          labels:${JSON.stringify(weekly.labels)},
          datasets:[
            { label:'Лавлагаа',   data:${JSON.stringify(weekly.lav)} },
            { label:'Үйлчилгээ',  data:${JSON.stringify(weekly.uil)} },
            { label:'Гомдол',     data:${JSON.stringify(weekly.gom)} }
          ]
        },
        options:{
          animation:false,
          plugins:{ legend:{ position:'bottom' } },
          scales:{ y:{ beginAtZero:true } }
        },
        plugins:[dataLabel]
      });
    })();</script>
  </section>`;
}

// APS sheet → Weekly by category (Лавлагаа/Үйлчилгээ/Гомдол)
function extractWeeklyByCategoryFromASS(
  wb,
  { sheetName = "APS", takeLast = 4 } = {}
) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet not found: ${sheetName}`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });

  const weekLike = (s) =>
    /(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)\s*[-–]\s*(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)/.test(
      String(s || "")
    );

  // 1) Sheet дээрх бүх week-like баганыг цуглуулаад баруун талын сүүлийн N-г авна
  const weekColsSet = new Set();
  rows.forEach((r) =>
    (r || []).forEach((v, i) => {
      if (weekLike(v)) weekColsSet.add(i);
    })
  );
  const weekCols = [...weekColsSet].sort((a, b) => a - b);
  if (!weekCols.length) throw new Error("[APS] Week columns not found");
  const pickCols = weekCols.slice(-takeLast);

  // 2) “Лавлагаа/Үйлчилгээ/Гомдол” мөрүүдийг аливаа баганад тааруулж олно
  const findRowByRe = (re) => {
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r] || [];
      for (let c = 0; c < row.length; c++) {
        const cell = String(row[c] ?? "")
          .replace(/\s+/g, " ")
          .trim();
        if (re.test(cell)) return r; // мөрийн индекс
      }
    }
    return -1;
  };

  const rLav = findRowByRe(/^\s*Лавлагаа\s*$/i);
  const rUil = findRowByRe(/^\s*Үйлчилгээ\s*$/i);
  const rGom = findRowByRe(/^\s*Гомдол\s*$/i);

  if (rLav < 0 || rUil < 0 || rGom < 0) {
    throw new Error(
      "[APS] Row(s) not found for Лавлагаа/Үйлчилгээ/Гомдол — мөрийн нэр өөр эсвэл өөр хэлбэртэй бичигдсэн байж болзошгүй."
    );
  }

  const toNum = (x) => Number(String(x ?? "").replace(/[^\d.-]/g, "")) || 0;

  // 3) Labels болон утгуудыг бэлтгэх
  const labels = pickCols.map((c) => {
    for (let r = 0; r < rows.length; r++) {
      const v = rows[r]?.[c];
      if (weekLike(v)) return String(v).trim();
    }
    return String(rows[0]?.[c] || "").trim(); // fallback
  });

  const lav = pickCols.map((c) => toNum(rows[rLav]?.[c]));
  const uil = pickCols.map((c) => toNum(rows[rUil]?.[c]));
  const gom = pickCols.map((c) => toNum(rows[rGom]?.[c]));

  return { labels, lav, uil, gom };
}

function renderCompanyTop10Section(top) {
  const pct = (v) => `${Math.abs(v * 100).toFixed(0)}%`;
  const arrow = (v) => (v >= 0 ? "▲" : "▼");
  const cls = (v) => (v >= 0 ? "up" : "down");

  const makeTable = (rows) => `
    <table class="cmp">
      <thead>
        <tr><th>Туслах ангилал</th><th>${top.labels[0]}</th><th>${
    top.labels[1]
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
            <td class="num ${cls(r.deltaPct)}">${arrow(r.deltaPct)} ${pct(
              r.deltaPct
            )}</td>
          </tr>
        `
          )
          .join("")}
      </tbody>
    </table>`;

  const block = (title, rows, id) => `
    <div class="card soft">
      <div class="card-title">${escapeHtml(title)}</div>
      <div class="tables">
        <div class="card soft">${makeTable(rows)}</div>
      </div>
    </div>`;

  return `
  <section class="company-top10">
    ${block("ТОП ЛАВЛАГАА / туслах ангилалаар /", top.lavlagaa, "topLav")}
    ${block("ТОП ҮЙЛЧИЛГЭЭ / туслах ангилалаар /", top.uilchilgee, "topUil")}
    ${block("ТОП ГОМДОЛ / туслах ангилалаар /", top.gomdol, "topGom")}
  </section>`;
}

function extractApsMiniBarsFromAPS(wb, sheetName = "APS", takeLast = 2) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`[APS] Sheet not found: ${sheetName}`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });

  // week-like толгой (ж/нь: 07.28-08.03)
  const isWeekLike = (s) =>
    /(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)\s*[-–]\s*(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)/.test(
      String(s || "")
    );

  // sheet дээрх бүх багануудаас week-like гарчигтай багануудыг олж сүүлийнхоос нь N-г авах
  const weekColsSet = new Set();
  rows.forEach((r) =>
    (r || []).forEach((v, i) => {
      if (isWeekLike(v)) weekColsSet.add(i);
    })
  );
  const weekCols = [...weekColsSet].sort((a, b) => a - b);
  if (weekCols.length < takeLast)
    throw new Error(`[APS] Not enough week columns (found ${weekCols.length})`);
  const pickCols = weekCols.slice(-takeLast);

  // мөрийн индексийг нэрээр нь ерөнхийд нь хайна (хаана ч байсан таарна)
  const findRowBy = (re) => {
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r] || [];
      for (let c = 0; c < row.length; c++) {
        const cell = String(row[c] ?? "")
          .replace(/\s+/g, " ")
          .trim();
        if (re.test(cell)) return r;
      }
    }
    return -1;
  };

  const rLav = findRowBy(/^\s*Лавлагаа\s*$/i);
  const rUil = findRowBy(/^\s*Үйлчилгээ\s*$/i);
  const rGom = findRowBy(/^\s*Гомдол\s*$/i);
  if (rLav < 0 || rUil < 0 || rGom < 0)
    throw new Error(
      "[APS] Row(s) not found for Лавлагаа/Үйлчилгээ/Гомдол (sheet дээрх нэрийг шалгана уу)."
    );

  const toNum = (x) => Number(String(x ?? "").replace(/[^\d.-]/g, "")) || 0;

  // labels – pickCols баганад бичигдсэн week шошгыг өөрөөс нь олж авна
  const labels = pickCols.map((c) => {
    for (let r = 0; r < rows.length; r++) {
      const v = rows[r]?.[c];
      if (isWeekLike(v)) return String(v).trim();
    }
    return String(rows[0]?.[c] || "").trim();
  });

  return {
    labels,
    lav: pickCols.map((c) => toNum(rows[rLav]?.[c])),
    uil: pickCols.map((c) => toNum(rows[rUil]?.[c])),
    gom: pickCols.map((c) => toNum(rows[rGom]?.[c])),
  };
}

function renderApsMiniBarsWithTable(mini, rightHtml = "") {
  // delta% (prev->last)
  const deltaPct = (arr) => {
    const n = arr.length;
    const prev = n > 1 ? arr[n - 2] || 0 : 0;
    const last = n ? arr[n - 1] || 0 : 0;
    const d = prev ? (last - prev) / prev : last ? 1 : 0;
    return { prev, last, d };
  };

  const fmtPct = (x) => `${(Math.abs(x) * 100).toFixed(0)}%`;
  const arrow = (x) => (x >= 0 ? "▲" : "▼");
  const badgeCls = (x) => (x >= 0 ? "bg-success" : "bg-danger");

  const lav = deltaPct(mini.lav);
  const uil = deltaPct(mini.uil);
  const gom = deltaPct(mini.gom);

  // datalabels
  const plugin = `
    const dataLabel={id:'dataLabel',afterDatasetsDraw(ch){
      const {ctx,getDatasetMeta,data:{datasets}}=ch; ctx.save();
      ctx.font='12px system-ui,-apple-system,Segoe UI,Roboto,Arial';
      ctx.textAlign='center'; ctx.textBaseline='bottom';
      datasets.forEach((ds,di)=>{const m=getDatasetMeta(di);
        (ds.data||[]).forEach((v,i)=>{if(v==null)return;const pt=m.data[i];
          ctx.fillStyle='#111'; ctx.fillText(String(v), pt.x, pt.y-6);
        });
      }); ctx.restore();
    }};`;

  const card = (id, title, arr, meta) => `
    <div class="card shadow-sm mb-3">
      <div class="card-body">
        <div class="d-flex justify-content-between align-items-center mb-2">
          <h6 class="m-0">${title}</h6>
          <span class="badge ${badgeCls(meta.d)}">
            ${arrow(meta.d)} ${fmtPct(meta.d)}
          </span>
        </div>
        <div class="text-muted small mb-2">
          ${mini.labels.at(-2) || ""}: <b>${meta.prev}</b> &nbsp;→&nbsp;
          ${mini.labels.at(-1) || ""}: <b>${meta.last}</b>
        </div>
        <canvas id="${id}" height="140"></canvas>
      </div>
    </div>
    <script>
      (function(){
        ${plugin}
        const ctx=document.getElementById('${id}').getContext('2d');
        new Chart(ctx,{
          type:'bar',
          data:{ labels:${JSON.stringify(mini.labels)},
                 datasets:[{label:'', data:${JSON.stringify(
                   arr
                 )}, borderWidth:0 }]},
          options:{ animation:false, plugins:{ legend:{display:false} },
                    scales:{ y:{ beginAtZero:true } } },
          plugins:[dataLabel]
        });
      })();
    </script>
  `;

  return `
  <div class="container-fluid">
    <div class="row g-3">
      <div class="col-lg-6">
        ${card("apsLavMini", "ЛАВЛАГАА", mini.lav, lav)}
        ${card("apsUilMini", "ҮЙЛЧИЛГЭЭ", mini.uil, uil)}
        ${card("apsGomMini", "ГОМДОЛ", mini.gom, gom)}
      </div>
      <div class="col-lg-6">
        ${rightHtml}
      </div>
    </div>
  </div>`;
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
    requireTLS: process.env.SMTP_PORT === "587",
    tls: { minVersion: "TLSv1.2" },
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

function wrapHtml(bodyHtml) {
  const css = fs.readFileSync(CONFIG.CSS_FILE, "utf-8");
  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  
  <!-- Bootstrap 5 CSS -->
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css"
        rel="stylesheet"
        integrity="sha384-QWTKZyjpPEjISv5WaRU9OFeRpok6YctnYmDr5pNlyT2bRjXh0JMhjY6hW+ALEwIH"
        crossorigin="anonymous">

  <style>${css}</style>
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
// MAIN
// ────────────────────────────────────────────────────────────────
async function runOnce() {
  // файлуудаа шалгана
  [
    CONFIG.CURR_FILE,
    CONFIG.PREV_FILE,
    CONFIG.CSS_FILE,
    CONFIG.GOMDOL_FILE,
  ].forEach((p) => {
    if (!fs.existsSync(p)) throw new Error(`Missing file: ${p}`);
  });
  if (!fs.existsSync(CONFIG.OUT_DIR))
    fs.mkdirSync(CONFIG.OUT_DIR, { recursive: true });

  const wbCurr = xlsx.readFile(CONFIG.CURR_FILE, { cellDates: true });
  const wbPrev = xlsx.readFile(CONFIG.PREV_FILE, { cellDates: true });

  const weekly = extractWeeklyByCategoryFromASS(wbCurr, {
    sheetName: "APS",
    takeLast: 4,
  });

  const ass = extractAPSLeftBlockLastMonths(wbCurr, {
    sheetName: "APS",
    yearLabel: "2025",
    takeLast: 4,
  });

  const topAA = extractOstTop10ForCompany(wbPrev, wbCurr, {
    sheetName: CONFIG.OST_SHEET,
    company: "Ардын Тэтгэврийн Данс",
  });

  const lastWeekLabel =
    getLastWeekLabelFromASS(wbCurr, CONFIG.ASS_SHEET) ||
    getSingleWeekLabelFromTotaly(wbCurr, "totaly") ||
    "";

  const cover = renderAssCover({
    company: CONFIG.ASS_COMPANY,
    periodText: lastWeekLabel ? `Долоо хоног: ${lastWeekLabel}` : "",
  });

  const assChart = renderAssMonthlyLineCard(ass);
  const apsMini = extractApsMiniBarsFromAPS(wbCurr, "APS", 2);
  const rightSide = renderCompanyTop10Section(topAA);
  const outbound = extractOutbound_AARD_HVL(
    CONFIG.OUTBOUND_FILE,
    CONFIG.OUTBOUND_SHEET
  );

  // HTML sections
  let body = "";
  body += cover;
  body += assChart;
  body += renderWeeklyByCategory(weekly);
  body += renderApsMiniBarsWithTable(apsMini, rightSide);
  body += renderOutboundTable(outbound);

  const html = wrapHtml(body);

  // PDF
  const monday = dayjs().tz(CONFIG.TIMEZONE).startOf("week").add(1, "day");
  const stamp = monday.format("YYYYMMDD");
  const pdfName = `ard-tetgever-${monday.format("YYYYMMDD")}.pdf`;
  const pdfPath = path.join(CONFIG.OUT_DIR, pdfName);
  // await htmlToPdf(html, pdfPath);
  await htmlToPdf(html, pdfPath);

  if (CONFIG.SAVE_HTML) {
    const htmlPath = path.join(
      CONFIG.OUT_DIR,
      `${CONFIG.HTML_NAME_PREFIX}-${stamp}.html`
    );
    fs.writeFileSync(htmlPath, html, "utf8");
    console.log(`[OK] HTML saved → ${htmlPath}`);
  }

  // Email
  const subject = `${CONFIG.SUBJECT_PREFIX} ${
    CONFIG.REPORT_TITLE
  } — ${monday.format("YYYY-MM-DD")}`;
  await sendEmailWithPdf(pdfPath, subject);

  console.log(`[OK] Sent ${pdfName} → ${process.env.RECIPIENTS}`);
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
const runNow = process.argv.includes("--once");
if (runNow) {
  runOnce().catch((e) => {
    console.error(e);
    process.exit(1);
  });
} else {
  startScheduler();
}
