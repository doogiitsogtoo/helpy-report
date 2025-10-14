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
  ASS_SHEET: "ASS",
  ASS_COMPANY: "Ард Актив",
  ASS_YEAR: "2025",
  ASS_TAKE_LAST_N_MONTHS: 4,
  OST_SHEET: "Osticket1",

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

function escapeHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function extractAssCompanyLatestMonths(
  wb,
  {
    sheetName = CONFIG.ASS_SHEET,
    company = CONFIG.ASS_COMPANY,
    yearLabel = CONFIG.ASS_YEAR,
    takeLast = CONFIG.ASS_TAKE_LAST_N_MONTHS,
  } = {}
) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`Sheet not found: ${sheetName}`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });

  // 1) "Ард Актив" гэсэн гарчгийг олно (ихэвчлэн зүүн дээд талд)
  const companyRow = findRowIndex(rows, (r) =>
    r.some((v) => _normTxt(v) === _normTxt(company))
  );
  if (companyRow < 0) throw new Error(`[ASS] Company not found: ${company}`);

  // 2) Жилийн толгой (2021, 2022 ... 2025) агуулсан мөрийг компанийн хэсгийн доорх эхний мөрүүдээс хайж олно
  const yearHeadRow = findRowIndex(
    rows.slice(companyRow, companyRow + 8),
    (r) => r.filter(Boolean).some((v) => /^\d{4}$/.test(String(v)))
  );
  if (yearHeadRow < 0)
    throw new Error("[ASS] Year header row not found near company block");
  const absYearHeadRow = companyRow + yearHeadRow;

  // 3) Сонгосон жилийн баганын индекс
  const header = rows[absYearHeadRow] || [];
  const yearCol = header.findIndex(
    (v) => String(v).trim() === String(yearLabel)
  );
  if (yearCol < 0) throw new Error(`[ASS] Year column not found: ${yearLabel}`);

  // 4) Сарын мөрүүд (1 сар..12 сар) — компанийн блокийн доод хэсгээс цуглуулна
  //    Зураг дээр сарууд B баганад “1 cap/сар, 2 сар ... 12 сар” хэлбэртэй.
  const monthRows = [];
  for (let r = absYearHeadRow + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const cell = String(row[1] ?? row[0] ?? "").trim(); // B эсвэл A
    if (/^\d+\s*сар$/i.test(cell) || /^\d+\s*cap$/i.test(cell)) {
      const label = cell.replace(/cap/i, "сар");
      const val =
        Number(String(row[yearCol] ?? "").replace(/[^\d.-]/g, "")) || 0;
      monthRows.push({ label, value: val });
      continue;
    }
    // блок дуусах нөхцөл – хоосон хуудас/дараагийн хэсэг эхлэхэд зогсоно
    if (monthRows.length && !cell) break;
  }

  // 5) 0 биш (идэвхтэй) сарын сүүлийн N-г авна
  const active = monthRows.filter((m) => m.value > 0);
  const lastN = active.slice(-takeLast);

  return {
    company,
    year: yearLabel,
    points: lastN, // [{label:'5 сар', value:611}, ...]
    allMonths: monthRows, // хүсвэл бүрэн жагсаалт
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
  const h = 200;

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
    <canvas id="assLine" height="${h}" style="height: 500px !important; display: flex; align-items: center; justify-content: center; margin:4rem 0;"></canvas>
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

function getSingleWeekLabelFromTotaly(wb, sheetName = "totaly") {
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  const header = rows[0] || [];
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

function extractOstTop10ForCompany(
  prevWb,
  currWb,
  { sheetName = CONFIG?.ASS_SHEET || "ASS", company = "Ард Актив" } = {}
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
    <div class="grid grid-2">
      <div>
        <div class="card-title">БҮРТГЭЛ /Ангиллаар/</div>
        <canvas id="byCategory" style="margin:4rem 0; height: 500px;"></canvas>
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

// ASS sheet → Weekly by category (Лавлагаа/Үйлчилгээ/Гомдол)
function extractWeeklyByCategoryFromASS(
  wb,
  { sheetName = "ASS", takeLast = 4 } = {}
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
  if (!weekCols.length) throw new Error("[ASS] Week columns not found");
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
      "[ASS] Row(s) not found for Лавлагаа/Үйлчилгээ/Гомдол — мөрийн нэр өөр эсвэл өөр хэлбэртэй бичигдсэн байж болзошгүй."
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
    <div class="card">
      <div class="card-title">БҮРТГЭЛ / Ангилал тус бүрээр / — ${escapeHtml(
        top.company
      )}</div>
      <p class="kpi-note">Харьцуулалт: <b>${top.labels[0]}</b> ↔ <b>${
    top.labels[1]
  }</b>.</p>
    </div>
    ${block("ТОП ЛАВЛАГАА / туслах ангилалаар /", top.lavlagaa, "topLav")}
    ${block("ТОП ҮЙЛЧИЛГЭЭ / туслах ангилалаар /", top.uilchilgee, "topUil")}
    ${block("ТОП ГОМДОЛ / туслах ангилалаар /", top.gomdol, "topGom")}
  </section>`;
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
    sheetName: "ASS",
    takeLast: 4,
  });

  const ass = extractAssCompanyLatestMonths(wbCurr, {
    sheetName: CONFIG.ASS_SHEET,
    company: CONFIG.ASS_COMPANY,
    yearLabel: CONFIG.ASS_YEAR,
    takeLast: CONFIG.ASS_TAKE_LAST_N_MONTHS,
  });

  const topAA = extractOstTop10ForCompany(wbPrev, wbCurr, {
    sheetName: CONFIG.OST_SHEET,
    company: "Ард Актив",
  });

  const cover = renderAssCover({
    company: CONFIG.ASS_COMPANY,
    periodText: `${ass.points[0]?.label ?? ""} – ${
      ass.points.at(-1)?.label ?? ""
    } (${ass.year})`,
  });
  const assChart = renderAssMonthlyLineCard(ass);

  // HTML sections
  let body = "";
  body += cover;
  body += assChart;
  body += renderWeeklyByCategory(weekly);
  body += renderCompanyTop10Section(topAA);

  const html = wrapHtml(body);

  // PDF
  const monday = dayjs().tz(CONFIG.TIMEZONE).startOf("week").add(1, "day");
  const stamp = monday.format("YYYYMMDD");
  const pdfName = `ard-aktiv-${monday.format("YYYYMMDD")}.pdf`;
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
