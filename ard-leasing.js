// ard-leasing.js — ALS (months) + totaly (weekly lav+gom) + Osticket1 Top10 (Lav+Gom) → HTML → PDF → Email
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
  CURR_FILE: "./ARD 10.13-10.19.xlsx",
  PREV_FILE: "./ARD 10.06-10.12.xlsx",
  GOMDOL_FILE: "./gomdol-weekly.xlsx", // weekly Гомдол override (хэрэглэж болно)

  // Sheets
  ASS_SHEET: "ALS",
  ASS_COMPANY: "Ард Лизинг",
  ASS_YEAR: "2025",
  ASS_TAKE_LAST_N_MONTHS: 4,
  OST_SHEET: "Osticket1",

  // PDF / Email
  OUT_DIR: "./out",
  CSS_FILE: "./css/template.css",
  SAVE_HTML: true,
  HTML_NAME_PREFIX: "report",
  REPORT_TITLE: process.env.REPORT_TITLE || "Ард Лизинг — 7 хоногийн тайлан",
  SUBJECT_PREFIX: process.env.SUBJECT_PREFIX || "[ArdLeasing Weekly]",
  EMAIL_ENABLED: String(process.env.EMAIL_ENABLED ?? "true") === "true",
  SCHED_ENABLED: String(process.env.SCHED_ENABLED ?? "true") === "true",
};

// ────────────────────────────────────────────────────────────────
// Utils
// ────────────────────────────────────────────────────────────────
const esc = (s) =>
  String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
const nnum = (v) => Number(String(v ?? "").replace(/[^\d.-]/g, "")) || 0;

// ────────────────────────────────────────────────────────────────
// Cover
// ────────────────────────────────────────────────────────────────
function renderAssCover({ company, periodText }) {
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
        <div style="font-size:36px;font-weight:800;line-height:1.1">${esc(
          company
        )}</div>
        <div style="opacity:.9;margin-top:8px">${esc(periodText || "")}</div>
      </div>
    </div>
  </section>`;
}

// ────────────────────────────────────────────────────────────────
// ALS → Сүүлийн N сар (months)
// ────────────────────────────────────────────────────────────────
function extractALSLatestMonths(
  wb,
  { sheetName = "ALS", yearLabel = "2025", takeLast = 4 } = {}
) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`[ALS] Sheet not found: ${sheetName}`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });

  // 4 оронт жилүүдтэй толгой мөр
  const yearHeadRowIdx = rows.findIndex(
    (r) => r && r.filter(Boolean).some((v) => /^\d{4}$/.test(String(v)))
  );
  if (yearHeadRowIdx < 0) throw new Error("[ALS] Year header row not found");

  const header = (rows[yearHeadRowIdx] || []).map((v) =>
    String(v || "").trim()
  );
  const yearCol = header.findIndex((v) => v === String(yearLabel));
  if (yearCol < 0) throw new Error(`[ALS] Year col not found: ${yearLabel}`);

  const monthLike = (s) =>
    /^\s*\d+\s*(сар|cap)\s*$/i.test(String(s || "").trim());
  const pickLabel = (row) => {
    for (const c of [0, 1, 2])
      if (monthLike(row?.[c])) return String(row[c]).replace(/cap/i, "сар");
    return null;
  };

  const monthRows = [];
  for (let r = yearHeadRowIdx + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const label = pickLabel(row);
    if (!label) {
      if (monthRows.length) break;
      continue;
    }
    const value = nnum(row[yearCol]);
    monthRows.push({ label, value });
  }

  const active = monthRows.filter((m) => m.value > 0);
  const points = active.slice(-takeLast);
  return { year: yearLabel, points, allMonths: monthRows };
}

// ────────────────────────────────────────────────────────────────
// totaly → 7 хоног (Лавлагаа + Гомдол) — Сүүлийн 4 долоо хоног
// ────────────────────────────────────────────────────────────────
function extractWeeklyLavGomFromTotaly(wb, sheetName = "totaly") {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`[totaly] Sheet not found: ${sheetName}`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  const head = rows[0] || [];

  const weekLike = (s) =>
    /(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)\s*[-–]\s*(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)/.test(
      String(s || "")
    );

  const idx = [];
  for (let i = head.length - 1; i >= 0 && idx.length < 4; i--) {
    const v = head[i];
    if (v && weekLike(v)) idx.push(i);
  }
  idx.reverse();
  const labels = idx.map((i) => String(head[i]).trim());

  const findRow = (re) => rows.find((r) => r && r[0] && re.test(String(r[0])));
  const toNum = (r, i) => nnum(r?.[i]);

  const rLav = findRow(/^Лавлагаа$/i);
  const rGom = findRow(/^Гомдол$/i);

  if (!rLav || !rGom)
    throw new Error("[totaly] Row(s) not found: Лавлагаа/Гомдол");

  return {
    labels,
    lav: idx.map((i) => toNum(rLav, i)),
    gom: idx.map((i) => toNum(rGom, i)),
  };
}

// ────────────────────────────────────────────────────────────────
// Гомдлын workbook → 7 хоногийн “Гомдол” override (хүсвэл)
// ────────────────────────────────────────────────────────────────
function extractWeeklyGomFromGomdol(file) {
  const wb = xlsx.readFile(file, { cellDates: true });
  const pick = (names) => names.find((n) => wb.SheetNames.includes(n));
  const target =
    pick(["WEEK", "Comp", "Osticket ", "Osticket", "Gomdol"]) ||
    wb.SheetNames[0];

  const ws = wb.Sheets[target];
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });

  const head = rows[0] || [];
  const weekLike = (s) =>
    /(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)\s*[-–]\s*(\д{1,2}[./-]\д{1,2}(?:[./-]\д{2,4})?)/.test(
      String(s || "")
    );
  const idx = [];
  for (let i = head.length - 1; i >= 0 && idx.length < 4; i--) {
    if (head[i] && weekLike(head[i])) idx.push(i);
  }
  idx.reverse();
  if (!idx.length) return null;

  const findRow = (re) =>
    rows.find((r) => r && r.some((v) => re.test(String(v || ""))));
  const rG = findRow(/^\s*Гомдол\s*$/i);
  if (!rG) return null;

  const labels = idx.map((i) => String(head[i]).trim());
  const gom = idx.map((i) => nnum(rG?.[i]));
  return { labels, gom };
}

// ────────────────────────────────────────────────────────────────
// Osticket1 → ТОП-10 (prev vs curr) “Лавлагаа” ба “Гомдол”
// ────────────────────────────────────────────────────────────────
function getSingleWeekLabelFromTotaly(wb, sheetName = "totaly") {
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  const header = rows[0] || [];
  const isWeekLike = (s) =>
    /(\d{1,2}[./-]\d{1,2}(?:[./-]\д{2,4})?)\s*[-–]\s*(\д{1,2}[./-]\д{1,2}(?:[./-]\д{2,4})?)/.test(
      String(s || "")
    );
  for (let i = header.length - 1; i >= 0; i--) {
    const v = String(header[i] || "").trim();
    if (!v) continue;
    if (/өрчлөлт|эзлэх\s*хув/i.test(v)) continue;
    if (isWeekLike(v)) return v;
  }
  return null;
}
function getLastWeekLikeFromHeader(header) {
  const isWeekLike = (s) =>
    /(\d{1,2}[./-]\д{1,2}(?:[./-]\д{2,4})?)\s*[-–]\s*(\д{1,2}[./-]\д{1,2}(?:[./-]\д{2,4})?)/.test(
      String(s || "")
    );
  for (let i = header.length - 1; i >= 0; i--) {
    if (isWeekLike(header[i])) return String(header[i]).trim();
  }
  return null;
}
function extractOstTop10_ALS(
  prevWb,
  currWb,
  { sheetName = "Osticket1", company = "Ард Лизинг" } = {}
) {
  const read = (wb) => {
    const ws = wb.Sheets[sheetName];
    if (!ws) throw new Error(`[Osticket1] Sheet not found: ${sheetName}`);
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
    const hdr = rows[0].map((v) =>
      String(v || "")
        .trim()
        .toLowerCase()
    );
    const idx = {
      comp: hdr.findIndex((h) => /компани/i.test(h)),
      cat: hdr.findIndex((h) => /^ангилал$/i.test(h)),
      sub: hdr.findIndex((h) => /туслах\s*ангилал/i.test(h)),
    };
    if (idx.comp < 0 || idx.cat < 0 || idx.sub < 0)
      throw new Error("[Ost] missing columns");

    const bag = { Лавлагаа: new Map(), Гомдол: new Map() };
    const T = (s) => String(s || "").trim();

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;
      if (T(row[idx.comp]) !== company) continue;
      const cat = T(row[idx.cat]);
      const sub = T(row[idx.sub]);
      if (!sub || !bag[cat]) continue;
      bag[cat].set(sub, (bag[cat].get(sub) || 0) + 1);
    }

    const top = (k) => {
      const arr = [...bag[k].entries()].map(([name, val]) => ({
        name,
        curr: val,
      }));
      arr.sort((a, b) => b.curr - a.curr);
      return arr.slice(0, 10);
    };
    return { lav: top("Лавлагаа"), gom: top("Гомдол") };
  };

  const prev = read(prevWb),
    curr = read(currWb);
  const join = (p, c) => {
    const names = new Set([...p.map((x) => x.name), ...c.map((x) => x.name)]);
    const MP = new Map(p.map((x) => [x.name, x.curr]));
    const MC = new Map(c.map((x) => [x.name, x.curr]));
    const out = [...names].map((n) => {
      const a = MP.get(n) || 0,
        b = MC.get(n) || 0;
      const d = a ? (b - a) / a : b ? 1 : 0;
      return { name: n, prev: a, curr: b, delta: d };
    });
    out.sort((x, y) => y.curr - x.curr);
    return out.slice(0, 10);
  };

  const prevLabel =
    getSingleWeekLabelFromTotaly(prevWb, "totaly") ||
    getLastWeekLikeFromHeader(
      xlsx.utils.sheet_to_json(prevWb.Sheets["Osticket1"], {
        header: 1,
        raw: false,
      })[0] || []
    ) ||
    "Өмнөх 7 хоног";

  const currLabel =
    getSingleWeekLabelFromTotaly(currWb, "totaly") ||
    getLastWeekLikeFromHeader(
      xlsx.utils.sheet_to_json(currWb.Sheets["Osticket1"], {
        header: 1,
        raw: false,
      })[0] || []
    ) ||
    "Одоогийн 7 хоног";

  return {
    labels: [prevLabel, currLabel],
    lav: join(prev.lav, curr.lav),
    gom: join(prev.gom, curr.gom),
  };
}

// ────────────────────────────────────────────────────────────────
// HTML Layout — plugin-гүй, тоог canvas дээр шууд зурах
// ────────────────────────────────────────────────────────────────
function renderLeasingLayout({ monthsALS, weekly, top10 }) {
  const totals = weekly.labels.map(
    (_, i) => (weekly.lav[i] || 0) + (weekly.gom[i] || 0)
  );
  const last = totals.at(-1) || 0;
  const prev = totals.length > 1 ? totals.at(-2) || 0 : 0;
  const deltaPct = prev ? ((last - prev) / prev) * 100 : 0;

  // Chart render болсны дараа тоонуудыг гар бичгээр зурна (plugin биш)
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
    }
  `;

  const monthsLabels = (monthsALS?.points || []).map((p) => p.label);
  const monthsData = (monthsALS?.points || []).map((p) => p.value);

  const lineCard = `
  <div class="card" style="height: 500px; margin-bottom: 4rem;">
    <div class="card-title">НИЙТ ХАНДАЛТ /Сүүлийн ${
      monthsLabels.length
    } сараар/</div>
    <canvas id="alsLine"></canvas>
  </div>
  <script>(function(){
    ${drawValuesFn}
    const ctx = document.getElementById('alsLine').getContext('2d');
    const ch  = new Chart(ctx, {
      type: 'line',
      data: { labels: ${JSON.stringify(monthsLabels)},
              datasets: [{ label:'', data: ${JSON.stringify(
                monthsData
              )}, tension:.3, pointRadius:4 }] },
      options: { animation:false, plugins:{legend:{display:false}}, scales:{ y:{beginAtZero:true} } }
    });
    setTimeout(() => drawValues(ch), 0);
  })();</script>`;

  const weeklyCard = `
  <div class="card">
    <div class="card-title">НИЙТ БҮРТГЭЛ /Сүүлийн ${
      weekly.labels.length
    } долоо хоног/</div>
    <div class="grid">
      <canvas id="weekly"></canvas>
      <ul style="margin:10px 0 0 18px;line-height:1.6">
        <li>Тайлант 7 хоногт нийт <b>${last.toLocaleString()}</b>.</li>
        <li>Өмнөх 7 хоногоос <b>${
          deltaPct >= 0 ? "өссөн" : "буурсан"
        }</b>: <b>${Math.abs(deltaPct).toFixed(0)}%</b>.</li>
      </ul>
    </div>
  </div>
  <script>(function(){
    ${drawValuesFn}
    const ctx = document.getElementById('weekly').getContext('2d');
    const ch  = new Chart(ctx, {
      type: 'bar',
      data: { labels: ${JSON.stringify(weekly.labels)},
              datasets: [
                { label:'Лавлагаа', data: ${JSON.stringify(weekly.lav)} },
                { label:'Гомдол',   data: ${JSON.stringify(weekly.gom)} }
              ] },
      options: { animation:false, plugins:{legend:{position:'bottom'}}, scales:{ y:{beginAtZero:true} } }
    });
    setTimeout(() => drawValues(ch), 0);
  })();</script>`;

  const topTable = (title, rows, labels) => `
  <div class="card">
    <div class="card-title">${title}</div>
    <table class="cmp">
      <thead>
        <tr><th>Туслах ангилал</th><th>${labels[0] || ""}</th><th>${
    labels[1] || ""
  }</th><th>Өөрчлөлт</th></tr>
      </thead>
      <tbody>
        ${(rows || [])
          .map(
            (r) => `
          <tr>
            <td>${esc(r.name)}</td>
            <td class="num">${(r.prev || 0).toLocaleString()}</td>
            <td class="num">${(r.curr || 0).toLocaleString()}</td>
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
    ${lineCard}
    ${weeklyCard}
  </div>
  <div class="grid grid-1" style="margin-top:8px">
    ${topTable("ТОП Лавлагаа", top10.lav, top10.labels)}
    ${topTable("ТОП Гомдол", top10.gom, top10.labels)}
  </div>
</section>`;
}

// ────────────────────────────────────────────────────────────────
// Shell HTML — Chart.js-г глобалаар нэг удаа include хийнэ
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
  <style>${css}
    .card{background:#fff;border:1px solid #e5e7eb;border-radius:10px;padding:12px;margin-bottom:12px;page-break-inside:avoid}
    .card-title{font-weight:700;margin-bottom:6px}
    table.cmp{table-layout:fixed;width:100%;border-collapse:collapse;font-size:12.5px}
    table.cmp th,table.cmp td{border:1px solid #e5e7eb;padding:6px 8px;vertical-align:top}
    table.cmp thead th{background:#ffe5d5}
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

    // Chart.js бэлэн, дор хаяж 2 canvas (line + bar) бий эсэх
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
  [CONFIG.CURR_FILE, CONFIG.PREV_FILE, CONFIG.CSS_FILE].forEach((p) => {
    if (!fs.existsSync(p)) throw new Error(`Missing file: ${p}`);
  });
  if (!fs.existsSync(CONFIG.OUT_DIR))
    fs.mkdirSync(CONFIG.OUT_DIR, { recursive: true });

  const wbCurr = xlsx.readFile(CONFIG.CURR_FILE, { cellDates: true });
  const wbPrev = xlsx.readFile(CONFIG.PREV_FILE, { cellDates: true });

  // weekly (лав + гом) totaly-оос
  const weekly_base = extractWeeklyLavGomFromTotaly(wbCurr, "totaly");

  // Гомдол override (байвал, ба шошгоны урт таарвал)
  let weekly = { ...weekly_base };
  try {
    if (fs.existsSync(CONFIG.GOMDOL_FILE)) {
      const g = extractWeeklyGomFromGomdol(CONFIG.GOMDOL_FILE);
      if (g && g.labels.length === weekly_base.labels.length) {
        weekly = {
          labels: weekly_base.labels,
          lav: weekly_base.lav,
          gom: g.gom,
        };
      }
    }
  } catch (e) {
    console.warn("[GOMDOL] override skipped:", e?.message || e);
  }

  // ТОП-10: зөвхөн Лавлагаа, Гомдол
  const top10 = extractOstTop10_ALS(wbPrev, wbCurr, {
    sheetName: CONFIG.OST_SHEET,
    company: CONFIG.ASS_COMPANY,
  });

  // Сарын график
  const alsMonths = extractALSLatestMonths(wbCurr, {
    sheetName: CONFIG.ASS_SHEET,
    yearLabel: CONFIG.ASS_YEAR,
    takeLast: CONFIG.ASS_TAKE_LAST_N_MONTHS,
  });

  const cover = renderAssCover({
    company: CONFIG.ASS_COMPANY.toUpperCase(),
    periodText: `${weekly.labels[0] ?? ""} – ${weekly.labels.at(-1) ?? ""}`,
  });

  const body =
    cover + renderLeasingLayout({ monthsALS: alsMonths, weekly, top10 });
  const html = wrapHtml(body);

  // PDF
  const monday = dayjs().tz(CONFIG.TIMEZONE).startOf("week").add(1, "day");
  const stamp = monday.format("YYYYMMDD");
  const pdfName = `ard-leasing-${monday.format("YYYYMMDD")}.pdf`;
  const pdfPath = path.join(CONFIG.OUT_DIR, pdfName);

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
if (process.argv.includes("--once")) {
  runOnce().catch((e) => {
    console.error(e);
    process.exit(1);
  });
} else {
  startScheduler();
}
