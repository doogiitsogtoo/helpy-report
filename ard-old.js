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

function escapeHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

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

function extractAPSLatestMonths(
  wb,
  { sheetName = "APS", yearLabel = "2025", takeLast = 4 } = {}
) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`[APS] Sheet not found: ${sheetName}`);

  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });

  // 1) Жилийн толгой мөрийг олох
  const yearHeadRowIdx = rows.findIndex(
    (r) => r && r.filter(Boolean).some((v) => /^\d{4}$/.test(String(v)))
  );
  if (yearHeadRowIdx < 0) throw new Error("[APS] Year header row not found");

  const header = (rows[yearHeadRowIdx] || []).map((v) =>
    String(v || "").trim()
  );
  const yearCol = header.findIndex((v) => v === String(yearLabel));
  if (yearCol < 0) throw new Error(`[APS] Year col not found: ${yearLabel}`);

  // 2) Сар (A/B/C баганаас аль ч талд байж болох)
  const monthLike = (s) =>
    /^\s*\d+\s*(сар|cap)\s*$/i.test(String(s || "").trim());
  const pickLabel = (row) => {
    for (const c of [0, 1, 2]) {
      if (monthLike(row[c])) return String(row[c]).replace(/cap/i, "сар");
    }
    return null;
  };

  const monthRows = [];
  for (let r = yearHeadRowIdx + 1; r < rows.length; r++) {
    const row = rows[r] || [];
    const label = pickLabel(row);
    if (!label) {
      if (monthRows.length) break; // сарын блок дууссан
      continue;
    }
    const value =
      Number(String(row[yearCol] ?? "").replace(/[^\d.-]/g, "")) || 0;
    monthRows.push({ label, value });
  }

  const active = monthRows.filter((m) => m.value > 0);
  const points = active.slice(-takeLast);

  return { year: yearLabel, points, allMonths: monthRows };
}

function extractWeeklyByCategoryFromTotaly(wb, sheetName = "totaly") {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`[totaly] Sheet not found: ${sheetName}`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  const head = rows[0] || [];

  // week-like толгойнуудын индекс (баруунаас 4)
  const weekLike = (s) =>
    /(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)\s*[-–]\s*(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)/.test(
      String(s || "")
    );
  const indexes = [];
  for (let i = head.length - 1; i >= 0 && indexes.length < 4; i--) {
    const v = head[i];
    if (v && weekLike(v)) indexes.push(i);
  }
  indexes.reverse(); // хуучнаас шинэ рүү
  const labels = indexes.map((i) => String(head[i]).trim());

  const findRow = (re) => rows.find((r) => r && r[0] && re.test(String(r[0])));
  const num = (r, i) =>
    Number(String(r?.[i] ?? "").replace(/[^\d.-]/g, "")) || 0;

  const rLav = findRow(/^Лавлагаа$/i);
  const rUil = findRow(/^Үйлчилгээ$/i);
  const rGom = findRow(/^Гомдол$/i);

  if (!rLav || !rUil || !rGom)
    throw new Error("[ASS] Row(s) not found: Лавлагаа/Үйлчилгээ/Гомдол");

  return {
    labels,
    lav: indexes.map((i) => num(rLav, i)),
    uil: indexes.map((i) => num(rUil, i)),
    gom: indexes.map((i) => num(rGom, i)),
  };
}

function extractOutbound3Weeks(wb, sheetName = "totaly") {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`[totaly] Sheet not found: ${sheetName}`);
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  const head = rows[0] || [];
  const weekLike = (s) =>
    /(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)\s*[-–]\s*(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)/.test(
      String(s || "")
    );
  const idx = [];
  for (let i = head.length - 1; i >= 0 && idx.length < 3; i--)
    if (head[i] && weekLike(head[i])) idx.push(i);
  idx.reverse();
  const findRow = (re) => rows.find((r) => r && r[0] && re.test(String(r[0])));
  const num = (r, i) =>
    Number(String(r?.[i] ?? "").replace(/[^\d.-]/g, "")) || 0;
  const pct = (s) => {
    const t = String(s ?? "").trim();
    return /^\d+(\.\d+)?%$/.test(t)
      ? Number(t.replace("%", ""))
      : Number(s) || 0;
  };

  const rOut = findRow(/^Outbound$/i);
  const rSR = findRow(/^success\s*outbound/i);
  return idx.map((i) => ({
    week: String(head[i]).trim(),
    total: num(rOut, i),
    success: Math.round(num(rOut, i) * (pct(rSR?.[i]) / 100)),
    sr: pct(rSR?.[i]) || 0,
  }));
}

function getSingleWeekLabelFromTotaly(wb, sheetName = "totaly") {
  const ws = wb.Sheets[sheetName];
  if (!ws) return null;
  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  const header = rows[0] || [];
  const isWeekLike = (s) =>
    /(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)\s*[-–]\s*(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)/.test(
      String(s || "")
    );
  // баруун талаас week-тэй хамгийн сүүлийг нь авна (өсөлт, эзлэх хув гэх мэт баганын гарчгийг алгасна)
  for (let i = header.length - 1; i >= 0; i--) {
    const v = String(header[i] || "").trim();
    if (!v) continue;
    if (/өрчлөлт|эзлэх\s*хув/i.test(v)) continue;
    if (isWeekLike(v)) return v;
  }
  return null;
}

// Нөөц хувилбар: дурын header массиваас “week-like” хамгийн сүүлийн шошго олох
function getLastWeekLikeFromHeader(header) {
  const isWeekLike = (s) =>
    /(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)\s*[-–]\s*(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)/.test(
      String(s || "")
    );
  for (let i = header.length - 1; i >= 0; i--) {
    if (isWeekLike(header[i])) return String(header[i]).trim();
  }
  return null;
}

function extractOstTop10_APS(
  prevWb,
  currWb,
  { sheetName = "Osticket1", company = "Ард Тэтгэврийн Сан" } = {}
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
      cat: hdr.findIndex((h) => /ангилал/i.test(h)),
      sub: hdr.findIndex((h) => /туслах\s*ангилал/i.test(h)),
    };
    if (idx.comp < 0 || idx.cat < 0 || idx.sub < 0)
      throw new Error("[Ost] missing columns");
    const bag = {
      Лавлагаа: new Map(),
      Үйлчилгээ: new Map(),
      Гомдол: new Map(),
    };
    const norm = (s) => String(s || "").trim();
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;
      if (norm(row[idx.comp]) !== company) continue;
      const cat = norm(row[idx.cat]);
      const sub = norm(row[idx.sub]);
      if (!sub) continue;
      if (!bag[cat]) continue;
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
    return { lav: top("Лавлагаа"), uil: top("Үйлчилгээ"), gom: top("Гомдол") };
  };

  const prev = read(prevWb),
    curr = read(currWb);
  const join = (p, c) => {
    const names = new Set([...p.map((x) => x.name), ...c.map((x) => x.name)]);
    const mapP = new Map(p.map((x) => [x.name, x.curr]));
    const mapC = new Map(c.map((x) => [x.name, x.curr]));
    const out = [...names].map((n) => {
      const a = mapP.get(n) || 0,
        b = mapC.get(n) || 0;
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

  // дараа нь return хийхдээ:
  return {
    labels: [prevLabel, currLabel],
    lav: join(prev.lav, curr.lav),
    uil: join(prev.uil, curr.uil),
    gom: join(prev.gom, curr.gom),
  };
}

function renderAPSLayout({ aps, weeklyCat, top10, outbound3w }) {
  const pct = (v) => `${(v * 100).toFixed(0)}%`;
  const sumArr = (a, b, c) =>
    a.map((_, i) => (a[i] || 0) + (b[i] || 0) + (c[i] || 0));
  const totals = sumArr(weeklyCat.lav, weeklyCat.uil, weeklyCat.gom);
  const last = totals.at(-1) || 0,
    prev = totals.at(-2) || 0;
  const delta = prev ? (last - prev) / prev : 0;
  const apsLabels = (aps?.points || []).map((p) => p.label);
  const apsData = (aps?.points || []).map((p) => p.value);

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
        apsLabels.length
      } сараар/</div>
      <canvas id="apsLine"></canvas>
    </div>
    <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
    <script>(function(){
      const ctx=document.getElementById('apsLine').getContext('2d');
      new Chart(ctx,{
        type:'line',
        data:{ labels:${JSON.stringify(apsLabels)},
               datasets:[{ label:'', data:${JSON.stringify(
                 apsData
               )}, tension:.3, pointRadius:4 }]},
        options:{ animation:false, plugins:{legend:{display:false}}, scales:{ y:{ beginAtZero:true } } }
      });
    })();</script>`;

  const weeklyCard = `
  <div class="card">
    <div class="card-title">НИЙТ БҮРТГЭЛ /Сүүлийн 4 долоо хоногоор/</div>
    <div class="grid">
      <canvas id="weeklyCat"></canvas>
      <ul style="margin:10px 0 0 18px;line-height:1.6">
        <li>Тайлант 7 хоногт нийт <b>${last}</b>
        (${pct((weeklyCat.lav.at(-1) || 0) / Math.max(1, last))} лавлагаа,
         ${pct((weeklyCat.uil.at(-1) || 0) / Math.max(1, last))} үйлчилгээ,
         ${pct((weeklyCat.gom.at(-1) || 0) / Math.max(1, last))} гомдол).</li>
        <li>Өмнөх 7 хоногоос <b>${
          delta >= 0 ? "өссөн" : "буурсан"
        }</b>: <b>${pct(Math.abs(delta))}</b>.</li>
      </ul>
    </div>
  </div>
  <script>(function(){
    ${dataLbl}
    const ctx=document.getElementById('weeklyCat').getContext('2d');
    new Chart(ctx,{type:'bar',
      data:{labels:${JSON.stringify(weeklyCat.labels)},
        datasets:[
          {label:'Лавлагаа', data:${JSON.stringify(weeklyCat.lav)}},
          {label:'Үйлчилгээ',data:${JSON.stringify(weeklyCat.uil)}},
          {label:'Гомдол',   data:${JSON.stringify(weeklyCat.gom)}},
        ]},
      options:{animation:false,plugins:{legend:{position:'bottom'}},scales:{y:{beginAtZero:true}}},
      plugins:[dataLabel]
    });
  })();</script>`;

  const smallMeter = (title, prev, curr) => {
    const d = prev ? (curr - prev) / prev : 0;
    const w = Math.min(100, Math.round((curr / Math.max(curr, prev, 1)) * 100));
    return `
      <div class="card soft" style="padding:10px 14px">
        <div style="font-weight:600;margin-bottom:4px">${title}</div>
        <div style="display:flex;align-items:center;gap:10px">
          <div style="flex:1;height:10px;background:#eee;border-radius:999px;overflow:hidden">
            <div style="width:${w}%;height:100%;background:#3b82f6"></div>
          </div>
          <div style="min-width:120px">${prev} → <b>${curr}</b> (${
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
          ${rows
            .map(
              (r) => `<tr><td>${escapeHtml(r.name)}</td>
                         <td class="num">${r.prev || 0}</td>
                         <td class="num">${r.curr || 0}</td>
                         <td class="num ${r.delta >= 0 ? "up" : "down"}">${
                r.delta >= 0 ? "▲" : "▼"
              } ${(Math.abs(r.delta) * 100).toFixed(0)}%</td></tr>`
            )
            .join("")}
        </tbody>
      </table>
    </div>`;

  const outboundTable = `
    <div class="card">
      <div class="card-title">OUTBOUND</div>
      <table class="cmp">
        <thead><tr><th>Онцлох</th><th>Залгасан</th><th>Амжилттай</th><th>SR</th></tr></thead>
        <tbody>
          ${outbound3w
            .map(
              (r) =>
                `<tr><td>${escapeHtml(r.week)}</td><td class="num">${
                  r.total
                }</td><td class="num">${r.success}</td><td class="num">${
                  r.sr
                }%</td></tr>`
            )
            .join("")}
          <tr><td><b>Нийт</b></td><td class="num"><b>${outbound3w.reduce(
            (a, b) => a + b.total,
            0
          )}</b></td>
              <td class="num"><b>${outbound3w.reduce(
                (a, b) => a + b.success,
                0
              )}</b></td>
              <td class="num"><b>${Math.round(
                (outbound3w.reduce((a, b) => a + b.success, 0) * 100) /
                  Math.max(
                    1,
                    outbound3w.reduce((a, b) => a + b.total, 0)
                  )
              )}%</b></td></tr>
        </tbody>
      </table>
    </div>`;

  return `
  <section>
    <div class="grid">
      ${lineCard}
      ${weeklyCard}
    </div>

    <div class="grid grid-3" style="margin-top:8px">
      ${smallMeter(
        "ЛАВЛАГАА",
        weeklyCat.lav.at(-2) || 0,
        weeklyCat.lav.at(-1) || 0
      )}
      ${smallMeter(
        "ҮЙЛЧИЛГЭЭ",
        weeklyCat.uil.at(-2) || 0,
        weeklyCat.uil.at(-1) || 0
      )}
      ${smallMeter(
        "ГОМДОЛ",
        weeklyCat.gom.at(-2) || 0,
        weeklyCat.gom.at(-1) || 0
      )}
    </div>

    <div class="grid grid-2" style="margin-top:8px">
      ${outboundTable}
      <div></div>
    </div>

    <div class="grid grid-1" style="margin-top:8px">
      ${topTable("ТОП Лавлагаа", top10.lav, top10.labels)}
      ${topTable("ТОП Үйлчилгээ", top10.uil, top10.labels)}
      ${topTable("ТОП Гомдол", top10.gom, top10.labels)}
    </div>
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

  const weeklyCat = extractWeeklyByCategoryFromTotaly(wbCurr, "totaly");
  const outbound3w = extractOutbound3Weeks(wbCurr, "totaly");
  const top10 = extractOstTop10_APS(wbPrev, wbCurr, {
    sheetName: "Osticket1",
    company: "Ард Тэтгэврийн Данс",
  });

  const aps = extractAPSLatestMonths(wbCurr, {
    sheetName: "APS",
    yearLabel: "2025",
    takeLast: 4,
  });
  const apsHtml = renderAPSLayout({
    aps, // << өмнө нь apsMonths гээд array явж байсан
    weeklyCat,
    top10,
    outbound3w,
  });
  const cover = renderAssCover({
    company: "АРД ТЭТГЭВРИЙН САН",
    periodText: `${aps.points[0]?.label ?? ""} – ${
      aps.points.at(-1)?.label ?? ""
    } (${aps.year})`,
  });

  // PDF
  const monday = dayjs().tz(CONFIG.TIMEZONE).startOf("week").add(1, "day");
  const stamp = monday.format("YYYYMMDD");
  const pdfName = `ard-tetgever-${monday.format("YYYYMMDD")}.pdf`;
  const pdfPath = path.join(CONFIG.OUT_DIR, pdfName);

  let body = "";
  body += cover;
  body += apsHtml;

  const html = wrapHtml(body);

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
