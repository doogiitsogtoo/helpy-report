// ard-app.company.js — Компанийн тайлан (APS + totaly + Osticket1) → HTML → PDF (+optional email)

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
  GOMDOL_FILE: "./gomdol-weekly.xlsx",

  // Sheets
  ASS_SHEET: "APS",
  ASS_COMPANY: process.env.ASS_COMPANY || "Ардын Тэтгэврийн Данс",
  ASS_YEAR: process.env.ASS_YEAR || "2025",
  ASS_TAKE_LAST_N_MONTHS: Number(process.env.ASS_TAKE_LAST_N_MONTHS || 4),
  OST_SHEET: "Osticket1",

  // PDF / Email
  OUT_DIR: "./out",
  CSS_FILE: "./css/template.css",
  SAVE_HTML: true,
  HTML_NAME_PREFIX: "report",
  REPORT_TITLE:
    process.env.REPORT_TITLE || "Ардын Тэтгэврийн Данс — 7 хоногийн тайлан",
  SUBJECT_PREFIX: process.env.SUBJECT_PREFIX || "[ArdApp Weekly]",
  EMAIL_ENABLED: String(process.env.EMAIL_ENABLED ?? "true") === "true",
  SCHED_ENABLED: String(process.env.SCHED_ENABLED ?? "true") === "true",
};

// ────────────────────────────────────────────────────────────────
// Helpers
// ────────────────────────────────────────────────────────────────
const escapeHtml = (s) =>
  String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");

const numStrict = (v) =>
  Number(
    String(v ?? "")
      .replace(/[\s,\u00A0]/g, "")
      .replace(/[^\d.-]/g, "")
  ) || 0;

const normalize = (s) =>
  String(s || "")
    .toLowerCase()
    .replace(/[\s\u00A0]+/g, " ")
    .replace(/ын|ийн/g, "")
    .trim();

function renderAssCover({ company, periodText }) {
  return `
  <section class="hero" style="margin-bottom:16px">
    <div style="background:linear-gradient(135deg,#ef4444,#f97316);
                border-radius:12px;padding:28px;display:flex;
                justify-content:space-between;align-items:center;min-height:200px;">
      <div style="background:#fff;border-radius:16px;padding:20px 24px;display:inline-block">
        <div style="font-weight:700;font-size:28px;letter-spacing:.5px;color:#ef4444">ARD</div>
        <div style="color:#666;margin-top:4px">Хүчтэй. Хамтдаа.</div>
      </div>
      <div style="color:#fff;text-align:right;padding:8px 16px">
        <div style="font-size:32px;font-weight:800;line-height:1.1">${escapeHtml(
          company
        )}</div>
        <div style="opacity:.9;margin-top:8px">${escapeHtml(
          periodText || ""
        )}</div>
      </div>
    </div>
  </section>`;
}

// ────────────────────────────────────────────────────────────────
// Extractors
// ────────────────────────────────────────────────────────────────
function extractAPSLatestMonths(
  wb,
  {
    sheetName = CONFIG.ASS_SHEET,
    yearLabel = CONFIG.ASS_YEAR,
    takeLast = CONFIG.ASS_TAKE_LAST_N_MONTHS,
  } = {}
) {
  const ws = wb.Sheets[sheetName];
  if (!ws) throw new Error(`[APS] Sheet not found: ${sheetName}`);

  const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });
  const yearHeadRowIdx = rows.findIndex(
    (r) => r && r.filter(Boolean).some((v) => /^\d{4}$/.test(String(v)))
  );
  if (yearHeadRowIdx < 0) throw new Error("[APS] Year header row not found");

  const header = (rows[yearHeadRowIdx] || []).map((v) =>
    String(v || "").trim()
  );
  const yearCol = header.findIndex((v) => v === String(yearLabel));
  if (yearCol < 0) throw new Error(`[APS] Year col not found: ${yearLabel}`);

  const monthLike = (s) =>
    /^\s*\d+\s*(сар|cap)\s*$/i.test(String(s || "").trim());
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
      continue;
    }
    monthRows.push({ label, value: numStrict(row[yearCol]) });
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

  const weekLike = (s) =>
    /(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)\s*[-–]\s*(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)/.test(
      String(s || "")
    );
  const idx = [];
  for (let i = head.length - 1; i >= 0 && idx.length < 4; i--) {
    if (head[i] && weekLike(head[i])) idx.push(i);
  }
  idx.reverse();
  const labels = idx.map((i) => String(head[i]).trim());

  const findRow = (re) => rows.find((r) => r && r[0] && re.test(String(r[0])));
  const rLav = findRow(/^Лавлагаа$/i);
  const rUil = findRow(/^Үйлчилгээ$/i);
  const rGom = findRow(/^Гомдол$/i);
  if (!rLav || !rUil || !rGom)
    throw new Error("[totaly] Missing rows (Лавлагаа/Үйлчилгээ/Гомдол)");

  const num = (r, i) => numStrict(r?.[i]);
  return {
    labels,
    lav: idx.map((i) => num(rLav, i)),
    uil: idx.map((i) => num(rUil, i)),
    gom: idx.map((i) => num(rGom, i)),
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
  for (let i = head.length - 1; i >= 0 && idx.length < 3; i--) {
    if (head[i] && weekLike(head[i])) idx.push(i);
  }
  idx.reverse();

  const findRow = (re) => rows.find((r) => r && r[0] && re.test(String(r[0])));
  const num = (r, i) => numStrict(r?.[i]);
  const pct = (s) => {
    const t = String(s ?? "").trim();
    return /^\d+(\.\d+)?%$/.test(t) ? Number(t.replace("%", "")) : numStrict(s);
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
  for (let i = header.length - 1; i >= 0; i--) {
    const v = String(header[i] || "").trim();
    if (!v) continue;
    if (/өрчлөлт|эзлэх\s*хув/i.test(v)) continue;
    if (isWeekLike(v)) return v;
  }
  return null;
}

// ────────────────────────────────────────────────────────────────
// Osticket TOP 10
// ────────────────────────────────────────────────────────────────
function extractOstTop10_APS(
  prevWb,
  currWb,
  { sheetName = CONFIG.OST_SHEET, company = CONFIG.ASS_COMPANY || "" } = {}
) {
  const read = (wb) => {
    const ws = wb.Sheets[sheetName];
    if (!ws) throw new Error(`[Osticket1] Sheet not found: ${sheetName}`);
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, raw: false });

    const hdr = (rows[0] || []).map((v) =>
      String(v || "")
        .trim()
        .toLowerCase()
    );
    const idx = {
      comp: hdr.findIndex((h) => /(компани|company)/i.test(h)),
      cat: hdr.findIndex((h) => /(ангилал|category)/i.test(h)),
      sub: hdr.findIndex((h) =>
        /(туслах\s*ангилал|дэд\s*ангилал|sub.?category|subcategory)/i.test(h)
      ),
    };
    if (idx.comp < 0 || idx.cat < 0 || idx.sub < 0) {
      throw new Error(
        "[Osticket1] 'Компани'/'Ангилал'/'Туслах(Дэд) ангилал' багана олдсонгүй"
      );
    }

    const want = normalize(company);
    const useCompanyFilter = Boolean(want);

    const bag = {
      Лавлагаа: new Map(),
      Үйлчилгээ: new Map(),
      Гомдол: new Map(),
    };
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const compCell = normalize(row[idx.comp]);
      if (useCompanyFilter && !compCell.includes(want)) continue;

      const cat = String(row[idx.cat] || "").trim();
      const sub = String(row[idx.sub] || "").trim();
      if (!sub || !bag[cat]) continue;

      bag[cat].set(sub, (bag[cat].get(sub) || 0) + 1);
    }

    const top = (k) =>
      [...bag[k].entries()]
        .map(([name, val]) => ({ name, curr: val }))
        .sort((a, b) => b.curr - a.curr)
        .slice(0, 10);

    return { lav: top("Лавлагаа"), uil: top("Үйлчилгээ"), gom: top("Гомдол") };
  };

  const prev = read(prevWb);
  const curr = read(currWb);

  const join = (p, c) => {
    const names = new Set([...p.map((x) => x.name), ...c.map((x) => x.name)]);
    const mP = new Map(p.map((x) => [x.name, x.curr]));
    const mC = new Map(c.map((x) => [x.name, x.curr]));
    const out = [...names].map((n) => {
      const a = mP.get(n) || 0,
        b = mC.get(n) || 0;
      const base = a > 0 ? a : b > 0 ? b : 1;
      return { name: n, prev: a, curr: b, delta: (b - a) / base };
    });
    out.sort((x, y) => y.curr - x.curr || y.prev - x.prev);
    return out.slice(0, 10);
  };

  const prevLabel =
    getSingleWeekLabelFromTotaly(prevWb, "totaly") || "Өмнөх 7 хоног";
  const currLabel =
    getSingleWeekLabelFromTotaly(currWb, "totaly") || "Одоогийн 7 хоног";

  return {
    labels: [prevLabel, currLabel],
    lav: join(prev.lav, curr.lav),
    uil: join(prev.uil, curr.uil),
    gom: join(prev.gom, curr.gom),
  };
}

// ────────────────────────────────────────────────────────────────
// Render (company chart numbers + stacked total labels)
// ────────────────────────────────────────────────────────────────
function renderAPSLayout({ aps, weeklyCat, top10, outbound3w }) {
  const pct = (v) => `${(v * 100).toFixed(0)}%`;
  const sumArr = (a, b, c) =>
    a.map((_, i) => (a[i] || 0) + (b[i] || 0) + (c[i] || 0));
  const totals = sumArr(weeklyCat.lav, weeklyCat.uil, weeklyCat.gom);
  const last = totals.at(-1) || 0;
  const prev = totals.at(-2) || 0;
  const delta = prev ? (last - prev) / prev : 0;

  const apsLabels = (aps?.points || []).map((p) => p.label);
  const apsData = (aps?.points || []).map((p) => p.value);

  // dataLabelPlus — сегментүүдийн тоо + stacked дээр НИЙТ дүн
  const dataLbl = `
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
        // сегментүүд
        sets.forEach((ds,di)=>{
          const meta = chart.getDatasetMeta(di);
          (ds.data||[]).forEach((v,i)=>{
            if(v==null) return;
            const el = meta.data?.[i]; if(!el) return;
            if(isBar){ const h=Math.abs((el.base??el.y)-el.y); if(h<18) return; }
            const pos = el.tooltipPosition ? el.tooltipPosition() : el;
            const x = pos.x;
            const yText = isBar ? (el.y + (el.base??el.y))/2 : (pos.y-6);
            const t = Number(v).toLocaleString('mn-MN');
            ctx.strokeText(t,x,yText); ctx.fillText(t,x,yText);
          });
        });
        // totals (stacked bar)
        if(isBar && isStacked && y && typeof y.getPixelForValue==='function'){
          const totals = labels.map((_,i)=>sets.reduce((s,ds)=>s+(+ds.data?.[i]||0),0));
          const metas = sets.map((_,di)=>chart.getDatasetMeta(di));
          totals.forEach((tot,i)=>{
            const base = metas[0]?.data?.[i]; if(!base) return;
            const x = (base.tooltipPosition?base.tooltipPosition():base).x;
            const yTop = y.getPixelForValue(tot);
            let minLabelY = Infinity;
            metas.forEach(m=>{
              const el=m.data?.[i]; if(!el) return;
              const h=Math.abs((el.base??el.y)-el.y);
              if(h>=18){ const yc=(el.y+(el.base??el.y))/2; minLabelY=Math.min(minLabelY,yc); }
            });
            let yText = yTop - 10; if(yText > (minLabelY-12)) yText = minLabelY - 12;
            const t = Number(tot).toLocaleString('mn-MN');
            ctx.strokeText(t,x,yText); ctx.fillText(t,x,yText);
          });
        }
        ctx.restore();
      }
    };
  `;

  const lineCard = apsLabels.length
    ? `
    <div class="card" style="height: 340px; margin-bottom: 12px;">
      <div class="card-title">НИЙТ ХАНДАЛТ /Сүүлийн ${
        apsLabels.length
      } сараар/</div>
      <canvas id="apsLine"></canvas>
    </div>
    <script>(function(){
      const ctx=document.getElementById('apsLine').getContext('2d');
      new Chart(ctx,{
        type:'line',
        data:{ labels:${JSON.stringify(apsLabels)},
               datasets:[{ label:'', data:${JSON.stringify(
                 apsData
               )}, tension:.3, pointRadius:4 }]},
        options:{ animation:false, plugins:{legend:{display:false}}, scales:{ y:{ beginAtZero:true } }, responsive:true, maintainAspectRatio:false }
      });
    })();</script>`
    : `
    <div class="card" style="min-height:120px;display:flex;align-items:center;"><div>
      <div class="card-title" style="margin-bottom:6px">НИЙТ ХАНДАЛТ</div>
      <div style="color:#666">APS шит дээр ${CONFIG.ASS_YEAR} оны сард идэвхтэй утга олдсонгүй.</div>
    </div></div>`;

  const weeklyCard = `
  <div class="card">
    <div class="card-title">НИЙТ БҮРТГЭЛ /Сүүлийн ${
      weeklyCat.labels.length
    } долоо хоногоор/</div>
    <div class="grid">
      <canvas id="weeklyCat"></canvas>
      <ul style="margin:10px 0 0 18px;line-height:1.6">
        <li>Тайлант 7 хоногт нийт <b>${last.toLocaleString("mn-MN")}</b>
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
    new Chart(ctx,{
      type:'bar',
      data:{labels:${JSON.stringify(weeklyCat.labels)},
        datasets:[
          {label:'Лавлагаа', data:${JSON.stringify(weeklyCat.lav)}},
          {label:'Үйлчилгээ',data:${JSON.stringify(weeklyCat.uil)}},
          {label:'Гомдол',   data:${JSON.stringify(weeklyCat.gom)}},
        ]},
      options:{
        animation:false,
        plugins:{legend:{position:'bottom'}},
        scales:{
          x:{ stacked:true, categoryPercentage:0.6, barPercentage:0.8 },
          y:{ stacked:true, beginAtZero:true, ticks:{ maxTicksLimit:6 } }
        },
        responsive:true, maintainAspectRatio:false
      },
      plugins:[dataLabelPlus]
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
          <div style="min-width:120px">${prev.toLocaleString(
            "mn-MN"
          )} → <b>${curr.toLocaleString("mn-MN")}</b> (${d >= 0 ? "+" : ""}${(
      d * 100
    ).toFixed(0)}%)</div>
        </div>
      </div>`;
  };

  const tableRowsHtml = (rows) =>
    rows.length
      ? rows
          .map(
            (r) => `<tr><td>${escapeHtml(r.name)}</td>
                      <td class="num">${(r.prev || 0).toLocaleString(
                        "mn-MN"
                      )}</td>
                      <td class="num">${(r.curr || 0).toLocaleString(
                        "mn-MN"
                      )}</td>
                      <td class="num ${r.delta >= 0 ? "up" : "down"}">
                        ${r.delta >= 0 ? "▲" : "▼"} ${(
              Math.abs(r.delta) * 100
            ).toFixed(0)}%
                      </td></tr>`
          )
          .join("")
      : `<tr><td colspan="4" style="text-align:center;color:#666">Мэдээлэл алга</td></tr>`;

  const topTable = (title, rows, labels) => `
    <div class="card">
      <div class="card-title">${title}</div>
      <table class="cmp">
        <thead><tr><th></th><th>${labels?.[0] || "Өмнөх"}</th><th>${
    labels?.[1] || "Одоогийн"
  }</th><th>%</th></tr></thead>
        <tbody>${tableRowsHtml(rows)}</tbody>
      </table>
    </div>`;

  const outboundTable = `
    <div class="card">
      <div class="card-title">OUTBOUND</div>
      <table class="cmp">
        <thead><tr><th>Онцлох</th><th>Залгасан</th><th>Амжилттай</th><th>SR</th></tr></thead>
        <tbody>
          ${
            outbound3w.length
              ? outbound3w
                  .map(
                    (r) =>
                      `<tr><td>${escapeHtml(r.week)}</td><td class="num">${
                        r.total
                      }</td><td class="num">${r.success}</td><td class="num">${
                        r.sr
                      }%</td></tr>`
                  )
                  .join("")
              : `<tr><td colspan="4" style="text-align:center;color:#666">Мэдээлэл алга</td></tr>`
          }
          ${
            outbound3w.length
              ? `<tr><td><b>Нийт</b></td><td class="num"><b>${outbound3w.reduce(
                  (a, b) => a + b.total,
                  0
                )}</b></td>
                   <td class="num"><b>${outbound3w.reduce(
                     (a, b) => a + b.success,
                     0
                   )}</b></td>
                   <td class="num"><b>${
                     Math.round(
                       (outbound3w.reduce((a, b) => a + b.success, 0) * 100) /
                         Math.max(
                           1,
                           outbound3w.reduce((a, b) => a + b.total, 0)
                         )
                     ) || 0
                   }%</b></td></tr>`
              : ""
          }
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
  let browser;
  try {
    browser = await puppeteer.launch({
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
      defaultViewport: { width: 1280, height: 900, deviceScaleFactor: 2 },
    });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: "networkidle0" });
    // charts ready guard (optional; quietly ignore if not set)
    await page
      .waitForFunction(() => window.__chartsReady === true, { timeout: 12000 })
      .catch(() => {});
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
    if (browser) await browser.close();
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
    html: `<p>Сайн байна уу,</p><p>Ардын Тэтгэврийн Данс 7 хоногийн тайланг хавсаргав.</p><p style="color:#666;font-size:12px">Автоматаар илгээв.</p>`,
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
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
  <style>
    ${css}
    .container { max-width: 1080px; margin: 0 auto; }
    .card { padding: 16px; border-radius: 12px; box-shadow: 0 2px 6px rgba(0,0,0,.05); }
    .card-title { font-weight: 700; margin-bottom: 8px; }
    .cmp { width: 100%; border-collapse: collapse; }
    .cmp th,.cmp td{ padding:8px 10px; border-bottom:1px solid #eee; }
    .cmp .num{ text-align:right; } .cmp .up{color:#16a34a} .cmp .down{color:#ef4444}
    #apsLine{ width:100% !important; height:340px !important; }
    #weeklyCat{ width:100% !important; height:300px !important; }
    @media print { .card{ break-inside: avoid; page-break-inside: avoid; } }
  </style>
  <script>window.__chartsReadyCount=0; window.__chartsReady=false; window._incReady=function(){window.__chartsReadyCount++; if(window.__chartsReadyCount>=2) window.__chartsReady=true;}</script>
</head>
<body>
 <div class="container py-3">
    <div class="row g-3">
      ${bodyHtml}
      <div class="footer">Автоматаар бэлтгэсэн тайлан (Node.js)</div>
    </div>
  </div>
  <script>_incReady();</script>
</body>
</html>`;
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

  const weeklyCat = extractWeeklyByCategoryFromTotaly(wbCurr, "totaly");
  const outbound3w = extractOutbound3Weeks(wbCurr, "totaly");
  const top10 = extractOstTop10_APS(wbPrev, wbCurr, {
    sheetName: CONFIG.OST_SHEET,
    company: CONFIG.ASS_COMPANY || "",
  });

  const aps = extractAPSLatestMonths(wbCurr, {
    sheetName: CONFIG.ASS_SHEET,
    yearLabel: CONFIG.ASS_YEAR,
    takeLast: CONFIG.ASS_TAKE_LAST_N_MONTHS,
  });

  const cover = renderAssCover({
    company: CONFIG.ASS_COMPANY || "АРД",
    periodText: `${weeklyCat.labels.at(-2) ?? weeklyCat.labels[0] ?? ""} – ${
      weeklyCat.labels.at(-1) ?? ""
    }`,
  });

  const body = cover + renderAPSLayout({ aps, weeklyCat, top10, outbound3w });
  const html = wrapHtml(body);

  const monday = dayjs().tz(CONFIG.TIMEZONE).startOf("week").add(1, "day");
  const stamp = monday.format("YYYYMMDD");
  const pdfPath = path.join(CONFIG.OUT_DIR, `ard-tetgever-${stamp}.pdf`);

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
      process.env.RECIPIENTS || "(recipients not set)"
    }`
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
