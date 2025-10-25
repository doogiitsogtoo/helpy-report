// ard-app.dynamic.js — Dynamic Excel → HTML → PDF → Email (with non-overlapping numbers)
// ---------------------------------------------------------------------------------
// ✔ Чартууд дээрх тоо давхцахгүй (stacked bar сегмент өндөр < 18px бол сегментийн тоо нуух)
// ✔ НИЙТ шошгыг баганын дээдээс 10px дээр байрлуулж, сегментийн шошготой давхцахаас сэргийлнэ
// ✔ APS/ArdApp сарын график хоосон бол мессеж, бусад карт хэвийн
// ✔ PDF хэвлэхээс өмнө чартууд бэлэн эсэхийг хүлээнэ
// ✔ PREV_FILE auto-олддог, ТОП хүснэгт, Outbound 3 долоо хоног, scheduler хэвээр
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

  // Excel files
  PREV_FILE: process.env.PREV_FILE || "./ARD 10.06-10.12.xlsx",
  CURR_FILE: process.env.CURR_FILE || "./ARD 10.13-10.19.xlsx",
  GOMDOL_FILE: process.env.GOMDOL_FILE || "./gomdol-weekly.xlsx",

  // Sheets
  ASS_SHEET: process.env.ASS_SHEET || "ArdApp",
  ASS_COMPANY: process.env.ASS_COMPANY || "", // empty → all companies
  ASS_YEAR: process.env.ASS_YEAR || "auto",
  ASS_TAKE_LAST_N_MONTHS: Number(process.env.ASS_TAKE_LAST_N_MONTHS || 4),
  OST_SHEET: process.env.OST_SHEET || "Osticket1",

  // PDF / Email
  OUT_DIR: process.env.OUT_DIR || "./out",
  CSS_FILE: process.env.CSS_FILE || "./css/template.css",
  SAVE_HTML: String(process.env.SAVE_HTML ?? "true") === "true",
  HTML_NAME_PREFIX: process.env.HTML_NAME_PREFIX || "report",
  REPORT_TITLE: process.env.REPORT_TITLE || "Ард Апп — 7 хоногийн тайлан",
  SUBJECT_PREFIX: process.env.SUBJECT_PREFIX || "[ArdApp Weekly]",
  EMAIL_ENABLED: String(process.env.EMAIL_ENABLED ?? "false") === "true",
  SCHED_ENABLED: String(process.env.SCHED_ENABLED ?? "true") === "true",
};

// ────────────────────────────────────────────────────────────────
// Helpers
// ────────────────────────────────────────────────────────────────
function escapeHtml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
const asNum = (v) =>
  Number(
    String(v ?? "")
      .replace(/[\s,\u00A0]/g, "")
      .replace(/[^\d.-]/g, "")
  ) || 0;
const pad2 = (n) => String(n).padStart(2, "0");

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
    raw: `${m[1]}.${m[2]}-${m[3]}.${m[4]}`,
  };
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

function autoFindPrevFile(currFile) {
  if (!currFile) return "";
  const dir = path.dirname(currFile);
  const base = path.basename(currFile);
  const files = fs.readdirSync(dir).filter((f) => /\.xlsx$/i.test(f));
  const wkCurr = parseWeekFromFilename(base);
  if (!wkCurr) return "";

  // infer year from current
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

// totaly helpers
function parseWeekLabelCell(s) {
  const m = String(s || "").match(
    /(\d{1,2})[./-](\d{1,2}).*?[-–].*?(\d{1,2})[./-](\d{1,2})/
  );
  if (!m) return null;
  return { m1: +m[1], d1: +m[2], m2: +m[3], d2: +m[4], raw: m[0] };
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

  // find year header row
  const yearHeadRowIdx = rows.findIndex(
    (r) => r && r.filter(Boolean).some((v) => /^\d{4}$/.test(String(v)))
  );
  if (yearHeadRowIdx < 0) throw new Error("[APS] Year header row not found");

  const header = (rows[yearHeadRowIdx] || []).map((v) =>
    String(v || "").trim()
  );

  // choose year
  let year = String(yearLabel);
  if (year === "auto") {
    const yCols = header
      .map((h, i) => (/^\d{4}$/.test(h) ? i : -1))
      .filter((i) => i >= 0);
    let pick = null;
    for (const i of yCols) {
      const has = rows.slice(yearHeadRowIdx + 1).some((r) => asNum(r?.[i]) > 0);
      if (has) pick = i;
    }
    if (pick == null) pick = yCols.at(-1) ?? -1;
    if (pick < 0) throw new Error("[APS] No usable year column");
    year = header[pick];
  }
  const yearCol = header.findIndex((v) => v === String(year));
  if (yearCol < 0) throw new Error(`[APS] Year col not found: ${year}`);

  // month rows
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
    const value = asNum(row[yearCol]);
    monthRows.push({ label, value });
  }

  const active = monthRows.filter((m) => m.value > 0);
  const points = active.slice(-takeLast);

  return { year, points, allMonths: monthRows };
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
    const v = head[i];
    if (v && weekLike(v)) idx.push(i);
  }
  idx.reverse();
  const labels = idx.map((i) => String(head[i]).trim());

  const findRow = (re) => rows.find((r) => r && r[0] && re.test(String(r[0])));
  const num = (r, i) => asNum(r?.[i]);

  const rLav = findRow(/^Лавлагаа$/i);
  const rUil = findRow(/^Үйлчилгээ$/i);
  const rGom = findRow(/^Гомдол$/i);

  if (!rLav || !rUil || !rGom)
    throw new Error("[totaly] Missing rows: Лавлагаа/Үйлчилгээ/Гомдол");

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
  for (let i = head.length - 1; i >= 0 && idx.length < 3; i--)
    if (head[i] && weekLike(head[i])) idx.push(i);
  idx.reverse();
  const findRow = (re) => rows.find((r) => r && r[0] && re.test(String(r[0])));
  const num = (r, i) => asNum(r?.[i]);
  const pct = (s) => {
    const t = String(s ?? "").trim();
    return /^\d+(\.\d+)?%$/.test(t) ? Number(t.replace("%", "")) : asNum(s);
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
  for (let i = header.length - 1; i >= 0; i--) {
    const v = String(header[i] || "").trim();
    if (!v) continue;
    if (/өрчлөлт|эзлэх\s*хув/i.test(v)) continue;
    if (/(\d{1,2}[./-]\d{1,2}).*[-–].*(\d{1,2}[./-]\d{1,2})/.test(v)) return v;
  }
  return null;
}

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
      comp: hdr.findIndex((h) => /(компани|company)/.test(h)),
      cat: hdr.findIndex((h) => /(ангилал|category)/.test(h)),
      sub: hdr.findIndex((h) =>
        /(туслах\s*ангилал|дэд\s*ангилал|sub.?category)/.test(h)
      ),
      created: hdr.findIndex((h) =>
        /(үүссэн\s*огноо|нээсэн\s*огноо|created|open)/.test(h)
      ),
      closed: hdr.findIndex((h) => /(хаагдсан\s*огноо|closed)/.test(h)),
    };
    if (idx.sub < 0 || idx.cat < 0 || idx.comp < 0) {
      throw new Error(
        "[Osticket1] 'Компани'/'Ангилал'/'Туслах(Дэд) ангилал' багана олдсонгүй"
      );
    }
    const dateCol = idx.created >= 0 ? idx.created : idx.closed;

    const bag = {
      Лавлагаа: new Map(),
      Үйлчилгээ: new Map(),
      Гомдол: new Map(),
    };
    const clean = (s) => String(s || "").trim();

    for (let r = 1; r < rows.length; r++) {
      const row = rows[r] || [];

      if (company && clean(row[idx.comp]) !== company) continue;

      if (dateCol >= 0) {
        const when = parseExcelDate(row[dateCol]);
        if (!when) continue; // bad date row skip
      }

      const cat = clean(row[idx.cat]);
      const sub = clean(row[idx.sub]);
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
// Render
// ────────────────────────────────────────────────────────────────
function renderAssCover({ company, periodText }) {
  return `
  <section class="hero" style="margin-bottom:16px">
    <div style="background:linear-gradient(135deg,#ef4444,#f97316);
                border-radius:12px;padding:28px;display:flex;
                justify-content:space-between;align-items:center;min-height:160px;">
      <div style="background:#fff;border-radius:16px;padding:16px 20px;display:inline-block">
        <div style="font-weight:700;font-size:24px;letter-spacing:.5px;color:#ef4444">ARD</div>
        <div style="color:#666;margin-top:4px">Хүчтэй. Хамтдаа.</div>
      </div>
      <div style="color:#fff;text-align:right;padding:8px 16px">
        <div style="font-size:30px;font-weight:800;line-height:1.1">${escapeHtml(
          company
        )}</div>
        <div style="opacity:.9;margin-top:6px">${escapeHtml(
          periodText || ""
        )}</div>
      </div>
    </div>
  </section>`;
}

function renderAPSLayout({ aps, weeklyCat, top10, outbound3w }) {
  const pctTxt = (v) => `${(v * 100).toFixed(0)}%`;
  const sum3 = (a, b, c) =>
    a.map((_, i) => (a[i] || 0) + (b[i] || 0) + (c[i] || 0));
  const totals = sum3(weeklyCat.lav, weeklyCat.uil, weeklyCat.gom);
  const last = totals.at(-1) || 0,
    prev = totals.at(-2) || 0;
  const delta = prev ? (last - prev) / prev : 0;

  const hasAps = Boolean(aps?.points?.length);
  const apsLabels = hasAps ? aps.points.map((p) => p.label) : [];
  const apsData = hasAps ? aps.points.map((p) => p.value) : [];

  const apsCard = hasAps
    ? `
      <div class="card">
        <div class="card-title">APS / Сүүлийн ${apsLabels.length} сар</div>
        <canvas id="apsLine"></canvas>
      </div>`
    : `
      <div class="card soft" style="display:flex;align-items:center;justify-content:center;min-height:220px">
        <div style="color:#666">APS (сарын) мэдээлэл байхгүй байна.</div>
      </div>`;

  const weeklyCard = `
  <div class="card">
    <div class="card-title">НИЙТ БҮРТГЭЛ /Сүүлийн ${
      weeklyCat.labels.length
    } долоо хоногоор/</div>
    <div style="display:grid;grid-template-columns:1fr 320px;gap:12px">
      <canvas id="weeklyCat"></canvas>
      <ul style="margin:10px 0 0 0;line-height:1.6">
        <li>Тайлант 7 хоногт нийт <b>${last.toLocaleString("en-US")}</b>
        (${pctTxt((weeklyCat.lav.at(-1) || 0) / Math.max(1, last))} лавлагаа,
         ${pctTxt((weeklyCat.uil.at(-1) || 0) / Math.max(1, last))} үйлчилгээ,
         ${pctTxt(
           (weeklyCat.gom.at(-1) || 0) / Math.max(1, last)
         )} гомдол).</li>
        <li>Өмнөх 7 хоногоос <b>${
          delta >= 0 ? "өссөн" : "буурсан"
        }</b>: <b>${pctTxt(Math.abs(delta))}</b>.</li>
      </ul>
    </div>
  </div>`;

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
          <div style="min-width:160px">${(prev || 0).toLocaleString(
            "en-US"
          )} → <b>${(curr || 0).toLocaleString("en-US")}</b> (${
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
                         <td class="num">${(r.prev || 0).toLocaleString(
                           "en-US"
                         )}</td>
                         <td class="num">${(r.curr || 0).toLocaleString(
                           "en-US"
                         )}</td>
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
                `<tr><td>${escapeHtml(r.week)}</td><td class="num">${(
                  r.total || 0
                ).toLocaleString("en-US")}</td><td class="num">${(
                  r.success || 0
                ).toLocaleString("en-US")}</td><td class="num">${
                  r.sr
                }%</td></tr>`
            )
            .join("")}
          <tr><td><b>Нийт</b></td>
              <td class="num"><b>${outbound3w
                .reduce((a, b) => a + (b.total || 0), 0)
                .toLocaleString("en-US")}</b></td>
              <td class="num"><b>${outbound3w
                .reduce((a, b) => a + (b.success || 0), 0)
                .toLocaleString("en-US")}</b></td>
              <td class="num"><b>${Math.round(
                (outbound3w.reduce((a, b) => a + (b.success || 0), 0) * 100) /
                  Math.max(
                    1,
                    outbound3w.reduce((a, b) => a + (b.total || 0), 0)
                  )
              )}%</b></td></tr>
        </tbody>
      </table>
    </div>`;

  return `
  <section>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
      ${apsCard}
      ${weeklyCard}
    </div>

    <div style="display:grid;grid-template-columns:repeat(3, 1fr);gap:12px;margin-top:8px">
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

    <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-top:8px">
      ${outboundTable}
      <div></div>
    </div>

    <div style="display:grid;grid-template-columns:1fr;gap:12px;margin-top:8px">
      ${topTable("ТОП Лавлагаа", top10.lav, top10.labels)}
      ${topTable("ТОП Үйлчилгээ", top10.uil, top10.labels)}
      ${topTable("ТОП Гомдол", top10.gom, top10.labels)}
    </div>

    <script>
      // --- Chart helpers (numbers on charts, non-overlapping) ---
      const fmtInt = v => (v == null ? '' : Number(v).toLocaleString('en-US'));

      const dataLabelPlus = {
        id: 'dataLabelPlus',
        afterDatasetsDraw(chart) {
          const { ctx, data, scales, config } = chart;
          const yScale = scales?.y;
          const labels = data?.labels || [];
          const datasets = data?.datasets || [];

          ctx.save();
          ctx.font = 'bold 12px system-ui,-apple-system,Segoe UI,Roboto,Arial';
          ctx.textAlign = 'center';
          ctx.lineWidth = 3;
          ctx.strokeStyle = 'rgba(255,255,255,.96)';
          ctx.fillStyle = '#111';

          const isBar = config.type === 'bar';
          const isStacked = !!(chart.options?.scales?.x?.stacked && chart.options?.scales?.y?.stacked);

          // 1) Dataset point/segment labels (skip tiny bar segments)
          datasets.forEach((ds, di) => {
            const meta = chart.getDatasetMeta(di);
            (ds.data || []).forEach((v, i) => {
              if (v == null) return;
              const el = meta.data?.[i];
              if (!el) return;

              if (isBar) {
                const h = Math.abs((el.base ?? el.y) - el.y);
                if (h < 18) return; // too small → don't draw to avoid overlap
              }

              const pos = el.tooltipPosition ? el.tooltipPosition() : el;
              const x = pos.x;
              const y = isBar ? (el.y + (el.base ?? el.y)) / 2 : (pos.y - 6);
              const text = fmtInt(v);
              ctx.strokeText(text, x, y);
              ctx.fillText(text, x, y);
            });
          });

          // 2) Stacked totals on top with clearance
          if (isBar && isStacked && yScale && typeof yScale.getPixelForValue === 'function') {
            const totals = labels.map((_, i) =>
              datasets.reduce((s, ds) => s + (Number(ds.data?.[i]) || 0), 0)
            );
            const metas = datasets.map((_, di) => chart.getDatasetMeta(di));

            totals.forEach((tot, i) => {
              const baseEl = metas[0]?.data?.[i];
              if (!baseEl) return;

              const x = (baseEl.tooltipPosition ? baseEl.tooltipPosition() : baseEl).x;
              const yTopStack = yScale.getPixelForValue(tot);

              // nearest segment label center (only those we actually drew)
              let minSegLabelY = Infinity;
              metas.forEach((m) => {
                const el = m.data?.[i];
                if (!el) return;
                const h = Math.abs((el.base ?? el.y) - el.y);
                if (h >= 18) {
                  const yCenter = (el.y + (el.base ?? el.y)) / 2;
                  minSegLabelY = Math.min(minSegLabelY, yCenter);
                }
              });

              let y = yTopStack - 10;                 // default offset above stack
              if (y > (minSegLabelY - 12)) y = minSegLabelY - 12; // keep clearance

              const text = fmtInt(tot);
              ctx.strokeText(text, x, y);
              ctx.fillText(text, x, y);
            });
          }

          ctx.restore();
        }
      };

      // charts-ready counter for PDF wait
      (function(){ window.__chartsReadyCount = (window.__chartsReadyCount || 0);
                   const need = ${
                     hasAps ? 2 : 1
                   }; window._incReady = function(){
                     window.__chartsReadyCount++; if (window.__chartsReadyCount >= need) window.__chartsReady = true; }; })();

      ${
        hasAps
          ? `
      (function(){
        const ctx = document.getElementById('apsLine').getContext('2d');
        new Chart(ctx,{
          type:'line',
          data:{ labels:${JSON.stringify(apsLabels)},
                 datasets:[{label:'APS', data:${JSON.stringify(
                   apsData
                 )}, tension:.3, pointRadius:3 }]},
          options:{ animation:false, plugins:{ legend:{ display:false } }, scales:{ y:{ beginAtZero:true } } },
          plugins:[dataLabelPlus]
        });
        _incReady();
      })();`
          : ""
      }

      (function(){
        const ctx = document.getElementById('weeklyCat').getContext('2d');
        new Chart(ctx,{
          type:'bar',
          data:{
            labels:${JSON.stringify(weeklyCat.labels)},
            datasets:[
              {label:'Лавлагаа', data:${JSON.stringify(weeklyCat.lav)}},
              {label:'Үйлчилгээ', data:${JSON.stringify(weeklyCat.uil)}},
              {label:'Гомдол',   data:${JSON.stringify(weeklyCat.gom)}},
            ]
          },
          options:{
            animation:false,
            plugins:{ legend:{ position:'bottom' } },
            scales:{ x:{ stacked:true }, y:{ stacked:true, beginAtZero:true } }
          },
          plugins:[dataLabelPlus]
        });
        _incReady();
      })();
    </script>
  </section>`;
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

    // charts ready guard (set by renderAPSLayout)
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
      margin: { top: "14mm", right: "12mm", bottom: "14mm", left: "12mm" },
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

// ────────────────────────────────────────────────────────────────
// HTML wrapper
// ────────────────────────────────────────────────────────────────
function wrapHtml(bodyHtml) {
  const css = fs.readFileSync(CONFIG.CSS_FILE, "utf-8");
  return `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4"></script>
  <style>${css}</style>
</head>
<body>
  <div class="container" style="padding:12px">
    ${bodyHtml}
    <div class="footer">Автоматаар бэлтгэсэн тайлан (Node.js)</div>
  </div>
</body>
</html>`;
}

// ────────────────────────────────────────────────────────────────
async function runOnce() {
  if (!CONFIG.CURR_FILE)
    throw new Error("Missing CURR_FILE (path to current Excel)");
  [CONFIG.CURR_FILE, CONFIG.CSS_FILE].forEach((p) => {
    if (!fs.existsSync(p)) throw new Error(`Missing file: ${p}`);
  });
  if (!fs.existsSync(CONFIG.OUT_DIR))
    fs.mkdirSync(CONFIG.OUT_DIR, { recursive: true });

  // prev auto-detect if needed
  let prevPath = CONFIG.PREV_FILE;
  if (!prevPath) {
    prevPath = autoFindPrevFile(CONFIG.CURR_FILE);
    if (prevPath) console.log(`[auto] PREV_FILE = ${prevPath}`);
  }
  if (!prevPath)
    throw new Error(
      "Prev файл олдсонгүй: PREV_FILE тохируулах эсвэл файлын нэрээ 09.29-10.05 маягаар өгнө үү."
    );

  const wbCurr = xlsx.readFile(CONFIG.CURR_FILE, { cellDates: true });
  const wbPrev = xlsx.readFile(prevPath, { cellDates: true });

  // sections
  const weeklyCat = extractWeeklyByCategoryFromTotaly(wbCurr, "totaly");
  const outbound3w = extractOutbound3Weeks(wbCurr, "totaly");
  const aps = extractAPSLatestMonths(wbCurr, {
    sheetName: CONFIG.ASS_SHEET,
    yearLabel: CONFIG.ASS_YEAR,
    takeLast: CONFIG.ASS_TAKE_LAST_N_MONTHS,
  });
  const top10 = extractOstTop10_APS(wbPrev, wbCurr, {
    sheetName: CONFIG.OST_SHEET,
    company: CONFIG.ASS_COMPANY || "",
  });

  // Cover period text
  const cover = renderAssCover({
    company: CONFIG.ASS_COMPANY || "АРД",
    periodText: `${weeklyCat.labels.at(-2) ?? weeklyCat.labels[0] ?? ""} – ${
      weeklyCat.labels.at(-1) ?? ""
    }`,
  });

  const html = wrapHtml(
    cover + renderAPSLayout({ aps, weeklyCat, top10, outbound3w })
  );

  // PDF
  const monday = dayjs().tz(CONFIG.TIMEZONE).startOf("week").add(1, "day");
  const stamp = monday.format("YYYYMMDD");
  const pdfName = `ard-tetgever-${stamp}.pdf`;
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

  console.log(
    `[OK] Sent ${pdfName} → ${process.env.RECIPIENTS || "(no recipients set)"}`
  );
}

// Scheduler (Mon 09:00)
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
