// Shared constants
const CONFIG = {
  PPTX_NAVY: "12284A",
  PPTX_BLUE: "1E4F9A",
  PPTX_ROWS_PER_SLIDE: 6,
  CONTACT_SLIDE_TEXT: "FOR FURTHER DETAILS:\nDiptesh Saha\nCISO & Practice Head - Cyber Security & Managed Security\nContact No. 7338882888\ndiptesh.s@snsin.com",
  CONTACT_FOOTER_TEXT: "For further details: Diptesh Saha | diptesh.s@snsin.com",
  COMPANY_NAME: "Secure Network Solutions India Pvt Ltd"
};

const fields = {
  customer: "customerName",
  range: "dateRange",
  preparedBy: "preparedBy",
  submittedOn: "submittedOn",
  executiveSummary: "executiveSummary"
};

const STATIC_REVIEWED_BY = "Kishore Kumar";
const STATIC_APPROVED_BY = "Diptesh Saha";

let snsLogoDataUrl = "";
let clientLogoDataUrl = "";
let puzzleLogoDataUrl = "";
let totPotIncChart;
let fortiSiemAlertsChart;
let truePositiveAlertsChart;
let falsePositiveAlertsChart;
let epsTotalChart;
let epsTopHostsChart;
let trendChart;
let ruleSeverityChart;
let responseSlaChart;
let remediationSlaChart;
let totPotPlotFramePluginRegistered = false;
let trendDataLabelsRegistered = false;
let slaPieCenterPluginRegistered = false;

/**
 * PptxGenJS layout names must match the library (lowercase x).
 * LAYOUT_WIDE is 13.33"×7.5" — do not use; export coords assume 10" wide (LAYOUT_16x9 / LAYOUT_4x3).
 * @see https://gitbrent.github.io/PptxGenJS/docs/usage-pres-options/
 */
const pptLayoutConfig = {
  LAYOUT_16X9: { pptxLayout: "LAYOUT_16x9", reportClass: "ppt-wide" },
  LAYOUT_4X3: { pptxLayout: "LAYOUT_4x3", reportClass: "ppt-standard" },
  A4: { pptxLayout: "CUSTOM_A4_PORTRAIT", reportClass: "ppt-a4" }
};

/** Slide size in inches — must match PptxGenJS presets + custom A4 from defineLayout below. */
const PPT_SLIDE_IN = {
  LAYOUT_16X9: { w: 10, h: 5.625 },
  LAYOUT_4X3: { w: 10, h: 7.5 },
  A4: { w: 8.27, h: 11.69 }
};

const PPT_EXPORT_REF = PPT_SLIDE_IN.LAYOUT_16X9;

function getPptSlideInches(layoutKey) {
  return PPT_SLIDE_IN[layoutKey] || PPT_SLIDE_IN.LAYOUT_16X9;
}

function setDynamicPrintPageSize() {
  const layoutKey = getPptLayoutKey();
  const { w, h } = getPptSlideInches(layoutKey);
  const styleId = "dynamicPrintPageSizeStyle";
  let styleEl = document.getElementById(styleId);
  if (!styleEl) {
    styleEl = document.createElement("style");
    styleEl.id = styleId;
    document.head.appendChild(styleEl);
  }
  styleEl.textContent = `@media print { @page { size: ${w}in ${h}in; margin: 0; } }`;
}

function getValue(id) {
  const el = document.getElementById(id);
  return el ? el.value.trim() : "";
}

function toHighResCanvasDataUrl(canvas, scale = 2) {
  if (!canvas || typeof canvas.toDataURL !== "function") return "";
  const width = canvas.width || canvas.clientWidth;
  const height = canvas.height || canvas.clientHeight;
  if (!width || !height) return canvas.toDataURL("image/png");

  const temp = document.createElement("canvas");
  temp.width = Math.max(1, Math.floor(width * scale));
  temp.height = Math.max(1, Math.floor(height * scale));
  const ctx = temp.getContext("2d");
  if (!ctx) return canvas.toDataURL("image/png");
  ctx.setTransform(scale, 0, 0, scale, 0, 0);
  ctx.drawImage(canvas, 0, 0, width, height);
  return temp.toDataURL("image/png");
}

function escapeHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatRichTextHTML(text) {
  if (!text) return "";
  let formatted = escapeHtml(text);
  formatted = formatted.replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>");
  formatted = formatted.replace(/\*(.*?)\*/g, "<em>$1</em>");
  formatted = formatted.replace(/\[MONTH\]/g, `<span class="fw-bold">${getReportMonth()}</span>`);
  return formatted;
}

function parseRichTextPptx(text, baseOptions = {}) {
  if (!text) return [];
  const parts = [];
  const monthReplaced = text.replace(/\[MONTH\]/g, getReportMonth());
  const segments = monthReplaced.split(/\*\*/);
  segments.forEach((segment, idx) => {
    if (segment === "") return;
    const isBold = idx % 2 !== 0;
    parts.push({
      text: segment,
      options: { ...baseOptions, bold: isBold || baseOptions.bold }
    });
  });
  return parts;
}

function parseDateMatch(str) {
  const match = str.match(/(\d{2})-(\d{2})-(\d{4})/);
  if (!match) return null;
  const monthIdx = Number(match[2]) - 1;
  const year = Number(match[3]);
  if (monthIdx < 0 || monthIdx > 11) return null;
  return new Date(year, monthIdx, 1).toLocaleString("en-US", { month: "long", year: "numeric" });
}

function sanitizeFilename(str) {
  return str.replace(/[^a-zA-Z0-9_\-]/g, "_").replace(/_+/g, "_");
}

function getReportMonth() {
  return (
    parseDateMatch(getValue("dateRange")) ||
    parseDateMatch(getValue("submittedOn")) ||
    new Date().toLocaleString("en-US", { month: "long", year: "numeric" })
  );
}

function getChartMonthLabel() {
  const reportMonth = getReportMonth();
  return reportMonth.split(" ")[0] || reportMonth;
}

function parseTrendCsvRows(csvText) {
  return String(csvText || "")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const parts = line.split(",").map((cell) => cell.trim());
      return {
        date: parts[0] || "",
        high: Math.max(0, parseFloat(parts[1]) || 0),
        medium: Math.max(0, parseFloat(parts[2]) || 0)
      };
    });
}

function parseCsvCells(line) {
  const out = [];
  let cur = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') {
        cur += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }
    if (ch === "," && !inQuotes) {
      out.push(cur.trim());
      cur = "";
      continue;
    }
    cur += ch;
  }
  out.push(cur.trim());
  return out;
}

function normalizeKey(text) {
  return String(text || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function parseRiskCsvRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];

  const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
  if (!rows.length) return [];

  const headerNorm = rows[0].map(normalizeKey);
  const hasHeader = headerNorm.some(
    (k) =>
      k.includes("attack") ||
      k.includes("scenario") ||
      k.includes("businessimpact") ||
      k.includes("riskrating") ||
      k.includes("ciatriad") ||
      k.includes("typeofrisk")
  );

  let idx = { attack: 0, scenario: 1, cia: 2, impact: 3, rating: 4 };
  let dataRows = rows;

  if (hasHeader) {
    dataRows = rows.slice(1);
    const findCol = (matchers, fallback) => {
      const pos = headerNorm.findIndex((k) => matchers.some((m) => k.includes(m)));
      return pos >= 0 ? pos : fallback;
    };
    idx = {
      attack: findCol(["attacktype", "attack", "alerttype", "eventname", "rulename"], 0),
      scenario: findCol(["riskscenario", "scenario", "description"], 1),
      cia: findCol(["ciatriad", "typeofrisk", "risktype", "cia"], 2),
      impact: findCol(["businessimpact", "impact", "consequence"], 3),
      rating: findCol(["riskrating", "rating", "severity"], 4)
    };
  }

  return dataRows
    .map((r, i) => ({
      sno: i + 1,
      attackType: r[idx.attack] || "",
      riskScenario: r[idx.scenario] || "",
      ciaTriad: r[idx.cia] || "",
      businessImpact: r[idx.impact] || "",
      riskRating: r[idx.rating] || ""
    }))
    .filter(
      (r) =>
        r.attackType || r.riskScenario || r.ciaTriad || r.businessImpact || r.riskRating
    );
}

function parseRuleSeverityCsvRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];
  const rows = lines.map(parseCsvCells);
  const head = rows[0].map(normalizeKey);
  const hasHeader = head.some(
    (k) => k.includes("rule") || k.includes("high") || k.includes("medium") || k.includes("low")
  );
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findCol = (keys, fallback) => {
    if (!hasHeader) return fallback;
    const idx = head.findIndex((k) => keys.some((s) => k.includes(s)));
    return idx >= 0 ? idx : fallback;
  };
  const iRule = findCol(["rulename", "rule", "attacktype", "eventname"], 0);
  const iHigh = findCol(["high"], 1);
  const iMed = findCol(["medium", "med"], 2);
  const iLow = findCol(["low"], 3);
  return dataRows
    .map((r) => ({
      rule: r[iRule] || "",
      high: Math.max(0, parseFloat(r[iHigh]) || 0),
      medium: Math.max(0, parseFloat(r[iMed]) || 0),
      low: Math.max(0, parseFloat(r[iLow]) || 0)
    }))
    .filter((r) => r.rule);
}

function parseEpsEventsRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];
  const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
  if (!rows.length) return [];
  const head = rows[0].map(normalizeKey);
  const hasHeader = head.some(
    (k) => k.includes("reportingdevice") || k.includes("eventtype") || k.includes("matchedevents")
  );
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findCol = (keys, fallback) => {
    if (!hasHeader) return fallback;
    const idx = head.findIndex((k) => keys.some((s) => k.includes(s)));
    return idx >= 0 ? idx : fallback;
  };
  const iSno = findCol(["sno", "sno."], 0);
  const iDevice = findCol(["reportingdevice", "device"], 1);
  const iType = findCol(["eventtype", "type"], 2);
  const iName = findCol(["eventname", "name"], 3);
  const iCount = findCol(["matchedevents", "count", "events"], 4);
  return dataRows
    .map((r, idx) => ({
      sno: r[iSno] || String(idx + 1),
      device: r[iDevice] || "",
      eventType: r[iType] || "",
      eventName: r[iName] || "",
      matchedEvents: r[iCount] || ""
    }))
    .filter((r) => r.device || r.eventType || r.eventName || r.matchedEvents);
}

function parseSupportMajorRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];
  const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
  const head = rows[0].map(normalizeKey);
  const hasHeader = head.some((k) => k.includes("created") || k.includes("closed") || k.includes("inprocess"));
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findCol = (keys, fallback) => {
    if (!hasHeader) return fallback;
    const idx = head.findIndex((k) => keys.some((s) => k.includes(s)));
    return idx >= 0 ? idx : fallback;
  };
  const iBy = findCol(["createdby", "created", "owner"], 0);
  const iClosed = findCol(["closed"], 1);
  const iInProcess = findCol(["inprocess", "in-progress", "inprogress"], 2);
  return dataRows
    .map((r) => ({
      by: r[iBy] || "",
      closed: Math.max(0, Number(r[iClosed]) || 0),
      inProcess: Math.max(0, Number(r[iInProcess]) || 0)
    }))
    .filter((r) => r.by)
    .slice(0, 20);
}

function parseSupportMinorRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];
  const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
  const head = rows[0].map(normalizeKey);
  const hasHeader = head.some((k) => k.includes("created") || k.includes("closed"));
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findCol = (keys, fallback) => {
    if (!hasHeader) return fallback;
    const idx = head.findIndex((k) => keys.some((s) => k.includes(s)));
    return idx >= 0 ? idx : fallback;
  };
  const iBy = findCol(["createdby", "created", "owner"], 0);
  const iClosed = findCol(["closed"], 1);
  return dataRows
    .map((r) => ({
      by: r[iBy] || "",
      closed: Math.max(0, Number(r[iClosed]) || 0)
    }))
    .filter((r) => r.by)
    .slice(0, 20);
}

function parseInventoryRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];
  const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
  const head = rows[0].map(normalizeKey);
  const hasHeader = head.some((k) => k.includes("device") || k.includes("count"));
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findCol = (keys, fallback) => {
    if (!hasHeader) return fallback;
    const idx = head.findIndex((k) => keys.some((s) => k.includes(s)));
    return idx >= 0 ? idx : fallback;
  };
  const iDevice = findCol(["device", "devicename", "name"], 0);
  const iCount = findCol(["count", "total"], 1);
  return dataRows
    .map((r) => ({
      device: r[iDevice] || "",
      count: Math.max(0, Number(String(r[iCount]).replace(/,/g, "")) || 0)
    }))
    .filter((r) => r.device);
}

function renderIntegratedInventorySlide() {
  const body = document.getElementById("inventoryTableBody");
  const foot = document.getElementById("inventoryTableFoot");
  const note = document.getElementById("inventoryNotePreview");
  if (!body || !foot || !note) return;
  const rows = parseInventoryRows(getValue("inventoryCsv"));
  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="3">No rows found. Upload Inventory CSV.</td></tr>';
    foot.innerHTML = "";
  } else {
    body.innerHTML = rows
      .map(
        (r, i) =>
          `<tr><td>${i + 1}</td><td>${escapeHtml(r.device)}</td><td>${r.count.toLocaleString("en-US")}</td></tr>`
      )
      .join("");
    const total = rows.reduce((a, r) => a + r.count, 0);
    foot.innerHTML = `<tr><td colspan="2">Total</td><td>${total.toLocaleString("en-US")}</td></tr>`;
  }
  note.innerHTML = formatRichTextHTML(getValue("inventoryNote"));
}

function renderSupportTicketsSlide() {
  const majorBody = document.getElementById("supportMajorBody");
  const minorBody = document.getElementById("supportMinorBody");
  if (!majorBody || !minorBody) return;

  const majorRows = parseSupportMajorRows(getValue("supportMajorCsv"));
  const minorRows = parseSupportMinorRows(getValue("supportMinorCsv"));

  if (!majorRows.length) {
    majorBody.innerHTML = '<tr><td colspan="4">No rows found.</td></tr>';
  } else {
    const majorHtml = majorRows
      .map((r) => {
        const total = r.closed + r.inProcess;
        return `<tr><td>${escapeHtml(r.by)}</td><td>${r.closed}</td><td>${r.inProcess}</td><td>${total}</td></tr>`;
      })
      .join("");
    const sumClosed = majorRows.reduce((a, r) => a + r.closed, 0);
    const sumInProcess = majorRows.reduce((a, r) => a + r.inProcess, 0);
    majorBody.innerHTML =
      majorHtml +
      `<tr><td>Grand Total</td><td>${sumClosed}</td><td>${sumInProcess}</td><td>${sumClosed + sumInProcess}</td></tr>`;
  }

  if (!minorRows.length) {
    minorBody.innerHTML = '<tr><td colspan="2">No rows found.</td></tr>';
  } else {
    const minorHtml = minorRows
      .map((r) => `<tr><td>${escapeHtml(r.by)}</td><td>${r.closed}</td></tr>`)
      .join("");
    const sumMinor = minorRows.reduce((a, r) => a + r.closed, 0);
    minorBody.innerHTML = minorHtml + `<tr><td>Grand Total</td><td>${sumMinor}</td></tr>`;
  }
}

function renderEpsEventsTable() {
  const body = document.getElementById("epsEventsTableBody");
  if (!body) return;
  const rows = parseEpsEventsRows(getValue("epsEventsCsv"));
  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="5" class="eps-num">No rows found. Upload EPS Events CSV.</td></tr>';
    return;
  }
  body.innerHTML = rows
    .map((r) => {
      const countNum = Number(String(r.matchedEvents).replace(/,/g, ""));
      const countText = Number.isFinite(countNum) ? countNum.toLocaleString("en-US") : r.matchedEvents;
      return `<tr>
        <td class="eps-num">${escapeHtml(String(r.sno))}</td>
        <td>${escapeHtml(r.device)}</td>
        <td>${escapeHtml(r.eventType)}</td>
        <td>${escapeHtml(r.eventName)}</td>
        <td class="eps-count">${escapeHtml(String(countText))}</td>
      </tr>`;
    })
    .join("");
}

function renderRiskSlides() {
  const host = document.getElementById("riskSlidesContainer");
  if (!host) return;
  host.innerHTML = "";

  const rows = parseRiskCsvRows(getValue("riskCsv"));
  const chunkSize = 3;
  const chunks = rows.length
    ? Array.from({ length: Math.ceil(rows.length / chunkSize) }, (_, i) =>
        rows.slice(i * chunkSize, i * chunkSize + chunkSize)
      )
    : [[]];
  const narrativeHtml = formatRichTextHTML(getValue("riskNarrative"));

  chunks.forEach((chunk, idx) => {
    const section = document.createElement("section");
    section.className = "page page-with-footer risk-slide";
    const title = idx === 0 ? "Potential Alerts - Risks Mitigated" : "Potential Alerts - Risks Mitigated (Contd.)";
    const tableRows = chunk.length
      ? chunk
          .map(
            (r) => `<tr>
  <td class="risk-col-sno">${r.sno}</td>
  <td class="risk-col-attack">${escapeHtml(r.attackType)}</td>
  <td class="risk-col-scenario">${escapeHtml(r.riskScenario)}</td>
  <td class="risk-col-cia">${escapeHtml(r.ciaTriad)}</td>
  <td class="risk-col-impact">${escapeHtml(r.businessImpact)}</td>
  <td class="risk-col-rating">${escapeHtml(r.riskRating)}</td>
</tr>`
          )
          .join("")
      : `<tr><td colspan="6">No risk rows found. Upload Risk CSV.</td></tr>`;

    const headerHtml = `<div class="page-header-right"><img class="nav-sns-logo" alt="SNS Logo" src="${snsLogoDataUrl}" style="display: ${snsLogoDataUrl ? "block" : "none"}" /></div>`;

    section.innerHTML = `
      ${headerHtml}
      <h2 class="revision-title risk-slide-title">${title}</h2>
      ${idx === 0 ? `<div class="narrative risk-slide-narrative">${narrativeHtml}</div>` : ""}
      <div class="risk-table-wrap">
        <table class="risk-table">
          <thead>
            <tr>
              <th class="risk-col-sno">S.No</th>
              <th class="risk-col-attack">Attack Type</th>
              <th class="risk-col-scenario">Risk Scenario</th>
              <th class="risk-col-cia">Type of Risk(s)<br/>CIA Triad</th>
              <th class="risk-col-impact">Potential Business Impact(s)</th>
              <th class="risk-col-rating">Risk Rating</th>
            </tr>
          </thead>
          <tbody>${tableRows}</tbody>
        </table>
      </div>
      <div class="page-footer-bar"></div>
    `;
    host.appendChild(section);
  });
}

function renderRuleSeverityChart() {
  const canvas = document.getElementById("ruleSeverityChart");
  if (!canvas || typeof Chart === "undefined") return;
  const rows = parseRuleSeverityCsvRows(getValue("ruleSeverityCsv"));
  if (ruleSeverityChart) {
    ruleSeverityChart.destroy();
    ruleSeverityChart = undefined;
  }
  if (!rows.length) return;

  const maxVal = Math.max(...rows.map((r) => Math.max(r.high, r.medium, r.low)), 1);
  const axisMax = Math.max(6, Math.ceil(maxVal));
  ensureTrendDataLabelsPlugin();
  const plugins = {
    title: {
      display: true,
      text: "Potential Alert - Severity Categories",
      font: { size: 16, weight: "bold" }
    },
    legend: {
      display: true,
      position: "bottom",
      labels: { boxWidth: 12, boxHeight: 8, usePointStyle: false, color: "#000" }
    }
  };
  if (trendDataLabelsRegistered) {
    plugins.datalabels = {
      display: (ctx) => Number(ctx.dataset.data[ctx.dataIndex]) > 0,
      color: "#000",
      anchor: "end",
      align: "right",
      clamp: true,
      formatter: (v) => Math.round(Number(v)),
      font: { size: 10 }
    };
  }

  ruleSeverityChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: rows.map((r) => r.rule),
      datasets: [
        { label: "High", data: rows.map((r) => r.high), backgroundColor: "#ff0000", borderWidth: 0 },
        { label: "Medium", data: rows.map((r) => r.medium), backgroundColor: "#ffff00", borderWidth: 0 },
        { label: "Low", data: rows.map((r) => r.low), backgroundColor: "#00b050", borderWidth: 0 }
      ]
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins,
      scales: {
        x: {
          beginAtZero: true,
          max: axisMax,
          grid: { color: "#d4d4d4", drawBorder: false },
          ticks: { stepSize: 1, color: "#000" }
        },
        y: {
          grid: { display: false, drawBorder: false },
          ticks: { color: "#000", font: { size: 11 } }
        }
      }
    }
  });
}

function ensureSlaPieCenterPlugin() {
  if (slaPieCenterPluginRegistered || typeof Chart === "undefined") return;
  slaPieCenterPluginRegistered = true;
  Chart.register({
    id: "slaPieCenter",
    afterDraw(chart, args, opts) {
      const txt = opts && opts.text ? String(opts.text) : "";
      if (!txt) return;
      const meta = chart.getDatasetMeta(0);
      if (!meta || !meta.data || !meta.data[0]) return;
      const arc = meta.data[0];
      const { ctx } = chart;
      ctx.save();
      ctx.fillStyle = "#000";
      ctx.font = "bold 12px Calibri, Segoe UI, Arial";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(txt, arc.x, arc.y);
      ctx.restore();
    }
  });
}

function classifySlaScore(rawPct) {
  const pct = Math.max(0, Math.min(100, Number(rawPct) || 0));
  if (pct >= 80) return { pct, label: "Within SLA", color: "#5b9bd5" };
  if (pct >= 30) return { pct, label: "Partially Breached", color: "#ffff00" };
  return { pct, label: "SLA Breached", color: "#ff0000" };
}

/** Light grey used outside the pie (box margin → circle edge); unfilled slice is white inside the circle. */
const SLA_HASH_GREY = "#ececec";

function renderSlaPieChart(canvasId, previousChart, info) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return null;
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;

  const dpr = window.devicePixelRatio || 1;
  const cssW = Math.max(10, canvas.clientWidth || 240);
  const cssH = Math.max(10, canvas.clientHeight || 140);
  canvas.width = Math.round(cssW * dpr);
  canvas.height = Math.round(cssH * dpr);
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, cssW, cssH);

  const cx = cssW * 0.42;
  const cy = cssH * 0.52;
  const r = Math.min(cssW, cssH) * 0.42;
  const start = -Math.PI / 2;
  const pct = Math.max(0, Math.min(100, Number(info.pct) || 0));
  const sweep = (pct / 100) * Math.PI * 2;
  const isFull = pct >= 99.999;

  const frameEl = canvas.parentElement;
  if (frameEl && frameEl.classList.contains("sla-pie-frame")) {
    frameEl.style.background = SLA_HASH_GREY;
  }

  ctx.fillStyle = SLA_HASH_GREY;
  ctx.fillRect(0, 0, cssW, cssH);

  ctx.beginPath();
  ctx.arc(cx, cy, r, 0, Math.PI * 2);
  ctx.fillStyle = "#ffffff";
  ctx.fill();

  if (pct > 0) {
    ctx.beginPath();
    if (isFull) {
      ctx.arc(cx, cy, r, 0, Math.PI * 2);
    } else {
      ctx.moveTo(cx, cy);
      ctx.arc(cx, cy, r, start, start + sweep, false);
      ctx.closePath();
    }
    ctx.fillStyle = info.color;
    ctx.fill();
  }

  ctx.beginPath();
  ctx.arc(cx, cy, r, 0, Math.PI * 2);
  ctx.strokeStyle = "#4c4c4c";
  ctx.lineWidth = 1;
  ctx.stroke();

  ctx.fillStyle = "#222";
  ctx.font = "bold 12px Calibri, Segoe UI, Arial";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(`${Math.round(pct)}%`, cx, cy);

  return null;
}

function renderSlaStatus() {
  const responseInfo = classifySlaScore(getValue("responseSlaPct"));
  const remediationInfo = classifySlaScore(getValue("remediationSlaPct"));
  const incidentCount = Math.max(0, parseInt(getValue("slaIncidentCount"), 10) || 0);
  const closedIncidentCount = Math.max(0, parseInt(getValue("slaClosedIncidentCount"), 10) || 0);
  const customer = getValue("customerName") || "customer";

  responseSlaChart = renderSlaPieChart("responseSlaChart", responseSlaChart, responseInfo);
  remediationSlaChart = renderSlaPieChart("remediationSlaChart", remediationSlaChart, remediationInfo);

  const responseLegend = document.getElementById("responseSlaLegend");
  if (responseLegend) responseLegend.textContent = responseInfo.label;
  const remediationLegend = document.getElementById("remediationSlaLegend");
  if (remediationLegend) remediationLegend.textContent = remediationInfo.label;

  const responseSummary = document.getElementById("responseSlaSummary");
  if (responseSummary) {
    responseSummary.textContent = `Overall, ${incidentCount} incidents were reported to ${customer}, response status: ${responseInfo.label}.`;
    responseSummary.style.color = "#111";
    responseSummary.style.fontWeight = "bold";
  }
  const remediationSummary = document.getElementById("remediationSlaSummary");
  if (remediationSummary) {
    remediationSummary.textContent = `Remediation Time for ${closedIncidentCount} closed tickets status: ${remediationInfo.label}.`;
    remediationSummary.style.color = "#111";
    remediationSummary.style.fontWeight = "bold";
  }
}

function autoCalculateSlaPercentages() {
  const totalEl = document.getElementById("slaIncidentCount");
  const closedEl = document.getElementById("slaClosedIncidentCount");
  const responsePctEl = document.getElementById("responseSlaPct");
  const remediationPctEl = document.getElementById("remediationSlaPct");
  if (!totalEl || !closedEl || !responsePctEl || !remediationPctEl) return;

  const total = Math.max(0, parseInt(totalEl.value, 10) || 0);
  const closed = Math.max(0, parseInt(closedEl.value, 10) || 0);
  const clampedClosed = Math.min(closed, total || closed);

  // Response SLA is tied to total incidents handled for the month.
  const responsePct = total > 0 ? 100 : 0;
  // Remediation SLA tracks closure against total incidents.
  const remediationPct = total > 0 ? Math.round((clampedClosed / total) * 100) : 0;

  responsePctEl.value = String(Math.max(0, Math.min(100, responsePct)));
  remediationPctEl.value = String(Math.max(0, Math.min(100, remediationPct)));
}

function initRiskTableResizing() {
  const tables = document.querySelectorAll(".risk-table");
  tables.forEach((table) => {
    const headerRow = table.querySelector("thead tr");
    if (!headerRow) return;
    const cols = headerRow.querySelectorAll("th");
    for (let i = 0; i < cols.length - 1; i += 1) {
      if (cols[i].querySelector(".resizer")) continue;
      const resizer = document.createElement("div");
      resizer.className = "resizer";
      cols[i].appendChild(resizer);

      let startX = 0;
      let startW = 0;
      let nextStartW = 0;
      const onMouseMove = (e) => {
        const dx = e.clientX - startX;
        const left = Math.max(60, startW + dx);
        const right = Math.max(60, nextStartW - dx);
        cols[i].style.width = `${left}px`;
        cols[i + 1].style.width = `${right}px`;
      };
      const onMouseUp = () => {
        document.removeEventListener("mousemove", onMouseMove);
        document.removeEventListener("mouseup", onMouseUp);
      };
      const onMouseDown = (e) => {
        startX = e.clientX;
        startW = cols[i].offsetWidth;
        nextStartW = cols[i + 1].offsetWidth;
        document.addEventListener("mousemove", onMouseMove);
        document.addEventListener("mouseup", onMouseUp);
      };
      resizer.addEventListener("mousedown", onMouseDown);
    }
  });
}

function getPptLayoutKey() {
  const sel = document.getElementById("pptLayout");
  return sel ? sel.value : "LAYOUT_16X9";
}

/** Keeps #reportRoot in sync with PPT Slide Size (see styles.css .ppt-wide / .ppt-standard / .ppt-a4). */
function syncReportRootLayoutClass() {
  const root = document.getElementById("reportRoot");
  if (!root) return;
  const { reportClass } = pptLayoutConfig[getPptLayoutKey()] || pptLayoutConfig.LAYOUT_16X9;
  root.className = `report ${reportClass}`;
}

function ensureTotPotPlotFramePlugin() {
  if (typeof Chart === "undefined" || totPotPlotFramePluginRegistered) return;
  totPotPlotFramePluginRegistered = true;
  Chart.register({
    id: "totPotPlotFrame",
    afterDraw(chart) {
      if (
        !chart.canvas ||
        ![
          "totPotIncChart",
          "fortiSiemAlertsChart",
          "truePositiveAlertsChart",
          "falsePositiveAlertsChart",
          "epsTotalChart",
          "epsTopHostsChart"
        ].includes(
          chart.canvas.id
        )
      )
        return;
      const { ctx, chartArea } = chart;
      if (!chartArea || chartArea.width <= 0 || chartArea.height <= 0) return;
      ctx.save();
      ctx.strokeStyle = "#000000";
      ctx.lineWidth = 1;
      ctx.strokeRect(
        chartArea.left + 0.5,
        chartArea.top + 0.5,
        chartArea.width - 1,
        chartArea.height - 1
      );
      ctx.restore();
    }
  });
}

function parseInventoryRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];
  const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
  const head = rows[0].map(normalizeKey);
  const hasHeader = head.some((k) => k.includes("device") || k.includes("count"));
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findCol = (keys, fallback) => {
    if (!hasHeader) return fallback;
    const idx = head.findIndex((k) => keys.some((s) => k.includes(s)));
    return idx >= 0 ? idx : fallback;
  };
  const iDevice = findCol(["device", "devicename", "name"], 0);
  const iCount = findCol(["count", "total"], 1);
  return dataRows
    .map((r, i) => ({
      sno: i + 1,
      device: r[iDevice] || "",
      count: Math.max(0, Number(String(r[iCount]).replace(/,/g, "")) || 0)
    }))
    .filter((r) => r.device);
}

function renderIntegratedInventorySlide() {
  const body = document.getElementById("inventoryTableBody");
  const foot = document.getElementById("inventoryTableFoot");
  const note = document.getElementById("inventoryNotePreview");
  if (!body || !foot || !note) return;
  const rows = parseInventoryRows(getValue("inventoryCsv"));
  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="3">No rows found. Upload Inventory CSV.</td></tr>';
    foot.innerHTML = "";
  } else {
    body.innerHTML = rows
      .map(
        (r, i) =>
          `<tr><td>${i + 1}</td><td>${escapeHtml(r.device)}</td><td>${r.count.toLocaleString("en-US")}</td></tr>`
      )
      .join("");
    const total = rows.reduce((a, r) => a + r.count, 0);
    foot.innerHTML = `<tr><td colspan="2">Total</td><td>${total.toLocaleString("en-US")}</td></tr>`;
  }
  note.innerHTML = formatRichTextHTML(getValue("inventoryNote"));
}

function renderTotPotIncChart() {
  const canvas = document.getElementById("totPotIncChart");
  if (!canvas || typeof Chart === "undefined") return;

  const high = Math.max(0, parseFloat(getValue("totPotIncHigh")) || 0);
  const medium = Math.max(0, parseFloat(getValue("totPotIncMedium")) || 0);
  const peak = Math.max(high, medium, 1);
  const yMax = Math.max(10, Math.ceil((peak * 1.1) / 5) * 5);

  if (totPotIncChart) totPotIncChart.destroy();
  ensureTotPotPlotFramePlugin();

  const axisFont = { family: "Calibri, Segoe UI, Arial, sans-serif", size: 11 };

  totPotIncChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: ["High", "Medium"],
      datasets: [
        {
          label: "Incident Count",
          data: [high, medium],
          backgroundColor: ["#ff0000", "#ffff00"],
          borderWidth: 0
        }
      ]
    },
    options: {
      // Default Chart.js bar animation is 1000ms; bars looked "wrong" (e.g. ~3 vs 14) until it finished.
      animation: false,
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { left: 0, right: 2, top: 6, bottom: 2 } },
      elements: { bar: { borderWidth: 0 } },
      plugins: {
        totPotPlotFrame: true,
        legend: { display: false },
        title: { display: false },
        tooltip: {
          enabled: true,
          callbacks: {
            label: (ctx) => `${ctx.label}: ${ctx.parsed.y}`
          }
        }
      },
      datasets: {
        bar: { categoryPercentage: 0.72, barPercentage: 0.72 }
      },
      scales: {
        x: {
          grid: { display: false, drawBorder: false },
          ticks: { display: false }
        },
        y: {
          beginAtZero: true,
          max: yMax,
          min: 0,
          ticks: {
            stepSize: 5,
            font: axisFont,
            color: "#000000",
            padding: 8,
            mirror: false
          },
          grid: {
            display: true,
            color: "#e0e0e0",
            lineWidth: 1,
            drawBorder: false,
            drawTicks: false
          }
        }
      }
    }
  });

  requestAnimationFrame(() => {
    if (totPotIncChart) totPotIncChart.resize();
  });
}

function renderFortiSiemAlertsChart() {
  const canvas = document.getElementById("fortiSiemAlertsChart");
  if (!canvas || typeof Chart === "undefined") return;

  const high = Math.max(0, parseFloat(getValue("fortiAlertHigh")) || 0);
  const medium = Math.max(0, parseFloat(getValue("fortiAlertMedium")) || 0);
  const low = Math.max(0, parseFloat(getValue("fortiAlertLow")) || 0);
  const peak = Math.max(high, medium, low, 1);
  const yMax = Math.max(2000, Math.ceil((peak * 1.1) / 2000) * 2000);
  const monthLabel = getChartMonthLabel();

  if (fortiSiemAlertsChart) fortiSiemAlertsChart.destroy();
  ensureTotPotPlotFramePlugin();

  const axisFont = { family: "Calibri, Segoe UI, Arial, sans-serif", size: 11 };

  fortiSiemAlertsChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: [monthLabel],
      datasets: [
        {
          label: "High",
          data: [high],
          backgroundColor: "#ff0000",
          borderWidth: 0
        },
        { label: "Medium", data: [medium], backgroundColor: "#ffff00", borderWidth: 0 },
        { label: "Low", data: [low], backgroundColor: "#00b050", borderWidth: 0 }
      ]
    },
    options: {
      animation: false,
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { left: 0, right: 4, top: 8, bottom: 4 } },
      elements: { bar: { borderWidth: 0 } },
      plugins: {
        totPotPlotFrame: true,
        legend: {
          display: true,
          position: "bottom",
          labels: { boxWidth: 12, boxHeight: 8, usePointStyle: false, color: "#000", padding: 12 }
        },
        title: { display: false },
        tooltip: {
          enabled: true,
          callbacks: {
            label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y}`
          }
        }
      },
      datasets: {
        bar: { categoryPercentage: 0.65, barPercentage: 0.85 }
      },
      scales: {
        x: {
          grid: { display: false, drawBorder: false },
          ticks: { font: axisFont, color: "#000000" }
        },
        y: {
          beginAtZero: true,
          max: yMax,
          min: 0,
          ticks: {
            stepSize: 2000,
            font: axisFont,
            color: "#000000",
            padding: 8,
            callback: (v) => v.toLocaleString("en-US")
          },
          grid: {
            display: true,
            color: "#e0e0e0",
            lineWidth: 1,
            drawBorder: false,
            drawTicks: false
          }
        }
      }
    }
  });

  requestAnimationFrame(() => {
    if (fortiSiemAlertsChart) fortiSiemAlertsChart.resize();
  });
}

function tfPosPickScale(peak) {
  const p = Math.max(peak * 1.05, 1);
  let step = 200;
  if (p > 6000) step = 2000;
  else if (p > 3000) step = 1000;
  else if (p > 1500) step = 500;
  else step = 200;
  const max = Math.ceil(p / step) * step;
  return { max, step };
}

function fillTfPosPlaceholders(template) {
  if (!template) return "";
  const tpH = Math.max(0, parseFloat(getValue("tfPosTrueHigh")) || 0);
  const tpM = Math.max(0, parseFloat(getValue("tfPosTrueMedium")) || 0);
  const tpL = Math.max(0, parseFloat(getValue("tfPosTrueLow")) || 0);
  const fpH = Math.max(0, parseFloat(getValue("tfPosFalseHigh")) || 0);
  const fpM = Math.max(0, parseFloat(getValue("tfPosFalseMedium")) || 0);
  const fpL = Math.max(0, parseFloat(getValue("tfPosFalseLow")) || 0);
  const fpTotal = fpH + fpM + fpL;
  return template
    .replace(/\[MONTH\]/g, getReportMonth())
    .replace(/\[TP_H\]/g, String(tpH))
    .replace(/\[TP_M\]/g, String(tpM))
    .replace(/\[TP_L\]/g, String(tpL))
    .replace(/\[FP_H\]/g, String(fpH))
    .replace(/\[FP_M\]/g, String(fpM))
    .replace(/\[FP_L\]/g, String(fpL))
    .replace(/\[FP_TOTAL\]/g, String(fpTotal));
}

function parseTfPosNoteLines(filled) {
  const lines = filled.split(/\r?\n/).map((l) => l.trim());
  const rawLines = lines.filter((l) => l.length > 0);
  if (!rawLines.length) return { items: [] };
  let items = [];
  if (/^Note:\s*$/i.test(rawLines[0])) {
    items = rawLines.slice(1);
  } else if (/^Note:\s+/i.test(rawLines[0])) {
    const after = rawLines[0].replace(/^Note:\s+/i, "").trim();
    items = after ? [after, ...rawLines.slice(1)] : rawLines.slice(1);
  } else {
    items = rawLines;
  }
  return { items };
}

function renderTfPosNotePreview(previewId, textareaId) {
  const el = document.getElementById(previewId);
  if (!el) return;
  const filled = fillTfPosPlaceholders(getValue(textareaId));
  const { items } = parseTfPosNoteLines(filled);
  const heading = '<p class="tfpos-note-heading"><strong>Note:</strong></p>';
  if (!items.length) {
    el.innerHTML = heading;
    return;
  }
  const lis = items.map((line) => `<li>${formatRichTextHTML(line)}</li>`).join("");
  el.innerHTML = `${heading}<ul class="tfpos-note-list">${lis}</ul>`;
}

/** Plain text for PPTX: "Note:" then bullet lines. */
function formatTfPosNoteForPptx(textareaId) {
  const filled = fillTfPosPlaceholders(getValue(textareaId));
  const { items } = parseTfPosNoteLines(filled);
  if (!items.length) return "Note:";
  return `Note:\n\n${items.map((line) => `• ${line}`).join("\n")}`;
}

function renderTfPosCharts() {
  const monthLabel = getChartMonthLabel();
  const axisFont = { family: "Calibri, Segoe UI, Arial, sans-serif", size: 10 };

  const build = (canvasId, oldChart, high, medium, low) => {
    const canvas = document.getElementById(canvasId);
    if (!canvas || typeof Chart === "undefined") return undefined;
    const peak = Math.max(high, medium, low, 1);
    const { max: yMax, step: yStep } = tfPosPickScale(peak);
    ensureTotPotPlotFramePlugin();

    if (oldChart) oldChart.destroy();
    return new Chart(canvas, {
      type: "bar",
      data: {
        labels: [monthLabel],
        datasets: [
          { label: "High", data: [high], backgroundColor: "#ff0000", borderWidth: 0 },
          { label: "Medium", data: [medium], backgroundColor: "#ffff00", borderWidth: 0 },
          { label: "Low", data: [low], backgroundColor: "#00b050", borderWidth: 0 }
        ]
      },
      options: {
        animation: false,
        responsive: true,
        maintainAspectRatio: false,
        layout: { padding: { left: 0, right: 2, top: 4, bottom: 2 } },
        elements: { bar: { borderWidth: 0 } },
        plugins: {
          totPotPlotFrame: true,
          legend: {
            display: true,
            position: "bottom",
            labels: { boxWidth: 10, boxHeight: 7, usePointStyle: false, color: "#000", padding: 8, font: { size: 9 } }
          },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${ctx.parsed.y.toLocaleString("en-US")}`
            }
          }
        },
        datasets: {
          bar: { categoryPercentage: 0.65, barPercentage: 0.85 }
        },
        scales: {
          x: {
            grid: { display: false, drawBorder: false },
            ticks: { font: axisFont, color: "#000000", maxRotation: 0 }
          },
          y: {
            beginAtZero: true,
            max: yMax,
            min: 0,
            ticks: {
              stepSize: yStep,
              font: axisFont,
              color: "#000000",
              padding: 6,
              callback: (v) => v.toLocaleString("en-US")
            },
            grid: {
              display: true,
              color: "#e0e0e0",
              lineWidth: 1,
              drawBorder: false,
              drawTicks: false
            }
          }
        }
      }
    });
  };

  const tpH = Math.max(0, parseFloat(getValue("tfPosTrueHigh")) || 0);
  const tpM = Math.max(0, parseFloat(getValue("tfPosTrueMedium")) || 0);
  const tpL = Math.max(0, parseFloat(getValue("tfPosTrueLow")) || 0);
  const fpH = Math.max(0, parseFloat(getValue("tfPosFalseHigh")) || 0);
  const fpM = Math.max(0, parseFloat(getValue("tfPosFalseMedium")) || 0);
  const fpL = Math.max(0, parseFloat(getValue("tfPosFalseLow")) || 0);

  truePositiveAlertsChart = build("truePositiveAlertsChart", truePositiveAlertsChart, tpH, tpM, tpL);
  falsePositiveAlertsChart = build("falsePositiveAlertsChart", falsePositiveAlertsChart, fpH, fpM, fpL);

  requestAnimationFrame(() => {
    if (truePositiveAlertsChart) truePositiveAlertsChart.resize();
    if (falsePositiveAlertsChart) falsePositiveAlertsChart.resize();
  });
}

function parseEpsTopHostsRows(csvText) {
  const lines = String(csvText || "")
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter(Boolean);
  if (!lines.length) return [];
  const rows = lines.map(parseCsvCells).filter((r) => r.some((c) => c !== ""));
  if (!rows.length) return [];
  const head = rows[0].map(normalizeKey);
  const hasHeader = head.some((k) => k.includes("host") || k.includes("device") || k.includes("eps"));
  const dataRows = hasHeader ? rows.slice(1) : rows;
  const findCol = (keys, fallback) => {
    if (!hasHeader) return fallback;
    const idx = head.findIndex((k) => keys.some((s) => k.includes(s)));
    return idx >= 0 ? idx : fallback;
  };
  const iHost = findCol(["host", "device", "reportingdevice", "name"], 0);
  const iEps = findCol(["eps", "avgeps", "value"], 1);
  return dataRows
    .map((r) => ({
      host: r[iHost] || "",
      eps: Math.max(0, Number(r[iEps]) || 0)
    }))
    .filter((r) => r.host)
    .slice(0, 10);
}

function renderEpsTrendPlot() {
  const totalCanvas = document.getElementById("epsTotalChart");
  const hostsCanvas = document.getElementById("epsTopHostsChart");
  if (!totalCanvas || !hostsCanvas || typeof Chart === "undefined") return;

  const monthLabel = getValue("epsMonthLabel") || getChartMonthLabel();
  const totalEps = Math.max(0, Number(getValue("epsTotalValue")) || 0);
  const hostRows = parseEpsTopHostsRows(getValue("epsTopHostsCsv"));
  const rows = hostRows.length ? hostRows : [{ host: "N/A", eps: 0 }];
  const palette = [
    "#4bc0c0",
    "#ffcd56",
    "#ff6384",
    "#ff9f40",
    "#9966ff",
    "#c45891",
    "#4db6ac",
    "#d4a017",
    "#d32f2f",
    "#e66a2c"
  ];

  ensureTrendDataLabelsPlugin();
  ensureTotPotPlotFramePlugin();
  if (epsTotalChart) epsTotalChart.destroy();
  if (epsTopHostsChart) epsTopHostsChart.destroy();

  const totalAxisMax = Math.max(10, Math.ceil((totalEps * 1.08) / 50) * 50);
  epsTotalChart = new Chart(totalCanvas, {
    type: "bar",
    data: {
      labels: ["EPS"],
      datasets: [{ label: monthLabel, data: [totalEps], backgroundColor: "#4bc0c0", borderWidth: 0 }]
    },
    options: {
      animation: false,
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false }, totPotPlotFrame: true },
      scales: {
        x: { grid: { display: false, drawBorder: false }, ticks: { color: "#000", font: { weight: "bold" } } },
        y: {
          beginAtZero: true,
          min: 0,
          max: totalAxisMax,
          ticks: { color: "#000", stepSize: Math.max(10, Math.round(totalAxisMax / 8)) },
          grid: { color: "#d9d9d9", drawBorder: false }
        }
      }
    }
  });

  const hostValues = rows.map((r) => r.eps);
  const hostAxisMax = Math.max(10, Math.ceil((Math.max(...hostValues, 1) * 1.15) / 10) * 10);
  epsTopHostsChart = new Chart(hostsCanvas, {
    type: "bar",
    data: {
      labels: rows.map((_, i) => String(i + 1)),
      datasets: [
        {
          label: "EPS",
          data: hostValues,
          backgroundColor: rows.map((_, i) => palette[i % palette.length]),
          borderWidth: 0
        }
      ]
    },
    options: {
      animation: false,
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        datalabels: {
          display: true,
          color: "#111",
          font: { weight: "bold", size: 10 },
          formatter: (v) => {
            const n = Number(v) || 0;
            return Number.isInteger(n) ? String(n) : n.toFixed(2).replace(/\.?0+$/, "");
          },
          anchor: "end",
          align: "top",
          offset: 2,
          clamp: true,
          clip: false
        },
        tooltip: {
          callbacks: {
            label: (ctx) => `${rows[ctx.dataIndex].host}: ${ctx.parsed.y}`
          }
        },
        totPotPlotFrame: true
      },
      scales: {
        x: {
          title: { display: true, text: "Hosts", color: "#000", font: { weight: "bold" } },
          grid: { display: false, drawBorder: false },
          ticks: { color: "#000", font: { weight: "bold" } }
        },
        y: {
          title: { display: true, text: "EPS", color: "#000", font: { weight: "bold" } },
          beginAtZero: true,
          min: 0,
          max: hostAxisMax,
          ticks: { color: "#000" },
          grid: { color: "#c9c9c9", drawBorder: false }
        }
      }
    }
  });

  const legendHost = document.getElementById("epsHostsLegend");
  if (legendHost) {
    legendHost.innerHTML = rows
      .map(
        (r, i) =>
          `<div class="eps-host-item"><span class="eps-host-color" style="background:${palette[i % palette.length]}"></span><span class="eps-host-text">${i + 1} ${escapeHtml(r.host)}</span></div>`
      )
      .join("");
  }
}

function ensureTrendDataLabelsPlugin() {
  if (trendDataLabelsRegistered || typeof Chart === "undefined") return;
  if (typeof ChartDataLabels === "undefined") return;
  Chart.register(ChartDataLabels);
  trendDataLabelsRegistered = true;
}

function renderTrendChart() {
  const canvas = document.getElementById("trendChart");
  if (!canvas || typeof Chart === "undefined") return;
  const rows = parseTrendCsvRows(getValue("trendCsv"));
  if (!rows.length) {
    if (trendChart) {
      trendChart.destroy();
      trendChart = undefined;
    }
    return;
  }

  const stackedMax = Math.max(...rows.map((r) => r.high + r.medium), 1);
  const yMax = Math.max(6, Math.ceil(stackedMax));
  ensureTrendDataLabelsPlugin();

  if (trendChart) trendChart.destroy();
  const plugins = {
    title: {
      display: true,
      text: `Potential Alert Summary - ${getReportMonth()}`,
      font: { size: 16, weight: "bold" }
    },
    legend: {
      display: true,
      position: "right",
      labels: { boxWidth: 18, boxHeight: 8, usePointStyle: false, color: "#000" },
      title: { display: true, text: "Severity", color: "#000", font: { weight: "bold" } }
    }
  };
  if (trendDataLabelsRegistered) {
    plugins.datalabels = {
      display: (ctx) => Number(ctx.dataset.data[ctx.dataIndex]) > 0,
      color: "#000",
      font: { weight: "bold", size: 10 },
      formatter: (v) => Math.round(Number(v)),
      anchor: "center",
      align: "center",
      clamp: true
    };
  }

  trendChart = new Chart(canvas, {
    type: "bar",
    data: {
      labels: rows.map((r) => r.date),
      datasets: [
        { label: "High", data: rows.map((r) => r.high), backgroundColor: "#ff0000", borderWidth: 0, stack: "s" },
        { label: "Medium", data: rows.map((r) => r.medium), backgroundColor: "#ffff00", borderWidth: 0, stack: "s" }
      ]
    },
    options: {
      animation: false,
      responsive: true,
      maintainAspectRatio: false,
      plugins,
      scales: {
        x: {
          stacked: true,
          title: { display: true, text: "Date", color: "#000" },
          grid: { display: false, drawBorder: false },
          ticks: { color: "#000", maxRotation: 45, minRotation: 45, autoSkip: true, maxTicksLimit: 16 }
        },
        y: {
          stacked: true,
          beginAtZero: true,
          min: 0,
          max: yMax,
          title: { display: true, text: "Count", color: "#000" },
          ticks: { stepSize: 1, color: "#000" },
          grid: { color: "#e0e0e0", drawBorder: false }
        }
      }
    }
  });

  requestAnimationFrame(() => {
    if (trendChart) trendChart.resize();
  });
}

function updateLogos() {
  const snsEl = document.getElementById("snsLogoPreview");
  const clientEl = document.getElementById("clientLogoPreview");
  const contactSnsEl = document.getElementById("contactSnsLogoPreview");
  if (snsEl) {
    snsEl.style.display = snsLogoDataUrl ? "block" : "none";
    snsEl.src = snsLogoDataUrl;
  }
  if (clientEl) {
    clientEl.style.display = clientLogoDataUrl ? "block" : "none";
    clientEl.src = clientLogoDataUrl;
  }
  if (contactSnsEl) {
    contactSnsEl.style.display = snsLogoDataUrl ? "block" : "none";
    contactSnsEl.src = snsLogoDataUrl;
  }
  document.querySelectorAll(".nav-sns-logo").forEach((img) => {
    img.src = snsLogoDataUrl;
    img.style.display = snsLogoDataUrl ? "block" : "none";
  });
}

function applyPremiumLayout() {
  document.querySelectorAll(".page").forEach((page) => {
    if (
      page.classList.contains("cover-hero") ||
      page.querySelector(".engagement-container") ||
      page.classList.contains("revision-page") ||
      page.classList.contains("slide-total-potential") ||
      page.classList.contains("slide-fortisiem-alerts") ||
      page.classList.contains("slide-tfpos-alerts") ||
      page.classList.contains("slide-eps-trend-plot") ||
      page.classList.contains("slide-eps-events") ||
      page.classList.contains("slide-support-tickets") ||
      page.classList.contains("slide-integrated-inventory") ||
      page.classList.contains("slide-key-points-summary") ||
      page.classList.contains("slide-contact-summary")
    ) {
      return;
    }
    if (!page.classList.contains("page-with-footer")) {
      page.classList.add("page-with-footer");
    }
    if (!page.querySelector(".page-header-right")) {
      const header = document.createElement("div");
      header.className = "page-header-right";
      header.innerHTML = '<img class="nav-sns-logo" alt="SNS Logo" />';
      page.insertBefore(header, page.firstChild);
    }
    if (!page.querySelector(".page-footer-bar")) {
      const footer = document.createElement("div");
      footer.className = "page-footer-bar";
      page.appendChild(footer);
    }
    const h2 = page.querySelector("h2");
    if (h2 && !h2.classList.contains("revision-title")) {
      h2.classList.add("revision-title");
      h2.style.cssText =
        "margin-left: 0; margin-bottom: 20px; font-family: Calibri, 'Segoe UI', Arial, sans-serif; font-size: 38px;";
    }
  });
  updateLogos();
}

function applyData() {
  setDynamicPrintPageSize();
  syncReportRootLayoutClass();
  const reportRoot = document.getElementById("reportRoot");
  if (reportRoot) void reportRoot.offsetHeight;

  Object.entries(fields).forEach(([key, inputId]) => {
    const value = getValue(inputId);
    document.querySelectorAll(`[data-field="${key}"]`).forEach((el) => {
      el.textContent = value;
    });
  });
  document.querySelectorAll('[data-field="reviewedBy"]').forEach((el) => {
    el.textContent = STATIC_REVIEWED_BY;
  });
  document.querySelectorAll('[data-field="approvedBy"]').forEach((el) => {
    el.textContent = STATIC_APPROVED_BY;
  });
  document.querySelectorAll('[data-field="month"]').forEach((el) => {
    el.textContent = getReportMonth();
  });

  const execEl = document.getElementById("executiveSummaryPreview");
  if (execEl) {
    execEl.innerHTML = getValue("executiveSummary")
      .split("\n")
      .filter((p) => p.trim())
      .map((p) => `<p>${formatRichTextHTML(p)}</p>`)
      .join("");
  }

  const puzzleEl = document.getElementById("puzzleGraphicPreview");
  if (puzzleEl && puzzleLogoDataUrl) {
    puzzleEl.src = puzzleLogoDataUrl;
    puzzleEl.style.display = "block";
    const ph = document.getElementById("puzzlePlaceholderLayout");
    if (ph) ph.style.display = "none";
  } else if (puzzleEl && !puzzleLogoDataUrl) {
    puzzleEl.style.display = "none";
    const ph = document.getElementById("puzzlePlaceholderLayout");
    if (ph) ph.style.display = "flex";
  }

  const monthLbl = getChartMonthLabel();
  const hi = getValue("totPotIncHigh");
  const med = getValue("totPotIncMedium");
  const hdr = document.getElementById("totPotIncTableColHdr");
  if (hdr) hdr.textContent = monthLbl;
  const th = document.getElementById("totPotIncTableHigh");
  if (th) th.textContent = hi !== "" ? hi : "0";
  const tm = document.getElementById("totPotIncTableMed");
  if (tm) tm.textContent = med !== "" ? med : "0";
  const trendNote = document.getElementById("trendNotePreview");
  if (trendNote) {
    trendNote.innerHTML = formatRichTextHTML(getValue("trendNote"));
  }
  const trendNarr = document.getElementById("trendNarrativePreview");
  if (trendNarr) {
    trendNarr.innerHTML = formatRichTextHTML(getValue("trendNarrative"));
  }
  updateLogos();
  renderRiskSlides();
  applyPremiumLayout();
  renderTotPotIncChart();
  renderTrendChart();
  renderRuleSeverityChart();
  renderSlaStatus();
  renderIntegratedInventorySlide();
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      if (totPotIncChart) totPotIncChart.resize();
      if (trendChart) trendChart.resize();
      window.dispatchEvent(new Event("resize"));
    });
  });
}

async function exportPptx() {
  applyData();
  await new Promise((r) => setTimeout(r, 900));

  const layoutKey = getPptLayoutKey();
  const { w: slideW, h: slideH } = getPptSlideInches(layoutKey);
  const sx = slideW / PPT_EXPORT_REF.w;
  const sy = slideH / PPT_EXPORT_REF.h;

  const pptx = new PptxGenJS();
  const layoutCfg = pptLayoutConfig[layoutKey] || pptLayoutConfig.LAYOUT_16X9;
  if (layoutKey === "A4") {
    pptx.defineLayout({ name: "CUSTOM_A4_PORTRAIT", width: PPT_SLIDE_IN.A4.w, height: PPT_SLIDE_IN.A4.h });
    pptx.layout = "CUSTOM_A4_PORTRAIT";
  } else {
    pptx.layout = layoutCfg.pptxLayout;
  }
  pptx.title = `Nocil Limited Monthly SOC Report`;

  const margin = 0.5 * sx;
  const contentW = slideW - 2 * margin;
  const logoW = 1.2;
  const logoX = slideW - logoW - 0.15 * sx;

  const titleStyle = { fontFace: "Calibri" };
  const barChartType = pptx.ChartType ? pptx.ChartType.bar : "bar";

  let slide = pptx.addSlide();
  // Cover background to match web preview tone
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 0,
    w: slideW,
    h: slideH,
    line: { color: "FAFAFC", pt: 0 },
    fill: { color: "FAFAFC" }
  });
  // Client logo card (same concept as web card)
  slide.addShape(pptx.ShapeType.rect, {
    x: 2.7 * sx,
    y: 1.0 * sy,
    w: 4.6 * sx,
    h: 1.55 * sy,
    line: { color: "EFEFF2", pt: 1 },
    fill: { color: "FFFFFF" }
  });
  slide.addText("MANAGED INCIDENT RESPONSE &\nREMEDIATION SERVICE", {
    x: margin,
    y: 3.15 * sy,
    w: contentW,
    h: 1.0 * sy,
    fontSize: 24,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  slide.addText(`Monthly SOC Report | ${getReportMonth()}`, {
    x: 1.4 * sx,
    y: 4.6 * sy,
    w: 7.2 * sx,
    h: 0.4 * sy,
    fontSize: 16,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (clientLogoDataUrl) {
    slide.addImage({ data: clientLogoDataUrl, x: 3.15 * sx, y: 1.2 * sy, w: 3.7 * sx, h: 1.15 * sy });
  }
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.2 * sy, w: logoW, h: 0.5 });
  }
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 5.05 * sy,
    w: slideW,
    h: 0.575 * sy,
    line: { color: "042E5F", pt: 0 },
    fill: { color: "042E5F" }
  });

  slide = pptx.addSlide();
  slide.addText("Document Revision History", {
    x: 0.7 * sx,
    y: 0.85 * sy,
    w: 8.6 * sx,
    h: 0.55 * sy,
    fontSize: 24,
    bold: true,
    color: "000000",
    align: "left",
    ...titleStyle
  });
  const revTable = [
    [{ text: "Version", options: { bold: true, fill: { color: "FFFFFF" } } }, { text: "1.0", options: { bold: true, fill: { color: "FFFFFF" } } }],
    [{ text: "Dated", options: { fill: { color: "E1F4F4" } } }, { text: getValue("dateRange"), options: { fill: { color: "E1F4F4" } } }],
    [{ text: "Prepared By", options: { fill: { color: "FFFFFF" } } }, { text: getValue("preparedBy"), options: { fill: { color: "FFFFFF" } } }],
    [{ text: "Reviewed By", options: { fill: { color: "E1F4F4" } } }, { text: STATIC_REVIEWED_BY, options: { fill: { color: "E1F4F4" } } }],
    [{ text: "Approved By", options: { fill: { color: "FFFFFF" } } }, { text: STATIC_APPROVED_BY, options: { fill: { color: "FFFFFF" } } }],
    [{ text: "Submitted On", options: { fill: { color: "E1F4F4" } } }, { text: getValue("submittedOn"), options: { fill: { color: "E1F4F4" } } }]
  ];
  slide.addTable(revTable, {
    x: 0.7 * sx,
    y: 1.65 * sy,
    w: 8.6 * sx,
    colW: [2.8 * sx, 5.8 * sx],
    fontSize: 12,
    border: { pt: 0, color: "FFFFFF" },
    color: "111111",
    valign: "mid"
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.2 * sy, w: logoW, h: 0.5 });
  }
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 4.95 * sy,
    w: slideW,
    h: 0.675 * sy,
    line: { color: "1E4F9A", pt: 0 },
    fill: { color: "1E4F9A" }
  });

  slide = pptx.addSlide();
  slide.addText("The Engagement", {
    x: margin,
    y: 0.4 * sy,
    w: contentW,
    h: 0.55 * sy,
    fontSize: 28,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  slide.addText(
    [{ text: `${getValue("customerName") || "Customer"} has engaged with SNS to monitor and review the entity's security.`, options: { bold: true } }],
    {
      x: 0.65 * sx,
      y: 1.2 * sy,
      w: 5.2 * sx,
      h: 0.5 * sy,
      fontSize: 11,
      color: "111111",
      ...titleStyle
    }
  );
  slide.addText(parseRichTextPptx(getValue("executiveSummary")), {
    x: margin,
    y: 1.8 * sy,
    w: 5.3 * sx,
    h: 3.45 * sy,
    fontSize: 10.5,
    color: "000000",
    valign: "top",
    ...titleStyle
  });
  if (puzzleLogoDataUrl) {
    slide.addImage({ data: puzzleLogoDataUrl, x: 6.2 * sx, y: 1.6 * sy, w: 3.1 * sx, h: 3.0 * sy });
  }
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.2 * sy, w: logoW, h: 0.5 });
  }
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: 4.95 * sy,
    w: slideW,
    h: 0.675 * sy,
    line: { color: "1E4F9A", pt: 0 },
    fill: { color: "1E4F9A" }
  });

  const monthLabel = getChartMonthLabel();
  const highVal = Math.max(0, parseFloat(getValue("totPotIncHigh")) || 0);
  const medVal = Math.max(0, parseFloat(getValue("totPotIncMedium")) || 0);
  const maxAlert = Math.max(highVal, medVal, 1);
  const yAxisMax = Math.max(20, Math.ceil((maxAlert * 1.1) / 2) * 2);

  // Slide 4: Total Potential Alerts
  slide = pptx.addSlide();
  slide.addText("Total Potential Alerts", {
    x: margin,
    y: 0.4 * sy,
    w: contentW,
    h: 0.55 * sy,
    fontSize: 28,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  slide.addChart(
    barChartType,
    [
      { name: "High", labels: [monthLabel], values: [highVal] },
      { name: "Medium", labels: [monthLabel], values: [medVal] }
    ],
    {
      x: 0.7 * sx,
      y: 0.95 * sy,
      w: 8.6 * sx,
      h: 2.95 * sy,
      barDir: "col",
      barGrouping: "clustered",
      chartColors: ["FF0000", "FFFF00"],
      showLegend: true,
      legendPos: "b",
      valAxisMaxVal: yAxisMax,
      valAxisMinVal: 0,
      valAxisMajorUnit: 2,
      showDataTable: true,
      dataTableFontSize: 9,
      catAxisLabelFontSize: 10,
      valAxisLabelFontSize: 10
    }
  );

  // Slide 5: Potential Alert Tickets Trend
  slide = pptx.addSlide();
  slide.addText("Potential Alert Tickets Trend", {
    x: margin,
    y: 0.4 * sy,
    w: contentW,
    h: 0.55 * sy,
    fontSize: 28,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  const trendCanvasForPpt = document.getElementById("trendChart");
  if (trendCanvasForPpt) {
    slide.addImage({ data: toHighResCanvasDataUrl(trendCanvasForPpt, 2), x: 0.45 * sx, y: 0.9 * sy, w: 9.1 * sx, h: 3.3 * sy });
  }
  slide.addText(parseRichTextPptx(getValue("trendNote")), {
    x: 0.5 * sx,
    y: 4.35 * sy,
    w: 9.0 * sx,
    h: 0.3 * sy,
    fontSize: 12,
    bold: true,
    color: "000000",
    fontFace: "Calibri"
  });
  slide.addText(parseRichTextPptx(getValue("trendNarrative")), {
    x: 0.5 * sx,
    y: 4.7 * sy,
    w: 9.0 * sx,
    h: 0.7 * sy,
    fontSize: 10.5,
    color: "000000",
    fontFace: "Calibri"
  });


  // Slide 6: Risk Mitigation
  const riskRowsPptx = parseRiskCsvRows(getValue("riskCsv"));
  const riskChunksPptx = riskRowsPptx.length
    ? Array.from({ length: Math.ceil(riskRowsPptx.length / 3) }, (_, i) =>
        riskRowsPptx.slice(i * 3, i * 3 + 3)
      )
    : [[]];
  const riskSlideTitle = "Potential Alerts - Risks Mitigated";
  const narrative = parseRichTextPptx(getValue("riskNarrative"));

  riskChunksPptx.forEach((rows, i) => {
    slide = pptx.addSlide();
    slide.addText(i === 0 ? riskSlideTitle : `${riskSlideTitle} (Contd.)`, {
      x: margin,
      y: 0.6 * sy,
      w: contentW,
      h: 0.5 * sy,
      ...titleStyle,
      align: "center",
      fontSize: 24,
      bold: true,
      color: "000000"
    });
    if (snsLogoDataUrl) {
      slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
    }

    let tableY = 2.05 * sy;
    if (i === 0) {
      slide.addText(narrative, {
        x: margin,
        y: 1.25 * sy,
        w: contentW,
        fontSize: 10.5,
        color: "111111",
        fontFace: "Calibri",
        align: "justify"
      });
      tableY = 2.05 * sy;
    }

    const riskTableData = [
      [
        { text: "S.No", options: { bold: true, fill: { color: "D9D9D9" }, align: "center" } },
        { text: "Type of Attack", options: { bold: true, fill: { color: "D9D9D9" }, align: "center" } },
        { text: "Scenario", options: { bold: true, fill: { color: "D9D9D9" }, align: "center" } },
        { text: "Impact (CIA)", options: { bold: true, fill: { color: "D9D9D9" }, align: "center" } },
        { text: "Business Impact/ Remediation", options: { bold: true, fill: { color: "D9D9D9" }, align: "center" } },
        { text: "Risk Rating", options: { bold: true, fill: { color: "D9D9D9" }, align: "center" } }
      ]
    ];
    
    rows.forEach((r) => {
        riskTableData.push([
          { text: String(r.sno), options: { align: "center" } },
          { text: r.attackType, options: { align: "left" } },
          { text: r.riskScenario, options: { align: "left" } },
          { text: r.ciaTriad, options: { align: "center" } },
          { text: r.businessImpact, options: { align: "left" } },
          { text: r.riskRating, options: { align: "center" } }
        ]);
    });

    for (let fillIdx = rows.length; fillIdx < 3; fillIdx++) {
      riskTableData.push([
        { text: "-", options: { align: "center" } },
        { text: "", options: { align: "left" } },
        { text: "", options: { align: "left" } },
        { text: "", options: { align: "center" } },
        { text: "", options: { align: "left" } },
        { text: "", options: { align: "left" } }
      ]);
    }
    
    slide.addTable(riskTableData, {
      x: margin,
      y: tableY,
      w: contentW,
      fontSize: 8.8,
      colW: [0.45 * sx, 1.45 * sx, 1.75 * sx, 1.25 * sx, 3.7 * sx, 0.9 * sx],
      border: { pt: 0, color: "FFFFFF" }
    });
  });

  // Slide 7: Rule-Based Severity Categories
  const ruleRows = parseRuleSeverityCsvRows(getValue("ruleSeverityCsv"));
  if (ruleRows.length) {
    slide = pptx.addSlide();
    slide.addText("Rule-Based Severity Categories For Potential Alerts", {
      x: margin,
      y: 0.6 * sy,
      w: contentW,
      h: 0.55 * sy,
      fontSize: 28,
      bold: true,
      color: "000000",
      align: "center",
      ...titleStyle
    });
    if (snsLogoDataUrl) {
      slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
    }

    const ruleCanvas = document.getElementById("ruleSeverityChart");
    if (ruleCanvas) {
      slide.addImage({
        data: ruleCanvas.toDataURL("image/png"),
        x: 0.5 * sx,
        y: 1.4 * sy,
        w: 9.0 * sx,
        h: 3.8 * sy
      });
    }
  }

  // Slide 8: SLA Performance
  const responseSlaPct = Number(getValue("responseSlaPct")) || 0;
  const remediationSlaPct = Number(getValue("remediationSlaPct")) || 0;
  const responseInfo = classifySlaScore(responseSlaPct);
  const remediationInfo = classifySlaScore(remediationSlaPct);
  const incidentCount = Math.max(0, parseInt(getValue("slaIncidentCount"), 10) || 0);
  const closedIncidentCount = Math.max(0, parseInt(getValue("slaClosedIncidentCount"), 10) || 0);
  const customer = getValue("customerName") || "customer";

  slide = pptx.addSlide();
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  slide.addText("Response Time SLA", {
    x: 0.5 * sx,
    y: 1.0 * sy, /* Shifted down from 0.6 */
    w: 2.8 * sx,
    h: 0.4 * sy,
    fontSize: 22,
    bold: true,
    color: "000000",
    fontFace: "Calibri"
  });
  slide.addText("Remediation Time SLA", {
    x: 0.5 * sx,
    y: 3.4 * sy, /* Shifted down from 3.0 */
    w: 3.2 * sx,
    h: 0.4 * sy,
    fontSize: 22,
    bold: true,
    color: "000000",
    fontFace: "Calibri"
  });

  const responseCanvas = document.getElementById("responseSlaChart");
  if (responseCanvas) {
    slide.addImage({
      data: responseCanvas.toDataURL("image/png"),
      x: 0.35 * sx,
      y: 1.4 * sy,
      w: 2.9 * sx,
      h: 1.8 * sy
    });
  }
  const remediationCanvas = document.getElementById("remediationSlaChart");
  if (remediationCanvas) {
    slide.addImage({
      data: remediationCanvas.toDataURL("image/png"),
      x: 0.35 * sx,
      y: 3.85 * sy,
      w: 2.9 * sx,
      h: 1.8 * sy
    });
  }

  const responseTbl = [
    [
      { text: "Severity Level Description", options: { bold: true, fill: { color: "FFFFFF" }, align: "center" } },
      { text: "Severity Level", options: { bold: true, fill: { color: "FFFFFF" }, align: "center" } },
      { text: "Response", options: { bold: true, fill: { color: "FFFFFF" }, align: "center" } },
      { text: "Resolution/Remediation", options: { bold: true, fill: { color: "FFFFFF" }, align: "center" } }
    ],
    [
      { text: "High Severity ticket resulted in extremely serious interruptions to Business system. It has affected the user community", options: { fill: { color: "EEF6F9" } } },
      { text: "S-1", options: { align: "center", fill: { color: "EEF6F9" } } },
      { text: "≤15 Minutes", options: { align: "center", fill: { color: "EEF6F9" } } },
      { text: "4 Hours", options: { align: "center", fill: { color: "EEF6F9" } } }
    ],
    [
      "Medium Severity ticket resulted in interruptions normal operations. It does not prevent Business Operations or minor ticket",
      { text: "S-2", options: { align: "center" } },
      { text: "≤15 Minutes", options: { align: "center" } },
      { text: "8 Hours", options: { align: "center" } }
    ],
    [
      { text: "Low Severity A Request/Query/Service that does not change existing service structure and no cost implication", options: { fill: { color: "EEF6F9" } } },
      { text: "S-3", options: { align: "center", fill: { color: "EEF6F9" } } },
      { text: "≤30 Minutes", options: { align: "center", fill: { color: "EEF6F9" } } },
      { text: "24 Hours", options: { align: "center", fill: { color: "EEF6F9" } } }
    ]
  ];
  slide.addTable(responseTbl, {
    x: 3.45 * sx,
    y: 1.15 * sy, /* Shifted down from 0.75 */
    w: 6.2 * sx,
    fontSize: 8.5,
    colW: [3.8 * sx, 1.1 * sx, 1.3 * sx, 1.35 * sx],
    border: { pt: 0.6, color: "C8D8DD" },
    valign: "middle"
  });
  slide.addText(`Overall, ${incidentCount} incidents were reported to ${customer}, response status: ${responseInfo.label}.`, {
    x: 3.45 * sx,
    y: 3.35 * sy, /* Shifted down from 2.95 */
    w: 6.2 * sx,
    h: 0.3 * sy,
    fontSize: 10,
    bold: true,
    color: "111111",
    fontFace: "Calibri"
  });

  const remediationTbl = [
    [
      { text: "Severity Level Description", options: { bold: true, fill: { color: "FFFFFF" }, align: "center" } },
      { text: "Severity Level", options: { bold: true, fill: { color: "FFFFFF" }, align: "center" } },
      { text: "Resolution/Remediation", options: { bold: true, fill: { color: "FFFFFF" }, align: "center" } }
    ],
    [
      { text: "High Severity ticket resulted in extremely serious interruptions to Business system. It has affected the user community", options: { fill: { color: "EEF6F9" } } },
      { text: "High", options: { align: "center", fill: { color: "EEF6F9" } } },
      { text: "4 Hours", options: { align: "center", fill: { color: "EEF6F9" } } }
    ],
    [
      "Medium Severity ticket resulted in interruptions normal operations. It does not prevent Business Operations or minor ticket",
      { text: "Medium", options: { align: "center" } },
      { text: "8 Hours", options: { align: "center" } }
    ],
    [
      { text: "Low A Request/Query/Service that does not change existing service structure and no cost implication", options: { fill: { color: "EEF6F9" } } },
      { text: "Low", options: { align: "center", fill: { color: "EEF6F9" } } },
      { text: "24 Hours", options: { align: "center", fill: { color: "EEF6F9" } } }
    ]
  ];
  slide.addTable(remediationTbl, {
    x: 3.45 * sx,
    y: 3.8 * sy, /* Increased from 3.65 */
    w: 6.2 * sx,
    fontSize: 8.5,
    colW: [4.2 * sx, 1.1 * sx, 1.3 * sx],
    border: { pt: 0.6, color: "C8D8DD" },
    valign: "middle"
  });
  slide.addText(`Remediation Time for ${closedIncidentCount} closed tickets status: ${remediationInfo.label}.`, {
    x: 3.45 * sx,
    y: 5.7 * sy, /* Increased from 5.55 */
    w: 6.2 * sx,
    h: 0.3 * sy,
    fontSize: 10,
    bold: true,
    color: "111111",
    fontFace: "Calibri"
  });

  // Slide 9: Integrated Device Inventory
  slide = pptx.addSlide();
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  slide.addText("Integrated Device Inventory", {
    x: margin,
    y: 0.4 * sy,
    w: contentW,
    h: 0.55 * sy,
    fontSize: 32,
    bold: true,
    color: "000000",
    align: "center",
    fontFace: "Calibri"
  });

  const invRows = parseInventoryRows(getValue("inventoryCsv"));
  const invTotal = invRows.reduce((a, r) => a + r.count, 0);
  const invTbl = [
    [
      { text: "S.NO", options: { bold: true, fill: { color: "40C4CC" }, color: "FFFFFF", align: "center" } },
      { text: "DEVICE NAME", options: { bold: true, fill: { color: "40C4CC" }, color: "FFFFFF", align: "center" } },
      { text: "COUNT", options: { bold: true, fill: { color: "40C4CC" }, color: "FFFFFF", align: "center" } }
    ],
    ...invRows.map((r, idx) => [
      { text: String(r.sno), options: { align: "center", fill: { color: idx % 2 !== 0 ? "F2F9FA" : "FFFFFF" } } },
      { text: r.device, options: { align: "center", fill: { color: idx % 2 !== 0 ? "F2F9FA" : "FFFFFF" } } },
      { text: r.count.toLocaleString("en-US"), options: { align: "center", fill: { color: idx % 2 !== 0 ? "F2F9FA" : "FFFFFF" } } }
    ]),
    [
      { text: "Total", options: { bold: true, fill: { color: "40C4CC" }, color: "FFFFFF", align: "center" }, colspan: 2 },
      { text: invTotal.toLocaleString("en-US"), options: { bold: true, fill: { color: "40C4CC" }, color: "FFFFFF", align: "center" } }
    ]
  ];

  slide.addTable(invTbl, {
    x: 1.5 * sx,
    y: 1.4 * sy,
    w: 7.0 * sx,
    fontSize: 14,
    colW: [1.0 * sx, 4.0 * sx, 2.0 * sx],
    border: { pt: 0.5, color: "D9EDF0" },
    valign: "middle"
  });

  const invNoteRaw = getValue("inventoryNote");
  slide.addText(invNoteRaw, {
    x: 0.5 * sx,
    y: 4.4 * sy,
    w: 9.0 * sx,
    fontSize: 18,
    color: "000000",
    align: "left",
    fontFace: "Calibri"
  });

  // Final contact slide.
  slide = pptx.addSlide();
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: 1.0 * sx, y: 1.35 * sy, w: 1.65 * sx, h: 0.95 * sy });
  }
  slide.addText("Your Trusted Security Advisor", {
    x: 1.0 * sx,
    y: 2.55 * sy,
    w: 7.8 * sx,
    h: 0.7 * sy,
    fontSize: 31,
    bold: true,
    color: "082A6F",
    fontFace: "Calibri"
  });
  slide.addText(CONFIG.CONTACT_SLIDE_TEXT, {
    x: 1.1 * sx,
    y: 3.32 * sy,
    w: 8.3 * sx,
    h: 1.65 * sy,
    fontSize: 14.5,
    bold: true,
    color: "1A6A9B",
    fontFace: "Calibri",
    breakLine: true
  });

  await pptx.writeFile(`Nocil_Monthly_SOC_Report.pptx`);
}

function bindImageInput(inputId, setter) {
  const input = document.getElementById(inputId);
  if (!input) return;
  input.addEventListener("change", (event) => {
    const file = event.target.files && event.target.files[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = (e) => {
        setter(String(e.target.result));
        applyData();
      };
      reader.readAsDataURL(file);
    }
  });
}

function bindCsvInput(fileInputId, textareaId) {
  const input = document.getElementById(fileInputId);
  const target = document.getElementById(textareaId);
  if (!input || !target) return;
  input.addEventListener("change", (event) => {
    const file = event.target.files && event.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      target.value = String(e.target.result || "");
      applyData();
    };
    reader.readAsText(file);
  });
}

async function init() {
  let panelApplyScheduled = false;
  function schedulePanelApply() {
    if (panelApplyScheduled) return;
    panelApplyScheduled = true;
    requestAnimationFrame(() => {
      panelApplyScheduled = false;
      applyData();
    });
  }
  document.querySelectorAll(".sidebar input, .sidebar textarea, .sidebar select").forEach((el) => {
    el.addEventListener("input", schedulePanelApply);
    el.addEventListener("change", schedulePanelApply);
  });
  const slaIncidentCountEl = document.getElementById("slaIncidentCount");
  const slaClosedIncidentCountEl = document.getElementById("slaClosedIncidentCount");
  const onSlaBaseChange = () => {
    autoCalculateSlaPercentages();
    schedulePanelApply();
  };
  if (slaIncidentCountEl) {
    slaIncidentCountEl.addEventListener("input", onSlaBaseChange);
    slaIncidentCountEl.addEventListener("change", onSlaBaseChange);
  }
  if (slaClosedIncidentCountEl) {
    slaClosedIncidentCountEl.addEventListener("input", onSlaBaseChange);
    slaClosedIncidentCountEl.addEventListener("change", onSlaBaseChange);
  }
  document.getElementById("printBtn").addEventListener("click", async () => {
    applyData();
    setDynamicPrintPageSize();
    await new Promise((resolve) => requestAnimationFrame(resolve));
    window.print();
  });
  document.getElementById("pptxBtn").addEventListener("click", exportPptx);
  
  const menuToggle = document.getElementById("menuToggle");
  const appShell = document.getElementById("appShell");
  if (menuToggle && appShell) {
    menuToggle.addEventListener("click", () => {
      appShell.classList.toggle("sidebar-collapsed");
    });
  }

  ["snsLogoInput", "clientLogoInput", "puzzleLogoInput"].forEach((id) => {
    bindImageInput(id, (v) => {
      if (id === "snsLogoInput") snsLogoDataUrl = v;
      if (id === "clientLogoInput") clientLogoDataUrl = v;
      if (id === "puzzleLogoInput") puzzleLogoDataUrl = v;
    });
  });
  bindCsvInput("trendCsvFile", "trendCsv");
  bindCsvInput("riskCsvFile", "riskCsv");
  bindCsvInput("ruleSeverityCsvFile", "ruleSeverityCsv");
  bindCsvInput("inventoryCsvFile", "inventoryCsv");

  autoCalculateSlaPercentages();
  applyData();
}

window.onload = init;
