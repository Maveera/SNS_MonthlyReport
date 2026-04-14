// Shared constants — update here instead of hunting through export functions
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
  executiveSummary: "executiveSummary",
  trendNote: "trendNote",
  trendNarrative: "trendNarrative",
  potIncidentsNarrative: "potIncidentsNarrative",
  riskNarrative: "riskNarrative",
  inventoryNote: "inventoryNote",
  kpEnhancement: "kpEnhancement",
  kpDecommissioning: "kpDecommissioning",
  kpRuleImplementation: "kpRuleImplementation"
};

const STATIC_REVIEWED_BY = "Kishore Kumar";
const STATIC_APPROVED_BY = "Diptesh Saha";

const templatePresets = {
  blupine: {
    customerName: "Blupine Energy",
    preparedBy: "Maveera"
  },
  nocil: {
    customerName: "Nocil Limited",
    preparedBy: "Maveera"
  },
  apraava: {
    customerName: "Apraava Energy",
    preparedBy: "Maveera"
  }
};

const templateCopy = {
  blupine: {
    summaryTitle: "Total Potential Incident",
    summarySubtitle: "Potential incidents identified during",
    trendTitle: "Potential Incident Tickets Trend"
  },
  nocil: {
    summaryTitle: "Total Potential Alerts",
    summarySubtitle: "Potential alerts identified during",
    trendTitle: "Potential Alert Tickets Trend"
  },
  apraava: {
    summaryTitle: "Potential Incidents (severity summary)",
    summarySubtitle: "Potential incidents and alert activity during",
    trendTitle: "Potential Incident Tickets Trend"
  },
  custom: {
    summaryTitle: "Total Potential Incident",
    summarySubtitle: "Potential incidents identified during",
    trendTitle: "Potential Incident Tickets Trend"
  }
};

let alertChart;
let tpfpChart;
let trendChart;
let potIncidentsChart;
let epsTrendChart;
let topEpsChart;
let slaPctChart;
let apraavaQuadCharts = [];
let snsLogoDataUrl = "";
let clientLogoDataUrl = "";
let puzzleLogoDataUrl = "";

const pptLayoutConfig = {
  LAYOUT_16X9: {
    pptxLayout: "LAYOUT_WIDE",
    reportClass: "ppt-wide"
  },
  LAYOUT_4X3: {
    pptxLayout: "LAYOUT_4X3",
    reportClass: "ppt-standard"
  },
  A4: {
    pptxLayout: "LAYOUT_4X3",
    reportClass: "ppt-a4"
  }
};

function getValue(id) {
  const el = document.getElementById(id);
  return el ? el.value.trim() : "";
}

// Escapes user-supplied text before inserting into innerHTML
function escapeHtml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Converts **text** to <strong>text</strong> and handles [MONTH] placeholder for HTML display.
 */
function formatRichTextHTML(text) {
  if (!text) return "";
  let formatted = escapeHtml(text);
  // Support **bold**
  formatted = formatted.replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>');
  // Support [MONTH]
  formatted = formatted.replace(/\[MONTH\]/g, `<span class="fw-bold">${getReportMonth()}</span>`);
  return formatted;
}

/**
 * Parses **text** into an array of PptxGenJS text objects.
 */
function parseRichTextPptx(text, baseOptions = {}) {
  if (!text) return [];
  const parts = [];
  const monthReplaced = text.replace(/\[MONTH\]/g, getReportMonth());
  
  // Split by ** delimiters
  const segments = monthReplaced.split(/\*\*/);
  
  segments.forEach((segment, idx) => {
    if (segment === "") return;
    const isBold = idx % 2 !== 0; // Every second segment is bold
    parts.push({
      text: segment,
      options: { ...baseOptions, bold: isBold || baseOptions.bold }
    });
  });
  
  return parts;
}

// Parses a DD-MM-YYYY date string and returns a locale month string, or null
function parseDateMatch(str) {
  const match = str.match(/(\d{2})-(\d{2})-(\d{4})/);
  if (!match) return null;
  const monthIdx = Number(match[2]) - 1;
  const year = Number(match[3]);
  if (monthIdx < 0 || monthIdx > 11) return null;
  return new Date(year, monthIdx, 1).toLocaleString("en-US", { month: "long", year: "numeric" });
}

// Removes characters unsafe in filenames
function sanitizeFilename(str) {
  return str.replace(/[^a-zA-Z0-9_\-]/g, "_").replace(/_+/g, "_");
}

function getTemplateKey() {
  const sel = document.getElementById("templatePreset");
  return sel ? sel.value : "custom";
}

function getReportMonth() {
  return parseDateMatch(getValue("dateRange"))
    || parseDateMatch(getValue("submittedOn"))
    || new Date().toLocaleString("en-US", { month: "long", year: "numeric" });
}

function getPptLayoutKey() {
  const sel = document.getElementById("pptLayout");
  return sel ? sel.value : "LAYOUT_16X9";
}


function parseCsvRows(csvText, columns) {
  return csvText
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line, idx) => {
      const parts = line.split(",").map((cell) => cell.trim());
      const row = { sno: idx + 1 };
      columns.forEach((col, colIdx) => {
        row[col] = parts[colIdx] || "";
      });
      return row;
    });
}


function renderTableRows(targetId, rows, renderRow) {
  const tableBody = document.getElementById(targetId);
  if (!tableBody) return;
  tableBody.innerHTML = "";
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = renderRow(row);
    tableBody.appendChild(tr);
  });
}

function updateLogos() {
  const snsEl = document.getElementById("snsLogoPreview");
  const clientEl = document.getElementById("clientLogoPreview");
  if (!snsEl || !clientEl) return;
  snsEl.style.display = snsLogoDataUrl ? "block" : "none";
  clientEl.style.display = clientLogoDataUrl ? "block" : "none";
  snsEl.src = snsLogoDataUrl;
  clientEl.src = clientLogoDataUrl;
  
  document.querySelectorAll(".nav-sns-logo").forEach(img => {
      img.src = snsLogoDataUrl;
      img.style.display = snsLogoDataUrl ? "block" : "none";
  });
  const lastSnsEl = document.getElementById("lastSnsLogoPreview");
  if (lastSnsEl) {
    lastSnsEl.src = snsLogoDataUrl;
    lastSnsEl.style.display = snsLogoDataUrl ? "block" : "none";
  }
  const puzzleGraphicEl = document.getElementById("puzzleGraphicPreview");
  const puzzleLayoutEl = document.getElementById("puzzlePlaceholderLayout");
  if (puzzleGraphicEl && puzzleLayoutEl) {
    if (puzzleLogoDataUrl) {
       puzzleGraphicEl.style.display = "block";
       puzzleGraphicEl.src = puzzleLogoDataUrl;
       puzzleLayoutEl.style.display = "none";
    } else {
       puzzleGraphicEl.style.display = "none";
       puzzleLayoutEl.style.display = "";
    }
  }
}

function parseChartPairs(csvText) {
  return csvText
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => {
      const [label, value] = line.split(",");
      return { label: (label || "").trim(), value: Number((value || "0").trim()) || 0 };
    });
}

function applyPremiumLayout() {
  document.querySelectorAll(".page").forEach(page => {
    // Skip cover page, rev history, or engagement if already handled
    if (page.classList.contains("cover-hero") || page.querySelector(".engagement-container") || page.classList.contains("revision-page")) {
      return;
    }
    
    // Upgrade to page container with footer padding
    if (!page.classList.contains("page-with-footer")) {
      page.classList.add("page-with-footer");
    }

    // Inject Logo Header
    if (!page.querySelector(".page-header-right")) {
      const header = document.createElement("div");
      header.className = "page-header-right";
      header.innerHTML = '<img class="nav-sns-logo" alt="SNS Logo" />';
      page.insertBefore(header, page.firstChild);
    }

    // Inject Footer Bar
    if (!page.querySelector(".page-footer-bar")) {
      const footer = document.createElement("div");
      footer.className = "page-footer-bar";
      page.appendChild(footer);
    }

    // Standardize H2 Titles
    const h2 = page.querySelector("h2");
    if (h2 && !h2.classList.contains("revision-title")) {
      h2.classList.add("revision-title");
      h2.style.cssText = "margin-left: 0; margin-bottom: 20px; font-family: Calibri, 'Segoe UI', Arial, sans-serif; font-size: 38px;";
    }
  });

  // Re-run updateLogos so the newly injected images get the logo URL
  updateLogos();
}

function applyDynamicTitles() {
  const key = getTemplateKey();
  const copy = templateCopy[key] || templateCopy.custom;
  document.querySelectorAll('[data-dynamic="summaryTitle"]').forEach((el) => {
    el.textContent = copy.summaryTitle;
  });
  document.querySelectorAll('[data-dynamic="summarySubtitle"]').forEach((el) => {
    el.textContent = copy.summarySubtitle;
  });
  document.querySelectorAll('[data-dynamic="trendTitle"]').forEach((el) => {
    el.textContent = copy.trendTitle;
  });
}

function applyTemplateVisibility() {
  const key = getTemplateKey();
  const root = document.getElementById("reportRoot");
  if (root) {
    const layoutClass = (pptLayoutConfig[getPptLayoutKey()] || pptLayoutConfig.LAYOUT_16X9).reportClass;
    root.className = `report template-${key} ${layoutClass}`;
  }

  document.querySelectorAll("[data-show]").forEach((el) => {
    const allowed = el.getAttribute("data-show").split(",").map((s) => s.trim());
    if (key === "custom") {
      el.style.display = "";
      return;
    }
    el.style.display = allowed.includes(key) ? "" : "none";
  });
}

function renderCharts() {
  const trendRows = parseCsvRows(getValue("trendCsv"), ["date", "high", "medium"]);
  const trendCtx = document.getElementById("trendChart");
  if (trendCtx) {
    if (trendChart) trendChart.destroy();
    trendChart = new Chart(trendCtx, {
      type: "bar",
      data: {
        labels: trendRows.map(r => r.date),
        datasets: [
          { label: "High", data: trendRows.map(r => parseFloat(r.high) || 0), backgroundColor: "#ff0000" },
          { label: "Medium", data: trendRows.map(r => parseFloat(r.medium) || 0), backgroundColor: "#ffff00" }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          title: {
            display: true,
            text: 'Incident Summary – ' + getReportMonth(),
            font: { size: 16, weight: 'bold' }
          },
          legend: { position: 'bottom', labels: { boxWidth: 12, usePointStyle: true } }
        },
        scales: {
          x: { stacked: true, grid: { display: false } },
          y: { stacked: true, beginAtZero: true }
        }
      }
    });
  }

  const potIncRows = parseCsvRows(getValue("potIncidentsCsv"), ["rule", "high", "medium"]);
  const potIncCtx = document.getElementById("potIncidentsChart");
  if (potIncCtx) {
    if (potIncidentsChart) potIncidentsChart.destroy();
    
    const potIncLabels = potIncRows.map(r => r.rule);
    potIncidentsChart = new Chart(potIncCtx, {
      type: "bar",
      data: {
        labels: potIncLabels,
        datasets: [
          { label: "High", data: potIncRows.map(r => parseFloat(r.high) || 0), backgroundColor: "#ff0000" },
          { label: "Medium", data: potIncRows.map(r => parseFloat(r.medium) || 0), backgroundColor: "#ffff00" }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          title: {
            display: true,
            text: 'Incident Summary – ' + getReportMonth(),
            font: { size: 16, weight: 'bold' }
          },
          legend: { position: 'right', labels: { boxWidth: 12, usePointStyle: true } }
        },
        scales: {
          x: { stacked: true, grid: { display: false }, ticks: { maxRotation: 45, minRotation: 45, font: { size: 9 } } },
          y: { stacked: true, beginAtZero: true }
        }
      }
    });
  }

  const epsJan = parseFloat(getValue("epsJan")) || 0;
  const epsFeb = parseFloat(getValue("epsFeb")) || 0;
  const epsMar = parseFloat(getValue("epsMar")) || 0;
  
  const epsJanTable = document.getElementById("epsJanTable");
  const epsFebTable = document.getElementById("epsFebTable");
  const epsMarTable = document.getElementById("epsMarTable");
  if (epsJanTable) epsJanTable.textContent = isNaN(parseFloat(getValue("epsJan"))) ? "" : epsJan;
  if (epsFebTable) epsFebTable.textContent = isNaN(parseFloat(getValue("epsFeb"))) ? "" : epsFeb;
  if (epsMarTable) epsMarTable.textContent = isNaN(parseFloat(getValue("epsMar"))) ? "" : epsMar;

  const epsCtx = document.getElementById("epsTrendChart");
  if (epsCtx) {
    if (epsTrendChart) epsTrendChart.destroy();
    epsTrendChart = new Chart(epsCtx, {
      type: "bar",
      data: {
        labels: ["Jan", "Feb", "Mar"],
        datasets: [{
          label: "AVG EPS",
          data: [epsJan, epsFeb, epsMar],
          backgroundColor: ["#4bc0c0", "#ffcd56", "#00b050"]
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false }
        },
        scales: {
          y: { beginAtZero: true }
        }
      }
    });
  }

  const epsTableRows = parseCsvRows(getValue("epsTableCsv"), ["reportingDevice", "eventName", "jan", "feb", "mar"]);
  renderTableRows(
    "epsTrendTableBody",
    epsTableRows,
    (row) => `
      <td style="border: 1px solid #000; padding: 4px;">${row.sno}</td>
      <td style="border: 1px solid #000; padding: 4px; text-align: left;">${row.reportingDevice}</td>
      <td style="border: 1px solid #000; padding: 4px; text-align: left;">${row.eventName}</td>
      <td style="border: 1px solid #000; padding: 4px;">${row.jan}</td>
      <td style="border: 1px solid #000; padding: 4px;">${row.feb}</td>
      <td style="border: 1px solid #000; padding: 4px;">${row.mar}</td>
    `
  );



  const aqKeys = [
    { csv: "overallAlertsCsv", chartId: "aqChart1", tableId: "aqTable1", narr: "overallAlertsNarrative", narrId: "aqNarrative1" },
    { csv: "truePositiveCsv", chartId: "aqChart2", tableId: "aqTable2", narr: "truePositiveNarrative", narrId: "aqNarrative2" },
    { csv: "falsePositiveCsv", chartId: "aqChart3", tableId: "aqTable3", narr: "falsePositiveNarrative", narrId: "aqNarrative3" },
    { csv: "apraavaIncidentsCsv", chartId: "aqChart4", tableId: "aqTable4", narr: "apraavaIncidentsNarrative", narrId: "aqNarrative4" }
  ];

  // Clear old charts
  apraavaQuadCharts.forEach(c => c && c.destroy());
  apraavaQuadCharts = [];

  aqKeys.forEach((config) => {
    // 1. Setup Narrative
    const narrEl = document.getElementById(config.narrId);
    if (narrEl) narrEl.textContent = getValue(config.narr);

    // 2. Parse CSV Data (High, Medium, Low)
    const csvLines = getValue(config.csv).split("\n").map(l => l.trim()).filter(Boolean);
    let dsHigh = [0,0,0], dsMed = [0,0,0], dsLow = [0,0,0];
    
    csvLines.forEach(line => {
      const parts = line.split(",").map(p => p.trim());
      const sev = (parts[0] || "").toLowerCase();
      const vals = [Number(parts[1])||0, Number(parts[2])||0, Number(parts[3])||0];
      if (sev === "high") dsHigh = vals;
      if (sev === "medium") dsMed = vals;
      if (sev === "low") dsLow = vals;
    });

    // 3. Render HTML Table Logic
    const tb = document.getElementById(config.tableId);
    if (tb) {
      tb.innerHTML = `
        <thead>
          <tr><th></th><th>Jan</th><th>Feb</th><th>Mar</th></tr>
        </thead>
        <tbody>
          <tr><td><span class="aq-sq high"></span> High</td><td>${dsHigh[0] || 0}</td><td>${dsHigh[1] || 0}</td><td>${dsHigh[2] || 0}</td></tr>
          <tr><td><span class="aq-sq med"></span> Medium</td><td>${dsMed[0] || 0}</td><td>${dsMed[1] || 0}</td><td>${dsMed[2] || 0}</td></tr>
          <tr><td><span class="aq-sq low"></span> Low</td><td>${dsLow[0] || 0}</td><td>${dsLow[1] || 0}</td><td>${dsLow[2] || 0}</td></tr>
        </tbody>
      `;
    }

    // 4. Render Chart.js
    const cv = document.getElementById(config.chartId);
    if (cv) {
      apraavaQuadCharts.push(new Chart(cv, {
        type: "bar",
        data: {
          labels: ["Jan", "Feb", "Mar"],
          datasets: [
            { label: "High", data: dsHigh, backgroundColor: "#ff0000" },
            { label: "Medium", data: dsMed, backgroundColor: "#ffff00" },
            { label: "Low", data: dsLow, backgroundColor: "#00b050" }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false } // Hidden since the HTML table acts as legend
          },
          scales: {
            x: { display: false }, // Hidden since the HTML table acts as X-axis labels
            y: { 
               beginAtZero: true,
               ticks: { font: { size: 10 } }
            }
          }
        }
      }));
    }
  });
}

function applyData() {
  applyTemplateVisibility();
  applyDynamicTitles();

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

  // alertSummaryNarrative removed since we partitioned it 


  const execEl = document.getElementById("executiveSummaryPreview");
  if (execEl) {
    const raw = getValue("executiveSummary");
    execEl.innerHTML = raw.split("\n")
      .filter(p => p.trim())
      .map(p => `<p>${formatRichTextHTML(p)}</p>`)
      .join("");
  }



  const trendNarrEl = document.getElementById("trendNarrativePreview");
  if (trendNarrEl) trendNarrEl.innerHTML = formatRichTextHTML(getValue("trendNarrative"));

  const potIncNarrEl = document.getElementById("potIncidentsNarrativePreview");
  if (potIncNarrEl) potIncNarrEl.innerHTML = formatRichTextHTML(getValue("potIncidentsNarrative"));

  const riskNarrEl = document.getElementById("riskNarrativePreview");
  if (riskNarrEl) riskNarrEl.innerHTML = formatRichTextHTML(getValue("riskNarrative"));

  const kpEnh = document.getElementById("kpEnhancementPreview");
  if (kpEnh) kpEnh.innerHTML = formatRichTextHTML(getValue("kpEnhancement"));

  const kpDecom = document.getElementById("kpDecommissioningPreview");
  if (kpDecom) kpDecom.innerHTML = formatRichTextHTML(getValue("kpDecommissioning"));

  const kpRule = document.getElementById("kpRuleImplementationPreview");
  if (kpRule) kpRule.innerHTML = formatRichTextHTML(getValue("kpRuleImplementation"));

  const deviceRows = parseCsvRows(getValue("deviceCsv"), ["device", "eps"]);
  renderTableRows(
    "deviceTableBody",
    deviceRows,
    (row) => `<tr><td style="border: 1px solid #00b0ba; padding: 6px;">${row.sno}</td><td style="border: 1px solid #00b0ba; padding: 6px; text-align: left;">${row.device}</td><td style="border: 1px solid #00b0ba; padding: 6px;">${row.eps}</td></tr>`
  );

  const topEpsCtx = document.getElementById("topEpsChart");
  if (topEpsCtx) {
    if (topEpsChart) topEpsChart.destroy();
    const colors = ["#4bc0c0", "#ffcd56", "#ff6384", "#ff9f40", "#9966ff", "#c9cbcf", "#00b050", "#be2f2f", "#1e4f9a", "#8e5ea2"];
    topEpsChart = new Chart(topEpsCtx, {
      type: "bar",
      data: {
        labels: deviceRows.map(r => r.device),
        datasets: [{
          data: deviceRows.map(r => parseFloat(r.eps) || 0),
          backgroundColor: colors.slice(0, deviceRows.length)
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          datalabels: {
            display: true,
            color: '#000',
            anchor: 'end',
            align: 'top',
            formatter: Math.round,
            font: { weight: 'bold' }
          }
        },
        scales: {
          x: { display: true, ticks: { maxRotation: 45, minRotation: 45, font: { size: 9 } } },
          y: { 
            beginAtZero: true,
            title: { display: true, text: 'EPS' }
          }
        }
      }
    });
  }

  const riskRows = parseCsvRows(getValue("riskCsv"), [
    "attackType",
    "riskScenario",
    "ciaTriad",
    "businessImpact",
    "riskRating"
  ]);
  renderTableRows(
    "riskTableBody",
    riskRows,
    (row) =>
      `<tr>
         <td style="border: 1px solid #7f7f7f; padding: 4px; text-align: center;">${row.sno}</td>
         <td style="border: 1px solid #7f7f7f; padding: 4px; text-align: left;">${row.attackType}</td>
         <td style="border: 1px solid #7f7f7f; padding: 4px; text-align: left;">${row.riskScenario}</td>
         <td style="border: 1px solid #7f7f7f; padding: 4px; text-align: center;">${row.ciaTriad}</td>
         <td style="border: 1px solid #7f7f7f; padding: 4px; text-align: left; font-size: 9px;">${row.businessImpact}</td>
         <td style="border: 1px solid #7f7f7f; padding: 4px; text-align: center; font-weight: bold;">${row.riskRating}</td>
       </tr>`
  );

  const slaTb = document.getElementById("slaTableBody");
  if (slaTb) {
    slaTb.innerHTML = `
      <tr>
        <td style="border: 1px solid #000; padding: 6px; font-weight: bold;">S-1 (High)</td>
        <td style="border: 1px solid #000; padding: 6px;">≤15 Min</td>
        <td style="border: 1px solid #000; padding: 6px;">4 Hours</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaHighJanCount")}</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaHighFebCount")}</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaHighMarCount")}</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000; padding: 6px; font-weight: bold;">S-2 (Medium)</td>
        <td style="border: 1px solid #000; padding: 6px;">≤15 Min</td>
        <td style="border: 1px solid #000; padding: 6px;">8 Hours</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaMedJanCount")}</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaMedFebCount")}</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaMedMarCount")}</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000; padding: 6px; font-weight: bold;">S-3 (Low)</td>
        <td style="border: 1px solid #000; padding: 6px;">≤30 Min</td>
        <td style="border: 1px solid #000; padding: 6px;">24 Hours</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaLowJanCount")}</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaLowFebCount")}</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaLowMarCount")}</td>
      </tr>
    `;
  }

  const slaPctCtx = document.getElementById("slaPctChart");
  if (slaPctCtx) {
    if (slaPctChart) slaPctChart.destroy();
    slaPctChart = new Chart(slaPctCtx, {
      type: "bar",
      data: {
        labels: ["Jan", "Feb", "Mar"],
        datasets: [
          {
            label: "S-1 (High) Within SLA (4 Hrs)",
            data: [getValue("slaHighJanPct"), getValue("slaHighFebPct"), getValue("slaHighMarPct")],
            backgroundColor: "#be2f2f"
          },
          {
            label: "S-2 (Medium) Within SLA (8 Hrs)",
            data: [getValue("slaMedJanPct"), getValue("slaMedFebPct"), getValue("slaMedMarPct")],
            backgroundColor: "#ffff00"
          },
          {
            label: "S-3 (Low) Within SLA (24 Hrs)",
            data: [getValue("slaLowJanPct"), getValue("slaLowFebPct"), getValue("slaLowMarPct")],
            backgroundColor: "#00b050"
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          y: { 
            beginAtZero: true, 
            max: 120,
            ticks: {
              callback: function(value) {
                return value + "%";
              }
            }
          }
        },
        plugins: {
          legend: { position: "bottom", labels: { font: { size: 10 } }, align: "center" }
        }
      }
    });
  }

  const inventoryRows = parseCsvRows(getValue("inventoryCsv"), ["deviceName", "count"]);
  let totalCount = 0;
  
  const inventoryHtml = inventoryRows.map((row, idx) => {
    const cnt = parseInt(row.count) || 0;
    totalCount += cnt;
    const bgColor = idx % 2 === 0 ? "#ffffff" : "#f1f9f9";
    return `<tr style="background: ${bgColor};">
      <td style="padding: 8px; border: 1px solid #4cc0c5;">${row.sno}</td>
      <td style="padding: 8px; border: 1px solid #4cc0c5; text-align: left;">${row.deviceName}</td>
      <td style="padding: 8px; border: 1px solid #4cc0c5;">${cnt}</td>
    </tr>`;
  }).join("");

  const invTb = document.getElementById("inventoryTableBody");
  if (invTb) invTb.innerHTML = inventoryHtml;

  const invTf = document.getElementById("inventoryTableFooter");
  if (invTf) {
    invTf.innerHTML = `
      <tr style="background: #ffffff; font-weight: bold;">
        <td style="padding: 10px; border: 1px solid #4cc0c5;" colspan="2">Total</td>
        <td style="padding: 10px; border: 1px solid #4cc0c5;">${totalCount}</td>
      </tr>
    `;
  }

  const invNotePreview = document.getElementById("inventoryNotePreview");
  if (invNotePreview) {
    invNotePreview.innerText = getValue("inventoryNote");
  }



  updateLogos();
  renderCharts();
  applyPremiumLayout();
}

function exportPdf() {
  window.print();
}

function canvasToPptxData(canvas) {
  if (!canvas) return null;
  return canvas.toDataURL("image/png");
}

function getPptxCtor() {
  return window.PptxGenJS || window.pptxgen;
}

function addPptxTitleBar(slide, title) {
  slide.addText(title, {
    x: 0,
    y: 0,
    w: "100%",
    h: 0.45,
    fontSize: 14,
    bold: true,
    color: "FFFFFF",
    fill: { color: CONFIG.PPTX_NAVY },
    align: "left",
    valign: "middle",
    margin: [0.1, 0.35, 0.1, 0.35]
  });
}

function addPptxFooter(slide) {
  slide.addText(CONFIG.COMPANY_NAME, {
    x: 0.35,
    y: 5,
    w: 9,
    h: 0.35,
    fontSize: 9,
    color: "666666"
  });
}

async function exportPptx() {
  const Ctor = getPptxCtor();
  if (!Ctor) {
    alert("PptxGenJS failed to load. Check network and refresh.");
    return;
  }

  applyData();
  // Ensure charts have enough time to render before capturing canvases
  await new Promise((r) => setTimeout(r, 600));

  const pptx = new Ctor();
  const selectedLayout = pptLayoutConfig[getPptLayoutKey()] || pptLayoutConfig.LAYOUT_16X9;
  pptx.layout = selectedLayout.pptxLayout;
  pptx.author = "SNS SOC";
  pptx.title = `${getValue("customerName")} Monthly SOC Report`;

  if (getPptLayoutKey() === "A4") {
    pptx.defineLayout({ name: "CUSTOM_A4_PORTRAIT", width: 8.27, height: 11.69 });
    pptx.layout = "CUSTOM_A4_PORTRAIT";
  }

  const key = getTemplateKey();
  const copy = templateCopy[key] || templateCopy.custom;

  let slide = pptx.addSlide();
  addPptxTitleBar(slide, "Monthly SOC Report");
  slide.addText("MANAGED INCIDENT RESPONSE & REMEDIATION SERVICE", {
    x: 0.5,
    y: 0.65,
    w: 9,
    h: 0.35,
    fontSize: 12,
    color: CONFIG.PPTX_BLUE,
    bold: true
  });
  slide.addText(`${getReportMonth()}`, { x: 0.5, y: 1.1, w: 9, h: 0.5, fontSize: 24, bold: true, color: CONFIG.PPTX_NAVY });
  slide.addText(`for ${getValue("customerName")}`, { x: 0.5, y: 1.65, w: 9, h: 0.4, fontSize: 16 });
  slide.addText(
    [
      `Version 1.0`,
      `Dated ${getValue("dateRange")}`,
      `Prepared By ${getValue("preparedBy")}`,
      `Reviewed By ${STATIC_REVIEWED_BY}`,
      `Approved By ${STATIC_APPROVED_BY}`,
      `Submitted On ${getValue("submittedOn")}`
    ].join("\n"),
    { x: 0.5, y: 2.3, w: 5.5, h: 2, fontSize: 11 }
  );
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: 6.8, y: 0.65, w: 1.2, h: 0.55 });
  }
  if (clientLogoDataUrl) {
    slide.addImage({ data: clientLogoDataUrl, x: 8.1, y: 0.65, w: 1.2, h: 0.55 });
  }
  addPptxFooter(slide);

  slide = pptx.addSlide();
  addPptxTitleBar(slide, "Document Revision History");
  slide.addText(
    getValue("documentRevisionHistory")
      .split("\n")
      .filter(Boolean)
      .map((l) => `• ${l}`)
      .join("\n"),
    { x: 0.5, y: 0.65, w: 9, h: 3, fontSize: 11 }
  );
  addPptxFooter(slide);

  slide = pptx.addSlide();
  addPptxTitleBar(slide, "The Engagement");
  
  const introText = `${getValue("customerName")} has engaged with SNS to monitor and review the entity's security.\n\n`;
  const execBody = getValue("executiveSummary");
  slide.addText([...parseRichTextPptx(introText), ...parseRichTextPptx(execBody)], { x: 0.5, y: 0.65, w: 9, h: 3.5, fontSize: 11 });
  addPptxFooter(slide);
  addPptxFooter(slide);

  if (key === "apraava") {
    slide = pptx.addSlide();
    addPptxTitleBar(slide, "Potential Incidents and Alert Summary");
    
    const quadIds = ["aqChart1", "aqChart2", "aqChart3", "aqChart4"];
    // We can extract images from the Chart.js canvases and position them in 4 quadrants
    quadIds.forEach((id, idx) => {
      const cImg = canvasToPptxData(document.getElementById(id));
      if (cImg) {
        const isRight = idx % 2 !== 0;
        const isBot = Math.floor(idx / 2) > 0;
        slide.addImage({ data: cImg, x: isRight ? 5.0 : 0.5, y: isBot ? 3.0 : 0.8, w: 4.4, h: 2.0 });
      }
    });
    addPptxFooter(slide);
  }



  slide = pptx.addSlide();
  addPptxTitleBar(slide, copy.trendTitle);
  const trendImg = canvasToPptxData(document.getElementById("trendChart"));
  if (trendImg) {
    slide.addImage({ data: trendImg, x: 0.8, y: 0.65, w: 8.4, h: 3.5 });
  }
  slide.addText(parseRichTextPptx(getValue("trendNote")), { x: 0.5, y: 4.3, w: 9, h: 0.3, fontSize: 13, bold: true });
  slide.addText(parseRichTextPptx(getValue("trendNarrative")), { x: 0.5, y: 4.6, w: 9, h: 0.7, fontSize: 11 });
  addPptxFooter(slide);

  slide = pptx.addSlide();
  addPptxTitleBar(slide, "Potential Incidents");
  const potIncImg = canvasToPptxData(document.getElementById("potIncidentsChart"));
  if (potIncImg) {
    slide.addImage({ data: potIncImg, x: 0.8, y: 0.65, w: 8.4, h: 3.5 });
  }
  slide.addText(parseRichTextPptx(getValue("potIncidentsNarrative")), { x: 0.5, y: 4.4, w: 9, h: 0.8, fontSize: 11 });
  addPptxFooter(slide);

  // New EPS Trend Slide
  slide = pptx.addSlide();
  addPptxTitleBar(slide, "EPS Trend Plot");
  const epsImg = canvasToPptxData(document.getElementById("epsTrendChart"));
  if (epsImg) {
    slide.addImage({ data: epsImg, x: 0.5, y: 0.8, w: 3.5, h: 4.0 });
  }
  const epsTblRows = parseCsvRows(getValue("epsTableCsv"), ["reportingDevice", "eventName", "jan", "feb", "mar"]);
  if (epsTblRows.length > 0) {
    const tableData = [
      [
        { text: "S.No", options: { bold: true, fill: { color: "F3F6FB" } } },
        { text: "Reporting Device", options: { bold: true, fill: { color: "F3F6FB" } } },
        { text: "Event Name", options: { bold: true, fill: { color: "F3F6FB" } } },
        { text: "Jan", options: { bold: true, fill: { color: "F3F6FB" } } },
        { text: "Feb", options: { bold: true, fill: { color: "F3F6FB" } } },
        { text: "Mar", options: { bold: true, fill: { color: "F3F6FB" } } }
      ],
      ...epsTblRows.map(r => [String(r.sno), r.reportingDevice, r.eventName, r.jan, r.feb, r.mar])
    ];
    // Place table on the right side of the slide
    slide.addTable(tableData, { x: 4.2, y: 0.8, w: 5.4, fontSize: 8, colW: [0.5, 1.2, 1.3, 0.8, 0.8, 0.8] });
  }
  addPptxFooter(slide);

  slide = pptx.addSlide();
  addPptxTitleBar(slide, "Top 10 Devices Contributing To Highest EPS");
  
  const devImg = canvasToPptxData(document.getElementById("topEpsChart"));
  if (devImg) {
    slide.addImage({ data: devImg, x: 0.5, y: 0.8, w: 4.5, h: 4.0 });
  }

  const devices = parseCsvRows(getValue("deviceCsv"), ["device", "eps"]);
  if (devices.length > 0) {
    const devRows = [
      [
        { text: "S.No", options: { bold: true, fill: { color: "8EEDE4" } } },
        { text: "Reporting Device", options: { bold: true, fill: { color: "8EEDE4" } } },
        { text: "AVG EPS", options: { bold: true, fill: { color: "8EEDE4" } } }
      ],
      ...devices.map((d) => [String(d.sno), d.device, d.eps])
    ];
    slide.addTable(devRows, { x: 5.2, y: 0.8, w: 4.3, fontSize: 10, colW: [0.6, 2.5, 1.2] });
  }
  addPptxFooter(slide);

  const risks = parseCsvRows(getValue("riskCsv"), [
    "attackType",
    "riskScenario",
    "ciaTriad",
    "businessImpact",
    "riskRating"
  ]);
  if (risks.length === 0) {
    slide = pptx.addSlide();
    addPptxTitleBar(slide, "Potential Incidents – Risks Mitigated");
    slide.addText("No risk rows in CSV. Add rows under Potential Incidents CSV.", {
      x: 0.5,
      y: 0.65,
      w: 9,
      h: 0.5,
      fontSize: 11
    });
    addPptxFooter(slide);
  }

  const risksRowsPerSlide = 4;
  for (let i = 0; i < risks.length; i += risksRowsPerSlide) {
    const chunk = risks.slice(i, i + risksRowsPerSlide);
    slide = pptx.addSlide();
    addPptxTitleBar(slide, i === 0 ? "Potential Incidents – Risks Mitigated" : "Contn.,");
    
    if (i === 0) {
      slide.addText(parseRichTextPptx(getValue("riskNarrative")), { x: 0.35, y: 0.65, w: 9.3, h: 0.8, fontSize: 10 });
    }

    const rows = [
      [
        { text: "S.No", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } },
        { text: "Attack Type", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } },
        { text: "Risk Scenario", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } },
        { text: "Type of Risk(s)\nCIA Triad", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } },
        { text: "Potential Business Impact(s)", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } },
        { text: "Risk Rating", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } }
      ],
      ...chunk.map((r) => [
        { text: String(r.sno), options: { align: "center", fontFace: "Calibri" } },
        { text: r.attackType, options: { fontFace: "Calibri" } },
        { text: r.riskScenario, options: { fontFace: "Calibri" } },
        { text: r.ciaTriad, options: { align: "center", fontFace: "Calibri" } },
        { text: r.businessImpact, options: { fontFace: "Calibri" } },  // No truncation so users can edit all text
        { text: r.riskRating, options: { align: "center", fontFace: "Calibri" } }
      ])
    ];
    
    const tableY = i === 0 ? 1.5 : 0.6;
    slide.addTable(rows, { 
       x: 0.35, y: tableY, w: 9.3, 
       fontSize: 9, 
       border: { type: 'solid', color: 'FFFFFF', pt: 1 }, 
       colW: [0.4, 1.4, 2.0, 1.2, 3.5, 0.8] 
    });
    addPptxFooter(slide);
  }



  slide = pptx.addSlide();
  addPptxTitleBar(slide, "Response Time SLA");
  const slaRows = [
    [
      { text: "Severity Level", options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } },
      { text: "SLA Response", options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } },
      { text: "SLA Resolution", options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } },
      { text: "Jan", options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } },
      { text: "Feb", options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } },
      { text: "Mar", options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } }
    ],
    [
      "S-1 (High)", "≤15 Min", "4 Hours",
      getValue("slaHighJanCount"), getValue("slaHighFebCount"), getValue("slaHighMarCount")
    ],
    [
      "S-2 (Medium)", "≤15 Min", "8 Hours",
      getValue("slaMedJanCount"), getValue("slaMedFebCount"), getValue("slaMedMarCount")
    ],
    [
      "S-3 (Low)", "≤30 Min", "24 Hours",
      getValue("slaLowJanCount"), getValue("slaLowFebCount"), getValue("slaLowMarCount")
    ]
  ];
  slide.addTable(slaRows, { x: 0.5, y: 0.8, w: 9.0, fontSize: 10, border: { pt: 1, color: "000000" } });

  slide.addText("Remediation Time SLA in %", { x: 0.5, y: 2.5, w: 9, h: 0.4, fontSize: 18, bold: true });
  const slaPctImg = canvasToPptxData(document.getElementById("slaPctChart"));
  if (slaPctImg) {
    slide.addImage({ data: slaPctImg, x: 1.0, y: 3.0, w: 8.0, h: 2.2 });
  }
  addPptxFooter(slide);

  slide = pptx.addSlide();
  addPptxTitleBar(slide, "Integrated Device Inventory");
  const invCsv = getValue("inventoryCsv");
  const inv = parseCsvRows(invCsv, ["deviceName", "count"]);
  let invTotal = 0;
  inv.forEach(r => { invTotal += (parseInt(r.count) || 0); });

  const invRows = [
    [
      { text: "S.NO", options: { bold: true, color: "FFFFFF", fill: { color: "4cc0c5" }, align: "center" } },
      { text: "DEVICE NAME", options: { bold: true, color: "FFFFFF", fill: { color: "4cc0c5" }, align: "center" } },
      { text: "COUNT", options: { bold: true, color: "FFFFFF", fill: { color: "4cc0c5" }, align: "center" } }
    ],
    ...inv.map((r, i) => [
      { text: String(r.sno), options: { align: "center" } },
      { text: r.deviceName, options: { align: "left" } },
      { text: String(r.count), options: { align: "center" } }
    ]),
    [
      { text: "Total", options: { bold: true, align: "center", colspan: 2 } },
      { text: String(invTotal), options: { bold: true, align: "center" } }
    ]
  ];
  slide.addTable(invRows, { 
    x: 0.5, y: 0.8, w: 9.0, 
    fontSize: 12, 
    border: { pt: 1, color: "4cc0c5" },
    valign: "middle"
  });

  const invNote = getValue("inventoryNote");
  if (invNote) {
    slide.addText(parseRichTextPptx(invNote), { 
      x: 0.5, y: 3.8, w: 9.0, h: 1.0, 
      fontSize: 11, 
      color: "000000",
      fontFace: "Calibri",
      valign: "top"
    });
  }
  addPptxFooter(slide);

  slide = pptx.addSlide();
  addPptxTitleBar(slide, "Key Points – Overall Summary");
  
  const yStart = 0.8;
  const sectionH = 1.3;
  
  // Section 1: Enhancement
  slide.addText("1. Report Enhancement:", { x: 0.5, y: yStart, w: 9, h: 0.3, fontSize: 13, bold: true, fontFace: "Calibri" });
  slide.addText(parseRichTextPptx(getValue("kpEnhancement")), { x: 0.5, y: yStart + 0.3, w: 9, h: 0.8, fontSize: 11, fontFace: "Calibri", valign: "top" });

  // Section 2: Decommissioning
  slide.addText("2. Device Decommissioning:", { x: 0.5, y: yStart + 1.2, w: 9, h: 0.3, fontSize: 13, bold: true, fontFace: "Calibri" });
  slide.addText(parseRichTextPptx(getValue("kpDecommissioning")), { x: 0.5, y: yStart + 1.5, w: 9, h: 1.2, fontSize: 11, fontFace: "Calibri", valign: "top" });

  // Section 3: Rule Implementation
  slide.addText("3. Rule Implementation:", { x: 0.5, y: yStart + 2.8, w: 9, h: 0.3, fontSize: 13, bold: true, fontFace: "Calibri" });
  slide.addText(parseRichTextPptx(getValue("kpRuleImplementation")), { x: 0.5, y: yStart + 3.1, w: 9, h: 1.2, fontSize: 11, fontFace: "Calibri", valign: "top" });

  addPptxFooter(slide);

  slide = pptx.addSlide();
  // Matching SS2: Left-aligned, large Branding, detailed contact info
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: 0.5, y: 0.5, w: 2.8 });
  }
  slide.addText("Your Trusted Security Advisor", { 
    x: 0.5, y: 1.8, w: 8.5, h: 1.0, 
    fontSize: 48, 
    bold: true, 
    color: "022a5b", 
    align: "left",
    fontFace: "Calibri"
  });
  
  const contactLines = [
    { text: "FOR FURTHER DETAILS:", options: { bold: true, color: "022a5b", fontSize: 13 } },
    { text: "Diptesh Saha", options: { bold: true, color: "1e4f9a", fontSize: 18 } },
    { text: "CISO & Practice Head - Cyber Security & Managed Security", options: { color: "1e4f9a", fontSize: 13 } },
    { text: "Contact No. 7338882888", options: { color: "1e4f9a", fontSize: 13 } },
    { text: "diptesh.s@snsin.com", options: { color: "1e4f9a", fontSize: 13 } }
  ];

  slide.addText(contactLines, { 
    x: 0.5, y: 2.8, w: 8.5, h: 2.5, 
    align: "left", 
    fontFace: "Calibri",
    valign: "top"
  });
  addPptxFooter(slide);

  const fname = `${sanitizeFilename(getValue("customerName"))}_${sanitizeFilename(getReportMonth())}_Monthly_SOC_Report.pptx`;
  await pptx.writeFile(fname);
}

async function exportDocx() {
  if (!window.docx || !window.saveAs) {
    alert("DOCX libraries failed to load.");
    return;
  }

  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType } = window.docx;

  const customer = getValue("customerName");
  const month = getReportMonth();
  const dateRange = getValue("dateRange");
  const preparedBy = getValue("preparedBy");
  const reviewedBy = STATIC_REVIEWED_BY;
  const approvedBy = STATIC_APPROVED_BY;
  const submittedOn = getValue("submittedOn");
  const highCount = getValue("highCount");
  const mediumCount = getValue("mediumCount");
  const points = getValue("keyPoints")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
  const devices = parseCsvRows(getValue("deviceCsv"), ["device", "eps"]);
  const risks = parseCsvRows(getValue("riskCsv"), [
    "attackType",
    "riskScenario",
    "ciaTriad",
    "businessImpact",
    "riskRating"
  ]);
  const slas = parseCsvRows(getValue("slaCsv"), ["severity", "response", "resolution", "count"]);
  const inventories = parseCsvRows(getValue("inventoryCsv"), ["deviceName", "count"]);
  const executiveSummary = getValue("executiveSummary");
  const incidentNote = getValue("incidentNote");
  const trendNote = getValue("trendNote");
  const revision = getValue("documentRevisionHistory");
  const alertNarrative = getValue("alertSummaryNarrative");

  const deviceRows = [
    new TableRow({
      children: ["S.No", "Device", "AVG EPS"].map(
        (col) =>
          new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: col, bold: true })] })]
          })
      )
    }),
    ...devices.map(
      (row) =>
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph(String(row.sno))] }),
            new TableCell({ children: [new Paragraph(row.device)] }),
            new TableCell({ children: [new Paragraph(row.eps)] })
          ]
        })
    )
  ];

  const key = getTemplateKey();
  const copy = templateCopy[key] || templateCopy.custom;

  const children = [
    new Paragraph({
      children: [new TextRun({ text: "MANAGED INCIDENT RESPONSE & REMEDIATION SERVICE", bold: true })]
    }),
    new Paragraph({
      children: [new TextRun({ text: `Monthly SOC Report | ${month}`, bold: true, size: 32 })]
    }),
    new Paragraph(`for ${customer}`),
    new Paragraph(""),
    new Paragraph(`Version 1.0`),
    new Paragraph(`Dated ${dateRange}`),
    new Paragraph(`Prepared By ${preparedBy}`),
    new Paragraph(`Reviewed By ${reviewedBy}`),
    new Paragraph(`Approved By ${approvedBy}`),
    new Paragraph(`Submitted On ${submittedOn}`),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "Document Revision History", bold: true })] }),
    new Paragraph(revision),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "The Engagement", bold: true })] }),
    new Paragraph(`${customer} has engaged with SNS to monitor and review the entity's security.`),
    new Paragraph(executiveSummary.replace(/\[MONTH\]/gi, month)),
    new Paragraph("")
  ];

  if (key === "apraava") {
    children.push(
      new Paragraph({ children: [new TextRun({ text: "Potential Incidents and Alert Summary", bold: true })] }),
      new Paragraph(alertNarrative),
      new Paragraph("")
    );
  }

  children.push(
    new Paragraph({ children: [new TextRun({ text: copy.summaryTitle, bold: true })] }),
    new Paragraph(`High: ${highCount}`),
    new Paragraph(`Medium: ${mediumCount}`),
    new Paragraph(`Note: ${incidentNote}`),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: copy.trendTitle, bold: true })] }),
    new Paragraph(trendNote),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "Potential Incidents - Risks Mitigated", bold: true })] }),
    ...risks.map(
      (risk) =>
        new Paragraph(
          `${risk.sno}. ${risk.attackType} | ${risk.riskScenario} | ${risk.ciaTriad} | ${risk.businessImpact} | ${risk.riskRating}`
        )
    ),
    new Paragraph("")
  );

  if (key === "blupine") {
    const hi = parseCsvRows(getValue("highestEpsCsv"), ["device", "eventType", "eventName", "matched"]);
    const hiRows = [
      new TableRow({
        children: ["S.No", "Reporting Device", "Event Type", "Event Name", "Matched Events"].map(
          (col) =>
            new TableCell({
              children: [new Paragraph({ children: [new TextRun({ text: col, bold: true })] })]
            })
        )
      }),
      ...hi.map(
        (row) =>
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph(String(row.sno))] }),
              new TableCell({ children: [new Paragraph(row.device)] }),
              new TableCell({ children: [new Paragraph(row.eventType)] }),
              new TableCell({ children: [new Paragraph(row.eventName)] }),
              new TableCell({ children: [new Paragraph(row.matched)] })
            ]
          })
      )
    ];
    children.push(
      new Paragraph({ children: [new TextRun({ text: "Highest EPS Consuming Events For Mentioned Firewalls", bold: true })] }),
      new Table({
        rows: hiRows,
        width: { size: 100, type: WidthType.PERCENTAGE }
      }),
      new Paragraph(""),
      new Paragraph({ children: [new TextRun({ text: "Overall Support Ticket Handled By SNS (Firewall Support)", bold: true })] }),
      new Paragraph(getValue("firewallTicketsText")),
      new Paragraph("")
    );
  }

  children.push(
    new Paragraph({ children: [new TextRun({ text: "Top EPS Devices", bold: true })] }),
    new Table({
      rows: deviceRows,
      width: { size: 100, type: WidthType.PERCENTAGE }
    }),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "Total Alerts / TP-FP Charts", bold: true })] }),
    new Paragraph("Charts appear in the on-screen report, PDF printout, and PPTX export."),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "Response & Remediation SLA", bold: true })] }),
    ...slas.map(
      (sla) => new Paragraph(`${sla.severity}: Response ${sla.response}, Resolution ${sla.resolution}, Count ${sla.count}`)
    ),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "Integrated Device Inventory", bold: true })] }),
    ...inventories.map((inv) => new Paragraph(`${inv.sno}. ${inv.deviceName}: ${inv.count}`)),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "Key Points - Overall Summary", bold: true })] }),
    ...points.map((point) => new Paragraph(`- ${point}`)),
    new Paragraph(""),
    new Paragraph("Your Trusted Security Advisor"),
    new Paragraph(CONFIG.CONTACT_FOOTER_TEXT),
    new Paragraph(CONFIG.COMPANY_NAME)
  );

  const doc = new Document({
    sections: [
      {
        children
      }
    ]
  });

  const blob = await Packer.toBlob(doc);
  saveAs(blob, `${sanitizeFilename(customer)}_${sanitizeFilename(month)}_Monthly_SOC_Report.docx`);
}

function setPresetData(presetKey) {
  const preset = templatePresets[presetKey];
  if (!preset) {
    applyData();
    return;
  }
  Object.entries(preset).forEach(([id, value]) => {
    const el = document.getElementById(id);
    if (el) {
      el.value = value;
    }
  });
  applyData();
}

function bindImageInput(inputId, setter) {
  const input = document.getElementById(inputId);
  if (!input) return;
  input.addEventListener("change", (event) => {
    const file = event.target.files && event.target.files[0];
    if (!file) {
      setter("");
      updateLogos();
      return;
    }
    const reader = new FileReader();
    reader.onload = (e) => {
      setter(String(e.target?.result || ""));
      updateLogos();
    };
    reader.readAsDataURL(file);
  });
}

function initTableResizing() {
  const tables = document.querySelectorAll("table");
  tables.forEach(table => {
    const headerRow = table.querySelector("thead tr");
    if (!headerRow) return;
    const cols = headerRow.querySelectorAll("th");
    // Only interior borders are resizable (length - 1 skips the right edge of the last column)
    for (let i = 0; i < cols.length - 1; i++) {
      if (cols[i].querySelector(".resizer")) continue;
      const resizer = document.createElement("div");
      resizer.className = "resizer";
      cols[i].appendChild(resizer);
      
      let x = 0, w = 0, nw = 0;
      const onMouseDown = (e) => {
        x = e.clientX;
        w = cols[i].offsetWidth;
        nw = cols[i+1].offsetWidth;
        document.addEventListener("mousemove", onMouseMove);
        document.addEventListener("mouseup", onMouseUp);
      };
      const onMouseMove = (e) => {
        const dx = e.clientX - x;
        cols[i].style.width = `${w + dx}px`;
        cols[i+1].style.width = `${nw - dx}px`;
      };
      const onMouseUp = () => {
        document.removeEventListener("mousemove", onMouseMove);
        document.removeEventListener("mouseup", onMouseUp);
      };
      resizer.addEventListener("mousedown", onMouseDown);
    }
  });
}

function applyData() {
  applyTemplateVisibility();
  applyDynamicTitles();

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
    const raw = getValue("executiveSummary");
    execEl.innerHTML = raw.split("\n")
      .filter(p => p.trim())
      .map(p => `<p>${formatRichTextHTML(p)}</p>`)
      .join("");
  }

  const trendNarrEl = document.getElementById("trendNarrativePreview");
  if (trendNarrEl) trendNarrEl.innerHTML = formatRichTextHTML(getValue("trendNarrative"));

  const potIncNarrEl = document.getElementById("potIncidentsNarrativePreview");
  if (potIncNarrEl) potIncNarrEl.innerHTML = formatRichTextHTML(getValue("potIncidentsNarrative"));

  const riskNarrEl = document.getElementById("riskNarrativePreview");
  if (riskNarrEl) riskNarrEl.innerHTML = formatRichTextHTML(getValue("riskNarrative"));

  const kpEnh = document.getElementById("kpEnhancementPreview");
  if (kpEnh) kpEnh.innerHTML = formatRichTextHTML(getValue("kpEnhancement"));

  const kpDecom = document.getElementById("kpDecommissioningPreview");
  if (kpDecom) kpDecom.innerHTML = formatRichTextHTML(getValue("kpDecommissioning"));

  const kpRule = document.getElementById("kpRuleImplementationPreview");
  if (kpRule) kpRule.innerHTML = formatRichTextHTML(getValue("kpRuleImplementation"));

  const deviceRows = parseCsvRows(getValue("deviceCsv"), ["device", "eps"]);
  renderTableRows(
    "deviceTableBody",
    deviceRows,
    (row) => `<tr><td style="border: 1px solid #00b0ba; padding: 6px;">${row.sno}</td><td style="border: 1px solid #00b0ba; padding: 6px; text-align: left;">${row.device}</td><td style="border: 1px solid #00b0ba; padding: 6px;">${row.eps}</td></tr>`
  );

  const topEpsCtx = document.getElementById("topEpsChart");
  if (topEpsCtx) {
    if (topEpsChart) topEpsChart.destroy();
    const colors = ["#4bc0c0", "#ffcd56", "#ff6384", "#ff9f40", "#9966ff", "#c9cbcf", "#00b050", "#be2f2f", "#1e4f9a", "#8e5ea2"];
    topEpsChart = new Chart(topEpsCtx, {
      type: "bar",
      data: {
        labels: deviceRows.map(r => r.device),
        datasets: [{
          data: deviceRows.map(r => parseFloat(r.eps) || 0),
          backgroundColor: colors.slice(0, deviceRows.length)
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          datalabels: {
            display: true,
            color: '#000',
            anchor: 'end',
            align: 'top',
            formatter: Math.round,
            font: { weight: 'bold' }
          }
        },
        scales: {
          x: { display: true, ticks: { maxRotation: 45, minRotation: 45, font: { size: 9 } } },
          y: { 
            beginAtZero: true,
            title: { display: true, text: 'EPS' }
          }
        }
      }
    });
  }

  const riskRows = parseCsvRows(getValue("riskCsv"), [
    "attackType",
    "riskScenario",
    "ciaTriad",
    "businessImpact",
    "riskRating"
  ]);
  renderTableRows(
    "riskTableBody",
    riskRows,
    (row) =>
      `<tr>
         <td style="border: 1px solid #7f7f7f; padding: 4px; text-align: center;">${row.sno}</td>
         <td style="border: 1px solid #7f7f7f; padding: 4px; text-align: left;">${row.attackType}</td>
         <td style="border: 1px solid #7f7f7f; padding: 4px; text-align: left;">${row.riskScenario}</td>
         <td style="border: 1px solid #7f7f7f; padding: 4px; text-align: center;">${row.ciaTriad}</td>
         <td style="border: 1px solid #7f7f7f; padding: 4px; text-align: left; font-size: 9px;">${row.businessImpact}</td>
         <td style="border: 1px solid #7f7f7f; padding: 4px; text-align: center; font-weight: bold;">${row.riskRating}</td>
       </tr>`
  );

  const slaTb = document.getElementById("slaTableBody");
  if (slaTb) {
    slaTb.innerHTML = `
      <tr>
        <td style="border: 1px solid #000; padding: 6px; font-weight: bold;">S-1 (High)</td>
        <td style="border: 1px solid #000; padding: 6px;">≤15 Min</td>
        <td style="border: 1px solid #000; padding: 6px;">4 Hours</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaHighJanCount")}</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaHighFebCount")}</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaHighMarCount")}</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000; padding: 6px; font-weight: bold;">S-2 (Medium)</td>
        <td style="border: 1px solid #000; padding: 6px;">≤15 Min</td>
        <td style="border: 1px solid #000; padding: 6px;">8 Hours</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaMedJanCount")}</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaMedFebCount")}</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaMedMarCount")}</td>
      </tr>
      <tr>
        <td style="border: 1px solid #000; padding: 6px; font-weight: bold;">S-3 (Low)</td>
        <td style="border: 1px solid #000; padding: 6px;">≤30 Min</td>
        <td style="border: 1px solid #000; padding: 6px;">24 Hours</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaLowJanCount")}</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaLowFebCount")}</td>
        <td style="border: 1px solid #000; padding: 6px;">${getValue("slaLowMarCount")}</td>
      </tr>
    `;
  }

  const slaPctCtx = document.getElementById("slaPctChart");
  if (slaPctCtx) {
    if (slaPctChart) slaPctChart.destroy();
    slaPctChart = new Chart(slaPctCtx, {
      type: "bar",
      data: {
        labels: ["Jan", "Feb", "Mar"],
        datasets: [
          {
            label: "S-1 (High) Within SLA (4 Hrs)",
            data: [getValue("slaHighJanPct"), getValue("slaHighFebPct"), getValue("slaHighMarPct")],
            backgroundColor: "#be2f2f"
          },
          {
            label: "S-2 (Medium) Within SLA (8 Hrs)",
            data: [getValue("slaMedJanPct"), getValue("slaMedFebPct"), getValue("slaMedMarPct")],
            backgroundColor: "#ffff00"
          },
          {
            label: "S-3 (Low) Within SLA (24 Hrs)",
            data: [getValue("slaLowJanPct"), getValue("slaLowFebPct"), getValue("slaLowMarPct")],
            backgroundColor: "#00b050"
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          y: { 
            beginAtZero: true, 
            max: 120,
            ticks: {
              callback: function(value) {
                return value + "%";
              }
            }
          }
        },
        plugins: {
          legend: { position: "bottom", labels: { font: { size: 10 } }, align: "center" }
        }
      }
    });
  }

  const inventoryRows = parseCsvRows(getValue("inventoryCsv"), ["deviceName", "count"]);
  let totalCount = 0;
  
  const inventoryHtml = inventoryRows.map((row, idx) => {
    const cnt = parseInt(row.count) || 0;
    totalCount += cnt;
    const bgColor = idx % 2 === 0 ? "#ffffff" : "#f1f9f9";
    return `<tr style="background: ${bgColor};">
      <td style="padding: 8px; border: 1px solid #4cc0c5;">${row.sno}</td>
      <td style="padding: 8px; border: 1px solid #4cc0c5; text-align: left;">${row.deviceName}</td>
      <td style="padding: 8px; border: 1px solid #4cc0c5;">${cnt}</td>
    </tr>`;
  }).join("");

  const invTb = document.getElementById("inventoryTableBody");
  if (invTb) invTb.innerHTML = inventoryHtml;

  const invTf = document.getElementById("inventoryTableFooter");
  if (invTf) {
    invTf.innerHTML = `
      <tr style="background: #ffffff; font-weight: bold;">
        <td style="padding: 10px; border: 1px solid #4cc0c5;" colspan="2">Total</td>
        <td style="padding: 10px; border: 1px solid #4cc0c5;">${totalCount}</td>
      </tr>
    `;
  }

  const invNotePreview = document.getElementById("inventoryNotePreview");
  if (invNotePreview) {
    invNotePreview.innerText = getValue("inventoryNote");
  }

  updateLogos();
  renderCharts();
  applyPremiumLayout();
  initTableResizing();
}

function exportPdf() {
  window.print();
}

function canvasToPptxData(canvas) {
  if (!canvas) return null;
  return canvas.toDataURL("image/png");
}

function getPptxCtor() {
  return window.PptxGenJS || window.pptxgen;
}

function addPptxTitleBar(slide, title) {
  slide.addText(title, {
    x: 0,
    y: 0,
    w: "100%",
    h: 0.45,
    fontSize: 14,
    bold: true,
    color: "FFFFFF",
    fill: { color: CONFIG.PPTX_NAVY },
    align: "left",
    valign: "middle",
    margin: [0.1, 0.35, 0.1, 0.35]
  });
}

function addPptxFooter(slide) {
  slide.addText(CONFIG.COMPANY_NAME, {
    x: 0.35,
    y: 5,
    w: 9,
    h: 0.35,
    fontSize: 9,
    color: "666666"
  });
}

async function exportPptx() {
  const Ctor = getPptxCtor();
  if (!Ctor) {
    alert("PptxGenJS failed to load. Check network and refresh.");
    return;
  }

  applyData();
  // Ensure charts have enough time to render before capturing canvases
  await new Promise((r) => setTimeout(r, 600));

  const pptx = new Ctor();
  const selectedLayout = pptLayoutConfig[getPptLayoutKey()] || pptLayoutConfig.LAYOUT_16X9;
  pptx.layout = selectedLayout.pptxLayout;
  pptx.author = "SNS SOC";
  pptx.title = `${getValue("customerName")} Monthly SOC Report`;

  if (getPptLayoutKey() === "A4") {
    pptx.defineLayout({ name: "CUSTOM_A4_PORTRAIT", width: 8.27, height: 11.69 });
    pptx.layout = "CUSTOM_A4_PORTRAIT";
  }

  const key = getTemplateKey();
  const copy = templateCopy[key] || templateCopy.custom;

  let slide = pptx.addSlide();
  addPptxTitleBar(slide, "Monthly SOC Report");
  slide.addText("MANAGED INCIDENT RESPONSE & REMEDIATION SERVICE", {
    x: 0.5,
    y: 0.65,
    w: 9,
    h: 0.35,
    fontSize: 12,
    color: CONFIG.PPTX_BLUE,
    bold: true
  });
  slide.addText(`${getReportMonth()}`, { x: 0.5, y: 1.1, w: 9, h: 0.5, fontSize: 24, bold: true, color: CONFIG.PPTX_NAVY });
  slide.addText(`for ${getValue("customerName")}`, { x: 0.5, y: 1.65, w: 9, h: 0.4, fontSize: 16 });
  slide.addText(
    [
      `Version 1.0`,
      `Dated ${getValue("dateRange")}`,
      `Prepared By ${getValue("preparedBy")}`,
      `Reviewed By ${STATIC_REVIEWED_BY}`,
      `Approved By ${STATIC_APPROVED_BY}`,
      `Submitted On ${getValue("submittedOn")}`
    ].join("\n"),
    { x: 0.5, y: 2.3, w: 5.5, h: 2, fontSize: 11 }
  );
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: 6.8, y: 0.65, w: 1.2, h: 0.55 });
  }
  if (clientLogoDataUrl) {
    slide.addImage({ data: clientLogoDataUrl, x: 8.1, y: 0.65, w: 1.2, h: 0.55 });
  }
  addPptxFooter(slide);

  slide = pptx.addSlide();
  addPptxTitleBar(slide, "Document Revision History");
  slide.addText(
    getValue("documentRevisionHistory")
      .split("\n")
      .filter(Boolean)
      .map((l) => `• ${l}`)
      .join("\n"),
    { x: 0.5, y: 0.65, w: 9, h: 3, fontSize: 11 }
  );
  addPptxFooter(slide);

  slide = pptx.addSlide();
  addPptxTitleBar(slide, "The Engagement");
  
  const introText = `${getValue("customerName")} has engaged with SNS to monitor and review the entity's security.\n\n`;
  const execBody = getValue("executiveSummary");
  slide.addText([...parseRichTextPptx(introText), ...parseRichTextPptx(execBody)], { x: 0.5, y: 0.65, w: 9, h: 3.5, fontSize: 11 });
  addPptxFooter(slide);

  if (key === "apraava") {
    slide = pptx.addSlide();
    addPptxTitleBar(slide, "Potential Incidents and Alert Summary");
    
    const quadIds = ["aqChart1", "aqChart2", "aqChart3", "aqChart4"];
    quadIds.forEach((id, idx) => {
      const cImg = canvasToPptxData(document.getElementById(id));
      if (cImg) {
        const isRight = idx % 2 !== 0;
        const isBot = Math.floor(idx / 2) > 0;
        slide.addImage({ data: cImg, x: isRight ? 5.0 : 0.5, y: isBot ? 3.0 : 0.8, w: 4.4, h: 2.0 });
      }
    });
    addPptxFooter(slide);
  }

  slide = pptx.addSlide();
  addPptxTitleBar(slide, copy.trendTitle);
  const trendImg = canvasToPptxData(document.getElementById("trendChart"));
  if (trendImg) {
    slide.addImage({ data: trendImg, x: 0.8, y: 0.65, w: 8.4, h: 3.5 });
  }
  slide.addText(parseRichTextPptx(getValue("trendNote")), { x: 0.5, y: 4.3, w: 9, h: 0.3, fontSize: 13, bold: true });
  slide.addText(parseRichTextPptx(getValue("trendNarrative")), { x: 0.5, y: 4.6, w: 9, h: 0.7, fontSize: 11 });
  addPptxFooter(slide);

  slide = pptx.addSlide();
  addPptxTitleBar(slide, "Potential Incidents");
  const potIncImg = canvasToPptxData(document.getElementById("potIncidentsChart"));
  if (potIncImg) {
    slide.addImage({ data: potIncImg, x: 0.8, y: 0.65, w: 8.4, h: 3.5 });
  }
  slide.addText(parseRichTextPptx(getValue("potIncidentsNarrative")), { x: 0.5, y: 4.4, w: 9, h: 0.8, fontSize: 11 });
  addPptxFooter(slide);

  // New EPS Trend Slide
  slide = pptx.addSlide();
  addPptxTitleBar(slide, "EPS Trend Plot");
  const epsImg = canvasToPptxData(document.getElementById("epsTrendChart"));
  if (epsImg) {
    slide.addImage({ data: epsImg, x: 0.5, y: 0.8, w: 3.5, h: 4.0 });
  }
  const epsTblRows = parseCsvRows(getValue("epsTableCsv"), ["reportingDevice", "eventName", "jan", "feb", "mar"]);
  if (epsTblRows.length > 0) {
    const tableData = [
      [
        { text: "S.No", options: { bold: true, fill: { color: "F3F6FB" } } },
        { text: "Reporting Device", options: { bold: true, fill: { color: "F3F6FB" } } },
        { text: "Event Name", options: { bold: true, fill: { color: "F3F6FB" } } },
        { text: "Jan", options: { bold: true, fill: { color: "F3F6FB" } } },
        { text: "Feb", options: { bold: true, fill: { color: "F3F6FB" } } },
        { text: "Mar", options: { bold: true, fill: { color: "F3F6FB" } } }
      ],
      ...epsTblRows.map(r => [String(r.sno), r.reportingDevice, r.eventName, r.jan, r.feb, r.mar])
    ];
    slide.addTable(tableData, { x: 4.2, y: 0.8, w: 5.4, fontSize: 8, colW: [0.5, 1.2, 1.3, 0.8, 0.8, 0.8] });
  }
  addPptxFooter(slide);

  slide = pptx.addSlide();
  addPptxTitleBar(slide, "Top 10 Devices Contributing To Highest EPS");
  
  const devImg = canvasToPptxData(document.getElementById("topEpsChart"));
  if (devImg) {
    slide.addImage({ data: devImg, x: 0.5, y: 0.8, w: 4.5, h: 4.0 });
  }

  const devices = parseCsvRows(getValue("deviceCsv"), ["device", "eps"]);
  if (devices.length > 0) {
    const devRows = [
      [
        { text: "S.No", options: { bold: true, fill: { color: "8EEDE4" } } },
        { text: "Reporting Device", options: { bold: true, fill: { color: "8EEDE4" } } },
        { text: "AVG EPS", options: { bold: true, fill: { color: "8EEDE4" } } }
      ],
      ...devices.map((d) => [String(d.sno), d.device, d.eps])
    ];
    slide.addTable(devRows, { x: 5.2, y: 0.8, w: 4.3, fontSize: 10, colW: [0.6, 2.5, 1.2] });
  }
  addPptxFooter(slide);

  const risks = parseCsvRows(getValue("riskCsv"), [
    "attackType",
    "riskScenario",
    "ciaTriad",
    "businessImpact",
    "riskRating"
  ]);
  const risksRowsPerSlide = 4;
  for (let i = 0; i < risks.length; i += risksRowsPerSlide) {
    const chunk = risks.slice(i, i + risksRowsPerSlide);
    slide = pptx.addSlide();
    addPptxTitleBar(slide, i === 0 ? "Potential Incidents – Risks Mitigated" : "Contn.,");
    if (i === 0) slide.addText(parseRichTextPptx(getValue("riskNarrative")), { x: 0.35, y: 0.65, w: 9.3, h: 0.8, fontSize: 10 });

    const rows = [
      [
        { text: "S.No", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } },
        { text: "Attack Type", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } },
        { text: "Risk Scenario", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } },
        { text: "Type of Risk(s)\nCIA Triad", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } },
        { text: "Potential Business Impact(s)", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } },
        { text: "Risk Rating", options: { bold: true, fill: { color: "E2E5E9" }, align: "center", fontFace: "Calibri" } }
      ],
      ...chunk.map((r) => [
        { text: String(r.sno), options: { align: "center", fontFace: "Calibri" } },
        { text: r.attackType, options: { fontFace: "Calibri" } },
        { text: r.riskScenario, options: { fontFace: "Calibri" } },
        { text: r.ciaTriad, options: { align: "center", fontFace: "Calibri" } },
        { text: r.businessImpact, options: { fontFace: "Calibri" } },
        { text: r.riskRating, options: { align: "center", fontFace: "Calibri" } }
      ])
    ];
    const tableY = i === 0 ? 1.5 : 0.6;
    slide.addTable(rows, { 
       x: 0.35, y: tableY, w: 9.3, fontSize: 9, 
       border: { type: 'solid', color: 'FFFFFF', pt: 1 }, 
       colW: [0.4, 1.4, 2.0, 1.2, 3.5, 0.8] 
    });
    addPptxFooter(slide);
  }

  slide = pptx.addSlide();
  addPptxTitleBar(slide, "Response Time SLA");
  const slaRows = [
    [{ text: "Severity Level", options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } }, { text: "SLA Response", options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } }, { text: "SLA Resolution", options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } }, { text: "Jan", options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } }, { text: "Feb", options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } }, { text: "Mar", options: { bold: true, fill: { color: "F3F6FB" }, align: "center" } }],
    ["S-1 (High)", "≤15 Min", "4 Hours", getValue("slaHighJanCount"), getValue("slaHighFebCount"), getValue("slaHighMarCount")],
    ["S-2 (Medium)", "≤15 Min", "8 Hours", getValue("slaMedJanCount"), getValue("slaMedFebCount"), getValue("slaMedMarCount")],
    ["S-3 (Low)", "≤30 Min", "24 Hours", getValue("slaLowJanCount"), getValue("slaLowFebCount"), getValue("slaLowMarCount")]
  ];
  slide.addTable(slaRows, { x: 0.5, y: 0.8, w: 9.0, fontSize: 10, border: { pt: 1, color: "000000" } });
  slide.addText("Remediation Time SLA in %", { x: 0.5, y: 2.5, w: 9, h: 0.4, fontSize: 18, bold: true });
  const slaPctImg = canvasToPptxData(document.getElementById("slaPctChart"));
  if (slaPctImg) slide.addImage({ data: slaPctImg, x: 1.0, y: 3.0, w: 8.0, h: 2.2 });
  addPptxFooter(slide);

  slide = pptx.addSlide();
  addPptxTitleBar(slide, "Integrated Device Inventory");
  const inv = parseCsvRows(getValue("inventoryCsv"), ["deviceName", "count"]);
  let invTotal = 0;
  inv.forEach(r => { invTotal += (parseInt(r.count) || 0); });
  const invRows = [
    [{ text: "S.NO", options: { bold: true, color: "FFFFFF", fill: { color: "4cc0c5" }, align: "center" } }, { text: "DEVICE NAME", options: { bold: true, color: "FFFFFF", fill: { color: "4cc0c5" }, align: "center" } }, { text: "COUNT", options: { bold: true, color: "FFFFFF", fill: { color: "4cc0c5" }, align: "center" } }],
    ...inv.map((r, i) => [{ text: String(r.sno), options: { align: "center" } }, { text: r.deviceName, options: { align: "left" } }, { text: String(r.count), options: { align: "center" } }]),
    [{ text: "Total", options: { bold: true, align: "center", colspan: 2 } }, { text: String(invTotal), options: { bold: true, align: "center" } }]
  ];
  slide.addTable(invRows, { x: 0.5, y: 0.8, w: 9.0, fontSize: 12, border: { pt: 1, color: "4cc0c5" }, valign: "middle" });
  const invNote = getValue("inventoryNote");
  if (invNote) slide.addText(parseRichTextPptx(invNote), { x: 0.5, y: 3.8, w: 9.0, h: 1.0, fontSize: 11, color: "000000", fontFace: "Calibri", valign: "top" });
  addPptxFooter(slide);

  slide = pptx.addSlide();
  addPptxTitleBar(slide, "Key Points – Overall Summary");
  const yStart = 0.8;
  slide.addText("1. Report Enhancement:", { x: 0.5, y: yStart, w: 9, h: 0.3, fontSize: 13, bold: true, fontFace: "Calibri" });
  slide.addText(parseRichTextPptx(getValue("kpEnhancement")), { x: 0.5, y: yStart + 0.3, w: 9, h: 0.8, fontSize: 11, fontFace: "Calibri", valign: "top" });
  slide.addText("2. Device Decommissioning:", { x: 0.5, y: yStart + 1.2, w: 9, h: 0.3, fontSize: 13, bold: true, fontFace: "Calibri" });
  slide.addText(parseRichTextPptx(getValue("kpDecommissioning")), { x: 0.5, y: yStart + 1.5, w: 9, h: 1.2, fontSize: 11, fontFace: "Calibri", valign: "top" });
  slide.addText("3. Rule Implementation:", { x: 0.5, y: yStart + 2.8, w: 9, h: 0.3, fontSize: 13, bold: true, fontFace: "Calibri" });
  slide.addText(parseRichTextPptx(getValue("kpRuleImplementation")), { x: 0.5, y: yStart + 3.1, w: 9, h: 1.2, fontSize: 11, fontFace: "Calibri", valign: "top" });
  addPptxFooter(slide);

  slide = pptx.addSlide();
  if (snsLogoDataUrl) slide.addImage({ data: snsLogoDataUrl, x: 0.5, y: 0.5, w: 2.8 });
  slide.addText("Your Trusted Security Advisor", { x: 0.5, y: 1.8, w: 8.5, h: 1.0, fontSize: 48, bold: true, color: "022a5b", align: "left", fontFace: "Calibri" });
  const contactLines = [{ text: "FOR FURTHER DETAILS:", options: { bold: true, color: "022a5b", fontSize: 13 } }, { text: "Diptesh Saha", options: { bold: true, color: "1e4f9a", fontSize: 18 } }, { text: "CISO & Practice Head - Cyber Security & Managed Security", options: { color: "1e4f9a", fontSize: 13 } }, { text: "Contact No. 7338882888", options: { color: "1e4f9a", fontSize: 13 } }, { text: "diptesh.s@snsin.com", options: { color: "1e4f9a", fontSize: 13 } }];
  slide.addText(contactLines, { x: 0.5, y: 2.8, w: 8.5, h: 2.5, align: "left", fontFace: "Calibri", valign: "top" });
  addPptxFooter(slide);

  const fname = `${sanitizeFilename(getValue("customerName"))}_${sanitizeFilename(getReportMonth())}_Monthly_SOC_Report.pptx`;
  await pptx.writeFile(fname);
}

async function exportDocx() {
  if (!window.docx || !window.saveAs) {
    alert("DOCX libraries failed to load.");
    return;
  }
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType } = window.docx;
  const customer = getValue("customerName");
  const month = getReportMonth();
  const dateRange = getValue("dateRange");
  const preparedBy = getValue("preparedBy");
  const revision = getValue("documentRevisionHistory");
  const executiveSummary = getValue("executiveSummary");
  const risks = parseCsvRows(getValue("riskCsv"), ["attackType", "riskScenario", "ciaTriad", "businessImpact", "riskRating"]);
  const inventories = parseCsvRows(getValue("inventoryCsv"), ["deviceName", "count"]);
  const slas = parseCsvRows(getValue("slaCsv"), ["severity", "response", "resolution", "count"]);
  
  const children = [
    new Paragraph({ children: [new TextRun({ text: "MANAGED INCIDENT RESPONSE & REMEDIATION SERVICE", bold: true })] }),
    new Paragraph({ children: [new TextRun({ text: `Monthly SOC Report | ${month}`, bold: true, size: 32 })] }),
    new Paragraph(`for ${customer}`),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "Document Revision History", bold: true })] }),
    new Paragraph(revision),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "The Engagement", bold: true })] }),
    new Paragraph(executiveSummary.replace(/\[MONTH\]/gi, month)),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "Potential Incidents - Risks Mitigated", bold: true })] }),
    ...risks.map(risk => new Paragraph(`${risk.sno}. ${risk.attackType} | ${risk.riskScenario} | ${risk.ciaTriad} | ${risk.businessImpact} | ${risk.riskRating}`)),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "Integrated Device Inventory", bold: true })] }),
    ...inventories.map(inv => new Paragraph(`${inv.sno}. ${inv.deviceName}: ${inv.count}`)),
    new Paragraph(""),
    new Paragraph({ children: [new TextRun({ text: "Response Time SLA", bold: true })] }),
    ...slas.map(sla => new Paragraph(`${sla.severity}: Response ${sla.response}, Resolution ${sla.resolution}, Count ${sla.count}`))
  ];

  const doc = new Document({ sections: [{ children }] });
  const blob = await Packer.toBlob(doc);
  saveAs(blob, `${sanitizeFilename(customer)}_${sanitizeFilename(month)}_Monthly_SOC_Report.docx`);
}

function initCustomDropdown() {
  const trigger = document.getElementById("dropdownTrigger");
  const menu = document.getElementById("dropdownMenu");
  const searchInput = document.getElementById("dropdownSearchInput");
  const options = document.querySelectorAll(".option");
  const hiddenInput = document.getElementById("templatePreset");
  const selectedSpan = document.getElementById("selectedTemplate");

  if (!trigger || !menu) return;

  trigger.addEventListener("click", () => {
    menu.classList.toggle("active");
    if (menu.classList.contains("active")) {
      searchInput.value = "";
      filterOptions("");
      searchInput.focus();
    }
  });

  searchInput.addEventListener("input", (e) => {
    filterOptions(e.target.value.toLowerCase());
  });

  function filterOptions(term) {
    options.forEach(opt => {
      const text = opt.textContent.toLowerCase();
      opt.style.display = text.includes(term) ? "block" : "none";
    });
  }

  options.forEach(opt => {
    opt.addEventListener("click", () => {
      const val = opt.getAttribute("data-value");
      const text = opt.textContent;
      hiddenInput.value = val;
      selectedSpan.textContent = text;
      options.forEach(o => o.classList.remove("selected"));
      opt.classList.add("selected");
      menu.classList.remove("active");
      setPresetData(val);
    });
  });

  document.addEventListener("click", (e) => {
    if (!trigger.contains(e.target) && !menu.contains(e.target)) {
      menu.classList.remove("active");
    }
  });
}

async function init() {
  await initCustomDropdown();
  document.querySelectorAll(".panel input, .panel textarea, .panel select").forEach(el => {
    if (el.id !== "dropdownSearchInput") el.addEventListener("input", applyData);
  });
  document.getElementById("applyBtn").addEventListener("click", applyData);
  document.getElementById("printBtn").addEventListener("click", exportPdf);
  document.getElementById("pptxBtn").addEventListener("click", exportPptx);
  document.getElementById("docxBtn").addEventListener("click", exportDocx);
  document.getElementById("pptLayout").addEventListener("change", applyData);
  
  ["snsLogoInput", "clientLogoInput", "puzzleLogoInput"].forEach(id => {
    bindImageInput(id, (v) => {
      if (id === "snsLogoInput") snsLogoDataUrl = v;
      if (id === "clientLogoInput") clientLogoDataUrl = v;
      if (id === "puzzleLogoInput") puzzleLogoDataUrl = v;
      applyData();
    });
  });

  const csvHandlers = ["trendCsv", "potIncidentsCsv", "epsTableCsv", "deviceCsv", "riskCsv", "inventoryCsv"];
  csvHandlers.forEach(id => {
    const input = document.getElementById(id + "File");
    if (input) {
      input.addEventListener("change", (e) => {
        const file = e.target.files[0];
        if (file) {
          const reader = new FileReader();
          reader.onload = (event) => {
            document.getElementById(id).value = event.target.result;
            applyData();
          };
          reader.readAsText(file);
        }
      });
    }
  });

  applyData();
}

window.onload = init;
