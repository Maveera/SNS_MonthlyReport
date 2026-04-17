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
let defaultPuzzleDataUrl = "";
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

function getValue(id) {
  const el = document.getElementById(id);
  return el ? el.value.trim() : "";
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
  const monthReplaced = text.replace(/\[MONTH\]/g, `**${getReportMonth()}**`);
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

const INLINE_EDITABLE_NARRATIVES = {
  executiveSummaryPreview: "executiveSummary",
  totPotIncNarrativePreview: "totPotIncNarrative",
  trendNotePreview: "trendNote",
  trendNarrativePreview: "trendNarrative",
  keyPointsSummaryPreview: "keyPointsSummary",
  inventoryNotePreview: "inventoryNote"
};

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function normalizeMonthPlaceholderFromMarkdown(source) {
  const m = getReportMonth();
  return String(source || "").replace(new RegExp(`\\*\\*${escapeRegExp(m)}\\*\\*`, "g"), "[MONTH]");
}

function shouldSkipApplyForNode(node) {
  if (!node || typeof node.contains !== "function") return false;
  const ae = document.activeElement;
  if (!ae || !node.contains(ae)) return false;
  return ae.isContentEditable === true;
}

function inlineHtmlToMarkdown(el) {
  let out = "";
  el.childNodes.forEach((node) => {
    if (node.nodeType === 3) {
      out += node.textContent;
    } else if (node.nodeType === 1) {
      const tag = node.tagName;
      if (tag === "STRONG" || tag === "B" || node.classList.contains("fw-bold")) {
        out += `**${node.textContent}**`;
      } else if (tag === "EM" || tag === "I") {
        out += `*${node.textContent}*`;
      } else if (tag === "BR") {
        out += "\n";
      } else if (tag === "P") {
        const inner = inlineHtmlToMarkdown(node);
        out += (out ? "\n\n" : "") + inner;
      } else {
        out += inlineHtmlToMarkdown(node);
      }
    }
  });
  return out;
}

function narrativeHtmlToMarkdownSource(html) {
  const wrap = document.createElement("div");
  wrap.innerHTML = String(html || "").trim();
  const ps = wrap.querySelectorAll(":scope > p");
  if (ps.length) {
    return Array.from(ps)
      .map((p) => inlineHtmlToMarkdown(p))
      .join("\n\n");
  }
  return inlineHtmlToMarkdown(wrap);
}

function addEngagementNativeSlide(pptx, slideW, slideH) {
  const slide = pptx.addSlide();
  const customer = getValue("customerName") || "Customer";
  const summaryLines = getValue("executiveSummary")
    .split("\n")
    .map((s) => s.trim())
    .filter(Boolean);
  const puzzleData = puzzleLogoDataUrl || defaultPuzzleDataUrl || "";
  const baseTxt = { fontFace: "Calibri", fontSize: 11, color: "000000" };
  const textParts = [
    { text: customer, options: { ...baseTxt, bold: true } },
    {
      text: " has engaged with SNS to monitor and review the entity's security.",
      options: { ...baseTxt }
    }
  ];
  summaryLines.forEach((line, idx) => {
    textParts.push({ text: idx === 0 ? "\n\n" : "\n\n", options: { ...baseTxt } });
    textParts.push(...parseRichTextPptx(line, baseTxt));
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: slideW - 1.35, y: 0.18, w: 1.2, h: 0.5 });
  }
  slide.addText("The Engagement", {
    x: 0.5,
    y: 0.35,
    w: 4.2,
    h: 0.5,
    fontSize: 24,
    bold: true,
    color: "000000",
    fontFace: "Calibri"
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 0.48,
    y: 0.86,
    w: 4.58,
    h: 3.92,
    line: { color: "C8E0EA", pt: 0.5 },
    fill: { color: "E1F4F8" }
  });
  slide.addText(textParts, {
    x: 0.58,
    y: 0.94,
    w: 4.38,
    h: 3.76,
    valign: "top",
    fontFace: "Calibri",
    color: "000000"
  });
  if (puzzleData) {
    slide.addImage({ data: puzzleData, x: 5.2, y: 0.9, w: 4.25, h: 3.85 });
  }
  slide.addShape(pptx.ShapeType.rect, {
    x: 0,
    y: slideH - 0.42,
    w: slideW,
    h: 0.42,
    line: { color: "1E4F9A", pt: 0 },
    fill: { color: "1E4F9A" }
  });
  return slide;
}

function initInlineEditableReport() {
  const report = document.getElementById("reportRoot");
  if (!report) return;
  if (report.dataset.inlineEditableBound === "1") return;
  report.dataset.inlineEditableBound = "1";

  const skipField = new Set(["month", "reviewedBy", "approvedBy"]);
  report.querySelectorAll("[data-field]").forEach((el) => {
    const key = el.getAttribute("data-field");
    if (!key || skipField.has(key)) return;
    if (!fields[key]) return;
    el.setAttribute("contenteditable", "true");
    el.setAttribute("spellcheck", "true");
    el.classList.add("report-editable-field");
  });

  Object.keys(INLINE_EDITABLE_NARRATIVES).forEach((id) => {
    const el = document.getElementById(id);
    if (!el) return;
    el.setAttribute("contenteditable", "true");
    el.setAttribute("spellcheck", "true");
    el.classList.add("report-editable-narrative");
    el.dataset.syncField = INLINE_EDITABLE_NARRATIVES[id];
  });

  const numericCells = [
    ["totPotIncTableHigh", "totPotIncHigh"],
    ["totPotIncTableMed", "totPotIncMedium"],
    ["fortiSiemTableHigh", "fortiAlertHigh"],
    ["fortiSiemTableMed", "fortiAlertMedium"],
    ["fortiSiemTableLow", "fortiAlertLow"],
    ["tfPosTrueTableHigh", "tfPosTrueHigh"],
    ["tfPosTrueTableMed", "tfPosTrueMedium"],
    ["tfPosTrueTableLow", "tfPosTrueLow"],
    ["tfPosFalseTableHigh", "tfPosFalseHigh"],
    ["tfPosFalseTableMed", "tfPosFalseMedium"],
    ["tfPosFalseTableLow", "tfPosFalseLow"]
  ];
  numericCells.forEach(([cellId, inputId]) => {
    const el = document.getElementById(cellId);
    if (!el) return;
    el.setAttribute("contenteditable", "true");
    el.classList.add("report-editable-field");
    el.dataset.syncInput = inputId;
  });

  report.addEventListener("focusout", (e) => {
    const t = e.target;
    if (!t.closest || !t.closest("#reportRoot")) return;
    const rel = e.relatedTarget;

    const syncCell = t.closest("[data-sync-input]");
    if (syncCell && syncCell.dataset.syncInput) {
      if (rel && syncCell.contains(rel)) return;
      const inp = document.getElementById(syncCell.dataset.syncInput);
      if (inp) inp.value = syncCell.textContent.replace(/\s+/g, " ").trim();
      applyData();
      return;
    }

    const narrative = t.closest(".report-editable-narrative");
    if (narrative && narrative.dataset.syncField) {
      if (rel && narrative.contains(rel)) return;
      const inp = document.getElementById(narrative.dataset.syncField);
      if (inp) {
        let src = narrativeHtmlToMarkdownSource(narrative.innerHTML);
        src = normalizeMonthPlaceholderFromMarkdown(src);
        inp.value = src;
      }
      applyData();
      return;
    }

    const fieldEl = t.closest("[data-field].report-editable-field");
    const key = fieldEl && fieldEl.getAttribute("data-field");
    if (fieldEl && key && fields[key]) {
      if (rel && fieldEl.contains(rel)) return;
      const inp = document.getElementById(fields[key]);
      if (inp) inp.value = fieldEl.textContent.trim();
      applyData();
    }
  });
}

const DRAWING_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

function pptxIsInTable(pEl) {
  let x = pEl.parentElement;
  while (x) {
    if (x.localName === "tbl" && x.namespaceURI === DRAWING_A_NS) return true;
    x = x.parentElement;
  }
  return false;
}

function extractParagraphsFromSlideXml(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const paras = doc.getElementsByTagNameNS(DRAWING_A_NS, "p");
  const lines = [];
  for (let i = 0; i < paras.length; i++) {
    const pEl = paras[i];
    if (pptxIsInTable(pEl)) continue;
    const ts = pEl.getElementsByTagNameNS(DRAWING_A_NS, "t");
    let line = "";
    for (let j = 0; j < ts.length; j++) line += ts[j].textContent || "";
    const t = line.replace(/\r/g, "").trim();
    if (t) lines.push(t);
  }
  return lines;
}

function extractTablesFromSlideXml(xmlString) {
  const doc = new DOMParser().parseFromString(xmlString, "text/xml");
  const tbls = doc.getElementsByTagNameNS(DRAWING_A_NS, "tbl");
  const tables = [];
  for (let ti = 0; ti < tbls.length; ti++) {
    const rows = [];
    const trs = tbls[ti].getElementsByTagNameNS(DRAWING_A_NS, "tr");
    for (let ri = 0; ri < trs.length; ri++) {
      const cells = [];
      const tcs = trs[ri].getElementsByTagNameNS(DRAWING_A_NS, "tc");
      for (let ci = 0; ci < tcs.length; ci++) {
        const txBody = tcs[ci].getElementsByTagNameNS(DRAWING_A_NS, "txBody")[0];
        let cellText = "";
        if (txBody) {
          const ps = txBody.getElementsByTagNameNS(DRAWING_A_NS, "p");
          for (let pi = 0; pi < ps.length; pi++) {
            const ts = ps[pi].getElementsByTagNameNS(DRAWING_A_NS, "t");
            for (let tj = 0; tj < ts.length; tj++) cellText += ts[tj].textContent || "";
            if (pi < ps.length - 1) cellText += "\n";
          }
        }
        cells.push(cellText.trim());
      }
      rows.push(cells);
    }
    tables.push(rows);
  }
  return tables;
}

function csvEscapeCell(v) {
  const s = String(v ?? "");
  if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function csvEscapeRow(cells) {
  return cells.map(csvEscapeCell).join(",");
}

function classifySlideKind(paragraphs) {
  const head = (paragraphs[0] || "").trim();
  if (head.includes("Document Revision History")) return "revision";
  if (head.includes("The Engagement")) return "engagement";
  if (head.includes("Total Potential Incident")) return "totPot";
  if (head.includes("Potential Incident Tickets Trend")) return "trend";
  if (head.includes("Potential Alerts - Risks Mitigated")) return "risk";
  if (head.includes("Rule-Based Severity")) return "ruleSeverity";
  if (head.includes("Response Time SLA") || head.includes("Remediation Time SLA")) return "sla";
  if (head.includes("Total Alerts Triggered in FortiSIEM")) return "forti";
  if (head.includes("Total Number Of True")) return "tfPos";
  if (head.includes("EPS Trend Plot")) return "epsTrend";
  if (head.includes("Highest EPS Consuming Events")) return "epsEvents";
  if (head.includes("Overall Support Ticket")) return "support";
  if (head.includes("Integrated Device Inventory")) return "inventory";
  if (head.includes("Key Points")) return "keyPoints";
  if (head.includes("Your Trusted Security Advisor")) return "contact";
  return "unknown";
}

function normalizeMonthForImportText(text) {
  const m = getReportMonth();
  return String(text || "").replace(new RegExp(escapeRegExp(m), "g"), "[MONTH]");
}

function importFromRevisionParagraphs(paragraphs, updates) {
  paragraphs.forEach((line) => {
    if (/^Dated\s+/i.test(line)) updates.dateRange = line.replace(/^Dated\s+/i, "").trim();
    if (/^Prepared By\s+/i.test(line)) updates.preparedBy = line.replace(/^Prepared By\s+/i, "").trim();
    if (/^Submitted On\s+/i.test(line)) updates.submittedOn = line.replace(/^Submitted On\s+/i, "").trim();
  });
}

function importFromEngagementParagraphs(paragraphs, updates) {
  const leadIdx = paragraphs.findIndex((p) => p.includes("has engaged with SNS"));
  if (leadIdx < 0) return;
  const lead = paragraphs[leadIdx];
  const m = lead.match(/^(.+?)\s+has engaged with SNS/i);
  if (m) updates.customerName = m[1].replace(/\*\*/g, "").trim();
  const rest = paragraphs
    .slice(leadIdx + 1)
    .filter((p) => p !== "The Engagement" && !/^Version\s+\d/i.test(p));
  if (rest.length) {
    updates.executiveSummary = normalizeMonthForImportText(rest.join("\n\n"));
  }
}

function importFromTotPotParagraphs(paragraphs, updates) {
  const tidx = paragraphs.findIndex((p) => p.includes("Total Potential Incident"));
  const slice = tidx >= 0 ? paragraphs.slice(tidx + 1) : paragraphs;
  const monthWords = /^(January|February|March|April|May|June|July|August|September|October|November|December)$/i;
  const junk = (p) =>
    /^(High|Medium|Low)$/i.test(p) ||
    monthWords.test(p) ||
    /^Incident Count$/i.test(p) ||
    /^\d{1,4}$/.test(p.trim());
  const candidates = slice.filter((p) => p.length > 40 && !junk(p));
  const narrative = candidates.length ? candidates[candidates.length - 1] : "";
  if (narrative) updates.totPotIncNarrative = normalizeMonthForImportText(narrative);
}

function importFromTrendParagraphs(paragraphs, updates) {
  const tidx = paragraphs.findIndex((p) => p.includes("Potential Incident Tickets Trend"));
  const detIdx = paragraphs.findIndex((p) => p.includes("Incident details are provided below"));
  const slice = paragraphs.slice(tidx + 1, detIdx >= 0 ? detIdx : undefined);
  const mid = slice.filter((p) => p.length > 3);
  if (mid.length >= 1) updates.trendNote = normalizeMonthForImportText(mid[0]);
  if (mid.length >= 2) updates.trendNarrative = normalizeMonthForImportText(mid.slice(1).join("\n\n"));
}

function importRiskTableToCsv(rows) {
  if (!rows.length) return "";
  const start = rows[0][0] && /S\.?No|Attack/i.test(rows[0][0]) ? 1 : 0;
  const out = [];
  for (let i = start; i < rows.length; i++) {
    const r = rows[i];
    if (r.length < 6) continue;
    if (/No risk rows/i.test(String(r[1] || ""))) continue;
    out.push(csvEscapeRow([r[1], r[2], r[3], r[4], r[5]]));
  }
  return out.join("\n");
}

function importFromRiskSlide(paragraphs, tables, updates, riskAccum) {
  const narIdx = paragraphs.findIndex((p) => p.includes("Potential Alerts - Risks Mitigated"));
  const slice = narIdx >= 0 ? paragraphs.slice(narIdx + 1) : paragraphs;
  const narrative = slice.filter((p) => p.length > 40).join("\n\n");
  if (narrative && !riskAccum.narrativeDone) {
    updates.riskNarrative = normalizeMonthForImportText(narrative);
    riskAccum.narrativeDone = true;
  }
  tables.forEach((tbl) => {
    const csv = importRiskTableToCsv(tbl);
    if (csv) riskAccum.csvParts.push(csv);
  });
}

function importFromSlaParagraphs(paragraphs, updates) {
  const t = paragraphs.join("\n");
  const m1 = t.match(/Overall,\s*(\d+)\s*incidents/i);
  if (m1) updates.slaIncidentCount = m1[1];
  const m2 = t.match(/Remediation Time for\s*(\d+)\s*closed/i);
  if (m2) updates.slaClosedIncidentCount = m2[1];
}

function wrapTfImportedNote(body) {
  const b = String(body || "").trim();
  if (!b) return "Note:";
  if (/^Note:/i.test(b)) return b;
  return `Note:\n${b}`;
}

function importTfNotesFromParagraphs(paragraphs, updates) {
  const blob = paragraphs.join("\n");
  const segs = `\n${blob}`.split(/\n\s*Note:\s*/i);
  if (segs.length < 2) return;
  const bodies = segs
    .slice(1)
    .map((s) => s.trim().replace(/^•\s*/gm, "").trim())
    .filter(Boolean);
  if (bodies[0]) updates.tfPosNoteTrue = wrapTfImportedNote(bodies[0]);
  if (bodies[1]) updates.tfPosNoteFalse = wrapTfImportedNote(bodies[1]);
}

function importFromEpsTrend(paragraphs, tables, updates) {
  if (tables.length && tables[0].length) {
    const r0 = tables[0][0];
    if (r0 && r0.length >= 2) {
      updates.epsMonthLabel = String(r0[0]).trim();
      const n = parseFloat(String(r0[1]).replace(/,/g, ""));
      if (Number.isFinite(n)) updates.epsTotalValue = String(n);
    }
  }
}

function importEpsEventsTableToCsv(rows) {
  if (!rows.length) return "";
  const start = rows[0][0] && /S\.?No|Reporting/i.test(rows[0][0]) ? 1 : 0;
  const out = [];
  for (let i = start; i < rows.length; i++) {
    const r = rows[i];
    if (r.length < 5) continue;
    if (/No rows found/i.test(String(r[1] || ""))) continue;
    out.push(csvEscapeRow([r[1], r[2], r[3], r[4], r[5]]));
  }
  return out.join("\n");
}

function importSupportFromTables(tables, updates) {
  if (tables.length < 1) return;
  const major = tables[0];
  const majorRows = [];
  for (let i = 0; i < major.length; i++) {
    const r = major[i];
    if (r.length < 4) continue;
    if (/Created by/i.test(String(r[0] || ""))) continue;
    if (/Grand Total/i.test(String(r[0] || ""))) continue;
    if (/^Major tickets$/i.test(String(r[0] || ""))) continue;
    if (/^Status$/i.test(String(r[0] || ""))) continue;
    if (/^Closed$/i.test(String(r[1] || "")) && i < 4) continue;
    const by = String(r[0] || "").trim();
    if (!by || by === "Grand Total") continue;
    majorRows.push(csvEscapeRow([by, r[1], r[2]]));
  }
  if (majorRows.length) updates.supportMajorCsv = majorRows.join("\n");

  if (tables.length >= 2) {
    const minor = tables[1];
    const minorRows = [];
    for (let i = 0; i < minor.length; i++) {
      const r = minor[i];
      if (r.length < 2) continue;
      if (/Created by/i.test(String(r[0] || ""))) continue;
      if (/Minor tickets$/i.test(String(r[0] || ""))) continue;
      if (/Grand Total/i.test(String(r[0] || ""))) continue;
      const by = String(r[0] || "").trim();
      if (!by) continue;
      minorRows.push(csvEscapeRow([by, r[1]]));
    }
    if (minorRows.length) updates.supportMinorCsv = minorRows.join("\n");
  }
}

function importInventoryFromSlide(paragraphs, tables, updates) {
  if (tables.length && tables[0].length) {
    const rows = tables[0];
    const start = rows[0][0] && /S\.?No|Device/i.test(rows[0][0]) ? 1 : 0;
    const out = [];
    for (let i = start; i < rows.length; i++) {
      const r = rows[i];
      if (r.length < 2) continue;
      if (/^Total$/i.test(String(r[1] || "").trim())) continue;
      out.push(csvEscapeRow([r[1], r[2]]));
    }
    if (out.length) updates.inventoryCsv = out.join("\n");
  }
  const tidx = paragraphs.findIndex((p) => p.includes("Integrated Device Inventory"));
  const tail = paragraphs.slice(tidx + 1).filter((p) => p.length > 20);
  if (tail.length) updates.inventoryNote = normalizeMonthForImportText(tail[tail.length - 1]);
}

function keyPointsFromParagraphs(paragraphs, updates) {
  const kidx = paragraphs.findIndex((p) => p.includes("Key Points"));
  const slice = kidx >= 0 ? paragraphs.slice(kidx + 1) : [];
  if (slice.length) updates.keyPointsSummary = normalizeMonthForImportText(slice.join("\n"));
}

const PPTX_IMPORT_FIELD_LABELS = {
  customerName: "Customer name",
  executiveSummary: "Executive summary (Engagement)",
  dateRange: "Report date range",
  preparedBy: "Prepared by",
  submittedOn: "Submitted on",
  totPotIncNarrative: "Total Potential Incident — narrative",
  trendNote: "Trend — note",
  trendNarrative: "Trend — narrative",
  riskNarrative: "Risk slide — narrative",
  riskCsv: "Risk CSV",
  slaIncidentCount: "SLA — incident count",
  slaClosedIncidentCount: "SLA — closed incident count",
  tfPosNoteTrue: "True positive — note",
  tfPosNoteFalse: "False positive — note",
  epsMonthLabel: "EPS — month label",
  epsTotalValue: "EPS — total value",
  epsEventsCsv: "EPS events CSV",
  supportMajorCsv: "Support — major tickets CSV",
  supportMinorCsv: "Support — minor tickets CSV",
  inventoryCsv: "Inventory CSV",
  inventoryNote: "Inventory note",
  keyPointsSummary: "Key points summary"
};

const LAST_PPTX_IMPORT_STORAGE_KEY = "bluPineLastPptxImport";

function renderLastImportLog(fileName, appliedIds) {
  const meta = document.getElementById("lastImportLogMeta");
  const list = document.getElementById("lastImportLogList");
  const box = document.getElementById("lastImportLog");
  if (!meta || !list || !box) return;

  const when = new Date().toLocaleString();
  meta.textContent = `${when} — ${fileName || "presentation.pptx"} — ${appliedIds.length} field(s)`;
  list.innerHTML = "";
  appliedIds.forEach((id) => {
    const li = document.createElement("li");
    li.textContent = PPTX_IMPORT_FIELD_LABELS[id] || id;
    list.appendChild(li);
  });
  box.hidden = false;

  try {
    sessionStorage.setItem(
      LAST_PPTX_IMPORT_STORAGE_KEY,
      JSON.stringify({ t: Date.now(), file: fileName || "", ids: appliedIds })
    );
  } catch (e) {
    // Storage may be unavailable (private mode).
  }
}

function restoreLastImportLogFromStorage() {
  try {
    const raw = sessionStorage.getItem(LAST_PPTX_IMPORT_STORAGE_KEY);
    if (!raw) return;
    const data = JSON.parse(raw);
    if (!data || !Array.isArray(data.ids) || !data.ids.length) return;
    const meta = document.getElementById("lastImportLogMeta");
    const list = document.getElementById("lastImportLogList");
    const box = document.getElementById("lastImportLog");
    if (!meta || !list || !box) return;
    const when = data.t ? new Date(data.t).toLocaleString() : "";
    meta.textContent = `${when}${when ? " — " : ""}${data.file || "presentation.pptx"} — ${data.ids.length} field(s) (restored from this session)`;
    list.innerHTML = "";
    data.ids.forEach((id) => {
      const li = document.createElement("li");
      li.textContent = PPTX_IMPORT_FIELD_LABELS[id] || id;
      list.appendChild(li);
    });
    box.hidden = false;
  } catch (e) {
    // ignore
  }
}

function clearLastImportLog() {
  const box = document.getElementById("lastImportLog");
  if (box) box.hidden = true;
  try {
    sessionStorage.removeItem(LAST_PPTX_IMPORT_STORAGE_KEY);
  } catch (e) {
    // ignore
  }
}

async function importEditablePptxIntoForm(file) {
  if (typeof JSZip === "undefined") {
    alert("JSZip did not load. Check your network and refresh the page.");
    return;
  }
  const buf = await file.arrayBuffer();
  const zip = await JSZip.loadAsync(buf);
  const slidePaths = Object.keys(zip.files)
    .filter((k) => /^ppt\/slides\/slide\d+\.xml$/i.test(k))
    .sort((a, b) => {
      const na = parseInt(a.match(/slide(\d+)/i)[1], 10);
      const nb = parseInt(b.match(/slide(\d+)/i)[1], 10);
      return na - nb;
    });

  if (!slidePaths.length) {
    alert("No slides found in this file.");
    return;
  }

  const updates = {};
  const riskAccum = { narrativeDone: false, csvParts: [] };

  for (const path of slidePaths) {
    const entry = zip.file(path);
    if (!entry) continue;
    const xml = await entry.async("string");
    const paragraphs = extractParagraphsFromSlideXml(xml);
    const tables = extractTablesFromSlideXml(xml);
    const kind = classifySlideKind(paragraphs);

    switch (kind) {
      case "revision":
        importFromRevisionParagraphs(paragraphs, updates);
        break;
      case "engagement":
        importFromEngagementParagraphs(paragraphs, updates);
        break;
      case "totPot":
        importFromTotPotParagraphs(paragraphs, updates);
        break;
      case "trend":
        importFromTrendParagraphs(paragraphs, updates);
        break;
      case "risk":
        importFromRiskSlide(paragraphs, tables, updates, riskAccum);
        break;
      case "sla":
        importFromSlaParagraphs(paragraphs, updates);
        break;
      case "tfPos":
        importTfNotesFromParagraphs(paragraphs, updates);
        break;
      case "epsTrend":
        importFromEpsTrend(paragraphs, tables, updates);
        break;
      case "epsEvents":
        if (tables[0]) updates.epsEventsCsv = importEpsEventsTableToCsv(tables[0]);
        break;
      case "support":
        importSupportFromTables(tables, updates);
        break;
      case "inventory":
        importInventoryFromSlide(paragraphs, tables, updates);
        break;
      case "keyPoints":
        keyPointsFromParagraphs(paragraphs, updates);
        break;
      default:
        break;
    }
  }

  if (riskAccum.csvParts.length) {
    updates.riskCsv = riskAccum.csvParts.join("\n");
  }

  const applied = [];
  Object.entries(updates).forEach(([id, val]) => {
    const el = document.getElementById(id);
    if (el && typeof val === "string") {
      el.value = val;
      applied.push(id);
    }
  });

  if (!applied.length) {
    alert(
      "No editable text fields were found. Use an PPTX exported with “Export PPTX (editable)” from this app, or edit text in PowerPoint (not only pictures)."
    );
    return;
  }

  applyData();
  renderLastImportLog(file && file.name ? file.name : "presentation.pptx", applied);
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

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(String(e.target?.result || ""));
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

function imageUrlToDataUrlViaCanvas(src) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      try {
        const c = document.createElement("canvas");
        c.width = img.naturalWidth;
        c.height = img.naturalHeight;
        const ctx = c.getContext("2d");
        if (!ctx) {
          resolve("");
          return;
        }
        ctx.drawImage(img, 0, 0);
        resolve(c.toDataURL("image/png"));
      } catch {
        resolve("");
      }
    };
    img.onerror = () => resolve("");
    img.src = src;
  });
}

async function ensureDefaultPuzzleDataUrl() {
  if (defaultPuzzleDataUrl) return defaultPuzzleDataUrl;
  const placeholderImg = document.querySelector("#puzzlePlaceholderLayout img");
  if (!placeholderImg) return "";
  const rawSrc = placeholderImg.getAttribute("src");
  if (!rawSrc) return "";
  const absUrl = new URL(rawSrc, window.location.href).toString();
  let dataUrl = "";
  try {
    const resp = await fetch(absUrl, { cache: "force-cache" });
    if (resp.ok) {
      const blob = await resp.blob();
      dataUrl = await blobToDataUrl(blob);
    }
  } catch (e) {
    console.warn("Puzzle fetch failed, trying canvas decode:", e);
  }
  if (!dataUrl) dataUrl = await imageUrlToDataUrlViaCanvas(absUrl);
  if (dataUrl) {
    defaultPuzzleDataUrl = dataUrl;
    placeholderImg.src = dataUrl;
  } else {
    console.warn("Unable to prepare default puzzle data URL for:", absUrl);
  }
  return defaultPuzzleDataUrl;
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
  if (!shouldSkipApplyForNode(note)) {
    note.innerHTML = formatRichTextHTML(getValue("inventoryNote"));
  }
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

    section.innerHTML = `
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
      text: "Potential Incident - Severity Categories",
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
  }
  const remediationSummary = document.getElementById("remediationSlaSummary");
  if (remediationSummary) {
    remediationSummary.textContent = `Remediation Time for ${closedIncidentCount} closed tickets status: ${remediationInfo.label}.`;
    remediationSummary.style.color = "#111";
  }
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
          borderWidth: 0,
          minBarLength: high > 0 ? 44 : 0
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
      text: `Potential Incident Summary - ${getReportMonth()}`,
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
  syncReportRootLayoutClass();
  const reportRoot = document.getElementById("reportRoot");
  if (reportRoot) void reportRoot.offsetHeight;

  Object.entries(fields).forEach(([key, inputId]) => {
    const value = getValue(inputId);
    document.querySelectorAll(`[data-field="${key}"]`).forEach((el) => {
      if (shouldSkipApplyForNode(el)) return;
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
  if (execEl && !shouldSkipApplyForNode(execEl)) {
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
    if (ph) {
      ph.style.display = "flex";
      const phImg = ph.querySelector("img");
      if (phImg && defaultPuzzleDataUrl) phImg.src = defaultPuzzleDataUrl;
    }
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
  const fHdr = document.getElementById("fortiSiemTableColHdr");
  if (fHdr) fHdr.textContent = monthLbl;
  const fH = getValue("fortiAlertHigh");
  const fM = getValue("fortiAlertMedium");
  const fL = getValue("fortiAlertLow");
  const thF = document.getElementById("fortiSiemTableHigh");
  if (thF) thF.textContent = fH !== "" ? fH : "0";
  const tmF = document.getElementById("fortiSiemTableMed");
  if (tmF) tmF.textContent = fM !== "" ? fM : "0";
  const tlF = document.getElementById("fortiSiemTableLow");
  if (tlF) tlF.textContent = fL !== "" ? fL : "0";

  const tfHdr = document.getElementById("tfPosTrueTableColHdr");
  if (tfHdr) tfHdr.textContent = monthLbl;
  const tfHdrF = document.getElementById("tfPosFalseTableColHdr");
  if (tfHdrF) tfHdrF.textContent = monthLbl;
  const tH = getValue("tfPosTrueHigh");
  const tM = getValue("tfPosTrueMedium");
  const tL = getValue("tfPosTrueLow");
  const elTtH = document.getElementById("tfPosTrueTableHigh");
  if (elTtH) elTtH.textContent = tH !== "" ? tH : "0";
  const elTtM = document.getElementById("tfPosTrueTableMed");
  if (elTtM) elTtM.textContent = tM !== "" ? tM : "0";
  const elTtL = document.getElementById("tfPosTrueTableLow");
  if (elTtL) elTtL.textContent = tL !== "" ? tL : "0";
  const fHt = getValue("tfPosFalseHigh");
  const fMt = getValue("tfPosFalseMedium");
  const fLt = getValue("tfPosFalseLow");
  const elFtH = document.getElementById("tfPosFalseTableHigh");
  if (elFtH) elFtH.textContent = fHt !== "" ? fHt : "0";
  const elFtM = document.getElementById("tfPosFalseTableMed");
  if (elFtM) elFtM.textContent = fMt !== "" ? fMt : "0";
  const elFtL = document.getElementById("tfPosFalseTableLow");
  if (elFtL) elFtL.textContent = fLt !== "" ? fLt : "0";
  renderTfPosNotePreview("tfPosNoteTruePreview", "tfPosNoteTrue");
  renderTfPosNotePreview("tfPosNoteFalsePreview", "tfPosNoteFalse");
  renderEpsEventsTable();
  renderSupportTicketsSlide();
  renderIntegratedInventorySlide();
  const keyPointsPreview = document.getElementById("keyPointsSummaryPreview");
  if (keyPointsPreview && !shouldSkipApplyForNode(keyPointsPreview)) {
    keyPointsPreview.innerHTML = formatRichTextHTML(getValue("keyPointsSummary"));
  }
  const epsMonth = getValue("epsMonthLabel") || monthLbl;
  const epsTotalVal = Math.max(0, Number(getValue("epsTotalValue")) || 0);
  const epsMonthEl = document.getElementById("epsTotalMonthLabelCell");
  if (epsMonthEl) epsMonthEl.textContent = epsMonth;
  const epsValEl = document.getElementById("epsTotalValueCell");
  if (epsValEl) epsValEl.textContent = String(epsTotalVal);

  const totNarr = document.getElementById("totPotIncNarrativePreview");
  if (totNarr && !shouldSkipApplyForNode(totNarr)) {
    totNarr.innerHTML = `<p>${formatRichTextHTML(getValue("totPotIncNarrative"))}</p>`;
  }
  const trendNote = document.getElementById("trendNotePreview");
  if (trendNote && !shouldSkipApplyForNode(trendNote)) {
    trendNote.innerHTML = formatRichTextHTML(getValue("trendNote"));
  }
  const trendNarr = document.getElementById("trendNarrativePreview");
  if (trendNarr && !shouldSkipApplyForNode(trendNarr)) {
    trendNarr.innerHTML = formatRichTextHTML(getValue("trendNarrative"));
  }

  renderRiskSlides();
  initRiskTableResizing();
  updateLogos();
  applyPremiumLayout();
  renderTotPotIncChart();
  renderFortiSiemAlertsChart();
  renderTfPosCharts();
  renderEpsTrendPlot();
  renderTrendChart();
  renderRuleSeverityChart();
  renderSlaStatus();
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      if (totPotIncChart) totPotIncChart.resize();
      if (fortiSiemAlertsChart) fortiSiemAlertsChart.resize();
      if (truePositiveAlertsChart) truePositiveAlertsChart.resize();
      if (falsePositiveAlertsChart) falsePositiveAlertsChart.resize();
      if (epsTotalChart) epsTotalChart.resize();
      if (epsTopHostsChart) epsTopHostsChart.resize();
      if (trendChart) trendChart.resize();
      if (ruleSeverityChart) ruleSeverityChart.resize();
      window.dispatchEvent(new Event("resize"));
    });
  });
}

function getHiResChartDataUrl(canvas) {
  if (!canvas) return "";
  const sourceW = canvas.width || Math.max(1, Math.round(canvas.clientWidth || 1));
  const sourceH = canvas.height || Math.max(1, Math.round(canvas.clientHeight || 1));
  const upscale = 2;
  const out = document.createElement("canvas");
  out.width = Math.max(1, Math.round(sourceW * upscale));
  out.height = Math.max(1, Math.round(sourceH * upscale));
  const outCtx = out.getContext("2d");
  if (!outCtx) return canvas.toDataURL("image/png");
  outCtx.imageSmoothingEnabled = true;
  outCtx.imageSmoothingQuality = "high";
  outCtx.drawImage(canvas, 0, 0, out.width, out.height);
  return out.toDataURL("image/png");
}

function engagementSlideOnClone(doc) {
  const engage = doc.getElementById("slideEngagement");
  if (!engage) return;
  const textBox = engage.querySelector(".engagement-text");
  if (textBox) {
    textBox.style.setProperty("background-color", "#e1f4f8", "important");
    textBox.style.setProperty("background-image", "none", "important");
    textBox.style.setProperty("border", "1px solid #c8e0ea", "important");
    textBox.style.setProperty("border-radius", "6px", "important");
    textBox.style.setProperty("box-sizing", "border-box", "important");
  }
  engage.querySelectorAll("strong, .fw-bold").forEach((node) => {
    node.style.setProperty("font-weight", "700", "important");
    node.style.setProperty("color", "#000000", "important");
  });
  const preview = engage.querySelector("#puzzleGraphicPreview");
  const previewSrc = preview && preview.getAttribute("src");
  if (preview && previewSrc && previewSrc.length > 8) {
    preview.style.setProperty("display", "block", "important");
    preview.style.setProperty("object-fit", "contain", "important");
  }
  engage.querySelectorAll(".puzzle-placeholder img").forEach((img) => {
    img.style.setProperty("display", "block", "important");
    img.style.setProperty("visibility", "visible", "important");
    img.style.setProperty("opacity", "1", "important");
    img.style.setProperty("object-fit", "contain", "important");
    img.style.setProperty("max-width", "100%", "important");
  });
}

async function waitForImagesInElement(root) {
  if (!root) return;
  const imgs = Array.from(root.querySelectorAll("img"));
  await Promise.all(
    imgs.map(
      (img) =>
        new Promise((resolve) => {
          if (img.complete && img.naturalWidth > 0) {
            resolve();
            return;
          }
          const done = () => resolve();
          img.addEventListener("load", done, { once: true });
          img.addEventListener("error", done, { once: true });
        })
    )
  );
}

async function captureElToPng(el, scale = 2) {
  if (!el) return "";
  if (typeof html2canvas !== "function") return "";
  try {
    const rect = el.getBoundingClientRect();
    const width = Math.max(
      1,
      Math.round(Math.max(el.scrollWidth || 0, rect.width, el.clientWidth || 0))
    );
    const height = Math.max(
      1,
      Math.round(Math.max(el.scrollHeight || 0, rect.height, el.clientHeight || 0))
    );
    const hasEngagement = Boolean(
      el.id === "slideEngagement" ||
        (el.classList && el.classList.contains("slide-engagement")) ||
        (el.querySelector && el.querySelector(".engagement-text"))
    );
    const capturePromise = html2canvas(el, {
      scale,
      backgroundColor: "#ffffff",
      useCORS: true,
      allowTaint: false,
      width,
      height,
      windowWidth: width,
      windowHeight: height,
      logging: false,
      onclone: hasEngagement ? (doc) => engagementSlideOnClone(doc) : undefined
    }).catch((err) => {
      console.warn("captureElToPng html2canvas failed:", err);
      return null;
    });

    const timeoutMs = 12000;
    const timeoutPromise = new Promise((resolve) =>
      setTimeout(() => resolve("__timeout__"), timeoutMs)
    );

    const result = await Promise.race([capturePromise, timeoutPromise]);
    if (!result || result === "__timeout__") return "";
    return result.toDataURL("image/png");
  } catch (e) {
    console.warn("captureElToPng failed:", e);
    return "";
  }
}

function canvasHasRenderedPixels(canvas) {
  try {
    if (!canvas || typeof canvas.getContext !== "function") return false;
    const w = canvas.width;
    const h = canvas.height;
    if (!w || !h) return false;
    if (w < 5 || h < 5) return false;
    const ctx = canvas.getContext("2d");
    if (!ctx || typeof ctx.getImageData !== "function") return false;
    const x = Math.floor(w / 2);
    const y = Math.floor(h / 2);
    const d = ctx.getImageData(x, y, 1, 1).data;
    // If the canvas hasn't been drawn, alpha is often 0 (transparent).
    return d && d.length >= 4 && d[3] > 0;
  } catch (e) {
    return true; // If we can't inspect, don't block export.
  }
}

async function waitForPageCanvases(page, timeoutMs = 4000) {
  const canvases = Array.from(page.querySelectorAll("canvas"));
  if (!canvases.length) return;

  const start = performance.now();
  while (performance.now() - start < timeoutMs) {
    const ok = canvases.every((c) => canvasHasRenderedPixels(c));
    if (ok) return;
    await new Promise((r) => setTimeout(r, 200));
  }
}

async function exportPptxFromDom() {
  await ensureDefaultPuzzleDataUrl();
  applyData();
  document.body.classList.add("pptx-export-capture");
  try {
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    if (totPotIncChart) totPotIncChart.resize();
    if (fortiSiemAlertsChart) fortiSiemAlertsChart.resize();
    if (truePositiveAlertsChart) truePositiveAlertsChart.resize();
    if (falsePositiveAlertsChart) falsePositiveAlertsChart.resize();
    if (epsTotalChart) epsTotalChart.resize();
    if (epsTopHostsChart) epsTopHostsChart.resize();
    if (trendChart) trendChart.resize();
    if (ruleSeverityChart) ruleSeverityChart.resize();
    if (responseSlaChart) responseSlaChart.resize();
    if (remediationSlaChart) remediationSlaChart.resize();
    window.dispatchEvent(new Event("resize"));
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    await new Promise((r) => setTimeout(r, 750));
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    if (document.fonts && document.fonts.ready) {
      try {
        await document.fonts.ready;
      } catch (e) {
        // Ignore font readiness errors; continue export.
      }
    }

    const layoutKey = getPptLayoutKey();
    const { w: slideW, h: slideH } = getPptSlideInches(layoutKey);
    const pptx = new PptxGenJS();
    const layoutCfg = pptLayoutConfig[layoutKey] || pptLayoutConfig.LAYOUT_16X9;

    if (layoutKey === "A4") {
      pptx.defineLayout({ name: "CUSTOM_A4_PORTRAIT", width: PPT_SLIDE_IN.A4.w, height: PPT_SLIDE_IN.A4.h });
      pptx.layout = "CUSTOM_A4_PORTRAIT";
    } else {
      pptx.layout = layoutCfg.pptxLayout;
    }

    pptx.title = "BluPine Monthly SOC Report";

    const addEngagementFallbackSlide = () => {
      addEngagementNativeSlide(pptx, slideW, slideH);
    };

    const pages = Array.from(document.querySelectorAll("#reportRoot .page"));
    for (const page of pages) {
      // Ensure the element is rendered with correct layout before html2canvas capture.
      page.scrollIntoView({ block: "center", inline: "nearest" });
      await new Promise((r) => setTimeout(r, 400));
      await waitForPageCanvases(page, 5000);
      await waitForImagesInElement(page);
      let img = await captureElToPng(page, 2);
      if (!img) img = await captureElToPng(page, 1.4);
      if (!img) {
        page.scrollIntoView({ block: "center", inline: "nearest" });
        await new Promise((r) => setTimeout(r, 400));
        img = await captureElToPng(page, 1);
      }
      if (!img) {
        console.warn("Page capture still blank, using low-scale fallback for:", page.id || page.className);
        img = await captureElToPng(page, 0.8);
      }
      if (img) {
        const slide = pptx.addSlide();
        slide.addImage({ data: img, x: 0, y: 0, w: slideW, h: slideH });
      } else if (page.id === "slideEngagement") {
        console.warn("Using engagement fallback slide generation.");
        addEngagementFallbackSlide();
      } else {
        const slide = pptx.addSlide();
        slide.addText("Slide capture failed for this page. Please re-export after refresh.", {
          x: 0.6,
          y: 2.6,
          w: slideW - 1.2,
          h: 0.5,
          fontSize: 16,
          bold: true,
          color: "AA0000",
          align: "center"
        });
      }
    }

    await pptx.writeFile({ fileName: "BluPine_Monthly_SOC_Report.pptx", compression: true });
  } finally {
    document.body.classList.remove("pptx-export-capture");
  }
}

async function exportPptx() {
  return exportPptxFromDom();
}

/** Native PptxGenJS slides (editable text); layout differs from the web preview. */
async function exportPptxEditableNative() {
  return exportPptxLegacy();
}

async function exportPptxLegacy() {
  await ensureDefaultPuzzleDataUrl();
  applyData();
  await new Promise((r) => setTimeout(r, 400));

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
  pptx.title = `BluPine Monthly SOC Report`;

  const margin = 0.5 * sx;
  const contentW = slideW - 2 * margin;
  const logoW = 1.2;
  const logoX = slideW - logoW - 0.15 * sx;

  const titleStyle = { fontFace: "Calibri" };
  const barChartType = pptx.ChartType ? pptx.ChartType.bar : "bar";

  let slide = pptx.addSlide();
  slide.addText("BluPine Monthly SOC Report", {
    x: margin,
    y: 1.1 * sy,
    w: contentW,
    fontSize: 24,
    bold: true,
    color: "12284A",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }

  slide = pptx.addSlide();
  slide.addText("Document Revision History", {
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
    [
      `Version 1.0`,
      `Dated ${getValue("dateRange")}`,
      `Prepared By ${getValue("preparedBy")}`,
      `Reviewed By ${STATIC_REVIEWED_BY}`,
      `Approved By ${STATIC_APPROVED_BY}`,
      `Submitted On ${getValue("submittedOn")}`
    ].join("\n"),
    { x: margin, y: 1.05 * sy, w: contentW * 0.92, fontSize: 11, color: "000000", ...titleStyle }
  );
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }

  addEngagementNativeSlide(pptx, slideW, slideH);

  const monthLabel = getChartMonthLabel();
  const highVal = Math.max(0, parseFloat(getValue("totPotIncHigh")) || 0);
  const medVal = Math.max(0, parseFloat(getValue("totPotIncMedium")) || 0);
  const fortiHigh = Math.max(0, parseFloat(getValue("fortiAlertHigh")) || 0);
  const fortiMed = Math.max(0, parseFloat(getValue("fortiAlertMedium")) || 0);
  const fortiLow = Math.max(0, parseFloat(getValue("fortiAlertLow")) || 0);
  const fortiPeak = Math.max(fortiHigh, fortiMed, fortiLow, 1);
  const fortiYMax = Math.max(2000, Math.ceil((fortiPeak * 1.1) / 2000) * 2000);

  slide = pptx.addSlide();
  slide.addText("Total Potential Incident", {
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
  const totBandEl = document.getElementById("totPotBandedBlock");
  const totBandImg = await captureElToPng(totBandEl, 2);
  if (totBandImg) {
    try {
      slide.addImage({
        data: totBandImg,
        x: 0.7 * sx,
        y: 0.95 * sy,
        w: 8.6 * sx,
        h: 2.55 * sy
      });
    } catch (e) {
      console.warn("totPotBandImg addImage failed, fallback:", e);
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
          h: 2.55 * sy,
          barDir: "col",
          barGrouping: "clustered",
          chartColors: ["FF0000", "FFFF00"],
          showLegend: true,
          legendPos: "b",
          valAxisMaxVal: 50,
          valAxisMinVal: 0,
          valAxisMajorUnit: 5,
          showDataTable: true,
          dataTableFontSize: 9,
          catAxisLabelFontSize: 10,
          valAxisLabelFontSize: 10
        }
      );
    }
  } else {
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
        h: 2.55 * sy,
        barDir: "col",
        barGrouping: "clustered",
        chartColors: ["FF0000", "FFFF00"],
        showLegend: true,
        legendPos: "b",
        valAxisMaxVal: 50,
        valAxisMinVal: 0,
        valAxisMajorUnit: 5,
        showDataTable: true,
        dataTableFontSize: 9,
        catAxisLabelFontSize: 10,
        valAxisLabelFontSize: 10
      }
    );
  }

  slide.addText(parseRichTextPptx(getValue("totPotIncNarrative")), {
    x: margin,
    y: 3.65 * sy,
    w: contentW,
    h: 1.45 * sy,
    fontSize: 11,
    color: "000000",
    valign: "top",
    fontFace: "Calibri"
  });

  slide = pptx.addSlide();
  slide.addText("Potential Incident Tickets Trend", {
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
  const trendCanvas = document.getElementById("trendChart");
  if (trendCanvas) {
    slide.addImage({
      data: getHiResChartDataUrl(trendCanvas, trendChart),
      x: 0.35 * sx,
      y: 0.9 * sy,
      w: 9.3 * sx,
      h: 3.0 * sy
    });
  }
  slide.addText(parseRichTextPptx(getValue("trendNote")), {
    x: margin,
    y: 4.0 * sy,
    w: contentW,
    h: 0.35 * sy,
    fontSize: 13,
    bold: true,
    color: "000000",
    fontFace: "Calibri"
  });
  slide.addText(parseRichTextPptx(getValue("trendNarrative")), {
    x: margin,
    y: 4.35 * sy,
    w: contentW,
    h: 0.55 * sy,
    fontSize: 11,
    color: "000000",
    fontFace: "Calibri"
  });
  slide.addText("Incident details are provided below.", {
    x: margin,
    y: 4.95 * sy,
    w: contentW,
    h: 0.3 * sy,
    fontSize: 11,
    color: "000000",
    fontFace: "Calibri"
  });

  const riskRows = parseRiskCsvRows(getValue("riskCsv"));
  const riskChunks = riskRows.length
    ? Array.from({ length: Math.ceil(riskRows.length / 3) }, (_, i) =>
        riskRows.slice(i * 3, i * 3 + 3)
      )
    : [[]];

  riskChunks.forEach((chunk, i) => {
    slide = pptx.addSlide();
    slide.addText(i === 0 ? "Potential Alerts - Risks Mitigated" : "Potential Alerts - Risks Mitigated (Contd.)", {
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

    let tableY = 1.0 * sy;
    if (i === 0) {
      slide.addText(parseRichTextPptx(getValue("riskNarrative")), {
        x: margin,
        y: 0.95 * sy,
        w: contentW,
        h: 0.85 * sy,
        fontSize: 10.5,
        color: "000000",
        valign: "top",
        fontFace: "Calibri"
      });
      tableY = 1.85 * sy;
    }

    const riskTableData = [
      [
        { text: "S.No", options: { bold: true, align: "center", fill: { color: "D7D7D7" } } },
        { text: "Attack Type", options: { bold: true, align: "center", fill: { color: "D7D7D7" } } },
        { text: "Risk Scenario", options: { bold: true, align: "center", fill: { color: "D7D7D7" } } },
        { text: "Type of Risk(s)\nCIA Triad", options: { bold: true, align: "center", fill: { color: "D7D7D7" } } },
        { text: "Potential Business Impact(s)", options: { bold: true, align: "center", fill: { color: "D7D7D7" } } },
        { text: "Risk Rating", options: { bold: true, align: "center", fill: { color: "D7D7D7" } } }
      ]
    ];
    if (chunk.length) {
      chunk.forEach((r) => {
        riskTableData.push([
          { text: String(r.sno), options: { align: "center" } },
          { text: r.attackType, options: { align: "left" } },
          { text: r.riskScenario, options: { align: "left" } },
          { text: r.ciaTriad, options: { align: "center" } },
          { text: r.businessImpact, options: { align: "left" } },
          { text: r.riskRating, options: { align: "center" } }
        ]);
      });
    } else {
      riskTableData.push([
        { text: "-", options: { align: "center" } },
        { text: "No risk rows found. Upload Risk CSV.", options: { align: "left" } },
        { text: "", options: { align: "left" } },
        { text: "", options: { align: "left" } },
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

  slide = pptx.addSlide();
  slide.addText("Rule-Based Severity Categories For Potential Incidents", {
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
  const ruleSeverityCanvas = document.getElementById("ruleSeverityChart");
  if (ruleSeverityCanvas) {
    slide.addImage({
      data: getHiResChartDataUrl(ruleSeverityCanvas, ruleSeverityChart),
      x: 0.35 * sx,
      y: 0.9 * sy,
      w: 9.3 * sx,
      h: 4.15 * sy
    });
  }

  const responseInfo = classifySlaScore(getValue("responseSlaPct"));
  const remediationInfo = classifySlaScore(getValue("remediationSlaPct"));
  const incidentCount = Math.max(0, parseInt(getValue("slaIncidentCount"), 10) || 0);
  const closedIncidentCount = Math.max(0, parseInt(getValue("slaClosedIncidentCount"), 10) || 0);
  const customer = getValue("customerName") || "customer";

  slide = pptx.addSlide();
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  slide.addText("Response Time SLA", {
    x: 0.5 * sx,
    y: 0.6 * sy,
    w: 2.8 * sx,
    h: 0.4 * sy,
    fontSize: 22,
    bold: true,
    color: "000000",
    fontFace: "Calibri"
  });
  slide.addText("Remediation Time SLA", {
    x: 0.5 * sx,
    y: 3.0 * sy,
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
      data: getHiResChartDataUrl(responseCanvas, responseSlaChart),
      x: 0.35 * sx,
      y: 1.0 * sy,
      w: 2.9 * sx,
      h: 1.8 * sy
    });
  }
  const remediationCanvas = document.getElementById("remediationSlaChart");
  if (remediationCanvas) {
    slide.addImage({
      data: getHiResChartDataUrl(remediationCanvas, remediationSlaChart),
      x: 0.35 * sx,
      y: 3.45 * sy,
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
      "High Severity ticket resulted in extremely serious interruptions to Business system. It has affected the user community",
      "S-1",
      "≤15 Minutes",
      "4 Hours"
    ],
    [
      "Medium Severity ticket resulted in interruptions normal operations. It does not prevent Business Operations or minor ticket",
      "S-2",
      "≤15 Minutes",
      "8 Hours"
    ],
    [
      "Low Severity A Request/Query/Service that does not change existing service structure and no cost implication",
      "S-3",
      "≤30 Minutes",
      "24 Hours"
    ]
  ];
  slide.addTable(responseTbl, {
    x: 3.45 * sx,
    y: 0.75 * sy,
    w: 6.2 * sx,
    fontSize: 8.5,
    colW: [3.8 * sx, 1.1 * sx, 1.3 * sx, 1.35 * sx],
    border: { pt: 0.6, color: "C8D8DD" }
  });
  slide.addText(`Overall, ${incidentCount} incidents were reported to ${customer}, response status: ${responseInfo.label}.`, {
    x: 3.45 * sx,
    y: 2.95 * sy,
    w: 6.2 * sx,
    h: 0.35 * sy,
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
      "High Severity ticket resulted in extremely serious interruptions to Business system. It has affected the user community",
      "High",
      "4 Hours"
    ],
    [
      "Medium Severity ticket resulted in interruptions normal operations. It does not prevent Business Operations or minor ticket",
      "Medium",
      "8 Hours"
    ],
    [
      "Low A Request/Query/Service that does not change existing service structure and no cost implication",
      "Low",
      "24 Hours"
    ]
  ];
  slide.addTable(remediationTbl, {
    x: 3.45 * sx,
    y: 3.2 * sy,
    w: 6.2 * sx,
    fontSize: 8.5,
    colW: [4.55 * sx, 1.0 * sx, 1.45 * sx],
    border: { pt: 0.6, color: "C8D8DD" }
  });
  slide.addText(`Remediation Time for ${closedIncidentCount} closed tickets status: ${remediationInfo.label}.`, {
    x: 3.45 * sx,
    y: 5.15 * sy,
    w: 6.2 * sx,
    h: 0.35 * sy,
    fontSize: 10,
    bold: true,
    color: "111111",
    fontFace: "Calibri"
  });

  slide = pptx.addSlide();
  slide.addText("Total Alerts Triggered in FortiSIEM", {
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
  const fortiBandEl = document.getElementById("fortiSiemBandedBlock");
  const fortiBandImg = await captureElToPng(fortiBandEl, 2);
  if (fortiBandImg) {
    try {
      slide.addImage({
        data: fortiBandImg,
        x: 0.7 * sx,
        y: 0.95 * sy,
        w: 8.6 * sx,
        h: 2.55 * sy
      });
    } catch (e) {
      console.warn("fortiBandImg addImage failed, fallback:", e);
      slide.addChart(
        barChartType,
        [
          { name: "High", labels: [monthLabel], values: [fortiHigh] },
          { name: "Medium", labels: [monthLabel], values: [fortiMed] },
          { name: "Low", labels: [monthLabel], values: [fortiLow] }
        ],
        {
          x: 0.7 * sx,
          y: 0.95 * sy,
          w: 8.6 * sx,
          h: 2.55 * sy,
          barDir: "col",
          barGrouping: "clustered",
          chartColors: ["FF0000", "FFFF00", "00B050"],
          showLegend: true,
          legendPos: "b",
          valAxisMaxVal: fortiYMax,
          valAxisMinVal: 0,
          valAxisMajorUnit: 2000,
          showDataTable: true,
          dataTableFontSize: 9,
          catAxisLabelFontSize: 10,
          valAxisLabelFontSize: 10
        }
      );
    }
  } else {
    slide.addChart(
      barChartType,
      [
        { name: "High", labels: [monthLabel], values: [fortiHigh] },
        { name: "Medium", labels: [monthLabel], values: [fortiMed] },
        { name: "Low", labels: [monthLabel], values: [fortiLow] }
      ],
      {
        x: 0.7 * sx,
        y: 0.95 * sy,
        w: 8.6 * sx,
        h: 2.55 * sy,
        barDir: "col",
        barGrouping: "clustered",
        chartColors: ["FF0000", "FFFF00", "00B050"],
        showLegend: true,
        legendPos: "b",
        valAxisMaxVal: fortiYMax,
        valAxisMinVal: 0,
        valAxisMajorUnit: 2000,
        showDataTable: true,
        dataTableFontSize: 9,
        catAxisLabelFontSize: 10,
        valAxisLabelFontSize: 10
      }
    );
  }

  slide = pptx.addSlide();
  slide.addText("Total Number Of True & False Positive Alerts", {
    x: margin,
    y: 0.4 * sy,
    w: contentW,
    h: 0.55 * sy,
    fontSize: 26,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  slide.addText("True Positive Alerts based on Severity", {
    x: 0.35 * sx,
    y: 0.88 * sy,
    w: 4.6 * sx,
    h: 0.35 * sy,
    fontSize: 12,
    bold: true,
    color: "000000",
    fontFace: "Calibri"
  });
  slide.addText("False Positive Alerts based on Severity", {
    x: 5.05 * sx,
    y: 0.88 * sy,
    w: 4.6 * sx,
    h: 0.35 * sy,
    fontSize: 12,
    bold: true,
    color: "000000",
    fontFace: "Calibri"
  });
  const tpCanvasEl = document.getElementById("truePositiveAlertsChart");
  const fpCanvasEl = document.getElementById("falsePositiveAlertsChart");
  if (tpCanvasEl) {
    slide.addImage({
      data: getHiResChartDataUrl(tpCanvasEl, truePositiveAlertsChart),
      x: 0.35 * sx,
      y: 1.15 * sy,
      w: 4.6 * sx,
      h: 2.65 * sy
    });
  }
  if (fpCanvasEl) {
    slide.addImage({
      data: getHiResChartDataUrl(fpCanvasEl, falsePositiveAlertsChart),
      x: 5.05 * sx,
      y: 1.15 * sy,
      w: 4.6 * sx,
      h: 2.65 * sy
    });
  }
  slide.addText(formatTfPosNoteForPptx("tfPosNoteTrue"), {
    x: 0.35 * sx,
    y: 3.95 * sy,
    w: 4.6 * sx,
    h: 1.85 * sy,
    fontSize: 9,
    color: "000000",
    valign: "top",
    fontFace: "Calibri"
  });
  slide.addText(formatTfPosNoteForPptx("tfPosNoteFalse"), {
    x: 5.05 * sx,
    y: 3.95 * sy,
    w: 4.6 * sx,
    h: 1.85 * sy,
    fontSize: 9,
    color: "000000",
    valign: "top",
    fontFace: "Calibri"
  });

  const epsMonthLabel = getValue("epsMonthLabel") || monthLabel;
  const epsTotalValue = Math.max(0, Number(getValue("epsTotalValue")) || 0);
  const epsRows = parseEpsTopHostsRows(getValue("epsTopHostsCsv"));
  const epsLegendRows = epsRows.length ? epsRows : [{ host: "N/A", eps: 0 }];

  slide = pptx.addSlide();
  slide.addText("EPS Trend Plot", {
    x: margin,
    y: 0.32 * sy,
    w: contentW,
    h: 0.5 * sy,
    fontSize: 30,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  const epsTotalCanvasEl = document.getElementById("epsTotalChart");
  if (epsTotalCanvasEl) {
    slide.addImage({
      data: getHiResChartDataUrl(epsTotalCanvasEl, epsTotalChart),
      x: 0.12 * sx,
      y: 0.82 * sy,
      w: 3.9 * sx,
      h: 3.95 * sy
    });
  }
  const epsTopCanvasEl = document.getElementById("epsTopHostsChart");
  if (epsTopCanvasEl) {
    slide.addImage({
      data: getHiResChartDataUrl(epsTopCanvasEl, epsTopHostsChart),
      x: 4.52 * sx,
      y: 0.82 * sy,
      w: 5.25 * sx,
      h: 2.45 * sy
    });
  }
  slide.addTable(
    [
      [{ text: epsMonthLabel, options: { align: "left" } }, { text: String(epsTotalValue), options: { align: "center", bold: true } }]
    ],
    {
      x: 0.12 * sx,
      y: 4.68 * sy,
      w: 3.9 * sx,
      fontSize: 11,
      colW: [1.9 * sx, 2.0 * sx],
      border: { pt: 0.6, color: "D0D0D0" }
    }
  );
  const legendTableRows = [];
  for (let i = 0; i < 5; i += 1) {
    const left = epsLegendRows[i];
    const right = epsLegendRows[i + 5];
    legendTableRows.push([
      left ? `${i + 1} ${left.host}` : "",
      right ? `${i + 6} ${right.host}` : ""
    ]);
  }
  slide.addTable(legendTableRows, {
    x: 4.52 * sx,
    y: 3.35 * sy,
    w: 5.25 * sx,
    h: 1.5 * sy,
    fontSize: 8.5,
    colW: [2.55 * sx, 2.55 * sx],
    border: { pt: 0.5, color: "D0D0D0" }
  });

  const epsEventsRows = parseEpsEventsRows(getValue("epsEventsCsv"));
  slide = pptx.addSlide();
  slide.addText("Highest EPS Consuming Events For Mentioned Firewalls", {
    x: margin,
    y: 0.42 * sy,
    w: contentW,
    h: 0.5 * sy,
    fontSize: 20,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  const epsEventsTableData = [
    [
      { text: "S.No", options: { bold: true, fill: { color: "8ED8DF" }, align: "center" } },
      { text: "Reporting Device", options: { bold: true, fill: { color: "8ED8DF" }, align: "center" } },
      { text: "Event Type", options: { bold: true, fill: { color: "8ED8DF" }, align: "center" } },
      { text: "Event Name", options: { bold: true, fill: { color: "8ED8DF" }, align: "center" } },
      { text: "Matched Events", options: { bold: true, fill: { color: "8ED8DF" }, align: "center" } }
    ]
  ];
  if (epsEventsRows.length) {
    epsEventsRows.slice(0, 12).forEach((r, idx) => {
      const n = Number(String(r.matchedEvents).replace(/,/g, ""));
      epsEventsTableData.push([
        String(r.sno || idx + 1),
        r.device,
        r.eventType,
        r.eventName,
        Number.isFinite(n) ? n.toLocaleString("en-US") : String(r.matchedEvents || "")
      ]);
    });
  } else {
    epsEventsTableData.push(["-", "No rows found. Upload EPS Events CSV.", "", "", ""]);
  }
  slide.addTable(epsEventsTableData, {
    x: 0.65 * sx,
    y: 1.05 * sy,
    w: 8.7 * sx,
    h: 4.2 * sy,
    fontSize: 8.6,
    colW: [0.65 * sx, 2.05 * sx, 1.95 * sx, 2.15 * sx, 1.9 * sx],
    border: { pt: 0.6, color: "6CC0C6" }
  });

  const majorRows = parseSupportMajorRows(getValue("supportMajorCsv"));
  const minorRows = parseSupportMinorRows(getValue("supportMinorCsv"));
  const majorDataRows = majorRows.length
    ? majorRows.map((r) => [r.by, String(r.closed), String(r.inProcess), String(r.closed + r.inProcess)])
    : [["-", "0", "0", "0"]];
  const majorClosedTotal = majorRows.reduce((a, r) => a + r.closed, 0);
  const majorInProcessTotal = majorRows.reduce((a, r) => a + r.inProcess, 0);
  majorDataRows.push(["Grand Total", String(majorClosedTotal), String(majorInProcessTotal), String(majorClosedTotal + majorInProcessTotal)]);

  const minorDataRows = minorRows.length ? minorRows.map((r) => [r.by, String(r.closed)]) : [["-", "0"]];
  const minorTotal = minorRows.reduce((a, r) => a + r.closed, 0);
  minorDataRows.push(["Grand Total", String(minorTotal)]);

  slide = pptx.addSlide();
  slide.addText("Overall Support Ticket Handled By SNS (Firewall Support)", {
    x: margin,
    y: 0.45 * sy,
    w: contentW,
    h: 0.5 * sy,
    fontSize: 23,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  slide.addTable(
    [
      [{ text: "Major tickets", options: { bold: true, align: "center", fill: { color: "082A6F" }, color: "FFFFFF" } }, "", "", ""],
      [
        { text: "Created by", options: { bold: true, align: "center" } },
        { text: "Status", options: { bold: true, align: "center" } },
        "",
        { text: "Grand Total", options: { bold: true, align: "center" } }
      ],
      [
        "",
        { text: "Closed", options: { bold: true, align: "center" } },
        { text: "In-process", options: { bold: true, align: "center" } },
        ""
      ],
      ...majorDataRows
    ],
    {
      x: 0.1 * sx,
      y: 1.2 * sy,
      w: 4.75 * sx,
      fontSize: 10,
      colW: [1.35 * sx, 1.1 * sx, 1.1 * sx, 1.2 * sx],
      border: { pt: 0.6, color: "888888" }
    }
  );
  slide.addTable(
    [
      [{ text: "Minor tickets", options: { bold: true, align: "center", fill: { color: "082A6F" }, color: "FFFFFF" } }, ""],
      [
        { text: "Created by", options: { bold: true, align: "center" } },
        { text: "Status", options: { bold: true, align: "center" } }
      ],
      [
        "",
        { text: "Closed", options: { bold: true, align: "center" } }
      ],
      ...minorDataRows
    ],
    {
      x: 5.15 * sx,
      y: 1.2 * sy,
      w: 4.75 * sx,
      fontSize: 10,
      colW: [2.35 * sx, 2.35 * sx],
      border: { pt: 0.6, color: "888888" }
    }
  );

  const inventoryRows = parseInventoryRows(getValue("inventoryCsv"));
  const inventoryTotal = inventoryRows.reduce((a, r) => a + r.count, 0);
  const inventoryTableData = [
    [
      { text: "S.NO", options: { bold: true, align: "center", fill: { color: "4CC0C5" }, color: "FFFFFF" } },
      { text: "DEVICE NAME", options: { bold: true, align: "center", fill: { color: "4CC0C5" }, color: "FFFFFF" } },
      { text: "COUNT", options: { bold: true, align: "center", fill: { color: "4CC0C5" }, color: "FFFFFF" } }
    ],
    ...(inventoryRows.length
      ? inventoryRows.map((r, i) => [String(i + 1), r.device, r.count.toLocaleString("en-US")])
      : [["-", "No rows found. Upload Inventory CSV.", ""]]),
    ["", "Total", inventoryTotal.toLocaleString("en-US")]
  ];
  slide = pptx.addSlide();
  slide.addText("Integrated Device Inventory", {
    x: margin,
    y: 0.35 * sy,
    w: contentW,
    h: 0.5 * sy,
    fontSize: 30,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  slide.addTable(inventoryTableData, {
    x: 1.4 * sx,
    y: 1.25 * sy,
    w: 7.2 * sx,
    fontSize: 10.5,
    colW: [1.15 * sx, 4.05 * sx, 2.0 * sx],
    border: { pt: 0.6, color: "E2EDF0" }
  });
  slide.addText(getValue("inventoryNote"), {
    x: 0.65 * sx,
    y: 4.9 * sy,
    w: 8.7 * sx,
    h: 0.55 * sy,
    fontSize: 11,
    color: "111111",
    fontFace: "Calibri"
  });

  const keyPointsText = (getValue("keyPointsSummary") || "")
    .replace(/\[MONTH\]/g, getReportMonth())
    .replace(/\*\*(.*?)\*\*/g, "$1")
    .replace(/\*(.*?)\*/g, "$1");
  slide = pptx.addSlide();
  slide.addText("Key Points - Overall Summary", {
    x: margin,
    y: 0.35 * sy,
    w: contentW,
    h: 0.5 * sy,
    fontSize: 28,
    bold: true,
    color: "000000",
    align: "center",
    ...titleStyle
  });
  if (snsLogoDataUrl) {
    slide.addImage({ data: snsLogoDataUrl, x: logoX, y: 0.35 * sy, w: logoW, h: 0.5 });
  }
  slide.addText(keyPointsText, {
    x: 0.65 * sx,
    y: 1.35 * sy,
    w: 8.7 * sx,
    h: 3.9 * sy,
    fontSize: 16,
    color: "111111",
    valign: "top",
    fontFace: "Calibri",
    breakLine: true
  });

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

  await pptx.writeFile({ fileName: "BluPine_Monthly_SOC_Report.pptx", compression: true });
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
  document.querySelectorAll(".panel input, .panel textarea, .panel select").forEach((el) => {
    el.addEventListener("input", schedulePanelApply);
    el.addEventListener("change", schedulePanelApply);
  });
  document.getElementById("printBtn").addEventListener("click", async () => {
    await ensureDefaultPuzzleDataUrl();
    applyData();
    await new Promise((r) => setTimeout(r, 800));
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
    window.print();
  });
  document.getElementById("pptxBtn").addEventListener("click", () => {
    exportPptx().catch((err) => console.error("PPTX export failed:", err));
  });
  const pptxScreenshotBtn = document.getElementById("pptxScreenshotBtn");
  if (pptxScreenshotBtn) {
    pptxScreenshotBtn.addEventListener("click", () => {
      exportPptxEditableNative().catch((err) => console.error("PPTX editable export failed:", err));
    });
  }
  const pptxImportInput = document.getElementById("pptxImportInput");
  const pptxImportBtn = document.getElementById("pptxImportBtn");
  if (pptxImportBtn && pptxImportInput) {
    pptxImportBtn.addEventListener("click", () => pptxImportInput.click());
    pptxImportInput.addEventListener("change", (e) => {
      const f = e.target.files && e.target.files[0];
      e.target.value = "";
      if (!f) return;
      importEditablePptxIntoForm(f).catch((err) => {
        console.error("PPTX import failed:", err);
        alert(`Import failed: ${err && err.message ? err.message : String(err)}`);
      });
    });
  }
  const lastImportLogClear = document.getElementById("lastImportLogClear");
  if (lastImportLogClear) {
    lastImportLogClear.addEventListener("click", clearLastImportLog);
  }

  const panelToggleBtn = document.getElementById("panelToggleBtn");
  const panelBackdrop = document.getElementById("panelBackdrop");
  const closePanel = () => document.body.classList.remove("panel-open");
  const openPanel = () => document.body.classList.add("panel-open");
  if (panelToggleBtn) {
    panelToggleBtn.addEventListener("click", () => {
      if (document.body.classList.contains("panel-open")) {
        closePanel();
      } else {
        openPanel();
      }
    });
  }
  if (panelBackdrop) {
    panelBackdrop.addEventListener("click", closePanel);
  }
  window.addEventListener("resize", () => {
    if (window.innerWidth > 1450) closePanel();
  });

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
  bindCsvInput("epsEventsCsvFile", "epsEventsCsv");
  bindCsvInput("supportMajorCsvFile", "supportMajorCsv");
  bindCsvInput("supportMinorCsvFile", "supportMinorCsv");
  bindCsvInput("inventoryCsvFile", "inventoryCsv");

  await ensureDefaultPuzzleDataUrl();
  applyData();
  initInlineEditableReport();
  restoreLastImportLogFromStorage();
}

window.onload = init;
