# Monthly SOC Report Automation

Standalone web app to generate monthly SOC reports with SNS-style branding, aligned with the Blupine, Nocil, and Apraava monthly report structures.

## How to use

1. Open `index.html` in a browser (double-click or serve locally).
2. Choose **Customer Template**: `Blupine`, `Nocil`, `Apraava`, or `Custom`.
3. Fill customer/report metadata (month, date range, prepared/reviewed/approved).
4. Upload **SNS** and **client** logos (PNG/JPG) for the cover.
5. Paste or edit CSV blocks (devices, risks, SLA, inventory, FortiSIEM alerts, TP/FP).
6. For **Apraava**, fill the alert narrative and EPS trend CSV; for **Blupine**, fill highest-EPS events and firewall support text when needed.
7. Click **Apply Data** to refresh the preview.
8. Export:
   - **Export PDF** — print dialog → Save as PDF (matches on-screen layout).
   - **Export PPTX** — downloads a PowerPoint with the same sections and chart images.
   - **Export DOCX** — downloads an editable Word document with tables and narrative.

## Exports

| Format | Purpose |
|--------|---------|
| PDF | Client-ready static report via browser print |
| PPTX | Editable slides; charts embedded from the preview |
| DOCX | Editable document; tables for risks, EPS, SLA, inventory |

## Customer-specific sections

- **Blupine**: includes highest EPS events table and firewall support block (matches the longer sample).
- **Nocil**: shorter flow; titles use “Total Potential Alerts” / “Potential Alert Tickets Trend”.
- **Apraava**: includes “Potential Incidents and Alert Summary” narrative and EPS trend chart.

Use **Custom** to show every optional input field in the panel.

## Exact visual match to your PDFs

Pixel-perfect parity with the original PDF decks depends on the same logo files, diagram assets, and fonts. Upload your official SNS/client logos in the panel; for full background graphics, add them as images in a future version or place them manually in PPT/DOCX after export.
