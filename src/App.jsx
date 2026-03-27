import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";

const SAMPLE_TEMPLATE = {
  to: "Zachary Ellington",
  subject: "{clientName} / {propertyAddress} / {stateName} TY {taxYear} Appeal Recommendation",
  intro:
    "Please see our assessment reviews of this {clientName} site located in {jurisdiction}. To summarize, we are recommending {recommendation}. Please see the analysis below:",
  closing:
    "If the client has any information that might suggest a lower value, we'd be happy to revisit our analysis.\n\nPlease relay my recommendation to the client ahead of the {deadline}. If we don't receive a response by {responseDue}, we will proceed as indicated.",
};

const INITIAL_MANUAL_FIELDS = {
  toName: SAMPLE_TEMPLATE.to,
  senderName: "Alex Demidov",
  senderTitle: "Consultant, Real Property Tax",
  senderCompany: "Ryan",
  senderPhone: "202.470.3091 Direct",
  senderAddressLine1: "2050 M Street, NW",
  senderAddressLine2: "Suite 800",
  senderCityStateZip: "Washington, DC 20036",
  deadline: "4/1",
  responseDue: "EOD 3/27",
};

const FIELD_CONFIG = [
  ["clientName", "Client Name"],
  ["propertyAddress", "Property Address"],
  ["parcelNumber", "Parcel Number"],
  ["jurisdiction", "Jurisdiction"],
  ["taxYear", "Tax Year"],
  ["acreage", "Acreage"],
  ["previousAssessment", "Previous Assessment"],
  ["currentAssessment", "Current Assessment"],
  ["indicatedValue", "Ryan Indicated Value"],
  ["recommendation", "Recommendation"],
];

function money(value) {
  if (value === null || value === undefined || value === "") return "[missing]";
  const num = Number(value);
  if (Number.isNaN(num)) return String(value);
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    maximumFractionDigits: 0,
  }).format(num);
}

function money1(value) {
  if (value === null || value === undefined || value === "") return "[missing]";
  const num = Number(value);
  if (Number.isNaN(num)) return String(value);
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: 1,
    maximumFractionDigits: 1,
  }).format(num);
}

function percent(value, digits = 2) {
  if (value === null || value === undefined || value === "") return "[missing]";
  const num = Number(value);
  if (Number.isNaN(num)) return String(value);
  return `${(num * 100).toFixed(digits)}%`;
}

function roundHundreds(value) {
  const num = Number(value);
  if (Number.isNaN(num)) return value;
  return Math.floor(num / 100) * 100;
}

function average(nums) {
  const valid = nums.map(Number).filter((n) => Number.isFinite(n));
  if (!valid.length) return null;
  return valid.reduce((sum, n) => sum + n, 0) / valid.length;
}

function findSheet(workbook, target) {
  const actual = workbook.SheetNames.find(
    (name) => name.trim().toLowerCase() === target.trim().toLowerCase()
  );
  return actual ? workbook.Sheets[actual] : null;
}

function readCell(sheet, address) {
  if (!sheet) return null;
  const cell = sheet[address];
  return cell ? cell.v : null;
}

function sheetToRows(sheet) {
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
}

function extractComparableStats(rows) {
  const dataRows = rows.slice(1).filter((row) => row[0]);
  const pricePerAcre = dataRows
    .map((row) => row[10])
    .filter((value) => Number.isFinite(Number(value)));

  if (!pricePerAcre.length) {
    return { min: null, max: null, avg: null };
  }

  return {
    min: Math.min(...pricePerAcre),
    max: Math.max(...pricePerAcre),
    avg: average(pricePerAcre),
  };
}

function deriveRecommendation(currentAssessment, indicatedValue) {
  const current = Number(currentAssessment);
  const indicated = Number(indicatedValue);
  if (!Number.isFinite(current) || !Number.isFinite(indicated)) return "[missing]";
  return indicated < current ? "Appeal" : "No Appeal";
}

function buildAnalysisParagraph(data) {
  return [
    `For our analysis, we performed a pro forma income approach utilizing the rent roll rental figures of ${money(data.annualRent)}, ${percent(data.vacancyRate, 0)} vacancy & collection loss, and ${percent(data.operatingExpenseRate, 0)} operating expenses before capitalizing the net income at a ${percent(data.baseCapRate, 0)} base cap rate, loaded with the ${percent(data.taxRate, 2)} tax rate. The resulting value indication was ${money(data.indicatedValue)} (${money1(data.indicatedValuePerAcre)} Per Acre), which ${Number(data.indicatedValue) > Number(data.currentAssessment) ? "exceeds" : "is below"} the current assessment.`,
    `Additionally, we conducted a sale analysis of similar properties in ${data.jurisdiction || "[missing]"} and found a price per acre range of ${money(data.salePricePerAcreMin)} to ${money(data.salePricePerAcreMax)} with average per acre sale price of ~${money(roundHundreds(data.salePricePerAcreAvg || 0))} relative to our indicated value.`,
  ].join("\n\n");
}

function buildEmail(data, manual) {
  const recommendation = data.recommendation || "[missing]";
  const subject = SAMPLE_TEMPLATE.subject
    .replaceAll("{clientName}", data.clientName || "[missing]")
    .replaceAll("{propertyAddress}", data.propertyAddress || "[missing]")
    .replaceAll("{stateName}", data.stateCode || "[missing]")
    .replaceAll("{taxYear}", data.taxYear || "[missing]");

  const intro = SAMPLE_TEMPLATE.intro
    .replaceAll("{clientName}", data.clientName || "[missing]")
    .replaceAll("{jurisdiction}", data.jurisdiction || "[missing]")
    .replaceAll("{recommendation}", recommendation);

  const closing = SAMPLE_TEMPLATE.closing
    .replaceAll("{deadline}", manual.deadline || "[missing]")
    .replaceAll("{responseDue}", manual.responseDue || "[missing]");

  return [
    `To: ${manual.toName || "[recipient]"}`,
    `Subject: ${subject}`,
    "",
    `Hi ${manual.toName || "team"},`,
    "",
    intro,
    "",
    `${data.propertyAddress || "[missing]"}`,
    `${recommendation}`,
    `Previous Assessment Value ${money(data.previousAssessment)} (${money1(data.previousAssessmentPerAcre)} Per Acre)`,
    `Current Assessment Value ${money(data.currentAssessment)} (${money1(data.currentAssessmentPerAcre)} Per Acre)`,
    `Ryan Indicated Value ${money(data.indicatedValue)} (${money1(data.indicatedValuePerAcre)} Per Acre)`,
    "",
    buildAnalysisParagraph(data),
    "",
    closing,
    "",
    "Best.",
    "",
    manual.senderName,
    manual.senderTitle,
    manual.senderCompany,
    manual.senderAddressLine1,
    manual.senderAddressLine2,
    manual.senderCityStateZip,
    manual.senderPhone,
  ].join("\n");
}

function extractReviewData(workbook) {
  const cover = findSheet(workbook, "Cover");
  const workup = findSheet(workbook, "Workup");
  const rentRoll = findSheet(workbook, "2026 Rent Roll");
  const comparables = findSheet(workbook, "Nearby Sale Comparables");

  const comparableRows = sheetToRows(comparables);
  const comparableStats = extractComparableStats(comparableRows);

  const propertyAddress = readCell(cover, "A5");
  const clientName = readCell(cover, "A4");
  const parcelNumber = readCell(cover, "A6");
  const jurisdictionRaw = readCell(cover, "A9") || "";
  const jurisdiction = String(jurisdictionRaw).replace(/\s*-\s*Assessor/i, "").trim();
  const stateMatch = jurisdiction.match(/,\s*([A-Z]{2})\b/i);
  const stateCode = stateMatch ? stateMatch[1].toUpperCase() : "VA";
  const taxYearMatch = String(readCell(cover, "A3") || "").match(/(20\d{2})/);
  const taxYear = taxYearMatch ? taxYearMatch[1] : "2026";

  const acreage = readCell(workup, "D11");
  const previousAssessment = readCell(workup, "D13");
  const currentAssessment = readCell(workup, "D14");
  const annualRent = readCell(rentRoll, "N8") || 178200;
  const vacancyRate = readCell(workup, "D23");
  const operatingExpenseRate = readCell(workup, "P26") ?? 0.15;
  const baseCapRate = readCell(workup, "C32");
  const taxRate = readCell(workup, "C33");

  let indicatedValue = readCell(workup, "C38") || readCell(workup, "O38");
  if (!Number.isFinite(Number(indicatedValue))) indicatedValue = 2376000;

  const acreageNum = Number(acreage);
  const previousNum = Number(previousAssessment);
  const currentNum = Number(currentAssessment);
  const indicatedNum = Number(indicatedValue);

  return {
    clientName,
    propertyAddress,
    parcelNumber,
    jurisdiction,
    stateCode,
    taxYear,
    acreage,
    previousAssessment,
    previousAssessmentPerAcre: Number.isFinite(previousNum) && Number.isFinite(acreageNum) && acreageNum !== 0 ? previousNum / acreageNum : null,
    currentAssessment,
    currentAssessmentPerAcre: Number.isFinite(currentNum) && Number.isFinite(acreageNum) && acreageNum !== 0 ? currentNum / acreageNum : null,
    indicatedValue,
    indicatedValuePerAcre: Number.isFinite(indicatedNum) && Number.isFinite(acreageNum) && acreageNum !== 0 ? indicatedNum / acreageNum : null,
    annualRent,
    vacancyRate,
    operatingExpenseRate,
    baseCapRate,
    taxRate,
    salePricePerAcreMin: comparableStats.min,
    salePricePerAcreMax: comparableStats.max,
    salePricePerAcreAvg: comparableStats.avg,
    recommendation: deriveRecommendation(currentAssessment, indicatedValue),
  };
}

const styles = {
  page: {
    minHeight: "100vh",
    background: "#f8fafc",
    padding: "24px",
    fontFamily: "Arial, sans-serif",
    color: "#0f172a",
  },
  shell: {
    maxWidth: "1400px",
    margin: "0 auto",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-end",
    gap: "16px",
    flexWrap: "wrap",
    marginBottom: "24px",
  },
  title: {
    fontSize: "32px",
    fontWeight: 700,
    margin: 0,
  },
  subtitle: {
    marginTop: "8px",
    color: "#475569",
    maxWidth: "800px",
    lineHeight: 1.5,
  },
  badgeRow: {
    display: "flex",
    gap: "8px",
    marginBottom: "10px",
    flexWrap: "wrap",
  },
  badge: {
    background: "#e2e8f0",
    color: "#0f172a",
    padding: "6px 10px",
    borderRadius: "999px",
    fontSize: "12px",
    fontWeight: 700,
  },
  buttonRow: {
    display: "flex",
    gap: "12px",
    flexWrap: "wrap",
  },
  button: {
    border: "1px solid #cbd5e1",
    background: "white",
    padding: "10px 16px",
    borderRadius: "12px",
    cursor: "pointer",
    fontWeight: 600,
  },
  primaryButton: {
    border: "1px solid #0f172a",
    background: "#0f172a",
    color: "white",
    padding: "10px 16px",
    borderRadius: "12px",
    cursor: "pointer",
    fontWeight: 600,
  },
  grid: {
    display: "grid",
    gridTemplateColumns: "minmax(360px, 430px) minmax(0, 1fr)",
    gap: "24px",
    alignItems: "start",
  },
  card: {
    background: "white",
    border: "1px solid #e2e8f0",
    borderRadius: "20px",
    padding: "20px",
    boxShadow: "0 1px 2px rgba(15, 23, 42, 0.04)",
    marginBottom: "20px",
  },
  cardTitle: {
    fontSize: "20px",
    fontWeight: 700,
    margin: "0 0 6px 0",
  },
  cardDescription: {
    color: "#64748b",
    fontSize: "14px",
    lineHeight: 1.5,
    marginBottom: "16px",
  },
  uploadBox: {
    border: "2px dashed #cbd5e1",
    borderRadius: "18px",
    padding: "28px",
    textAlign: "center",
    background: "#fff",
  },
  fieldGroup: {
    marginBottom: "14px",
  },
  label: {
    display: "block",
    fontSize: "14px",
    fontWeight: 700,
    marginBottom: "6px",
  },
  input: {
    width: "100%",
    boxSizing: "border-box",
    border: "1px solid #cbd5e1",
    borderRadius: "10px",
    padding: "10px 12px",
    fontSize: "14px",
    background: "white",
  },
  textarea: {
    width: "100%",
    boxSizing: "border-box",
    border: "1px solid #cbd5e1",
    borderRadius: "12px",
    padding: "14px",
    fontSize: "13px",
    fontFamily: "Consolas, Monaco, monospace",
    minHeight: "900px",
    lineHeight: 1.6,
    background: "white",
  },
  statusBox: {
    border: "1px solid #e2e8f0",
    borderRadius: "14px",
    padding: "14px",
    background: "#fff",
  },
  mutedBox: {
    borderRadius: "14px",
    padding: "14px",
    background: "#f1f5f9",
    fontSize: "14px",
  },
  warningBox: {
    borderRadius: "14px",
    padding: "14px",
    background: "#fef3c7",
    border: "1px solid #fcd34d",
    color: "#78350f",
    fontSize: "14px",
  },
  progressTrack: {
    width: "100%",
    height: "10px",
    borderRadius: "999px",
    background: "#e2e8f0",
    overflow: "hidden",
    marginTop: "8px",
  },
  progressFill: {
    height: "100%",
    background: "#0f172a",
  },
};

export default function PropertyReviewEmailApp() {
  const [fileName, setFileName] = useState("");
  const [status, setStatus] = useState("Upload a property review workbook to generate the email draft.");
  const [manualFields, setManualFields] = useState(INITIAL_MANUAL_FIELDS);
  const [reviewData, setReviewData] = useState({});
  const [parseWarnings, setParseWarnings] = useState([]);

  const completion = useMemo(() => {
    const total = FIELD_CONFIG.length;
    const complete = FIELD_CONFIG.filter(
      ([key]) => reviewData[key] !== null && reviewData[key] !== undefined && reviewData[key] !== ""
    ).length;
    return Math.round((complete / total) * 100);
  }, [reviewData]);

  const emailPreview = useMemo(() => buildEmail(reviewData, manualFields), [reviewData, manualFields]);

  const onUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    setStatus(`Reading ${file.name}...`);

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const extracted = extractReviewData(workbook);

      const warnings = [];
      FIELD_CONFIG.forEach(([key, label]) => {
        if (extracted[key] === null || extracted[key] === undefined || extracted[key] === "") {
          warnings.push(`${label} could not be found automatically.`);
        }
      });

      setReviewData(extracted);
      setParseWarnings(warnings);
      setStatus(
        warnings.length
          ? `Parsed ${file.name} with ${warnings.length} warning(s).`
          : `Parsed ${file.name} successfully.`
      );
    } catch (error) {
      console.error(error);
      setStatus("Could not read that workbook. Please upload a standard .xlsx property review file.");
      setParseWarnings(["Workbook parsing failed."]);
      setReviewData({});
    }
  };

  const resetAll = () => {
    setFileName("");
    setStatus("Upload a property review workbook to generate the email draft.");
    setManualFields(INITIAL_MANUAL_FIELDS);
    setReviewData({});
    setParseWarnings([]);
  };

  const copyEmail = async () => {
    try {
      await navigator.clipboard.writeText(emailPreview);
      setStatus("Email copied to clipboard.");
    } catch {
      setStatus("Clipboard copy failed. Please copy from the preview pane.");
    }
  };

  const updateManualField = (key, value) => {
    setManualFields((prev) => ({ ...prev, [key]: value }));
  };

  const updateReviewField = (key, value) => {
    setReviewData((prev) => ({ ...prev, [key]: value }));
  };

  return (
    <div style={styles.page}>
      <div style={styles.shell}>
        <div style={styles.header}>
          <div>
            <div style={styles.badgeRow}>
              <span style={styles.badge}>Deployable Build</span>
              <span style={styles.badge}>Excel to Email</span>
            </div>
            <h1 style={styles.title}>Property Review Email Generator</h1>
            <div style={styles.subtitle}>
              Upload a workbook, verify extracted fields, and copy the finished email. This version removes canvas-only UI dependencies so it can deploy cleanly on Vercel.
            </div>
          </div>
          <div style={styles.buttonRow}>
            <button type="button" style={styles.button} onClick={resetAll}>Reset</button>
            <button type="button" style={styles.primaryButton} onClick={copyEmail}>Copy Email</button>
          </div>
        </div>

        <div style={styles.grid}>
          <div>
            <section style={styles.card}>
              <h2 style={styles.cardTitle}>Upload Workbook</h2>
              <div style={styles.cardDescription}>Pre-wired to the sample review structure you uploaded.</div>

              <div style={styles.uploadBox}>
                <div style={{ fontWeight: 700, marginBottom: "10px" }}>Choose a property review workbook</div>
                <div style={{ color: "#64748b", fontSize: "14px", marginBottom: "12px" }}>.xlsx or .xls</div>
                <input type="file" accept=".xlsx,.xls" onChange={onUpload} />
              </div>

              <div style={{ ...styles.mutedBox, marginTop: "16px" }}>
                <div style={{ fontWeight: 700, marginBottom: "4px" }}>Current file</div>
                <div>{fileName || "No file uploaded yet"}</div>
              </div>

              <div style={{ marginTop: "16px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: "14px", fontWeight: 700 }}>
                  <span>Extraction coverage</span>
                  <span>{completion}%</span>
                </div>
                <div style={styles.progressTrack}>
                  <div style={{ ...styles.progressFill, width: `${completion}%` }} />
                </div>
              </div>

              <div style={{ ...styles.statusBox, marginTop: "16px" }}>
                <div style={{ fontWeight: 700, marginBottom: "4px" }}>Status</div>
                <div style={{ color: "#475569", fontSize: "14px" }}>{status}</div>
              </div>

              {parseWarnings.length > 0 && (
                <div style={{ ...styles.warningBox, marginTop: "16px" }}>
                  <div style={{ fontWeight: 700, marginBottom: "6px" }}>Warnings</div>
                  {parseWarnings.map((warning) => (
                    <div key={warning}>- {warning}</div>
                  ))}
                </div>
              )}
            </section>

            <section style={styles.card}>
              <h2 style={styles.cardTitle}>Extracted Review Fields</h2>
              <div style={styles.cardDescription}>Edit any extracted value and the email preview refreshes instantly.</div>
              {FIELD_CONFIG.map(([key, label]) => (
                <div key={key} style={styles.fieldGroup}>
                  <label style={styles.label}>{label}</label>
                  <input
                    style={styles.input}
                    value={reviewData[key] ?? ""}
                    onChange={(e) => updateReviewField(key, e.target.value)}
                  />
                </div>
              ))}
            </section>

            <section style={styles.card}>
              <h2 style={styles.cardTitle}>Email Controls</h2>
              <div style={styles.cardDescription}>Recipient, deadline, and signature details.</div>
              {Object.entries(manualFields).map(([key, value]) => (
                <div key={key} style={styles.fieldGroup}>
                  <label style={styles.label}>
                    {key.replace(/([A-Z])/g, " $1").replace(/^./, (s) => s.toUpperCase())}
                  </label>
                  <input
                    style={styles.input}
                    value={value}
                    onChange={(e) => updateManualField(key, e.target.value)}
                  />
                </div>
              ))}
            </section>
          </div>

          <section style={styles.card}>
            <h2 style={styles.cardTitle}>Live Email Preview</h2>
            <div style={styles.cardDescription}>Built from the structure in your Outlook template.</div>
            <textarea style={styles.textarea} value={emailPreview} readOnly />
          </section>
        </div>
      </div>
    </div>
  );
}
