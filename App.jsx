import React, { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import { motion } from "framer-motion";
import {
  AlertCircle,
  CheckCircle2,
  Copy,
  FileSpreadsheet,
  Mail,
  RefreshCw,
  Upload,
  Wand2,
} from "lucide-react";

const TEMPLATE = {
  to: "Zachary Ellington",
  subject: "{clientName} / {propertyAddress} / {stateName} TY {taxYear} Appeal Recommendation",
  intro:
    "Please see our assessment reviews of this {clientName} site located in {jurisdiction}. To summarize, we are recommending {recommendation}. Please see the analysis below:",
  closing:
    "If the client has any information that might suggest a lower value, we'd be happy to revisit our analysis.\n\nPlease relay my recommendation to the client ahead of the {deadline}. If we don't receive a response by {responseDue}, we will proceed as indicated.",
};

const INITIAL_MANUAL_FIELDS = {
  toName: TEMPLATE.to,
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

function formatMoney(value, decimals = 0) {
  if (value === null || value === undefined || value === "") return "[missing]";
  const num = Number(value);
  if (!Number.isFinite(num)) return String(value);
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals,
  }).format(num);
}

function formatPercent(value, digits = 2) {
  if (value === null || value === undefined || value === "") return "[missing]";
  const num = Number(value);
  if (!Number.isFinite(num)) return String(value);
  return `${(num * 100).toFixed(digits)}%`;
}

function roundDownHundreds(value) {
  const num = Number(value);
  if (!Number.isFinite(num)) return value;
  return Math.floor(num / 100) * 100;
}

function average(values) {
  const nums = values.map(Number).filter(Number.isFinite);
  if (!nums.length) return null;
  return nums.reduce((sum, n) => sum + n, 0) / nums.length;
}

function findSheet(workbook, targetName) {
  const sheetName = workbook.SheetNames.find(
    (name) => name.trim().toLowerCase() === targetName.trim().toLowerCase()
  );
  return sheetName ? workbook.Sheets[sheetName] : null;
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
    `For our analysis, we performed a pro forma income approach utilizing the rent roll rental figures of ${formatMoney(data.annualRent)}, ${formatPercent(data.vacancyRate, 0)} vacancy & collection loss, and ${formatPercent(data.operatingExpenseRate, 0)} operating expenses before capitalizing the net income at a ${formatPercent(data.baseCapRate, 0)} base cap rate, loaded with the ${formatPercent(data.taxRate, 2)} tax rate. The resulting value indication was ${formatMoney(data.indicatedValue)} (${formatMoney(data.indicatedValuePerAcre, 1)} Per Acre), which ${Number(data.indicatedValue) > Number(data.currentAssessment) ? "exceeds" : "is below"} the current assessment.`,
    `Additionally, we conducted a sale analysis of similar properties in ${data.jurisdiction || "[missing]"} and found a price per acre range of ${formatMoney(data.salePricePerAcreMin)} to ${formatMoney(data.salePricePerAcreMax)} with average per acre sale price of ~${formatMoney(roundDownHundreds(data.salePricePerAcreAvg || 0))} relative to our indicated value.`,
  ].join("\n\n");
}

function buildEmail(data, manual) {
  const recommendation = data.recommendation || "[missing]";
  const subject = TEMPLATE.subject
    .replaceAll("{clientName}", data.clientName || "[missing]")
    .replaceAll("{propertyAddress}", data.propertyAddress || "[missing]")
    .replaceAll("{stateName}", data.stateCode || "[missing]")
    .replaceAll("{taxYear}", data.taxYear || "[missing]");

  const intro = TEMPLATE.intro
    .replaceAll("{clientName}", data.clientName || "[missing]")
    .replaceAll("{jurisdiction}", data.jurisdiction || "[missing]")
    .replaceAll("{recommendation}", recommendation);

  const closing = TEMPLATE.closing
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
    `Previous Assessment Value ${formatMoney(data.previousAssessment)} (${formatMoney(data.previousAssessmentPerAcre, 1)} Per Acre)`,
    `Current Assessment Value ${formatMoney(data.currentAssessment)} (${formatMoney(data.currentAssessmentPerAcre, 1)} Per Acre)`,
    `Ryan Indicated Value ${formatMoney(data.indicatedValue)} (${formatMoney(data.indicatedValuePerAcre, 1)} Per Acre)`,
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

  const comparableStats = extractComparableStats(sheetToRows(comparables));

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

  const previousAssessmentPerAcre = Number.isFinite(previousNum) && Number.isFinite(acreageNum) && acreageNum !== 0
    ? previousNum / acreageNum
    : null;
  const currentAssessmentPerAcre = Number.isFinite(currentNum) && Number.isFinite(acreageNum) && acreageNum !== 0
    ? currentNum / acreageNum
    : null;
  const indicatedValuePerAcre = Number.isFinite(indicatedNum) && Number.isFinite(acreageNum) && acreageNum !== 0
    ? indicatedNum / acreageNum
    : null;

  return {
    clientName,
    propertyAddress,
    parcelNumber,
    jurisdiction,
    stateCode,
    taxYear,
    acreage,
    previousAssessment,
    previousAssessmentPerAcre,
    currentAssessment,
    currentAssessmentPerAcre,
    indicatedValue,
    indicatedValuePerAcre,
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

function Badge({ children, secondary = false }) {
  return <span className={`badge ${secondary ? "badge-secondary" : ""}`}>{children}</span>;
}

function Panel({ title, description, icon, children, right }) {
  const Icon = icon;
  return (
    <section className="panel">
      <div className="panel-header">
        <div>
          <h2 className="panel-title">{Icon ? <Icon size={18} /> : null}{title}</h2>
          {description ? <p className="panel-description">{description}</p> : null}
        </div>
        {right}
      </div>
      <div className="panel-body">{children}</div>
    </section>
  );
}

export default function App() {
  const [fileName, setFileName] = useState("");
  const [status, setStatus] = useState("Upload a property review workbook to generate the email draft.");
  const [manualFields, setManualFields] = useState(INITIAL_MANUAL_FIELDS);
  const [reviewData, setReviewData] = useState({});
  const [warnings, setWarnings] = useState([]);

  const completion = useMemo(() => {
    const complete = FIELD_CONFIG.filter(([key]) => reviewData[key] !== null && reviewData[key] !== undefined && reviewData[key] !== "").length;
    return Math.round((complete / FIELD_CONFIG.length) * 100);
  }, [reviewData]);

  const emailPreview = useMemo(() => buildEmail(reviewData, manualFields), [reviewData, manualFields]);

  async function handleUpload(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    setFileName(file.name);
    setStatus(`Reading ${file.name}...`);

    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
      const extracted = extractReviewData(workbook);
      const missing = FIELD_CONFIG
        .filter(([key]) => extracted[key] === null || extracted[key] === undefined || extracted[key] === "")
        .map(([, label]) => `${label} could not be found automatically.`);

      setReviewData(extracted);
      setWarnings(missing);
      setStatus(missing.length ? `Parsed ${file.name} with ${missing.length} warning(s).` : `Parsed ${file.name} successfully.`);
    } catch (error) {
      console.error(error);
      setStatus("Could not read that workbook. Please upload a standard .xlsx property review file.");
      setWarnings(["Workbook parsing failed."]);
      setReviewData({});
    }
  }

  async function handleCopy() {
    try {
      await navigator.clipboard.writeText(emailPreview);
      setStatus("Email copied to clipboard.");
    } catch {
      setStatus("Clipboard copy failed. Please copy directly from the preview.");
    }
  }

  function handleReset() {
    setFileName("");
    setStatus("Upload a property review workbook to generate the email draft.");
    setManualFields(INITIAL_MANUAL_FIELDS);
    setReviewData({});
    setWarnings([]);
  }

  return (
    <div className="app-shell">
      <motion.div initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} className="hero">
        <div>
          <div className="badge-row">
            <Badge>Deployable Web App</Badge>
            <Badge secondary>Excel to Email</Badge>
          </div>
          <h1>Property Review Email Generator</h1>
          <p>
            Upload a workbook, verify the extracted values, and generate a standardized email that updates live as you edit.
          </p>
        </div>
        <div className="hero-actions">
          <button className="button button-secondary" onClick={handleReset}><RefreshCw size={16} />Reset</button>
          <button className="button" onClick={handleCopy}><Copy size={16} />Copy Email</button>
        </div>
      </motion.div>

      <div className="layout-grid">
        <div className="left-column">
          <Panel
            title="Upload Workbook"
            description="Pre-wired to the sample review structure you uploaded."
            icon={FileSpreadsheet}
          >
            <label className="upload-zone">
              <Upload size={30} />
              <div className="upload-title">Click to upload .xlsx</div>
              <div className="upload-subtitle">The preview updates immediately after parsing.</div>
              <input type="file" accept=".xlsx,.xls" onChange={handleUpload} />
            </label>

            <div className="info-box">
              <strong>Current file</strong>
              <span>{fileName || "No file uploaded yet"}</span>
            </div>

            <div className="progress-block">
              <div className="progress-row">
                <span>Extraction coverage</span>
                <span>{completion}%</span>
              </div>
              <div className="progress-track">
                <div className="progress-fill" style={{ width: `${completion}%` }} />
              </div>
            </div>

            <div className="status-box">
              {warnings.length ? <AlertCircle size={18} /> : <CheckCircle2 size={18} />}
              <div>
                <strong>Status</strong>
                <div>{status}</div>
              </div>
            </div>

            {warnings.length > 0 && (
              <div className="warning-box">
                <strong>Warnings</strong>
                {warnings.map((warning) => (
                  <div key={warning}>• {warning}</div>
                ))}
              </div>
            )}
          </Panel>

          <Panel
            title="Extracted Review Fields"
            description="Edit any value manually and the preview refreshes live."
            icon={Wand2}
          >
            <div className="form-grid">
              {FIELD_CONFIG.map(([key, label]) => (
                <label key={key} className="field">
                  <span>{label}</span>
                  <input
                    value={reviewData[key] ?? ""}
                    onChange={(event) => setReviewData((prev) => ({ ...prev, [key]: event.target.value }))}
                  />
                </label>
              ))}
            </div>
          </Panel>

          <Panel
            title="Email Controls"
            description="These stay editable for recipient, deadline, and sender details."
            icon={Mail}
          >
            <div className="form-grid">
              {Object.entries(manualFields).map(([key, value]) => (
                <label key={key} className="field">
                  <span>{key.replace(/([A-Z])/g, " $1").replace(/^./, (s) => s.toUpperCase())}</span>
                  <input
                    value={value}
                    onChange={(event) => setManualFields((prev) => ({ ...prev, [key]: event.target.value }))}
                  />
                </label>
              ))}
            </div>
          </Panel>
        </div>

        <div className="right-column">
          <Panel
            title="Live Email Preview"
            description="Built from the structure in your finished Outlook email template."
            icon={Mail}
          >
            <textarea className="email-preview" value={emailPreview} readOnly />
          </Panel>
        </div>
      </div>
    </div>
  );
}
