/**
 * Endeavor Class of '26 — Expression of Interest form → Google Sheets
 *
 * Setup:
 * 1. Open the spreadsheet:
 *    https://docs.google.com/spreadsheets/d/1-Se8DLwMzzVJZqTV5CV1aMz-4Pz-z28Dl5JvdXkAAWM/edit
 * 2. Extensions → Apps Script → paste this file.
 * 3. Run setupSheetHeaders() once to update row 1 headers (does not clear data).
 * 4. Deploy → New deployment → Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy the Web App URL into index.html → EOI_WEBHOOK_URL
 *
 * Optional: set DRIVE_FOLDER_ID to store uploaded PDFs in a specific Drive folder.
 */

const SPREADSHEET_ID = "1-Se8DLwMzzVJZqTV5CV1aMz-4Pz-z28Dl5JvdXkAAWM";
const SHEET_NAME = "Submissions";
const DRIVE_FOLDER_ID = "1rEiYu9YpsJw12N4RrbS7CEW5LTU9Asi5"; // e.g. folder ID for pitch decks; leave empty to skip file upload

const HEADERS = [
  "Submitted at",
  "Full name",
  "LinkedIn",
  "Email",
  "Company name",
  "Role",
  "Company website or LinkedIn",
  "Company description",
  "Full-time team size",
  "Last 12 months total revenue (USDm)",
  "Revenue growth last 12 months (%)",
  "Funding status",
  "Pitch deck URL",
  "Hardest problem",
  "Why Stanford Class of 2026",
  "Available full week 16–21 Aug 2026",
  "US visa status",
  "Source URL",
];

function setupSheetHeaders() {
  const sheet = getOrCreateSheet_();
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
}

function doGet() {
  return jsonResponse_({ ok: true, service: "Endeavor EOI webhook" });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Missing request body");
    }

    const data = JSON.parse(e.postData.contents);
    const row = buildRow_(data);
    const sheet = getOrCreateSheet_();
    sheet.appendRow(row);

    return jsonResponse_({ ok: true });
  } catch (error) {
    return jsonResponse_({ ok: false, error: String(error) });
  }
}

function buildRow_(data) {
  const pitchDeckUrl = savePitchDeck_(
    data.pitchDeck,
    data.fullName,
    data.companyName,
  );

  return [
    data.submittedAt || new Date().toISOString(),
    data.fullName || "",
    data.linkedIn || "",
    data.email || "",
    data.companyName || "",
    data.role || "",
    data.companyUrl || "",
    data.companyDescription || "",
    data.teamSize || "",
    data.revenueLast12MonthsUsd || "",
    data.revenueGrowthPct || "",
    data.fundingStatus || "",
    pitchDeckUrl,
    data.hardestProblem || "",
    data.whyStanford || "",
    data.availabilityWeek || "",
    data.visaStatus || "",
    data.source || "",
  ];
}

function savePitchDeck_(pitchDeck, fullName, companyName) {
  if (!pitchDeck || !pitchDeck.base64 || !DRIVE_FOLDER_ID) {
    return "";
  }

  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const safeName = String(companyName || fullName || "submission")
    .replace(/[^\w\s-]/g, "")
    .trim()
    .replace(/\s+/g, "-")
    .slice(0, 60);
  const fileName = `${safeName || "pitch-deck"}-${Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyyMMdd-HHmmss",
  )}.pdf`;

  const blob = Utilities.newBlob(
    Utilities.base64Decode(pitchDeck.base64),
    pitchDeck.mimeType || "application/pdf",
    pitchDeck.fileName || fileName,
  );

  const file = folder.createFile(blob);
  return file.getUrl();
}

function getOrCreateSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
  }

  return sheet;
}

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
