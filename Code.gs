/**
 * IDEAL FASTENER — CLIENT VISIT REPORT — Backend
 * ------------------------------------------------
 * What this does:
 * 1. Receives the form data from your mobile web form
 * 2. Appends it as a new row directly into your existing "Meeting Discussions" tab
 *    (same columns as your original file — no format change)
 * 3. Marks "Y" + highlights the day-column in your existing "DSR" tab
 * 4. Sends you an email (Outlook) with your FULL workbook attached as a real .xlsx
 *    file — same tabs, same formatting as your original Excel file
 *
 * SETUP STEPS (one time only):
 * 1. Upload your original DSR.xls to Google Drive
 * 2. Right-click it > Open with > Google Sheets
 *    (this converts it to a Google Sheet, keeping all tabs & formatting)
 * 3. In that converted Sheet: Extensions > Apps Script
 * 4. Delete any starter code, paste this whole file in
 * 5. Update the CONFIG section below with real emails
 * 6. Deploy > New deployment > Type "Web app"
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 7. Copy the Web app URL — paste it into the HTML form file
 * 8. First deploy, Google will ask to authorize — allow it
 *    (this script needs Drive access too, since it exports the file as Excel —
 *    you'll see an extra permission screen for that, just approve it)
 */

// ============ CONFIG — EDIT THESE ============
const MY_EMAIL   = "towhid@idealfastener.com";   // <-- your Outlook/email address — this is the ONLY recipient
const SHEET_NAME = "Meeting Discussions";        // your existing tab — do not rename
const DSR_SHEET_NAME = "DSR";                    // your existing tab — do not rename
const DSR_FIRST_DAY_COLUMN = 6;                  // column F = Day 1 (A=Name,B=Category,C=Product,D=blank,E=Buyers)

// Status colors — matches your legend exactly
const STATUS_COLORS = {
  visit:   "#00B050",  // green  — On Visit
  office:  "#FFA500",  // orange — In Office
  leave:   "#7030A0",  // purple — On Leave
  holiday: "#FF0000"   // red    — Official Holiday
};
const STATUS_LABELS = {
  visit: "On Visit", office: "In Office", leave: "On Leave", holiday: "Official Holiday"
};
// ================================================

// Serves the list of existing customer names (for the form's autocomplete dropdown)
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DSR_SHEET_NAME);
  let names = [];
  if (sheet) {
    const dataStartRow = findDSRDataStartRow(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow >= dataStartRow) {
      const values = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, 1).getValues();
      names = values.map(r => (r[0] || "").toString().trim()).filter(n => n !== "");
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify(names))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const kind = data.reportType || "visit"; // "visit" or "status"

    if (kind === "visit") {
      const visits = data.visits || [data]; // backward-compatible: supports old single-visit payloads too
      visits.forEach(v => {
        logVisit(data.meetingDate, v);
        markCustomerVisit(data.meetingDate, v);
      });
      markDayColumn(data.meetingDate, "visit");
      sendReportEmail(data.meetingDate, visits);
    } else {
      // status-only day: office / leave / holiday — no customer, no meeting recap
      markDayColumn(data.meetingDate, data.status);
      sendStatusEmail(data);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function logVisit(meetingDate, v) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Tab "' + SHEET_NAME + '" not found — check the tab name matches exactly.');

  // Same column order as your original file: Meeting dt, Customer Name,
  // Concern Person/Designation, Brand Name, Meeting Recap, Remarks
  sheet.appendRow([
    meetingDate,
    v.customerName,
    v.concernPerson,
    v.brandName,
    v.meetingRecap,
    v.remarks
  ]);
}

function sendStatusEmail(data) {
  const label = STATUS_LABELS[data.status] || data.status;
  const subject = `Daily Status — ${label} — ${data.meetingDate}`;

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background: #4338CA; color: #fff; padding: 16px 20px; border-radius: 6px 6px 0 0;">
        <h2 style="margin:0; font-size: 18px;">Daily Status</h2>
      </div>
      <table style="width:100%; border-collapse: collapse; margin-top: 12px;">
        <tr><td style="padding:8px; font-weight:bold; width:140px; border-bottom:1px solid #eee;">Date</td><td style="padding:8px; border-bottom:1px solid #eee;">${escapeHtml(data.meetingDate)}</td></tr>
        <tr><td style="padding:8px; font-weight:bold;">Status</td><td style="padding:8px;">${escapeHtml(label)}</td></tr>
      </table>
      <p style="color:#888; font-size:13px; margin-top:14px;">Full report workbook (all tabs) attached as Excel.</p>
      <p style="color:#888; font-size:12px; margin-top:6px;">Submitted via mobile visit report tool · System designed & engineered by Towhid</p>
    </div>
  `;

  const excelBlob = getWorkbookAsExcelBlob();
  const options = { htmlBody: htmlBody, attachments: [excelBlob] };

  MailApp.sendEmail(MY_EMAIL, subject, "Please see attached report (also viewable in the email body).", options);
}

function sendReportEmail(meetingDate, visits) {
  const names = visits.map(v => v.customerName).join(", ");
  const subject = visits.length === 1
    ? `Client Visit Report — ${visits[0].customerName} — ${meetingDate}`
    : `Client Visit Report — ${visits.length} visits — ${meetingDate}`;

  const visitSections = visits.map((v, i) => `
    <div style="margin-bottom:${i < visits.length - 1 ? '20' : '0'}px; ${i < visits.length - 1 ? 'padding-bottom:16px; border-bottom:2px solid #E5E5EF;' : ''}">
      <div style="color:#4338CA; font-weight:bold; font-size:13px; text-transform:uppercase; letter-spacing:0.04em; margin-bottom:6px;">Visit ${i + 1} of ${visits.length}</div>
      <table style="width:100%; border-collapse: collapse;">
        <tr><td style="padding:6px 8px; font-weight:bold; width:180px; border-bottom:1px solid #eee;">Customer Name</td><td style="padding:6px 8px; border-bottom:1px solid #eee;">${escapeHtml(v.customerName)}</td></tr>
        <tr><td style="padding:6px 8px; font-weight:bold; border-bottom:1px solid #eee;">Concern Person / Designation</td><td style="padding:6px 8px; border-bottom:1px solid #eee;">${escapeHtml(v.concernPerson)}</td></tr>
        <tr><td style="padding:6px 8px; font-weight:bold; border-bottom:1px solid #eee;">Brand Name</td><td style="padding:6px 8px; border-bottom:1px solid #eee;">${escapeHtml(v.brandName)}</td></tr>
        <tr><td style="padding:6px 8px; font-weight:bold; vertical-align:top; border-bottom:1px solid #eee;">Meeting Recap</td><td style="padding:6px 8px; border-bottom:1px solid #eee; white-space: pre-wrap;">${escapeHtml(v.meetingRecap)}</td></tr>
        <tr><td style="padding:6px 8px; font-weight:bold; vertical-align:top;">Remarks</td><td style="padding:6px 8px; white-space: pre-wrap;">${escapeHtml(v.remarks || "-")}</td></tr>
      </table>
    </div>
  `).join("");

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background: #4338CA; color: #fff; padding: 16px 20px; border-radius: 6px 6px 0 0;">
        <h2 style="margin:0; font-size: 18px;">Client Visit Report — ${escapeHtml(meetingDate)}</h2>
        <p style="margin:4px 0 0; font-size:13px; opacity:0.9;">${visits.length} visit${visits.length > 1 ? "s" : ""}: ${escapeHtml(names)}</p>
      </div>
      <div style="padding-top:14px;">
        ${visitSections}
      </div>
      <p style="color:#888; font-size:13px; margin-top:16px;">Full report workbook (all tabs) attached as Excel.</p>
      <p style="color:#888; font-size:12px; margin-top:6px;">Submitted via mobile visit report tool · System designed & engineered by Towhid</p>
    </div>
  `;

  const excelBlob = getWorkbookAsExcelBlob();

  const options = {
    htmlBody: htmlBody,
    attachments: [excelBlob]
  };

  MailApp.sendEmail(MY_EMAIL, subject, "Please see attached report (also viewable in the email body).", options);
}

// Exports the ENTIRE spreadsheet (all tabs, all formatting) as a real .xlsx file
function getWorkbookAsExcelBlob() {
  SpreadsheetApp.flush(); // ensure all pending edits (colors, values) are committed before export
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx";
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { "Authorization": "Bearer " + token }
  });
  const dateStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  return response.getBlob().setName(ss.getName() + " - " + dateStamp + ".xlsx");
}

/**
 * Writes "Y" in the visited customer's row for this day.
 * Adds a new row automatically if it's a new customer.
 */
/**
 * Normalizes a name for matching: trims, collapses internal whitespace,
 * strips invisible characters, and lowercases — so "ABC  Ltd." , "abc ltd.",
 * and " ABC Ltd. " all match as the same customer.
 */
function normalizeName(str) {
  if (!str) return "";
  return str.toString()
    .replace(/[\u200B-\u200D\uFEFF]/g, "")  // strip zero-width/invisible chars
    .replace(/\s+/g, " ")                   // collapse multiple spaces/tabs
    .trim()
    .toLowerCase();
}

/**
 * Finds the row right after the "CUSTOMER NAME" header in column A.
 * Falls back to row 2 if no header text is found (keeps things working
 * even if your sheet's title/header rows shift slightly).
 */
function findDSRDataStartRow(sheet) {
  const scanRows = Math.min(sheet.getLastRow(), 10);
  if (scanRows < 1) return 2;
  const colA = sheet.getRange(1, 1, scanRows, 1).getValues();
  for (let i = 0; i < colA.length; i++) {
    if (normalizeName(colA[i][0]) === "customer name") {
      return i + 2; // data starts right after the header row
    }
  }
  return 2; // fallback default
}

/**
 * Writes "Y" in the visited customer's row for this day.
 * Matches EXISTING customers (case/whitespace-insensitive) so repeat
 * visits never create duplicate rows — only genuinely new customers
 * get a new row added.
 */
function markCustomerVisit(meetingDate, v) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DSR_SHEET_NAME);
  if (!sheet) return;

  const day = parseInt(meetingDate.split("-")[2], 10);
  if (!day || day < 1 || day > 31) return;
  const dayCol = DSR_FIRST_DAY_COLUMN + (day - 1);

  const dataStartRow = findDSRDataStartRow(sheet);
  const lastRow = Math.max(sheet.getLastRow(), dataStartRow - 1);
  const customerName = normalizeName(v.customerName);
  let targetRow = -1;

  if (lastRow >= dataStartRow) {
    const names = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, 1).getValues();

    // Pass 1: exact match (case/whitespace-insensitive)
    for (let i = 0; i < names.length; i++) {
      if (names[i][0] && normalizeName(names[i][0]) === customerName) {
        targetRow = dataStartRow + i;
        break;
      }
    }

    // Pass 2: safe fallback — only if typed name is a distinct prefix of EXACTLY ONE
    // existing customer (e.g. "sparrow" -> "sparrow apparel"). If it matches more than
    // one, we don't guess — a new row is safer than merging two different customers.
    if (targetRow === -1 && customerName.length >= 4) {
      const candidates = [];
      for (let i = 0; i < names.length; i++) {
        const n = normalizeName(names[i][0]);
        if (n && (n.startsWith(customerName) || customerName.startsWith(n))) {
          candidates.push(dataStartRow + i);
        }
      }
      if (candidates.length === 1) targetRow = candidates[0];
    }
  }

  if (targetRow === -1) {
    targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1).setValue((v.customerName || "").trim());
    if (v.brandName) sheet.getRange(targetRow, 5).setValue(v.brandName); // column E = Buyers
  }

  sheet.getRange(targetRow, dayCol).setValue("Y");
}

/**
 * Colors the ENTIRE day-column (all customer rows) based on the day's status:
 * green = visit, orange = office, purple = leave, red = holiday.
 * Works even when no customer visit happened that day.
 */
function markDayColumn(meetingDate, kind) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DSR_SHEET_NAME);
  if (!sheet) return;

  const color = STATUS_COLORS[kind];
  if (!color) return;

  const day = parseInt(meetingDate.split("-")[2], 10);
  if (!day || day < 1 || day > 31) return;
  const dayCol = DSR_FIRST_DAY_COLUMN + (day - 1);

  const dataStartRow = findDSRDataStartRow(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow >= dataStartRow) {
    sheet.getRange(dataStartRow, dayCol, lastRow - dataStartRow + 1, 1).setBackground(color);
  }
  // also color the header cell for that day so it's visible even with 0 customer rows
  sheet.getRange(1, dayCol).setBackground(color);
}

function escapeHtml(str) {
  if (!str) return "";
  return str.toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
