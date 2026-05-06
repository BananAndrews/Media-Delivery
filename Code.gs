// ============================================================
//  Blue Belle Weddings – Media Delivery Portal
//  Server-side Apps Script (Code.gs)
//
//  HOW TO DEPLOY
//  1. Open any Google Spreadsheet (or create a standalone project):
//     Extensions → Apps Script
//  2. Rename the default file to Code.gs and paste this content.
//  3. Create a new HTML file named "Index" and paste Index.html content.
//  4. Deploy → New Deployment → Web App
//       Execute as : Me (so the script can write to the sheet)
//       Who has access: Anyone
//  5. Copy the Web App URL and embed it in the Creatives Portal Google Site.
// ============================================================

// ── Configuration ─────────────────────────────────────────────
var SPREADSHEET_ID = "1Tf-h96pkJ1JZ_3TJAsYvFt7d-PHt5w-jm7TaH3Atq7A";
var SHEET_NAME     = "Media Delivery";
var STATUS_FILTER  = "Media to Upload";
var STATUS_DONE    = "Media Uploaded";

// "Media Delivery" sheet – column indices (0-based, A=0)
var COL_TYPE   = 4;   // E – "Photos" or "Video"
var COL_NAME   = 5;   // F – Project / couple name
var COL_LINK   = 8;   // I – Upload link
var COL_SONGS         = 9;   // J – Songs choice (Video only)
var COL_SPECIAL_NOTES = 10;  // K – Special notes (shared field, written by whoever submits)
var COL_LEAD          = 11;  // L – Lead creative name
var COL_SECOND = 12;  // M – Second creative name (may be empty)
var COL_STATUS = 13;  // N – "Media to Upload" / "Media Uploaded"

// Lead creative write columns
var COL_LEAD_FILES   = 14;  // O
var COL_LEAD_EXTRAS  = 16;  // Q
var COL_LEAD_INVOICE = 18;  // S
var COL_LEAD_NOTES   = 20;  // U

// Second creative write columns
var COL_SECOND_FILES   = 15;  // P
var COL_SECOND_EXTRAS  = 17;  // R
var COL_SECOND_INVOICE = 19;  // T
var COL_SECOND_NOTES   = 21;  // V

// Read up to column W (index 22) to cover all submission columns
var READ_WIDTH = 23;

// ── Web app entry point ────────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Blue Belle – Media Delivery")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── getProjects ────────────────────────────────────────────────
/**
 * Returns all rows from the Media Delivery sheet where
 * Column N = "Media to Upload". Called from the client on page load.
 */
function getProjects() {
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Sheet '" + SHEET_NAME + "' not found.");

  var data     = sheet.getDataRange().getValues();
  var projects = [];

  for (var r = 1; r < data.length; r++) {   // row 0 = header
    var row    = data[r];
    var status = String(row[COL_STATUS]).trim();
    if (status !== STATUS_FILTER) continue;

    var name = String(row[COL_NAME]).trim();
    if (!name) continue;

    projects.push({
      sheetRow: r + 1,                          // 1-based sheet row
      name:     name,
      type:     String(row[COL_TYPE]).trim(),
      link:     String(row[COL_LINK]).trim(),
      songs:    String(row[COL_SONGS]).trim(),
      lead:     String(row[COL_LEAD]).trim(),
      second:   String(row[COL_SECOND]).trim()
    });
  }

  return projects;
}

// ── submitForm ─────────────────────────────────────────────────
/**
 * Writes form data to the correct columns and updates Column N status.
 * data = { sheetRow, role ("lead"|"second"), filesAmount, extras,
 *          invoiceAmount, notes, songs }
 */
function submitForm(data) {
  var sheet    = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  var sheetRow = Number(data.sheetRow);
  var isLead   = data.role === "lead";

  Logger.log("submitForm received: " + JSON.stringify(data));
  Logger.log("Writing special notes '" + data.specialNotes + "' to row " + sheetRow + ", col " + (COL_SPECIAL_NOTES + 1));

  // Write submission data to the correct set of columns
  if (isLead) {
    sheet.getRange(sheetRow, COL_LEAD_FILES   + 1).setValue(data.filesAmount);
    sheet.getRange(sheetRow, COL_LEAD_EXTRAS  + 1).setValue(data.extras   || 0);
    sheet.getRange(sheetRow, COL_LEAD_INVOICE + 1).setValue(data.invoiceAmount);
    sheet.getRange(sheetRow, COL_LEAD_NOTES   + 1).setValue(data.notes    || "");
  } else {
    sheet.getRange(sheetRow, COL_SECOND_FILES   + 1).setValue(data.filesAmount);
    sheet.getRange(sheetRow, COL_SECOND_EXTRAS  + 1).setValue(data.extras   || 0);
    sheet.getRange(sheetRow, COL_SECOND_INVOICE + 1).setValue(data.invoiceAmount);
    sheet.getRange(sheetRow, COL_SECOND_NOTES   + 1).setValue(data.notes    || "");
  }

  // Append special notes to Column K — preserve existing content from the other creative
  var newSpecialNotes = (data.specialNotes || "").trim();
  if (newSpecialNotes) {
    var existingSpecialNotes = String(sheet.getRange(sheetRow, COL_SPECIAL_NOTES + 1).getValue()).trim();
    var combinedSpecialNotes = existingSpecialNotes ? existingSpecialNotes + "\n" + newSpecialNotes : newSpecialNotes;
    sheet.getRange(sheetRow, COL_SPECIAL_NOTES + 1).setValue(combinedSpecialNotes);
  }

  // Append songs choice to Column J — preserve existing content from the other creative
  var newSongs = (data.songs || "").trim();
  if (newSongs) {
    var existingSongs = String(sheet.getRange(sheetRow, COL_SONGS + 1).getValue()).trim();
    var combinedSongs = existingSongs ? existingSongs + "\n" + newSongs : newSongs;
    sheet.getRange(sheetRow, COL_SONGS + 1).setValue(combinedSongs);
  }

  SpreadsheetApp.flush();  // Commit all writes before reading back

  // Re-read the row to evaluate status-change scenario
  var fresh     = sheet.getRange(sheetRow, 1, 1, READ_WIDTH).getValues()[0];
  var hasSecond = String(fresh[COL_SECOND]).trim() !== "";

  if (!hasSecond) {
    // Scenario A: single creative → mark done immediately
    sheet.getRange(sheetRow, COL_STATUS + 1).setValue(STATUS_DONE);
  } else {
    var leadDone   = String(fresh[COL_LEAD_FILES]).trim()     !== "" &&
                     String(fresh[COL_LEAD_INVOICE]).trim()   !== "";
    var secondDone = String(fresh[COL_SECOND_FILES]).trim()   !== "" &&
                     String(fresh[COL_SECOND_INVOICE]).trim() !== "";

    if (leadDone && secondDone) {
      // Scenario C: both submitted → mark done
      sheet.getRange(sheetRow, COL_STATUS + 1).setValue(STATUS_DONE);
    }
    // Scenario B: first of two submitted → no status change
  }

  return { success: true };
}
