// ═══════════════════════════════════════════════════════════
// 5-FACTOR WORK ASSESSMENT — Google Apps Script
// Paste this entire file into your Apps Script editor
// ═══════════════════════════════════════════════════════════

const SHEET_NAME = "Results";
const HEADERS = ["Timestamp","Name","F1","F2","F3","F4","F5","Top Match","Fit Score (%)"];

// ── GET: Teacher dashboard reads all data ──────────────────
function doGet(e) {
  const action = e.parameter.action;
  
  if (action === "getAll") {
    const sheet = getOrCreateSheet();
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) {
      return jsonResponse([]);
    }
    const headers = rows[0];
    const data = rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h.toLowerCase().replace(/\s+/g,"_").replace(/[()%]/g,"")] = row[i]; });
      // normalize keys to match HTML expectations
      return {
        name:     obj["name"],
        timestamp:obj["timestamp"],
        F1:       obj["f1"],
        F2:       obj["f2"],
        F3:       obj["f3"],
        F4:       obj["f4"],
        F5:       obj["f5"],
        topMatch: obj["top_match"],
        fitScore: obj["fit_score_"],
      };
    });
    return jsonResponse(data);
  }
  
  return jsonResponse({ status: "ok", message: "5-Factor Assessment API running" });
}

// ── POST: Student submits result ───────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();
    
    sheet.appendRow([
      payload.timestamp || new Date().toISOString(),
      payload.name      || "Anonymous",
      parseFloat(payload.F1) || 0,
      parseFloat(payload.F2) || 0,
      parseFloat(payload.F3) || 0,
      parseFloat(payload.F4) || 0,
      parseFloat(payload.F5) || 0,
      payload.topMatch  || "",
      parseFloat(payload.fitScore) || 0,
    ]);
    
    return jsonResponse({ status: "success", message: "Result saved" });
  } catch(err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

// ── HELPERS ────────────────────────────────────────────────
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    // Format header row
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setBackground("#1a1a2e")
      .setFontColor("#818cf8")
      .setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
