/**
 * Google Apps Script Web App for Scholars' Card Verification
 * 
 * Target Sheet Structure:
 * - Column D (index 3): ID Number (e.g. 256390)
 * - Column E (index 4): Full Name [Last Name, First Name M.I.]
 */
function doGet(e) {
  try {
    const rawId = (e && e.parameter && e.parameter.id) ? e.parameter.id.trim() : "";
    
    if (!rawId) {
      return jsonResponse({ valid: false, error: "Missing 'id' parameter" });
    }

    // Input format is expected to be 'YYYY-YYYY-StudentID' (e.g., '2025-2026-256390')
    // Extract everything after the second hyphen as the StudentID
    const parts = rawId.split("-");
    const studentId = (parts.length >= 3 ? parts.slice(2).join("-") : rawId).trim();

    if (!studentId) {
      return jsonResponse({ valid: false, error: "Invalid ID format" });
    }

    // Access the active spreadsheet and the target sheet (first sheet or by name)
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data = sheet.getDataRange().getValues();

    // Iterate through rows (skipping header row at index 0)
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const sheetIdNumber = String(row[3] !== undefined && row[3] !== null ? row[3] : "").trim();

      // Check for exact match against Column D
      if (sheetIdNumber === studentId) {
        const fullName = String(row[4] || "").trim();
        return jsonResponse({
          valid: true,
          name: fullName
        });
      }
    }

    // If no match found
    return jsonResponse({ valid: false });

  } catch (err) {
    return jsonResponse({
      valid: false,
      error: err.toString()
    });
  }
}

/**
 * Helper to serialize and return JSON response with CORS headers
 */
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
