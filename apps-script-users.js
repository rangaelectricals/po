// =================================================================
// ADD THIS TO YOUR EXISTING Google Apps Script (Code.gs)
// =================================================================
// This adds user management endpoints to your existing Apps Script.
// It stores users in a "Users" sheet in the same spreadsheet.
//
// HOW TO ADD:
// 1. Open your Google Apps Script project
// 2. Paste the functions below into your existing Code.gs file
// 3. Add the cases to your existing doGet() and doPost() switch/if blocks
// 4. Click "Deploy" → "Manage Deployments" → edit → bump version → Deploy
// =================================================================

// --- ADD THESE CASES TO YOUR EXISTING doGet() FUNCTION ---
//
//   if (action === 'read_users') {
//     return ContentService.createTextOutput(JSON.stringify(readUsers()))
//       .setMimeType(ContentService.MimeType.JSON);
//   }

// --- ADD THESE CASES TO YOUR EXISTING doPost() FUNCTION ---
//
//   if (action === 'save_users') {
//     var users = payload.users || [];
//     saveUsers(users);
//     return ContentService.createTextOutput(JSON.stringify({ success: true }))
//       .setMimeType(ContentService.MimeType.JSON);
//   }


// ============================================================
// USER MANAGEMENT FUNCTIONS — paste these in your Code.gs
// ============================================================

/**
 * Get or create the "Users" sheet.
 * Columns: A=username, B=password, C=role
 */
function getUsersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Users');
  if (!sheet) {
    sheet = ss.insertSheet('Users');
    sheet.appendRow(['username', 'password', 'role']); // header
    sheet.appendRow(['ranga', 'Ranga@7677', 'admin']);  // default admin
  }
  return sheet;
}

/**
 * Read all users from the Users sheet.
 * Returns an array of { username, password, role } objects.
 */
function readUsers() {
  var sheet = getUsersSheet();
  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) { // skip header
    if (data[i][0]) {
      users.push({
        username: String(data[i][0]).trim(),
        password: String(data[i][1]),
        role: String(data[i][2]).trim() || 'viewer'
      });
    }
  }
  return users;
}

/**
 * Save the entire users array to the Users sheet (overwrites existing data).
 * @param {Array} users - Array of { username, password, role }
 */
function saveUsers(users) {
  var sheet = getUsersSheet();
  // Clear everything except header
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clear();
  }
  // Write users
  for (var i = 0; i < users.length; i++) {
    sheet.getRange(i + 2, 1, 1, 3).setValues([
      [users[i].username, users[i].password, users[i].role || 'viewer']
    ]);
  }
}
