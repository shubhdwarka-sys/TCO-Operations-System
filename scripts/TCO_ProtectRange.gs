// ═══════════════════════════════════════════════════════════════════
// TCO_ProtectRange.gs
// Company: Loan 11 Possible Pvt Ltd (The Car Owner)
// Author: Shubh — Automation Manager | March 2026
// Script 2 of 2
//
// Purpose: Column protection + Editor email management
// Alag file hai — Auth changes yahan karo, main script touch mat karo
// ═══════════════════════════════════════════════════════════════════


// ═══════════════════════════════════════
// SECTION 1 — EDITOR CONFIG
// Emails yahan change karo — phir setupAllProtections() run karo
// ═══════════════════════════════════════

var PROTECT_CONFIG = {
  // Admin — full access har jagah
  ADMIN: 'admin.loan11@gmail.com',

  // User 1 — DM_SHEET manual entry (Col 2, 4-24, 28-30)
  USER_1: 'dalveernegi9@gmail.com',

  // User 2 — DM_SHEET approval only (Col 31-33)
  USER_2: 'shubhamkumar41664@gmail.com',

  // User 3 — DM_SHEET approval (31-33) + ACCOUNT manual entry (Col 14, 19-41)
  USER_3: 'acc.loan11possible@gmail.com',

  // User 4 — RTO_TRACKER manual entry (Col 12-15, 19-27)
  USER_4: 'deepakrto99988@gmail.com',

  HEADER_ROW: 2,
};


// ═══════════════════════════════════════
// SECTION 2 — SETUP ALL PROTECTIONS
// Ek baar run karo — sab lock ho jayega
// ═══════════════════════════════════════

function setupAllProtections() {
  var ss = SpreadsheetApp.getActive();

  // Pehle purani protections hatao
  removeAllProtections();
  SpreadsheetApp.flush();

  // ─── DM_SHEET ───
  setupDMProtections_(ss);

  // ─── ACCOUNT_PAYMENT_TRACKER ───
  setupAccountProtections_(ss);

  // ─── RTO_TRACKER ───
  setupRTOProtections_(ss);

  // ─── Full sheet locks ───
  lockFullSheet_(ss, 'MASTER_DATA', 'Script sync — read only');
  lockFullSheet_(ss, 'AUTH_AUDIT_LOG', 'Audit log — tamper-proof');
  lockFullSheet_(ss, 'PAYMENT_TO_RTO_LOG', 'Payment log — tamper-proof');

  SpreadsheetApp.getActive().toast('Sab protections lag gayin! ✅', 'Protection Setup', 5);
}


// ═══════════════════════════════════════
// SECTION 3 — DM_SHEET PROTECTIONS
// ═══════════════════════════════════════

function setupDMProtections_(ss) {
  var sheet = ss.getSheetByName('DM_SHEET');
  if (!sheet) return;
  var maxRow = sheet.getMaxRows();

  // ─── LOCKED COLUMNS (sirf Admin) ───
  // Col 1 — DM NO (auto-generate)
  lockColumn_(sheet, 1, maxRow, 'DM NO — Auto Generate');

  // Col 3 — MONTH (auto-fill)
  lockColumn_(sheet, 3, maxRow, 'MONTH — Auto Fill');

  // Col 25 — EXECUTIVE NAME (VLOOKUP)
  lockColumn_(sheet, 25, maxRow, 'EXEC NAME — VLOOKUP');

  // Col 26 — BRANCH (VLOOKUP)
  lockColumn_(sheet, 26, maxRow, 'BRANCH — VLOOKUP');

  // Col 27 — TEAM LEADER (VLOOKUP)
  lockColumn_(sheet, 27, maxRow, 'TEAM LEADER — VLOOKUP');

  // ─── USER 1 COLUMNS (manual entry) ───
  // Col 2 (LOGIN DATE), Col 4-24, Col 28-30
  var user1Cols = [2, 4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24, 28,29,30];
  for (var i = 0; i < user1Cols.length; i++) {
    protectColumnForUsers_(sheet, user1Cols[i], maxRow,
      'DM Manual — Col ' + user1Cols[i],
      [PROTECT_CONFIG.ADMIN, PROTECT_CONFIG.USER_1]);
  }

  // ─── USER 2 + USER 3 COLUMNS (approval) ───
  // Col 31 (APPROVAL STATUS), Col 32 (APPROVAL DATE), Col 33 (CHALLAN)
  var approvalCols = [31, 32, 33];
  for (var j = 0; j < approvalCols.length; j++) {
    protectColumnForUsers_(sheet, approvalCols[j], maxRow,
      'DM Approval — Col ' + approvalCols[j],
      [PROTECT_CONFIG.ADMIN, PROTECT_CONFIG.USER_2, PROTECT_CONFIG.USER_3]);
  }
}


// ═══════════════════════════════════════
// SECTION 4 — ACCOUNT PROTECTIONS
// ═══════════════════════════════════════

function setupAccountProtections_(ss) {
  var sheet = ss.getSheetByName('ACCOUNT_PAYMENT_TRACKER');
  if (!sheet) return;
  var maxRow = sheet.getMaxRows();

  // ─── LOCKED COLUMNS — Auto-pull zone (sirf Admin) ───
  // Col 1-13, 15-18 — auto-pull from DM
  var autoLockCols = [1,2,3,4,5,6,7,8,9,10,11,12,13, 15,16,17,18];
  for (var i = 0; i < autoLockCols.length; i++) {
    lockColumn_(sheet, autoLockCols[i], maxRow, 'Account Auto-Pull — Col ' + autoLockCols[i]);
  }

  // Col 42-45 — RC/Vendor auto-pull
  var rcLockCols = [42, 43, 44, 45];
  for (var r = 0; r < rcLockCols.length; r++) {
    lockColumn_(sheet, rcLockCols[r], maxRow, 'Account RC Auto-Pull — Col ' + rcLockCols[r]);
  }

  // ─── USER 3 COLUMNS (manual entry) ───
  // Col 14 (UTR), Col 19-41 (Payers, Hold, Dealer Payment, Scores, Remarks)
  var user3Cols = [14, 19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41];
  for (var j = 0; j < user3Cols.length; j++) {
    protectColumnForUsers_(sheet, user3Cols[j], maxRow,
      'Account Manual — Col ' + user3Cols[j],
      [PROTECT_CONFIG.ADMIN, PROTECT_CONFIG.USER_3]);
  }
}


// ═══════════════════════════════════════
// SECTION 5 — RTO PROTECTIONS
// ═══════════════════════════════════════

function setupRTOProtections_(ss) {
  var sheet = ss.getSheetByName('RTO_TRACKER');
  if (!sheet) return;
  var maxRow = sheet.getMaxRows();

  // ─── LOCKED COLUMNS — Auto-pull zone (sirf Admin) ───
  // Col 1-8 (auto-pull from Account)
  for (var i = 1; i <= 8; i++) {
    lockColumn_(sheet, i, maxRow, 'RTO Auto-Pull — Col ' + i);
  }

  // Col 9 — RTO CODE (auto-generate)
  lockColumn_(sheet, 9, maxRow, 'RTO CODE — Auto Generate');

  // Col 10-11 (Vehicle Model, Mfg Year — auto-pull)
  lockColumn_(sheet, 10, maxRow, 'RTO Auto-Pull — Col 10');
  lockColumn_(sheet, 11, maxRow, 'RTO Auto-Pull — Col 11');

  // Col 16-18 (Buyer Name/Phone, Bank — auto-pull)
  lockColumn_(sheet, 16, maxRow, 'RTO Auto-Pull — Col 16');
  lockColumn_(sheet, 17, maxRow, 'RTO Auto-Pull — Col 17');
  lockColumn_(sheet, 18, maxRow, 'RTO Auto-Pull — Col 18');

  // ─── USER 4 COLUMNS (manual entry) ───
  // Col 12-15 (Chassis, Engine, Seller Name/Phone)
  // Col 19-27 (Case Type, Vendor, Scan, Receipt, Transfer, Status, RC, Remarks)
  var user4Cols = [12,13,14,15, 19,20,21,22,23,24,25,26,27];
  for (var j = 0; j < user4Cols.length; j++) {
    protectColumnForUsers_(sheet, user4Cols[j], maxRow,
      'RTO Manual — Col ' + user4Cols[j],
      [PROTECT_CONFIG.ADMIN, PROTECT_CONFIG.USER_4]);
  }
}


// ═══════════════════════════════════════
// SECTION 6 — HELPER FUNCTIONS
// ═══════════════════════════════════════

// Lock column — sirf Admin edit kar sake
function lockColumn_(sheet, col, maxRow, desc) {
  var range = sheet.getRange(PROTECT_CONFIG.HEADER_ROW + 1, col, maxRow - PROTECT_CONFIG.HEADER_ROW, 1);
  var protection = range.protect().setDescription(desc);
  protection.removeEditors(protection.getEditors());
  protection.addEditor(PROTECT_CONFIG.ADMIN);
  protection.setWarningOnly(false);
}

// Protect column — specific users can edit
function protectColumnForUsers_(sheet, col, maxRow, desc, allowedEmails) {
  var range = sheet.getRange(PROTECT_CONFIG.HEADER_ROW + 1, col, maxRow - PROTECT_CONFIG.HEADER_ROW, 1);
  var protection = range.protect().setDescription(desc);
  protection.removeEditors(protection.getEditors());
  for (var i = 0; i < allowedEmails.length; i++) {
    protection.addEditor(allowedEmails[i]);
  }
  protection.setWarningOnly(false);
}

// Lock full sheet — sirf Admin
function lockFullSheet_(ss, sheetName, desc) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  var protection = sheet.protect().setDescription(desc);
  protection.removeEditors(protection.getEditors());
  protection.addEditor(PROTECT_CONFIG.ADMIN);
  protection.setWarningOnly(false);
}


// ═══════════════════════════════════════
// SECTION 7 — REMOVE ALL PROTECTIONS
// Emergency — sab unlock karna ho
// ═══════════════════════════════════════

function removeAllProtections() {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    // Sheet-level protections
    var sheetProtections = sheets[s].getProtections(SpreadsheetApp.ProtectionType.SHEET);
    for (var i = 0; i < sheetProtections.length; i++) {
      sheetProtections[i].remove();
    }
    // Range-level protections
    var rangeProtections = sheets[s].getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var j = 0; j < rangeProtections.length; j++) {
      rangeProtections[j].remove();
    }
  }
  SpreadsheetApp.getActive().toast('Sab protections hat gayin!', 'Protections Removed', 5);
}


// ═══════════════════════════════════════
// SECTION 8 — UPDATE EDITORS
// Emails change karo CONFIG mein, phir ye run karo
// ═══════════════════════════════════════

function updateEditors() {
  // Pehle sab hatao
  removeAllProtections();
  SpreadsheetApp.flush();

  // Phir naye lagao
  setupAllProtections();
  SpreadsheetApp.getActive().toast('Editors updated + Protections refreshed!', 'Done', 5);
}


// ═══════════════════════════════════════
// SECTION 9 — VIEW CURRENT PROTECTIONS
// Check karo kya-kya locked hai
// ═══════════════════════════════════════

function viewProtections() {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  var report = [];

  for (var s = 0; s < sheets.length; s++) {
    var sheetName = sheets[s].getName();
    var sheetProts = sheets[s].getProtections(SpreadsheetApp.ProtectionType.SHEET);
    var rangeProts = sheets[s].getProtections(SpreadsheetApp.ProtectionType.RANGE);

    if (sheetProts.length > 0) {
      report.push('🔒 ' + sheetName + ' — FULL SHEET LOCKED');
    }

    if (rangeProts.length > 0) {
      report.push('📋 ' + sheetName + ' — ' + rangeProts.length + ' protected ranges');
      for (var r = 0; r < rangeProts.length; r++) {
        var desc = rangeProts[r].getDescription();
        var editors = rangeProts[r].getEditors().map(function(e) { return e.getEmail(); });
        report.push('   ' + desc + ' → ' + editors.join(', '));
      }
    }
  }

  if (report.length === 0) {
    report.push('❌ Koi protection nahi hai!');
  }

  SpreadsheetApp.getUi().alert(
    '🔐 PROTECTION REPORT\n━━━━━━━━━━━━━━━━━━━━\n\n' +
    report.join('\n') +
    '\n\n━━━━━━━━━━━━━━━━━━━━\nTCO Operations System'
  );
}
