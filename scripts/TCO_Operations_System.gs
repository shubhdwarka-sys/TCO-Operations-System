// ═══════════════════════════════════════════════════════════════════
// TCO_Operations_System.gs — FINAL v2
// Company: Loan 11 Possible Pvt Ltd (The Car Owner)
// Author: Shubh — Automation Manager | March 2026
// Script 1 of 2
//
// FONT COLORS:
// Manual Entry = Black (#000000) Bold
// Auto-Pull    = Deep Blue (#1A4D8F) Bold Italic
// Auto-Generate = Dark Green (#1B5E20) Bold
// DM NO + REG NO = Size 12
// Email ID = lowercase (no CAPS)
// All = Center + Middle + Wrap
// ═══════════════════════════════════════════════════════════════════

// SECTION 1 — CONFIG
var CONFIG = {
  DM_SHEET: 'DM_SHEET',
  ACC_SHEET: 'ACCOUNT_PAYMENT_TRACKER',
  RTO_SHEET: 'RTO_TRACKER',
  MASTER_DATA: 'MASTER_DATA',
  AUDIT_LOG: 'AUTH_AUDIT_LOG',
  PAYMENT_LOG: 'PAYMENT_TO_RTO_LOG',
  BACKUP_DM: '_BACKUP_DM',
  BACKUP_ACC: '_BACKUP_ACC',
  BACKUP_RTO: '_BACKUP_RTO',
  HEADER_ROW: 2,
  ADMIN_EMAIL: 'admin.loan11@gmail.com',
  ADMIN_NAME: 'Shubh',

  DM: { DM_NO:1, LOGIN_DATE:2, MONTH:3, BUYER_NAME:4, PHONE:5, PHONE2:6, EMAIL:7, REG_NO:8, VEHICLE_MODEL:9, MFG_YEAR:10, BANK_NAME:11, PRODUCT:12, LOAN_AMT_APPLIED:13, LOAN_APPROVED_AMT:14, INSURANCE_AMT:15, LOAN_SURAKHSA:16, FILE_CHARGE:17, ROI:18, EMI:19, TENURE:20, DISBURSEMENT_AMT:21, DISBURSEMENT_DATE:22, RC_TRANSFER_CHARGES:23, EMP_ID:24, EXEC_NAME:25, BRANCH:26, TEAM_LEADER:27, DEALERSHIP_NAME:28, CASE_STATUS:29, REMARKS:30, APPROVAL_STATUS:31, APPROVAL_DATE:32, CHALLAN_AMT:33, TOTAL_COLS:33 },

  ACC: { DISBURSEMENT_DATE:1, DM_NO:2, CASE_STATUS:3, APPROVAL_STATUS:4, BUYER_NAME:5, PHONE:6, PHONE2:7, REG_NO:8, VEHICLE_MODEL:9, MFG_YEAR:10, BANK_NAME:11, PRODUCT:12, DISBURSEMENT_AMT:13, UTR_NO:14, EXEC_NAME:15, BRANCH:16, DEALERSHIP_NAME:17, APPROVAL_DATE:18, PAYER1_NAME:19, PAYER1_DATE:20, PAYER1_AMT:21, PAYER1_ACC:22, PAYER1_STATUS:23, PAYER2_NAME:24, PAYER2_DATE:25, PAYER2_AMT:26, PAYER2_ACC:27, PAYER2_STATUS:28, PAYER3_NAME:29, PAYER3_DATE:30, PAYER3_AMT:31, PAYER3_ACC:32, PAYER3_STATUS:33, HOLD_AMT:34, DEALER_PAY_STATUS:35, DEALER_PAY_DATE:36, PAYOUT_FROM_BANK:37, PAYOUT_TO_DEALER:38, EXEC_INCENTIVE:39, LOAN_SCORE:40, REMARKS:41, RC_CHARGES:42, RC_STATUS:43, VENDOR_NAME:44, VENDOR_PAYMENT:45, TOTAL_COLS:45 },

  RTO: { DISBURSEMENT_DATE:1, DM_NO:2, DEALER_PAY_STATUS:3, PRODUCT:4, BRANCH:5, EXEC_NAME:6, DEALERSHIP_NAME:7, REG_NO:8, RTO_CODE:9, VEHICLE_MODEL:10, MFG_YEAR:11, CHASSIS_NO:12, ENGINE_NO:13, SELLER_NAME:14, SELLER_PHONE:15, BUYER_NAME:16, BUYER_PHONE:17, BANK_NAME:18, CASE_TYPE:19, RTO_VENDOR:20, SCAN_FILE:21, RECEIPT_NO:22, TRANSFER_DATE:23, RTO_STATUS:24, RC_RECEIVED:25, SYSTEM_REMARKS:26, REMARKS_ISSUE:27, TOTAL_COLS:27 },

  DM_TO_ACC: { 22:1, 1:2, 29:3, 31:4, 4:5, 5:6, 6:7, 8:8, 9:9, 10:10, 11:11, 12:12, 21:13, 25:15, 26:16, 28:17, 32:18 },
  ACC_TO_RTO: { 1:1, 2:2, 35:3, 12:4, 16:5, 15:6, 17:7, 8:8, 9:10, 10:11, 5:16, 6:17, 11:18 },

  // ─── 3 COLORS ───
  COLOR_MANUAL: '#000000',      // Black — manual entry
  COLOR_AUTO_PULL: '#1A4D8F',   // Deep Blue — auto-pull from other sheet
  COLOR_AUTO_GEN: '#1B5E20',    // Dark Green — auto-generate by script
  COLOR_DM_NO: '#3E2723',       // Dark Brown — DM NO (all 3 sheets)
  COLOR_REG_NO: '#E65100',      // Dark Orange — Vehicle/Reg NO (all 3 sheets)

  // ─── Column types per sheet ───
  // DM: Auto-generate cols
  AUTO_GEN_DM: [1, 3, 25, 26, 27],  // DM_NO, MONTH, EXEC, BRANCH, TL
  // DM: Manual cols (everything else)
  MANUAL_DM: [2, 4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24, 28,29,30,31,32,33],

  // ACC: Auto-pull cols
  AUTO_PULL_ACC: [1,2,3,4,5,6,7,8,9,10,11,12,13,15,16,17,18],
  // ACC: Auto-pull RC/Vendor
  AUTO_PULL_ACC_RC: [42,43,44,45],
  // ACC: Manual cols
  MANUAL_ACC: [14, 19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41],

  // RTO: Auto-pull cols
  AUTO_PULL_RTO: [1,2,3,4,5,6,7,10,11,16,17,18],
  // RTO: Auto-generate cols
  AUTO_GEN_RTO: [8, 9],  // REG_NO (pulled but displayed as auto), RTO_CODE
  // RTO: Manual cols
  MANUAL_RTO: [12,13,14,15, 19,20,21,22,23,24,25,26,27],

  // ─── CAPS SKIP cols (Email + Dropdowns) ───
  CAPS_SKIP_DM: [7, 11,12,26,28,29,31],   // 7=Email
  CAPS_SKIP_ACC: [3,4,11,12,16,17,23,28,33,35,43],
  CAPS_SKIP_RTO: [3,4,5,18,19,24,25],

  // ─── Lowercase cols (Email — always lowercase) ───
  LOWERCASE_DM: [7],   // EMAIL ID
  LOWERCASE_ACC: [],
  LOWERCASE_RTO: [],

  // ─── Special size cols ───
  SIZE_12_DM: [1, 8],    // DM_NO, REG_NO
  SIZE_12_ACC: [2, 8],   // DM_NO, REG_NO
  SIZE_12_RTO: [2, 8],   // DM_NO, REG_NO

  // DM_NO column per sheet (Dark Brown)
  DM_NO_COL_DM: 1,
  DM_NO_COL_ACC: 2,
  DM_NO_COL_RTO: 2,

  // REG_NO column per sheet (Dark Orange)
  REG_NO_COL_DM: 8,
  REG_NO_COL_ACC: 8,
  REG_NO_COL_RTO: 8,

  DM_PREFIX: 'TCO',
  STALE_DAYS: 7,
  MASTER_ACC_START: 34,
  MASTER_RTO_START: 60,
  ACC_EXTRA_COLS: [19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45],
  RTO_EXTRA_COLS: [9,12,13,14,15,20,21,22,23,24,25,26,27],
};


// ═══════════════════════════════════════
// SECTION 2 — DM_NO AUTO-GENERATION
// ═══════════════════════════════════════

function generateDMNo_(sheet, row) {
  var loginDate = sheet.getRange(row, CONFIG.DM.LOGIN_DATE).getValue();
  if (!loginDate) return;
  var existingDM = sheet.getRange(row, CONFIG.DM.DM_NO).getValue();
  if (existingDM) return;
  var date = new Date(loginDate);
  if (isNaN(date.getTime())) return;

  var months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
  var prefix = CONFIG.DM_PREFIX + months[date.getMonth()] + String(date.getFullYear()).slice(-2);
  var lastRow = sheet.getLastRow();
  var allDmNos = lastRow > CONFIG.HEADER_ROW ? sheet.getRange(CONFIG.HEADER_ROW+1, CONFIG.DM.DM_NO, lastRow-CONFIG.HEADER_ROW, 1).getValues().flat() : [];

  var maxSerial = 0;
  for (var i = 0; i < allDmNos.length; i++) {
    if (allDmNos[i] && allDmNos[i].toString().startsWith(prefix)) {
      var s = parseInt(allDmNos[i].toString().slice(prefix.length));
      if (s > maxSerial) maxSerial = s;
    }
  }

  var newDmNo = prefix + String(maxSerial + 1).padStart(3, '0');
  sheet.getRange(row, CONFIG.DM.DM_NO).setValue(newDmNo);
  var firstOfMonth = new Date(date.getFullYear(), date.getMonth(), 1);
  sheet.getRange(row, CONFIG.DM.MONTH).setValue(firstOfMonth);

  // Format — Auto-Generate = Dark Green, Size 12
  formatCellStyle_(sheet, row, CONFIG.DM.DM_NO, 'auto_gen', true);
  formatCellStyle_(sheet, row, CONFIG.DM.MONTH, 'auto_gen', false);

  logAudit_('DM_NO_GENERATED', CONFIG.DM_SHEET, row, 'DM_NO: ' + newDmNo);
}

function guardDMNoSeries() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.DM_SHEET);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow <= CONFIG.HEADER_ROW) return;
  for (var r = CONFIG.HEADER_ROW+1; r <= lastRow; r++) {
    var loginDate = sheet.getRange(r, CONFIG.DM.LOGIN_DATE).getValue();
    var dmNo = sheet.getRange(r, CONFIG.DM.DM_NO).getValue();
    if (loginDate && !dmNo) generateDMNo_(sheet, r);
  }
}


// ═══════════════════════════════════════
// SECTION 3 — AUTO-PULL: DM → ACCOUNT
// ═══════════════════════════════════════

function checkAndPullDMToAccount_(sheet, row) {
  var ss = SpreadsheetApp.getActive();
  var dmNo = sheet.getRange(row, CONFIG.DM.DM_NO).getValue();
  var buyer = sheet.getRange(row, CONFIG.DM.BUYER_NAME).getValue();
  var phone = sheet.getRange(row, CONFIG.DM.PHONE).getValue();
  var regNo = sheet.getRange(row, CONFIG.DM.REG_NO).getValue();
  if (!dmNo || !buyer || !phone || !regNo) return;

  var accSheet = ss.getSheetByName(CONFIG.ACC_SHEET);
  if (!accSheet) return;

  var accLastRow = accSheet.getLastRow();
  var existingRow = 0;
  if (accLastRow > CONFIG.HEADER_ROW) {
    var accDmNos = accSheet.getRange(CONFIG.HEADER_ROW+1, CONFIG.ACC.DM_NO, accLastRow-CONFIG.HEADER_ROW, 1).getValues().flat();
    var idx = accDmNos.indexOf(dmNo);
    if (idx >= 0) existingRow = CONFIG.HEADER_ROW + 1 + idx;
  }

  var sourceData = sheet.getRange(row, 1, 1, CONFIG.DM.TOTAL_COLS).getValues()[0];
  var targetRow = existingRow > 0 ? existingRow : accSheet.getLastRow() + 1;

  var map = CONFIG.DM_TO_ACC;
  var keys = Object.keys(map);
  for (var i = 0; i < keys.length; i++) {
    var value = sourceData[parseInt(keys[i]) - 1];
    if (value !== '' && value !== null && value !== undefined) {
      accSheet.getRange(targetRow, map[keys[i]]).setValue(value);
    }
  }

  // Format auto-pull cols
  formatRowByType_(accSheet, targetRow, CONFIG.ACC_SHEET);

  if (existingRow === 0) {
    logPayment_('DM_TO_ACC', dmNo, 'DM → Account: new row');
    checkAndPullAccountToRTO_(accSheet, targetRow);
  } else {
    logPayment_('DM_TO_ACC_UPDATE', dmNo, 'DM → Account: updated');
    syncAccountToRTO_(accSheet, targetRow);
  }
}


// ═══════════════════════════════════════
// SECTION 4 — AUTO-PULL: ACCOUNT → RTO
// ═══════════════════════════════════════

function checkAndPullAccountToRTO_(accSheet, accRow) {
  var ss = SpreadsheetApp.getActive();
  var rtoSheet = ss.getSheetByName(CONFIG.RTO_SHEET);
  if (!rtoSheet) return;
  var dmNo = accSheet.getRange(accRow, CONFIG.ACC.DM_NO).getValue();
  if (!dmNo) return;

  var rtoLastRow = rtoSheet.getLastRow();
  var existingRow = 0;
  if (rtoLastRow > CONFIG.HEADER_ROW) {
    var rtoDmNos = rtoSheet.getRange(CONFIG.HEADER_ROW+1, CONFIG.RTO.DM_NO, rtoLastRow-CONFIG.HEADER_ROW, 1).getValues().flat();
    var idx = rtoDmNos.indexOf(dmNo);
    if (idx >= 0) existingRow = CONFIG.HEADER_ROW + 1 + idx;
  }

  var sourceData = accSheet.getRange(accRow, 1, 1, CONFIG.ACC.TOTAL_COLS).getValues()[0];
  var targetRow = existingRow > 0 ? existingRow : rtoSheet.getLastRow() + 1;

  var map = CONFIG.ACC_TO_RTO;
  var keys = Object.keys(map);
  for (var i = 0; i < keys.length; i++) {
    var value = sourceData[parseInt(keys[i]) - 1];
    if (value !== '' && value !== null && value !== undefined) {
      rtoSheet.getRange(targetRow, map[keys[i]]).setValue(value);
    }
  }

  var regNo = sourceData[CONFIG.ACC.REG_NO - 1];
  if (regNo) {
    var rtoCode = regNo.toString().replace(/[0-9]/g, '').substring(0, 4).toUpperCase();
    rtoSheet.getRange(targetRow, CONFIG.RTO.RTO_CODE).setValue(rtoCode);
  }
  var product = sourceData[CONFIG.ACC.PRODUCT - 1];
  if (product) rtoSheet.getRange(targetRow, CONFIG.RTO.CASE_TYPE).setValue(product);
  if (existingRow === 0) {
    rtoSheet.getRange(targetRow, CONFIG.RTO.SYSTEM_REMARKS).setValue('AUTO-TRANSFERRED FROM ACCOUNT');
    logPayment_('ACC_TO_RTO', dmNo, 'Account → RTO: new row');
  }

  // Format
  formatRowByType_(rtoSheet, targetRow, CONFIG.RTO_SHEET);
}

function syncAccountToRTO_(accSheet, accRow) { checkAndPullAccountToRTO_(accSheet, accRow); }
function syncDMToAccountAndRTO_(sheet, row) { checkAndPullDMToAccount_(sheet, row); }


// ═══════════════════════════════════════
// SECTION 5 — MASTER_DATA SYNC
// ═══════════════════════════════════════

function syncMasterData() {
  var ss = SpreadsheetApp.getActive();
  var dmSheet = ss.getSheetByName(CONFIG.DM_SHEET);
  var accSheet = ss.getSheetByName(CONFIG.ACC_SHEET);
  var rtoSheet = ss.getSheetByName(CONFIG.RTO_SHEET);
  var masterSheet = ss.getSheetByName(CONFIG.MASTER_DATA);
  if (!dmSheet || !accSheet || !rtoSheet || !masterSheet) return;

  var dmLastRow = dmSheet.getLastRow();
  if (dmLastRow <= CONFIG.HEADER_ROW) return;

  var dmData = dmSheet.getRange(CONFIG.HEADER_ROW + 1, 1, dmLastRow - CONFIG.HEADER_ROW, CONFIG.DM.TOTAL_COLS).getValues();

  var accLookup = {};
  var accLastRow = accSheet.getLastRow();
  if (accLastRow > CONFIG.HEADER_ROW) {
    var accData = accSheet.getRange(CONFIG.HEADER_ROW + 1, 1, accLastRow - CONFIG.HEADER_ROW, CONFIG.ACC.TOTAL_COLS).getValues();
    for (var a = 0; a < accData.length; a++) {
      if (accData[a][CONFIG.ACC.DM_NO - 1]) accLookup[accData[a][CONFIG.ACC.DM_NO - 1]] = accData[a];
    }
  }

  var rtoLookup = {};
  var rtoLastRow = rtoSheet.getLastRow();
  if (rtoLastRow > CONFIG.HEADER_ROW) {
    var rtoData = rtoSheet.getRange(CONFIG.HEADER_ROW + 1, 1, rtoLastRow - CONFIG.HEADER_ROW, CONFIG.RTO.TOTAL_COLS).getValues();
    for (var r = 0; r < rtoData.length; r++) {
      if (rtoData[r][CONFIG.RTO.DM_NO - 1]) rtoLookup[rtoData[r][CONFIG.RTO.DM_NO - 1]] = rtoData[r];
    }
  }

  var masterRows = [];
  for (var d = 0; d < dmData.length; d++) {
    var dmNo = dmData[d][CONFIG.DM.DM_NO - 1];
    if (!dmNo) continue;
    var row = [];
    for (var c = 0; c < CONFIG.DM.TOTAL_COLS; c++) row.push(dmData[d][c]);
    var accRow = accLookup[dmNo];
    for (var ae = 0; ae < CONFIG.ACC_EXTRA_COLS.length; ae++) row.push(accRow ? accRow[CONFIG.ACC_EXTRA_COLS[ae] - 1] : '');
    var rtoRow = rtoLookup[dmNo];
    for (var re = 0; re < CONFIG.RTO_EXTRA_COLS.length; re++) row.push(rtoRow ? rtoRow[CONFIG.RTO_EXTRA_COLS[re] - 1] : '');
    masterRows.push(row);
  }

  var masterLastRow = masterSheet.getLastRow();
  var masterLastCol = masterSheet.getLastColumn();
  if (masterLastRow > CONFIG.HEADER_ROW && masterLastCol > 0) {
    masterSheet.getRange(CONFIG.HEADER_ROW + 1, 1, masterLastRow - CONFIG.HEADER_ROW, masterLastCol).clearContent();
  }

  if (masterRows.length > 0) {
    masterSheet.getRange(CONFIG.HEADER_ROW + 1, 1, masterRows.length, masterRows[0].length).setValues(masterRows);
    masterSheet.getRange(CONFIG.HEADER_ROW + 1, 1, masterRows.length, masterRows[0].length)
      .setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  }
}


// ═══════════════════════════════════════
// SECTION 6 — BACKUP SYSTEM
// ═══════════════════════════════════════

function createBackupSheets() {
  var ss = SpreadsheetApp.getActive();
  var pairs = [[CONFIG.DM_SHEET, CONFIG.BACKUP_DM], [CONFIG.ACC_SHEET, CONFIG.BACKUP_ACC], [CONFIG.RTO_SHEET, CONFIG.BACKUP_RTO]];
  for (var p = 0; p < pairs.length; p++) {
    var oldBk = ss.getSheetByName(pairs[p][1]);
    if (oldBk) { ss.deleteSheet(oldBk); SpreadsheetApp.flush(); }
    var bkSheet = ss.insertSheet(pairs[p][1]);
    bkSheet.hideSheet();
    var srcSheet = ss.getSheetByName(pairs[p][0]);
    if (srcSheet && srcSheet.getLastColumn() > 0) {
      var numCols = srcSheet.getLastColumn();
      bkSheet.getRange(1, 1, CONFIG.HEADER_ROW, numCols).setValues(srcSheet.getRange(1, 1, CONFIG.HEADER_ROW, numCols).getValues());
    }
  }
  fullBackupSync();
  SpreadsheetApp.getActive().toast('Backup sheets ready!', 'Setup', 5);
}

function fullBackupSync() {
  var ss = SpreadsheetApp.getActive();
  var pairs = [[CONFIG.DM_SHEET, CONFIG.BACKUP_DM], [CONFIG.ACC_SHEET, CONFIG.BACKUP_ACC], [CONFIG.RTO_SHEET, CONFIG.BACKUP_RTO]];
  for (var p = 0; p < pairs.length; p++) {
    var srcSheet = ss.getSheetByName(pairs[p][0]);
    var bkSheet = ss.getSheetByName(pairs[p][1]);
    if (!srcSheet || !bkSheet) continue;
    var lastRow = srcSheet.getLastRow(); var lastCol = srcSheet.getLastColumn();
    if (lastRow <= CONFIG.HEADER_ROW || lastCol === 0) continue;
    var bkLR = bkSheet.getLastRow(); var bkLC = bkSheet.getLastColumn();
    if (bkLR > CONFIG.HEADER_ROW && bkLC > 0) bkSheet.getRange(CONFIG.HEADER_ROW+1, 1, bkLR-CONFIG.HEADER_ROW, bkLC).clearContent();
    var dataRows = lastRow - CONFIG.HEADER_ROW;
    if (dataRows > 0) bkSheet.getRange(CONFIG.HEADER_ROW+1, 1, dataRows, lastCol).setValues(srcSheet.getRange(CONFIG.HEADER_ROW+1, 1, dataRows, lastCol).getValues());
  }
  syncMasterData();
}


// ═══════════════════════════════════════
// SECTION 7 — DELETE PROTECTION & RESTORE
// ═══════════════════════════════════════

function checkAndRestore() {
  var ss = SpreadsheetApp.getActive();
  var pairs = [[CONFIG.DM_SHEET, CONFIG.BACKUP_DM, CONFIG.DM.DM_NO], [CONFIG.ACC_SHEET, CONFIG.BACKUP_ACC, CONFIG.ACC.DM_NO], [CONFIG.RTO_SHEET, CONFIG.BACKUP_RTO, CONFIG.RTO.DM_NO]];
  for (var p = 0; p < pairs.length; p++) {
    var srcSheet = ss.getSheetByName(pairs[p][0]);
    var bkSheet = ss.getSheetByName(pairs[p][1]);
    var dmCol = pairs[p][2];
    if (!srcSheet || !bkSheet) continue;
    var bkLastRow = bkSheet.getLastRow();
    if (bkLastRow <= CONFIG.HEADER_ROW) continue;
    var srcLastRow = srcSheet.getLastRow();
    var currentDmNos = srcLastRow > CONFIG.HEADER_ROW ? srcSheet.getRange(CONFIG.HEADER_ROW+1, dmCol, srcLastRow-CONFIG.HEADER_ROW, 1).getValues().flat() : [];
    var bkLC = bkSheet.getLastColumn();
    if (bkLC === 0) continue;
    var bkData = bkSheet.getRange(CONFIG.HEADER_ROW+1, 1, bkLastRow-CONFIG.HEADER_ROW, bkLC).getValues();
    for (var i = 0; i < bkData.length; i++) {
      var dmNo = bkData[i][dmCol - 1];
      if (dmNo && currentDmNos.indexOf(dmNo) === -1) {
        var restoreRow = srcSheet.getLastRow() + 1;
        srcSheet.getRange(restoreRow, 1, 1, bkData[i].length).setValues([bkData[i]]);
        formatRowByType_(srcSheet, restoreRow, pairs[p][0]);
        logAudit_('ROW_RESTORED', pairs[p][0], restoreRow, 'DM_NO: ' + dmNo);
      }
    }
  }
}

function onChangeHandler(e) {
  if (e.changeType === 'REMOVE_ROW') {
    checkAndRestore();
    logAudit_('ROW_DELETE_DETECTED', 'UNKNOWN', 0, 'Row deletion — by: ' + (Session.getActiveUser().getEmail() || 'UNKNOWN'));
  }
}


// ═══════════════════════════════════════
// SECTION 8 — FORMAT ENGINE (3 COLORS)
// Manual = Black Bold
// Auto-Pull = Deep Blue Bold Italic
// Auto-Generate = Dark Green Bold
// DM NO + REG NO = Size 12
// Email = lowercase
// ═══════════════════════════════════════

function formatCellStyle_(sheet, row, col, type, isSize12) {
  var cell = sheet.getRange(row, col);
  var fontSize = isSize12 ? 12 : 10;
  cell.setFontFamily('Arial').setFontSize(fontSize).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  if (type === 'auto_pull') {
    cell.setFontStyle('italic').setFontColor(CONFIG.COLOR_AUTO_PULL);
  } else if (type === 'auto_gen') {
    cell.setFontStyle('normal').setFontColor(CONFIG.COLOR_AUTO_GEN);
  } else {
    cell.setFontStyle('normal').setFontColor(CONFIG.COLOR_MANUAL);
  }
}

function formatRowByType_(sheet, row, sheetName) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;

  // Base format — all cells
  sheet.getRange(row, 1, 1, lastCol).setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setFontStyle('normal').setFontColor(CONFIG.COLOR_MANUAL)
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // Auto-pull cols — Deep Blue Italic
  var autoPullCols = [];
  if (sheetName === CONFIG.ACC_SHEET) autoPullCols = CONFIG.AUTO_PULL_ACC.concat(CONFIG.AUTO_PULL_ACC_RC);
  else if (sheetName === CONFIG.RTO_SHEET) autoPullCols = CONFIG.AUTO_PULL_RTO;
  for (var i = 0; i < autoPullCols.length; i++) {
    if (autoPullCols[i] <= lastCol) {
      sheet.getRange(row, autoPullCols[i]).setFontStyle('italic').setFontColor(CONFIG.COLOR_AUTO_PULL);
    }
  }

  // Auto-generate cols — Dark Green
  var autoGenCols = [];
  if (sheetName === CONFIG.DM_SHEET) autoGenCols = CONFIG.AUTO_GEN_DM;
  else if (sheetName === CONFIG.RTO_SHEET) autoGenCols = CONFIG.AUTO_GEN_RTO;
  for (var g = 0; g < autoGenCols.length; g++) {
    if (autoGenCols[g] <= lastCol) {
      sheet.getRange(row, autoGenCols[g]).setFontStyle('normal').setFontColor(CONFIG.COLOR_AUTO_GEN);
    }
  }

  // Size 12 cols — DM NO + REG NO
  var size12Cols = [];
  if (sheetName === CONFIG.DM_SHEET) size12Cols = CONFIG.SIZE_12_DM;
  else if (sheetName === CONFIG.ACC_SHEET) size12Cols = CONFIG.SIZE_12_ACC;
  else if (sheetName === CONFIG.RTO_SHEET) size12Cols = CONFIG.SIZE_12_RTO;
  for (var s = 0; s < size12Cols.length; s++) {
    if (size12Cols[s] <= lastCol) {
      sheet.getRange(row, size12Cols[s]).setFontSize(12);
    }
  }

  // DM NO — Dark Brown color (overrides auto-pull/auto-gen color)
  var dmNoCol = 0;
  if (sheetName === CONFIG.DM_SHEET) dmNoCol = CONFIG.DM_NO_COL_DM;
  else if (sheetName === CONFIG.ACC_SHEET) dmNoCol = CONFIG.DM_NO_COL_ACC;
  else if (sheetName === CONFIG.RTO_SHEET) dmNoCol = CONFIG.DM_NO_COL_RTO;
  if (dmNoCol > 0 && dmNoCol <= lastCol) {
    sheet.getRange(row, dmNoCol).setFontColor(CONFIG.COLOR_DM_NO).setFontSize(12).setFontWeight('bold');
  }

  // REG NO — Dark Orange color (overrides auto-pull/auto-gen color)
  var regNoCol = 0;
  if (sheetName === CONFIG.DM_SHEET) regNoCol = CONFIG.REG_NO_COL_DM;
  else if (sheetName === CONFIG.ACC_SHEET) regNoCol = CONFIG.REG_NO_COL_ACC;
  else if (sheetName === CONFIG.RTO_SHEET) regNoCol = CONFIG.REG_NO_COL_RTO;
  if (regNoCol > 0 && regNoCol <= lastCol) {
    sheet.getRange(row, regNoCol).setFontColor(CONFIG.COLOR_REG_NO).setFontSize(12).setFontWeight('bold');
  }

  // CAPS — all text uppercase (skip dropdown + email + lowercase cols)
  var capsSkip = [];
  var lowercaseCols = [];
  if (sheetName === CONFIG.DM_SHEET) { capsSkip = CONFIG.CAPS_SKIP_DM; lowercaseCols = CONFIG.LOWERCASE_DM; }
  else if (sheetName === CONFIG.ACC_SHEET) { capsSkip = CONFIG.CAPS_SKIP_ACC; lowercaseCols = CONFIG.LOWERCASE_ACC; }
  else if (sheetName === CONFIG.RTO_SHEET) { capsSkip = CONFIG.CAPS_SKIP_RTO; lowercaseCols = CONFIG.LOWERCASE_RTO; }

  for (var c = 1; c <= lastCol; c++) {
    var cell = sheet.getRange(row, c);
    var val = cell.getValue();
    if (val && typeof val === 'string' && val.trim()) {
      if (lowercaseCols.indexOf(c) >= 0) {
        // Email — always lowercase
        cell.setValue(val.toLowerCase());
      } else if (capsSkip.indexOf(c) === -1) {
        // CAPS — not dropdown
        cell.setValue(val.toUpperCase());
      }
    }
  }
}

function formatGuard_(e) {
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  if ([CONFIG.DM_SHEET, CONFIG.ACC_SHEET, CONFIG.RTO_SHEET].indexOf(sheetName) === -1) return;
  var row = e.range.getRow();
  var col = e.range.getColumn();
  if (row <= CONFIG.HEADER_ROW) return;

  // Base format
  e.range.setFontFamily('Arial').setFontSize(10).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  // Determine color type
  var autoPullCols = [];
  if (sheetName === CONFIG.ACC_SHEET) autoPullCols = CONFIG.AUTO_PULL_ACC.concat(CONFIG.AUTO_PULL_ACC_RC);
  else if (sheetName === CONFIG.RTO_SHEET) autoPullCols = CONFIG.AUTO_PULL_RTO;

  var autoGenCols = [];
  if (sheetName === CONFIG.DM_SHEET) autoGenCols = CONFIG.AUTO_GEN_DM;
  else if (sheetName === CONFIG.RTO_SHEET) autoGenCols = CONFIG.AUTO_GEN_RTO;

  if (autoPullCols.indexOf(col) >= 0) {
    e.range.setFontStyle('italic').setFontColor(CONFIG.COLOR_AUTO_PULL);
  } else if (autoGenCols.indexOf(col) >= 0) {
    e.range.setFontStyle('normal').setFontColor(CONFIG.COLOR_AUTO_GEN);
  } else {
    e.range.setFontStyle('normal').setFontColor(CONFIG.COLOR_MANUAL);
  }

  // Size 12
  var size12 = [];
  if (sheetName === CONFIG.DM_SHEET) size12 = CONFIG.SIZE_12_DM;
  else if (sheetName === CONFIG.ACC_SHEET) size12 = CONFIG.SIZE_12_ACC;
  else if (sheetName === CONFIG.RTO_SHEET) size12 = CONFIG.SIZE_12_RTO;
  if (size12.indexOf(col) >= 0) e.range.setFontSize(12);

  // DM NO — Dark Brown override
  var dmNoCol = 0;
  if (sheetName === CONFIG.DM_SHEET) dmNoCol = CONFIG.DM_NO_COL_DM;
  else if (sheetName === CONFIG.ACC_SHEET) dmNoCol = CONFIG.DM_NO_COL_ACC;
  else if (sheetName === CONFIG.RTO_SHEET) dmNoCol = CONFIG.DM_NO_COL_RTO;
  if (col === dmNoCol) e.range.setFontColor(CONFIG.COLOR_DM_NO).setFontSize(12);

  // REG NO — Dark Orange override
  var regNoCol = 0;
  if (sheetName === CONFIG.DM_SHEET) regNoCol = CONFIG.REG_NO_COL_DM;
  else if (sheetName === CONFIG.ACC_SHEET) regNoCol = CONFIG.REG_NO_COL_ACC;
  else if (sheetName === CONFIG.RTO_SHEET) regNoCol = CONFIG.REG_NO_COL_RTO;
  if (col === regNoCol) e.range.setFontColor(CONFIG.COLOR_REG_NO).setFontSize(12);

  // CAPS / lowercase / dropdown
  var capsSkip = [];
  var lowercaseCols = [];
  if (sheetName === CONFIG.DM_SHEET) { capsSkip = CONFIG.CAPS_SKIP_DM; lowercaseCols = CONFIG.LOWERCASE_DM; }
  else if (sheetName === CONFIG.ACC_SHEET) { capsSkip = CONFIG.CAPS_SKIP_ACC; lowercaseCols = CONFIG.LOWERCASE_ACC; }
  else if (sheetName === CONFIG.RTO_SHEET) { capsSkip = CONFIG.CAPS_SKIP_RTO; lowercaseCols = CONFIG.LOWERCASE_RTO; }

  var val = e.range.getValue();
  if (val && typeof val === 'string' && val.trim()) {
    if (lowercaseCols.indexOf(col) >= 0) {
      if (val !== val.toLowerCase()) e.range.setValue(val.toLowerCase());
    } else if (capsSkip.indexOf(col) === -1) {
      if (val !== val.toUpperCase()) e.range.setValue(val.toUpperCase());
    }
  }
}


// ═══════════════════════════════════════
// SECTION 9 — VALIDATION
// ═══════════════════════════════════════

function validateData_(e) {
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CONFIG.DM_SHEET) return;
  var col = e.range.getColumn(); var row = e.range.getRow();
  if (row <= CONFIG.HEADER_ROW) return;
  var value = e.range.getValue();

  if ((col === CONFIG.DM.PHONE || col === CONFIG.DM.PHONE2) && value) {
    if (!/^\d{10}$/.test(value.toString().trim())) {
      SpreadsheetApp.getActive().toast('Phone must be 10 digits!', 'Validation', 5);
      e.range.setBackground('#FDEAEA');
    } else { e.range.setBackground(null); }
  }
  if (col === CONFIG.DM.EMAIL && value) {
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value.toString().trim())) {
      SpreadsheetApp.getActive().toast('Invalid email format!', 'Validation', 5);
      e.range.setBackground('#FDEAEA');
    } else { e.range.setBackground(null); }
  }
  if (col === CONFIG.DM.REG_NO && value) {
    var clean = value.toString().toUpperCase().replace(/\s/g, '');
    if (clean !== value.toString()) e.range.setValue(clean);
  }
  if (col === CONFIG.DM.ROI && value) {
    var n = parseFloat(value);
    if (!isNaN(n)) { if (n > 1) e.range.setValue(n / 100); e.range.setNumberFormat('0.00%'); }
  }
  if (col === CONFIG.DM.LOGIN_DATE && value) {
    var date = new Date(value);
    if (!isNaN(date.getTime())) generateDMNo_(sheet, row);
  }
}


// ═══════════════════════════════════════
// SECTION 10 — LOGGING
// ═══════════════════════════════════════

function logAudit_(action, sheetName, row, detail) {
  try {
    var ss = SpreadsheetApp.getActive();
    var ls = ss.getSheetByName(CONFIG.AUDIT_LOG);
    if (!ls) { ls = ss.insertSheet(CONFIG.AUDIT_LOG); ls.appendRow(['TIMESTAMP','EMAIL','ACTION','SHEET','ROW','DETAIL']); }
    ls.appendRow([new Date(), Session.getActiveUser().getEmail() || 'SYSTEM', action, sheetName, row, detail]);
  } catch(e) { Logger.log('Audit error: ' + e.message); }
}

function logPayment_(action, dmNo, detail) {
  try {
    var ss = SpreadsheetApp.getActive();
    var ls = ss.getSheetByName(CONFIG.PAYMENT_LOG);
    if (!ls) { ls = ss.insertSheet(CONFIG.PAYMENT_LOG); ls.appendRow(['TIMESTAMP','USER','ACTION','DM_NO','DETAIL']); }
    ls.appendRow([new Date(), Session.getActiveUser().getEmail() || 'SYSTEM', action, dmNo, detail]);
  } catch(e) { Logger.log('Payment log error: ' + e.message); }
}


// ═══════════════════════════════════════
// SECTION 11 — ADMIN MENU
// ═══════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🔧 TCO Admin')
    .addItem('📅 Month View Filter', 'showMonthFilter')
    .addSeparator()
    .addItem('📊 Refresh All Data', 'adminRefreshData')
    .addItem('📋 Refresh Master Data', 'adminRefreshMaster')
    .addItem('💾 Run Full Backup', 'adminBackup')
    .addItem('🔍 System Health Check', 'adminHealthCheck')
    .addItem('🔄 Restore Deleted Rows', 'adminRestore')
    .addSeparator()
    .addItem('⚡ Setup Triggers', 'adminSetupTriggers')
    .addItem('🛠️ Create Backup Sheets', 'adminCreateBackup')
    .addToUi();
}

function isAdmin_() { return Session.getActiveUser().getEmail() === CONFIG.ADMIN_EMAIL; }
function adminRefreshData() { if(!isAdmin_()){SpreadsheetApp.getActive().toast('Access Denied','Error',5);return;} guardDMNoSeries(); fullBackupSync(); SpreadsheetApp.getActive().toast('Done!','Refresh',5); }
function adminRefreshMaster() { if(!isAdmin_()){SpreadsheetApp.getActive().toast('Access Denied','Error',5);return;} syncMasterData(); SpreadsheetApp.getActive().toast('Master synced!','Done',5); }
function adminBackup() { if(!isAdmin_()){SpreadsheetApp.getActive().toast('Access Denied','Error',5);return;} fullBackupSync(); SpreadsheetApp.getActive().toast('Backup done!','Done',5); }
function adminHealthCheck() { if(!isAdmin_()){SpreadsheetApp.getActive().toast('Access Denied','Error',5);return;} systemHealthCheck(); }
function adminRestore() { if(!isAdmin_()){SpreadsheetApp.getActive().toast('Access Denied','Error',5);return;} checkAndRestore(); SpreadsheetApp.getActive().toast('Restore done!','Done',5); }
function adminSetupTriggers() { if(!isAdmin_()){SpreadsheetApp.getActive().toast('Access Denied','Error',5);return;} setupTriggers(); }
function adminCreateBackup() { if(!isAdmin_()){SpreadsheetApp.getActive().toast('Access Denied','Error',5);return;} createBackupSheets(); }


// ═══════════════════════════════════════
// SECTION 12 — MONTH VIEW FILTER
// ═══════════════════════════════════════

function showMonthFilter() {
  var html = HtmlService.createHtmlOutput(
    '<style>body{font-family:Arial;padding:15px}h3{color:#1B3A5C}select,button{width:100%;padding:8px;font-size:14px;margin-bottom:8px;border-radius:4px}.apply{background:#1B5E20;color:white;border:none;cursor:pointer}.clear{background:#B71C1C;color:white;border:none;cursor:pointer}</style>' +
    '<h3>📅 Month Filter</h3><p>Filter will apply on active sheet:</p>' +
    '<select id="m"><option value="1">January</option><option value="2">February</option><option value="3" selected>March</option><option value="4">April</option><option value="5">May</option><option value="6">June</option><option value="7">July</option><option value="8">August</option><option value="9">September</option><option value="10">October</option><option value="11">November</option><option value="12">December</option></select>' +
    '<select id="y"><option value="2025">2025</option><option value="2026" selected>2026</option><option value="2027">2027</option></select>' +
    '<button class="apply" onclick="google.script.run.withSuccessHandler(function(){google.script.host.close()}).applyMonthFilter(+document.getElementById(\'m\').value,+document.getElementById(\'y\').value)">✅ Apply Filter</button>' +
    '<button class="clear" onclick="google.script.run.withSuccessHandler(function(){google.script.host.close()}).clearMonthFilter()">❌ Clear Filter</button>'
  ).setWidth(280).setHeight(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function applyMonthFilter(month, year) {
  var ss = SpreadsheetApp.getActive(); var sheet = ss.getActiveSheet(); var sheetName = sheet.getName();
  var filterCol = 0;
  if (sheetName === CONFIG.DM_SHEET) filterCol = CONFIG.DM.MONTH;
  else if (sheetName === CONFIG.ACC_SHEET) filterCol = CONFIG.ACC.DISBURSEMENT_DATE;
  else if (sheetName === CONFIG.RTO_SHEET) filterCol = CONFIG.RTO.DISBURSEMENT_DATE;
  else { SpreadsheetApp.getActive().toast('No filter available for this sheet','Info',5); return; }
  var lastRow = sheet.getLastRow(); var lastCol = sheet.getLastColumn();
  if (lastRow <= CONFIG.HEADER_ROW) return;
  if (sheet.getFilter()) sheet.getFilter().remove();
  var range = sheet.getRange(CONFIG.HEADER_ROW, 1, lastRow-CONFIG.HEADER_ROW+1, lastCol);
  range.createFilter().setColumnFilterCriteria(filterCol, SpreadsheetApp.newFilterCriteria().whenDateAfter(new Date(year,month-1,0)).whenDateBefore(new Date(year,month,1)).build());
  var months = ['','January','February','March','April','May','June','July','August','September','October','November','December'];
  SpreadsheetApp.getActive().toast(months[month]+' '+year+' filter applied!','Filter',5);
}

function clearMonthFilter() {
  var sheet = SpreadsheetApp.getActive().getActiveSheet();
  if (sheet.getFilter()) { sheet.getFilter().remove(); SpreadsheetApp.getActive().toast('Filter cleared!','Done',5); }
}


// ═══════════════════════════════════════
// SECTION 13 — STALE CASE ALERT
// ═══════════════════════════════════════

function staleCaseAlert() {
  var ss = SpreadsheetApp.getActive(); var sheet = ss.getSheetByName(CONFIG.DM_SHEET);
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  if (lastRow <= CONFIG.HEADER_ROW) return;
  var data = sheet.getRange(CONFIG.HEADER_ROW+1, 1, lastRow-CONFIG.HEADER_ROW, CONFIG.DM.TOTAL_COLS).getValues();
  var today = new Date(); var staleCases = [];
  for (var i = 0; i < data.length; i++) {
    var status = (data[i][CONFIG.DM.CASE_STATUS-1]||'').toString().trim().toUpperCase();
    var loginDate = new Date(data[i][CONFIG.DM.LOGIN_DATE-1]);
    var dmNo = data[i][CONFIG.DM.DM_NO-1];
    if ((status==='IN PROCESS'||status==='LOGIN') && loginDate && !isNaN(loginDate.getTime())) {
      var days = Math.floor((today-loginDate)/(86400000));
      if (days > CONFIG.STALE_DAYS) {
        staleCases.push({dmNo:dmNo, buyer:data[i][CONFIG.DM.BUYER_NAME-1], branch:data[i][CONFIG.DM.BRANCH-1], exec:data[i][CONFIG.DM.EXEC_NAME-1], days:days, status:status});
        sheet.getRange(CONFIG.HEADER_ROW+1+i, 1, 1, CONFIG.DM.TOTAL_COLS).setBackground('#FFF3E0');
      }
    }
  }
  if (staleCases.length > 0) {
    var html = '<h3>⚠️ TCO Stale Cases — '+staleCases.length+'</h3><table border="1" cellpadding="6" style="border-collapse:collapse;font-family:Arial;font-size:12px"><tr style="background:#1B3A5C;color:#C8A951"><th>DM NO</th><th>BUYER</th><th>BRANCH</th><th>EXEC</th><th>DAYS</th><th>STATUS</th></tr>';
    for (var j = 0; j < staleCases.length; j++) { var c = staleCases[j]; html += '<tr><td><b>'+c.dmNo+'</b></td><td>'+c.buyer+'</td><td>'+c.branch+'</td><td>'+c.exec+'</td><td style="color:'+(c.days>15?'#C0392B':'#E65100')+';font-weight:bold">'+c.days+'</td><td>'+c.status+'</td></tr>'; }
    html += '</table><p style="color:gray;font-size:11px">TCO Operations System</p>';
    GmailApp.sendEmail(CONFIG.ADMIN_EMAIL, '⚠️ TCO: '+staleCases.length+' cases '+CONFIG.STALE_DAYS+'+ days pending', '', {htmlBody:html});
  }
}


// ═══════════════════════════════════════
// SECTION 14 — HEALTH CHECK
// ═══════════════════════════════════════

function systemHealthCheck() {
  var ss = SpreadsheetApp.getActive(); var checks = [];
  var names = [CONFIG.DM_SHEET, CONFIG.ACC_SHEET, CONFIG.RTO_SHEET, CONFIG.MASTER_DATA, CONFIG.BACKUP_DM, CONFIG.BACKUP_ACC, CONFIG.BACKUP_RTO, CONFIG.AUDIT_LOG, CONFIG.PAYMENT_LOG];
  for (var i = 0; i < names.length; i++) checks.push((ss.getSheetByName(names[i])?'✅':'❌')+' '+names[i]);
  checks.push('⚡ Triggers: '+ScriptApp.getProjectTriggers().length+' (limit 20)');
  var dm=ss.getSheetByName(CONFIG.DM_SHEET); if(dm) checks.push('📊 DM: '+Math.max(dm.getLastRow()-CONFIG.HEADER_ROW,0)+' rows');
  var acc=ss.getSheetByName(CONFIG.ACC_SHEET); if(acc) checks.push('📊 Account: '+Math.max(acc.getLastRow()-CONFIG.HEADER_ROW,0)+' rows');
  var rto=ss.getSheetByName(CONFIG.RTO_SHEET); if(rto) checks.push('📊 RTO: '+Math.max(rto.getLastRow()-CONFIG.HEADER_ROW,0)+' rows');
  var master=ss.getSheetByName(CONFIG.MASTER_DATA); if(master) checks.push('📋 Master: '+Math.max(master.getLastRow()-CONFIG.HEADER_ROW,0)+' rows');
  var bk=ss.getSheetByName(CONFIG.BACKUP_DM); if(bk) checks.push('💾 Backup: '+Math.max(bk.getLastRow()-CONFIG.HEADER_ROW,0)+' rows');
  SpreadsheetApp.getUi().alert('🔍 SYSTEM HEALTH CHECK\n━━━━━━━━━━━━━━━━━━━━\n\n'+checks.join('\n')+'\n\n━━━━━━━━━━━━━━━━━━━━\nTCO Operations System');
}


// ═══════════════════════════════════════
// SECTION 15 — MASTER onEdit + TRIGGERS
// ═══════════════════════════════════════

function masterOnEdit(e) {
  try {
    var sheet = e.source.getActiveSheet(); var sheetName = sheet.getName();
    var row = e.range.getRow();
    if (row <= CONFIG.HEADER_ROW) return;

    if (sheetName === CONFIG.DM_SHEET) {
      validateData_(e);
      var dmNo = sheet.getRange(row, CONFIG.DM.DM_NO).getValue();
      var buyer = sheet.getRange(row, CONFIG.DM.BUYER_NAME).getValue();
      var phone = sheet.getRange(row, CONFIG.DM.PHONE).getValue();
      var regNo = sheet.getRange(row, CONFIG.DM.REG_NO).getValue();
      if (dmNo && buyer && phone && regNo) syncDMToAccountAndRTO_(sheet, row);
    }

    if (sheetName === CONFIG.ACC_SHEET) {
      var accDmNo = sheet.getRange(row, CONFIG.ACC.DM_NO).getValue();
      if (accDmNo) syncAccountToRTO_(sheet, row);
    }

    formatGuard_(e);

    if ([CONFIG.DM_SHEET, CONFIG.ACC_SHEET, CONFIG.RTO_SHEET].indexOf(sheetName) >= 0) {
      logAudit_('EDIT', sheetName, row, e.range.getA1Notation()+': '+(e.oldValue||'')+' → '+(e.value||''));
    }
  } catch(err) { Logger.log('masterOnEdit error: '+err.message); }
}

function setupTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) ScriptApp.deleteTrigger(triggers[i]);
  var ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('masterOnEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('onChangeHandler').forSpreadsheet(ss).onChange().create();
  ScriptApp.newTrigger('guardDMNoSeries').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('checkAndRestore').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('fullBackupSync').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('staleCaseAlert').timeBased().everyDays(1).atHour(9).create();
  SpreadsheetApp.getActive().toast('6 triggers installed!','Setup Complete',5);
}
