/**********************************************************************
 * NHẬT KÝ HỆ THỐNG
 **********************************************************************/

function ensureSystemLogSheet_(ss) {
  const sh = ss.getSheetByName(LAB.SHEETS.SYSTEM_LOG);
  ensureSheetSize_(sh, 30, 6);
  if (!normalizeLabel_(sh.getRange('A1').getValue())) {
    sh.getRange('A1:F1').setValues([[
      'Thời điểm', 'Mức', 'Hàm', 'Tuần', 'Thông báo', 'Chi tiết'
    ]]);
    sh.getRange('A1:F1')
      .setFontWeight('bold')
      .setBackground('#d9d9d9')
      .setHorizontalAlignment('center');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 175);
    sh.setColumnWidth(2, 80);
    sh.setColumnWidth(3, 220);
    sh.setColumnWidth(4, 100);
    sh.setColumnWidth(5, 320);
    sh.setColumnWidth(6, 500);
  }
  return sh;
}

function logSystem_(level, functionName, message, detail) {
  try {
    const ss = getSpreadsheet_();
    const sh = ss.getSheetByName(LAB.SHEETS.SYSTEM_LOG) || ss.insertSheet(LAB.SHEETS.SYSTEM_LOG);
    ensureSystemLogSheet_(ss);
    const monday = getRegisterWeekMonday_(ss);
    sh.appendRow([
      new Date(),
      level,
      functionName,
      monday ? weekKey_(monday) : '',
      message,
      String(detail || '').slice(0, 5000)
    ]);
    sh.getRange(sh.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm:ss');
    sh.getDataRange().setVerticalAlignment('middle').setWrap(true);

    if (sh.getLastRow() > LAB.MAX_LOG_ROWS + 1) {
      sh.deleteRows(2, sh.getLastRow() - LAB.MAX_LOG_ROWS - 1);
    }
  } catch (err) {
    console.error('Không ghi được nhật ký: ' + readableError_(err));
  }
}

function safeLogError_(functionName, err) {
  try {
    logSystem_('ERROR', functionName, readableError_(err), err && err.stack ? err.stack : '');
  } catch (ignored) {
    console.error(functionName + ': ' + readableError_(err));
  }
}

function readableError_(err) {
  if (!err) return 'Lỗi không xác định';
  return err.message ? String(err.message) : String(err);
}

