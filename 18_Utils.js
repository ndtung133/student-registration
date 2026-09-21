/**********************************************************************
 * HÀM TIỆN ÍCH
 **********************************************************************/

function rangeTouchesColumn_(range, column) {
  return range.getColumn() <= column && range.getLastColumn() >= column;
}

function rangesOverlapColumns_(range, firstColumn, lastColumn) {
  return range.getColumn() <= lastColumn && range.getLastColumn() >= firstColumn;
}

function rangeTouchesCell_(range, row, column) {
  return range.getRow() <= row && range.getLastRow() >= row &&
    range.getColumn() <= column && range.getLastColumn() >= column;
}

function confirm_(title, message) {
  const ui = SpreadsheetApp.getUi();
  return ui.alert(title, message, ui.ButtonSet.YES_NO) === ui.Button.YES;
}

function toast_(message, title, seconds) {
  try {
    getSpreadsheet_().toast(message, title || 'Vệ sinh lab', seconds || 5);
  } catch (err) {}
}

function setActiveSheetSafe_(ss, sheetName) {
  try {
    const sh = ss.getSheetByName(sheetName);
    if (sh) ss.setActiveSheet(sh);
  } catch (err) {}
}

function updateWorkflowBanner_(ss) {
  const phaseText = workflowPhaseText_(getWorkflowPhase_());
  const config = ss.getSheetByName(LAB.SHEETS.CONFIG);
  if (config) config.getRange('B5').setValue(phaseText);

  const home = ss.getSheetByName(LAB.SHEETS.HOME);
  if (home) home.getRange('B8').setValue(phaseText);

  const roster = ss.getSheetByName(LAB.SHEETS.STUDENTS);
  if (roster) {
    roster.getRange('A1').setValue('DANH SÁCH SINH VIÊN DÙNG CHO CẢ HỌC KỲ');
  }
}

function twoDigits_(value) {
  return String(value).padStart(2, '0');
}

