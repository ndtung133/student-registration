/**********************************************************************
 * DỰNG TOÀN BỘ MỘT TUẦN
 **********************************************************************/

function beginWeekBuild_(ss) {
  setWorkflowPhase_(LAB.PHASE_CLOSED);
  setRegistrationFullLocked_(false);
  applyRegistrationAccessProtections_(ss, true);

  const reg = ss.getSheetByName(LAB.SHEETS.REGISTER);
  if (reg) {
    reg.getRange('C2:N2').merge()
      .setValue('HỆ THỐNG ĐANG CHUẨN BỊ LỊCH TUẦN MỚI - VUI LÒNG CHỜ THÔNG BÁO MỞ ĐĂNG KÝ')
      .setBackground('#cfe2f3')
      .setFontColor('#073763')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
  updateWorkflowBanner_(ss);
  SpreadsheetApp.flush();
}

function rebuildWeekThenOpen_(ss, monday, students, source) {
  beginWeekBuild_(ss);
  buildCompleteWeek_(ss, monday, students, source);
  if (!openPreparedRegistration_(ss, monday, students.length)) {
    throw new Error('Đã dựng lịch nhưng chưa thể mở đăng ký vì số slot không khớp danh sách.');
  }
}

function openPreparedRegistration_(ss, targetMonday, expectedStudentCount) {
  const reg = ss.getSheetByName(LAB.SHEETS.REGISTER);
  if (!reg) return false;
  const currentMonday = getRegisterWeekMonday_(ss);
  if (
    targetMonday &&
    (!currentMonday || weekKey_(currentMonday) !== weekKey_(targetMonday))
  ) return false;

  const expected = typeof expectedStudentCount === 'number'
    ? expectedStudentCount
    : getCurrentWeekStudents_(ss).length;
  const rowCount = getRegisterRowCount_(reg);
  if (rowCount !== expected) return false;

  setRegistrationFullLocked_(false);
  setWorkflowPhase_(expected > 0 ? LAB.PHASE_REGISTRATION_OPEN : LAB.PHASE_CLOSED);
  applyRegistrationAccessProtections_(ss, false);
  refreshRegistrationStateFast_(ss);
  updateWorkflowBanner_(ss);
  SpreadsheetApp.flush();
  return true;
}

function buildCompleteWeek_(ss, monday, students, source) {
  runSelfTests_();
  ensureAllSheets_(ss);
  refreshSystemStudentList_(ss, students, monday);

  buildConfigSheet_(ss, students);
  buildRegisterSheet_(ss, monday, students);
  buildDashboardSheet_(ss);
  buildHomeSheet_(ss, students);
  buildRulesSheet_(ss, students);

  if (getWeekMode_() === LAB.MODE_GENERAL_CLEANING) {
    buildGeneralCleaningSheet_(ss, monday, students);
  } else {
    hideGeneralCleaningSheet_(ss);
  }

  ensureWeeklyHistorySheet_(ss);
  ensureGeneralHistorySheet_(ss);
  ensureSystemLogSheet_(ss);

  setDocumentProperty_(LAB.PROP.ROSTER_HASH, rosterHash_(students));
  setDocumentProperty_(LAB.PROP.PENDING_ROSTER_HASH, '');
  setDocumentProperty_(LAB.PROP.LAST_RESET_WEEK, weekKey_(monday));

  refreshAllStatuses_(ss, false);
  applyAllProtections_(ss);
  SpreadsheetApp.flush();

  logSystem_('INFO', 'buildCompleteWeek',
    'Đã dựng lịch tuần',
    'Nguồn=' + source + '; tuần=' + weekKey_(monday) + '; SV=' + students.length +
      '; chế độ=' + getWeekMode_() + '; phân công=' + getAssignmentMode_());
}

function ensureAllSheets_(ss) {
  Object.keys(LAB.SHEETS).forEach(function (key) {
    const name = LAB.SHEETS[key];
    if (!ss.getSheetByName(name)) ss.insertSheet(name);
  });
}

function ensureSheetSize_(sheet, requiredRows, requiredColumns) {
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }
  if (sheet.getMaxColumns() < requiredColumns) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredColumns - sheet.getMaxColumns());
  }
}

function resetSheet_(sheet) {
  if (!sheet) return;
  removeManagedProtectionsFromSheet_(sheet);
  try {
    const filter = sheet.getFilter();
    if (filter) filter.remove();
  } catch (err) {}
  try {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
  } catch (err) {}
  sheet.clear();
  sheet.clearConditionalFormatRules();
}

