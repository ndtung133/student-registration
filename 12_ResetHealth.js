/**********************************************************************
 * RESET TUẦN VÀ KIỂM TRA SỨC KHỎE - CHỐNG CHẠY LẶP
 **********************************************************************/

function resetForWeek_(ss, targetMonday, source, force, targetMode, openAfterBuild) {
  const shouldOpen = openAfterBuild !== false;
  const validation = validateRoster_(ss, true);
  const appliedStudents = getAppliedStudents_(ss);
  const students = validation.valid
    ? validation.students
    : mergeRosterSnapshots_(appliedStudents, validation.eligibleStudents);
  const usedFallbackRoster = !validation.valid;
  const targetKey = weekKey_(targetMonday);
  const desiredMode = targetMode || getWeekMode_();
  const currentMonday = getRegisterWeekMonday_(ss);
  const currentKey = currentMonday ? weekKey_(currentMonday) : '';
  const currentRosterHash = getDocumentProperty_(LAB.PROP.ROSTER_HASH) || '';
  const latestRosterHash = rosterHash_(students);
  const reg = ss.getSheetByName(LAB.SHEETS.REGISTER);
  const rowCount = getRegisterRowCount_(reg);

  if (
    !force &&
    currentKey === targetKey &&
    currentRosterHash === latestRosterHash &&
    rowCount === students.length &&
    getWeekMode_() === desiredMode
  ) {
    repairSystemWithoutRebuild_(ss, students);
    if (shouldOpen) openPreparedRegistration_(ss, targetMonday, students.length);
    if (usedFallbackRoster) {
      setDocumentProperty_(LAB.PROP.PENDING_ROSTER_HASH, validation.hash);
    }
    logSystem_('INFO', 'resetForWeek', 'Bỏ qua reset lặp vì tuần đã được tạo', targetKey + '; nguồn=' + source);
    return false;
  }

  archiveCurrentWeek_(ss, 'Lưu tự động trước reset - ' + source, false);
  beginWeekBuild_(ss);
  setWeekMode_(desiredMode);
  setAssignmentMode_(LAB.ASSIGN_SELF);
  buildCompleteWeek_(ss, targetMonday, students, source);
  if (shouldOpen && !openPreparedRegistration_(ss, targetMonday, students.length)) {
    throw new Error('Không thể mở đăng ký vì lịch vừa tạo chưa khớp danh sách sinh viên.');
  }
  if (usedFallbackRoster) {
    setDocumentProperty_(LAB.PROP.PENDING_ROSTER_HASH, validation.hash);
    logSystem_(
      'WARN',
      'resetForWeek',
      'Danh sách học kỳ có lỗi; lịch mới dùng bản hợp lệ gần nhất',
      validation.message + '; số SV áp dụng=' + students.length
    );
  }
  setDocumentProperty_(LAB.PROP.LAST_RESET_WEEK, targetKey);
  logSystem_('INFO', 'resetForWeek', 'Reset tuần thành công', targetKey + '; nguồn=' + source);
  return true;
}

function healthCheckInternal_(ss, showToast) {
  ss.setSpreadsheetTimeZone(LAB.TIMEZONE);
  ensureAllSheets_(ss);
  bootstrapAdmins_(ss);

  const validation = validateRoster_(ss, true);
  if (!validation.valid) {
    setDocumentProperty_(LAB.PROP.PENDING_ROSTER_HASH, validation.hash);
    logSystem_('ERROR', 'healthCheckInternal', 'Danh sách SV không hợp lệ', validation.message);
  }

  const liveStudents = validation.valid ? validation.students : validation.eligibleStudents;
  const now = new Date();
  const expectedMonday = getMonday_(now);
  const currentMonday = getRegisterWeekMonday_(ss);
  const localDay = getLocalWeekday_(now);

  if (!currentMonday || (localDay !== 7 && weekKey_(currentMonday) !== weekKey_(expectedMonday))) {
    resetForWeek_(ss, expectedMonday, 'HEALTH_RECOVERY', false, LAB.MODE_REGULAR);
    ensureManagedTriggers_();
    if (showToast) {
      toast_(
        validation.valid
          ? 'Đã phục hồi lịch đúng tuần hiện tại.'
          : 'Đã phục hồi lịch bằng danh sách hợp lệ gần nhất; các ô đỏ cần cán bộ kiểm tra.',
        'Tự sửa hoàn tất',
        8
      );
    }
    return;
  }

  const latestHash = validation.hash;
  const appliedHash = getDocumentProperty_(LAB.PROP.ROSTER_HASH) || '';
  const appliedStudents = getCurrentWeekStudents_(ss);
  const reg = ss.getSheetByName(LAB.SHEETS.REGISTER);
  const rowCount = getRegisterRowCount_(reg);
  const countMismatch = rowCount !== appliedStudents.length;
  const rosterChanged = latestHash !== appliedHash;

  if (countMismatch && !hasWeekActivity_(ss)) {
    rebuildWeekThenOpen_(ss, currentMonday, appliedStudents, 'HEALTH_SLOT_REPAIR');
    if (rosterChanged) setDocumentProperty_(LAB.PROP.PENDING_ROSTER_HASH, latestHash);
    logSystem_('INFO', 'healthCheckInternal', 'Tự sửa số slot theo danh sách đang áp dụng', 'Số SV=' + appliedStudents.length);
  } else {
    if (rosterChanged) setDocumentProperty_(LAB.PROP.PENDING_ROSTER_HASH, latestHash);
    repairSystemWithoutRebuild_(ss, appliedStudents.length || hasAppliedRosterSnapshot_(ss)
      ? appliedStudents
      : liveStudents);
  }

  ensureManagedTriggers_();
  if (showToast) {
    toast_(
      !validation.valid
        ? 'Lịch đang chạy vẫn được giữ ổn định; danh sách có ô đỏ cần sửa trước Chủ nhật.'
        : rosterChanged
          ? 'Hệ thống ổn; thay đổi danh sách sẽ tự áp dụng Chủ nhật 16:00.'
          : countMismatch
            ? 'Đã kiểm tra và phục hồi cấu trúc lịch tuần hiện tại.'
            : 'Trigger, trạng thái, validation, khóa và số slot đều đã được kiểm tra.',
      'Kiểm tra hoàn tất',
      8
    );
  }
}

function repairSystemWithoutRebuild_(ss, students) {
  const reg = ss.getSheetByName(LAB.SHEETS.REGISTER);
  const rowCount = getRegisterRowCount_(reg);
  const hasApplied = hasAppliedRosterSnapshot_(ss);
  const appliedStudents = getAppliedStudents_(ss);
  const effectiveStudents = hasApplied ? appliedStudents : (students || []);

  if (!hasApplied && rowCount > 0 && effectiveStudents.length === rowCount) {
    refreshSystemStudentList_(ss, effectiveStudents);
  }

  if (rowCount > 0) {
    applyRegisterValidation_(ss, rowCount, effectiveStudents.length);
    repairRegisterStatuses_(reg, rowCount);
    applyRegisterConditionalFormatting_(reg, rowCount);
  }

  buildConfigSheet_(ss, effectiveStudents);
  buildDashboardSheet_(ss);
  buildHomeSheet_(ss, effectiveStudents);
  buildRulesSheet_(ss, effectiveStudents);
  ensureWeeklyHistorySheet_(ss);
  ensureGeneralHistorySheet_(ss);
  ensureSystemLogSheet_(ss);

  const general = ss.getSheetByName(LAB.SHEETS.GENERAL_CLEANING);
  const generalCount = getGeneralRowCount_(general);
  if (getWeekMode_() === LAB.MODE_GENERAL_CLEANING) {
    if (generalCount === 0 && !hasAnyCompletion_(ss)) {
      const monday = getRegisterWeekMonday_(ss) || getExpectedOpenMonday_(new Date());
      buildGeneralCleaningSheet_(ss, monday, effectiveStudents);
    } else if (generalCount > 0) {
      general.showSheet();
      applyGeneralCleaningValidation_(ss, generalCount, effectiveStudents.length);
    }
  } else {
    hideGeneralCleaningSheet_(ss);
  }

  updateStudentPolicyGuides_(ss);
  refreshAllStatuses_(ss, false);
  applyAllProtections_(ss);
  SpreadsheetApp.flush();
}

function repairRegisterStatuses_(sh, rowCount) {
  if (!sh || rowCount <= 0) return;
  const assignmentMode = getAssignmentMode_();
  const width = LAB.COL.DA_TRUC - LAB.COL.SV + 1;
  const rows = sh.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.SV, rowCount, width).getValues();
  const values = rows.map(function (row) {
    const student = row[0];
    const done = row[LAB.COL.DA_TRUC - LAB.COL.SV] === true;
    return [registrationStatusText_(student, done, assignmentMode)];
  });
  sh.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.TRANGTHAI, rowCount, 1).setValues(values);
}

function ensureManagedTriggers_() {
  const expected = {};
  expected[LAB.TRIGGER.EDIT] = 1;
  expected[LAB.TRIGGER.PREPARE] = 1;
  expected[LAB.TRIGGER.OPEN] = 1;
  expected[LAB.TRIGGER.RECOVERY] = 1;
  expected[LAB.TRIGGER.HEALTH] = 1;
  expected[LAB.TRIGGER.SYNC] = 1;
  const counts = {};
  let hasLegacyReset = false;

  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    const name = trigger.getHandlerFunction();
    if (name === 'weeklyResetLabCleaning') hasLegacyReset = true;
    if (Object.prototype.hasOwnProperty.call(expected, name)) {
      counts[name] = (counts[name] || 0) + 1;
    }
  });

  const invalid = Object.keys(expected).some(function (name) {
    return (counts[name] || 0) !== expected[name];
  });
  if (invalid || hasLegacyReset) createLabCleaningTriggersInternal_(getSpreadsheet_());
}

