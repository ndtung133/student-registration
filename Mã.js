/**********************************************************************
 * VIỆC LÀM VÀ VỊ TRÍ TRỰC THƯỜNG - HỖ TRỢ MỌI SỐ NGƯỜI/PHÒNG
 **********************************************************************/

function getPositionNameByLab_(labName, position, count) {
  if (labName === 'Phòng phụ') {
    if (count === 1) return 'Vị trí 1 - Lau + quét/lau phòng phụ';
    if (position === count) return 'Vị trí ' + position + ' - Quét/lau sàn phòng phụ';
    return 'Vị trí ' + position + ' - Lau khu vực phòng phụ ' + position;
  }

  if (count === 1) return 'Vị trí 1 - Lau + quét/lau phòng';
  if (position === count) return 'Vị trí ' + position + ' - Quét/lau sàn phòng';
  return 'Vị trí ' + position + ' - Lau ' + cleaningZoneName_(position, count - 1);
}

function getTaskByLab_(labName, position, count) {
  if (labName === 'Phòng phụ') {
    if (count === 1) return 'Phòng phụ - Lau + quét/lau toàn bộ phòng.';
    if (position === count) return 'Phòng phụ - Quét/lau sàn phòng.';
    return 'Phòng phụ - Lau ' + cleaningZoneName_(position, count - 1) + '.';
  }

  if (count === 1) return labName + ' - Lau + quét/lau toàn bộ phòng.';
  if (position === count) return labName + ' - Quét/lau sàn phòng.';
  return labName + ' - Lau ' + cleaningZoneName_(position, count - 1) + '.';
}

function cleaningZoneName_(position, cleaningCount) {
  if (cleaningCount === 1) return 'bàn và các bề mặt trong phòng';
  if (cleaningCount === 2) return position === 1 ? 'khu vực bên trái' : 'khu vực bên phải';
  if (cleaningCount === 3) {
    return ['khu vực bên trái', 'khu vực ở giữa', 'khu vực bên phải'][position - 1];
  }
  return 'khu vực số ' + position + '/' + cleaningCount;
}

function applyAlternatingDayColors_(sh, rowCount) {
  if (rowCount <= 0) return;
  const days = sh.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.THU, rowCount, 1).getValues();
  const backgrounds = [];
  let previous = null;
  let colorIndex = 0;

  days.forEach(function (row) {
    if (row[0] !== previous) {
      previous = row[0];
      colorIndex = 1 - colorIndex;
    }
    backgrounds.push(Array(14).fill(colorIndex ? '#f3f6f4' : '#ffffff'));
  });

  sh.getRange(LAB.FIRST_REGISTER_ROW, 1, rowCount, 14).setBackgrounds(backgrounds);
  sh.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.SV, rowCount, 1).setBackground('#fff2cc');
  sh.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.DA_TRUC, rowCount, 1).setBackground('#d9ead3');
}

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

/**********************************************************************
 * LƯU LỊCH SỬ CHI TIẾT TRƯỚC KHI RESET
 **********************************************************************/

function ensureWeeklyHistorySheet_(ss) {
  const sh = ss.getSheetByName(LAB.SHEETS.WEEK_HISTORY);
  ensureSheetSize_(sh, 30, 20);
  if (normalizeLabel_(sh.getRange('A1').getValue())) return sh;

  sh.getRange(1, 1, 1, 20).setValues([[
    'Mã lưu', 'Loại lịch', 'Tuần bắt đầu', 'Chế độ tuần', 'STT',
    'Thứ', 'Ngày', 'Khu vực', 'Công việc', 'Vị trí/Yêu cầu',
    'Sinh viên', 'Thời điểm đăng ký/phân công', 'Ghi chú SV',
    'Trạng thái', 'Đã xong', 'Người kiểm tra', 'Thời điểm kiểm tra',
    'Ghi chú kiểm tra', 'Thời điểm lưu', 'Lý do lưu'
  ]]);
  sh.getRange(1, 1, 1, 20)
    .setFontWeight('bold')
    .setBackground('#cccccc')
    .setHorizontalAlignment('center')
    .setWrap(true);
  sh.setFrozenRows(1);
  return sh;
}

function archiveCurrentWeek_(ss, reason, snapshot) {
  const reg = ss.getSheetByName(LAB.SHEETS.REGISTER);
  const rowCount = getRegisterRowCount_(reg);
  const monday = getRegisterWeekMonday_(ss);
  if (!reg || rowCount === 0 || !monday) return false;

  SpreadsheetApp.flush();
  const history = ss.getSheetByName(LAB.SHEETS.WEEK_HISTORY) || ss.insertSheet(LAB.SHEETS.WEEK_HISTORY);
  ensureWeeklyHistorySheet_(ss);

  const baseKey = weekKey_(monday) + '|' + getWeekMode_();
  const archiveKey = snapshot
    ? baseKey + '|SNAPSHOT|' + timestampKey_(new Date())
    : baseKey + '|FINAL';

  if (!snapshot && historyContainsArchiveKey_(history, archiveKey)) return false;

  const savedAt = new Date();
  const rows = [];
  const data = reg.getRange(LAB.FIRST_REGISTER_ROW, 1, rowCount, 14).getValues();
  data.forEach(function (row) {
    const student = normalizeLabel_(row[LAB.COL.SV - 1]);
    const done = row[LAB.COL.DA_TRUC - 1] === true;
    const status = !student
      ? 'Còn trống'
      : done
        ? 'Đã trực xong'
        : getAssignmentMode_() === LAB.ASSIGN_AUTO
          ? 'Đã phân công - chưa trực'
          : 'Đã đăng ký - chưa trực';
    rows.push([
      archiveKey,
      'Trực thường',
      monday,
      weekModeText_(getWeekMode_()),
      row[LAB.COL.STT - 1],
      row[LAB.COL.THU - 1],
      row[LAB.COL.NGAY - 1],
      row[LAB.COL.LAB - 1],
      row[LAB.COL.VIEC - 1],
      row[LAB.COL.VITRI - 1],
      student,
      row[LAB.COL.TIME_DANGKY - 1],
      row[LAB.COL.GHICHU_SV - 1],
      status,
      done,
      row[LAB.COL.NGUOI_KIEMTRA - 1],
      row[LAB.COL.TIME_KIEMTRA - 1],
      row[LAB.COL.GHICHU_KIEMTRA - 1],
      savedAt,
      reason
    ]);
  });

  if (getWeekMode_() === LAB.MODE_GENERAL_CLEANING) {
    const general = ss.getSheetByName(LAB.SHEETS.GENERAL_CLEANING);
    const generalCount = getGeneralRowCount_(general);
    if (generalCount > 0) {
      const targetDate = addDays_(monday, 4);
      const generalData = general.getRange(LAB.FIRST_GENERAL_ROW, 1, generalCount, 10).getValues();
      generalData.forEach(function (row) {
        const student = normalizeLabel_(row[LAB.GENERAL_COL.SV - 1]);
        const done = row[LAB.GENERAL_COL.DONE - 1] === true;
        rows.push([
          archiveKey,
          'Tổng vệ sinh',
          monday,
          weekModeText_(getWeekMode_()),
          row[LAB.GENERAL_COL.STT - 1],
          'Thứ 6',
          targetDate,
          row[LAB.GENERAL_COL.AREA - 1],
          row[LAB.GENERAL_COL.JOB - 1],
          row[LAB.GENERAL_COL.REQUIREMENT - 1],
          student,
          '',
          '',
          !student ? 'Còn trống' : done ? 'Đã hoàn thành' : 'Chưa hoàn thành',
          done,
          row[LAB.GENERAL_COL.INSPECTOR - 1],
          row[LAB.GENERAL_COL.CHECK_TIME - 1],
          row[LAB.GENERAL_COL.NOTE - 1],
          savedAt,
          reason
        ]);
      });
    }
  }

  if (rows.length > 0) {
    ensureSheetSize_(history, history.getLastRow() + rows.length + 5, 20);
    history.getRange(history.getLastRow() + 1, 1, rows.length, 20).setValues(rows);
    history.getRange(2, 3, history.getLastRow() - 1, 1).setNumberFormat('dd/MM/yyyy');
    history.getRange(2, 7, history.getLastRow() - 1, 1).setNumberFormat('dd/MM/yyyy');
    history.getRange(2, 12, history.getLastRow() - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    history.getRange(2, 17, history.getLastRow() - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    history.getRange(2, 19, history.getLastRow() - 1, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    history.getDataRange().setVerticalAlignment('middle').setWrap(true);
  }

  return rows.length > 0;
}

function historyContainsArchiveKey_(history, archiveKey) {
  if (history.getLastRow() < 2) return false;
  return history.getRange(2, 1, history.getLastRow() - 1, 1)
    .getValues()
    .some(function (row) { return String(row[0]) === archiveKey; });
}

/**********************************************************************
 * LỊCH SỬ TỔNG VỆ SINH - UPSERT, KHÔNG APPEND TRÙNG
 **********************************************************************/

function ensureGeneralHistorySheet_(ss) {
  const sh = ss.getSheetByName(LAB.SHEETS.GENERAL_HISTORY);
  ensureSheetSize_(sh, 30, 4);
  if (!normalizeLabel_(sh.getRange('A1').getValue())) {
    sh.getRange('A1:D1').setValues([[
      'Ngày tổng vệ sinh', 'Giờ', 'Thời điểm tạo/cập nhật bảng', 'Ghi chú'
    ]]);
    sh.getRange('A1:D1')
      .setFontWeight('bold')
      .setBackground('#cccccc')
      .setHorizontalAlignment('center');
  }
  dedupeGeneralHistory_(sh);
  return sh;
}

function dedupeGeneralHistory_(sh) {
  if (sh.getLastRow() < 3) return;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  const seen = {};
  const unique = [];
  values.forEach(function (row) {
    const key = historyDateKey_(row[0]);
    if (!key || seen[key]) return;
    seen[key] = true;
    unique.push(row);
  });

  if (unique.length !== values.length) {
    sh.getRange(2, 1, values.length, 4).clearContent();
    if (unique.length > 0) sh.getRange(2, 1, unique.length, 4).setValues(unique);
    logSystem_('WARN', 'dedupeGeneralHistory', 'Đã loại bản ghi tổng vệ sinh trùng', 'Xóa ' + (values.length - unique.length) + ' dòng trùng');
  }
}

function upsertGeneralCleaningHistory_(ss, targetDate) {
  const sh = ensureGeneralHistorySheet_(ss);
  const targetKey = historyDateKey_(targetDate);
  let existingRow = 0;

  if (sh.getLastRow() >= 2) {
    const dates = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    dates.some(function (row, index) {
      if (historyDateKey_(row[0]) === targetKey) {
        existingRow = index + 2;
        return true;
      }
      return false;
    });
  }

  const values = [[
    targetDate,
    '14:00',
    new Date(),
    getWeekMode_() === LAB.MODE_GENERAL_CLEANING
      ? 'Tuần có tổng vệ sinh Thứ 6 - ' + assignmentModeText_(getAssignmentMode_())
      : 'Tạo thủ công - ' + assignmentModeText_(getAssignmentMode_())
  ]];

  const row = existingRow || sh.getLastRow() + 1;
  sh.getRange(row, 1, 1, 4).setValues(values);
  sh.getRange(row, 1).setNumberFormat('dd/MM/yyyy');
  sh.getRange(row, 3).setNumberFormat('dd/MM/yyyy HH:mm');
  sh.getRange('A:D').setVerticalAlignment('middle').setWrap(true);
  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 80);
  sh.setColumnWidth(3, 195);
  sh.setColumnWidth(4, 340);
}

function historyDateKey_(value) {
  if (isValidDate_(value)) return Utilities.formatDate(value, LAB.TIMEZONE, 'yyyy-MM-dd');
  const text = normalizeLabel_(value);
  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (match) return match[3] + '-' + twoDigits_(match[2]) + '-' + twoDigits_(match[1]);
  return text;
}

/**********************************************************************
 * BẢO VỆ SHEET - SV CHỈ ĐƯỢC SỬA ĐÚNG VÙNG CẦN THIẾT
 **********************************************************************/

function applyAllProtections_(ss) {
  applyManagedSheetBaseProtection_(ss.getSheetByName(LAB.SHEETS.HOME), [], true);
  applyManagedSheetBaseProtection_(ss.getSheetByName(LAB.SHEETS.CONFIG), [], true);
  applyRosterProtectionsOnly_(ss, true);
  applyManagedSheetBaseProtection_(ss.getSheetByName(LAB.SHEETS.DASHBOARD), [], true);
  applyManagedSheetBaseProtection_(ss.getSheetByName(LAB.SHEETS.RULES), [], true);
  applyManagedSheetBaseProtection_(ss.getSheetByName(LAB.SHEETS.GENERAL_HISTORY), [], true);
  applyManagedSheetBaseProtection_(ss.getSheetByName(LAB.SHEETS.WEEK_HISTORY), [], true);
  applyManagedSheetBaseProtection_(ss.getSheetByName(LAB.SHEETS.SYSTEM_LOG), [], true);
  applyManagedSheetBaseProtection_(ss.getSheetByName(LAB.SHEETS.SYSTEM), [], true);
  applyRegistrationAccessProtections_(ss, true);

  try { ss.getSheetByName(LAB.SHEETS.SYSTEM).hideSheet(); } catch (err) {}
}

function applyRegistrationAccessProtections_(ss, cleanupLegacyOwners) {
  const phase = getWorkflowPhase_();
  const full = getDocumentProperty_(LAB.PROP.FULL_LOCK) === 'YES';
  const reg = ss.getSheetByName(LAB.SHEETS.REGISTER);
  const regCount = getRegisterRowCount_(reg);
  const registerOpenRanges = [];

  if (phase === LAB.PHASE_REGISTRATION_OPEN && !full && reg && regCount > 0) {
    registerOpenRanges.push(reg.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.SV, regCount, 1));
    registerOpenRanges.push(reg.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.GHICHU_SV, regCount, 1));
  }
  applyManagedSheetBaseProtection_(reg, registerOpenRanges, cleanupLegacyOwners);

  const general = ss.getSheetByName(LAB.SHEETS.GENERAL_CLEANING);
  const generalCount = getGeneralRowCount_(general);
  const generalOpenRanges = [];
  const generalMode = getWeekMode_() === LAB.MODE_GENERAL_CLEANING;
  if (
    phase === LAB.PHASE_REGISTRATION_OPEN &&
    !full &&
    generalCount > 0 &&
    generalMode
  ) {
    generalOpenRanges.push(general.getRange(LAB.FIRST_GENERAL_ROW, LAB.GENERAL_COL.SV, generalCount, 1));
  }
  if (generalMode || cleanupLegacyOwners) {
    applyManagedSheetBaseProtection_(general, generalOpenRanges, cleanupLegacyOwners);
  }

  if (cleanupLegacyOwners) {
    removeManagedRangeProtectionsByPrefix_(reg, LAB.REGISTER_ROW_PREFIX);
    removeManagedRangeProtectionsByPrefix_(general, LAB.GENERAL_ROW_PREFIX);
  }
}

function removeManagedRangeProtectionsByPrefix_(sheet, prefix) {
  if (!sheet) return;
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (protection) {
    const description = protection.getDescription() || '';
    if (description.indexOf(prefix) === 0) {
      try { protection.remove(); } catch (err) {}
    }
  });
}

function applyRosterProtectionsOnly_(ss, cleanupLegacyOwners) {
  const sh = ss.getSheetByName(LAB.SHEETS.STUDENTS);
  if (!sh) return;
  const inputRows = sh.getMaxRows() - LAB.FIRST_STUDENT_ROW + 1;
  const openRanges = inputRows > 0
    ? [sh.getRange(LAB.FIRST_STUDENT_ROW, 2, inputRows, 1)]
    : [];
  applyManagedSheetBaseProtection_(sh, openRanges, cleanupLegacyOwners);
  if (cleanupLegacyOwners) {
    removeManagedRangeProtectionsByPrefix_(sh, LAB.ROSTER_ROW_PREFIX);
  }
}

function applyManagedSheetBaseProtection_(sheet, unprotectedRanges, refreshEditors) {
  if (!sheet) return;
  const managed = sheet
    .getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .filter(function (protection) {
      const description = protection.getDescription() || '';
      return description.indexOf(LAB.PROTECTION_PREFIX) === 0 ||
        description === LAB.LEGACY_LOCK_DESC;
    });
  let protection = managed.shift();
  const created = !protection;
  if (!protection) protection = sheet.protect();
  managed.forEach(function (duplicate) {
    try { duplicate.remove(); } catch (err) {}
  });
  if (created || refreshEditors) {
    configureManagedSheetProtection_(protection, sheet, unprotectedRanges);
  } else {
    protection.setUnprotectedRanges(unprotectedRanges || []);
  }
}

function configureManagedSheetProtection_(protection, sheet, unprotectedRanges) {
  protection.setDescription(LAB.PROTECTION_PREFIX + sheet.getName());
  protection.setWarningOnly(false);

  const admins = getAdminEmails_();
  admins.forEach(function (email) {
    try { protection.addEditor(email); } catch (err) {}
  });

  try {
    protection.getEditors().forEach(function (editor) {
      const email = normalizeEmail_(editor.getEmail());
      if (email && admins.length > 0 && admins.indexOf(email) === -1) {
        protection.removeEditor(editor);
      }
    });
  } catch (err) {}

  try {
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
  } catch (err) {}

  protection.setUnprotectedRanges(unprotectedRanges || []);
}

function removeManagedProtectionsFromSheet_(sheet) {
  if (!sheet) return;
  [SpreadsheetApp.ProtectionType.SHEET, SpreadsheetApp.ProtectionType.RANGE]
    .forEach(function (type) {
      sheet.getProtections(type).forEach(function (protection) {
        const description = protection.getDescription() || '';
        if (
          description.indexOf(LAB.PROTECTION_PREFIX) === 0 ||
          description.indexOf(LAB.ROSTER_ROW_PREFIX) === 0 ||
          description.indexOf(LAB.REGISTER_ROW_PREFIX) === 0 ||
          description.indexOf(LAB.GENERAL_ROW_PREFIX) === 0 ||
          description === LAB.LEGACY_LOCK_DESC
        ) {
          try { protection.remove(); } catch (err) {}
        }
      });
    });
}

function removeAllManagedProtections_(ss) {
  ss.getSheets().forEach(function (sheet) {
    removeManagedProtectionsFromSheet_(sheet);
  });
}

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

/**********************************************************************
 * CHẾ ĐỘ VÀ NỘI DUNG HIỂN THỊ
 **********************************************************************/

function getWeekMode_() {
  return getDocumentProperty_(LAB.PROP.WEEK_MODE) || LAB.MODE_REGULAR;
}

function setWeekMode_(mode) {
  if ([LAB.MODE_REGULAR, LAB.MODE_GENERAL_CLEANING].indexOf(mode) === -1) {
    throw new Error('Chế độ tuần không hợp lệ: ' + mode);
  }
  setDocumentProperty_(LAB.PROP.WEEK_MODE, mode);
}

function getAssignmentMode_() {
  return LAB.ASSIGN_SELF;
}

function setAssignmentMode_(mode) {
  setDocumentProperty_(LAB.PROP.ASSIGNMENT_MODE, LAB.ASSIGN_SELF);
}

function weekModeText_(mode) {
  return mode === LAB.MODE_GENERAL_CLEANING
    ? 'Tuần có tổng vệ sinh Thứ 6'
    : 'Tuần thường';
}

function assignmentModeText_(mode) {
  return 'Sinh viên tự đăng ký';
}

function getHomeTimeText_() {
  return getWeekMode_() === LAB.MODE_GENERAL_CLEANING
    ? 'Thứ 2-Thứ 5 trực thường; Thứ 6 lúc 14:00 tổng vệ sinh'
    : 'Thứ 2 đến Thứ 6';
}

function getHomeScheduleText_(schedule) {
  if (!schedule || schedule.length === 0) return 'Chưa có sinh viên';
  const parts = schedule.map(function (day) { return day.day + ': ' + day.total + ' SV'; });
  if (getWeekMode_() === LAB.MODE_GENERAL_CLEANING) parts.push('Thứ 6: tổng vệ sinh');
  return parts.join('; ');
}

function getConfigNoteText_(studentCount, schedule) {
  if (studentCount === 0) {
    return 'Chưa có sinh viên áp dụng cho tuần này. Sinh viên nhập MSSV - họ tên vào cột B tại DANH SÁCH SV, không cần email; danh sách áp dụng vào lịch tuần mới.';
  }
  return 'Tuần này hệ thống dùng ' + studentCount + ' SV, tạo đúng số slot và chia tương đối đều cho ' + schedule.length +
    ' ngày. Danh sách mới không làm đổi lịch đang chạy; sẽ tự áp dụng Chủ nhật 16:00.';
}

function getRegisterTitle_() {
  const action = getAssignmentMode_() === LAB.ASSIGN_AUTO ? 'PHÂN CÔNG' : 'ĐĂNG KÝ';
  return getWeekMode_() === LAB.MODE_GENERAL_CLEANING
    ? action + ' TRỰC VỆ SINH LAB - TUẦN CÓ TỔNG VỆ SINH THỨ 6'
    : action + ' TRỰC VỆ SINH LAB';
}

function getRegisterGuideText_() {
  const modeText = 'Chọn tên tại cột G; mỗi tên chỉ 1 ca. Không đối chiếu email. Muốn đổi/hủy ca thì nhờ cán bộ.';
  const weekText = getWeekMode_() === LAB.MODE_GENERAL_CLEANING
    ? ' Thứ 6 thực hiện tổng vệ sinh theo sheet riêng.'
    : '';
  return modeText + ' Cán bộ tick cột K sau khi kiểm tra.' + weekText;
}

function getOpenMessage_(signed, total) {
  if (total === 0) return 'CHƯA CÓ SINH VIÊN - hãy nhập danh sách tại sheet DANH SÁCH SV';
  if (getWorkflowPhase_() === LAB.PHASE_ROSTER_OPEN) return 'CHỜ HỆ THỐNG MỞ ĐĂNG KÝ TUẦN MỚI';
  if (getWorkflowPhase_() === LAB.PHASE_CLOSED) return 'ĐĂNG KÝ ĐÃ KHÓA: ' + signed + '/' + total + ' sinh viên';
  return 'Đang mở đăng ký: ' + signed + '/' + total + ' sinh viên';
}

function getFullMessage_(signed, total) {
  return 'ĐÃ ĐỦ ' + signed + '/' + total + ' SINH VIÊN - ĐÃ KHÓA ĐĂNG KÝ';
}

/**********************************************************************
 * NGÀY GIỜ - LUÔN TÍNH THEO MÚI GIỜ VIỆT NAM
 **********************************************************************/

function getRegisterWeekMonday_(ss) {
  const reg = ss.getSheetByName(LAB.SHEETS.REGISTER);
  if (!reg) return null;
  try {
    return parseSheetDate_(reg.getRange('B2').getValue());
  } catch (err) {
    return null;
  }
}

function parseSheetDate_(value) {
  if (isValidDate_(value)) return getMonday_(value);
  const text = String(value || '').trim();
  const match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!match) return null;
  return localDateFromYmd_(match[3] + '-' + twoDigits_(match[2]) + '-' + twoDigits_(match[1]));
}

function getExpectedOpenMonday_(date) {
  const monday = getMonday_(date);
  const day = getLocalWeekday_(date);
  const hour = Number(Utilities.formatDate(date, LAB.TIMEZONE, 'H'));
  return day === 7 && hour >= LAB.RESET_HOUR ? addDays_(monday, 7) : monday;
}

function getNextMonday_(date) {
  return addDays_(getMonday_(date), 7);
}

function getMonday_(date) {
  const ymd = Utilities.formatDate(date, LAB.TIMEZONE, 'yyyy-MM-dd');
  const parts = ymd.split('-').map(Number);
  const utc = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2], 12, 0, 0));
  const day = utc.getUTCDay();
  const diff = day === 0 ? 6 : day - 1;
  utc.setUTCDate(utc.getUTCDate() - diff);
  const mondayYmd = [
    utc.getUTCFullYear(),
    twoDigits_(utc.getUTCMonth() + 1),
    twoDigits_(utc.getUTCDate())
  ].join('-');
  return localDateFromYmd_(mondayYmd);
}

function getLocalWeekday_(date) {
  const ymd = Utilities.formatDate(date, LAB.TIMEZONE, 'yyyy-MM-dd');
  const parts = ymd.split('-').map(Number);
  const day = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2], 12, 0, 0)).getUTCDay();
  return day === 0 ? 7 : day;
}

function localDateFromYmd_(ymd) {
  return new Date(ymd + 'T00:00:00+07:00');
}

function addDays_(date, days) {
  return new Date(date.getTime() + Number(days || 0) * 24 * 60 * 60 * 1000);
}

function weekKey_(date) {
  return Utilities.formatDate(date, LAB.TIMEZONE, 'yyyyMMdd');
}

function timestampKey_(date) {
  return Utilities.formatDate(date, LAB.TIMEZONE, 'yyyyMMdd_HHmmss_SSS');
}

function formatDate_(value) {
  if (isValidDate_(value)) return Utilities.formatDate(value, LAB.TIMEZONE, 'dd/MM/yyyy');
  return String(value || '');
}

function isValidDate_(value) {
  return Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime());
}

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

/**********************************************************************
 * KIỂM THỬ LOGIC THUẦN - KHÔNG GHI DỮ LIỆU
 **********************************************************************/

function runSelfTests_() {
  for (let count = 0; count <= 300; count++) {
    [LAB.MODE_REGULAR, LAB.MODE_GENERAL_CLEANING].forEach(function (mode) {
      const schedule = getScheduleConfig_(count, mode);
      const expectedDays = mode === LAB.MODE_GENERAL_CLEANING ? 4 : 5;
      assertTest_(schedule.length === expectedDays, 'Sai số ngày với ' + count + ' SV');
      assertTest_(getScheduleTotal_(schedule) === count, 'Sai tổng slot với ' + count + ' SV');

      const daily = schedule.map(function (day) {
        assertTest_(
          day.lab404 + day.lab403 + day.phongPhu === day.total,
          'Sai tổng phòng tại ' + day.day + ', ' + count + ' SV'
        );
        return day.total;
      });
      const max = Math.max.apply(null, daily);
      const min = Math.min.apply(null, daily);
      assertTest_(max - min <= 1, 'Phân ngày không đều với ' + count + ' SV');
    });

    const tasks = buildDynamicGeneralTasks_(count);
    assertTest_(tasks.length === count, 'Sai số việc tổng vệ sinh với ' + count + ' SV');
  }

  assertTest_(
    registrationStatusText_('', false, LAB.ASSIGN_SELF) === 'Còn trống',
    'Sai trạng thái slot trống'
  );
  assertTest_(
    registrationStatusText_('SV A', false, LAB.ASSIGN_SELF) === 'Đã đăng ký - chưa trực',
    'Sai trạng thái đã đăng ký'
  );
  assertTest_(
    registrationStatusText_('SV A', true, LAB.ASSIGN_SELF) === 'Đã trực xong',
    'Sai trạng thái hoàn thành'
  );
}

function assertTest_(condition, message) {
  if (!condition) throw new Error('Kiểm thử thất bại: ' + message);
}
