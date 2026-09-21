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
