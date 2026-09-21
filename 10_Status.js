/**********************************************************************
 * CẬP NHẬT TRẠNG THÁI VÀ THEO DÕI
 **********************************************************************/

function registrationStatusText_(student, done, assignmentMode) {
  if (!normalizeLabel_(student)) return 'Còn trống';
  if (done === true) return 'Đã trực xong';
  return (assignmentMode || getAssignmentMode_()) === LAB.ASSIGN_AUTO
    ? 'Đã phân công - chưa trực'
    : 'Đã đăng ký - chưa trực';
}

function hasOrphanRegistrationMetadata_(row) {
  if (normalizeLabel_(row[LAB.COL.SV - 1])) return false;
  return Boolean(
    normalizeLabel_(row[LAB.COL.TIME_DANGKY - 1]) ||
    normalizeLabel_(row[LAB.COL.GHICHU_SV - 1]) ||
    row[LAB.COL.DA_TRUC - 1] === true ||
    normalizeLabel_(row[LAB.COL.NGUOI_KIEMTRA - 1]) ||
    normalizeLabel_(row[LAB.COL.TIME_KIEMTRA - 1]) ||
    normalizeLabel_(row[LAB.COL.GHICHU_KIEMTRA - 1])
  );
}

function getRegistrationBannerState_(
  rowCount,
  rosterCount,
  signed,
  full,
  generalSigned,
  generalCount,
  phase,
  weekMode
) {
  if (rowCount !== rosterCount) {
    return {
      message: 'LỊCH TUẦN HIỆN TẠI CẦN PHỤC HỒI - hãy chạy kiểm tra hệ thống',
      background: '#f4cccc',
      fontColor: '#990000'
    };
  }

  if (phase === LAB.PHASE_CLOSED && !full) {
    return {
      message: rowCount === 0
        ? 'CHỜ MỞ ĐĂNG KÝ TUẦN MỚI - tự động Chủ nhật khoảng 16:00'
        : 'ĐĂNG KÝ ĐÃ KHÓA - liên hệ cán bộ nếu cần điều chỉnh',
      background: rowCount === 0 ? '#cfe2f3' : '#f4cccc',
      fontColor: rowCount === 0 ? '#073763' : '#990000'
    };
  }

  if (full) {
    return {
      message: getFullMessage_(signed, rowCount),
      background: '#d9ead3',
      fontColor: '#274e13'
    };
  }

  let message = getOpenMessage_(signed, rowCount);
  if (weekMode === LAB.MODE_GENERAL_CLEANING) {
    message += ' | Tổng vệ sinh: ' + generalSigned + '/' + generalCount;
  }
  return {
    message: message,
    background: '#fff2cc',
    fontColor: '#7f6000'
  };
}

/**
 * Đường xử lý nhanh cho lượt đăng ký.
 * Chỉ cập nhật trạng thái dòng vừa sửa và banner. Bảng theo dõi cùng dashboard
 * được trigger nền đồng bộ sau để không giữ khóa đăng ký quá lâu.
 */
function refreshRegistrationStateFast_(ss, editedRange) {
  const reg = ss.getSheetByName(LAB.SHEETS.REGISTER);
  if (!reg) return;

  const rowCount = getRegisterRowCount_(reg);
  const rosterCount = getCurrentWeekStudents_(ss).length;
  const oldFull = isRegistrationFullLocked_();
  const oldPhase = getWorkflowPhase_();
  const assignmentMode = getAssignmentMode_();
  const statusValues = [];
  const statusColors = [];
  let signed = 0;

  if (rowCount > 0) {
    const width = LAB.COL.DA_TRUC - LAB.COL.SV + 1;
    const data = reg
      .getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.SV, rowCount, width)
      .getValues();

    data.forEach(function (row) {
      const student = normalizeLabel_(row[0]);
      const done = row[LAB.COL.DA_TRUC - LAB.COL.SV] === true;
      statusValues.push([registrationStatusText_(student, done, assignmentMode)]);
      statusColors.push([student ? (done ? '#d9ead3' : '#fff2cc') : '#eeeeee']);
      if (student) signed++;
    });

    let firstStatusRow = LAB.FIRST_REGISTER_ROW;
    let statusRowCount = rowCount;
    const isRegisterEdit = editedRange &&
      editedRange.getSheet().getName() === LAB.SHEETS.REGISTER;

    if (isRegisterEdit) {
      firstStatusRow = Math.max(editedRange.getRow(), LAB.FIRST_REGISTER_ROW);
      const lastStatusRow = Math.min(
        editedRange.getLastRow(),
        LAB.FIRST_REGISTER_ROW + rowCount - 1
      );
      statusRowCount = Math.max(lastStatusRow - firstStatusRow + 1, 0);
    } else if (editedRange) {
      statusRowCount = 0;
    }

    if (statusRowCount > 0) {
      const offset = firstStatusRow - LAB.FIRST_REGISTER_ROW;
      reg.getRange(firstStatusRow, LAB.COL.TRANGTHAI, statusRowCount, 1)
        .setValues(statusValues.slice(offset, offset + statusRowCount))
        .setBackgrounds(statusColors.slice(offset, offset + statusRowCount))
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
    }
  }

  const weekMode = getWeekMode_();
  let generalCount = 0;
  let generalSigned = 0;
  if (weekMode === LAB.MODE_GENERAL_CLEANING) {
    const general = ss.getSheetByName(LAB.SHEETS.GENERAL_CLEANING);
    generalCount = getGeneralRowCount_(general);
    if (generalCount > 0) {
      generalSigned = general
        .getRange(LAB.FIRST_GENERAL_ROW, LAB.GENERAL_COL.SV, generalCount, 1)
        .getValues()
        .filter(function (row) { return normalizeLabel_(row[0]) !== ''; })
        .length;
    }
  }

  const regularFull = rowCount > 0 && signed === rowCount && rowCount === rosterCount;
  const generalFull = weekMode !== LAB.MODE_GENERAL_CLEANING ||
    (generalCount === rosterCount && generalSigned === generalCount);
  const full = regularFull && generalFull;
  if (oldFull !== full) setRegistrationFullLocked_(full);
  if (full && oldPhase === LAB.PHASE_REGISTRATION_OPEN) {
    setWorkflowPhase_(LAB.PHASE_CLOSED);
  }

  const newPhase = getWorkflowPhase_();
  const banner = getRegistrationBannerState_(
    rowCount,
    rosterCount,
    signed,
    full,
    generalSigned,
    generalCount,
    newPhase,
    weekMode
  );
  reg.getRange('C2:N2').merge()
    .setValue(banner.message)
    .setBackground(banner.background)
    .setFontColor(banner.fontColor)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  const accessChanged = oldFull !== full || oldPhase !== newPhase;
  setDocumentProperty_(LAB.PROP.STATUS_DIRTY, accessChanged ? 'ACCESS' : 'YES');
  if (oldPhase !== newPhase) updateWorkflowBanner_(ss);
}

function refreshAllStatuses_(ss, forceProtection) {
  const reg = ss.getSheetByName(LAB.SHEETS.REGISTER);
  const sv = ss.getSheetByName(LAB.SHEETS.STUDENTS);
  if (!reg || !sv) return;

  const rowCount = getRegisterRowCount_(reg);
  const appliedStudents = getCurrentWeekStudents_(ss);
  const rosterCount = appliedStudents.length;
  const oldFull = getDocumentProperty_(LAB.PROP.FULL_LOCK) === 'YES';
  const oldPhase = getWorkflowPhase_();
  const studentMap = {};
  let signed = 0;
  let completed = 0;
  const assignmentMode = getAssignmentMode_();

  if (rowCount > 0) {
    const data = reg.getRange(LAB.FIRST_REGISTER_ROW, 1, rowCount, 14).getValues();
    const statusValues = [];
    const statusColors = [];
    const orphanRows = [];

    data.forEach(function (row, index) {
      const day = row[LAB.COL.THU - 1];
      const date = row[LAB.COL.NGAY - 1];
      const lab = row[LAB.COL.LAB - 1];
      const student = normalizeLabel_(row[LAB.COL.SV - 1]);
      const done = row[LAB.COL.DA_TRUC - 1] === true;
      const status = registrationStatusText_(student, done, assignmentMode);
      let color = '#eeeeee';

      if (student && done) {
        signed++;
        completed++;
        color = '#d9ead3';
      } else if (student) {
        signed++;
        color = '#fff2cc';
      }

      statusValues.push([status]);
      statusColors.push([color]);
      if (student) {
        studentMap[canonicalLabel_(student)] = {
          status: status,
          schedule: day + ' - ' + formatDate_(date) + ' - ' + lab
        };
      } else if (hasOrphanRegistrationMetadata_(row)) {
        orphanRows.push(LAB.FIRST_REGISTER_ROW + index);
      }
    });

    orphanRows.forEach(function (row) {
      clearRegistrationRowInput_(reg, row);
    });

    reg.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.TRANGTHAI, rowCount, 1)
      .setValues(statusValues)
      .setBackgrounds(statusColors)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }

  updateStudentStatusSheet_(sv, studentMap, appliedStudents);
  let generalCount = 0;
  let generalSigned = 0;
  if (getWeekMode_() === LAB.MODE_GENERAL_CLEANING) {
    const general = ss.getSheetByName(LAB.SHEETS.GENERAL_CLEANING);
    generalCount = getGeneralRowCount_(general);
    if (generalCount > 0) {
      generalSigned = general
        .getRange(LAB.FIRST_GENERAL_ROW, LAB.GENERAL_COL.SV, generalCount, 1)
        .getValues()
        .filter(function (row) { return normalizeLabel_(row[0]) !== ''; })
        .length;
    }
  }

  const regularFull = rowCount > 0 && signed === rowCount && rowCount === rosterCount;
  const generalFull = getWeekMode_() !== LAB.MODE_GENERAL_CLEANING ||
    (generalCount === rosterCount && generalSigned === generalCount);
  const full = regularFull && generalFull;
  setRegistrationFullLocked_(full);
  if (full && oldPhase === LAB.PHASE_REGISTRATION_OPEN) {
    setWorkflowPhase_(LAB.PHASE_CLOSED);
  }

  const currentPhase = getWorkflowPhase_();
  const banner = getRegistrationBannerState_(
    rowCount,
    rosterCount,
    signed,
    full,
    generalSigned,
    generalCount,
    currentPhase,
    getWeekMode_()
  );

  reg.getRange('C2:N2').merge()
    .setValue(banner.message)
    .setBackground(banner.background)
    .setFontColor(banner.fontColor)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  updateDashboardValues_(ss);
  if (forceProtection || oldFull !== full || oldPhase !== currentPhase) {
    applyRegistrationAccessProtections_(ss, false);
  }
  updateWorkflowBanner_(ss);
  setDocumentProperty_(LAB.PROP.STATUS_DIRTY, 'NO');
}

function updateStudentStatusSheet_(sheet, studentMap, appliedStudents) {
  const last = Math.max(sheet.getLastRow(), LAB.FIRST_STUDENT_ROW);
  const count = last - LAB.FIRST_STUDENT_ROW + 1;
  const labels = sheet.getRange(LAB.FIRST_STUDENT_ROW, 2, count, 1).getValues();
  const numbers = [];
  const statuses = [];
  const schedules = [];
  const colors = [];
  const applied = {};
  let stt = 0;

  (appliedStudents || []).forEach(function (student) {
    applied[canonicalLabel_(student.label)] = true;
  });

  labels.forEach(function (row) {
    const label = normalizeLabel_(row[0]);
    if (!label) {
      numbers.push(['']);
      statuses.push(['']);
      schedules.push(['']);
      colors.push(['#ffffff']);
      return;
    }

    stt++;
    numbers.push([stt]);
    const item = studentMap[canonicalLabel_(label)];
    if (item) {
      statuses.push([item.status]);
      schedules.push([item.schedule]);
      colors.push([item.status === 'Đã trực xong' ? '#d9ead3' : '#fff2cc']);
    } else {
      statuses.push([
        applied[canonicalLabel_(label)]
          ? 'Chưa đăng ký'
          : 'Chờ áp dụng Chủ nhật 16:00'
      ]);
      schedules.push(['']);
      colors.push(['#f4cccc']);
    }
  });

  sheet.getRange(LAB.FIRST_STUDENT_ROW, 1, count, 1).setValues(numbers);
  sheet.getRange(LAB.FIRST_STUDENT_ROW, 5, count, 1)
    .setValues(statuses)
    .setBackgrounds(colors)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange(LAB.FIRST_STUDENT_ROW, 6, count, 1).setValues(schedules).setWrap(true);
}

function getRegisterRowCount_(sheet) {
  if (!sheet || sheet.getLastRow() < LAB.FIRST_REGISTER_ROW) return 0;
  const count = sheet.getLastRow() - LAB.FIRST_REGISTER_ROW + 1;
  return sheet.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.STT, count, 1)
    .getValues()
    .filter(function (row) { return row[0] !== '' && row[0] !== null; })
    .length;
}

function getGeneralRowCount_(sheet) {
  if (!sheet || sheet.getLastRow() < LAB.FIRST_GENERAL_ROW) return 0;
  const count = sheet.getLastRow() - LAB.FIRST_GENERAL_ROW + 1;
  return sheet.getRange(LAB.FIRST_GENERAL_ROW, LAB.GENERAL_COL.STT, count, 1)
    .getValues()
    .filter(function (row) { return row[0] !== '' && row[0] !== null; })
    .length;
}

function hasAnyCompletion_(ss) {
  const reg = ss.getSheetByName(LAB.SHEETS.REGISTER);
  const count = getRegisterRowCount_(reg);
  if (count > 0 && reg.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.DA_TRUC, count, 1)
    .getValues().some(function (row) { return row[0] === true; })) return true;

  if (getWeekMode_() !== LAB.MODE_GENERAL_CLEANING) return false;
  const general = ss.getSheetByName(LAB.SHEETS.GENERAL_CLEANING);
  const generalCount = getGeneralRowCount_(general);
  return generalCount > 0 && general.getRange(LAB.FIRST_GENERAL_ROW, LAB.GENERAL_COL.DONE, generalCount, 1)
    .getValues().some(function (row) { return row[0] === true; });
}

function hasAnySignupOrCompletion_(ss) {
  const reg = ss.getSheetByName(LAB.SHEETS.REGISTER);
  const count = getRegisterRowCount_(reg);
  if (count > 0) {
    const data = reg.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.SV, count, 5).getValues();
    if (data.some(function (row) { return normalizeLabel_(row[0]) !== '' || row[4] === true; })) return true;
  }

  if (getWeekMode_() !== LAB.MODE_GENERAL_CLEANING) return false;
  const general = ss.getSheetByName(LAB.SHEETS.GENERAL_CLEANING);
  const generalCount = getGeneralRowCount_(general);
  if (generalCount > 0) {
    const data = general.getRange(LAB.FIRST_GENERAL_ROW, LAB.GENERAL_COL.SV, generalCount, 6).getValues();
    if (data.some(function (row) { return normalizeLabel_(row[0]) !== '' || row[5] === true; })) return true;
  }
  return false;
}

function hasWeekActivity_(ss) {
  return getAssignmentMode_() === LAB.ASSIGN_AUTO
    ? hasAnyCompletion_(ss)
    : hasAnySignupOrCompletion_(ss);
}

