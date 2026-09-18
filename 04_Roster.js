/**********************************************************************
 * THAY ĐỔI CHẾ ĐỘ VÀ ĐỒNG BỘ DANH SÁCH
 **********************************************************************/

function changeAssignmentModeFromMenu_(mode) {
  runManualAction_('changeAssignmentModeFromMenu', function () {
    const ss = getSpreadsheet_();
    assertAdmin_(ss);
    if (getAssignmentMode_() === mode) {
      toast_('Hệ thống đã ở chế độ này.', 'Không thay đổi', 4);
      return;
    }

    if (hasWeekActivity_(ss) && !confirm_(
      'Đổi chế độ phân công?',
      'Lịch hiện tại đã có dữ liệu. Hệ thống sẽ lưu một bản lịch sử và tạo lại lịch. Tiếp tục?'
    )) return;

    if (hasWeekActivity_(ss)) {
      archiveCurrentWeek_(ss, 'Ảnh chụp trước khi đổi chế độ phân công', true);
    }

    setAssignmentMode_(mode);
    const students = getValidatedStudents_(ss, true);
    const monday = getRegisterWeekMonday_(ss) || getExpectedOpenMonday_(new Date());
    rebuildWeekThenOpen_(ss, monday, students, 'CHANGE_ASSIGNMENT_MODE');
    logSystem_('INFO', 'changeAssignmentModeFromMenu', 'Đổi chế độ phân công', assignmentModeText_(mode));
    toast_('Đã chuyển sang: ' + assignmentModeText_(mode) + '.', 'Hoàn tất', 6);
  });
}

function changeWeekModeFromMenu_(mode) {
  runManualAction_('changeWeekModeFromMenu', function () {
    const ss = getSpreadsheet_();
    assertAdmin_(ss);
    const targetText = weekModeText_(mode);

    if (hasWeekActivity_(ss) && !confirm_(
      'Tạo ' + targetText.toLowerCase() + '?',
      'Lịch hiện tại đã có dữ liệu. Hệ thống sẽ lưu một bản lịch sử và tạo lại lịch theo danh sách mới nhất. Tiếp tục?'
    )) return;

    if (hasWeekActivity_(ss)) {
      archiveCurrentWeek_(ss, 'Ảnh chụp trước khi đổi loại tuần', true);
    }

    setWeekMode_(mode);
    const students = getValidatedStudents_(ss, true);
    const monday = getRegisterWeekMonday_(ss) || getExpectedOpenMonday_(new Date());
    rebuildWeekThenOpen_(ss, monday, students, 'CHANGE_WEEK_MODE');
    logSystem_('INFO', 'changeWeekModeFromMenu', 'Đổi loại tuần', targetText);
    toast_('Đã tạo ' + targetText.toLowerCase() + ' cho ' + students.length + ' SV.', 'Hoàn tất', 7);
  });
}

function handleConfigWeekModeEdit_(ss, e) {
  const value = String(e.range.getValue() || '').trim();
  const requested = value === weekModeText_(LAB.MODE_GENERAL_CLEANING)
    ? LAB.MODE_GENERAL_CLEANING
    : value === weekModeText_(LAB.MODE_REGULAR)
      ? LAB.MODE_REGULAR
      : null;

  if (!requested) {
    e.range.setValue(weekModeText_(getWeekMode_()));
    toast_('Giá trị chế độ tuần không hợp lệ.', 'Đã hoàn tác', 5);
    return;
  }

  if (requested === getWeekMode_()) return;

  if (hasWeekActivity_(ss)) {
    e.range.setValue(weekModeText_(getWeekMode_()));
    toast_('Tuần đã có dữ liệu. Hãy đổi chế độ bằng menu để hệ thống lưu lịch sử trước.', 'Chưa thay đổi', 8);
    return;
  }

  setWeekMode_(requested);
  const students = getValidatedStudents_(ss, true);
  const monday = getRegisterWeekMonday_(ss) || getExpectedOpenMonday_(new Date());
  rebuildWeekThenOpen_(ss, monday, students, 'CONFIG_WEEK_MODE');
}

function handleConfigAssignmentModeEdit_(ss, e) {
  setAssignmentMode_(LAB.ASSIGN_SELF);
  e.range.setValue(assignmentModeText_(LAB.ASSIGN_SELF));
  toast_('Hệ thống này cố định ở chế độ sinh viên tự đăng ký.', 'Không thay đổi', 5);
}

function handleRosterEdit_(ss, e) {
  if (!claimRosterRows_(ss, e)) return;
  const sh = ss.getSheetByName(LAB.SHEETS.STUDENTS);
  const nearBottom = sh && e.range.getLastRow() > sh.getMaxRows() - LAB.ROSTER_BUFFER_ROWS;
  const rosterGrew = nearBottom ? ensureRosterInputCapacity_(ss) : false;
  if (rosterGrew) {
    applyRosterProtectionsOnly_(ss, false);
  }
  setDocumentProperty_(LAB.PROP.ROSTER_DIRTY, 'YES');
  toast_(
    'Đã lưu dòng khai báo. Hệ thống sẽ tự kiểm tra và cập nhật tổng số trong ít phút.',
    'Đã ghi nhận',
    5
  );
}

function refreshRosterViews_(ss) {
  renumberRoster_(ss);
  const validation = validateRoster_(ss, true);
  setDocumentProperty_(LAB.PROP.PENDING_ROSTER_HASH, validation.hash);
  updateRosterCountDisplays_(ss, validation.students.length);
  setDocumentProperty_(LAB.PROP.ROSTER_DIRTY, 'NO');
}

function ensureRosterInputCapacity_(ss) {
  const sh = ss.getSheetByName(LAB.SHEETS.STUDENTS);
  if (!sh) return false;
  const maxRows = sh.getMaxRows();
  const inputCount = maxRows - LAB.FIRST_STUDENT_ROW + 1;
  const values = inputCount > 0
    ? sh.getRange(LAB.FIRST_STUDENT_ROW, 2, inputCount, 3).getValues()
    : [];
  let lastInputRow = LAB.FIRST_STUDENT_ROW - 1;

  values.forEach(function (row, index) {
    if (row.some(function (value) { return normalizeLabel_(value) !== ''; })) {
      lastInputRow = LAB.FIRST_STUDENT_ROW + index;
    }
  });

  if (maxRows - lastInputRow >= LAB.ROSTER_BUFFER_ROWS) return false;

  const firstNewRow = maxRows + 1;
  sh.insertRowsAfter(maxRows, LAB.ROSTER_GROWTH_ROWS);
  sh.getRange(firstNewRow, 1, LAB.ROSTER_GROWTH_ROWS, 7)
    .setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle')
    .setWrap(true);
  sh.getRange(firstNewRow, 2, LAB.ROSTER_GROWTH_ROWS, 1).setBackground('#fff2cc');
  sh.getRange(firstNewRow, 3, LAB.ROSTER_GROWTH_ROWS, 2).setBackground('#f3f3f3');
  return true;
}

function claimRosterRows_(ss, e) {
  const sh = ss.getSheetByName(LAB.SHEETS.STUDENTS);
  // Email chỉ dùng nhận diện cán bộ, KHÔNG là điều kiện sinh viên được ghi danh.
  const admin = isAdminEmail_(getEventUserEmail_(e));
  if (!sh || !isEditEventCurrent_(e)) return false;

  if (admin) return true;

  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1 ||
      e.range.getColumn() !== 2 || e.range.getRow() < LAB.FIRST_STUDENT_ROW) {
    revertEditedRange_(e);
    toast_('Chỉ nhập một ô trống ở cột B: MSSV - Họ và tên.', 'Đã hoàn tác', 7);
    return false;
  }

  const label = normalizeLabel_(e.range.getValue());
  const previous = normalizeLabel_(e.oldValue);
  if (previous) {
    if (canonicalLabel_(previous) !== canonicalLabel_(label)) {
      revertEditedRange_(e);
      toast_('Danh sách khai một lần cho cả học kỳ. Cần sửa hoặc xóa tên thì nhờ cán bộ.', 'Không thể thay đổi', 8);
    }
    return false;
  }
  if (!label) return false;

  const last = Math.max(sh.getLastRow(), LAB.FIRST_STUDENT_ROW);
  const allRows = sh.getRange(LAB.FIRST_STUDENT_ROW, 2, last - LAB.FIRST_STUDENT_ROW + 1, 1).getValues();
  const duplicate = allRows.some(function (row, index) {
    return LAB.FIRST_STUDENT_ROW + index !== e.range.getRow() &&
      canonicalLabel_(row[0]) === canonicalLabel_(label);
  });
  if (duplicate) {
    revertEditedRange_(e);
    toast_('MSSV - họ tên đã có trong danh sách. Không cần nhập lại.', 'Đã chặn trùng', 7);
    return false;
  }

  // Không ghi email và không yêu cầu e.user. Email cũ, nếu có, được giữ nguyên.
  return isEditEventCurrent_(e);
}

function revertEditedRange_(e) {
  try {
    // Không cho trigger chạy trễ ghi đè một thay đổi mới hơn của người khác.
    if (!isEditEventCurrent_(e)) return;
    if (e.range.getNumRows() === 1 && e.range.getNumColumns() === 1) {
      if (typeof e.oldValue === 'undefined') e.range.clearContent();
      else e.range.setValue(e.oldValue);
    } else {
      e.range.clearContent();
    }
  } catch (err) {}
}

function isEditEventCurrent_(e) {
  if (!e || !e.range) return false;
  if (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1) return true;

  const current = e.range.getValue();
  const currentText = current === null || typeof current === 'undefined'
    ? ''
    : String(current).trim();
  const eventText = typeof e.value === 'undefined' || e.value === null
    ? ''
    : String(e.value).trim();

  if (/^(true|false)$/i.test(currentText) || /^(true|false)$/i.test(eventText)) {
    return currentText.toLowerCase() === eventText.toLowerCase();
  }
  return currentText === eventText;
}

function updateRosterCountDisplays_(ss, count) {
  const safeCount = Math.max(0, Number(count) || 0);
  const home = ss.getSheetByName(LAB.SHEETS.HOME);
  if (home) {
    home.getRange('B4').setValue(
      safeCount + ' sinh viên (thay đổi áp dụng Chủ nhật 16:00)'
    );
  }
  const config = ss.getSheetByName(LAB.SHEETS.CONFIG);
  if (config) config.getRange('B2').setValue(safeCount);
  updateWorkflowBanner_(ss);
}

function countRosterEntries_(ss) {
  const sh = ss.getSheetByName(LAB.SHEETS.STUDENTS);
  if (!sh || sh.getLastRow() < LAB.FIRST_STUDENT_ROW) return 0;
  const count = sh.getLastRow() - LAB.FIRST_STUDENT_ROW + 1;
  return sh.getRange(LAB.FIRST_STUDENT_ROW, 2, count, 1)
    .getValues()
    .filter(function (row) { return normalizeLabel_(row[0]) !== ''; })
    .length;
}


function findStudentByLabel_(students, label) {
  const target = canonicalLabel_(label);
  for (let i = 0; i < students.length; i++) {
    if (canonicalLabel_(students[i].label) === target) return students[i];
  }
  return null;
}

/**********************************************************************
 * DANH SÁCH SINH VIÊN ĐỘNG
 **********************************************************************/

function readRosterForPreservation_(sheet) {
  if (!sheet || sheet.getLastRow() < LAB.FIRST_STUDENT_ROW) return [];
  const count = sheet.getLastRow() - LAB.FIRST_STUDENT_ROW + 1;
  const isV4 = canonicalLabel_(sheet.getRange('C2').getValue()).indexOf('email') !== -1;
  const values = sheet.getRange(
    LAB.FIRST_STUDENT_ROW,
    1,
    count,
    Math.min(sheet.getMaxColumns(), 7)
  ).getValues();

  return values
    .filter(function (row) { return normalizeLabel_(row[1]) !== ''; })
    .map(function (row) {
      return {
        label: normalizeLabel_(row[1]),
        email: isV4 ? normalizeEmail_(row[2]) : '',
        group: String(isV4 ? row[3] || '' : row[2] || '').trim(),
        note: String(isV4 ? row[6] || '' : row[5] || '').trim()
      };
    });
}

function getValidatedStudents_(ss, throwOnInvalid) {
  const result = validateRoster_(ss, true);
  if (!result.valid && throwOnInvalid) throw new Error(result.message);
  return result.students;
}

function validateRoster_(ss, paintErrors) {
  const sh = ss.getSheetByName(LAB.SHEETS.STUDENTS);
  if (!sh) {
    return {
      valid: false,
      students: [],
      eligibleStudents: [],
      hash: '',
      message: 'Thiếu sheet DANH SÁCH SV.'
    };
  }

  const last = Math.max(sh.getLastRow(), LAB.FIRST_STUDENT_ROW);
  const count = last - LAB.FIRST_STUDENT_ROW + 1;
  const values = sh.getRange(LAB.FIRST_STUDENT_ROW, 2, count, 3).getValues();
  const students = [];
  const seenLabels = Object.create(null);
  const errorRows = {};
  let duplicateLabels = 0;
  let incompleteRows = 0;

  values.forEach(function (row, index) {
    const label = normalizeLabel_(row[0]);
    const email = normalizeEmail_(row[1]);
    const group = String(row[2] || '').trim();
    const sheetRow = LAB.FIRST_STUDENT_ROW + index;

    if (!label && !email && !group) return;
    if (!label) {
      errorRows[sheetRow] = true;
      incompleteRows++;
      return;
    }

    const key = canonicalLabel_(label);
    if (seenLabels[key]) {
      errorRows[seenLabels[key].row] = true;
      errorRows[sheetRow] = true;
      duplicateLabels++;
    } else {
      seenLabels[key] = { row: sheetRow };
    }

    // Email là thông tin tùy chọn; chỉ tên/MSSV quyết định tính hợp lệ.

    students.push({
      row: sheetRow,
      label: label,
      email: email,
      group: group
    });
  });

  if (paintErrors && count > 0) {
    const nameBackgrounds = values.map(function (row, index) {
      const sheetRow = LAB.FIRST_STUDENT_ROW + index;
      if (errorRows[sheetRow]) return ['#f4cccc'];
      return [normalizeLabel_(row[0]) ? '#ffffff' : '#fff2cc'];
    });
    const emailBackgrounds = values.map(function (row, index) {
      const sheetRow = LAB.FIRST_STUDENT_ROW + index;
      if (errorRows[sheetRow]) return ['#f4cccc'];
      return [normalizeEmail_(row[1]) ? '#ffffff' : '#f8f9fa'];
    });
    sh.getRange(LAB.FIRST_STUDENT_ROW, 2, count, 1).setBackgrounds(nameBackgrounds);
    sh.getRange(LAB.FIRST_STUDENT_ROW, 3, count, 1).setBackgrounds(emailBackgrounds);
  }

  const hash = rosterHash_(students);
  const eligibleStudents = students.filter(function (student) {
    return !errorRows[student.row];
  });
  if (Object.keys(errorRows).length > 0) {
    const details = [];
    if (incompleteRows) details.push(incompleteRows + ' dòng thiếu tên');
    if (duplicateLabels) details.push(duplicateLabels + ' tên/MSSV trùng');
    return {
      valid: false,
      students: students,
      eligibleStudents: eligibleStudents,
      hash: hash,
      message: 'Danh sách chưa hợp lệ: ' + details.join(', ') + '. Hãy sửa các ô màu đỏ.'
    };
  }

  return {
    valid: true,
    students: students,
    eligibleStudents: students.slice(),
    hash: hash,
    message: ''
  };
}

function renumberRoster_(ss) {
  const sh = ss.getSheetByName(LAB.SHEETS.STUDENTS);
  if (!sh) return;
  const last = Math.max(sh.getLastRow(), LAB.FIRST_STUDENT_ROW);
  const count = last - LAB.FIRST_STUDENT_ROW + 1;
  const labels = sh.getRange(LAB.FIRST_STUDENT_ROW, 2, count, 1).getValues();
  let number = 0;
  const numbers = labels.map(function (row) {
    if (!normalizeLabel_(row[0])) return [''];
    number++;
    return [number];
  });
  sh.getRange(LAB.FIRST_STUDENT_ROW, 1, count, 1).setValues(numbers);
}

function normalizeLabel_(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .trim();
}

function canonicalLabel_(value) {
  const normalized = normalizeLabel_(value);
  try {
    return normalized.normalize('NFC').toLocaleLowerCase('vi');
  } catch (err) {
    return normalized.toLowerCase();
  }
}

function rosterHash_(students) {
  const source = students
    .map(function (student) {
      return canonicalLabel_(student.label) + '|' + normalizeEmail_(student.email);
    })
    .sort()
    .join('|');
  return hashString_(source);
}

function refreshSystemStudentList_(ss, students, mondayOverride) {
  const sh = ss.getSheetByName(LAB.SHEETS.SYSTEM);
  resetSheet_(sh);
  ensureSheetSize_(sh, Math.max(students.length + 3, 20), 5);

  sh.getRange('A1:E1').setValues([[
    'DANH SÁCH SV HỢP LỆ', 'EMAIL', 'ROSTER_HASH', 'WEEK_KEY', 'VERSION'
  ]]).setFontWeight('bold');

  if (students.length > 0) {
    sh.getRange(2, 1, students.length, 2)
      .setValues(students.map(function (student) { return [student.label, student.email]; }));
  }

  sh.getRange('C2').setValue(rosterHash_(students));
  const monday = mondayOverride || getRegisterWeekMonday_(ss);
  sh.getRange('D2').setValue(monday ? weekKey_(monday) : '');
  sh.getRange('E2').setValue(LAB.VERSION);
  sh.hideSheet();
}

function getAppliedStudents_(ss) {
  const sh = ss.getSheetByName(LAB.SHEETS.SYSTEM);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  return values
    .map(function (row) {
      return { label: normalizeLabel_(row[0]), email: normalizeEmail_(row[1]) };
    })
    .filter(function (student) {
      return student.label !== '';
    });
}

function getCurrentWeekStudents_(ss) {
  const applied = getAppliedStudents_(ss);
  if (hasAppliedRosterSnapshot_(ss)) return applied;
  const validation = validateRoster_(ss, false);
  return validation.valid ? validation.students : validation.eligibleStudents;
}

function mergeRosterSnapshots_(baseStudents, addedStudents) {
  const result = [];
  const labels = Object.create(null);
  (baseStudents || []).concat(addedStudents || []).forEach(function (student) {
    const label = normalizeLabel_(student && student.label);
    const key = canonicalLabel_(label);
    if (!label || labels[key]) return;
    labels[key] = true;
    result.push({
      row: student.row,
      label: label,
      email: normalizeEmail_(student.email),
      group: String(student.group || '').trim()
    });
  });
  return result;
}

function hasAppliedRosterSnapshot_(ss) {
  const sh = ss.getSheetByName(LAB.SHEETS.SYSTEM);
  if (!sh) return false;
  return normalizeLabel_(sh.getRange('E2').getValue()) !== '' ||
    normalizeLabel_(sh.getRange('C2').getValue()) !== '';
}

