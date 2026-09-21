/**********************************************************************
 * TỔNG VỆ SINH THỨ 6 - SỐ VIỆC TỰ ĐỔI THEO SỐ SINH VIÊN
 **********************************************************************/

function buildGeneralCleaningSheet_(ss, monday, students) {
  const sh = ss.getSheetByName(LAB.SHEETS.GENERAL_CLEANING);
  resetSheet_(sh);
  ensureSheetSize_(sh, Math.max(students.length + 15, 40), 10);
  sh.showSheet();
  sh.setTabColor('#990000');
  sh.setHiddenGridlines(true);

  const targetDate = addDays_(monday, 4);
  const dateText = formatDate_(targetDate);
  const tasks = buildDynamicGeneralTasks_(students.length);
  const auto = getAssignmentMode_() === LAB.ASSIGN_AUTO;
  const orderedStudents = auto
    ? orderStudentsForWeek_(students, monday, 'GENERAL_CLEANING')
    : [];
  const rows = tasks.map(function (task, index) {
    return [
      index + 1,
      auto && orderedStudents[index] ? orderedStudents[index].label : '',
      task.area,
      task.job,
      task.requirement,
      '14:00',
      false,
      '',
      '',
      ''
    ];
  });

  sh.getRange('A1:J1')
    .merge()
    .setValue('TỔNG VỆ SINH LAB - THỨ 6 NGÀY ' + dateText)
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#990000')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 44);

  sh.getRange('A2:B5').setValues([
    ['Thời gian', '14:00, Thứ 6 ngày ' + dateText],
    ['Số lượng', students.length + ' sinh viên (tự động theo danh sách)'],
    ['Hình thức', auto ? 'Đã tự động phân công' : 'Sinh viên tự chọn 01 dòng công việc'],
    ['An toàn', 'Không tự ý mở máy, tháo lắp thiết bị hoặc di chuyển mẫu/hóa chất.']
  ]);
  sh.getRange('A2:A5').setFontWeight('bold').setBackground('#f4cccc');
  sh.getRange('A2:B5').setBorder(true, true, true, true, true, true).setWrap(true);

  const headers = [
    'STT',
    auto ? 'Sinh viên được phân công' : 'Sinh viên đăng ký',
    'Phòng/khu vực',
    'Việc cần làm',
    'Yêu cầu hoàn thành',
    'Giờ thực hiện',
    'Đã xong?',
    'Người kiểm tra',
    'Thời điểm kiểm tra',
    'Ghi chú kiểm tra'
  ];
  sh.getRange(6, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#ea9999')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  if (rows.length > 0) {
    sh.getRange(LAB.FIRST_GENERAL_ROW, 1, rows.length, headers.length).setValues(rows);
    sh.getRange(LAB.FIRST_GENERAL_ROW, LAB.GENERAL_COL.DONE, rows.length, 1)
      .insertCheckboxes()
      .setHorizontalAlignment('center');
    sh.getRange(LAB.FIRST_GENERAL_ROW, LAB.GENERAL_COL.CHECK_TIME, rows.length, 1)
      .setNumberFormat('dd/MM/yyyy HH:mm');
    applyGeneralCleaningValidation_(ss, rows.length, students.length);
    colorGeneralCleaningRows_(sh, rows.length);
    sh.getRange(6, 1, rows.length + 1, headers.length)
      .setBorder(true, true, true, true, true, true)
      .setVerticalAlignment('middle')
      .setWrap(true);
  }

  sh.getRange('A:A').setHorizontalAlignment('center');
  sh.getRange('F:I').setHorizontalAlignment('center');
  sh.setFrozenRows(6);
  [50, 290, 130, 390, 400, 90, 90, 190, 155, 230]
    .forEach(function (width, index) { sh.setColumnWidth(index + 1, width); });

  upsertGeneralCleaningHistory_(ss, targetDate);
}

function hideGeneralCleaningSheet_(ss) {
  const sh = ss.getSheetByName(LAB.SHEETS.GENERAL_CLEANING);
  if (!sh) return;
  try { sh.hideSheet(); } catch (err) {}
}

function applyGeneralCleaningValidation_(ss, rowCount, studentCount) {
  const sh = ss.getSheetByName(LAB.SHEETS.GENERAL_CLEANING);
  if (!sh || rowCount <= 0) return;
  const range = sh.getRange(LAB.FIRST_GENERAL_ROW, LAB.GENERAL_COL.SV, rowCount, 1);
  if (studentCount <= 0) {
    range.clearDataValidations();
    return;
  }

  const source = ss.getSheetByName(LAB.SHEETS.SYSTEM).getRange(2, 1, studentCount, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(source, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(rule);
}

function handleGeneralCleaningEdit_(ss, e) {
  const sh = e.range.getSheet();
  if (e.range.getLastRow() < LAB.FIRST_GENERAL_ROW) return;

  if (rangeTouchesColumn_(e.range, LAB.GENERAL_COL.SV)) {
    if (getWorkflowPhase_() !== LAB.PHASE_REGISTRATION_OPEN) {
      revertEditedRange_(e);
      toast_('Bảng đăng ký tổng vệ sinh chưa mở hoặc đã bị khóa.', 'Không thể đăng ký', 7);
      return;
    }
    if (sanitizeGeneralSelfRegistration_(ss, e) === false) return;
  }

  if (rangeTouchesColumn_(e.range, LAB.GENERAL_COL.DONE)) {
    processGeneralCompletion_(ss, e);
  }

  if (
    rangeTouchesColumn_(e.range, LAB.GENERAL_COL.SV) ||
    rangeTouchesColumn_(e.range, LAB.GENERAL_COL.DONE)
  ) {
    refreshRegistrationStateFast_(ss, e.range);
  }
}

function restoreAutoGeneralNames_(ss) {
  const sh = ss.getSheetByName(LAB.SHEETS.GENERAL_CLEANING);
  const students = getCurrentWeekStudents_(ss);
  const monday = getRegisterWeekMonday_(ss) || getExpectedOpenMonday_(new Date());
  const ordered = orderStudentsForWeek_(students, monday, 'GENERAL_CLEANING');
  const rowCount = getGeneralRowCount_(sh);
  if (rowCount !== ordered.length) {
    throw new Error('Số dòng tổng vệ sinh không khớp danh sách. Hãy tạo lại bảng tổng vệ sinh từ menu.');
  }
  if (rowCount === 0) return;
  sh.getRange(LAB.FIRST_GENERAL_ROW, LAB.GENERAL_COL.SV, rowCount, 1)
    .setValues(ordered.map(function (student) { return [student.label]; }));
}

function sanitizeGeneralSelfRegistration_(ss, e) {
  const sh = ss.getSheetByName(LAB.SHEETS.GENERAL_CLEANING);
  const students = getCurrentWeekStudents_(ss);
  const count = getGeneralRowCount_(sh);
  if (!sh || count === 0 || !isEditEventCurrent_(e)) return false;
  const admin = isAdminEmail_(getEventUserEmail_(e));

  if (!admin && (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1 ||
      e.range.getColumn() !== LAB.GENERAL_COL.SV)) {
    revertEditedRange_(e);
    toast_('Mỗi lần chỉ chọn một ô tên trong bảng tổng vệ sinh.', 'Đã hoàn tác', 7);
    return false;
  }

  if (!admin && normalizeLabel_(e.oldValue)) {
    if (canonicalLabel_(e.oldValue) !== canonicalLabel_(e.range.getValue())) {
      revertEditedRange_(e);
      toast_('Dòng này đã có tên. Muốn đổi hoặc hủy thì nhờ cán bộ.', 'Không thể thay đổi', 7);
    }
    return false;
  }

  const start = Math.max(e.range.getRow(), LAB.FIRST_GENERAL_ROW);
  const end = Math.min(e.range.getLastRow(), LAB.FIRST_GENERAL_ROW + count - 1);
  if (start > end) {
    revertEditedRange_(e);
    return false;
  }
  const occupied = Object.create(null);
  const allNames = sh.getRange(LAB.FIRST_GENERAL_ROW, LAB.GENERAL_COL.SV, count, 1).getValues();
  allNames.forEach(function (row, index) {
    const sheetRow = LAB.FIRST_GENERAL_ROW + index;
    if (sheetRow >= start && sheetRow <= end) return;
    const key = canonicalLabel_(row[0]);
    if (key) occupied[key] = sheetRow;
  });

  if (!admin) {
    const chosen = normalizeLabel_(allNames[start - LAB.FIRST_GENERAL_ROW][0]);
    const student = chosen ? findStudentByLabel_(students, chosen) : null;
    if (!chosen) return false;
    if (!student || occupied[canonicalLabel_(student.label)]) {
      revertEditedRange_(e);
      toast_(
        student ? 'Tên này đã đăng ký tổng vệ sinh rồi.' : 'Tên chưa có trong danh sách áp dụng cho tuần này.',
        'Không thể đăng ký',
        8
      );
      return false;
    }
  }

  if (!isEditEventCurrent_(e)) return false;
  let corrections = 0;
  for (let row = start; row <= end; row++) {
    const chosen = normalizeLabel_(allNames[row - LAB.FIRST_GENERAL_ROW][0]);
    if (!chosen) {
      sh.getRange(row, LAB.GENERAL_COL.DONE, 1, 4).setValues([[false, '', '', '']]);
      continue;
    }
    const student = findStudentByLabel_(students, chosen);
    const key = student ? canonicalLabel_(student.label) : '';
    if (!student || occupied[key]) {
      sh.getRange(row, LAB.GENERAL_COL.SV).clearContent();
      sh.getRange(row, LAB.GENERAL_COL.DONE, 1, 4).setValues([[false, '', '', '']]);
      corrections++;
    } else {
      occupied[key] = row;
      if (chosen !== student.label) sh.getRange(row, LAB.GENERAL_COL.SV).setValue(student.label);
      sh.getRange(row, LAB.GENERAL_COL.DONE, 1, 4).setValues([[false, '', '', '']]);
    }
  }
  if (corrections > 0) toast_('Đã loại tên không thuộc danh sách hoặc đăng ký trùng.', 'Đã kiểm tra', 7);
  return true;
}

function processGeneralCompletion_(ss, e) {
  const sh = ss.getSheetByName(LAB.SHEETS.GENERAL_CLEANING);
  const start = Math.max(e.range.getRow(), LAB.FIRST_GENERAL_ROW);
  const end = e.range.getLastRow();
  const editor = getEventUserEmail_(e);
  const now = new Date();

  if (!editor || !isAdminEmail_(editor)) {
    for (let row = start; row <= end; row++) {
      sh.getRange(row, LAB.GENERAL_COL.DONE).setValue(false);
      sh.getRange(row, LAB.GENERAL_COL.INSPECTOR, 1, 2).clearContent();
    }
    toast_('Chỉ cán bộ quản trị mới được xác nhận hoàn thành.', 'Không đủ quyền', 7);
    return;
  }

  for (let row = start; row <= end; row++) {
    const checked = sh.getRange(row, LAB.GENERAL_COL.DONE).getValue() === true;
    const student = normalizeLabel_(sh.getRange(row, LAB.GENERAL_COL.SV).getValue());
    if (checked && !student) {
      sh.getRange(row, LAB.GENERAL_COL.DONE).setValue(false);
      sh.getRange(row, LAB.GENERAL_COL.INSPECTOR, 1, 2).clearContent();
      toast_('Dòng chưa có sinh viên nên không thể xác nhận hoàn thành.', 'Thiếu sinh viên', 5);
    } else if (checked) {
      sh.getRange(row, LAB.GENERAL_COL.INSPECTOR).setValue(editor);
      sh.getRange(row, LAB.GENERAL_COL.CHECK_TIME).setValue(now).setNumberFormat('dd/MM/yyyy HH:mm');
    } else {
      sh.getRange(row, LAB.GENERAL_COL.INSPECTOR, 1, 2).clearContent();
    }
  }
}

function buildDynamicGeneralTasks_(studentCount) {
  const total = Math.max(0, Number(studentCount) || 0);
  if (total === 0) return [];

  const base = getBaseGeneralCleaningTasks_();
  const areaOrder = ['Lab 4.04', 'Lab 4.03', 'Phòng phụ'];
  const groups = areaOrder.map(function (area) {
    return {
      area: area,
      tasks: base.filter(function (task) { return task.area === area; })
    };
  });
  const allocations = allocateByWeights_(
    total,
    groups.map(function (group) { return group.tasks.length; })
  );
  const result = [];

  groups.forEach(function (group, groupIndex) {
    const slots = allocations[groupIndex];
    if (slots <= 0) return;

    if (slots <= group.tasks.length) {
      const sizes = distributeTotal_(
        group.tasks.length,
        slots,
        Array.from({ length: slots }, function (_, i) { return i; })
      );
      let offset = 0;
      sizes.forEach(function (size) {
        const bundle = group.tasks.slice(offset, offset + size);
        offset += size;
        result.push({
          area: group.area,
          job: bundle.map(function (task) { return task.job; }).join(' + '),
          requirement: bundle.length === 1
            ? bundle[0].requirement
            : 'Hoàn thành đầy đủ ' + bundle.length + ' hạng mục được giao; phối hợp an toàn và báo cán bộ nếu có thiết bị/mẫu cản trở.'
        });
      });
    } else {
      group.tasks.forEach(function (task) { result.push(task); });
      for (let extra = group.tasks.length; extra < slots; extra++) {
        result.push({
          area: group.area,
          job: 'Hỗ trợ tổng vệ sinh khu vực ' + group.area + ' - vị trí ' + (extra - group.tasks.length + 1),
          requirement: 'Phối hợp với các vị trí trong cùng khu vực theo hướng dẫn của cán bộ; không tự ý thao tác thiết bị.'
        });
      }
    }
  });

  return result;
}

function allocateByWeights_(total, weights) {
  if (total <= 0) return weights.map(function () { return 0; });
  const weightSum = weights.reduce(function (sum, value) { return sum + value; }, 0);
  if (weightSum <= 0) return distributeTotal_(total, weights.length, weights.map(function (_, i) { return i; }));

  const raw = weights.map(function (weight) { return total * weight / weightSum; });
  const result = raw.map(function (value) { return Math.floor(value); });
  let remainder = total - result.reduce(function (sum, value) { return sum + value; }, 0);
  const order = raw
    .map(function (value, index) { return { index: index, fraction: value - Math.floor(value) }; })
    .sort(function (a, b) { return b.fraction - a.fraction || a.index - b.index; });
  for (let i = 0; i < remainder; i++) result[order[i].index]++;
  return result;
}

function getBaseGeneralCleaningTasks_() {
  return [
    { area: 'Lab 4.04', job: 'Dãy 1 bên trái: lau bàn + lau kệ 2 tầng', requirement: 'Lau sạch mặt bàn và hai tầng kệ; không tự ý di chuyển mẫu/thiết bị.' },
    { area: 'Lab 4.04', job: 'Dãy 1 bên trái: lau thiết bị bên ngoài', requirement: 'Chỉ lau bên ngoài; không mở máy, tháo lắp hoặc chỉnh nút.' },
    { area: 'Lab 4.04', job: 'Dãy 2 ở giữa: lau bàn + lau kệ 2 tầng', requirement: 'Lau sạch mặt bàn và hai tầng kệ; không tự ý di chuyển mẫu/thiết bị.' },
    { area: 'Lab 4.04', job: 'Dãy 2 ở giữa: lau thiết bị bên ngoài', requirement: 'Chỉ lau bên ngoài; không mở máy, tháo lắp hoặc chỉnh nút.' },
    { area: 'Lab 4.04', job: 'Dãy 3 bên phải: lau bàn + lau kệ 2 tầng', requirement: 'Lau sạch mặt bàn và hai tầng kệ; không tự ý di chuyển mẫu/thiết bị.' },
    { area: 'Lab 4.04', job: 'Dãy 3 bên phải: lau thiết bị bên ngoài', requirement: 'Chỉ lau bên ngoài; không mở máy, tháo lắp hoặc chỉnh nút.' },
    { area: 'Lab 4.04', job: 'Vệ sinh bồn rửa số 1', requirement: 'Làm sạch bồn rửa, vòi nước và khu vực xung quanh.' },
    { area: 'Lab 4.04', job: 'Vệ sinh bồn rửa số 2', requirement: 'Làm sạch bồn rửa, vòi nước và khu vực xung quanh.' },
    { area: 'Lab 4.04', job: 'Vệ sinh bồn rửa số 3', requirement: 'Làm sạch bồn rửa, vòi nước và khu vực xung quanh.' },
    { area: 'Lab 4.04', job: 'Lau 2 cửa + tay nắm cửa', requirement: 'Lau sạch cửa, tay nắm và các bề mặt thường tiếp xúc.' },
    { area: 'Lab 4.04', job: 'Quét/lau sàn khu vực chung', requirement: 'Quét sạch bụi/rác, lau sàn; không để sàn quá ướt.' },

    { area: 'Lab 4.03', job: 'Dãy 1 bên trái: lau bàn + lau kệ 2 tầng', requirement: 'Lau sạch mặt bàn và hai tầng kệ; sắp xếp gọn vật dụng.' },
    { area: 'Lab 4.03', job: 'Dãy 2 ở giữa: lau bàn + lau kệ 2 tầng', requirement: 'Lau sạch mặt bàn và hai tầng kệ; sắp xếp gọn vật dụng.' },
    { area: 'Lab 4.03', job: 'Dãy 3 bên phải: lau bàn + lau kệ 2 tầng', requirement: 'Lau sạch mặt bàn và hai tầng kệ; sắp xếp gọn vật dụng.' },
    { area: 'Lab 4.03', job: 'Lau các thiết bị bên ngoài', requirement: 'Chỉ lau bên ngoài; không mở máy, tháo lắp hoặc chỉnh nút.' },
    { area: 'Lab 4.03', job: 'Vệ sinh bồn rửa số 1', requirement: 'Làm sạch bồn rửa, vòi nước và khu vực xung quanh.' },
    { area: 'Lab 4.03', job: 'Vệ sinh bồn rửa số 2', requirement: 'Làm sạch bồn rửa, vòi nước và khu vực xung quanh.' },
    { area: 'Lab 4.03', job: 'Vệ sinh bồn rửa số 3', requirement: 'Làm sạch bồn rửa, vòi nước và khu vực xung quanh.' },
    { area: 'Lab 4.03', job: 'Lau 2 cửa + tay nắm cửa', requirement: 'Lau sạch cửa, tay nắm và các bề mặt thường tiếp xúc.' },
    { area: 'Lab 4.03', job: 'Quét/lau sàn khu vực chung', requirement: 'Quét sạch bụi/rác, lau sàn; không để sàn quá ướt.' },

    { area: 'Phòng phụ', job: 'Dãy trái: lau bàn + lau thiết bị', requirement: 'Lau bàn và bên ngoài thiết bị; không tự ý di chuyển thiết bị.' },
    { area: 'Phòng phụ', job: 'Dãy phải: lau bàn + lau thiết bị', requirement: 'Lau bàn và bên ngoài thiết bị; không tự ý di chuyển thiết bị.' },
    { area: 'Phòng phụ', job: 'Lau/sắp xếp khu vực để dụng cụ', requirement: 'Lau sạch và sắp xếp dụng cụ đúng vị trí.' },
    { area: 'Phòng phụ', job: 'Vệ sinh bồn rửa số 1', requirement: 'Làm sạch bồn rửa, vòi nước và khu vực xung quanh.' },
    { area: 'Phòng phụ', job: 'Vệ sinh bồn rửa số 2', requirement: 'Làm sạch bồn rửa, vòi nước và khu vực xung quanh.' },
    { area: 'Phòng phụ', job: 'Lau cửa + quét/lau sàn', requirement: 'Lau cửa, tay nắm; quét và lau sàn phòng phụ.' }
  ];
}

function colorGeneralCleaningRows_(sh, rowCount) {
  if (rowCount <= 0) return;
  const areas = sh.getRange(LAB.FIRST_GENERAL_ROW, LAB.GENERAL_COL.AREA, rowCount, 1).getValues();
  const colors = {
    'Lab 4.04': '#d9ead3',
    'Lab 4.03': '#cfe2f3',
    'Phòng phụ': '#eadcf8'
  };
  const backgrounds = areas.map(function (row) {
    return Array(10).fill(colors[row[0]] || '#ffffff');
  });
  sh.getRange(LAB.FIRST_GENERAL_ROW, 1, rowCount, 10).setBackgrounds(backgrounds);
  sh.getRange(LAB.FIRST_GENERAL_ROW, LAB.GENERAL_COL.SV, rowCount, 1).setBackground('#fff2cc');
  sh.getRange(LAB.FIRST_GENERAL_ROW, LAB.GENERAL_COL.DONE, rowCount, 1).setBackground('#d9ead3');
}

