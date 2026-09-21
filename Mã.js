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

/**********************************************************************
 * SHEET DANH SÁCH SV
 **********************************************************************/

function buildStudentSheet_(ss, records) {
  const sh = ss.getSheetByName(LAB.SHEETS.STUDENTS);
  resetSheet_(sh);
  ensureSheetSize_(sh, Math.max(LAB.MIN_ROSTER_INPUT_ROWS + 2, records.length + 22), 7);
  sh.setTabColor('#1155cc');
  sh.setHiddenGridlines(true);

  sh.getRange('A1:G1')
    .merge()
    .setValue('DANH SÁCH SINH VIÊN DÙNG CHO CẢ HỌC KỲ')
    .setFontSize(15)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#1155cc')
    .setHorizontalAlignment('center');

  sh.getRange('A2:G2')
    .setValues([[
      'STT',
      'MSSV - Họ và tên (SV chỉ nhập 1 lần)',
      'Email (tùy chọn, không đối chiếu)',
      'Lớp/Nhóm (cán bộ bổ sung nếu cần)',
      'Tình trạng tuần',
      'Lịch trực',
      'Ghi chú'
    ]])
    .setFontWeight('bold')
    .setBackground('#cfe2f3')
    .setHorizontalAlignment('center')
    .setWrap(true);

  if (records.length > 0) {
    sh.getRange(LAB.FIRST_STUDENT_ROW, 1, records.length, 7)
      .setValues(records.map(function (record, index) {
        return [index + 1, record.label, record.email || '', record.group, '', '', record.note];
      }));
  }

  const inputRows = Math.max(LAB.MIN_ROSTER_INPUT_ROWS, records.length + 20);
  sh.getRange(2, 1, inputRows + 1, 7)
    .setBorder(true, true, true, true, true, true)
    .setVerticalAlignment('middle')
    .setWrap(true);

  sh.getRange(LAB.FIRST_STUDENT_ROW, 2, inputRows, 1).setBackground('#fff2cc');
  sh.getRange(LAB.FIRST_STUDENT_ROW, 3, inputRows, 2).setBackground('#f3f3f3');
  sh.getRange('A:A').setHorizontalAlignment('center');
  sh.getRange('E:E').setHorizontalAlignment('center');
  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 60);
  sh.setColumnWidth(2, 290);
  sh.setColumnWidth(3, 250);
  sh.setColumnWidth(4, 130);
  sh.setColumnWidth(5, 180);
  sh.setColumnWidth(6, 360);
  sh.setColumnWidth(7, 210);

  sh.getRange('B2').setNote(
    getRosterGuideText_()
  );
}

/**********************************************************************
 * SHEET CẤU HÌNH
 **********************************************************************/

function buildConfigSheet_(ss, students) {
  const sh = ss.getSheetByName(LAB.SHEETS.CONFIG);
  resetSheet_(sh);
  ensureSheetSize_(sh, 30, 6);
  sh.setTabColor('#674ea7');
  sh.setHiddenGridlines(true);

  const schedule = getScheduleConfig_(students.length, getWeekMode_());

  sh.getRange('A1:F1')
    .merge()
    .setValue('CẤU HÌNH SINH VIÊN TỰ ĐĂNG KÝ CA')
    .setFontSize(15)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#674ea7')
    .setHorizontalAlignment('center');

  sh.getRange('A2:A7').setValues([
    ['Số SV trong danh sách học kỳ'],
    ['Chế độ tuần'],
    ['Hình thức phân công'],
    ['Giai đoạn hiện tại'],
    ['Reset tự động'],
    ['Phiên bản hệ thống']
  ]).setFontWeight('bold').setBackground('#f3f0fa');

  sh.getRange('B2').setValue(countRosterEntries_(ss));
  sh.getRange('B3').setValue(weekModeText_(getWeekMode_()));
  sh.getRange('B4').setValue(assignmentModeText_(LAB.ASSIGN_SELF));
  sh.getRange('B5').setValue(workflowPhaseText_(getWorkflowPhase_()));
  sh.getRange('B6').setValue('Chủ nhật khoảng ' + twoDigits_(LAB.RESET_HOUR) + ':00');
  sh.getRange('B7').setValue(LAB.VERSION);

  sh.getRange('B3').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList([
        weekModeText_(LAB.MODE_REGULAR),
        weekModeText_(LAB.MODE_GENERAL_CLEANING)
      ], true)
      .setAllowInvalid(false)
      .build()
  );

  sh.getRange('B4').setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList([assignmentModeText_(LAB.ASSIGN_SELF)], true)
      .setAllowInvalid(false)
      .build()
  );

  sh.getRange('A8:F8')
    .setValues([['Thứ', 'Số SV/ngày', 'Lab 4.04', 'Lab 4.03', 'Phòng phụ', 'Ghi chú']])
    .setFontWeight('bold')
    .setBackground('#d9d2e9')
    .setHorizontalAlignment('center');

  if (schedule.length > 0) {
    sh.getRange(9, 1, schedule.length, 6).setValues(schedule.map(function (day) {
      return [day.day, day.total, day.lab404, day.lab403, day.phongPhu, day.note];
    }));
  }

  const totalRow = 9 + schedule.length;
  sh.getRange(totalRow, 1, 1, 6).setValues([[
    'Tổng kiểm tra',
    getScheduleTotal_(schedule),
    getTotalByField_(schedule, 'lab404'),
    getTotalByField_(schedule, 'lab403'),
    getTotalByField_(schedule, 'phongPhu'),
    getScheduleTotal_(schedule) === students.length ? 'Khớp danh sách áp dụng tuần này' : 'LỆCH LỊCH'
  ]]).setFontWeight('bold').setBackground('#e2f0d9');

  sh.getRange(totalRow + 2, 1, 1, 6)
    .merge()
    .setValue(getConfigNoteText_(students.length, schedule))
    .setBackground('#fff2cc')
    .setFontWeight('bold')
    .setWrap(true);

  sh.getRange('A2:B7').setBorder(true, true, true, true, true, true);
  sh.getRange(8, 1, schedule.length + 2, 6).setBorder(true, true, true, true, true, true);
  sh.getRange('A:F').setVerticalAlignment('middle').setWrap(true);
  sh.setColumnWidth(1, 180);
  sh.setColumnWidth(2, 145);
  sh.setColumnWidth(3, 105);
  sh.setColumnWidth(4, 105);
  sh.setColumnWidth(5, 105);
  sh.setColumnWidth(6, 520);
}

/**********************************************************************
 * SHEET ĐĂNG KÝ / PHÂN CÔNG TRỰC
 **********************************************************************/

function buildRegisterSheet_(ss, monday, students) {
  const sh = ss.getSheetByName(LAB.SHEETS.REGISTER);
  resetSheet_(sh);
  ensureSheetSize_(sh, Math.max(students.length + 15, 40), 14);
  sh.setTabColor('#cc0000');
  sh.setHiddenGridlines(true);

  const assignmentMode = getAssignmentMode_();
  const isAuto = assignmentMode === LAB.ASSIGN_AUTO;
  const schedule = getScheduleConfig_(students.length, getWeekMode_());
  const orderedStudents = isAuto
    ? orderStudentsForWeek_(students, monday, 'REGULAR')
    : [];
  const assignedAt = new Date();
  const pendingText = isAuto ? 'Đã phân công - chưa trực' : 'Đã đăng ký - chưa trực';
  const rows = [];
  let stt = 1;
  let studentIndex = 0;

  schedule.forEach(function (dayObj) {
    const date = addDays_(monday, dayObj.dayOffset);
    [
      ['Lab 4.04', dayObj.lab404],
      ['Lab 4.03', dayObj.lab403],
      ['Phòng phụ', dayObj.phongPhu]
    ].forEach(function (labEntry) {
      const labName = labEntry[0];
      const count = Number(labEntry[1] || 0);
      for (let position = 1; position <= count; position++) {
        const student = isAuto ? orderedStudents[studentIndex++] : null;
        rows.push([
          stt++,
          dayObj.day,
          date,
          labName,
          getTaskByLab_(labName, position, count),
          getPositionNameByLab_(labName, position, count),
          student ? student.label : '',
          student ? assignedAt : '',
          '',
          student ? pendingText : 'Còn trống',
          false,
          '',
          '',
          ''
        ]);
      }
    });
  });

  sh.getRange('A1:N1')
    .merge()
    .setValue(getRegisterTitle_())
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#cc0000')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 44);

  sh.getRange('A2').setValue('Tuần bắt đầu');
  sh.getRange('B2').setValue(monday).setNumberFormat('dd/MM/yyyy');
  const registrationOpen = getWorkflowPhase_() === LAB.PHASE_REGISTRATION_OPEN;
  sh.getRange('C2:N2').merge()
    .setValue(registrationOpen
      ? getOpenMessage_(isAuto ? students.length : 0, students.length)
      : 'HỆ THỐNG ĐANG CHUẨN BỊ LỊCH TUẦN MỚI - CHƯA MỞ ĐĂNG KÝ')
    .setBackground(registrationOpen && students.length > 0 && isAuto ? '#d9ead3' : '#cfe2f3')
    .setFontColor(registrationOpen && students.length > 0 && isAuto ? '#274e13' : '#073763')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  sh.getRange('A3:N3').merge()
    .setValue(getRegisterGuideText_())
    .setBackground('#f3f6f4')
    .setFontStyle('italic')
    .setHorizontalAlignment('center')
    .setWrap(true);

  const headers = [
    'STT', 'Thứ', 'Ngày', 'Khu vực lab', 'Việc cần làm', 'Vị trí',
    isAuto ? 'Sinh viên được phân công' : 'MSSV - Họ tên SV đăng ký',
    isAuto ? 'Thời điểm phân công' : 'Thời điểm đăng ký',
    'Ghi chú SV', 'Trạng thái', 'Đã trực xong?', 'Người kiểm tra',
    'Thời điểm kiểm tra', 'Ghi chú kiểm tra'
  ];

  sh.getRange(5, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#f4cccc')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  if (rows.length > 0) {
    sh.getRange(LAB.FIRST_REGISTER_ROW, 1, rows.length, headers.length).setValues(rows);
    sh.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.NGAY, rows.length, 1).setNumberFormat('dd/MM/yyyy');
    sh.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.TIME_DANGKY, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    sh.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.TIME_KIEMTRA, rows.length, 1).setNumberFormat('dd/MM/yyyy HH:mm');
    sh.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.DA_TRUC, rows.length, 1).insertCheckboxes();

    applyRegisterValidation_(ss, rows.length, students.length);
    applyAlternatingDayColors_(sh, rows.length);
    applyRegisterConditionalFormatting_(sh, rows.length);
    sh.getRange(5, 1, rows.length + 1, headers.length).setBorder(true, true, true, true, true, true);
  }

  sh.getRange('A2:N3').setBorder(true, true, true, true, true, true);
  sh.getRange('A:N').setVerticalAlignment('middle').setWrap(true);
  sh.getRange('A:C').setHorizontalAlignment('center');
  sh.getRange('F:F').setHorizontalAlignment('center');
  sh.getRange('H:N').setHorizontalAlignment('center');
  sh.setFrozenRows(5);

  const widths = [50, 75, 100, 115, 360, 220, 290, 150, 180, 190, 120, 190, 155, 220];
  widths.forEach(function (width, index) { sh.setColumnWidth(index + 1, width); });

  setRegistrationFullLocked_(isAuto && students.length > 0);
}

function applyRegisterValidation_(ss, rowCount, studentCount) {
  const reg = ss.getSheetByName(LAB.SHEETS.REGISTER);
  if (!reg || rowCount <= 0) return;
  const target = reg.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.SV, rowCount, 1);

  if (studentCount <= 0) {
    target.clearDataValidations();
    return;
  }

  const system = ss.getSheetByName(LAB.SHEETS.SYSTEM);
  const source = system.getRange(2, 1, studentCount, 1);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(source, true)
    .setAllowInvalid(false)
    .setHelpText('Chỉ chọn sinh viên có trong DANH SÁCH SV.')
    .build();
  target.setDataValidation(rule);
}

function applyRegisterConditionalFormatting_(sh, rowCount) {
  if (rowCount <= 0) return;
  const statusRange = sh.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.TRANGTHAI, rowCount, 1);
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Đã trực xong')
      .setBackground('#d9ead3')
      .setFontColor('#274e13')
      .setRanges([statusRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('chưa trực')
      .setBackground('#fff2cc')
      .setFontColor('#7f6000')
      .setRanges([statusRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Còn trống')
      .setBackground('#eeeeee')
      .setFontColor('#666666')
      .setRanges([statusRange])
      .build()
  ];
  sh.setConditionalFormatRules(rules);
}

/**********************************************************************
 * SHEET TRANG CHỦ
 **********************************************************************/

function buildHomeSheet_(ss, students) {
  const sh = ss.getSheetByName(LAB.SHEETS.HOME);
  resetSheet_(sh);
  ensureSheetSize_(sh, 25, 6);
  sh.setTabColor('#38761d');
  sh.setHiddenGridlines(true);

  const schedule = getScheduleConfig_(students.length, getWeekMode_());

  sh.getRange('A1:F1')
    .merge()
    .setValue('HỆ THỐNG SINH VIÊN TỰ ĐĂNG KÝ TRỰC VỆ SINH LAB')
    .setFontSize(18)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#38761d')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 46);

  sh.getRange('A3:B9').setValues([
    ['Cấu trúc lab', 'Lab 4.04 + Lab 4.03 + Phòng phụ'],
    ['Số SV trong danh sách học kỳ', ''],
    ['Hình thức', 'Sinh viên tự khai danh sách và tự đăng ký ca'],
    ['Thời gian trực', getHomeTimeText_()],
    ['Phân bổ hiện tại', getHomeScheduleText_(schedule)],
    ['Trạng thái', ''],
    ['Reset tự động', 'Chủ nhật hằng tuần khoảng ' + twoDigits_(LAB.RESET_HOUR) + ':00']
  ]);

  sh.getRange('B4').setValue(
    countRosterEntries_(ss) + ' sinh viên (thay đổi áp dụng Chủ nhật 16:00)'
  );
  sh.getRange('B8').setValue(workflowPhaseText_(getWorkflowPhase_()));
  sh.getRange('A3:A9').setFontWeight('bold').setBackground('#f3f6f4');
  sh.getRange('A3:B9').setBorder(true, true, true, true, true, true);

  sh.getRange('A11:F11')
    .merge()
    .setValue('CÁCH DÙNG NHANH')
    .setFontWeight('bold')
    .setBackground('#d9ead3')
    .setHorizontalAlignment('center');

  const steps = [
    ['1', 'Cán bộ chạy setup một lần rồi chia sẻ file với quyền Người chỉnh sửa cho sinh viên.'],
    ['2', getRosterGuideText_()],
    ['3', 'Không cần cán bộ chốt danh sách. Chủ nhật khoảng 16:00 hệ thống tự lấy danh sách hợp lệ, tạo đúng số slot và mở đăng ký tuần mới.'],
    ['4', 'SV tự chọn tên tại cột G; mỗi tên chỉ một ca. Không đối chiếu email nên không ngăn được chọn tên hộ. Cần đổi/hủy ca thì nhờ cán bộ.'],
    ['5', 'Khi đủ người, bảng tự khóa. Chỉ cán bộ quản trị được tick cột K và ghi nhận hoàn thành.'],
    ['6', 'Danh sách được dùng suốt học kỳ. SV thêm/bớt giữa tuần không làm đổi lịch đang chạy; hệ thống áp dụng vào Chủ nhật kế tiếp.']
  ];

  sh.getRange(12, 1, steps.length, 2).setValues(steps);
  sh.getRange(12, 1, steps.length, 1).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(11, 1, steps.length + 1, 6).setBorder(true, true, true, true, true, true);
  sh.getRange('A:F').setVerticalAlignment('middle').setWrap(true);
  sh.setColumnWidth(1, 165);
  sh.setColumnWidth(2, 720);
  sh.setColumnWidths(3, 4, 110);
}

/**********************************************************************
 * SHEET THEO DÕI - GHI GIÁ TRỊ TRỰC TIẾP, KHÔNG PHỤ THUỘC LOCALE
 **********************************************************************/

function buildDashboardSheet_(ss) {
  const sh = ss.getSheetByName(LAB.SHEETS.DASHBOARD);
  resetSheet_(sh);
  ensureSheetSize_(sh, 30, 8);
  sh.setTabColor('#f1c232');
  sh.setHiddenGridlines(true);

  sh.getRange('A1:H1')
    .merge()
    .setValue('THEO DÕI PHÂN CÔNG VÀ KIỂM TRA TRỰC VỆ SINH LAB')
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#bf9000')
    .setHorizontalAlignment('center');

  sh.getRange('A3:D3')
    .setValues([['Chỉ tiêu', 'Số lượng', 'Ghi chú', 'Trạng thái']])
    .setFontWeight('bold')
    .setBackground('#fff2cc')
    .setHorizontalAlignment('center');

  sh.getRange('A4:D9').setValues([
    ['Tổng slot của tuần', 0, 'Tự động bằng số SV đang áp dụng', ''],
    ['Đã có sinh viên', 0, 'Đã phân công hoặc đã đăng ký', ''],
    ['Còn trống', 0, 'Slot chưa có sinh viên', ''],
    ['Đã trực xong', 0, 'Đã được cán bộ xác nhận', ''],
    ['Có tên - chưa trực', 0, 'Đã có lịch nhưng chưa hoàn thành', ''],
    ['Chưa hoàn tất', 0, 'Bao gồm còn trống hoặc chưa trực', '']
  ]);

  sh.getRange('A11:C11')
    .setValues([['Thứ', 'Đã có sinh viên', 'Đã trực xong']])
    .setFontWeight('bold')
    .setBackground('#d9ead3')
    .setHorizontalAlignment('center');
  const days = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6'];
  sh.getRange(12, 1, days.length, 3).setValues(days.map(function (day) { return [day, 0, 0]; }));

  sh.getRange('E11:G11')
    .setValues([['Phòng', 'Đã có sinh viên', 'Đã trực xong']])
    .setFontWeight('bold')
    .setBackground('#cfe2f3')
    .setHorizontalAlignment('center');
  const labs = ['Lab 4.04', 'Lab 4.03', 'Phòng phụ'];
  sh.getRange(12, 5, labs.length, 3).setValues(labs.map(function (lab) { return [lab, 0, 0]; }));

  sh.getRange('A3:D9').setBorder(true, true, true, true, true, true);
  sh.getRange('A11:C16').setBorder(true, true, true, true, true, true);
  sh.getRange('E11:G14').setBorder(true, true, true, true, true, true);
  sh.getRange('A:G').setVerticalAlignment('middle').setWrap(true);
  sh.setColumnWidth(1, 180);
  sh.setColumnWidth(2, 110);
  sh.setColumnWidth(3, 290);
  sh.setColumnWidth(4, 250);
  sh.setColumnWidth(5, 140);
  sh.setColumnWidth(6, 120);
  sh.setColumnWidth(7, 120);

  const statusRange = sh.getRange('D4:D9');
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('OK')
      .setBackground('#d9ead3')
      .setRanges([statusRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Lệch')
      .setBackground('#f4cccc')
      .setRanges([statusRange])
      .build()
  ]);

  updateDashboardValues_(ss);
}

function updateDashboardValues_(ss) {
  const sh = ss.getSheetByName(LAB.SHEETS.DASHBOARD);
  const reg = ss.getSheetByName(LAB.SHEETS.REGISTER);
  if (!sh || !reg) return;

  const days = ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6'];
  const labs = ['Lab 4.04', 'Lab 4.03', 'Phòng phụ'];
  const dayCounts = {};
  const labCounts = {};
  days.forEach(function (day) { dayCounts[day] = { signed: 0, completed: 0 }; });
  labs.forEach(function (lab) { labCounts[lab] = { signed: 0, completed: 0 }; });

  const rowCount = getRegisterRowCount_(reg);
  let signed = 0;
  let completed = 0;
  if (rowCount > 0) {
    const data = reg.getRange(LAB.FIRST_REGISTER_ROW, 1, rowCount, 14).getValues();
    data.forEach(function (row) {
      const student = normalizeLabel_(row[LAB.COL.SV - 1]);
      if (!student) return;
      const done = row[LAB.COL.DA_TRUC - 1] === true;
      const day = normalizeLabel_(row[LAB.COL.THU - 1]);
      const lab = normalizeLabel_(row[LAB.COL.LAB - 1]);
      signed++;
      if (done) completed++;
      if (dayCounts[day]) {
        dayCounts[day].signed++;
        if (done) dayCounts[day].completed++;
      }
      if (labCounts[lab]) {
        labCounts[lab].signed++;
        if (done) labCounts[lab].completed++;
      }
    });
  }

  const appliedCount = getCurrentWeekStudents_(ss).length;
  const empty = Math.max(rowCount - signed, 0);
  const pending = Math.max(signed - completed, 0);
  const unfinished = Math.max(rowCount - completed, 0);

  sh.getRange('B4:B9').setValues([
    [rowCount], [signed], [empty], [completed], [pending], [unfinished]
  ]);
  sh.getRange('D4:D9').setValues([
    [rowCount === appliedCount ? 'OK' : 'Lệch lịch đang áp dụng'],
    [signed + '/' + rowCount],
    [empty === 0 ? 'Không còn slot' : 'Còn ' + empty + ' slot'],
    [completed + '/' + rowCount],
    [pending === 0 ? 'Không còn' : pending + ' bạn chưa trực'],
    [rowCount > 0 && completed === rowCount
      ? 'Tất cả đã trực xong'
      : 'Còn ' + unfinished + ' slot chưa hoàn tất']
  ]);
  sh.getRange(12, 2, days.length, 2).setValues(days.map(function (day) {
    return [dayCounts[day].signed, dayCounts[day].completed];
  }));
  sh.getRange(12, 6, labs.length, 2).setValues(labs.map(function (lab) {
    return [labCounts[lab].signed, labCounts[lab].completed];
  }));
}

/**********************************************************************
 * SHEET QUY ĐỊNH
 **********************************************************************/

function buildRulesSheet_(ss, students) {
  const sh = ss.getSheetByName(LAB.SHEETS.RULES);
  resetSheet_(sh);
  ensureSheetSize_(sh, 30, 6);
  sh.setTabColor('#6aa84f');
  sh.setHiddenGridlines(true);

  sh.getRange('A1:F1')
    .merge()
    .setValue('QUY ĐỊNH TRỰC VỆ SINH LAB')
    .setFontSize(16)
    .setFontWeight('bold')
    .setFontColor('#ffffff')
    .setBackground('#6aa84f')
    .setHorizontalAlignment('center');

  sh.getRange('A3:F3').merge()
    .setValue('Phạm vi công việc ngày thường')
    .setFontWeight('bold')
    .setBackground('#d9ead3');
  sh.getRange('A4:F8').setValues([
    ['1', 'Lab 4.04', 'Lau phòng hoặc quét/lau sàn theo vị trí được phân công.', '', '', ''],
    ['2', 'Lab 4.03', 'Lau phòng hoặc quét/lau sàn theo vị trí được phân công.', '', '', ''],
    ['3', 'Phòng phụ', 'Lau và quét/lau sàn phòng phụ theo vị trí được phân công.', '', '', ''],
    ['4', 'Không làm ngày thường', 'Không tự ý mở máy, tháo thiết bị, di chuyển mẫu hoặc hóa chất.', '', '', ''],
    ['5', 'Tổng vệ sinh', 'Có bảng công việc riêng; số vị trí tự đổi theo số sinh viên.', '', '', '']
  ]);

  sh.getRange('A10:F10').merge()
    .setValue('Nguyên tắc phân công và kiểm tra')
    .setFontWeight('bold')
    .setBackground('#fff2cc');

  const rules = [
    ['1', 'Số SV không cố định; lịch tuần này đang áp dụng ' + students.length + ' SV.', '', '', '', ''],
    ['2', getRosterGuideText_(), '', '', '', ''],
    ['3', 'Không cần chốt danh sách. Sau khi hệ thống mở lịch, mỗi sinh viên tự đăng ký đúng 1 ca tại cột G.', '', '', '', ''],
    ['4', 'Cán bộ tick cột K sau khi kiểm tra; sinh viên không tự ý tick hoàn thành.', '', '', '', ''],
    ['5', 'Chủ nhật khoảng ' + twoDigits_(LAB.RESET_HOUR) + ':00 hệ thống lưu lịch sử và tạo tuần mới.', '', '', '', ''],
    ['6', 'Danh sách dùng cho cả học kỳ. Thay đổi giữa tuần chỉ áp dụng vào lịch mới Chủ nhật kế tiếp.', '', '', '', ''],
    ['7', 'Tuần có tổng vệ sinh: Thứ 6 lúc 14:00, công việc được phân bổ động theo số SV.', '', '', '', '']
  ];
  sh.getRange(11, 1, rules.length, 6).setValues(rules);

  sh.getRange(3, 1, 15, 6).setBorder(true, true, true, true, true, true);
  sh.getRange('A:F').setVerticalAlignment('middle').setWrap(true);
  sh.setColumnWidth(1, 50);
  sh.setColumnWidth(2, 720);
  sh.setColumnWidth(3, 120);
}

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

/**********************************************************************
 * XỬ LÝ ĐĂNG KÝ / XÁC NHẬN HOÀN THÀNH
 **********************************************************************/

function handleRegisterEdit_(ss, e) {
  if (e.range.getLastRow() < LAB.FIRST_REGISTER_ROW) return;

  if (rangesOverlapColumns_(e.range, LAB.COL.SV, LAB.COL.GHICHU_SV)) {
    if (getWorkflowPhase_() !== LAB.PHASE_REGISTRATION_OPEN) {
      revertEditedRange_(e);
      toast_('Bảng đăng ký chưa được mở hoặc đã bị khóa.', 'Không thể đăng ký', 7);
      return;
    }
    if (sanitizeSelfRegistration_(ss, e) === false) return;
  }

  if (rangeTouchesColumn_(e.range, LAB.COL.DA_TRUC)) {
    processRegisterCompletion_(ss, e);
  }

  if (
    rangeTouchesColumn_(e.range, LAB.COL.SV) ||
    rangeTouchesColumn_(e.range, LAB.COL.DA_TRUC)
  ) {
    refreshRegistrationStateFast_(ss, e.range);
  }
}

function restoreAutoAssignedNames_(ss) {
  const sh = ss.getSheetByName(LAB.SHEETS.REGISTER);
  const students = getCurrentWeekStudents_(ss);
  const monday = getRegisterWeekMonday_(ss) || getExpectedOpenMonday_(new Date());
  const ordered = orderStudentsForWeek_(students, monday, 'REGULAR');
  const rowCount = getRegisterRowCount_(sh);

  if (rowCount !== ordered.length) {
    throw new Error('Số dòng lịch không khớp danh sách. Hãy dùng “Đồng bộ danh sách và phân công ngay”.');
  }

  if (rowCount === 0) return;
  sh.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.SV, rowCount, 1)
    .setValues(ordered.map(function (student) { return [student.label]; }));

  const timesRange = sh.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.TIME_DANGKY, rowCount, 1);
  const times = timesRange.getValues();
  const now = new Date();
  timesRange.setValues(times.map(function (row) { return [row[0] || now]; }))
    .setNumberFormat('dd/MM/yyyy HH:mm');
}

function sanitizeSelfRegistration_(ss, e) {
  const sh = ss.getSheetByName(LAB.SHEETS.REGISTER);
  const students = getCurrentWeekStudents_(ss);
  const rowCount = getRegisterRowCount_(sh);
  if (!sh || rowCount === 0 || !isEditEventCurrent_(e)) return false;

  // Không xác minh sinh viên theo email. Email chỉ quyết định quyền sửa của cán bộ.
  const admin = isAdminEmail_(getEventUserEmail_(e));
  const column = e.range.getColumn();
  if (!admin && (e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1 ||
      (column !== LAB.COL.SV && column !== LAB.COL.GHICHU_SV))) {
    revertEditedRange_(e);
    toast_('Mỗi lần chỉ chọn một ô tên ở cột G hoặc nhập một ô ghi chú ở cột I.', 'Đã hoàn tác', 7);
    return false;
  }

  const start = Math.max(e.range.getRow(), LAB.FIRST_REGISTER_ROW);
  const end = Math.min(e.range.getLastRow(), LAB.FIRST_REGISTER_ROW + rowCount - 1);
  if (start > end) {
    revertEditedRange_(e);
    return false;
  }

  if (!admin && column === LAB.COL.GHICHU_SV) {
    const label = sh.getRange(start, LAB.COL.SV).getValue();
    if (!findStudentByLabel_(students, label)) {
      revertEditedRange_(e);
      toast_('Hãy đăng ký tên trước khi nhập ghi chú.', 'Chưa có đăng ký', 6);
      return false;
    }
    // Không có xác minh danh tính nên không xác định được chủ nhân ghi chú.
    return true;
  }

  if (!admin && normalizeLabel_(e.oldValue)) {
    if (canonicalLabel_(e.oldValue) !== canonicalLabel_(e.range.getValue())) {
      revertEditedRange_(e);
      toast_('Ca này đã có tên. Muốn đổi hoặc hủy đăng ký thì nhờ cán bộ.', 'Không thể thay đổi', 8);
    }
    return false;
  }

  const allNames = sh.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.SV, rowCount, 1).getValues();
  const occupied = Object.create(null);
  allNames.forEach(function (row, index) {
    const sheetRow = LAB.FIRST_REGISTER_ROW + index;
    if (sheetRow >= start && sheetRow <= end) return;
    const key = canonicalLabel_(row[0]);
    if (key) occupied[key] = sheetRow;
  });

  if (!admin) {
    const chosen = normalizeLabel_(allNames[start - LAB.FIRST_REGISTER_ROW][0]);
    const student = chosen ? findStudentByLabel_(students, chosen) : null;
    if (!chosen) return false;
    if (!student || occupied[canonicalLabel_(student.label)]) {
      revertEditedRange_(e);
      toast_(
        student ? 'Tên này đã đăng ký một ca rồi.' : 'Tên chưa có trong danh sách áp dụng cho tuần này.',
        'Không thể đăng ký',
        8
      );
      return false;
    }
  }

  if (!isEditEventCurrent_(e)) return false;
  const now = new Date();
  let corrections = 0;
  for (let row = start; row <= end; row++) {
    const chosen = normalizeLabel_(allNames[row - LAB.FIRST_REGISTER_ROW][0]);
    if (!chosen) {
      clearRegistrationRowInput_(sh, row);
      continue;
    }

    const student = findStudentByLabel_(students, chosen);
    const key = student ? canonicalLabel_(student.label) : '';
    if (!student || occupied[key]) {
      clearRegistrationRowInput_(sh, row);
      corrections++;
      continue;
    }

    occupied[key] = row;
    if (chosen !== student.label) sh.getRange(row, LAB.COL.SV).setValue(student.label);
    if (rangeTouchesColumn_(e.range, LAB.COL.SV)) {
      sh.getRange(row, LAB.COL.TIME_DANGKY).setValue(now);
      sh.getRange(row, LAB.COL.DA_TRUC, 1, 3).setValues([[false, '', '']]);
    }
  }

  if (corrections > 0) toast_('Đã loại tên không thuộc danh sách hoặc bị trùng ca.', 'Đã kiểm tra', 7);
  return true;
}

function clearRegistrationRowInput_(sh, row) {
  sh.getRange(row, LAB.COL.SV, 1, 3).clearContent();
  sh.getRange(row, LAB.COL.DA_TRUC).setValue(false);
  sh.getRange(row, LAB.COL.NGUOI_KIEMTRA, 1, 3).clearContent();
}

function processRegisterCompletion_(ss, e) {
  const sh = ss.getSheetByName(LAB.SHEETS.REGISTER);
  const editor = normalizeEmail_(getEditorEmail_(e));
  const admins = getAdminEmails_();
  const identifiable = editor && editor !== 'không xác định';

  if (!identifiable || admins.length === 0 || admins.indexOf(editor) === -1) {
    const startUnauthorized = Math.max(e.range.getRow(), LAB.FIRST_REGISTER_ROW);
    const endUnauthorized = e.range.getLastRow();
    for (let row = startUnauthorized; row <= endUnauthorized; row++) {
      sh.getRange(row, LAB.COL.DA_TRUC).setValue(false);
      sh.getRange(row, LAB.COL.NGUOI_KIEMTRA, 1, 2).clearContent();
    }
    toast_('Chỉ cán bộ quản trị đăng nhập đúng email mới được xác nhận hoàn thành.', 'Không đủ quyền', 7);
    return;
  }

  const start = Math.max(e.range.getRow(), LAB.FIRST_REGISTER_ROW);
  const end = e.range.getLastRow();
  const now = new Date();
  const displayEditor = getEditorEmail_(e);

  for (let row = start; row <= end; row++) {
    const checked = sh.getRange(row, LAB.COL.DA_TRUC).getValue() === true;
    const student = normalizeLabel_(sh.getRange(row, LAB.COL.SV).getValue());

    if (checked && !student) {
      sh.getRange(row, LAB.COL.DA_TRUC).setValue(false);
      sh.getRange(row, LAB.COL.NGUOI_KIEMTRA, 1, 2).clearContent();
      toast_('Dòng chưa có sinh viên nên không thể xác nhận đã trực.', 'Thiếu sinh viên', 5);
    } else if (checked) {
      sh.getRange(row, LAB.COL.NGUOI_KIEMTRA).setValue(displayEditor);
      sh.getRange(row, LAB.COL.TIME_KIEMTRA).setValue(now).setNumberFormat('dd/MM/yyyy HH:mm');
    } else {
      sh.getRange(row, LAB.COL.NGUOI_KIEMTRA, 1, 2).clearContent();
    }
  }
}

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
