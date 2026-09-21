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

