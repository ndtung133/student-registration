
/**********************************************************************
 * MENU VÀ CÁC HÀM CÔNG KHAI
 **********************************************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Vệ sinh lab')
    .addItem('Cài đặt / sửa chữa toàn bộ hệ thống', 'setupLabCleaningSystem')
    .addSeparator()
    .addItem('Mở đăng ký tuần này ngay', 'moDangKyTuanNayNgay')
    .addItem('Áp dụng SV mới vào tuần hiện tại', 'apDungSinhVienMoiVaoTuanHienTai')
    .addItem('Khóa đăng ký ca', 'khoaDangKy')
    .addItem('Mở lại đăng ký ca', 'moLaiDangKy')
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('Loại tuần')
        .addItem('Tuần thường: Thứ 2-Thứ 6', 'taoTuanThuong')
        .addItem('Tuần có tổng vệ sinh Thứ 6', 'taoTuanCoTongVeSinhThu6')
        .addItem('Tạo lại bảng tổng vệ sinh', 'taoBangTongVeSinhThu6')
    )
    .addSeparator()
    .addItem('Reset sang tuần kế tiếp ngay', 'manualResetLabCleaning')
    .addItem('Kiểm tra và tự sửa lỗi', 'healthCheckLabCleaning')
    .addItem('Kiểm thử logic phân ca', 'testLabCleaningLogic')
    .addSeparator()
    .addItem('Tạo lại trigger tự động', 'createLabCleaningTriggers')
    .addItem('Xóa trigger tự động', 'deleteLabCleaningTriggers')
    .addToUi();
}

function setupLabCleaningSystem() {
  runManualAction_('setupLabCleaningSystem', function () {
    const ss = bindAndGetSpreadsheet_();
    bootstrapAdmins_(ss);
    assertAdmin_(ss);

    const preservedRoster = readRosterForPreservation_(
      ss.getSheetByName(LAB.SHEETS.STUDENTS)
    );

    archiveCurrentWeek_(ss, 'Ảnh chụp trước khi cài đặt/sửa chữa', true);
    ensureAllSheets_(ss);
    removeAllManagedProtections_(ss);
    ss.setSpreadsheetTimeZone(LAB.TIMEZONE);

    if (!getDocumentProperty_(LAB.PROP.WEEK_MODE)) {
      setWeekMode_(LAB.MODE_REGULAR);
    }
    setAssignmentMode_(LAB.ASSIGN_SELF);
    setWorkflowPhase_(LAB.PHASE_CLOSED);

    buildStudentSheet_(ss, preservedRoster);
    const rosterValidation = validateRoster_(ss, true);
    const students = rosterValidation.valid
      ? rosterValidation.students
      : rosterValidation.eligibleStudents;
    const existingMonday = getRegisterWeekMonday_(ss);
    const monday = existingMonday || getExpectedOpenMonday_(new Date());

    buildCompleteWeek_(ss, monday, students, 'SETUP');
    if (!rosterValidation.valid) {
      setDocumentProperty_(LAB.PROP.PENDING_ROSTER_HASH, rosterValidation.hash);
    }
    createLabCleaningTriggersInternal_(ss);
    applyAllProtections_(ss);
    setActiveSheetSafe_(ss, LAB.SHEETS.HOME);

    logSystem_('INFO', 'setupLabCleaningSystem',
      'Cài đặt/sửa chữa hệ thống thành công',
      'Số SV=' + students.length + '; tuần=' + weekKey_(monday));
    toast_(
      rosterValidation.valid
        ? 'Đã cài đặt. Sinh viên khai danh sách một lần; Chủ nhật khoảng 16:00 hệ thống tự mở đăng ký tuần mới.'
        : 'Đã cài đặt. Các dòng thiếu hoặc trùng MSSV - họ tên cần được cán bộ sửa; dòng hợp lệ sẽ được dùng khi mở lịch.',
      'Hoàn tất',
      10
    );
  });
}

function syncRosterAndSchedule() {
  moDangKyTuanNayNgay();
}

/** Nâng cấp tại chỗ, không reset tuần và không ghi vào các ô tên đã đăng ký. */
function capNhatDangKyKhongEmail() {
  runManualAction_('capNhatDangKyKhongEmail', function () {
    const ss = getSpreadsheet_();
    assertAdmin_(ss);
    if (!ss.getSheetByName(LAB.SHEETS.STUDENTS) || !ss.getSheetByName(LAB.SHEETS.REGISTER)) {
      throw new Error('Thiếu bảng danh sách hoặc bảng đăng ký. Hàm này chỉ nâng cấp file đã cài đặt.');
    }
    updateStudentPolicyGuides_(ss);
    refreshRosterViews_(ss);
    refreshRegistrationStateFast_(ss);
    applyRosterProtectionsOnly_(ss, true);
    applyRegistrationAccessProtections_(ss, true);
    ensureManagedTriggers_();
    SpreadsheetApp.flush();
    toast_('Đã bỏ đối chiếu email sinh viên. Tên đã đăng ký được giữ nguyên; tải lại Sheet để sử dụng.', 'Cập nhật hoàn tất', 9);
  });
}

function getRosterGuideText_() {
  return 'Mỗi SV nhập MSSV - họ tên vào một ô trống cột B đúng 1 lần cho cả học kỳ. Không cần email. Cần sửa hoặc xóa tên thì nhờ cán bộ; thay đổi danh sách áp dụng vào lịch tuần mới.';
}

function updateStudentPolicyGuides_(ss) {
  const roster = ss.getSheetByName(LAB.SHEETS.STUDENTS);
  if (roster) {
    roster.getRange('C2').setValue('Email (tùy chọn, không đối chiếu)');
    roster.getRange('B2').setNote(getRosterGuideText_());
  }
  const reg = ss.getSheetByName(LAB.SHEETS.REGISTER);
  if (reg) reg.getRange('A3').setValue(getRegisterGuideText_());
  const home = ss.getSheetByName(LAB.SHEETS.HOME);
  if (home) {
    home.getRange('B13').setValue(getRosterGuideText_());
    home.getRange('B15').setValue('SV tự chọn tên tại cột G; mỗi tên chỉ một ca. Không đối chiếu email nên không ngăn được chọn tên hộ. Cần đổi/hủy ca thì nhờ cán bộ.');
  }
  const rules = ss.getSheetByName(LAB.SHEETS.RULES);
  if (rules) rules.getRange('B12').setValue(getRosterGuideText_());
  const config = ss.getSheetByName(LAB.SHEETS.CONFIG);
  if (config) config.getRange('B7').setValue(LAB.VERSION);
  const system = ss.getSheetByName(LAB.SHEETS.SYSTEM);
  if (system && hasAppliedRosterSnapshot_(ss)) system.getRange('E2').setValue(LAB.VERSION);
}

function moDanhSachSinhVien() {
  runManualAction_('moDanhSachSinhVien', function () {
    const ss = getSpreadsheet_();
    assertAdmin_(ss);
    setActiveSheetSafe_(ss, LAB.SHEETS.STUDENTS);
    toast_('Danh sách luôn mở ở các dòng trống. Mỗi SV chỉ khai MSSV - họ tên một lần; không cần email.', 'Danh sách học kỳ', 8);
  });
}

function chotDanhSachVaMoDangKy() {
  moDangKyTuanNayNgay();
}

function moDangKyTuanNayNgay() {
  runManualAction_('moDangKyTuanNayNgay', function () {
    const ss = getSpreadsheet_();
    assertAdmin_(ss);
    const validation = validateRoster_(ss, true);
    if (!validation.valid) throw new Error(validation.message);
    if (validation.students.length === 0) throw new Error('Danh sách chưa có sinh viên.');

    const started = hasWeekActivity_(ss);

    if (started && !confirm_(
      'Tạo lại đăng ký tuần này?',
      'Đang có dữ liệu đăng ký. Hệ thống sẽ lưu lịch sử rồi tạo lại bảng theo danh sách học kỳ hiện tại. Tiếp tục?'
    )) return;

    if (started) {
      archiveCurrentWeek_(ss, 'Ảnh chụp trước khi đồng bộ thủ công', true);
    }

    setAssignmentMode_(LAB.ASSIGN_SELF);
    const students = validation.students;
    const monday = getRegisterWeekMonday_(ss) || getExpectedOpenMonday_(new Date());
    rebuildWeekThenOpen_(ss, monday, students, 'MANUAL_OPEN_REGISTRATION');
    logSystem_('INFO', 'moDangKyTuanNayNgay', 'Đã tạo lịch và mở đăng ký thủ công', 'Số SV=' + students.length);
    setActiveSheetSafe_(ss, LAB.SHEETS.REGISTER);
    toast_('Đã mở đăng ký cho ' + students.length + ' SV. Mỗi bạn tự chọn tên của mình; mỗi tên chỉ được một ca, không đối chiếu email.', 'Đã mở đăng ký', 10);
  });
}
function apDungSinhVienMoiVaoTuanHienTai() {
  runManualAction_('apDungSinhVienMoiVaoTuanHienTai', function () {
    const ss = getSpreadsheet_();
    assertAdmin_(ss);

    applyNewRosterStudentsToCurrentWeek_(ss);

    setActiveSheetSafe_(ss, LAB.SHEETS.REGISTER);
  });
}

function chonTuDongPhanCong() {
  throw new Error('Bản này chỉ dùng quy trình sinh viên tự đăng ký, không tự động phân công.');
}

function chonTuDangKy() {
  chotDanhSachVaMoDangKy();
}

function khoaDangKy() {
  runManualAction_('khoaDangKy', function () {
    const ss = getSpreadsheet_();
    assertAdmin_(ss);
    setWorkflowPhase_(LAB.PHASE_CLOSED);
    applyRegistrationAccessProtections_(ss);
    updateWorkflowBanner_(ss);
    toast_('Đã khóa đăng ký. Chỉ cán bộ còn quyền kiểm tra và tick hoàn thành.', 'Đã khóa', 7);
  });
}

function moLaiDangKy() {
  runManualAction_('moLaiDangKy', function () {
    const ss = getSpreadsheet_();
    assertAdmin_(ss);
    const students = getCurrentWeekStudents_(ss);
    const reg = ss.getSheetByName(LAB.SHEETS.REGISTER);
    if (getRegisterRowCount_(reg) !== students.length) {
      throw new Error('Số slot tuần này đang lỗi. Hãy chạy “Kiểm tra và tự sửa lỗi” hoặc “Mở đăng ký tuần này ngay”.');
    }
    setWorkflowPhase_(LAB.PHASE_REGISTRATION_OPEN);
    setRegistrationFullLocked_(false);
    applyRegistrationAccessProtections_(ss);
    updateWorkflowBanner_(ss);
    toast_('Đã mở lại đăng ký ca.', 'Đã mở', 6);
  });
}

function taoTuanThuong() {
  changeWeekModeFromMenu_(LAB.MODE_REGULAR);
}

function taoTuanCoTongVeSinhThu6() {
  changeWeekModeFromMenu_(LAB.MODE_GENERAL_CLEANING);
}

function taoBangTongVeSinhThu6() {
  runManualAction_('taoBangTongVeSinhThu6', function () {
    const ss = getSpreadsheet_();
    assertAdmin_(ss);
    const students = getCurrentWeekStudents_(ss);
    const monday = getRegisterWeekMonday_(ss) || getExpectedOpenMonday_(new Date());
    buildGeneralCleaningSheet_(ss, monday, students);
    applyAllProtections_(ss);
    toast_('Đã tạo bảng tổng vệ sinh cho ' + students.length + ' sinh viên.', 'Hoàn tất', 6);
  });
}

function manualResetLabCleaning() {
  runManualAction_('manualResetLabCleaning', function () {
    const ss = getSpreadsheet_();
    assertAdmin_(ss);

    if (!confirm_(
      'Reset sang tuần kế tiếp?',
      'Hệ thống sẽ lưu lịch tuần hiện tại, lấy danh sách SV mới nhất và tạo lịch cho tuần kế tiếp. Tiếp tục?'
    )) return;

    resetForWeek_(ss, getNextMonday_(new Date()), 'MANUAL_RESET', true, LAB.MODE_REGULAR);
    toast_('Đã mở lịch cho tuần kế tiếp.', 'Reset hoàn tất', 6);
  });
}

function checkAndLockIfFull() {
  runManualAction_('checkAndLockIfFull', function () {
    const ss = getSpreadsheet_();
    assertAdmin_(ss);
    refreshAllStatuses_(ss, true);
    toast_('Đã kiểm tra trạng thái và cập nhật khóa.', 'Hoàn tất', 5);
  });
}

function unlockRegistration() {
  runManualAction_('unlockRegistration', function () {
    const ss = getSpreadsheet_();
    assertAdmin_(ss);
    setWorkflowPhase_(LAB.PHASE_REGISTRATION_OPEN);
    setRegistrationFullLocked_(false);
    applyRegistrationAccessProtections_(ss);
    updateWorkflowBanner_(ss);
    toast_('Đã mở lại vùng sinh viên đăng ký.', 'Đã mở khóa', 5);
  });
}

function weeklyPrepareLabCleaning() {
  runTriggeredAction_('weeklyPrepareLabCleaning', function () {
    const ss = getSpreadsheet_();
    resetForWeek_(ss, getNextMonday_(new Date()), 'WEEKLY_PREPARE', false, LAB.MODE_REGULAR, false);
  });
}

function weeklyOpenLabCleaning() {
  runTriggeredAction_('weeklyOpenLabCleaning', function () {
    const ss = getSpreadsheet_();
    const targetMonday = getNextMonday_(new Date());
    if (!openPreparedRegistration_(ss, targetMonday)) {
      resetForWeek_(ss, targetMonday, 'WEEKLY_OPEN_FALLBACK', false, LAB.MODE_REGULAR, true);
    }
  });
}

// Tên cũ được giữ để trigger cũ không làm hỏng hệ thống trước khi được tạo lại.
function weeklyResetLabCleaning() {
  weeklyOpenLabCleaning();
}

function sundayRecoveryLabCleaning() {
  runTriggeredAction_('sundayRecoveryLabCleaning', function () {
    const ss = getSpreadsheet_();
    resetForWeek_(ss, getNextMonday_(new Date()), 'SUNDAY_RECOVERY', false, LAB.MODE_REGULAR, true);
  });
}

function refreshLabCleaningViews() {
  const dirtyState = getDocumentProperty_(LAB.PROP.STATUS_DIRTY);
  const rosterDirty = getDocumentProperty_(LAB.PROP.ROSTER_DIRTY) === 'YES';
  if (dirtyState !== 'YES' && dirtyState !== 'ACCESS' && !rosterDirty) return;
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const ss = getSpreadsheet_();
    if (rosterDirty) refreshRosterViews_(ss);
    if (dirtyState === 'YES' || dirtyState === 'ACCESS') {
      refreshAllStatuses_(ss, dirtyState === 'ACCESS');
    }
  } catch (err) {
    safeLogError_('refreshLabCleaningViews', err);
  } finally {
    lock.releaseLock();
  }
}

function healthCheckLabCleaning(e) {
  const isManual = !(e && e.triggerUid);
  const runner = isManual ? runManualAction_ : runTriggeredAction_;

  runner('healthCheckLabCleaning', function () {
    const ss = getSpreadsheet_();
    if (isManual) assertAdmin_(ss);
    healthCheckInternal_(ss, isManual);
  });
}

function testLabCleaningLogic() {
  runManualAction_('testLabCleaningLogic', function () {
    const ss = getSpreadsheet_();
    assertAdmin_(ss);
    runSelfTests_();
    toast_('Tất cả kiểm thử phân ca từ 0 đến 300 SV đều đạt.', 'Kiểm thử đạt', 7);
  });
}

