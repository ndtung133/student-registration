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

