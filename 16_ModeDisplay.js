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

