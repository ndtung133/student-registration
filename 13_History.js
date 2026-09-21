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

