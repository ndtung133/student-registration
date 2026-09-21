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
