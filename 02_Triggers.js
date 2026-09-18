/**********************************************************************
 * TRIGGER
 **********************************************************************/

function createLabCleaningTriggers() {
  runManualAction_('createLabCleaningTriggers', function () {
    const ss = getSpreadsheet_();
    assertAdmin_(ss);
    createLabCleaningTriggersInternal_(ss);
    toast_('Đã tạo trigger chuẩn bị lịch 15:30, mở đăng ký 16:00, đồng bộ nền và dự phòng 17:00.', 'Trigger', 8);
  });
}

function deleteLabCleaningTriggers() {
  runManualAction_('deleteLabCleaningTriggers', function () {
    const ss = getSpreadsheet_();
    assertAdmin_(ss);
    deleteLabCleaningTriggersInternal_();
    toast_('Đã xóa các trigger do hệ thống này quản lý.', 'Trigger', 5);
  });
}

function createLabCleaningTriggersInternal_(ss) {
  deleteLabCleaningTriggersInternal_();

  ScriptApp.newTrigger(LAB.TRIGGER.EDIT)
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ScriptApp.newTrigger(LAB.TRIGGER.PREPARE)
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(LAB.PREPARE_HOUR)
    .nearMinute(LAB.PREPARE_MINUTE)
    .inTimezone(LAB.TIMEZONE)
    .create();

  ScriptApp.newTrigger(LAB.TRIGGER.OPEN)
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(LAB.RESET_HOUR)
    .nearMinute(0)
    .inTimezone(LAB.TIMEZONE)
    .create();

  ScriptApp.newTrigger(LAB.TRIGGER.RECOVERY)
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(LAB.RECOVERY_HOUR)
    .nearMinute(0)
    .inTimezone(LAB.TIMEZONE)
    .create();

  ScriptApp.newTrigger(LAB.TRIGGER.HEALTH)
    .timeBased()
    .everyDays(1)
    .atHour(LAB.HEALTH_CHECK_HOUR)
    .nearMinute(15)
    .inTimezone(LAB.TIMEZONE)
    .create();

  ScriptApp.newTrigger(LAB.TRIGGER.SYNC)
    .timeBased()
    .everyMinutes(5)
    .create();
}

function deleteLabCleaningTriggersInternal_() {
  const managed = [
    LAB.TRIGGER.EDIT,
    LAB.TRIGGER.PREPARE,
    LAB.TRIGGER.OPEN,
    LAB.TRIGGER.RECOVERY,
    LAB.TRIGGER.HEALTH,
    LAB.TRIGGER.SYNC,
    'weeklyResetLabCleaning'
  ];
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (managed.indexOf(trigger.getHandlerFunction()) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function handleLabCleaningEdit(e) {
  if (!e || !e.range) return;
  try {
    withDocumentLock_(function () {
      dispatchEdit_(e);
    });
  } catch (err) {
    revertEditedRange_(e);
    setDocumentProperty_(LAB.PROP.STATUS_DIRTY, 'YES');
    const busy = readableError_(err).indexOf('Hệ thống đang xử lý') === 0;
    if (busy) {
      console.warn('handleLabCleaningEdit: ' + readableError_(err));
    } else {
      safeLogError_('handleLabCleaningEdit', err);
    }
    toast_(
      busy
        ? 'Có nhiều bạn đăng ký cùng lúc. Thay đổi vừa rồi đã được hoàn tác; vui lòng chọn lại sau vài giây.'
        : 'Thay đổi vừa rồi đã được hoàn tác vì hệ thống gặp lỗi: ' + readableError_(err),
      'Chưa ghi nhận',
      9
    );
  }
}

function dispatchEdit_(e) {
  const ss = e.source || getSpreadsheet_();
  const sh = e.range.getSheet();
  const name = sh.getName();
  if (!isEditEventCurrent_(e)) return;

  if (
    name === LAB.SHEETS.STUDENTS &&
    e.range.getLastRow() >= LAB.FIRST_STUDENT_ROW &&
    rangesOverlapColumns_(e.range, 2, 4)
  ) {
    handleRosterEdit_(ss, e);
    return;
  }

  if (name === LAB.SHEETS.CONFIG && rangeTouchesCell_(e.range, 3, 2)) {
    handleConfigWeekModeEdit_(ss, e);
    return;
  }

  if (name === LAB.SHEETS.CONFIG && rangeTouchesCell_(e.range, 4, 2)) {
    handleConfigAssignmentModeEdit_(ss, e);
    return;
  }

  if (name === LAB.SHEETS.REGISTER) {
    handleRegisterEdit_(ss, e);
    return;
  }

  if (name === LAB.SHEETS.GENERAL_CLEANING) {
    handleGeneralCleaningEdit_(ss, e);
  }
}

