/**********************************************************************
 * KHÓA, QUYỀN QUẢN TRỊ VÀ CHẠY AN TOÀN
 **********************************************************************/

function runManualAction_(name, action) {
  try {
    return withDocumentLock_(action);
  } catch (err) {
    safeLogError_(name, err);
    toast_('Lỗi: ' + readableError_(err), 'Không hoàn tất', 10);
    throw err;
  }
}

function runTriggeredAction_(name, action) {
  try {
    return withDocumentLock_(action);
  } catch (err) {
    safeLogError_(name, err);
    throw err;
  }
}

function withDocumentLock_(action) {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(60000)) {
    throw new Error('Hệ thống đang xử lý nhiều lượt đăng ký cùng lúc. Vui lòng thử lại sau vài giây.');
  }

  try {
    return action();
  } finally {
    lock.releaseLock();
  }
}

function bindAndGetSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Không tìm thấy Google Sheet đang gắn với dự án Apps Script.');
  PropertiesService.getScriptProperties().setProperty(LAB.PROP.SPREADSHEET_ID, ss.getId());
  return ss;
}

function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const storedId = props.getProperty(LAB.PROP.SPREADSHEET_ID);
  if (storedId) return SpreadsheetApp.openById(storedId);
  return bindAndGetSpreadsheet_();
}

function bootstrapAdmins_(ss) {
  const emails = {};

  LAB.ADMIN_EMAILS.forEach(function (email) {
    const normalized = normalizeEmail_(email);
    if (normalized) emails[normalized] = true;
  });

  try {
    const owner = ss.getOwner();
    if (owner) {
      const email = normalizeEmail_(owner.getEmail());
      if (email) emails[email] = true;
    }
  } catch (err) {}

  if (Object.keys(emails).length === 0) {
    const current = normalizeEmail_(getCurrentUserEmail_());
    if (current) emails[current] = true;
  }

  const previous = getAdminEmails_();
  previous.forEach(function (email) { emails[email] = true; });
  setDocumentProperty_(LAB.PROP.ADMIN_EMAILS, JSON.stringify(Object.keys(emails)));
}

function getAdminEmails_() {
  const result = {};
  LAB.ADMIN_EMAILS.forEach(function (email) {
    const normalized = normalizeEmail_(email);
    if (normalized) result[normalized] = true;
  });

  try {
    const stored = JSON.parse(getDocumentProperty_(LAB.PROP.ADMIN_EMAILS) || '[]');
    stored.forEach(function (email) {
      const normalized = normalizeEmail_(email);
      if (normalized) result[normalized] = true;
    });
  } catch (err) {}

  return Object.keys(result);
}

function assertAdmin_(ss) {
  bootstrapAdmins_(ss);
  const admins = getAdminEmails_();
  if (admins.length === 0) return;

  const current = normalizeEmail_(getCurrentUserEmail_());
  if (!current || admins.indexOf(current) === -1) {
    throw new Error('Chỉ quản trị viên hệ thống mới được chạy thao tác này.');
  }
}

function getCurrentUserEmail_() {
  try {
    const active = Session.getActiveUser().getEmail();
    if (active) return active;
  } catch (err) {}

  try {
    return Session.getEffectiveUser().getEmail() || '';
  } catch (err) {
    return '';
  }
}

function getEditorEmail_(e) {
  try {
    if (e && e.user) {
      const email = e.user.getEmail();
      if (email) return email;
    }
  } catch (err) {}

  try {
    const active = Session.getActiveUser().getEmail();
    if (active) return active;
  } catch (err) {}

  return 'Không xác định';
}

function getEventUserEmail_(e) {
  try {
    if (e && e.user) return normalizeEmail_(e.user.getEmail());
  } catch (err) {}

  try {
    return normalizeEmail_(Session.getActiveUser().getEmail());
  } catch (err) {
    return '';
  }
}

function isAdminEmail_(email) {
  const normalized = normalizeEmail_(email);
  return normalized !== '' && getAdminEmails_().indexOf(normalized) !== -1;
}

function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function getDocumentProperty_(key) {
  return PropertiesService.getDocumentProperties().getProperty(key);
}

function setDocumentProperty_(key, value) {
  PropertiesService.getDocumentProperties().setProperty(key, String(value));
}

function isRegistrationFullLocked_() {
  return getDocumentProperty_(LAB.PROP.FULL_LOCK) === 'YES';
}

function setRegistrationFullLocked_(value) {
  setDocumentProperty_(LAB.PROP.FULL_LOCK, value ? 'YES' : 'NO');
}

function getWorkflowPhase_() {
  return getDocumentProperty_(LAB.PROP.WORKFLOW_PHASE) || LAB.PHASE_CLOSED;
}

function setWorkflowPhase_(phase) {
  const allowed = [
    LAB.PHASE_ROSTER_OPEN,
    LAB.PHASE_REGISTRATION_OPEN,
    LAB.PHASE_CLOSED
  ];
  if (allowed.indexOf(phase) === -1) throw new Error('Giai đoạn hệ thống không hợp lệ: ' + phase);
  setDocumentProperty_(LAB.PROP.WORKFLOW_PHASE, phase);
}

function workflowPhaseText_(phase) {
  if (phase === LAB.PHASE_REGISTRATION_OPEN) return 'Đang mở đăng ký ca';
  if (phase === LAB.PHASE_CLOSED) return 'Đã khóa đăng ký';
  return 'Chờ mở đăng ký tuần mới';
}

