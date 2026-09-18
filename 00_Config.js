/**
 * HỆ THỐNG SINH VIÊN TỰ KHAI DANH SÁCH VÀ TỰ ĐĂNG KÝ TRỰC
 * Phiên bản 4.4.0 - không đối chiếu email sinh viên, vẫn chống trùng và ghi đè ca
 *
 * NÂNG CẤP FILE ĐANG DÙNG (GIỮ NGUYÊN CÁC BẠN ĐÃ ĐĂNG KÝ):
 * 1) Dán TOÀN BỘ tệp này thay cho mã cũ rồi lưu.
 * 2) Chạy capNhatDangKyKhongEmail đúng 1 lần, sau đó tải lại Google Sheet.
 * 3) KHÔNG chạy setupLabCleaningSystem khi tuần hiện tại đã có đăng ký.
 *
 * CÀI MỚI: chạy setupLabCleaningSystem đúng 1 lần, cấp quyền rồi tải lại Sheet.
 *
 * Quy trình: SV khai danh sách một lần cho cả học kỳ; Chủ nhật hệ thống
 * chuẩn bị lịch khoảng 15:30 và mở cho SV tự chọn ca khoảng 16:00.
 * Hệ thống không đối chiếu tài khoản Google của sinh viên khi đăng ký.
 * Vì vậy sinh viên phải tự giác chọn đúng tên của mình.
 *
 * Tài khoản chạy setup lần đầu (thường là chủ file) được lưu là quản trị viên.
 * Nếu có thêm cán bộ cần được phép tick hoàn thành, điền email vào ADMIN_EMAILS.
 */

const LAB = {
  VERSION: '4.4.0',
  TIMEZONE: 'Asia/Ho_Chi_Minh',
  PREPARE_HOUR: 15,
  PREPARE_MINUTE: 30,
  RESET_HOUR: 16,
  RECOVERY_HOUR: 17,
  HEALTH_CHECK_HOUR: 6,
  DEFAULT_ASSIGNMENT_MODE: 'SELF',
  ADMIN_EMAILS: ['ndtung133@gmail.com'],
  MIN_ROSTER_INPUT_ROWS: 100,
  ROSTER_BUFFER_ROWS: 20,
  ROSTER_GROWTH_ROWS: 50,
  MAX_LOG_ROWS: 2000,

  SHEETS: {
    HOME: 'TRANG CHỦ',
    CONFIG: 'CẤU HÌNH',
    STUDENTS: 'DANH SÁCH SV',
    REGISTER: 'ĐĂNG KÝ TRỰC',
    DASHBOARD: 'THEO DÕI',
    RULES: 'QUY ĐỊNH',
    GENERAL_CLEANING: 'TỔNG VỆ SINH',
    GENERAL_HISTORY: 'LỊCH SỬ TỔNG VS',
    WEEK_HISTORY: 'LỊCH SỬ TUẦN',
    SYSTEM_LOG: 'NHẬT KÝ HỆ THỐNG',
    SYSTEM: '_HỆ THỐNG'
  },

  FIRST_STUDENT_ROW: 3,
  FIRST_REGISTER_ROW: 6,
  FIRST_GENERAL_ROW: 7,

  COL: {
    STT: 1,
    THU: 2,
    NGAY: 3,
    LAB: 4,
    VIEC: 5,
    VITRI: 6,
    SV: 7,
    TIME_DANGKY: 8,
    GHICHU_SV: 9,
    TRANGTHAI: 10,
    DA_TRUC: 11,
    NGUOI_KIEMTRA: 12,
    TIME_KIEMTRA: 13,
    GHICHU_KIEMTRA: 14
  },

  GENERAL_COL: {
    STT: 1,
    SV: 2,
    AREA: 3,
    JOB: 4,
    REQUIREMENT: 5,
    TIME: 6,
    DONE: 7,
    INSPECTOR: 8,
    CHECK_TIME: 9,
    NOTE: 10
  },

  MODE_REGULAR: 'REGULAR',
  MODE_GENERAL_CLEANING: 'GENERAL_CLEANING',
  ASSIGN_AUTO: 'AUTO',
  ASSIGN_SELF: 'SELF',
  PHASE_ROSTER_OPEN: 'ROSTER_OPEN',
  PHASE_REGISTRATION_OPEN: 'REGISTRATION_OPEN',
  PHASE_CLOSED: 'CLOSED',

  PROP: {
    SPREADSHEET_ID: 'LAB_SPREADSHEET_ID',
    WEEK_MODE: 'LAB_WEEK_MODE',
    ASSIGNMENT_MODE: 'LAB_ASSIGNMENT_MODE',
    WORKFLOW_PHASE: 'LAB_WORKFLOW_PHASE',
    FULL_LOCK: 'LAB_REG_FULL_LOCKED',
    ADMIN_EMAILS: 'LAB_ADMIN_EMAILS',
    ROSTER_HASH: 'LAB_ROSTER_HASH',
    PENDING_ROSTER_HASH: 'LAB_PENDING_ROSTER_HASH',
    LAST_RESET_WEEK: 'LAB_LAST_RESET_WEEK',
    STATUS_DIRTY: 'LAB_STATUS_DIRTY',
    ROSTER_DIRTY: 'LAB_ROSTER_DIRTY'
  },

  TRIGGER: {
    EDIT: 'handleLabCleaningEdit',
    PREPARE: 'weeklyPrepareLabCleaning',
    OPEN: 'weeklyOpenLabCleaning',
    RECOVERY: 'sundayRecoveryLabCleaning',
    HEALTH: 'healthCheckLabCleaning',
    SYNC: 'refreshLabCleaningViews'
  },

  PROTECTION_PREFIX: 'LAB_SYSTEM_PROTECTION::',
  ROSTER_ROW_PREFIX: 'LAB_ROSTER_OWNER::',
  REGISTER_ROW_PREFIX: 'LAB_REGISTER_OWNER::',
  GENERAL_ROW_PREFIX: 'LAB_GENERAL_OWNER::',
  LEGACY_LOCK_DESC: 'AUTO_LOCK_DANG_KY_TRUC_VE_SINH_LAB'
};
