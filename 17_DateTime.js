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

