/**********************************************************************
 * VIỆC LÀM VÀ VỊ TRÍ TRỰC THƯỜNG - HỖ TRỢ MỌI SỐ NGƯỜI/PHÒNG
 **********************************************************************/

function getPositionNameByLab_(labName, position, count) {
  if (labName === 'Phòng phụ') {
    if (count === 1) return 'Vị trí 1 - Lau + quét/lau phòng phụ';
    if (position === count) return 'Vị trí ' + position + ' - Quét/lau sàn phòng phụ';
    return 'Vị trí ' + position + ' - Lau khu vực phòng phụ ' + position;
  }

  if (count === 1) return 'Vị trí 1 - Lau + quét/lau phòng';
  if (position === count) return 'Vị trí ' + position + ' - Quét/lau sàn phòng';
  return 'Vị trí ' + position + ' - Lau ' + cleaningZoneName_(position, count - 1);
}

function getTaskByLab_(labName, position, count) {
  if (labName === 'Phòng phụ') {
    if (count === 1) return 'Phòng phụ - Lau + quét/lau toàn bộ phòng.';
    if (position === count) return 'Phòng phụ - Quét/lau sàn phòng.';
    return 'Phòng phụ - Lau ' + cleaningZoneName_(position, count - 1) + '.';
  }

  if (count === 1) return labName + ' - Lau + quét/lau toàn bộ phòng.';
  if (position === count) return labName + ' - Quét/lau sàn phòng.';
  return labName + ' - Lau ' + cleaningZoneName_(position, count - 1) + '.';
}

function cleaningZoneName_(position, cleaningCount) {
  if (cleaningCount === 1) return 'bàn và các bề mặt trong phòng';
  if (cleaningCount === 2) return position === 1 ? 'khu vực bên trái' : 'khu vực bên phải';
  if (cleaningCount === 3) {
    return ['khu vực bên trái', 'khu vực ở giữa', 'khu vực bên phải'][position - 1];
  }
  return 'khu vực số ' + position + '/' + cleaningCount;
}

function applyAlternatingDayColors_(sh, rowCount) {
  if (rowCount <= 0) return;
  const days = sh.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.THU, rowCount, 1).getValues();
  const backgrounds = [];
  let previous = null;
  let colorIndex = 0;

  days.forEach(function (row) {
    if (row[0] !== previous) {
      previous = row[0];
      colorIndex = 1 - colorIndex;
    }
    backgrounds.push(Array(14).fill(colorIndex ? '#f3f6f4' : '#ffffff'));
  });

  sh.getRange(LAB.FIRST_REGISTER_ROW, 1, rowCount, 14).setBackgrounds(backgrounds);
  sh.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.SV, rowCount, 1).setBackground('#fff2cc');
  sh.getRange(LAB.FIRST_REGISTER_ROW, LAB.COL.DA_TRUC, rowCount, 1).setBackground('#d9ead3');
}

