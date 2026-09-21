/**********************************************************************
 * PHÂN BỔ SỐ LƯỢNG ĐỘNG VÀ THỨ TỰ SINH VIÊN
 **********************************************************************/

function getScheduleConfig_(studentCount, weekMode) {
  const mode = weekMode || getWeekMode_();
  const days = mode === LAB.MODE_GENERAL_CLEANING
    ? ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5']
    : ['Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6'];

  const extraOrder = mode === LAB.MODE_GENERAL_CLEANING
    ? [0, 1, 2, 3]
    : [4, 3, 2, 1, 0];

  const dailyCounts = distributeTotal_(studentCount, days.length, extraOrder);

  return days.map(function (day, index) {
    const labs = allocateLabs_(dailyCounts[index]);
    return {
      day: day,
      dayOffset: index,
      total: dailyCounts[index],
      lab404: labs.lab404,
      lab403: labs.lab403,
      phongPhu: labs.phongPhu,
      note: dailyCounts[index] + ' người: Lab 4.04 = ' + labs.lab404 +
        '; Lab 4.03 = ' + labs.lab403 + '; Phòng phụ = ' + labs.phongPhu
    };
  });
}

function distributeTotal_(total, bucketCount, extraOrder) {
  total = Math.max(0, Math.floor(Number(total) || 0));
  bucketCount = Math.max(1, Math.floor(Number(bucketCount) || 1));
  const base = Math.floor(total / bucketCount);
  const result = new Array(bucketCount).fill(base);
  const remainder = total % bucketCount;
  const order = (extraOrder || []).filter(function (index) {
    return index >= 0 && index < bucketCount;
  });

  for (let i = 0; i < remainder; i++) {
    result[order[i] !== undefined ? order[i] : i]++;
  }
  return result;
}

function allocateLabs_(dailyTotal) {
  const total = Math.max(0, Math.floor(Number(dailyTotal) || 0));
  if (total === 0) return { lab404: 0, lab403: 0, phongPhu: 0 };
  if (total === 1) return { lab404: 1, lab403: 0, phongPhu: 0 };
  if (total === 2) return { lab404: 1, lab403: 1, phongPhu: 0 };

  const phongPhu = 1;
  const remaining = total - phongPhu;
  return {
    lab404: Math.ceil(remaining / 2),
    lab403: Math.floor(remaining / 2),
    phongPhu: phongPhu
  };
}

function orderStudentsForWeek_(students, monday, salt) {
  if (!students.length) return [];
  const ordered = students.slice().sort(function (a, b) {
    return a.label.localeCompare(b.label, 'vi', { sensitivity: 'base' });
  });
  const shift = parseInt(hashString_(weekKey_(monday) + '|' + (salt || 'REGULAR')).slice(-8), 16) % ordered.length;
  return ordered.slice(shift).concat(ordered.slice(0, shift));
}

function hashString_(text) {
  let hash = 2166136261;
  const value = String(text || '');
  for (let i = 0; i < value.length; i++) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return ('00000000' + (hash >>> 0).toString(16)).slice(-8);
}

function getTotalByField_(schedule, fieldName) {
  return schedule.reduce(function (sum, day) {
    return sum + Number(day[fieldName] || 0);
  }, 0);
}

function getScheduleTotal_(schedule) {
  return schedule.reduce(function (sum, day) {
    return sum + Number(day.total || 0);
  }, 0);
}

