// ==========================================
// 【本番用】03_Helpers.gs v25 (変更なし・再掲)
// ==========================================

function getTrueData(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow === 0 || lastCol === 0) return [];
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  let trueLast = 1;
  for (let i = lastRow - 1; i >= 0; i--) {
    if (data[i].join("").length > 0) {
      trueLast = i + 1;
      break;
    }
  }
  return data.slice(0, trueLast);
}

function extractPlannedShifts(cellText) {
  const shifts = [];
  SHIFT_REGEX.lastIndex = 0; 
  let match;
  while ((match = SHIFT_REGEX.exec(cellText)) !== null) {
    shifts.push({
      loc: match[1],
      start: `${match[2].padStart(2, '0')}:${match[3] ? match[3].padStart(2, '0') : '00'}`,
      endHour: parseInt(match[4], 10)
    });
  }
  return shifts;
}

function calculatePlannedHours(plannedShifts) {
  let total = 0;
  plannedShifts.forEach(shift => {
    const startHour = parseInt(shift.start.split(":")[0], 10);
    let hours = shift.endHour - startHour;
    if (hours >= 8) hours -= 1; 
    total += hours > 0 ? hours : 0;
  });
  return total;
}

function normalizeLocation(rawLoc, masterMap, dept = "") {
  let base = String(rawLoc).replace(/[\s　\/／・]+/g, "");

  if (masterMap[base]) {
    base = masterMap[base];
  } else if (/^\d+$/.test(base)) {
    let num = parseInt(base, 10);
    let numStr = num.toString();
    let padStr = numStr.padStart(2, '0');
    let modStr = (num % 1000).toString();
    let padModStr = modStr.padStart(2, '0');

    if (masterMap[numStr]) base = masterMap[numStr];
    else if (masterMap[padStr]) base = masterMap[padStr];
    else if (masterMap[modStr]) base = masterMap[modStr];
    else if (masterMap[padModStr]) base = masterMap[padModStr];
  }
  
  if (base === "亀有" || base === "北葛西") {
    if (dept.includes("内科") || rawLoc.includes("内科")) {
      return base + "内科";
    }
  }
  return base;
}

function timeToMins(tStr) {
  if (!tStr) return -1;
  const match = tStr.match(/(\d{1,2}):(\d{2})/);
  if (match) return parseInt(match[1], 10) * 60 + parseInt(match[2], 10);
  return -1;
}

function calcHoursFromTimes(startStr, endStr) {
  let sM = timeToMins(startStr);
  let eM = timeToMins(endStr);
  if (sM === -1 || eM === -1) return 0;
  if (eM < sM) eM += 24*60;
  let dur = eM - sM;
  let bS = 13*60, bE = 15*60;
  let overS = Math.max(sM, bS);
  let overE = Math.min(eM, bE);
  if(overE > overS) dur -= (overE - overS);
  return Math.max(0, dur / 60);
}

function parseDateToString(val) {
  if (!val) return null;
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}/${m}/${d}`;
  }
  const str = String(val);
  const match = str.match(DATE_REGEX);
  if (match) return `${match[1]}/${match[2].padStart(2, '0')}/${match[3].padStart(2, '0')}`;
  return null;
}

function extractTimeOnly(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return `${String(val.getHours()).padStart(2, '0')}:${String(val.getMinutes()).padStart(2, '0')}`;
  }
  const str = String(val);
  const match = str.match(/(\d{1,2}):(\d{2})/);
  if (match) return `${match[1].padStart(2, '0')}:${match[2]}`;
  return str;
}