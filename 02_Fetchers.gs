// ==========================================
// 【本番用】02_Fetchers.gs (完全動的検索・アーカイブ自動補完・ID正規化対応版)
// ==========================================

const ARCHIVE_SHIFT_SS_ID = "10Z8jg4o7Ri9Ggf7u_jauyDDZa6bVnc-iuuuqmkNDafw";

/**
 * 医籍番号の正規化（空白除去・全角半角統一・小数点除去・先頭ゼロ統一）
 */
function normalizeMedId(val) {
  if (val === null || val === undefined) return "";
  let s = String(val)
    .replace(/[\s ]+/g, "")
    .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .replace(/\.0+$/, "");
  if (/^\d+$/.test(s)) s = String(parseInt(s, 10));
  return s;
}

/**
 * 拠点マスタの取得（列番号ハードコードなし・動的検索）
 */
function fetchLocationMaster() {
  const map = {};
  try {
    const ss = SpreadsheetApp.openById(MASTER_SS_ID);
    const data = getTrueData(ss.getSheetByName("拠点名"));
    if (data.length === 0) return map;

    const headers = data[0].map(h => String(h).replace(/[\s ]+/g, ""));
    let formalIdx = headers.findIndex(h => /正式|拠点名|名称/.test(h));
    if (formalIdx === -1) formalIdx = 0;
    const noIdx = headers.findIndex(h => /クリニックNo|番号|コード/i.test(h));
    const areaIdx = headers.findIndex(h => /エリア|地域|都道府県/.test(h));

    for (let i = 1; i < data.length; i++) {
      const formal = String(data[i][formalIdx] || "").replace(/[\s ]+/g, "");
      if (!formal) continue;
      map[formal] = formal;

      for (let c = 0; c < headers.length; c++) {
        if (c === formalIdx || c === areaIdx || !data[i][c]) continue;
        const val = String(data[i][c]).replace(/[\s ]+/g, "");
        if (!val) continue;

        if (c === noIdx) {
          map[val] = formal;
          const num = parseInt(val, 10);
          if (!isNaN(num)) map[num.toString()] = formal;
        } else {
          map[val] = formal;
        }
      }
    }
  } catch (e) {
    console.error("fetchLocationMaster エラー: " + e.message);
  }
  return map;
}

function findColumnIndex(headers, priorities) {
  for (let i = 0; i < priorities.length; i++) {
    const idx = headers.findIndex(h => priorities[i].test(String(h).replace(/[\s ]/g, "")));
    if (idx !== -1) return idx;
  }
  return -1;
}

/**
 * 確定シフトシート（現行 または アーカイブ月別シート）1枚分を動的に解析してMapに格納する共通関数
 */
function parseShiftSheetDynamic(sheet, targetMap) {
  let minDate = "9999/99/99";
  let maxDate = "0000/00/00";
  let rowCount = 0;

  if (!sheet) return { rowCount, minDate, maxDate };
  const data = getTrueData(sheet);
  if (data.length < 2) return { rowCount, minDate, maxDate };

  // 見出し行を1〜5行目から動的に特定
  let headerRowIdx = 0;
  for (let r = 0; r < Math.min(5, data.length); r++) {
    const joined = data[r].join("");
    if (/勤務日|日付/.test(joined) && /医籍番号|クリニック|拠点/.test(joined)) {
      headerRowIdx = r;
      break;
    }
  }

  const headers = data[headerRowIdx].map(h => String(h).trim());
  const normHeaders = headers.map(h => h.replace(/[\s ()（）]/g, ""));

  const medIdIdx = headers.findIndex(h => /医籍番号/.test(h));
  const clinicNameIdx = headers.findIndex(h => /クリニック名|拠点名/.test(h));
  const clinicNoIdx = headers.findIndex(h => /クリニックNo|拠点No/i.test(h));
  const fallbackClinicIdx = headers.findIndex(h => /クリニック|拠点|勤務先|店舗/.test(h));
  const deptIdx = headers.findIndex(h => /診療科|科目/.test(h));
  const dateIdx = headers.findIndex(h => /勤務日|日付|対象日/.test(h));
  const remarksIdx = headers.findIndex(h => /備考|コメント/.test(h));

  let startIdx = normHeaders.indexOf("元のシフト勤務開始時間");
  if (startIdx === -1) startIdx = normHeaders.findIndex(h => /勤務開始時間|開始時間/.test(h));
  let endIdx = normHeaders.indexOf("元のシフト勤務終了時間");
  if (endIdx === -1) endIdx = normHeaders.findIndex(h => /勤務終了時間|終了時間/.test(h));

  const allowanceIndices = [];
  normHeaders.forEach((h, idx) => {
    if (/^追加支給額\d+$/.test(h)) allowanceIndices.push(idx);
  });

  if (dateIdx === -1 || medIdIdx === -1 || (clinicNameIdx === -1 && clinicNoIdx === -1 && fallbackClinicIdx === -1)) {
    console.error(`シート「${sheet.getName()}」に必須列（日付・医籍番号・拠点）が見つかりません。`);
    return { rowCount, minDate, maxDate };
  }

  for (let i = headerRowIdx + 1; i < data.length; i++) {
    const row = data[i];
    // クリニック名を最優先、空ならクリニックNoを使用
    const rawName = clinicNameIdx > -1 ? String(row[clinicNameIdx] || "").trim() : "";
    const rawNo = clinicNoIdx > -1 ? String(row[clinicNoIdx] || "").trim() : "";
    let clinicName = rawName || rawNo || (fallbackClinicIdx > -1 ? String(row[fallbackClinicIdx] || "").trim() : "");

    const cleanCode = (rawNo || clinicName).replace(/[\s ]+/g, "");
    if (SPECIAL_CODE_MAP[cleanCode]) {
      clinicName = SPECIAL_CODE_MAP[cleanCode];
    }

    if (IGNORE_SHIFTS.some(ig => clinicName.includes(ig))) continue;

    const dStr = parseDateToString(row[dateIdx]);
    const medId = normalizeMedId(row[medIdIdx]);

    if (!dStr || !medId) continue;

    if (dStr < minDate) minDate = dStr;
    if (dStr > maxDate) maxDate = dStr;
    rowCount++;

    const key = `${dStr}_${medId}`;
    if (!targetMap.has(key)) targetMap.set(key, []);

    let hasAllowance = false;
    for (const aIdx of allowanceIndices) {
      const val = row[aIdx];
      if (val && !isNaN(parseFloat(val)) && parseFloat(val) > 0) {
        hasAllowance = true;
        break;
      }
    }

    const isGuaranteedPay = clinicName.includes("院外勤務");

    targetMap.get(key).push({
      clinic: clinicName,
      dept: deptIdx > -1 ? String(row[deptIdx] || "") : "",
      start: startIdx > -1 ? extractTimeOnly(row[startIdx]) : "",
      end: endIdx > -1 ? extractTimeOnly(row[endIdx]) : "",
      remarks: remarksIdx > -1 ? String(row[remarksIdx] || "") : "",
      hasAllowance: hasAllowance,
      isGuaranteedPay: isGuaranteedPay
    });
  }

  return { rowCount, minDate, maxDate };
}

/**
 * 現行の「確定シフト」を取得（01C main3.gs 等の既存呼び出しとも完全互換）
 * ※戻り値の Map に .minDate プロパティを付与して収録開始日を伝える
 */
function fetchActualShifts(locMaster, neededArchiveMonths) {
  const map = new Map();
  let minCurrentDate = "9999/99/99";

  try {
    const ss = SpreadsheetApp.openById(SHIFT_SS_ID);
    const sheet = ss.getSheetByName("確定シフト");
    const res = parseShiftSheetDynamic(sheet, map);
    minCurrentDate = res.minDate;
    console.log(`[確定シフト読込] 現行シート: ${res.rowCount}行 (期間: ${res.minDate} ～ ${res.maxDate})`);
  } catch (e) {
    console.error("fetchActualShifts(現行) エラー: " + e.message);
  }

  // 必要な過去アーカイブ月（"2026/04" など）が指定されている場合のみ、アーカイブスプシから補完読込
  if (neededArchiveMonths && neededArchiveMonths.size > 0 && ARCHIVE_SHIFT_SS_ID) {
    try {
      const archSs = SpreadsheetApp.openById(ARCHIVE_SHIFT_SS_ID);
      neededArchiveMonths.forEach(ym => {
        const archSheet = archSs.getSheetByName(ym);
        if (archSheet) {
          const res = parseShiftSheetDynamic(archSheet, map);
          console.log(`[アーカイブ補完読込] シート「${ym}」: ${res.rowCount}行`);
        }
      });
    } catch (e) {
      console.error("fetchActualShifts(アーカイブ) エラー: " + e.message);
    }
  }

  map.minDate = minCurrentDate;
  return map;
}

function fetchAbsenceData(ss) {
  const map = new Map();
  try {
    const data = getTrueData(ss.getSheetByName("お休み情報"));
    if (data.length < 2) return map;
    const headers = data[0].map(h => String(h).trim());
    const dateIdx = headers.findIndex(h => /日付|勤務日|対象日/.test(h));
    const medIdIdx = headers.findIndex(h => /医籍番号/.test(h));
    const typeIdx = headers.findIndex(h => /休み種別|種別/.test(h));

    if (dateIdx > -1 && medIdIdx > -1 && typeIdx > -1) {
      for (let i = 1; i < data.length; i++) {
        const dStr = parseDateToString(data[i][dateIdx]);
        const medId = normalizeMedId(data[i][medIdIdx]);
        if (dStr && medId) map.set(`${dStr}_${medId}`, data[i][typeIdx]);
      }
    } else {
      console.warn("お休み情報シートに必須列（日付・医籍番号・種別）が見つかりません。");
    }
  } catch (e) {}
  return map;
}

function fetchSubstitutionData(ss) {
  const map = new Map();
  try {
    const data = getTrueData(ss.getSheetByName("振替勤務"));
    if (data.length < 2) return map;
    const headers = data[0].map(h => String(h).trim());
    const dateIdx = headers.findIndex(h => /日付|勤務日|対象日/.test(h));
    const medIdIdx = headers.findIndex(h => /医籍番号/.test(h));
    const startIdx = headers.findIndex(h => /開始時間|開始/.test(h));
    const endIdx = headers.findIndex(h => /終了時間|終了/.test(h));
    const locIdx = headers.findIndex(h => /拠点|クリニック|勤務先|店舗/.test(h));

    if (dateIdx > -1 && medIdIdx > -1) {
      for (let i = 1; i < data.length; i++) {
        const dStr = parseDateToString(data[i][dateIdx]);
        const medId = normalizeMedId(data[i][medIdIdx]);

        if (dStr && medId) {
          let shiftText = "";
          let hours = 0;
          const loc = (locIdx > -1 && data[i][locIdx]) ? String(data[i][locIdx]).replace(/[\s ]+/g, "") : "拠点不明";

          if (startIdx > -1 && endIdx > -1 && data[i][startIdx] && data[i][endIdx]) {
            const sTime = extractTimeOnly(data[i][startIdx]);
            const eTime = extractTimeOnly(data[i][endIdx]);
            const cleanSTime = sTime.replace(/:00$/, "");
            const cleanETime = eTime.replace(/:00$/, "");

            shiftText = `【${loc}】${cleanSTime}-${cleanETime}※振替`;
            hours = calcHoursFromTimes(sTime, eTime);
          } else {
            shiftText = `【${loc}】※振替`;
          }
          map.set(`${dStr}_${medId}`, { text: shiftText, hours: hours });
        }
      }
    } else {
      console.warn("振替勤務シートに必須列（日付・医籍番号）が見つかりません。");
    }
  } catch (e) {}
  return map;
}

function fetchExternalClosedDays() {
  const map = new Map();
  try {
    const extSs = SpreadsheetApp.openById(CLOSED_SS_ID);
    const data = getTrueData(extSs.getSheetByName("休館日"));
    if (data.length < 2) return map;
    const headers = data[0].map(h => String(h).trim());
    const dateIdx = headers.findIndex(h => /日付/.test(h));
    const locIdx = headers.findIndex(h => /拠点名|拠点/.test(h));

    if (dateIdx > -1 && locIdx > -1) {
      for (let i = 1; i < data.length; i++) {
        const dStr = parseDateToString(data[i][dateIdx]);
        if (dStr && data[i][locIdx]) {
          if (!map.has(dStr)) map.set(dStr, []);
          map.get(dStr).push(String(data[i][locIdx]).replace(/[\s ]+/g, ""));
        }
      }
    }
  } catch (e) {}
  return map;
}