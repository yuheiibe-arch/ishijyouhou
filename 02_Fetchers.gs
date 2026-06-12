// ==========================================
// 【本番用】02_Fetchers.gs v26 (完全ID照合版)
// ==========================================

function fetchLocationMaster() {
  const map = {};
  try {
    const ss = SpreadsheetApp.openById(MASTER_SS_ID);
    const data = getTrueData(ss.getSheetByName("拠点名"));
    if (data.length === 0) return map;
    
    const headers = data[0].map(h => String(h).replace(/[\s　]+/g, ""));
    const noIdx = headers.findIndex(h => h.includes("クリニックNo"));
    
    for (let i = 1; i < data.length; i++) {
      const formal = String(data[i][0]).replace(/[\s　]+/g, "");
      if (!formal) continue;
      
      for (let c = 1; c <= 4; c++) {
        if (data[i][c]) {
          const variation = String(data[i][c]).replace(/[\s　]+/g, "");
          map[variation] = formal;
        }
      }
      
      if (noIdx !== -1 && data[i][noIdx]) {
        let clinicNoStr = String(data[i][noIdx]).replace(/[\s　]+/g, "");
        map[clinicNoStr] = formal; 
        
        let num = parseInt(clinicNoStr, 10);
        if (!isNaN(num)) {
          map[num.toString()] = formal; 
        }
      }
      
      map[formal] = formal; 
    }
  } catch(e) {}
  return map;
}

function findColumnIndex(headers, priorities) {
  for (let i = 0; i < priorities.length; i++) {
    const idx = headers.findIndex(h => priorities[i].test(String(h).replace(/[\s　]/g, "")));
    if (idx !== -1) return idx;
  }
  return -1;
}

function fetchActualShifts(locMaster) {
  const map = new Map();
  try {
    const ss = SpreadsheetApp.openById(SHIFT_SS_ID);
    const data = getTrueData(ss.getSheetByName("確定シフト"));
    const headers = data[0] || [];
    
    // ★氏名ではなく医籍番号を探す
    const medIdIdx = headers.findIndex(h => /医籍番号/.test(String(h)));
    const clinicIdx = headers.findIndex(h => /クリニック|拠点|勤務先/.test(String(h)));
    const deptIdx = headers.findIndex(h => /診療科/.test(String(h)));
    const dateIdx = headers.findIndex(h => /勤務日|日付/.test(String(h)));
    const remarksIdx = headers.findIndex(h => /備考|コメント/.test(String(h)));

    const normalizedHeaders = headers.map(h => String(h).replace(/[\s　()（）]/g, ""));
    const startIdx = normalizedHeaders.indexOf("元のシフト勤務開始時間");
    const endIdx = normalizedHeaders.indexOf("元のシフト勤務終了時間");

    const allowanceIndices = [];
    for (let i = 1; i <= 5; i++) {
      let idx = normalizedHeaders.indexOf(`追加支給額${i}`);
      if (idx > -1) allowanceIndices.push(idx);
    }

    if (dateIdx === -1 || medIdIdx === -1 || clinicIdx === -1) {
       console.error("確定シフトに「医籍番号」または必須列が見つかりません。");
       return map;
    }

    for (let i = 1; i < data.length; i++) {
      let clinicName = String(data[i][clinicIdx]);
      
      let cleanCode = clinicName.replace(/[\s　]+/g, "");
      if (SPECIAL_CODE_MAP[cleanCode]) {
        clinicName = SPECIAL_CODE_MAP[cleanCode];
      }

      if (IGNORE_SHIFTS.some(ig => clinicName.includes(ig))) continue;

      const dStr = parseDateToString(data[i][dateIdx]);
      // ★名前ではなく医籍番号を取得
      const medId = String(data[i][medIdIdx]).replace(/[\s　]+/g, "");

      if (dStr && medId) {
        // ★キーが「日付_医籍番号」になる
        const key = `${dStr}_${medId}`;
        if (!map.has(key)) map.set(key, []);
        
        let hasAllowance = false;
        for (let aIdx of allowanceIndices) {
          let val = data[i][aIdx];
          if (val && !isNaN(parseFloat(val)) && parseFloat(val) > 0) {
            hasAllowance = true;
            break;
          }
        }

        let isGuaranteedPay = clinicName.includes("院外勤務");

        map.get(key).push({
          clinic: clinicName, 
          dept: deptIdx > -1 ? String(data[i][deptIdx]) : "",
          start: startIdx > -1 ? extractTimeOnly(data[i][startIdx]) : "",
          end: endIdx > -1 ? extractTimeOnly(data[i][endIdx]) : "",
          remarks: remarksIdx > -1 ? String(data[i][remarksIdx]) : "",
          hasAllowance: hasAllowance,
          isGuaranteedPay: isGuaranteedPay
        });
      }
    }
  } catch (e) { console.error("fetchActualShifts エラー: " + e.message); }
  return map;
}

function fetchAbsenceData(ss) {
  const map = new Map();
  try {
    const data = getTrueData(ss.getSheetByName("お休み情報"));
    const headers = data[0] || [];
    const dateIdx = headers.findIndex(h => /日付|勤務日/.test(String(h)));
    // ★氏名ではなく医籍番号を探す
    const medIdIdx = headers.findIndex(h => /医籍番号/.test(String(h)));
    const typeIdx = headers.findIndex(h => /休み種別|種別/.test(String(h)));

    if (dateIdx > -1 && medIdIdx > -1) {
      for (let i = 1; i < data.length; i++) {
        const dStr = parseDateToString(data[i][dateIdx]);
        // ★医籍番号を取得
        const medId = String(data[i][medIdIdx]).replace(/[\s　]+/g, "");
        if (dStr && medId) map.set(`${dStr}_${medId}`, data[i][typeIdx]);
      }
    } else {
       console.warn("お休み情報シートに「医籍番号」列が見つかりません。");
    }
  } catch(e) {}
  return map;
}

function fetchSubstitutionData(ss) {
  const map = new Map();
  try {
    const data = getTrueData(ss.getSheetByName("振替勤務"));
    const headers = data[0] || [];
    const dateIdx = headers.findIndex(h => /日付|勤務日/.test(String(h)));
    // ★氏名ではなく医籍番号を探す
    const medIdIdx = headers.findIndex(h => /医籍番号/.test(String(h)));
    const startIdx = headers.findIndex(h => /開始時間|開始/.test(String(h)));
    const endIdx = headers.findIndex(h => /終了時間|終了/.test(String(h)));
    const locIdx = headers.findIndex(h => /拠点|クリニック|勤務先|店舗/.test(String(h)));

    if (dateIdx > -1 && medIdIdx > -1) {
      for (let i = 1; i < data.length; i++) {
        const dStr = parseDateToString(data[i][dateIdx]);
        // ★医籍番号を取得
        const medId = String(data[i][medIdIdx]).replace(/[\s　]+/g, "");
        
        if (dStr && medId) {
          let shiftText = "";
          let hours = 0;
          let loc = (locIdx > -1 && data[i][locIdx]) ? String(data[i][locIdx]).replace(/[\s　]+/g, "") : "拠点不明";

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
      console.warn("振替勤務シートに「医籍番号」列が見つかりません。");
    }
  } catch(e) {}
  return map;
}

function fetchExternalClosedDays() {
  const map = new Map();
  try {
    const extSs = SpreadsheetApp.openById(CLOSED_SS_ID);
    const data = getTrueData(extSs.getSheetByName("休館日"));
    const headers = data[0] || [];
    const dateIdx = headers.findIndex(h => /日付/.test(String(h)));
    const locIdx = headers.findIndex(h => /拠点名|拠点/.test(String(h)));

    if (dateIdx > -1 && locIdx > -1) {
      for (let i = 1; i < data.length; i++) {
        const dStr = parseDateToString(data[i][dateIdx]);
        if (dStr && data[i][locIdx]) {
          if (!map.has(dStr)) map.set(dStr, []);
          map.get(dStr).push(String(data[i][locIdx]).replace(/[\s　]+/g, ""));
        }
      }
    }
  } catch (e) {}
  return map;
}