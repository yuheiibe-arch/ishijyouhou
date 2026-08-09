// ==========================================
// 【本番用・完全統一版】実績データ（KING OF TIME等）と勤怠予定の同期
// （✅ 医籍番号判定 ＋ 院外ブロック ＋ 辞書プルダウン ＋ 日付型トラップ回避 ＋ ARRAYFORMULA対応 ＋ 行追加 ＋ 強行書き込み ＋ 降順ソート ＋ 過去日非表示）
// ==========================================

function syncActualsToExceptions() {
  const API_URL = "https://script.google.com/a/macros/mnys.jp/s/AKfycbzzMGqyKvNvc5daaRAxx0mKE5Ipn5mE2Ghuh6i0j-vOOws1xgFF1HBFC57od09B2-o5/exec";
  
  const today = new Date();
  const currentMonth = today.getMonth() + 1;
  const currentFY = currentMonth >= 4 ? today.getFullYear() : today.getFullYear() - 1;
  const activeYears = [String(currentFY), String(currentFY + 1)];
  
  const HARDCODED_START_DATE = "2026/04/01";

  const activeSs = SpreadsheetApp.getActiveSpreadsheet();
  let targetSs = activeSs;
  if (typeof ATTENDANCE_SS_ID !== "undefined" && ATTENDANCE_SS_ID !== "") {
    try { targetSs = SpreadsheetApp.openById(ATTENDANCE_SS_ID); } catch (e) {}
  }

  console.log(`=== 🚀 欠勤・例外の抽出連動（対象年度: ${activeYears.join(", ")}）を開始 ===`);

  const locMaster = typeof fetchLocationMaster === 'function' ? fetchLocationMaster() : {};
  const actualShifts = typeof fetchActualShifts === 'function' ? fetchActualShifts(locMaster) : new Map();
  
  let exceptionSs = activeSs; 
  if (typeof EXCEPTION_SS_ID !== "undefined" && EXCEPTION_SS_ID !== "") {
    try { exceptionSs = SpreadsheetApp.openById(EXCEPTION_SS_ID); } catch (e) {}
  } else if (targetSs.getId() !== activeSs.getId()) {
    exceptionSs = targetSs; 
  }
  
  const absenceSheet = exceptionSs.getSheetByName("お休み情報");
  const subSheet = exceptionSs.getSheetByName("振替勤務");

  if (!absenceSheet || !subSheet) {
    console.error("❌ お休み情報 または 振替勤務 シートが見つかりません。");
    return;
  }

  // ★時間フォーマットを元の正しい仕様に復元（半角スペース埋め）
  const padTime = (t) => {
    if (!t) return "";
    let str = String(t).trim();
    if (str.match(/^0\d:\d{2}$/)) str = str.substring(1); 
    return str.padStart(5, ' ');
  };

  const cleanStr = (str) => String(str).replace(/[\s 【】\u200B\n\r]/g, "").trim(); 

  const exceptionHeaders = absenceSheet.getRange(1, 1, 1, absenceSheet.getLastColumn()).getValues()[0];
  const cLocIndex = exceptionHeaders.map(cleanStr).findIndex(h => /拠点|クリニック|勤務先|店舗/.test(h));
  let allowedValues = [];
  if (cLocIndex > -1) {
    const rule = absenceSheet.getRange(2, cLocIndex + 1).getDataValidation();
    if (rule) {
      if (rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) allowedValues = rule.getCriteriaValues()[0];
      else if (rule.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) allowedValues = rule.getCriteriaValues()[0].getValues().flat().filter(String);
    }
  }

  let dict = {};
  try {
    const mapUrl = "https://docs.google.com/spreadsheets/d/14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs/edit";
    const ssMap = SpreadsheetApp.openByUrl(mapUrl);
    const sheetMap = ssMap.getSheetByName("拠点名");
    const dataMap = sheetMap.getDataRange().getValues();
    for (let i = 1; i < dataMap.length; i++) {
      let canonical = dataMap[i][0] ? String(dataMap[i][0]).trim() : "";
      if (!canonical) continue;
      dict[canonical] = canonical;
      for (let j = 1; j <= 4; j++) {
        let variant = dataMap[i][j] ? String(dataMap[i][j]).replace(/[\s ]+/g, "") : "";
        if (variant) dict[variant] = canonical;
      }
    }
  } catch(e) { console.error("❌ マスタ辞書構築エラー: " + e.message); }

  const getDropdownFormalName = (rawCalendarLoc, deptName) => {
    let cleanRaw = String(rawCalendarLoc).replace(/[【】\(（]?(内科|小児科)[\)）]?/g, "").replace(/\/.*/, "").replace(/[\s ]+/g, "");
    
    // ★修正：クリニックNoが含まれる辞書(locMaster)も参照させるように変更
    let canonicalLoc = dict[cleanRaw] || locMaster[cleanRaw] || cleanRaw;

    if (canonicalLoc === "亀有" || canonicalLoc === "北葛西") {
      let suffix = (deptName === "内科") ? "（内科）" : "（小児科）";
      if (allowedValues.includes(canonicalLoc + suffix)) return canonicalLoc + suffix;
    }

    for (let i = 0; i < allowedValues.length; i++) {
      let dropVal = allowedValues[i];
      if (!dropVal) continue;
      let cleanDrop = String(dropVal).replace(/[【】\(（]?(内科|小児科)[\)）]?/g, "").replace(/\/.*/, "").replace(/[\s ]+/g, "");
      let dropCanonical = dict[cleanDrop] || cleanDrop;
      if (dropCanonical === canonicalLoc) return dropVal; 
    }
    return canonicalLoc;
  };

  const buildExistingSet = (sheet) => {
    let set = new Set();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return set;

    const headers = data[0].map(h => cleanStr(h));
    const cDate = headers.findIndex(h => /日付|勤務日|対象日|年月日|日時/.test(h));
    const cMedId = headers.findIndex(h => /医籍番号|ID|id/i.test(h));
    const cName = headers.findIndex(h => /氏名|名前|医師名|ドクター/.test(h));

    if (cDate === -1) return set; 

    for (let i = 1; i < data.length; i++) {
      let dVal = data[i][cDate];
      let dStr = "";
      
      if (dVal instanceof Date) {
        dStr = Utilities.formatDate(dVal, "JST", "yyyy/MM/dd");
      } else if (dVal) {
        let strVal = String(dVal).replace(/[年月]/g, '/').replace(/日/g, '');
        let d = new Date(strVal);
        if (!isNaN(d.getTime())) dStr = Utilities.formatDate(d, "JST", "yyyy/MM/dd");
      }
      
      if (!dStr) continue;
      
      let medId = cMedId > -1 ? cleanStr(data[i][cMedId]) : "";
      if (!medId) medId = "NO_ID";

      let docName = cName > -1 ? cleanStr(data[i][cName]).replace(/先生$/, "") : "NO_NAME";
      
      set.add(`${dStr}_${medId}`);
      set.add(`${dStr}_NAME:${docName}`);
    }
    return set;
  };

  const existingAbsenceSet = buildExistingSet(absenceSheet);
  const existingSubSet = buildExistingSet(subSheet);

  const targetSheets = targetSs.getSheets().filter(sheet => {
    const match = sheet.getName().match(/^(常勤|定期非常勤)勤怠(\d{4})/);
    return match && activeYears.includes(match[2]);
  });

  let newAbsenceData = [];
  let newSubData = [];
  let affectedLocationsByYear = {};
  activeYears.forEach(year => affectedLocationsByYear[year] = new Set());

  targetSheets.forEach(sheet => {
    const sheetYear = sheet.getName().match(/\d{4}/)[0];
    const fullValues = typeof getTrueData === 'function' ? getTrueData(sheet) : sheet.getDataRange().getValues();
    if (fullValues.length < 3) return; 

    let originalSummaryStartRow = fullValues.length;
    for (let i = 2; i < fullValues.length; i++) {
      let cellA = fullValues[i][0];
      if (cellA == null || String(cellA).replace(/[\s \n\r]/g, "") === "") {
        originalSummaryStartRow = i;
        break;
      }
      let strA = String(cellA);
      if (strA.includes("合計") || strA.includes("年間") || strA.includes("備考")) {
        originalSummaryStartRow = i;
        break;
      }
    }

    let nameRowIdx = 0, idRowIdx = 1, textCountRow0 = 0, textCountRow1 = 0;
    let dateIdx = fullValues[0].findIndex(h => String(h).includes("日付"));
    if (dateIdx === -1 && fullValues.length > 1) dateIdx = fullValues[1].findIndex(h => String(h).includes("日付"));
    if (dateIdx === -1) dateIdx = 0;

    for(let c = dateIdx + 1; c < fullValues[0].length; c++) {
      if (fullValues[0][c] && isNaN(Number(fullValues[0][c]))) textCountRow0++;
      if (fullValues[1][c] && isNaN(Number(fullValues[1][c]))) textCountRow1++;
    }
    if (textCountRow1 > textCountRow0) { nameRowIdx = 1; idRowIdx = 0; }

    const topHeaders = fullValues[nameRowIdx].map(h => cleanStr(h));
    const idHeaders = fullValues[idRowIdx].map(h => cleanStr(h));
    
    const calValues = fullValues.slice(Math.max(1, nameRowIdx, idRowIdx) + 1, originalSummaryStartRow);

    for (let r = 0; r < calValues.length; r++) {
      const dateCell = calValues[r][dateIdx];
      if (!(dateCell instanceof Date)) continue;
      
      const dateKey = Utilities.formatDate(dateCell, "JST", "yyyy/MM/dd");
      if (dateKey < HARDCODED_START_DATE) continue; 
      
      for (let c = dateIdx + 1; c < topHeaders.length; c++) {
        const docNameClean = topHeaders[c];
        const docMedId = idHeaders[c]; 
        if (!docNameClean || !docMedId || docNameClean === "休" || isNaN(Number(docMedId))) continue;

        let rawVal = String(calValues[r][c] || "");
        if (!rawVal || rawVal === "休") continue;

        if (rawVal.includes('契約：')) {
          const match = rawVal.match(/契約：([^\n]+)/);
          if (match) rawVal = match[1]; 
        }

        let isManualAbsence = rawVal.includes("欠勤");
        let cleanLines = rawVal.split('\n').filter(line => !line.startsWith('→') && !line.includes('※振替') && !line.includes('半日有給') && !line.includes('有給') && !line.includes('欠勤') && !line.includes('移動依頼') && !line.includes('確定：'));
        let currentVal = cleanLines.join('\n').trim();
        if (!currentVal) continue;

        const plannedShifts = typeof extractPlannedShifts === 'function' ? extractPlannedShifts(currentVal) : [];
        const compositeKey = `${dateKey}_${docMedId}`; 
        const actualsForDoctorDay = actualShifts.get(compositeKey) || [];
        
        const hasAbsenceShiftInActual = isManualAbsence || actualsForDoctorDay.some(act => act.clinic.includes("欠勤") || act.dept.includes("欠勤") || act.remarks.includes("欠勤"));
        
        const hasAllowance = actualsForDoctorDay.some(act => {
          const isAllowance = act.hasAllowance || act.isGuaranteedPay;
          const isIngai = String(act.clinic).includes("院外") || String(act.remarks).includes("院外");
          return isAllowance && !isIngai;
        });

        plannedShifts.forEach(plan => {
          if (!plan.loc) return; 

          let deptName = "小児科"; 
          if (actualsForDoctorDay.length > 0 && actualsForDoctorDay[0].dept) {
            deptName = actualsForDoctorDay[0].dept.includes("内科") ? "内科" : "小児科";
          } else if (plan.loc.includes("内科")) {
            deptName = "内科";
          }

          const safeEnd = plan.end ? plan.end : (plan.endHour ? `${plan.endHour}:00` : "");
          let locForDropdown = getDropdownFormalName(plan.loc, deptName);

          let planLocNorm = typeof normalizeLocation === 'function' ? normalizeLocation(plan.loc, locMaster) : plan.loc;
          planLocNorm = planLocNorm.replace(/ヶ/g, "が"); 
          const cleanLocStr = cleanStr(planLocNorm).split('_')[0]; 
          
          const docNameCleanStr = cleanStr(docNameClean).replace(/先生$/, "");
          
          const exactMatchKey = `${dateKey}_${docMedId}`;
          const fallbackMatchKey = `${dateKey}_NAME:${docNameCleanStr}`;
          
          if (existingAbsenceSet.has(exactMatchKey) || existingAbsenceSet.has(fallbackMatchKey) || 
              existingSubSet.has(exactMatchKey) || existingSubSet.has(fallbackMatchKey)) return; 

          let matchFound = false;
          if (actualsForDoctorDay.length > 0) {
            matchFound = actualsForDoctorDay.some(act => {
              if (act.clinic.includes("欠勤") || act.dept.includes("欠勤") || act.remarks.includes("欠勤")) return false;
              const actLocNorm = typeof normalizeLocation === 'function' ? normalizeLocation(act.clinic, locMaster, act.dept) : act.clinic;
              return actLocNorm.includes(planLocNorm) || planLocNorm.includes(actLocNorm);
            });
          }

          if (isManualAbsence) matchFound = false;

          if (!matchFound) {
            if (hasAbsenceShiftInActual) {
              newAbsenceData.push({ date: dateKey, medId: docMedId, start: padTime(plan.start), end: padTime(safeEnd), loc: locForDropdown, type: "欠勤", name: docNameClean });
              existingAbsenceSet.add(exactMatchKey);
              existingAbsenceSet.add(fallbackMatchKey); 
              affectedLocationsByYear[sheetYear].add(cleanLocStr); 
            } else if (hasAllowance && actualsForDoctorDay.length > 0) {
              const actDest = actualsForDoctorDay.find(act => (act.hasAllowance || act.isGuaranteedPay) && !String(act.clinic).includes("院外") && !String(act.remarks).includes("院外"));
              
              if (actDest) {
                let rawClinicName = String(actDest.clinic).trim();
                if (rawClinicName === "8") rawClinicName = "東品川";

                let actDestDept = actDest.dept ? actDest.dept : (rawClinicName.includes("内科") ? "内科" : "小児科");
                let destForDropdown = getDropdownFormalName(rawClinicName, actDestDept);

                newSubData.push({ date: dateKey, medId: docMedId, start: padTime(actDest.start), end: padTime(actDest.end), loc: destForDropdown, type: "事務局移動", name: docNameClean });
                existingSubSet.add(`${dateKey}_${docMedId}`);
                existingSubSet.add(`${dateKey}_NAME:${docNameCleanStr}`);
                affectedLocationsByYear[sheetYear].add(cleanStr(destForDropdown).split('_')[0]); 
              }
            }
          }
        });
      }
    }
  });

  function getRealLastRow(sheet) {
    const aVals = sheet.getRange("A1:A").getValues();
    for (let i = aVals.length - 1; i >= 0; i--) {
      if (aVals[i][0] !== "" && aVals[i][0] != null) return i + 1;
    }
    return 1;
  }

  function appendDataWithFormat(sheet, dataArray) {
    if (dataArray.length === 0) return 0;
    
    let lastCol = sheet.getLastColumn();
    let realLastRow = Math.max(2, getRealLastRow(sheet)); 
    let startRow = realLastRow + 1;
    let numRows = dataArray.length;

    // ★シートの行数が足りない場合は末尾に追加してクラッシュを防ぐ
    let maxRows = sheet.getMaxRows();
    if (startRow + numRows - 1 > maxRows) {
      sheet.insertRowsAfter(maxRows, (startRow + numRows - 1) - maxRows + 2);
    }

    const sheetHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => cleanStr(h));
    
    const cDate = sheetHeaders.findIndex(h => /日付|勤務日|対象日|年月日|日時/.test(h));
    const cName = sheetHeaders.findIndex(h => /氏名|名前|医師名|ドクター/.test(h));
    const cMedId = sheetHeaders.findIndex(h => /医籍番号|ID|id/i.test(h)); 
    const cStart = sheetHeaders.findIndex(h => /開始|出勤|入/.test(h));
    const cEnd = sheetHeaders.findIndex(h => /終了|退勤|退/.test(h));
    const cLoc = sheetHeaders.findIndex(h => /拠点|クリニック|勤務先|店舗|勤務地|場所/.test(h));
    const cType = sheetHeaders.findIndex(h => /種別|タイプ|理由|事由|区分|休み/.test(h));

    const sourceRange = sheet.getRange(realLastRow, 1, 1, lastCol);
    const targetRange = sheet.getRange(startRow, 1, numRows, lastCol);
    
    // ★フォーマットと入力規則のコピー
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    
    // ★数式のコピーとARRAYFORMULAの保護
    sourceRange.copyTo(targetRange); 
    const formulas = targetRange.getFormulas();
    const values = targetRange.getValues();
    for (let r = 0; r < numRows; r++) {
      for (let c = 0; c < lastCol; c++) {
        if (formulas[r][c] === "") {
          if (values[r][c] === true || values[r][c] === false) values[r][c] = false;
          else values[r][c] = null; // 空文字ではなくnullを入れてARRAYFORMULAを貫通させる
        }
      }
    }
    targetRange.setValues(values); 

    let writeData = { date: [], name: [], start: [], end: [], loc: [], type: [] };

    for (let i = 0; i < numRows; i++) {
      let rowObj = dataArray[i];
      writeData.date.push([rowObj.date]);
      writeData.name.push([rowObj.name]);
      // ★時間にアポストロフィを復元し、プルダウン規則の型と完全一致させる
      writeData.start.push([rowObj.start ? `'${rowObj.start}` : null]);
      writeData.end.push([rowObj.end ? `'${rowObj.end}` : null]);
      writeData.loc.push([rowObj.loc || null]); 
      writeData.type.push([rowObj.type || null]);
    }
    
    const safeSetValues = (colIdx, dataArr) => {
      if (colIdx > -1) {
        const range = sheet.getRange(startRow, colIdx + 1, numRows, 1);
        const rules = range.getDataValidations();
        range.clearDataValidations(); // ドロップダウンを一時解除
        range.setValues(dataArr);     // データ書き込み
        try { range.setDataValidations(rules); } catch(e) {} // 復元
      }
    };

    safeSetValues(cDate, writeData.date);
    safeSetValues(cName, writeData.name);
    // 医籍番号は書き込まない
    safeSetValues(cStart, writeData.start);
    safeSetValues(cEnd, writeData.end);
    safeSetValues(cLoc, writeData.loc);
    safeSetValues(cType, writeData.type);

    return numRows;
  }

  // ★ソートと非表示を同時に行う統合関数
  const sortAndHidePastDates = (sheet) => {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) return;
    
    sheet.showRows(2, lastRow - 1);

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(cleanStr);
    const dateCol = headers.findIndex(h => /日付|勤務日|対象日|年月日|日時/.test(h)) + 1;
    
    if (dateCol > 0) {
      sheet.getRange(2, 1, lastRow - 1, lastCol).sort({ column: dateCol, ascending: false });
      
      const dateValues = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues();
      const todayDate = new Date();
      todayDate.setHours(0, 0, 0, 0);
      
      for (let i = 0; i < dateValues.length; i++) {
        const dVal = dateValues[i][0];
        let rowDate = null;
        
        if (dVal instanceof Date) {
          rowDate = new Date(dVal.getTime());
        } else if (dVal) {
          let strVal = String(dVal).replace(/[年月]/g, '/').replace(/日/g, '');
          let d = new Date(strVal);
          if (!isNaN(d.getTime())) {
            rowDate = d;
          }
        }
        
        if (rowDate) {
          rowDate.setHours(0, 0, 0, 0);
          if (rowDate.getTime() < todayDate.getTime()) {
            sheet.hideRows(i + 2);
          }
        }
      }
    }
  };

  let isUpdated = false;

  if (newAbsenceData.length > 0) {
    let addedCount = appendDataWithFormat(absenceSheet, newAbsenceData);
    console.log(`\n✅ 【お休み情報】に ${addedCount} 件追記しました。`);
    isUpdated = true;
  }

  if (newSubData.length > 0) {
    let addedCount = appendDataWithFormat(subSheet, newSubData);
    console.log(`\n✅ 【振替勤務】に ${addedCount} 件追記しました。`);
    isUpdated = true;
  }

  if (isUpdated) {
    sortAndHidePastDates(absenceSheet);
    sortAndHidePastDates(subSheet);
    console.log("✅ シートを降順にソートし、過去の日付を非表示にしました。");
  }

  let apiSentCount = 0;
  activeYears.forEach(year => {
    if (affectedLocationsByYear[year].size > 0) {
      const targetsStr = Array.from(affectedLocationsByYear[year]).join(",");
      console.log(`🌐 [${year}年度] 変更を検知！ カレンダーのAPIを呼び出します... [対象拠点: ${targetsStr}]`);
      const url = `${API_URL}?year=${year}&targets=${encodeURIComponent(targetsStr)}`;
      
      try {
        const response = UrlFetchApp.fetch(url, { method: "get", muteHttpExceptions: true, headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() } });
        console.log(`🚀 APIレスポンス(${year}): ${response.getContentText()}`);
        apiSentCount++;
      } catch (e) {
        console.error(`❌ API呼び出しエラー(${year}): ${e.message}`);
      }
    }
  });

  if (apiSentCount === 0 && !isUpdated) {
    console.log("\n✨ 新たに追記すべき差分はありませんでした。処理をスキップしました。");
  }
}