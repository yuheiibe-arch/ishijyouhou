/**
 * 【ステップ5：最終シミュレーション（ドライラン）】
 * ※シートへの書き込み・削除は一切行いません（完全読み取り専用）
 * 新ロジックで4月〜9月が正しく照合されるか、4月からの勤怠遵守率がどう計算されるかを検証します。
 */
function debug_Step5_DryRunSimulation() {
  console.log("==========================================");
  console.log("🚀 新ロジック ドライラン検証開始（シート変更なし）");
  console.log("==========================================");

  const ARCHIVE_SS_ID = "10Z8jg4o7Ri9Ggf7u_jauyDDZa6bVnc-iuuuqmkNDafw";
  const locMaster = fetchLocationMaster_Dynamic_();

  // 1. 現行の確定シフト ＋ アーカイブ(2026/04〜2026/09) を動的カラム検索で統合取得
  const actualShifts = new Map();
  
  // (A) まず現行の「確定シフト」から取得
  const currentSs = SpreadsheetApp.openById(SHIFT_SS_ID);
  const currentSheet = currentSs.getSheetByName("確定シフト");
  let minCurrentDate = "9999/99/99";
  if (currentSheet) {
    const res = parseShiftSheet_Dynamic_(currentSheet, actualShifts);
    if (res.minDate < minCurrentDate) minCurrentDate = res.minDate;
    console.log(`✅ 現行「確定シフト」読込完了: ${res.rowCount}行 (収録開始日: ${minCurrentDate})`);
  }

  // (B) 現行データより前の月（例: 2026/04〜2026/09）をアーカイブスプシから動的取得
  const archiveSs = SpreadsheetApp.openById(ARCHIVE_SS_ID);
  archiveSs.getSheets().forEach(s => {
    const sName = s.getName().trim();
    // "2026/04" 形式のシートで、かつ現行データの開始月より前のものを読み込む
    if (/^\d{4}\/\d{2}$/.test(sName) && (sName + "/01") < minCurrentDate) {
      const res = parseShiftSheet_Dynamic_(s, actualShifts);
      console.log(`✅ アーカイブ「${sName}」読込完了: ${res.rowCount}行`);
    }
  });
  console.log(`📦 統合後の確定シフト総キー数: ${actualShifts.size}件\n`);

  // 2. 例外データ（お休み情報・振替勤務・休館日）の動的取得
  const exceptionSs = SpreadsheetApp.openById(EXCEPTION_SS_ID);
  const absenceData = fetchAbsenceData_Dynamic_(exceptionSs);
  const subData = fetchSubstitutionData_Dynamic_(exceptionSs);

  // 3. 勤怠シートでのシミュレーション（常勤・定期非常勤）
  const activeSs = SpreadsheetApp.getActiveSpreadsheet();
  ["常勤勤怠2026", "定期非常勤勤怠2026"].forEach(sheetName => {
    const sheet = activeSs.getSheetByName(sheetName);
    if (!sheet) return;

    const fullValues = getTrueData(sheet);
    const headerInfo = detectAttendanceHeaders_Dynamic_(fullValues);
    const { docRowIdx, idRowIdx, calStartRowIdx, dateColIdx, firstDocColIdx, calendarEndRowIdx } = headerInfo;

    const cleanHeaders = fullValues[docRowIdx].map(h => String(h).replace(/[\s ]+/g, ""));
    const idHeaders = fullValues[idRowIdx].map(h => normalizeId_Dynamic_(h));

    console.log(`========================================`);
    console.log(`📊 シミュレーション結果: 【${sheetName}】`);
    console.log(`  ・動的検出 -> 医師名行:${docRowIdx+1}行目, 医籍番号行:${idRowIdx+1}行目, 日付列:${dateColIdx+1}列目, 医師開始列:${firstDocColIdx+1}列目, カレンダー範囲:${calStartRowIdx+1}〜${calendarEndRowIdx}行目`);

    let statsAprSep = { matched: 0, mismatch: 0, missing: 0, exception: 0 };
    let sampleRestored = [];
    const summaryDiffs = {};

    for (let r = calStartRowIdx; r < calendarEndRowIdx; r++) {
      const dateCell = fullValues[r][dateColIdx];
      if (!(dateCell instanceof Date)) continue;
      const dateKey = Utilities.formatDate(dateCell, "JST", "yyyy/MM/dd");
      const monthKey = `${dateCell.getFullYear()}年${dateCell.getMonth() + 1}月`;
      const isAprToSep = (dateKey >= "2026/04/01" && dateKey <= "2026/09/30");

      for (let c = firstDocColIdx; c < fullValues[0].length; c++) {
        const docNameClean = cleanHeaders[c];
        const docMedId = idHeaders[c];
        if (!docNameClean || !docMedId) continue;

        if (!summaryDiffs[docNameClean]) {
          summaryDiffs[docNameClean] = { annual: { absH: 0, absD: 0, subH: 0, subD: 0, paid: 0 }, months: {} };
        }
        if (!summaryDiffs[docNameClean].months[monthKey]) {
          summaryDiffs[docNameClean].months[monthKey] = { absH: 0, absD: 0, subH: 0, subD: 0 };
        }

        const compositeKey = `${dateKey}_${docMedId}`;
        let rawCellVal = String(fullValues[r][c] || "");
        let currentVal = rawCellVal;

        if (currentVal.includes("契約：")) {
          const match = currentVal.match(/契約：([\s\S]*?)(?:\n確定：|$)/);
          if (match) currentVal = match[1].trim();
        }
        let cleanLines = currentVal.split('\n').filter(line => !line.startsWith('→') && !line.includes('※振替') && !line.includes('半日有給') && !line.includes('有給') && !line.includes('欠勤') && !line.includes('移動依頼'));
        currentVal = cleanLines.join('\n');

        const isOriginallyOff = (!currentVal || currentVal === "休" || currentVal === "-" || currentVal === "");
        const plannedShifts = isOriginallyOff ? [] : extractPlannedShifts(currentVal);
        if (plannedShifts.length === 0 && !subData.has(compositeKey)) continue;

        const actualsForDoctorDay = actualShifts.get(compositeKey) || [];
        const hasPaidInAct = actualsForDoctorDay.some(a => /有給|有休/.test(a.clinic + a.remarks));
        const hasAbsInAct = actualsForDoctorDay.some(a => /欠勤/.test(a.clinic + a.dept + a.remarks));

        let hasException = false;
        let finalCellPreview = currentVal;

        if (absenceData.has(compositeKey)) {
          const absType = String(absenceData.get(compositeKey));
          hasException = true;
          finalCellPreview += `\n→${absType}`;
          if (absType.includes("欠勤")) {
            const h = calculatePlannedHours(plannedShifts);
            summaryDiffs[docNameClean].months[monthKey].absH += h;
            summaryDiffs[docNameClean].months[monthKey].absD += 1;
            summaryDiffs[docNameClean].annual.absH += h;
            summaryDiffs[docNameClean].annual.absD += 1;
          } else if (/有給|有休/.test(absType)) {
            summaryDiffs[docNameClean].annual.paid += (/半/.test(absType) ? 0.5 : 1.0);
          }
        } else if (hasPaidInAct) {
          hasException = true;
          summaryDiffs[docNameClean].annual.paid += 1.0;
          finalCellPreview += `\n→有給`;
        } else if (hasAbsInAct) {
          hasException = true;
          finalCellPreview += `\n→欠勤`;
          const h = calculatePlannedHours(plannedShifts);
          summaryDiffs[docNameClean].months[monthKey].absH += h;
          summaryDiffs[docNameClean].months[monthKey].absD += 1;
          summaryDiffs[docNameClean].annual.absH += h;
          summaryDiffs[docNameClean].annual.absD += 1;
        }

        if (subData.has(compositeKey)) {
          const sub = subData.get(compositeKey);
          hasException = true;
          finalCellPreview = sub.text;
          summaryDiffs[docNameClean].months[monthKey].subH += sub.hours;
          summaryDiffs[docNameClean].months[monthKey].subD += 1;
          summaryDiffs[docNameClean].annual.subH += sub.hours;
          summaryDiffs[docNameClean].annual.subD += 1;
        }

        if (hasException) {
          if (isAprToSep) statsAprSep.exception++;
        } else if (plannedShifts.length > 0) {
          let isMismatch = false, isMissing = false;
          plannedShifts.forEach(plan => {
            const planLocNorm = normalizeLocation(plan.loc, locMaster);
            const matchFound = actualsForDoctorDay.some(act => {
              const actLocNorm = normalizeLocation(act.clinic, locMaster, act.dept);
              const locMatch = actLocNorm.includes(planLocNorm) || planLocNorm.includes(actLocNorm);
              const sM_act = timeToMins(act.start), eM_act = timeToMins(act.end), sM_plan = timeToMins(plan.start);
              const timeMatch = (act.start === plan.start) || (sM_act <= sM_plan && sM_plan < eM_act);
              return locMatch && timeMatch;
            });
            if (!matchFound && !(actualsForDoctorDay.some(a => a.hasAllowance || a.isGuaranteedPay))) {
              if (actualsForDoctorDay.length > 0) isMismatch = true;
              else isMissing = true;
            }
          });

          if (isMismatch) {
            if (isAprToSep) statsAprSep.mismatch++;
            const actTexts = actualsForDoctorDay.map(a => `【${normalizeLocation(a.clinic, locMaster, a.dept)}】${a.start.replace(/:00$/,"")}-${a.end.replace(/:00$/,"")}`).join("、");
            finalCellPreview = `契約：${currentVal}\n確定：${actTexts}`;
          } else if (isMissing) {
            if (isAprToSep) statsAprSep.missing++;
            finalCellPreview = `契約：${currentVal}\n確定：未登録`;
          } else {
            if (isAprToSep) statsAprSep.matched++;
          }
        }

        if (isAprToSep && sampleRestored.length < 5 && rawCellVal.includes("未登録") && !finalCellPreview.includes("未登録")) {
          sampleRestored.push(`  ・${dateKey} ${docNameClean}: [旧] ${JSON.stringify(rawCellVal)}  ==>  [新] ${JSON.stringify(finalCellPreview)}`);
        }
      }
    }

    console.log(`📈 4月〜9月 予定シフト照合結果:`);
    console.log(`  ・⭕ 契約通り一致(正常): ${statsAprSep.matched}件`);
    console.log(`  ・🔄 有給/欠勤/振替反映: ${statsAprSep.exception}件`);
    console.log(`  ・🟥 契約と確定の不一致(赤): ${statsAprSep.mismatch}件`);
    console.log(`  ・🟨 確定シフト未登録(黄): ${statsAprSep.missing}件`);
    console.log(`✨ 「未登録」から正常復元されるセルのサンプル(先頭5件):\n` + sampleRestored.join("\n"));

    // 4月〜3月を通算した勤怠遵守率のシミュレーション（変化があった医師トップ5を表示）
    const origSummaryVals = fullValues.slice(calendarEndRowIdx);
    const origDaysRow = origSummaryVals.find(row => String(row[dateColIdx]).includes("年間勤務日数"));
    let rateSamples = [];
    if (origDaysRow) {
      for (let c = firstDocColIdx; c < Math.min(firstDocColIdx + 15, fullValues[0].length); c++) {
        const docName = cleanHeaders[c];
        if (!docName) continue;
        const origDays = parseFloat(String(origDaysRow[c]).replace(/[^\d.]/g, "")) || 0;
        const d = summaryDiffs[docName] ? summaryDiffs[docName].annual : { absD: 0, subD: 0, paid: 0 };
        const netAbsD = Math.max(0, d.absD - d.subD);
        const rate = origDays > 0 ? Math.min(100, Math.max(0, Math.round(((origDays - netAbsD) / origDays) * 100))) : 0;
        rateSamples.push(`${docName}: 予定${origDays}日 / 欠勤${d.absD}日(振替${d.subD}日) / 有給${d.paid}日 -> 遵守率 ${rate}％`);
      }
    }
    console.log(`🏆 通年(4月〜3月) 集計＆勤怠遵守率シミュレーション(先頭15名):\n  ` + rateSamples.join("\n  "));
  });
}

// --- 以下、動的検索ヘルパー群（ハードコード完全ゼロ） ---
function normalizeId_Dynamic_(val) {
  if (val === null || val === undefined) return "";
  let s = String(val).replace(/[\s ]+/g, "").replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)).replace(/\.0+$/, "");
  if (/^\d+$/.test(s)) s = String(parseInt(s, 10));
  return s;
}

function fetchLocationMaster_Dynamic_() {
  const map = {};
  const ss = SpreadsheetApp.openById(MASTER_SS_ID);
  const data = getTrueData(ss.getSheetByName("拠点名"));
  if (data.length === 0) return map;
  const headers = data[0].map(h => String(h).replace(/[\s ]+/g, ""));
  const formalIdx = headers.findIndex(h => /正式|拠点名|名称/.test(h)) !== -1 ? headers.findIndex(h => /正式|拠点名|名称/.test(h)) : 0;
  const noIdx = headers.findIndex(h => /クリニックNo|番号|コード/i.test(h));
  const areaIdx = headers.findIndex(h => /エリア|地域|都道府県/.test(h));

  for (let i = 1; i < data.length; i++) {
    const formal = String(data[i][formalIdx] || "").replace(/[\s ]+/g, "");
    if (!formal) continue;
    map[formal] = formal;
    for (let c = 0; c < headers.length; c++) {
      if (c === formalIdx || c === areaIdx || !data[i][c]) continue;
      const val = String(data[i][c]).replace(/[\s ]+/g, "");
      if (c === noIdx) {
        map[val] = formal;
        const num = parseInt(val, 10);
        if (!isNaN(num)) map[num.toString()] = formal;
      } else {
        map[val] = formal;
      }
    }
  }
  return map;
}

function parseShiftSheet_Dynamic_(sheet, targetMap) {
  const data = getTrueData(sheet);
  if (data.length < 2) return { rowCount: 0, minDate: "9999/99/99" };
  let hIdx = 0;
  for (let r = 0; r < Math.min(5, data.length); r++) {
    if (/勤務日|日付/.test(data[r].join("")) && /医籍番号|クリニック/.test(data[r].join(""))) { hIdx = r; break; }
  }
  const headers = data[hIdx].map(h => String(h).trim());
  const normHeaders = headers.map(h => h.replace(/[\s ()（）]/g, ""));

  const medIdIdx = headers.findIndex(h => /医籍番号/.test(h));
  const clinicNameIdx = headers.findIndex(h => /クリニック名|拠点名/.test(h));
  const clinicNoIdx = headers.findIndex(h => /クリニックNo|拠点No/i.test(h));
  const fallbackClinicIdx = headers.findIndex(h => /クリニック|拠点|勤務先/.test(h));
  const deptIdx = headers.findIndex(h => /診療科|科目/.test(h));
  const dateIdx = headers.findIndex(h => /勤務日|日付|対象日/.test(h));
  const remarksIdx = headers.findIndex(h => /備考|コメント/.test(h));

  let startIdx = normHeaders.indexOf("元のシフト勤務開始時間");
  if (startIdx === -1) startIdx = normHeaders.findIndex(h => /勤務開始時間|開始時間/.test(h));
  let endIdx = normHeaders.indexOf("元のシフト勤務終了時間");
  if (endIdx === -1) endIdx = normHeaders.findIndex(h => /勤務終了時間|終了時間/.test(h));

  const allowanceIndices = [];
  normHeaders.forEach((h, idx) => { if (/^追加支給額\d+$/.test(h)) allowanceIndices.push(idx); });

  let minDate = "9999/99/99", count = 0;
  for (let i = hIdx + 1; i < data.length; i++) {
    const row = data[i];
    // ★クリニック名を最優先、空ならクリニックNoを使用
    let rawName = clinicNameIdx > -1 ? String(row[clinicNameIdx] || "").trim() : "";
    let rawNo = clinicNoIdx > -1 ? String(row[clinicNoIdx] || "").trim() : "";
    let clinicName = rawName || rawNo || (fallbackClinicIdx > -1 ? String(row[fallbackClinicIdx] || "").trim() : "");
    let cleanCode = (rawNo || clinicName).replace(/[\s ]+/g, "");
    if (SPECIAL_CODE_MAP[cleanCode]) clinicName = SPECIAL_CODE_MAP[cleanCode];
    if (IGNORE_SHIFTS.some(ig => clinicName.includes(ig))) continue;

    const dStr = parseDateToString(row[dateIdx]);
    const medId = medIdIdx > -1 ? normalizeId_Dynamic_(row[medIdIdx]) : "";
    if (!dStr || !medId) continue;

    if (dStr < minDate) minDate = dStr;
    count++;
    const key = `${dStr}_${medId}`;
    if (!targetMap.has(key)) targetMap.set(key, []);

    const hasAllowance = allowanceIndices.some(aIdx => row[aIdx] && !isNaN(parseFloat(row[aIdx])) && parseFloat(row[aIdx]) > 0);
    targetMap.get(key).push({
      clinic: clinicName,
      dept: deptIdx > -1 ? String(row[deptIdx] || "") : "",
      start: startIdx > -1 ? extractTimeOnly(row[startIdx]) : "",
      end: endIdx > -1 ? extractTimeOnly(row[endIdx]) : "",
      remarks: remarksIdx > -1 ? String(row[remarksIdx] || "") : "",
      hasAllowance: hasAllowance,
      isGuaranteedPay: clinicName.includes("院外勤務")
    });
  }
  return { rowCount: count, minDate: minDate };
}

function fetchAbsenceData_Dynamic_(ss) {
  const map = new Map();
  const data = getTrueData(ss.getSheetByName("お休み情報"));
  if (data.length < 2) return map;
  const headers = data[0].map(h => String(h).trim());
  const dateIdx = headers.findIndex(h => /日付|勤務日/.test(h));
  const medIdIdx = headers.findIndex(h => /医籍番号/.test(h));
  const typeIdx = headers.findIndex(h => /休み種別|種別/.test(h));
  if (dateIdx > -1 && medIdIdx > -1 && typeIdx > -1) {
    for (let i = 1; i < data.length; i++) {
      const dStr = parseDateToString(data[i][dateIdx]);
      const medId = normalizeId_Dynamic_(data[i][medIdIdx]);
      if (dStr && medId) map.set(`${dStr}_${medId}`, data[i][typeIdx]);
    }
  }
  return map;
}

function fetchSubstitutionData_Dynamic_(ss) {
  const map = new Map();
  const data = getTrueData(ss.getSheetByName("振替勤務"));
  if (data.length < 2) return map;
  const headers = data[0].map(h => String(h).trim());
  const dateIdx = headers.findIndex(h => /日付|勤務日/.test(h));
  const medIdIdx = headers.findIndex(h => /医籍番号/.test(h));
  const startIdx = headers.findIndex(h => /開始時間|開始/.test(h));
  const endIdx = headers.findIndex(h => /終了時間|終了/.test(h));
  const locIdx = headers.findIndex(h => /拠点|クリニック|勤務先|店舗/.test(h));
  if (dateIdx > -1 && medIdIdx > -1) {
    for (let i = 1; i < data.length; i++) {
      const dStr = parseDateToString(data[i][dateIdx]);
      const medId = normalizeId_Dynamic_(data[i][medIdIdx]);
      if (dStr && medId) {
        const loc = (locIdx > -1 && data[i][locIdx]) ? String(data[i][locIdx]).replace(/[\s ]+/g, "") : "拠点不明";
        let shiftText = `【${loc}】※振替`, hours = 0;
        if (startIdx > -1 && endIdx > -1 && data[i][startIdx] && data[i][endIdx]) {
          const sTime = extractTimeOnly(data[i][startIdx]), eTime = extractTimeOnly(data[i][endIdx]);
          shiftText = `【${loc}】${sTime.replace(/:00$/, "")}-${eTime.replace(/:00$/, "")}※振替`;
          hours = calcHoursFromTimes(sTime, eTime);
        }
        map.set(`${dStr}_${medId}`, { text: shiftText, hours: hours });
      }
    }
  }
  return map;
}

function detectAttendanceHeaders_Dynamic_(fullValues) {
  let dateColIdx = 0, calHeaderRowIdx = 1;
  for (let r = 0; r < Math.min(5, fullValues.length); r++) {
    const idx = fullValues[r].findIndex(cell => String(cell).trim() === "日付");
    if (idx !== -1) { calHeaderRowIdx = r; dateColIdx = idx; break; }
  }
  const docRowIdx = calHeaderRowIdx > 0 ? calHeaderRowIdx - 1 : 0;
  const idRowIdx = calHeaderRowIdx;
  const calStartRowIdx = Math.max(docRowIdx, idRowIdx) + 1;

  // カレンダー見出し（日付・曜日・平日判定・祝日判定・祝日備考・年末年始判定）以外の最初の列＝医師開始列を動的特定
  let firstDocColIdx = dateColIdx + 1;
  for (let c = dateColIdx + 1; c < fullValues[idRowIdx].length; c++) {
    const hVal = String(fullValues[idRowIdx][c]).trim();
    const docVal = String(fullValues[docRowIdx][c]).trim();
    if (docVal !== "" || (/^\d+$/.test(normalizeId_Dynamic_(hVal)) && !/曜日|判定|備考/.test(hVal))) {
      firstDocColIdx = c;
      break;
    }
  }

  let calendarEndRowIdx = fullValues.length;
  for (let r = calStartRowIdx; r < fullValues.length; r++) {
    if (!(fullValues[r][dateColIdx] instanceof Date)) {
      calendarEndRowIdx = r;
      break;
    }
  }
  return { docRowIdx, idRowIdx, calStartRowIdx, dateColIdx, firstDocColIdx, calendarEndRowIdx };
}