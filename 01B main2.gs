// ==========================================
// 【機能分割A】カレンダー実績反映（色塗り＆契約・確定併記）
// ※ 完全動的列検索・過去確定データ保護ロック・アーカイブ自動修復搭載
// ==========================================

const EXCEPTION_SS_ID = "10yPdoOOgqSSGKwoiPAXi83YM9vb_Em8r6Ex-bLfg28M"; 
const SHIFT_SS_ID = "1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co";     
const MASTER_SS_ID = "14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs";    
const CLOSED_SS_ID = "1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA";    
const ATTENDANCE_SS_ID = ""; 

const IGNORE_SHIFTS = [
  "出張インフルエンザワクチン", "【埼玉】ワクチンバックアップシフト", 
  "職域新型コロナワクチン接種", "【千葉】ワクチンバックアップシフト", 
  "【東京】ワクチンバックアップシフト", 
  "【関東】バックアップシフト", "嘱託医業務", "医師会業務"
];

const SPECIAL_CODE_MAP = {
  "1000": "院外勤務（小児科）", "1002": "嘱託医業務", "1004": "医師会業務",
  "1005": "【関東】バックアップシフト", "1017": "有給", "1018": "欠勤"
};

const SHIFT_REGEX = /【(.*?)】(\d{1,2})[:]?(\d{0,2})\s*[-~～]\s*(\d{1,2})[:]?(\d{0,2})/g;
const DATE_REGEX = /(\d{4})[\/年](\d{1,2})[\/月](\d{1,2})/;
const TIME_REGEX = /(\d{1,2}):(\d{2})/;

/**
 * 勤怠シートの見出し行・日付列・医師開始列・カレンダー終了行をすべて動的に特定するヘルパー
 */
function detectAttendanceSheetStructure(fullValues) {
  let dateColIdx = 0;
  let calHeaderRowIdx = 1;

  for (let r = 0; r < Math.min(5, fullValues.length); r++) {
    const idx = fullValues[r].findIndex(cell => String(cell).trim() === "日付");
    if (idx !== -1) {
      calHeaderRowIdx = r;
      dateColIdx = idx;
      break;
    }
  }

  const docRowIdx = calHeaderRowIdx > 0 ? calHeaderRowIdx - 1 : 0;
  const idRowIdx = calHeaderRowIdx;
  const calStartRowIdx = Math.max(docRowIdx, idRowIdx) + 1;

  let firstDocColIdx = dateColIdx + 1;
  for (let c = dateColIdx + 1; c < fullValues[idRowIdx].length; c++) {
    const hVal = String(fullValues[idRowIdx][c] || "").trim();
    const docVal = String(fullValues[docRowIdx][c] || "").trim();
    const normId = typeof normalizeMedId === "function" ? normalizeMedId(hVal) : hVal;
    if (docVal !== "" || (/^\d+$/.test(normId) && !/曜日|判定|備考/.test(hVal))) {
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

function updateAttendanceActuals() {
  const activeSs = SpreadsheetApp.getActiveSpreadsheet();
  let targetSs = activeSs;

  if (ATTENDANCE_SS_ID !== "") {
    try { targetSs = SpreadsheetApp.openById(ATTENDANCE_SS_ID); } 
    catch (e) { console.error("勤怠ファイルの読み込みエラー: " + e.message); return; }
  }

  const allTargetSheets = targetSs.getSheets().filter(sheet => sheet.getName().match(/^(常勤|定期非常勤)勤怠\d{4}(年度)?$/));
  let latestFullTime = null;
  let latestPartTime = null;
  let maxFullYear = 0;
  let maxPartYear = 0;

  allTargetSheets.forEach(sheet => {
    const match = sheet.getName().match(/^(常勤|定期非常勤)勤怠(\d{4})/);
    if (match) {
      const type = match[1];
      const year = parseInt(match[2], 10);
      if (type === "常勤" && year > maxFullYear) {
        maxFullYear = year;
        latestFullTime = sheet;
      } else if (type === "定期非常勤" && year > maxPartYear) {
        maxPartYear = year;
        latestPartTime = sheet;
      }
    }
  });

  const targetSheets = [];
  if (latestFullTime) targetSheets.push(latestFullTime);
  if (latestPartTime) targetSheets.push(latestPartTime);

  if (targetSheets.length === 0) return;

  console.log("=== [色塗り処理] 外部データの読み込み開始 ===");
  const locMaster = fetchLocationMaster();
  const closedData = fetchExternalClosedDays();

  // まず現行の「確定シフト」を読み込み、その最小日付（例: 2026/10/01）を確認
  let actualShifts = fetchActualShifts(locMaster);
  const minActiveShiftDate = actualShifts.minDate || "9999/99/99";

  // 対象シートの過去日（minActiveShiftDate より前）に「確定：未登録」等で壊れている月があるか事前スキャン
  // ※また、勤怠シート一括作成の直後（まだ一度も実績集計区切りが作られていない状態）も過去月アーカイブを読み込む
  const neededArchiveMonths = new Set();
  targetSheets.forEach(sheet => {
    const vals = getTrueData(sheet);
    const hasExistingActualSummary = vals.some(row => row.some(cell => String(cell).trim() === "実績集計区切り"));
    const st = detectAttendanceSheetStructure(vals);

    for (let r = st.calStartRowIdx; r < st.calendarEndRowIdx; r++) {
      const dCell = vals[r][st.dateColIdx];
      if (!(dCell instanceof Date)) continue;
      const dKey = Utilities.formatDate(dCell, "JST", "yyyy/MM/dd");
      if (dKey >= minActiveShiftDate) continue;

      const ym = dKey.substring(0, 7); // "2026/04"
      if (!hasExistingActualSummary) {
        neededArchiveMonths.add(ym);
      } else {
        for (let c = st.firstDocColIdx; c < vals[r].length; c++) {
          if (String(vals[r][c] || "").includes("確定：未登録")) {
            neededArchiveMonths.add(ym);
            break;
          }
        }
      }
    }
  });

  // 修復・初期照合が必要な過去月があれば、アーカイブスプシから該当月だけ追加読込
  if (neededArchiveMonths.size > 0) {
    console.log(`📦 過去月の復元・初期照合のためアーカイブシートを読み込みます: ${Array.from(neededArchiveMonths).sort().join(", ")}`);
    actualShifts = fetchActualShifts(locMaster, neededArchiveMonths);
  } else {
    console.log(`🛡️ 過去月(${minActiveShiftDate}より前)はすべて確定済みのため保護ロックを適用します（アーカイブ読込不要）。`);
  }

  const exceptionSs = SpreadsheetApp.openById(EXCEPTION_SS_ID);
  const absenceData = fetchAbsenceData(exceptionSs);
  const subData = fetchSubstitutionData(exceptionSs);

  targetSheets.forEach(sheet => {
    console.log(`\n処理中: ${sheet.getName()}`);
    
    // ① 過去の「実績集計区切り」以降があれば先に削除してから正確なデータを取得する
    const preValues = getTrueData(sheet);
    const oldSepIdx = preValues.findIndex(row => row.some(cell => String(cell).trim() === "実績集計区切り"));
    if (oldSepIdx !== -1) {
      sheet.deleteRows(oldSepIdx + 1, sheet.getLastRow() - oldSepIdx);
    }

    const lastCol = sheet.getLastColumn();
    const fullValues = getTrueData(sheet);
    const st = detectAttendanceSheetStructure(fullValues);
    const { docRowIdx, idRowIdx, calStartRowIdx, dateColIdx, firstDocColIdx, calendarEndRowIdx } = st;

    if (calendarEndRowIdx <= calStartRowIdx) return;

    const topHeaders = fullValues[docRowIdx];
    const cleanHeaders = topHeaders.map(h => String(h || "").replace(/[\s ]+/g, ""));
    const idHeaders = fullValues[idRowIdx].map(h => typeof normalizeMedId === "function" ? normalizeMedId(h) : String(h || "").replace(/[\s ]+/g, ""));
    
    const calRowCount = calendarEndRowIdx - calStartRowIdx;
    const calRange = sheet.getRange(calStartRowIdx + 1, 1, calRowCount, lastCol);
    const calValues = calRange.getValues();
    const calBackgrounds = calRange.getBackgrounds();
    
    const summaryDiffs = {};

    for (let r = 0; r < calValues.length; r++) {
      const dateCell = calValues[r][dateColIdx];
      if (!(dateCell instanceof Date)) continue;
      
      const dateKey = Utilities.formatDate(dateCell, "JST", "yyyy/MM/dd");
      const monthKey = `${dateCell.getFullYear()}年${dateCell.getMonth() + 1}月`;
      const ymKey = dateKey.substring(0, 7);
      const isClosedDay = closedData.has(dateKey);
      const closedLocations = isClosedDay ? closedData.get(dateKey) : [];

      // ★過去月の保護判定：現行確定シフトより前の日付で、かつアーカイブ再読込対象ではない確定済み月か？
      const isProtectedPastMonth = (dateKey < minActiveShiftDate) && !neededArchiveMonths.has(ymKey);

      for (let c = firstDocColIdx; c < lastCol; c++) {
        const docNameClean = cleanHeaders[c];
        const docMedId = idHeaders[c]; 
        
        if (!docMedId || !docNameClean) continue;
        
        if (!summaryDiffs[docNameClean]) {
          summaryDiffs[docNameClean] = { annual: { absH: 0, absD: 0, subH: 0, subD: 0, paid: 0 }, months: {} };
        }
        if (!summaryDiffs[docNameClean].months[monthKey]) {
          summaryDiffs[docNameClean].months[monthKey] = { absH: 0, absD: 0, subH: 0, subD: 0, dates: [] };
        }

        const existingCellStr = String(calValues[r][c] || "");

        // =========================================================================
        // 🛡️ 【過去確定セルの保護ロック】
        // 既に確定している過去月（かつ「確定：未登録」ではないセル）は一切書き換えず、
        // セル内の確定表記（またはお休み・振替情報）から集計値だけをカウントして次へ進む
        // =========================================================================
        if (isProtectedPastMonth && !existingCellStr.includes("確定：未登録")) {
          let baseShiftStr = existingCellStr;
          if (baseShiftStr.includes("契約：")) {
            const m = baseShiftStr.match(/契約：([\s\S]*?)(?:\n確定：|$)/);
            if (m) baseShiftStr = m[1].trim();
          }
          const baseClean = baseShiftStr.split('\n').filter(line => !line.startsWith('→') && !line.includes('※振替') && !line.includes('半日有給') && !line.includes('有給') && !line.includes('欠勤') && !line.includes('移動依頼')).join('\n');
          const plannedForCalc = (!baseClean || baseClean === "休" || baseClean === "-") ? [] : extractPlannedShifts(baseClean);

          if (existingCellStr.includes("※振替")) {
            const subInfo = subData.get(`${dateKey}_${docMedId}`);
            const subH = subInfo ? subInfo.hours : calculatePlannedHours(extractPlannedShifts(existingCellStr));
            summaryDiffs[docNameClean].months[monthKey].subH += subH;
            summaryDiffs[docNameClean].months[monthKey].subD += 1;
            summaryDiffs[docNameClean].annual.subH += subH;
            summaryDiffs[docNameClean].annual.subD += 1;
          } else if (existingCellStr.includes("→欠勤") || existingCellStr.includes("欠勤")) {
            const deductedHours = calculatePlannedHours(plannedForCalc);
            summaryDiffs[docNameClean].months[monthKey].absH += deductedHours;
            summaryDiffs[docNameClean].months[monthKey].absD += 1;
            summaryDiffs[docNameClean].months[monthKey].dates.push(`${dateCell.getMonth() + 1}/${dateCell.getDate()}`);
            summaryDiffs[docNameClean].annual.absH += deductedHours;
            summaryDiffs[docNameClean].annual.absD += 1;
          } else if (existingCellStr.includes("→半日有給")) {
            summaryDiffs[docNameClean].annual.paid += 0.5;
          } else if (existingCellStr.includes("→有給") || existingCellStr.includes("有給") || existingCellStr.includes("有休")) {
            summaryDiffs[docNameClean].annual.paid += 1.0;
          }
          continue; // セルの文字・背景色は一切変更しない
        }

        // =========================================================================
        // 🔄 通常の照合・更新処理（当月以降 または アーカイブからの修復対象月）
        // =========================================================================
        const compositeKey = `${dateKey}_${docMedId}`;
        let currentVal = existingCellStr;

        // 期間外（#d9d9d9）や休日（#f3f3f3）の背景色を維持しつつ、前回の赤・黄・青をリセット
        if (calBackgrounds[r][c] !== "#d9d9d9") {
          calBackgrounds[r][c] = (currentVal === "休") ? "#f3f3f3" : null;
        }

        if (currentVal.includes('契約：')) {
          const match = currentVal.match(/契約：([\s\S]*?)(?:\n確定：|$)/);
          if (match) currentVal = match[1].trim(); 
        }

        let cleanLines = currentVal.split('\n').filter(line => !line.startsWith('→') && !line.includes('※振替') && !line.includes('半日有給') && !line.includes('有給') && !line.includes('欠勤') && !line.includes('移動依頼'));
        currentVal = cleanLines.join('\n');
        
        const isOriginallyOff = (!currentVal || currentVal === "休" || currentVal === "-" || currentVal === "");
        const plannedShifts = isOriginallyOff ? [] : extractPlannedShifts(currentVal);
        
        let hasException = false;
        let appendText = "";
        let paidCountForDay = 0;

        const actualsForDoctorDay = actualShifts.get(compositeKey) || [];
        const hasPaidShiftInActual = actualsForDoctorDay.some(act => 
          act.clinic.includes("有給") || act.remarks.includes("有給") || act.clinic.includes("有休") || act.remarks.includes("有休")
        );
        const hasAbsenceShiftInActual = actualsForDoctorDay.some(act => 
          act.clinic.includes("欠勤") || act.dept.includes("欠勤") || act.remarks.includes("欠勤")
        );
        const hasWorkShiftInActual = actualsForDoctorDay.some(act => 
          !act.clinic.includes("有給") && !act.clinic.includes("有休") && !act.remarks.includes("有給") && !act.remarks.includes("有休") && !act.clinic.includes("欠勤") && !act.remarks.includes("欠勤")
        );

        if (absenceData.has(compositeKey)) {
          const absenceType = String(absenceData.get(compositeKey));
          if (absenceType.includes("欠勤")) {
            appendText += `\n→${absenceType}`;
            hasException = true;
            const deductedHours = calculatePlannedHours(plannedShifts);
            summaryDiffs[docNameClean].months[monthKey].absH += deductedHours;
            summaryDiffs[docNameClean].months[monthKey].absD += 1;
            summaryDiffs[docNameClean].months[monthKey].dates.push(`${dateCell.getMonth() + 1}/${dateCell.getDate()}`);
            summaryDiffs[docNameClean].annual.absH += deductedHours;
            summaryDiffs[docNameClean].annual.absD += 1;

            if (!hasAbsenceShiftInActual && deductedHours > 0 && !subData.has(compositeKey)) {
              calBackgrounds[r][c] = "#e8f0fe"; 
            }
          } else if (absenceType.includes("有給") || absenceType.includes("有休")) {
            hasException = true;
            if (hasWorkShiftInActual || absenceType.includes("半休") || absenceType.includes("半日") || absenceType.includes("午前") || absenceType.includes("午後")) {
              paidCountForDay = 0.5;
              appendText += `\n→半日有給`;
            } else {
              paidCountForDay = 1.0;
              appendText += `\n→${absenceType}`;
            }
          } else {
            appendText += `\n→${absenceType}`;
            hasException = true;
          }
        } 
        else if (hasPaidShiftInActual) {
          hasException = true;
          if (hasWorkShiftInActual) {
            paidCountForDay = 0.5;
            appendText += `\n→半日有給`;
          } else {
            paidCountForDay = 1.0;
            appendText += `\n→有給`;
          }
        }
        else if (hasAbsenceShiftInActual) {
          hasException = true;
          appendText += `\n→欠勤`;
          const deductedHours = calculatePlannedHours(plannedShifts);
          summaryDiffs[docNameClean].months[monthKey].absH += deductedHours;
          summaryDiffs[docNameClean].months[monthKey].absD += 1;
          summaryDiffs[docNameClean].months[monthKey].dates.push(`${dateCell.getMonth() + 1}/${dateCell.getDate()}`);
          summaryDiffs[docNameClean].annual.absH += deductedHours;
          summaryDiffs[docNameClean].annual.absD += 1;
        }
        else if (isClosedDay) {
          const matchedLocation = closedLocations.find(loc => currentVal.includes(loc));
          if (matchedLocation) {
            appendText += `\n→休館日`;
            hasException = true;
          }
        }

        if (subData.has(compositeKey)) {
          const subInfo = subData.get(compositeKey);
          currentVal = subInfo.text; 
          appendText = ""; 
          hasException = true;
          
          summaryDiffs[docNameClean].months[monthKey].subH += subInfo.hours;
          summaryDiffs[docNameClean].months[monthKey].subD += 1;
          summaryDiffs[docNameClean].annual.subH += subInfo.hours;
          summaryDiffs[docNameClean].annual.subD += 1;
        }

        summaryDiffs[docNameClean].annual.paid += paidCountForDay;

        let isMismatchActual = false;
        let isMissingActual = false;
        let transferAppendText = "";

        const dayHasAllowance = actualsForDoctorDay.some(act => act.hasAllowance);
        const dayHasGuaranteedPay = actualsForDoctorDay.some(act => act.isGuaranteedPay);
        const dayIsSpecial = dayHasAllowance || dayHasGuaranteedPay;

        if (plannedShifts.length > 0 && !hasException) {
          plannedShifts.forEach(plan => {
            const planLocNorm = normalizeLocation(plan.loc, locMaster);
            const matchFound = actualsForDoctorDay.some(act => {
              const actLocNorm = normalizeLocation(act.clinic, locMaster, act.dept);
              const locMatch = actLocNorm.includes(planLocNorm) || planLocNorm.includes(actLocNorm);
              const sM_act = timeToMins(act.start);
              const eM_act = timeToMins(act.end);
              const sM_plan = timeToMins(plan.start);
              const timeMatch = (act.start === plan.start) || (sM_act <= sM_plan && sM_plan < eM_act);
              return locMatch && timeMatch;
            });
            
            if (!matchFound) {
              if (dayIsSpecial) {
                actualsForDoctorDay.forEach(act => {
                  const sTime = act.start.replace(/:00$/, "");
                  const eTime = act.end.replace(/:00$/, "");
                  const formalClinic = normalizeLocation(act.clinic, locMaster, act.dept);
                  const actText = `\n→移動依頼【${formalClinic}】${sTime}-${eTime}`;
                  if (!transferAppendText.includes(actText)) transferAppendText += actText;
                });
              } else {
                if (actualsForDoctorDay.length > 0) isMismatchActual = true; 
                else isMissingActual = true;  
              }
            }
          });
        }

        if (!hasException && isMismatchActual) {
          let actualShiftTexts = actualsForDoctorDay.map(act => {
            const sTime = act.start.replace(/:00$/, "");
            const eTime = act.end.replace(/:00$/, "");
            const formalClinic = normalizeLocation(act.clinic, locMaster, act.dept);
            return `【${formalClinic}】${sTime}-${eTime}`;
          }).join("、");
          
          if (!actualShiftTexts) actualShiftTexts = "情報なし";

          currentVal = `契約：${currentVal}\n確定：${actualShiftTexts}`;
          calBackgrounds[r][c] = "red";
          
        } else if (!hasException && isMissingActual) {
          currentVal = `契約：${currentVal}\n確定：未登録`;
          calBackgrounds[r][c] = "yellow"; 
        }

        calValues[r][c] = currentVal + appendText + transferAppendText;
      }
    }

    calRange.setValues(calValues);
    calRange.setBackgrounds(calBackgrounds);

    console.log(`-> カレンダー更新完了。新・集計表を作成します...`);

    // ② 元の集計表を calendarEndRowIdx から1行のズレもなく取得する
    const origSummaryVals = fullValues.slice(calendarEndRowIdx);
    
    const appendStartRow = sheet.getLastRow() + 1;
    sheet.getRange(appendStartRow, dateColIdx + 1).setValue("実績集計区切り").setFontColor("white").setBackground("black");
    sheet.getRange(appendStartRow, 1, 1, lastCol).setBackground("black");

    const newSummaryStart = appendStartRow + 1;
    const newSummaryVals = [];
    const newSummaryRichTexts = []; 
    const newSummaryBackgrounds = []; 

    const emptyRichCell = SpreadsheetApp.newRichTextValue().setText("").build();

    for (let sr = 0; sr < origSummaryVals.length; sr++) {
      let rawLabel = origSummaryVals[sr][dateColIdx];
      if (!rawLabel) continue;
      let rowLabel = (rawLabel instanceof Date) ? `${rawLabel.getFullYear()}年${rawLabel.getMonth() + 1}月` : String(rawLabel).trim();

      let isMonthRow = rowLabel.match(/\d{4}年\d{1,2}月/);
      if (isMonthRow) {
        let isAllZeroOrEmpty = true;
        for (let c = firstDocColIdx; c < lastCol; c++) {
          if (!cleanHeaders[c]) continue;
          const origVal = String(origSummaryVals[sr][c] || "").replace(/\s/g, "");
          if (origVal !== "" && origVal !== "0h(0日)" && origVal !== "0h(0.0日)" && origVal !== "0" && origVal !== "-") {
            isAllZeroOrEmpty = false;
            break;
          }
        }
        if (isAllZeroOrEmpty) continue; 
      }

      const newRowVals = new Array(lastCol).fill("");
      const newRowRich = new Array(lastCol).fill(emptyRichCell);
      const newRowBgs = new Array(lastCol).fill(null);

      newRowVals[dateColIdx] = rowLabel;
      newRowRich[dateColIdx] = SpreadsheetApp.newRichTextValue().setText(rowLabel).build();
      for (let i = 0; i < firstDocColIdx; i++) newRowBgs[i] = "#f3f3f3";

      for (let c = firstDocColIdx; c < lastCol; c++) {
        const docNameClean = cleanHeaders[c];
        const origVal = String(origSummaryVals[sr][c] || "");
        
        if (!docNameClean) continue;

        const diffs = summaryDiffs[docNameClean] || { annual: { absH: 0, absD: 0, subH: 0, subD: 0, paid: 0 }, months: {} };
        let baseText = origVal;
        let suffixText = "";
        let isRed = false;
        let isBlue = false;

        let calcObj = null;
        if (isMonthRow) calcObj = diffs.months[rowLabel];
        else if (rowLabel.includes("年間労働時間") || rowLabel.includes("年間勤務日数")) calcObj = diffs.annual;

        if (calcObj && (calcObj.absH > 0 || calcObj.subH > 0)) {
          let netH = parseFloat((calcObj.subH - calcObj.absH).toFixed(2));
          let netD = calcObj.absD - calcObj.subD; 
          
          if (netH < 0) {
            let absDaysStr = netD > 0 ? `${netD}` : `${Math.abs(netD)}`;
            suffixText = `${netH}h/-${absDaysStr}日`;
            isRed = true;
          } else if (netH === 0) {
            suffixText = `(±0h/完全相殺)`;
          } else {
            suffixText = `(+${netH}h/振替増)`;
            isBlue = true; 
          }
        }
        
        let fullText = baseText + suffixText;
        newRowVals[c] = fullText;

        const richBuilder = SpreadsheetApp.newRichTextValue().setText(fullText);
        if (suffixText !== "") {
          const startIdx = fullText.lastIndexOf(suffixText);
          const endIdx = fullText.length;
          let color = isRed ? "red" : (isBlue ? "blue" : "green");
          richBuilder.setTextStyle(startIdx, endIdx, SpreadsheetApp.newTextStyle().setForegroundColor(color).build());
        }
        newRowRich[c] = richBuilder.build();
      }
      newSummaryVals.push(newRowVals);
      newSummaryRichTexts.push(newRowRich);
      newSummaryBackgrounds.push(newRowBgs);
    }

    const origDaysRow = origSummaryVals.find(row => String(row[dateColIdx]).includes("年間勤務日数"));

    const extraLabels = ["年間欠勤数", "有給取得日数", "勤怠遵守率"];
    extraLabels.forEach(label => {
      const newRowVals = new Array(lastCol).fill("");
      const newRowRich = new Array(lastCol).fill(emptyRichCell);
      const newRowBgs = new Array(lastCol).fill(null);

      newRowVals[dateColIdx] = label;
      newRowRich[dateColIdx] = SpreadsheetApp.newRichTextValue().setText(label).build();
      for (let i = 0; i < firstDocColIdx; i++) newRowBgs[i] = "#f3f3f3";

      for (let c = firstDocColIdx; c < lastCol; c++) {
        const docNameClean = cleanHeaders[c];
        if (!docNameClean) continue;

        const diffs = summaryDiffs[docNameClean] || { annual: { absH: 0, absD: 0, subH: 0, subD: 0, paid: 0 }, months: {} };
        let valText = "";
        let cellBg = null;

        let origDays = 0;
        if (origDaysRow && origDaysRow[c] !== undefined) {
          origDays = parseFloat(String(origDaysRow[c]).replace(/[^\d.]/g, "")) || 0;
        }
        
        if (origDays > 0) {
          let netAbsD = Math.max(0, diffs.annual.absD - diffs.annual.subD);
          if (label === "年間欠勤数") {
            valText = netAbsD > 0 ? `${netAbsD}日` : "0日";
          } else if (label === "有給取得日数") {
            valText = diffs.annual.paid > 0 ? `${parseFloat(diffs.annual.paid.toFixed(1))}日` : "0日";
          } else if (label === "勤怠遵守率") {
            const actualDays = origDays - netAbsD;
            const rate = Math.min(100, Math.max(0, Math.round((actualDays / origDays) * 100)));
            valText = `${rate}％`;
            if (rate < 100) cellBg = "#fce8e6"; 
          }
        }
        newRowVals[c] = valText;
        newRowRich[c] = SpreadsheetApp.newRichTextValue().setText(valText).build();
        newRowBgs[c] = cellBg;
      }
      newSummaryVals.push(newRowVals);
      newSummaryRichTexts.push(newRowRich);
      newSummaryBackgrounds.push(newRowBgs);
    });

    if (newSummaryRichTexts.length > 0) {
      const targetRange = sheet.getRange(newSummaryStart, 1, newSummaryRichTexts.length, lastCol);
      targetRange.setRichTextValues(newSummaryRichTexts);
      targetRange.setBackgrounds(newSummaryBackgrounds);
      targetRange.setFontColor("black");
      targetRange.setFontWeight("normal");
      
      sheet.getRange(newSummaryStart, dateColIdx + 1, newSummaryRichTexts.length, 1).setFontWeight("bold");
    }

    console.log(`-> シート「${sheet.getName()}」の処理完了`);
  });
  
  console.log("=== 色塗り処理が完了しました ===");
}