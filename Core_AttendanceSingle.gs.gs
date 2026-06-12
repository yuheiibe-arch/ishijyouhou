// ==========================================
// 【Core_AttendanceSingle】単独追加（フォーム連動）の共通ロジック
// ==========================================

/**
 * 単独追加の共通処理
 * @param {string} targetYear - 対象年度
 * @param {Array} headers - 契約情報シートのヘッダー配列
 * @param {Array} rowData - 契約情報シートのデータ行
 * @param {string} type - "常勤" または "定期非常勤"
 * @param {string} contractText - 自動生成されたシフト契約テキスト
 */
function appendDoctor_Single_Core(targetYear, headers, rowData, type, contractText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const yearNum = String(targetYear).replace("年度", "");
  const attSheetName = `${type}勤怠${yearNum}`;
  const attSheet = ss.getSheetByName(attSheetName);
  
  if (!attSheet) {
    console.warn(`勤怠シート「${attSheetName}」がないため、追加処理をスキップしました。`);
    return; 
  }

  console.log(`[単独追加] ${attSheetName} 処理開始`);

  // 1. 医師情報の取得
  const name = getValueByHeader(headers, rowData, "氏名") || getValueByHeader(headers, rowData, "医師名");
  if (!name) return; 

  // ★追加: 医籍番号の取得
  const medId = getValueByHeader(headers, rowData, "医籍番号") || "番号不明";

  const joinDateVal = getValueByHeader(headers, rowData, "入職日");
  
  // 引数で渡された contractText を最優先で取得
  let remarksRaw = contractText || "";
  if (!remarksRaw) {
    remarksRaw = getValueByHeader(headers, rowData, "契約情報(自動)") || getValueByHeader(headers, rowData, "勤務備考");
  }

  // 祝日・年末年始ルールの取得
  const hKey = (type === "常勤") ? "【常勤】 祝日" : "【定期非常勤】 祝日";
  const yKey = (type === "常勤") ? "【常勤】 年末年始" : "【定期非常勤】 年末年始";

  const holidayRule = unifyWording(getValueByHeader(headers, rowData, hKey) || getValueByHeader(headers, rowData, "祝日"));
  const newYearRule = unifyWording(getValueByHeader(headers, rowData, yKey) || getValueByHeader(headers, rowData, "年末年始"));
  
  const FY_START = new Date(`${yearNum}/04/01`);
  const FY_END = new Date(`${parseInt(yearNum) + 1}/03/31`);
  
  let entryDate = (joinDateVal instanceof Date) ? joinDateVal : FY_START;
  if (entryDate < FY_START) entryDate = FY_START;

  const remarks = standardizeRemarksFormat(remarksRaw);
  const blocks = parseRemarksToTextBlocks(remarks); 

  // 2. 列追加（★修正: 【案A】に合わせて 1行目に医師名、2行目に医籍番号）
  const lastCol = attSheet.getLastColumn();
  const targetCol = lastCol + 1;
  
  // 1行目：医師名
  attSheet.getRange(1, targetCol).setValue(name).setBackground("#fff2cc").setFontWeight("bold").setHorizontalAlignment("center");
  // 2行目：医籍番号
  attSheet.getRange(2, targetCol).setValue(medId).setBackground("#fce5cd").setHorizontalAlignment("center");

  // 3. カレンダー計算（★変更: 3行目から読み込み）
  const lastRow = attSheet.getLastRow();
  // 3行目から下を読み込む
  const dateValues = attSheet.getRange(3, 1, lastRow - 2, 6).getValues();

  const outputValues = [];
  const outputBackgrounds = [];

  let totalHours = 0;
  let totalDays = 0;
  let holidayCount = 0;
  const monthlyStats = new Map(); 

  for (let i = 0; i < dateValues.length; i++) {
    const row = dateValues[i];
    const dateVal = new Date(row[0]);
    if (isNaN(dateVal.getTime())) break; // 日付エリア終了でBreak

    const isHoliday = (row[3] === "祝日" || row[3] === "祝");
    const isNewYear = (row[5] === "年末年始" || row[5] === "年末");

    // 期間外の場合は、集計Mapを作成せずにスキップする
    if (dateVal < entryDate || dateVal > FY_END) {
      outputValues.push([""]); 
      outputBackgrounds.push(["#d9d9d9"]); 
      continue;
    }

    // 期間内であることが確定してから集計Mapを作る
    const monthKey = `${dateVal.getFullYear()}/${dateVal.getMonth() + 1}`;
    if (!monthlyStats.has(monthKey)) monthlyStats.set(monthKey, { h: 0, d: 0 });
    const mStat = monthlyStats.get(monthKey);

    const activeBlock = blocks.find(b => {
      if (!b.startDate) return true;
      return dateVal >= b.startDate && dateVal <= b.endDate;
    });

    if (!activeBlock) {
      outputValues.push(["休"]);
      outputBackgrounds.push(["#f3f3f3"]);
      holidayCount++;
      continue;
    }

    const shiftText = getShiftForDay(dateVal, activeBlock.lines, isHoliday, holidayRule, isNewYear, newYearRule);
    
    if (shiftText) {
      outputValues.push([shiftText]);
      outputBackgrounds.push([null]);
      const hours = calculateWorkHours_AbsoluteBreak(shiftText);
      totalHours += hours;
      totalDays++;
      mStat.h += hours;
      mStat.d++;
    } else {
      outputValues.push(["休"]);
      outputBackgrounds.push(["#f3f3f3"]);
      holidayCount++;
    }
  }

  // 4. データ書き込み（★変更: 3行目から書き込み）
  if (outputValues.length > 0) {
    attSheet.getRange(3, targetCol, outputValues.length, 1).setValues(outputValues);
    attSheet.getRange(3, targetCol, outputValues.length, 1).setBackgrounds(outputBackgrounds);
  }

  // 5. 集計欄更新（★変更: スタート行を1行下に調整）
  const summaryStartRow = 3 + outputValues.length; // ヘッダー2行(1,2) + カレンダー行数
  const summaryAreaHeight = lastRow - summaryStartRow + 1;
  
  if (summaryAreaHeight > 0) {
    const summaryLabels = attSheet.getRange(summaryStartRow, 1, summaryAreaHeight, 1).getValues().flat();
    const summaryWrite = [];
    
    for (let k = 0; k < summaryLabels.length; k++) {
      const cellVal = summaryLabels[k];
      let val = "-";
      let key = "";

      let targetYear = 0;
      let targetMonth = 0;

      if (cellVal instanceof Date) {
        targetYear = cellVal.getFullYear();
        targetMonth = cellVal.getMonth() + 1;
        key = `${targetYear}/${targetMonth}`;
      } else {
        const strVal = String(cellVal);
        const match = strVal.match(/(\d{4})[\/\年](\d{1,2})/);
        if (match) {
          targetYear = parseInt(match[1], 10);
          targetMonth = parseInt(match[2], 10);
          key = `${targetYear}/${targetMonth}`;
        }
      }

      if (String(cellVal) === "") {
        val = "";
      } else if (key) {
        if (monthlyStats.has(key)) {
          const s = monthlyStats.get(key);
          val = `${parseFloat(s.h.toFixed(2))}h(${s.d}日)`;
        } else {
          val = "-";
        }
      } else if (String(cellVal).includes("年間労働時間")) {
        val = parseFloat(totalHours.toFixed(2)) + "h";
      } else if (String(cellVal).includes("年間勤務日数")) {
        val = totalDays + "日";
      } else if (String(cellVal).includes("年間休日数")) {
        val = holidayCount + "日";
      }
      
      summaryWrite.push([val]);
    }
    
    attSheet.getRange(summaryStartRow, targetCol, summaryWrite.length, 1).setValues(summaryWrite);
    if (summaryWrite.length > 0) {
       attSheet.getRange(summaryStartRow, targetCol).setBackground("black");
    }
  }
  console.log(`[単独追加] 完了: ${name} (${type})`);
}