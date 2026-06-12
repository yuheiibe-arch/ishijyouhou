// ==========================================
// 【Core_AttendanceBatch】勤怠シート一括作成（設計完全一致・カレンダー自動転記版）
// ==========================================

function generateAttendance_Core(targetYear, type) {
  if (!targetYear) { console.error("年度未指定"); return; }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const yearNum = String(targetYear).replace("年度", "");
  const contractSheetName = `${type}${yearNum}年度`;
  const attSheetName = `${type}勤怠${yearNum}`;

  const contractSheet = ss.getSheetByName(contractSheetName);
  let attSheet = ss.getSheetByName(attSheetName);

  if (!contractSheet) {
    throw new Error(`マスタシート未検出: ${contractSheetName}`);
  }

  // --- 1. マスタデータの読み込み（動的取得 ＆ 空行スキップ） ---
  const contractAllData = contractSheet.getDataRange().getValues();
  const contractHeaders = contractAllData[0];
  
  const idxName = contractHeaders.indexOf("医師名");
  const idxMedId = contractHeaders.indexOf("医籍番号"); // ★追加：医籍番号
  const idxEntry = contractHeaders.indexOf("入職日");
  const idxRetire = contractHeaders.indexOf("退職日"); 
  const idxRemarks = contractHeaders.indexOf("勤務備考");
  const idxHoliday = contractHeaders.indexOf("祝日"); 
  const idxNewYear = contractHeaders.indexOf("年末年始");

  if (idxName === -1 || idxRemarks === -1) throw new Error("マスタ必須列（医師名 / 勤務備考）が見つかりません");
  if (idxMedId === -1) throw new Error("マスタ必須列（医籍番号）が見つかりません。勤務シートを確認してください。");

  // プルダウン等の空行を無視する「真の最終行」を取得
  let trueLastRow = 1;
  for (let i = contractAllData.length - 1; i >= 1; i--) {
    if (contractAllData[i][idxName] !== "") {
      trueLastRow = i + 1;
      break;
    }
  }

  const masterDoctors = [];
  const contractMap = new Map(); 
  const FY_START = new Date(`${yearNum}/04/01`);
  const FY_END = new Date(`${parseInt(yearNum) + 1}/03/31`);

  for (let i = 1; i < trueLastRow; i++) {
    const row = contractAllData[i];
    const rawName = row[idxName];
    if (!rawName) continue;
    const docName = String(rawName).trim().replace(/\s+/g, ""); 
    const medId = String(row[idxMedId]).trim(); // ★追加：医籍番号取得
    
    if (!contractMap.has(docName)) {
        masterDoctors.push({ name: String(rawName).trim(), key: docName, medId: medId });
    }

    let entryDate = row[idxEntry] instanceof Date ? row[idxEntry] : FY_START;
    if (entryDate < FY_START) entryDate = FY_START;
    
    let retireDate = FY_END;
    if (idxRetire !== -1 && row[idxRetire] instanceof Date) {
      retireDate = row[idxRetire];
    }
    const hRule = idxHoliday !== -1 ? String(row[idxHoliday]) : "ー";
    const nRule = idxNewYear !== -1 ? String(row[idxNewYear]) : "ー";
    
    let remarks = String(row[idxRemarks]);
    remarks = standardizeRemarksFormat(remarks); 
    const blocks = parseRemarksToTextBlocks(remarks);
    contractMap.set(docName, { medId, entryDate, retireDate, hRule, nRule, blocks }); // ★追加: medId
  }

  // --- 2. 勤怠シートの準備（なければ外部から自動生成） ---
  if (!attSheet) {
    console.log(`勤怠シート「${attSheetName}」を作成し、外部カレンダーから日付を取得します。`);
    try {
      // 外部ソースID
      const sourceSS = SpreadsheetApp.openById('1WlmirSDOPnIcV2cwY5ClXkMWBFMM-4zrAw4XFDNUPGw');
      const sourceSheetName = `${yearNum}年度`; 
      const sourceSheet = sourceSS.getSheetByName(sourceSheetName);
      
      if (!sourceSheet) throw new Error(`外部カレンダーソースに「${sourceSheetName}」が見つかりません。`);

      // 新しいシートを作成
      attSheet = ss.insertSheet(attSheetName);
      
      // 外部の2行目以降（A2:F〜末尾）を取得
      const sourceLastRow = sourceSheet.getLastRow();
      const sourceData = sourceSheet.getRange(2, 1, sourceLastRow - 1, 6).getValues();
      
      // ★修正：新シートの2行目（A2:F）から貼り付け（これにより2行目が見出し、3行目から日付になる）
      attSheet.getRange(2, 1, sourceData.length, 6).setValues(sourceData);
      
      // ★修正：見出し行（2行目A〜F列）のデザイン調整のみ実行（余計な文字上書きを削除）
      attSheet.getRange(2, 1, 1, 6).setBackground("#d9ead3").setFontWeight("bold").setHorizontalAlignment("center");
      
    } catch (e) {
      throw new Error(`カレンダーの自動生成に失敗しました: ${e.message}`);
    }
  }

  // ★重要対策：入力規則（プルダウン）による書き込みエラーを防止
  const maxR = attSheet.getMaxRows();
  const maxC = attSheet.getMaxColumns();
  if (maxC >= 7) {
    // ★変更：3行目・G列目以降の入力エリアの制限を完全解除
    attSheet.getRange(3, 7, maxR - 2, maxC - 6).setDataValidation(null);
  }

  // --- 3. 勤怠シートの列チェックと医師の追加 ---
  let attLastCol = attSheet.getLastColumn();
  if (attLastCol < 6) attLastCol = 6; 
  // ★変更：医師名が見出しの「1行目」になるので、1行目を取得
  let attHeaders = attSheet.getRange(1, 1, 1, attLastCol).getValues()[0];
  
  const existingDocKeys = new Set();
  for (let c = 6; c < attLastCol; c++) {
    const hName = String(attHeaders[c]).trim().replace(/\s+/g, "");
    if (hName) existingDocKeys.add(hName);
  }

  const missingDoctors = masterDoctors.filter(d => !existingDocKeys.has(d.key));

  if (missingDoctors.length > 0) {
    const newNames = missingDoctors.map(d => d.name); 
    const newMedIds = missingDoctors.map(d => d.medId); 

    // ★変更：1行目に医師名（名前）をセット
    attSheet.getRange(1, attLastCol + 1, 1, newNames.length)
            .setValues([newNames])
            .setBackground("#fff2cc") // 名前の背景色
            .setFontWeight("bold")
            .setHorizontalAlignment("center");

    // ★変更：2行目に医籍番号をセット
    attSheet.getRange(2, attLastCol + 1, 1, newMedIds.length)
            .setValues([newMedIds])
            .setBackground("#fce5cd") // 医籍番号の背景色
            .setHorizontalAlignment("center");
    
    attLastCol = attSheet.getLastColumn();
    attHeaders = attSheet.getRange(1, 1, 1, attLastCol).getValues()[0];
  }

  // ★変更：すでにシートにいる医師も含めて、全員の「医籍番号（2行目）」を強制上書き！
  if (attLastCol >= 7) {
    const currentMedIds = [];
    for (let c = 6; c < attLastCol; c++) {
      const docName = String(attHeaders[c]).trim().replace(/\s+/g, "");
      if (docName && contractMap.has(docName)) {
        currentMedIds.push(contractMap.get(docName).medId);
      } else {
        currentMedIds.push("");
      }
    }
    attSheet.getRange(2, 7, 1, currentMedIds.length)
            .setValues([currentMedIds])
            .setBackground("#fce5cd")
            .setHorizontalAlignment("center");
  }

  // --- 4. 勤怠シートデータの読み込みと計算 ---
  const attLastRow = attSheet.getLastRow();
  // ★変更：3行目から下を読み込み
  const attData = attSheet.getRange(3, 1, attLastRow - 2, attLastCol).getValues();

  const doctorColMap = new Map();
  const doctorYearlyStats = new Map();
  const monthlyStats = new Map();

  // 医師名と列番号（G列=6以降）をマッピング
  for (let c = 6; c < attLastCol; c++) { 
    const docName = String(attHeaders[c]).trim().replace(/\s+/g, "");
    if (docName) {
      doctorColMap.set(docName, c);
      doctorYearlyStats.set(docName, { totalHours: 0, totalDays: 0, holidayCount: 0 });
    }
  }

  const calendarGrid = [];    
  const calendarBgGrid = [];

  for (let r = 0; r < attData.length; r++) {
    const rowData = attData[r];
    const dateVal = new Date(rowData[0]); 
    const isHoliday = (rowData[3] === "祝日" || rowData[3] === "祝"); 
    const isNewYear = (rowData[5] === "年末年始" || rowData[5] === "年末"); 
    
    if (isNaN(dateVal.getTime())) break;

    const dayShifts = new Array(attLastCol - 6).fill(""); 
    const dayBgs = new Array(attLastCol - 6).fill(null); 

    const thisMonthStr = `${dateVal.getFullYear()}/${dateVal.getMonth() + 1}`;
    if (!monthlyStats.has(thisMonthStr)) monthlyStats.set(thisMonthStr, new Map());
    const currentMonthMap = monthlyStats.get(thisMonthStr);

    doctorColMap.forEach((colIndex, docName) => {
      const contract = contractMap.get(docName);
      if (!contract) return; 

      if (!currentMonthMap.has(docName)) currentMonthMap.set(docName, { hours: 0, days: 0 });
      const mStats = currentMonthMap.get(docName);
      const yStats = doctorYearlyStats.get(docName);
      const arrIdx = colIndex - 6;

      if (dateVal < contract.entryDate || dateVal > contract.retireDate) {
        dayShifts[arrIdx] = "";
        dayBgs[arrIdx] = "#d9d9d9"; 
        return;
      }
      
      const activeBlock = contract.blocks.find(b => {
         if (!b.startDate) return true;
         return dateVal >= b.startDate && dateVal <= b.endDate;
      });

      if (!activeBlock) {
        dayShifts[arrIdx] = "休";
        dayBgs[arrIdx] = "#f3f3f3"; 
        yStats.holidayCount++; 
        return;
      }

      const shiftText = getShiftForDay(dateVal, activeBlock.lines, isHoliday, contract.hRule, isNewYear, contract.nRule);

      if (shiftText) {
        dayShifts[arrIdx] = shiftText;
        dayBgs[arrIdx] = null; 
        const dailyHours = calculateWorkHours_AbsoluteBreak(shiftText);
        yStats.totalHours += dailyHours;
        yStats.totalDays++;
        mStats.hours += dailyHours;
        mStats.days++;
      } else {
        dayShifts[arrIdx] = "休";
        dayBgs[arrIdx] = "#f3f3f3"; 
        yStats.holidayCount++; 
      }
    });
    calendarGrid.push(dayShifts);
    calendarBgGrid.push(dayBgs);
  }

  // --- 5. 書き込み ---
  const calendarRowCount = calendarGrid.length;
  if (calendarRowCount > 0) {
    const numCols = calendarGrid[0].length;
    // ★変更：3行目7列目から結果を書き込み
    const targetRange = attSheet.getRange(3, 7, calendarRowCount, numCols);
    targetRange.setValues(calendarGrid);
    targetRange.setBackgrounds(calendarBgGrid);
  }

  // --- 6. 集計欄の書き込み ---
  writeSummarySection(attSheet, attLastCol, attHeaders, monthlyStats, doctorYearlyStats, contractMap, calendarRowCount);

  // ★★★ 確実な空行削除処理 ★★★
  const finalMaxRows = attSheet.getMaxRows();
  const dataLastRow = attSheet.getLastRow();

  if (dataLastRow > 0) {
    const allValues = attSheet.getRange(1, 1, dataLastRow, attSheet.getLastColumn()).getValues();
    let trueLastRow = 1;
    
    for (let i = allValues.length - 1; i >= 0; i--) {
      if (allValues[i].join("").trim() !== "") {
        trueLastRow = i + 1;
        break;
      }
    }
    
    if (finalMaxRows > trueLastRow) {
      attSheet.deleteRows(trueLastRow + 1, finalMaxRows - trueLastRow);
    }
  }

  console.log(`[${type}] ${yearNum}年度 勤怠書き込み完了`);
}

/**
 * 集計欄書き込みヘルパー（一括処理専用）
 */
function writeSummarySection(sheet, lastCol, headers, monthlyStats, yearlyStats, contractMap, startRowOffset) {
  const summaryValues = [];
  const summaryBackgrounds = [];
  
  const createEmptyRow = (valA, bgColorA, bgColorData) => {
    const rowVal = new Array(lastCol).fill("");
    rowVal[0] = valA;
    const rowBg = new Array(lastCol).fill(null);
    for(let i=0; i<6; i++) rowBg[i] = bgColorA; 
    for(let i=6; i<lastCol; i++) rowBg[i] = bgColorData || null; 
    return { val: rowVal, bg: rowBg };
  };

  const separatorRow = new Array(lastCol).fill("");
  const separatorBg = new Array(lastCol).fill("black");
  summaryValues.push(separatorRow);
  summaryBackgrounds.push(separatorBg);

  const monthEntries = Array.from(monthlyStats.entries());
  for (const [monthStr, docMap] of monthEntries) {
    const [mYear, mMonth] = monthStr.split('/').map(Number);
    const monthStart = new Date(mYear, mMonth - 1, 1);
    const monthEnd = new Date(mYear, mMonth, 0);

    const rowObj = createEmptyRow(`${mYear}年${mMonth}月`, "#e0f7fa", null);
    
    for (let c = 6; c < lastCol; c++) {
      const docName = String(headers[c]).trim().replace(/\s+/g, "");
      if (docName && contractMap.has(docName)) {
        const contract = contractMap.get(docName);
        if (contract.retireDate < monthStart) {
          rowObj.val[c] = "退職";
        } else if (contract.entryDate > monthEnd) {
          rowObj.val[c] = "-";
        } else {
          const stats = docMap.has(docName) ? docMap.get(docName) : { hours: 0, days: 0 };
          rowObj.val[c] = `${parseFloat(stats.hours.toFixed(2))}h(${stats.days}日)`;
        }
      } else {
        rowObj.val[c] = "-";
      }
    }
    summaryValues.push(rowObj.val);
    summaryBackgrounds.push(rowObj.bg);
  }

  const labels = ["年間労働時間", "年間勤務日数", "年間休日数"];
  const keys = ["totalHours", "totalDays", "holidayCount"];
  
  labels.forEach((label, idx) => {
    const rowObj = createEmptyRow(label, "#fff2cc", "#fff2cc");
    for (let c = 6; c < lastCol; c++) {
      const docName = String(headers[c]).trim().replace(/\s+/g, "");
      if (docName && yearlyStats.has(docName)) {
        const val = yearlyStats.get(docName)[keys[idx]];
        if (keys[idx] === "totalHours") {
          rowObj.val[c] = parseFloat(val.toFixed(2)) + "h";
        } else {
          rowObj.val[c] = val + "日";
        }
      }
    }
    summaryValues.push(rowObj.val);
    summaryBackgrounds.push(rowObj.bg);
  });

  // ★変更：カレンダー開始位置が1行下にズレたため、集計欄の開始位置も+1（3行目ベース）に調整
  const startRow = 3 + startRowOffset; // ヘッダー行(1,2) + カレンダー行数
  const totalSummaryRows = summaryValues.length;
  
  const maxRows = sheet.getMaxRows();
  if (maxRows >= startRow) {
    const rangeToClear = sheet.getRange(startRow, 1, maxRows - startRow + 1, lastCol);
    rangeToClear.clear();        
    rangeToClear.breakApart();   
  }

  const targetRange = sheet.getRange(startRow, 1, totalSummaryRows, lastCol);
  targetRange.setValues(summaryValues);
  targetRange.setBackgrounds(summaryBackgrounds);

  const yearlyStartRow = startRow + totalSummaryRows - 3;
  sheet.getRange(yearlyStartRow, 1, 3, lastCol).setFontWeight("bold");
  sheet.getRange(startRow, 1, totalSummaryRows, 6).setHorizontalAlignment("center"); 

  // ★修正：セル結合（merge）をコメントアウトして解除
  // for (let i = 1; i < totalSummaryRows; i++) {
  //   sheet.getRange(startRow + i, 1, 1, 6).merge(); 
  // }
}