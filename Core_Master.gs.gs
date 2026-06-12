// ==========================================
// 【Core_Master】マスタ整備の共通ロジック（空行スキップ・空白バッファ対応・末尾空行削除版）
// ==========================================

function generateMaster_Core(targetYear, type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const yearNum = String(targetYear).replace("年度", "");
  const sheetName = `${type}${yearNum}年度`; 
  const sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    console.error(`シート「${sheetName}」が見つかりません。マスタ整備をスキップします。`);
    return;
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();
  const headers = values[0]; 

  // プルダウンだけの空行を無視するため、「医師名」が入っている本当の最終行を探す
  const idxName = headers.indexOf("医師名");
  let trueLastRow = 1;
  for (let i = 1; i < values.length; i++) {
    if (idxName > -1 && values[i][idxName] !== "") {
      trueLastRow = i + 1; 
    }
  }

  const idxEntry = headers.indexOf("入職日");
  const idxRetire = headers.indexOf("退職日"); 
  const idxRemarks = headers.indexOf("勤務備考");
  const idxWeeklyHours = headers.indexOf("週労働"); 
  let idxOutputStart = headers.indexOf("当初シフト(期間)");

  if (idxEntry === -1 || idxRemarks === -1) { 
    Browser.msgBox(`[${type}] エラー：必須列（入職日/勤務備考）が見つかりません。`); 
    return; 
  }

  let startColNum;
  if (idxOutputStart === -1) {
    startColNum = sheet.getLastColumn() + 2; 
    console.warn(`「当初シフト(期間)」列が見つかりません。空白列を挟んで ${startColNum}列目 から追加します。`);
  } else {
    startColNum = idxOutputStart + 1;
  }

  const FY_START = new Date(`${yearNum}/04/01`);
  const FY_END = new Date(`${parseInt(yearNum) + 1}/03/31`);

  const outputData = [];
  const backgroundColors = [];
  const weeklyHoursOutput = [];
  let maxBlocks = 0; 

  const analyzedRows = [];
  
  for (let i = 1; i < trueLastRow; i++) {
    const entryDateVal = values[i][idxEntry]; 
    const retireDateVal = (idxRetire > -1) ? values[i][idxRetire] : null; 
    let remarks = values[i][idxRemarks];      
    
    let entryDate = FY_START;
    if (entryDateVal instanceof Date) {
      entryDate = (entryDateVal > FY_START) ? entryDateVal : FY_START;
    }

    let validEndDate = FY_END;
    if (retireDateVal instanceof Date) {
      if (retireDateVal < FY_END) {
        validEndDate = retireDateVal;
      }
    }

    let validBlocks = [];
    if (remarks) {
      remarks = standardizeRemarksFormat(remarks); 
      const rawBlocks = parseRemarksRobust(remarks); 
      validBlocks = filterBlocksByYearAndEntry(rawBlocks, entryDate, validEndDate);
    }
    
    analyzedRows.push({ blocks: validBlocks, endDate: validEndDate });
    
    if (validBlocks.length > maxBlocks) maxBlocks = validBlocks.length;

    let latestTotal = "";
    if (validBlocks.length > 0) {
      const latestBlock = validBlocks[validBlocks.length - 1];
      const sum = latestBlock.hours.reduce((acc, cur) => acc + cur, 0);
      latestTotal = parseFloat(sum.toFixed(2));
    }
    weeklyHoursOutput.push([latestTotal]);
  }

  const BLOCK_WIDTH = 9; 
  const totalCols = maxBlocks * BLOCK_WIDTH;

  for (let i = 0; i < analyzedRows.length; i++) {
    const { blocks, endDate: rowEndDate } = analyzedRows[i]; 
    const rowData = [];
    const rowColors = [];

    for (let b = 0; b < maxBlocks; b++) {
      if (b < blocks.length) {
        const block = blocks[b];
        let currentBlockEnd;
        
        if (b < blocks.length - 1) {
          const nextStart = new Date(blocks[b + 1].startDate);
          currentBlockEnd = new Date(nextStart.setDate(nextStart.getDate() - 1));
        } else {
          currentBlockEnd = rowEndDate;
        }

        rowData.push(`${formatDate(block.startDate)}～${formatDate(currentBlockEnd)}`);
        rowColors.push(null); 
        
        let weeklyTotal = 0;
        for (let d = 0; d < 7; d++) {
          if (block.hours[d] > 0) {
            rowData.push(block.hours[d] + "h");
            weeklyTotal += block.hours[d];
          } else {
            rowData.push("休"); 
          }
          rowColors.push(null);
        }
        rowData.push(weeklyTotal > 0 ? parseFloat(weeklyTotal.toFixed(2)) + "h" : "0h");
        rowColors.push(null);
      } else {
        for (let k = 0; k < BLOCK_WIDTH; k++) {
          rowData.push("");
          rowColors.push("#f3f3f3"); 
        }
      }
    }
    outputData.push(rowData);
    backgroundColors.push(rowColors);
  }

  if (outputData.length > 0 && maxBlocks > 0) {
    const numRows = outputData.length;

    const requiredCols = startColNum + totalCols - 1;
    const currentMaxCols = sheet.getMaxColumns();
    if (currentMaxCols < requiredCols) {
      sheet.insertColumnsAfter(currentMaxCols, requiredCols - currentMaxCols);
    }

    const requiredRows = 1 + numRows;
    const currentMaxRows = sheet.getMaxRows();
    if (currentMaxRows < requiredRows) {
      sheet.insertRowsAfter(currentMaxRows, requiredRows - currentMaxRows);
    }
    
    if (idxWeeklyHours > -1) {
      sheet.getRange(2, idxWeeklyHours + 1, numRows, 1).setValues(weeklyHoursOutput);
    }

    const targetRange = sheet.getRange(2, startColNum, numRows, totalCols);
    const headerRange = sheet.getRange(1, startColNum, 1, totalCols);

    targetRange.clearDataValidations();
    headerRange.clearDataValidations();

    targetRange.setValues(outputData);
    targetRange.setBackgrounds(backgroundColors);

    const headerRow = [];
    for (let b = 0; b < maxBlocks; b++) {
      let label = (b === 0) ? "当初シフト" : `変更シフト${b}`;
      let suffix = (b === 0) ? "" : (b === 1 ? "(B)" : `(${String.fromCharCode(66 + b)})`);
      headerRow.push(`${label}(期間)`);
      ["月", "火", "水", "木", "金", "土", "日"].forEach(d => headerRow.push(`${d}${suffix}`));
      headerRow.push("合計");
    }
    headerRange.setValues([headerRow]);
  }

  // ★★★ 今回の追加機能：末尾の余分な空欄行をすべて削除する ★★★
  const finalMaxRows = sheet.getMaxRows();
  const finalLastRow = sheet.getLastRow();
  if (finalMaxRows > finalLastRow) {
    // 最終行の次の行から、余分な行数分を一気に削除します
    sheet.deleteRows(finalLastRow + 1, finalMaxRows - finalLastRow);
  }
  // ★★★ ここまで ★★★

  console.log(`[${type}] ${targetYear}年度 マスタ整備完了`);
}