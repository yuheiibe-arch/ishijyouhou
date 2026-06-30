/**
 * ====================================================================
 * 【ファイル: J_ContractUpdater.gs】
 * 定期非常勤医師の契約（勤務備考）を一括で次期に更新・追記するツール
 * ====================================================================
 */

function cu_showContractUpdater() {
  const html = HtmlService.createHtmlOutputFromFile('ContractUpdate_UI')
      .setWidth(900)
      .setHeight(650)
      .setTitle('定期非常勤 契約一括更新ツール');
  SpreadsheetApp.getUi().showModalDialog(html, '定期非常勤 契約一括更新ツール');
}

function cu_getContractData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ptSheet = ss.getSheets().find(s => /^定期非常勤\d{4}年度$/.test(s.getName()));
  
  if (!ptSheet) throw new Error("「定期非常勤〇〇年度」のシートが見つかりません。");

  const fyMatch = ptSheet.getName().match(/(\d{4})年度/);
  const fiscalYear = fyMatch ? parseInt(fyMatch[1], 10) : new Date().getFullYear();

  const data = ptSheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('医師名');
  const noteCol = headers.findIndex(h => String(h).trim() === "勤務備考" || String(h).trim() === "提案シフト");
  const breakNoteCol = headers.findIndex(h => String(h).trim() === "勤務備考（休憩あり）"); 
  const retireCol = headers.indexOf('退職日');

  if (nameCol === -1 || noteCol === -1) throw new Error("必要なカラムが見つかりません。");

  const today = new Date();
  today.setHours(0,0,0,0);
  const results = [];

  for (let i = 1; i < data.length; i++) {
    const name = data[i][nameCol];
    const currentText = data[i][noteCol] ? String(data[i][noteCol]).trim() : "";
    const currentBreakText = (breakNoteCol !== -1 && data[i][breakNoteCol]) ? String(data[i][breakNoteCol]).trim() : "";
    const retireDate = data[i][retireCol];

    if (!name || name === "★雛形") continue;
    if (!currentText) continue;

    // すでに退職済みの人（過去の日付）はリストから完全に除外
    const isPastRetire = (retireDate instanceof Date && !isNaN(retireDate.getTime()) && retireDate < today);
    if (isPastRetire) continue;

    // 退職日の文字列化
    let retireDateStr = "";
    if (retireDate instanceof Date && !isNaN(retireDate.getTime())) {
      retireDateStr = `${retireDate.getFullYear()}/${String(retireDate.getMonth() + 1).padStart(2, '0')}/${String(retireDate.getDate()).padStart(2, '0')}`;
    } else if (String(retireDate).trim() !== "") {
      retireDateStr = String(retireDate).trim();
    }

    const hasRetireDate = retireDateStr !== "";

    // 退職予定がある場合は、新しいブロックを生成しない（今のまま維持）
    const proposedText = hasRetireDate ? currentText : _cu_generateNextContractBlock(currentText, fiscalYear);
    
    let proposedBreakText = "";
    if (currentBreakText) {
      proposedBreakText = hasRetireDate ? currentBreakText : _cu_generateNextContractBlock(currentBreakText, fiscalYear);
    } else {
      proposedBreakText = proposedText;
    }

    results.push({
      name: name,
      current: currentText,
      proposed: proposedText,
      proposedBreak: proposedBreakText,
      checked: !hasRetireDate,       // 退職予定者は初期でチェックを外す
      disabled: hasRetireDate,       // UIで選択不可にするフラグ
      retireDateStr: retireDateStr   // UI表示用の退職日文字列
    });
  }
  return { sheetName: ptSheet.getName(), doctors: results };
}

function cu_saveContracts(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(payload.sheetName);
  if (!sheet) throw new Error(`シート「${payload.sheetName}」が見つかりません。`);

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameCol = headers.indexOf('医師名');
  const noteCol = headers.findIndex(h => String(h).trim() === "勤務備考" || String(h).trim() === "提案シフト");
  const breakNoteCol = headers.findIndex(h => String(h).trim() === "勤務備考（休憩あり）");

  const outRange = sheet.getRange(2, noteCol + 1, data.length - 1, 1);
  const outValues = outRange.getValues();
  
  let outRangeBreak, outValuesBreak;
  if (breakNoteCol !== -1) {
    outRangeBreak = sheet.getRange(2, breakNoteCol + 1, data.length - 1, 1);
    outValuesBreak = outRangeBreak.getValues();
  }
  
  const updateMap = new Map();
  payload.updates.forEach(u => {
    if (u.checked && !u.disabled) updateMap.set(u.name, u); 
  });

  let count = 0;
  for (let i = 1; i < data.length; i++) {
    const name = data[i][nameCol];
    if (updateMap.has(name)) {
      const u = updateMap.get(name);
      
      outValues[i - 1][0] = u.proposed;
      
      if (breakNoteCol !== -1 && outValuesBreak) {
        outValuesBreak[i - 1][0] = u.proposedBreak;
      }
      count++;
    }
  }

  outRange.setValues(outValues);
  if (breakNoteCol !== -1 && outRangeBreak) {
    outRangeBreak.setValues(outValuesBreak);
  }

  sheet.autoResizeRows(2, data.length - 1);

  ss.toast(`${count} 名の契約更新が完了しました！`, "更新完了", 5);
  return count;
}

function _cu_generateNextContractBlock(text, fiscalYear) {
  if (!text) return "";

  let cleanText = text.replace(/^"|"$/g, '').trim();

  const dateHeaderRegex = /([0-9０-９]{4})\s*[\/\-／年]\s*([0-9０-９]{1,2})\s*[\/\-／月]\s*([0-9０-９]{1,2})\s*日?\s*[～~〜\-−ー—–]\s*([0-9０-９]{4})\s*[\/\-／年]\s*([0-9０-９]{1,2})\s*[\/\-／月]\s*([0-9０-９]{1,2})\s*日?/g;
  
  let matches = [];
  let match;
  while ((match = dateHeaderRegex.exec(cleanText)) !== null) {
    matches.push({ 
      fullText: match[0], 
      index: match.index,
      endY: match[4],
      endM: match[5],
      endD: match[6]
    });
  }

  if (matches.length === 0) return cleanText;

  const lastMatch = matches[matches.length - 1];
  const lastBlockText = cleanText.substring(lastMatch.index);

  const toHalf = (str) => str.replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0));
  
  const y = parseInt(toHalf(lastMatch.endY), 10);
  const m = parseInt(toHalf(lastMatch.endM), 10);
  const d = parseInt(toHalf(lastMatch.endD), 10);

  const nextStartDate = new Date(y, m - 1, d + 1);
  const fiscalEndDate = new Date(fiscalYear + 1, 2, 31); 

  if (nextStartDate > fiscalEndDate) {
    return cleanText;
  }

  const sy = nextStartDate.getFullYear();
  const sm = nextStartDate.getMonth() + 1;

  let nextEndDate;
  if (sm >= 4 && sm <= 9) {
    nextEndDate = new Date(sy, 8, 30);
  } else {
    const endY = (sm >= 10) ? sy + 1 : sy;
    nextEndDate = new Date(endY, 2, 31);
  }

  if (nextEndDate > fiscalEndDate) {
    nextEndDate = fiscalEndDate;
  }

  const fmt = (dt) => `${dt.getFullYear()}/${String(dt.getMonth()+1).padStart(2,'0')}/${String(dt.getDate()).padStart(2,'0')}`;
  const nextPeriodStr = `${fmt(nextStartDate)}～${fmt(nextEndDate)}`;

  const remainingText = lastBlockText.substring(lastMatch.fullText.length).replace(/^[ \t\r\n]*/, '');
  const newBlockText = nextPeriodStr + "\n" + remainingText;

  return cleanText + "\n\n" + newBlockText.trim();
}