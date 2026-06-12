// ==========================================
// 【MasterSync】マスタへの転記・シフト展開・計算 (完全動的・医師免許対応版)
// ==========================================

function syncAndExpandSchedule(headers, rowData, contractText, fiscalYear, contractWithBreakText) {
  if (!fiscalYear || !contractText) {
    console.warn("年度または契約テキストがないため、マスタ転記をスキップしました。");
    return;
  }

  // 1. 転記先シートの決定
  const typeIdx = headers.findIndex(h => String(h).includes("採用区分"));
  const type = (typeIdx > -1) ? rowData[typeIdx] : "";
  
  const yearNum = String(fiscalYear).replace("年度", "");
  const sheetYearName = `${yearNum}年度`;

  let targetSheetName = "";
  if (type === "常勤") {
    targetSheetName = `常勤${sheetYearName}`;
  } else if (type === "定期非常勤") {
    targetSheetName = `定期非常勤${sheetYearName}`;
  } else {
    console.warn(`採用区分「${type}」に対応するシート定義がありません。`);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetSheet = ss.getSheetByName(targetSheetName);

  if (!targetSheet) {
    console.error(`エラー: 転記先シート「${targetSheetName}」が見つかりません。`);
    return;
  }

  const lastCol = targetSheet.getLastColumn();
  const targetHeaders = targetSheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // 3. 基本データの準備（見出し名で動的取得）
  const medId = getValueByHeader(headers, rowData, "医籍番号");
  const name = getValueByHeader(headers, rowData, "氏名") || getValueByHeader(headers, rowData, "医師名");
  const joinDate = getValueByHeader(headers, rowData, "入職日");
  const specialty = getValueByHeader(headers, rowData, "専門科") || getValueByHeader(headers, rowData, "専門");
  
  // ★追加：医師免許取得日の動的取得
  const licenseDate = getValueByHeader(headers, rowData, "医師免許取得日") || getValueByHeader(headers, rowData, "医師免許取得");

  // ★追加：jinjer番号とシメイの取得・整形
  const jinjerId = getValueByHeader(headers, rowData, "jinjer番号") || "";
  const kanaRaw = getValueByHeader(headers, rowData, "シメイ") || getValueByHeader(headers, rowData, "フリガナ") || getValueByHeader(headers, rowData, "カナ") || "";
  const cleanKana = String(kanaRaw).replace(/[\s　\/／]/g, "").trim();
  
  // 祝日・年末年始
  const holidayHeader = type === "常勤" ? "【常勤】 祝日" : "【定期非常勤】 祝日";
  const newYearHeader = type === "常勤" ? "【常勤】 年末年始" : "【定期非常勤】 年末年始";
  
  let rawHoliday = getValueByHeader(headers, rowData, holidayHeader);
  if (!rawHoliday) rawHoliday = getValueByHeader(headers, rowData, "祝日");
  
  let rawNewYear = getValueByHeader(headers, rowData, newYearHeader);
  if (!rawNewYear) rawNewYear = getValueByHeader(headers, rowData, "年末年始");
  
  const holidaySimple = convertToYesNoSimple(rawHoliday);
  const newYearSimple = convertToYesNoSimple(rawNewYear);

  // 時給の取得ロジック（あいまい検索対応）
  let wage = "";
  if (type === "定期非常勤") {
    wage = getValueByHeader(headers, rowData, "契約時給");
    if (!wage) {
      const wIdx = headers.findIndex(h => String(h).includes("定期非常勤") && String(h).includes("契約時給"));
      if (wIdx > -1) wage = rowData[wIdx];
    }
    if (!wage) {
      const wIdx = headers.findIndex(h => String(h).includes("契約時給"));
      if (wIdx > -1) wage = rowData[wIdx];
    }
  }

  // 解析と計算
  const parsedBlocks = parseRemarksRobust(contractText);
  const calcResult = calculateMainDutyAndWeeklyHours(parsedBlocks, contractText);
  const mainDuty = calcResult.mainDuty;   
  const weeklyHours = calcResult.weeklyHours; 

  // 転記用データマップ（シートの列名と紐付け）
  const dataMap = {
    "医籍番号": medId,
    "jinjer番号": jinjerId,         // ★追加
    "医師名": name,
    "氏名": name,
    "シメイ": cleanKana,            // ★追加
    "入職日": joinDate,
    "専門": specialty,      
    "専門科": specialty,    
    "医師免許取得": licenseDate, // ★追加
    "医師免許取得日": licenseDate, // ★追加
    "主務": mainDuty,
    "祝日": holidaySimple,      
    "年末年始": newYearSimple,  
    "週労働": weeklyHours,
    "勤務備考": contractText,   
    "勤務備考（休憩あり）": contractWithBreakText, // ★追加
    "時給": wage,
    "契約時給": wage
  };

  // 5. 新規行の追加と番号採番
  const lastRow = targetSheet.getLastRow();
  let targetRowIndex = lastRow + 1; 
  const noColIdx = targetHeaders.indexOf("番号");
  if (noColIdx > -1) {
    const newNo = getNextNumber(targetSheet); 
    targetSheet.getRange(targetRowIndex, noColIdx + 1).setValue(newNo);
  }

  // 6. 行データの書き込み
  for (let i = 0; i < lastCol; i++) {
    const headerName = String(targetHeaders[i]).trim();
    if (headerName.includes("シフト") || headerName === "合計" || headerName === "月") continue;

    if (dataMap.hasOwnProperty(headerName) && dataMap[headerName] !== undefined && dataMap[headerName] !== "") {
      targetSheet.getRange(targetRowIndex, i + 1).setValue(dataMap[headerName]);
    }
  }

  // 7. シフト表展開
  const shiftStartIdx = targetHeaders.indexOf("当初シフト(期間)");
  if (shiftStartIdx > -1) {
    writeShiftScheduleRow_Dynamic(targetSheet, targetRowIndex, parsedBlocks, shiftStartIdx + 1);
  } else {
    console.warn("「当初シフト(期間)」列が見つかりません。");
  }

  console.log(`[${type}] ${name}様 を「${targetSheetName}」の行${targetRowIndex}に追加しました。`);
}

function convertToYesNoSimple(text) {
  if (!text) return "無"; 
  const t = String(text);
  if (t.includes("なし") || t.includes("休") || t.includes("無")) return "無";
  if (t.includes("あり") || t.includes("勤") || t.includes("有")) return "有";
  return t; 
}

function calculateMainDutyAndWeeklyHours(blocks, text) {
  if (!blocks || blocks.length === 0) {
    return { mainDuty: "", weeklyHours: "" };
  }
  const firstBlock = blocks[0];
  let totalHours = 0;
  firstBlock.hours.forEach(h => totalHours += h);
  const weeklyStr = totalHours > 0 ? `${parseFloat(totalHours.toFixed(2))}h` : "";

  let duty = "";
  const match = text.match(/【([^】]+)】/);
  if (match) duty = match[1]; 

  return { mainDuty: duty, weeklyHours: weeklyStr };
}

function writeShiftScheduleRow_Dynamic(sheet, row, blocks, startCol) {
  if (!blocks || blocks.length === 0) return;

  const rowData = [];
  const fyEnd = new Date("2026/03/31"); 

  blocks.forEach((block, index) => {
    let endDate;
    if (index < blocks.length - 1) {
      const nextStart = new Date(blocks[index + 1].startDate);
      endDate = new Date(nextStart.setDate(nextStart.getDate() - 1));
    } else {
      endDate = fyEnd;
    }

    rowData.push(`${formatDate(block.startDate)}～${formatDate(endDate)}`);
    
    let subTotal = 0;
    block.hours.forEach(h => {
      if (h > 0) {
        rowData.push(h + "h");
        subTotal += h;
      } else {
        rowData.push("休");
      }
    });
    rowData.push(subTotal > 0 ? parseFloat(subTotal.toFixed(2)) + "h" : "0h");
  });

  if (rowData.length > 0) {
    sheet.getRange(row, startCol, 1, rowData.length).setValues([rowData]);
  }
}

function formatDate(date) {
  if (!date) return "";
  return Utilities.formatDate(date, "JST", "yyyy/MM/dd");
}

function getNextNumber(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const max = values.reduce((a, b) => Math.max(Number(a) || 0, Number(b) || 0), 0);
  return max + 1;
}

/**
 * ==========================================
 * 外部シートからの自動更新リクエストを受け取るWeb API
 * ==========================================
 */
function doGet(e) {
  try {
    // 外部シートから指定された「年度」を受け取る
    let targetYear = e.parameter.year;
    
    // 指定がない場合は、シート名から最新年度を自動判別
    if (!targetYear) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let maxYear = 0;
      ss.getSheets().forEach(s => {
        const match = s.getName().match(/^(\d{4})/);
        if (match) {
          const y = parseInt(match[1], 10);
          if (y > maxYear) maxYear = y;
        }
      });
      targetYear = String(maxYear);
    }

    if (targetYear !== "0") {
      // 私たちが作った「スマート更新（差分チェック）」を裏でスタートさせる！
      checkAndStartSmartSync(targetYear, "通年");
    }
    
    return ContentService.createTextOutput("Success: SmartSync Started for " + targetYear);
  } catch (error) {
    return ContentService.createTextOutput("Error: " + error.message);
  }
}