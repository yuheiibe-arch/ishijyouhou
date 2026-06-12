// ==========================================
// 【Utils】共通ヘルパー関数群 (新旧全スクリプトで共有)
// ==========================================

/**
 * 備考欄解析（完全版）
 * 日付、時間(MAX集計)、曜日を解析してブロック配列を返す
 */
function parseRemarksRobust(text) {
  if (!text) return [];
  text = String(text).replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = text.split("\n");
  const blocks = [];
   
  let currentBlock = { startDate: new Date("2025/04/01"), hours: [0,0,0,0,0,0,0], hasHours: false };
  let isDateFound = false;

  lines.forEach(line => {
    line = line.trim();
    if (!line) return;

    // 日付検出 (2025/06/01, 2025年6月 etc)
    let dateMatch = line.match(/(\d{4})[\/\.\-](\d{1,2})[\/\.\-](\d{1,2})/);
    if (!dateMatch) {
      const jpDateMatch = line.match(/(\d{4})\s*年\s*(\d{1,2})\s*月(?:\s*(\d{1,2})\s*日)?/);
      if (jpDateMatch) {
        const day = jpDateMatch[3] ? jpDateMatch[3] : "1";
        dateMatch = [jpDateMatch[0], jpDateMatch[1], jpDateMatch[2], day];
      }
    }

    if (dateMatch) {
      if (isDateFound || currentBlock.hasHours) {
        blocks.push(currentBlock);
      }
      const newDate = new Date(`${dateMatch[1]}/${dateMatch[2]}/${dateMatch[3]}`);
      currentBlock = { startDate: newDate, hours: [0,0,0,0,0,0,0], hasHours: false };
      isDateFound = true;
      return; 
    }

    // 時間検出 (18:00-21:00, 18-21時 etc)
    let startH, startM, endH, endM;
    let validTime = false;
    const timeRegexColon = /(\d{1,2})[:：](\d{2})\s*[～~〜\-\u2010\u2011\u2013\u2014\u2212\uFF0D]\s*(\d{1,2})[:：](\d{2})/;
    const matchColon = line.match(timeRegexColon);

    if (matchColon) {
      startH = parseInt(matchColon[1], 10); startM = parseInt(matchColon[2], 10);
      endH = parseInt(matchColon[3], 10); endM = parseInt(matchColon[4], 10);
      validTime = true;
    } else {
      const timeRegexHour = /(\d{1,2})\s*[～~〜\-\u2010\u2011\u2013\u2014\u2212\uFF0D]\s*(\d{1,2})\s*時/;
      const matchHour = line.match(timeRegexHour);
      if (matchHour) {
        startH = parseInt(matchHour[1], 10); startM = 0;
        endH = parseInt(matchHour[2], 10); endM = 0;
        validTime = true;
      }
    }
    
    if (!validTime) return;

    const workMinutes = calculateNetWorkMinutesStrict(startH, startM, endH, endM);
    const workHours = parseFloat((workMinutes / 60).toFixed(2));
    if (workHours <= 0) return;

    // 曜日検出
    const targetDays = [];
    if (checkDayContext(line, "月")) targetDays.push(0);
    if (checkDayContext(line, "火")) targetDays.push(1);
    if (checkDayContext(line, "水")) targetDays.push(2);
    if (checkDayContext(line, "木")) targetDays.push(3);
    if (checkDayContext(line, "金")) targetDays.push(4);
    if (checkDayContext(line, "土")) targetDays.push(5);

    let sundayCheckLine = line.replace(/月曜日|火曜日|水曜日|木曜日|金曜日|土曜日/g, "").replace(/月曜|火曜|水曜|木曜|金曜|土曜/g, "");
    if (checkDayContext(sundayCheckLine, "日")) targetDays.push(6);

    // MAX集計
    [...new Set(targetDays)].forEach(dayIndex => {
      currentBlock.hours[dayIndex] = Math.max(currentBlock.hours[dayIndex], workHours);
    });
    currentBlock.hasHours = true;
  });
   
  if (currentBlock.hasHours) blocks.push(currentBlock);
  return blocks.filter(b => b.hasHours);
}

// 列検索
function findTargetIndex(headers, keywords, isFull, is2nd, isPartTime) {
  const candidates = [];
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i]);
    if (!keywords.every(k => h.includes(k))) continue;
    const has2ndTag = h.includes("第2期");
    if ((is2nd && !has2ndTag) || (!is2nd && has2ndTag)) continue;
    const hasPartTimeTag = h.includes("定期非常勤");
    const isShiftRelated = keywords.some(k => k.includes("勤務") || k.includes("拠点") || k.includes("曜日"));
    if (isShiftRelated) {
        if ((isPartTime && !hasPartTimeTag) || (!isPartTime && hasPartTimeTag)) continue;
    }
    candidates.push(i);
  }
  return isFull ? candidates[candidates.length - 1] : candidates[0];
}

// 休憩計算（既存：一律2時間引くロジック ※勤怠表では使用しませんが残します）
function calculateNetWorkMinutesStrict(startH, startM, endH, endM) {
  const start = startH * 60 + startM;
  let end = endH * 60 + endM;
  if (end < start) end += 24 * 60;
  let duration = end - start;
  const breakStart = 13 * 60; 
  const breakEnd = 15 * 60;    
  if (start <= breakStart && end >= breakEnd) duration -= 120;
  return duration;
}

// その他ヘルパー
function getValueByHeader(headers, rowData, keyword) {
  const index = headers.findIndex(h => String(h).includes(keyword));
  return (index > -1) ? rowData[index] : "";
}
function removePrefix(s) {
  if (!s) return "";
  const str = String(s);
  const cleaned = str.replace(/^【[^】]+】/, '');
  return (cleaned === "") ? str.replace(/【|】/g, '') : cleaned;
}
function unifyWording(rawVal) {
  if (!rawVal) return "ー";
  const s = String(rawVal).trim();
  if (s === "無" || s === "なし" || s === "勤務なし" || s === "無し") return "勤務なし";
  if (s === "有" || s === "あり" || s === "勤務あり" || s === "有り") return "勤務あり";
  return s;
}
function getDayBefore(dVal) {
  if (!dVal) return "";
  const d = new Date(dVal);
  if (isNaN(d.getTime())) return "";
  d.setTime(d.getTime() - (24*60*60*1000));
  return formatDateJp(d);
}
function formatDateJp(d) {
  if (!d) return "";
  const dt = new Date(d);
  if (isNaN(dt.getTime())) return String(d);
  return `${dt.getFullYear()}/${dt.getMonth()+1}/${dt.getDate()}`;
}
function formatDate(date) { return Utilities.formatDate(date, "JST", "yyyy/MM/dd"); }
function checkDayContext(line, char) { return line.includes(`${char}曜`) || line.includes(`(${char})`) || line.includes(`${char}:`) || line.includes(`${char}　`) || line.includes(`${char} `); }
function getNextNumber(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  const lastVal = sheet.getRange(lastRow, 1).getValue();
  return (typeof lastVal === 'number') ? lastVal + 1 : 1;
}
function filterBlocksByYearAndEntry(blocks, entryDate, fyEnd) {
  const result = [];
  blocks.sort((a, b) => a.startDate - b.startDate);
  for (let i = 0; i < blocks.length; i++) {
    let block = blocks[i];
    if (block.startDate > fyEnd) continue;
    let nextBlockStart = (i + 1 < blocks.length) ? blocks[i+1].startDate : new Date("2100/01/01");
    if (nextBlockStart <= entryDate) continue;
    if (block.startDate < entryDate) {
      let adjustedBlock = JSON.parse(JSON.stringify(block));
      adjustedBlock.startDate = new Date(entryDate);
      result.push(adjustedBlock);
    } else {
      result.push(block);
    }
  }
  return result;
}

// =======================================================
// ★以下、勤怠自動入力用に追加した共通ヘルパー関数群★
// =======================================================

function standardizeRemarksFormat(text) {
  if (!text) return "";
  let s = String(text);
  s = s.replace(/　/g, " ");
  s = s.replace(/(\d{1,2}[:：]\d{2})\s*[~~\-ｰ\u2010\u2011\u2013\u2014\u2212]\s*(\d{1,2}[:：]\d{2})/g, "$1～$2");
  s = s.replace(/(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})\s*[~~\-ｰ\u2010\u2011\u2013\u2014\u2212]\s*(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})/g, "$1～$2");
  s = s.replace(/(曜日)\s+(\d)/g, "$1：$2");
  s = s.replace(/(曜日)(\d{1,2}[:：])/g, "$1：$2");
  return s;
}

function parseRemarksToTextBlocks(text) {
  if (!text) return [];
  const lines = text.split("\n");
  const blocks = [];
  const FY_START = new Date("2025/04/01");
  const FY_END = new Date("2026/03/31");
  let currentBlock = { startDate: null, endDate: null, lines: [] };
  let hasDate = false;

  lines.forEach(line => {
    line = line.trim();
    if (!line) return;
    const dateMatch = line.match(/(\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2})/);
    if (dateMatch) {
      if (currentBlock.lines.length > 0 || hasDate) {
        blocks.push(currentBlock);
      }
      const parts = line.split("～");
      let sDate = new Date(parts[0]);
      let eDate = parts[1] ? new Date(parts[1]) : FY_END;
      currentBlock = { startDate: sDate, endDate: eDate, lines: [] };
      hasDate = true;
    } else {
      if (!line.startsWith("（祝日")) {
        currentBlock.lines.push(line);
      }
    }
  });
  if (currentBlock.lines.length > 0 || hasDate) blocks.push(currentBlock);
  if (blocks.length === 0 && currentBlock.lines.length > 0) {
      currentBlock.startDate = FY_START;
      currentBlock.endDate = FY_END;
      blocks.push(currentBlock);
  }
  return blocks;
}

// ★修正版：1日複数シフトの抽出対応
function getShiftForDay(date, shiftLines, isHoliday, hRule, isNewYear, nRule) {
  if (isHoliday) {
    if (hRule.match(/(無|なし|×|不可|勤務なし)/)) return "";
  } 
  else if (isNewYear) {
    if (nRule.match(/(無|なし|×|不可|勤務なし)/)) return ""; 
  }
  
  const dayOfWeek = date.getDay();
  const weekNum = Math.floor((date.getDate() - 1) / 7) + 1;
  const jpDays = ["日", "月", "火", "水", "木", "金", "土"];
  const targetDayStr = jpDays[dayOfWeek]; 

  let matchedShifts = []; // 複数シフトを格納する配列

  for (let line of shiftLines) {
    if (line.indexOf(targetDayStr + "曜日") === -1 && line.indexOf(targetDayStr + "曜") === -1) continue;
    let isMatch = false;
    
    if (line.includes("毎週")) {
      isMatch = true;
    } else if (line.match(/第[\d・･]+/)) {
      const match = line.match(/第([\d・･]+)/);
      if (match) {
        const weeks = match[1].split(/[・･]/).map(Number); 
        if (weeks.includes(weekNum)) isMatch = true;
      }
    }
    
    // 条件に合致すれば配列に追加（即終了しない）
    if (isMatch) {
      const formatted = formatShiftLine(line);
      if (formatted) matchedShifts.push(formatted);
    }
  }
  
  // 見つかったすべてのシフトを改行で繋いで返す（1つならそのまま、0なら空文字）
  return matchedShifts.join('\n'); 
}

function formatShiftLine(line) {
  const locMatch = line.match(/(【.*?】)/);
  const location = locMatch ? locMatch[1] : "";
  const timeMatch = line.match(/(\d{1,2})[:：](\d{2})\s*～\s*(\d{1,2})[:：](\d{2})/);
  if (timeMatch) {
    let startH = timeMatch[1];
    let startM = timeMatch[2];
    let endH = timeMatch[3];
    let endM = timeMatch[4];
    let startStr = (startM === "00") ? startH : `${startH}:${startM}`;
    let endStr = (endM === "00") ? endH : `${endH}:${endM}`;
    return `${location}${startStr}-${endStr}`;
  }
  return "";
}

// ★修正版：複数シフトの労働時間を合算する対応
function calculateWorkHours_AbsoluteBreak(shiftText) {
  if (!shiftText) return 0;
  
  // 改行区切りで複数シフトが来る可能性があるため分割
  const shiftLines = shiftText.split('\n');
  let totalHours = 0;

  for (let shift of shiftLines) {
    const match = shift.match(/(\d{1,2})(:(\d{2}))?-(\d{1,2})(:(\d{2}))?/);
    if (!match) continue;
    
    const startH = parseInt(match[1], 10);
    const startM = match[3] ? parseInt(match[3], 10) : 0;
    const startMin = startH * 60 + startM;
    
    const endH = parseInt(match[4], 10);
    const endM = match[6] ? parseInt(match[6], 10) : 0;
    let endMin = endH * 60 + endM;
    
    if (endMin < startMin) endMin += 24 * 60;
    
    const duration = (endMin - startMin);
    const breakStartMin = 13 * 60; 
    const breakEndMin = 15 * 60;   
    const overlapStart = Math.max(startMin, breakStartMin);
    const overlapEnd = Math.min(endMin, breakEndMin);
    
    let breakDuration = 0;
    if (overlapEnd > overlapStart) {
      breakDuration = overlapEnd - overlapStart;
    }
    
    const netWorkMinutes = duration - breakDuration;
    totalHours += Math.max(0, netWorkMinutes / 60);
  }
  
  return totalHours;
}