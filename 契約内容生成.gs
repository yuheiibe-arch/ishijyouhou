// ==========================================
// 【ContractGenerator】契約テキスト生成ロジック
// ==========================================

function generateTextForRow(headers, rowData) {
  // 1. 基礎情報の取得
  const typeIdx = headers.findIndex(h => String(h).includes("採用区分"));
  const joinDateIdx = headers.findIndex(h => String(h).includes("入職日"));
  const type = (typeIdx > -1) ? rowData[typeIdx] : "";
  const joinDateVal = (joinDateIdx > -1) ? rowData[joinDateIdx] : "";
    
  let lines = [];
  let linesWithBreak = []; // ★休憩ありテキスト用
  
  // --- 年度判定ロジック ---
  let fyVal = getValueByHeader(headers, rowData, "採用年度");
  let currentYear = 2025; 

  if (fyVal) {
    currentYear = parseInt(String(fyVal).replace("年度", ""));
  } else if (joinDateVal) {
    const d = new Date(joinDateVal);
    if (!isNaN(d.getTime())) {
      const y = d.getFullYear();
      const m = d.getMonth() + 1;
      const fyJoin = (m < 4) ? y - 1 : y;
      
      if (fyJoin >= 2025) {
        currentYear = fyJoin;
      } else {
        currentYear = 2025;
      }
    }
  }

  const fiscalYear = `${currentYear}年度`;
  const fyStartDate = new Date(currentYear, 3, 1); // 4月1日
  const fyEndDate = new Date(currentYear + 1, 2, 31); // 翌3月31日

  // 勤務備考（シフト情報）
  let remarksRaw = getValueByHeader(headers, rowData, "勤務備考");
  if (!remarksRaw) remarksRaw = "";

  // 採用区分によるヘッダー特定
  let holidayHeader = "祝日";
  let newYearHeader = "年末年始";
  
  if (type === "常勤") {
    holidayHeader = "【常勤】 祝日";
    newYearHeader = "【常勤】 年末年始";
  } else if (type === "定期非常勤") {
    holidayHeader = "【定期非常勤】 祝日";
    newYearHeader = "【定期非常勤】 年末年始";
  }

  // 値取得
  let holidayVal = getValueByHeader(headers, rowData, holidayHeader);
  if (!holidayVal) holidayVal = getValueByHeader(headers, rowData, "祝日"); 
  
  let newYearVal = getValueByHeader(headers, rowData, newYearHeader);
  if (!newYearVal) newYearVal = getValueByHeader(headers, rowData, "年末年始"); 

  // 表記統一（勤務あり/なし）
  holidayVal = unifyWording(holidayVal); 
  newYearVal = unifyWording(newYearVal);

  // --- 第2期 変更判定 ---
  const isFull = (type === '常勤');
  const isPartTime = !isFull; 
  const changeKW = isPartTime ? ["年度内の契約変更", "定期非常勤"] : ["年度内の契約変更"];
  const changeIdx = findTargetIndex(headers, changeKW, isFull, false, isPartTime);
  const changeVal = (changeIdx > -1) ? rowData[changeIdx] : "";
  const hasChange = String(changeVal).includes('ある');

  const start2Idx = findTargetIndex(headers, ["第2期", "契約開始日"], isFull, true, isPartTime);
  const end2Idx = findTargetIndex(headers, ["第2期", "契約終了日"], isFull, true, isPartTime);
  const start2Val = (start2Idx > -1) ? rowData[start2Idx] : "";
  const end2Val = (end2Idx > -1) ? rowData[end2Idx] : "";

  // 3. 契約テキストの構築
  const stdRemarks = typeof standardizeRemarksFormat === 'function' ? standardizeRemarksFormat(remarksRaw) : remarksRaw;
  const blocks = typeof parseRemarksToTextBlocks === 'function' ? parseRemarksToTextBlocks(stdRemarks) : [];

  // --- A. 第1期 生成 ---
  const autoTerm1 = buildShiftTextAuto(headers, rowData, isFull, false, isPartTime);
  
  if (joinDateVal || blocks.length > 0 || autoTerm1.normal) {
    let s1Date = null;
    
    if (joinDateVal) {
      s1Date = new Date(joinDateVal);
      if (s1Date < fyStartDate) {
        s1Date = fyStartDate;
      }
    } else {
      s1Date = fyStartDate;
    }

    let e1Str = "";
    if (hasChange && start2Val) {
      e1Str = getDayBefore(start2Val);
    } else {
      e1Str = formatDateJp(fyEndDate);
    }

    const pStr = `${formatDateJp(s1Date)}～${e1Str}`;
    lines.push(pStr);
    linesWithBreak.push(pStr);

    if (autoTerm1.normal) {
      lines.push(autoTerm1.normal);
      linesWithBreak.push(autoTerm1.withBreak);
    } else if (remarksRaw) {
      lines.push(remarksRaw);
      linesWithBreak.push(remarksRaw);
    }

    if (isFull) {
      const footer = `（祝日：${holidayVal}／年末年始：${newYearVal}）`;
      lines.push(footer);
      linesWithBreak.push(footer);
    }
  }

  // --- B. 第2期 生成 ---
  if (hasChange) {
    lines.push(""); 
    linesWithBreak.push("");
    
    if (start2Val) {
      const s2 = formatDateJp(start2Val);
      let e2 = end2Val ? formatDateJp(end2Val) : "";
      if (!e2) e2 = formatDateJp(fyEndDate);
      const p2Str = `${s2}～${e2}`;
      lines.push(p2Str);
      linesWithBreak.push(p2Str);
    } else {
      lines.push("（期間未定）");
      linesWithBreak.push("（期間未定）");
    }
    
    const autoTerm2 = buildShiftTextAuto(headers, rowData, isFull, true, isPartTime);
    if (autoTerm2.normal) {
      lines.push(autoTerm2.normal);
      linesWithBreak.push(autoTerm2.withBreak);
    }

    const h2Idx = findTargetIndex(headers, isPartTime ? ["第2期", "祝日"] : ["第2期", "祝日"], isFull, true, isPartTime);
    const y2Idx = findTargetIndex(headers, isPartTime ? ["第2期", "年末年始"] : ["第2期", "年末年始"], isFull, true, isPartTime);
    const h2Val = (h2Idx > -1) ? rowData[h2Idx] : "";
    const y2Val = (y2Idx > -1) ? rowData[y2Idx] : "";
    
    if (h2Val || y2Val) {
        const footer2 = `（祝日：${unifyWording(h2Val)}／年末年始：${unifyWording(y2Val)}）`;
        lines.push(footer2);
        linesWithBreak.push(footer2);
    }
  }
  
  if (isPartTime && !hasChange) {
      const footerPT = `（祝日：${holidayVal}／年末年始：${newYearVal}）`;
      lines.push(footerPT);
      linesWithBreak.push(footerPT);
  }
  
  return { 
    contract: lines.join('\n'), 
    contractWithBreak: linesWithBreak.join('\n'), 
    fy: fiscalYear 
  };
}

// ------------------------------------------
// ヘルパー関数
// ------------------------------------------

function buildShiftTextAuto(headers, rowData, isFull, is2nd, isPartTime) {
  const days = ['月', '火', '水', '木', '金', '土', '日'];
  let texts = [];
  let textsWithBreak = [];

  days.forEach(day => {
    const kwTime = ["勤務時刻", `[${day}]`]; 
    if (is2nd) kwTime.push("第2期");
    const tIdx = findTargetIndex(headers, kwTime, isFull, is2nd, isPartTime);

    const kwFreq = ["勤務曜日", `[${day}]`];
    if (is2nd) kwFreq.push("第2期");
    const fIdx = findTargetIndex(headers, kwFreq, isFull, is2nd, isPartTime);

    const dayKw = `${day}曜日`;
    const kwAM = [dayKw, "拠点", "午前"];
    const kwPM = [dayKw, "拠点", "午後"];
    if (is2nd) { kwAM.push("第2期"); kwPM.push("第2期"); }
    const amIdx = findTargetIndex(headers, kwAM, isFull, is2nd, isPartTime);
    const pmIdx = findTargetIndex(headers, kwPM, isFull, is2nd, isPartTime);

    if (tIdx > -1) {
      let time = rowData[tIdx];
      if (time && String(time) !== '休' && String(time) !== '') {
        let timeStr = String(time);
        let hasBreak = false;
        
        if (timeStr.includes("変則")) {
            timeStr = "変則";
        } else {
            timeStr = timeStr.replace(/[－ー〜～-]/g, '～');
            
            // 休憩時間の自動判定（13:00～15:00を跨ぐか）
            let [sStr, eStr] = timeStr.split('～');
            if (sStr && eStr) {
              let sHour = parseInt(sStr.split(':')[0], 10);
              let eHour = parseInt(eStr.split(':')[0], 10);
              if (!isNaN(sHour) && !isNaN(eHour) && sHour <= 13 && eHour >= 15) {
                hasBreak = true;
              }
            }
            timeStr = timeStr.replace(/^0/, '').replace(/～0/g, '～');
        }

        // 頻度の表記ブレ統一
        let freqStr = "毎週";
        if (fIdx > -1) {
          freqStr = normalizeFrequency(rowData[fIdx]);
        }
        
        let locAM = (amIdx > -1) ? rowData[amIdx] : "";
        let locPM = (pmIdx > -1) ? rowData[pmIdx] : "";
        const cAM = removePrefix(locAM).trim();
        const cPM = removePrefix(locPM).trim();
        
        let finalLocStr = "";
        const isAMHoliday = (!cAM || cAM === '休日' || cAM === '勤務なし' || cAM === '');
        const isPMValid = (cPM && cPM !== '同拠点' && cPM !== '休日' && cPM !== '勤務なし' && cPM !== '');

        if (isPMValid) {
          if (isAMHoliday) finalLocStr = cPM;
          else if (cAM !== cPM) finalLocStr = `${cAM}/${cPM}`;
          else finalLocStr = cAM;
        } else {
          finalLocStr = cAM;
        }
        const locDisplay = finalLocStr ? `【${finalLocStr}】` : "";
        const baseLine = `${locDisplay}${freqStr}${day}曜日：${timeStr}`;
        
        texts.push(baseLine);
        textsWithBreak.push(baseLine + (hasBreak ? "　(休憩13:00～15:00)" : ""));
      }
    }
  });
  
  return {
    normal: texts.join('\n'),
    withBreak: textsWithBreak.join('\n')
  };
}

function normalizeFrequency(rawFreq) {
  if (!rawFreq) return "毎週";
  let str = String(rawFreq).trim();
  if (str.includes("毎週")) return "毎週";
  
  let nums = str.match(/[1-5１-５]/g);
  if (nums) {
    let normalizedNums = [...new Set(nums.map(v => parseInt(v.replace(/[１-５]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)), 10)))].sort((a,b) => a-b);
    return `第${normalizedNums.join('・')}`;
  }
  return str;
}

function findTargetIndex(headers, keywords, isFull, is2nd, isPartTime) {
  return headers.findIndex(header => {
    const h = String(header);
    const allMatch = keywords.every(kw => h.includes(kw));
    if (!allMatch) return false;
    if (!is2nd && h.includes("第2期")) return false;
    if (isFull && h.includes("定期非常勤")) return false;
    if (isPartTime && h.includes("常勤") && !h.includes("定期非常勤")) return false; 
    return true;
  });
}

function formatDateJp(dateVal) {
  if (!dateVal) return "";
  const d = new Date(dateVal);
  if (isNaN(d.getTime())) return String(dateVal);
  return `${d.getFullYear()}/${d.getMonth() + 1}/${d.getDate()}`;
}

function getDayBefore(dateVal) {
  if (!dateVal) return "";
  const d = new Date(dateVal);
  d.setDate(d.getDate() - 1);
  return formatDateJp(d);
}

function removePrefix(text) {
  if (!text) return "";
  return String(text).replace(/^【.*?】/, "");
}

function unifyWording(text) {
  if (!text) return "勤務なし";
  const t = String(text);
  if (t.includes("なし") || t.includes("休") || t.includes("無")) return "勤務なし";
  if (t.includes("あり") || t.includes("勤") || t.includes("有")) return "勤務あり";
  return t;
}

function getValueByHeader(headers, rowData, headerName) {
  const idx = headers.indexOf(headerName);
  if (idx === -1) return "";
  return rowData[idx];
}