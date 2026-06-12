/**
 * 【本番用】契約情報シートからの自動転記スクリプト
 * 契約情報シートの最新回答を読み取り、常勤・定期非常勤シートへ自動で振り分けて転記します。
 */
function transferContractToMaster() {
  const targetSs = SpreadsheetApp.getActiveSpreadsheet(); // 転記先（現在のファイル）
  
  // ========================================================
  // 1. 外部マスタ（契約情報）の取得
  // ========================================================
  const extUrl = 'https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit';
  let extSs;
  try { extSs = SpreadsheetApp.openByUrl(extUrl); } 
  catch(e) { console.error("❌ 外部マスタが開けません。アクセス権限を確認してください。"); return; }

  const sourceSheet = extSs.getSheetByName('契約情報');
  if (!sourceSheet) { console.error("❌ 契約情報シートが見つかりません。"); return; }

  const sourceLastRow = sourceSheet.getLastRow();
  const sourceLastCol = sourceSheet.getLastColumn();
  if (sourceLastRow < 2) {
    console.error("⚠️ 転記するデータ（回答）がありません。");
    return;
  }

  const sourceHeaders = sourceSheet.getRange(1, 1, 1, sourceLastCol).getValues()[0].map(h => String(h).trim());
  const rawForm = sourceSheet.getRange(sourceLastRow, 1, 1, sourceLastCol).getValues()[0];

  const getVal = (colName) => {
    const idx = sourceHeaders.findIndex(h => h.includes(colName));
    return idx !== -1 ? rawForm[idx] : "";
  };

  // ========================================================
  // 2. 基本情報の取得と転記先シートの判定
  // ========================================================
  const doctorName = String(getVal('医師名') || getVal('氏名') || rawForm[3] || "").replace(/先生$/, "").trim();
  if (!doctorName) {
    console.error("❌ 医師名が取得できませんでした。処理を終了します。");
    return;
  }

  const employmentType = String(getVal('主務') || getVal('雇用形態') || rawForm[7] || "").trim();
  const isJoukin = employmentType.includes("常勤") && !employmentType.includes("非常勤");
  const targetSheetName = isJoukin ? '常勤2026年度' : '定期非常勤2026年度';
  const targetSheet = targetSs.getSheetByName(targetSheetName);
  
  if (!targetSheet) { 
    console.error(`❌ 転記先シート「${targetSheetName}」が見つかりません。`); 
    return; 
  }

  const medId = String(getVal('医籍番号') || rawForm[2] || "").trim();
  const jinjerId = String(getVal('jinjer番号') || rawForm[8] || "").trim();
  const kanaRaw = String(getVal('シメイ') || getVal('フリガナ') || getVal('カナ') || rawForm[4] || "").trim();
  const entryDateRaw = getVal('入職日') || rawForm[5];
  const specialty = String(getVal('専門') || rawForm[6] || "").trim();
  const holChoice = String(getVal('祝日') || "").trim();      
  const nyChoice = String(getVal('年末年始') || "").trim();
  const kanriChoice = String(getVal('管理医師') || "").trim();

  // ========================================================
  // 3. データの整形・正規化
  // ========================================================
  const cleanKana = kanaRaw.replace(/[\s　\/／]/g, "").trim(); // 空白・スラッシュ完全除去
  const kanriStatus = (kanriChoice.includes("管理") || kanriChoice.includes("有")) ? "管理" : "";
  const bikouHolText = holChoice.includes("無") ? "勤務なし" : "勤務あり";
  const bikouNyText = nyChoice.includes("無") ? "勤務なし" : "勤務あり";

  // 入職日の整形 (YYYY/MM/DD)
  let formattedStart = "";
  if (entryDateRaw instanceof Date) {
    formattedStart = Utilities.formatDate(entryDateRaw, Session.getScriptTimeZone(), "yyyy/MM/dd");
  } else if (entryDateRaw) {
    const parsed = new Date(entryDateRaw);
    formattedStart = !isNaN(parsed.getTime()) ? Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy/MM/dd") : String(entryDateRaw).trim();
  }
  
  // 期間の整形
  let initialShiftPeriod = "";
  let periodNoZero = "";
  if (formattedStart) {
    initialShiftPeriod = `${formattedStart}～2027/03/31`;
    const pParts = formattedStart.split("/");
    if (pParts.length === 3) {
      periodNoZero = `${pParts[0]}/${parseInt(pParts[1], 10)}/${parseInt(pParts[2], 10)}～2027/3/31`;
    } else {
      periodNoZero = `${formattedStart}～2027/3/31`;
    }
  }

  // 頻度の正規化関数（「第2週, 第4週」➔「第2・4」）
  const normalizeFrequency = (rawFreq) => {
    if (!rawFreq) return "";
    if (rawFreq.includes("毎週")) return "毎週";
    let nums = rawFreq.match(/[1-5１-５]/g);
    if (nums) {
      let normalizedNums = [...new Set(nums.map(v => parseInt(v.replace(/[１-５]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)), 10)))].sort((a,b) => a-b);
      return `第${normalizedNums.join('・')}`;
    }
    return rawFreq;
  };

  // ========================================================
  // 4. 各曜日のシフト時間計算・勤務備考の構築
  // ========================================================
  const dows = ["月", "火", "水", "木", "金", "土", "日"];
  let bikouLines = periodNoZero ? [periodNoZero] : [];
  let breakBikouLines = periodNoZero ? [periodNoZero] : [];
  let uniqueLocs = new Set();
  
  let dailyHoursStr = { "月": "休", "火": "休", "水": "休", "木": "休", "金": "休", "土": "休", "日": "休" };
  let totalHoursNum = 0;

  dows.forEach((d, i) => {
    let freqIdx = sourceHeaders.findIndex(h => h.includes(d) && (h.includes('頻度') || h.includes('ペース') || h.includes('週')));
    let timeIdx = sourceHeaders.findIndex(h => h.includes(d) && (h.includes('時間') || h.includes('何時から')));
    let locIdx  = sourceHeaders.findIndex(h => h.includes(d) && (h.includes('拠点') || h.includes('勤務地')));

    // 見つからなかった場合のフォールバック（平井先生の配列をベースに担保）
    if (freqIdx === -1) freqIdx = 9 + i;
    if (timeIdx === -1) timeIdx = 16 + i;
    if (locIdx === -1)  locIdx = 23 + i;

    let rawFreq = String(rawForm[freqIdx] || "").trim();
    let time = String(rawForm[timeIdx] || "").trim();
    let rawLoc = String(rawForm[locIdx] || "").trim();

    if (time && time !== "休" && time !== "休日" && time.includes("-")) {
      let cleanLoc = rawLoc.replace(/【.*?】/, "").trim();
      if (cleanLoc) uniqueLocs.add(cleanLoc);

      let timeTrimmed = time.replace(/^0/, "").replace(/-0/, "-").replace("-", "～");
      let cleanFreq = normalizeFrequency(rawFreq);
      
      let baseLine = `【${cleanLoc}】${cleanFreq}${d}曜日：${timeTrimmed}`;
      bikouLines.push(baseLine);

      // 労働時間の計算と休憩マイナス処理
      let [sStr, eStr] = time.split("-");
      let sHour = parseInt(sStr.split(":")[0], 10);
      let eHour = parseInt(eStr.split(":")[0], 10);
      let sMin = parseInt(sStr.split(":")[1], 10) / 60;
      let eMin = parseInt(eStr.split(":")[1], 10) / 60;

      let sDecimal = sHour + (isNaN(sMin) ? 0 : sMin);
      let eDecimal = eHour + (isNaN(eMin) ? 0 : eMin);
      let hours = eDecimal - sDecimal;

      let hasBreak = false;
      if (sDecimal <= 13 && eDecimal >= 15) {
        hasBreak = true;
        hours -= 2; 
      }

      dailyHoursStr[d] = `${hours}h`;
      totalHoursNum += hours;

      if (hasBreak) {
        breakBikouLines.push(`${baseLine}　(休憩13:00～15:00)`);
      } else {
        breakBikouLines.push(baseLine);
      }
    }
  });

  const footer = `（祝日：${bikouHolText}／年末年始：${bikouNyText}）`;
  bikouLines.push(footer);
  breakBikouLines.push(footer);

  // 主務の判定
  const shumu = uniqueLocs.size === 1 ? Array.from(uniqueLocs)[0] : "";

  // ========================================================
  // 5. 転記先シートへの動的マッピングと一括書き込み
  // ========================================================
  const targetLastCol = targetSheet.getLastColumn();
  const actualTargetLastCol = targetLastCol > 0 ? targetLastCol : 36;
  const targetHeaders = targetLastCol > 0 ? targetSheet.getRange(1, 1, 1, targetLastCol).getValues()[0].map(h => String(h).trim()) : [];
  
  const getTargetCol = (name) => targetHeaders.indexOf(name);
  let newRowData = new Array(actualTargetLastCol).fill("");
  
  const mapData = (colName, val) => {
    const idx = getTargetCol(colName);
    if (idx !== -1) newRowData[idx] = val;
  };

  mapData('医籍番号', medId);
  mapData('jinjer番号', jinjerId);
  mapData('医師名', doctorName);
  mapData('シメイ', cleanKana);
  mapData('入職日', formattedStart);
  mapData('専門', specialty);
  mapData('主務', shumu);
  mapData('祝日', holChoice);
  mapData('年末年始', nyChoice);
  mapData('管理医師', kanriStatus);
  mapData('週労働', totalHoursNum); // ※定期非常勤の場合は既存の計算スクリプトで後で上書き可能
  mapData('勤務備考', bikouLines.join('\n'));
  mapData('勤務備考（休憩あり）', breakBikouLines.join('\n'));
  mapData('当初シフト(期間)', initialShiftPeriod);
  
  dows.forEach(d => {
    mapData(d, dailyHoursStr[d]);
  });
  mapData('合計', totalHoursNum + "h");

  // 最下行へ追加
  const appendRowIdx = targetSheet.getLastRow() + 1;
  targetSheet.getRange(appendRowIdx, 1, 1, actualTargetLastCol).setValues([newRowData]);

  console.log(`✅ 【転記成功】${doctorName} 先生のデータを「${targetSheetName}」の ${appendRowIdx} 行目に書き込みました。`);
}