// ==========================================
// 【AutoUpdate】退職ステータスの自動同期（常勤・非常勤対応版）
// ==========================================

/**
 * シートの退職日を監視し、ステータスを自動更新する関数
 * トリガー設定：時間主導型（例: 15分〜30分おき）
 */
function syncRetirementStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  // 対象：シート名に「常勤」または「定期非常勤」を含み、かつ「xxxx年度」で終わるもの
  const targetSheets = sheets.filter(s => s.getName().match(/(常勤|定期非常勤)\d{4}年度$/));

  targetSheets.forEach(sheet => {
    processSheetForRetirement(sheet);
  });
}

/**
 * 各シートごとの処理ロジック
 */
function processSheetForRetirement(sheet) {
  const sheetName = sheet.getName();
  const isPartTime = sheetName.includes("定期非常勤"); // 定期非常勤かどうかの判定フラグ
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  // --- 1. 列インデックスの特定 ---
  const idxRetire = headers.indexOf("退職日"); // VLOOKUPの結果
  
  // 変更対象1: "次年度用...変更"（常勤・非常勤 共通）
  const idxNextChange = headers.findIndex(h => String(h).includes("次年度用") && String(h).includes("前年度からの変更")); 
  
  // 変更対象2: "年度内更新"（定期非常勤のみ）
  const idxUpdate = (isPartTime) ? headers.indexOf("年度内更新") : -1;

  // 必須列（退職日、次年度変更）がなければスキップ
  if (idxRetire === -1 || idxNextChange === -1) return;

  // --- 2. データ取得 ---
  // 行数が多いため、列ごとに取得してメモリ節約
  const rangeRetire = sheet.getRange(2, idxRetire + 1, lastRow - 1, 1);
  const rangeNext = sheet.getRange(2, idxNextChange + 1, lastRow - 1, 1);
  
  const valRetire = rangeRetire.getValues();
  const valNext = rangeNext.getValues();
  
  let rangeUpdate, valUpdate;
  if (idxUpdate > -1) {
    rangeUpdate = sheet.getRange(2, idxUpdate + 1, lastRow - 1, 1);
    valUpdate = rangeUpdate.getValues();
  }

  let isChanged = false; // 変更があったかどうかのフラグ

  // --- 3. 行ごとに判定 ---
  for (let i = 0; i < valRetire.length; i++) {
    const retireDate = valRetire[i][0];
    
    // 現在の値
    const currentNext = valNext[i][0];
    
    // 退職日が入っている場合のみ判定
    if (retireDate instanceof Date) {
      
      // A. 【共通】次年度変更列を「退職」にする
      const targetNextVal = "退職";
      if (currentNext !== targetNextVal) {
        valNext[i][0] = targetNextVal;
        isChanged = true;
      }

      // B. 【定期非常勤のみ】年度内更新列の判定
      if (isPartTime && valUpdate) {
        const currentUpdate = valUpdate[i][0];
        
        // 年度末かどうか判定 (3月31日なら契約満了)
        // ※JavaScriptの月は0始まり (0=1月, ... 2=3月)
        const isFyEnd = (retireDate.getMonth() === 2 && retireDate.getDate() === 31);
        const targetUpdateVal = isFyEnd ? "契約満了" : "中途解除";

        if (currentUpdate !== targetUpdateVal) {
          valUpdate[i][0] = targetUpdateVal;
          isChanged = true;
        }
      }
    }
  }

  // --- 4. 変更があった場合のみ書き込み ---
  if (isChanged) {
    rangeNext.setValues(valNext);
    if (isPartTime && rangeUpdate) {
      rangeUpdate.setValues(valUpdate);
    }
    console.log(`[自動更新] ${sheetName} のステータスを更新しました。`);
  }
}