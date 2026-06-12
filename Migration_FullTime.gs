// ==========================================
// 【Migration_FullTime】常勤 次年度移行ロジック（元シート変更なし版）
// ==========================================

function createFullTimeMigrationSheet(targetYear) {
  const currentYearNum = parseInt(targetYear.replace("年度", "")) - 1; 
  const nextYearNum = parseInt(targetYear.replace("年度", ""));

  const srcSheetName = `常勤${currentYearNum}年度`;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = ss.getSheetByName(srcSheetName);
  
  if (!srcSheet) throw new Error(`現在のシート「${srcSheetName}」が見つかりません。`);

  const data = srcSheet.getDataRange().getValues();
  const headers = data[0];
  const hMap = {};
  headers.forEach((h, i) => hMap[h] = i);
  
  // 必須列チェック
  if (hMap["退職日"] === undefined || hMap["次年度用\n前年度からの変更"] === undefined) {
    throw new Error(`シート「${srcSheetName}」に必須の列（退職日 または 次年度用...）がありません。`);
  }

  const outputRows = [];
  // グレーアウト用の配列は削除しました

  // --- ループ処理 ---
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // 除外判定
    const retireDate = row[hMap["退職日"]];
    const nextChange = row[hMap["次年度用\n前年度からの変更"]];
    
    if (retireDate instanceof Date) continue;
    if (String(nextChange) === "退職") continue;

    // 転記データの作成
    const newRow = [...row]; 
    newRow[0] = outputRows.length + 1; // No再番

    const isChange = (String(nextChange) === "あり");

    if (isChange) {
      // パターンB: 変更あり -> 備考とシフト列を空に
      newRow[hMap["勤務備考"]] = "";
      if (hMap["当初シフト(期間)"] !== undefined) newRow[hMap["当初シフト(期間)"]] = "";
      
    } else {
      // パターンA: 変更なし -> 日付更新
      const currentRemarks = row[hMap["勤務備考"]];
      const lastBlock = extractLastShiftBlock(currentRemarks); 
      
      if (lastBlock) {
        const newTerm = `${nextYearNum}/04/01～${nextYearNum + 1}/03/31`;
        newRow[hMap["勤務備考"]] = `${newTerm}\n${lastBlock.content}`;
      }
    }

    outputRows.push(newRow);
  }

  // --- 移行先への書き込み（シートコピー方式） ---
  const destSS = SpreadsheetApp.openById(MIGRATION_DEST_ID);
  const destMasterName = `(調整)常勤${nextYearNum}年度`;
  
  // 既に同名シートがあれば削除
  const existingSheet = destSS.getSheetByName(destMasterName);
  if (existingSheet) destSS.deleteSheet(existingSheet);

  // 1. 元シートを移行先スプシにコピー
  const newSheet = srcSheet.copyTo(destSS);
  newSheet.setName(destMasterName);

  // 2. データをクリア（ヘッダー1行目は残す）
  if (newSheet.getLastRow() > 1) {
    newSheet.getRange(2, 1, newSheet.getLastRow() - 1, newSheet.getLastColumn()).clearContent();
    newSheet.getRange(2, 1, newSheet.getLastRow() - 1, newSheet.getLastColumn()).setBackground(null); 
  }

  // 3. 新しいデータを書き込み
  if (outputRows.length > 0) {
    newSheet.getRange(2, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
  }

  // ※元シートへの操作（グレーアウト等）は一切行いません

  return outputRows.length;
}

/** 勤務備考抽出ヘルパー */
function extractLastShiftBlock(text) {
  if (!text) return null;
  const blocks = text.split(/\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}.*?[\n\r]/g);
  const validBlocks = blocks.filter(b => b.trim().length > 0);
  if (validBlocks.length === 0) return { content: text.trim() };
  return { content: validBlocks[validBlocks.length - 1].trim() };
}