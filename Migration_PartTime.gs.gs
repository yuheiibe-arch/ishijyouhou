// ==========================================
// 【Migration_PartTime】定期非常勤 次年度移行ロジック（テンプレート活用・行数最適化版）
// ==========================================

function createPartTimeMigrationSheet(targetYear) {
  // 入力: "2026年度" -> 現在: 2025, 次: 2026
  const currentYearNum = parseInt(targetYear.replace("年度", "")) - 1;
  const nextYearNum = parseInt(targetYear.replace("年度", ""));
  
  const srcSheetName = `定期非常勤${currentYearNum}年度`;
  const destSheetName = `(調整中)定期非常勤${nextYearNum}年度`;
  const templateSheetName = "テンプレート"; // 固定のテンプレートシート名

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = ss.getSheetByName(srcSheetName);
  
  if (!srcSheet) throw new Error(`現在のシート「${srcSheetName}」が見つかりません。`);

  // --- 1. テンプレート情報の取得（移行先スプシから） ---
  const destSS = SpreadsheetApp.openById(MIGRATION_DEST_ID);
  
  const templateSheet = destSS.getSheetByName(templateSheetName);
  if (!templateSheet) {
    throw new Error(`移行先スプレッドシートに「${templateSheetName}」シートが見つかりません。`);
  }

  // ★動的化：テンプレートのヘッダーを読み取り、列数と配置を自動把握する
  const tmplMaxCols = templateSheet.getLastColumn();
  const tmplHeaders = templateSheet.getRange(1, 1, 1, tmplMaxCols).getValues()[0];

  // M列（13列目）の「書き方ルール」テキストを取得 (テンプレートの2行目から取得)
  // ★動的化：13列目という数字の決め打ちをやめ、「ルール」という文字がある列から取得
  let templateRuleText = "";
  const ruleIdx = tmplHeaders.findIndex(h => String(h).includes("ルール"));
  if (ruleIdx > -1 && templateSheet.getLastRow() >= 2) {
    templateRuleText = templateSheet.getRange(2, ruleIdx + 1).getValue();
  }

  // --- 2. データ作成 ---
  const data = srcSheet.getDataRange().getValues();
  const headers = data[0];
  const hMap = {};
  headers.forEach((h, i) => hMap[h] = i);

  const outputRows = [];
  let rowCount = 1;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // 列名から動的に値を取得
    const retireDate = hMap["退職日"] !== undefined ? row[hMap["退職日"]] : "";
    const nextChange = hMap["次年度用\n前年度からの変更"] !== undefined ? row[hMap["次年度用\n前年度からの変更"]] : "";
    
    // 除外判定
    if (retireDate instanceof Date) continue;
    if (String(nextChange).includes("退職") || String(nextChange).includes("満了")) continue;

    const currentShiftText = hMap["勤務備考"] !== undefined ? row[hMap["勤務備考"]] : "";
    const isChange = (String(nextChange) === "あり");

    // ★動的化：27列固定ではなく、テンプレートの列数に合わせた空箱を作る
    const newRow = new Array(tmplMaxCols).fill("");

    // テンプレートのヘッダー名に合わせて、正しい位置にデータを流し込む
    for (let c = 0; c < tmplMaxCols; c++) {
      const th = String(tmplHeaders[c]).trim();
      
      if (th === "番号") {
        newRow[c] = rowCount;
      } else if (th === "退職日") {
        newRow[c] = ""; // 退職日はリセット
      } else if (th.includes("保留")) {
        newRow[c] = isChange ? true : false;
      } else if (th.includes("対応不要")) {
        newRow[c] = "";
      } else if (th.includes("記入例")) {
        newRow[c] = currentShiftText;
      } else if (th.includes("ルール")) {
        newRow[c] = templateRuleText;
      } else if (th.includes("提案シフト")) {
        // 変更なし -> 保留OFF、上期固定で日付更新 / 変更あり -> 提案シフト空欄
        newRow[c] = isChange ? "" : generateKamikiShiftText(currentShiftText, nextYearNum);
      } else if (th === "契約時給") {
        newRow[c] = "時給表どおり";
      } else if (hMap[th] !== undefined) {
        // マスタに同じ名前の列（医籍番号、jinjer番号、シメイ等）があれば、すべてそのままコピー
        newRow[c] = row[hMap[th]];
      }
    }
    
    outputRows.push(newRow);
    rowCount++;
  }

  // --- 3. シート作成と設定 ---
  
  // 既存の同名シートがあれば削除
  const existingSheet = destSS.getSheetByName(destSheetName);
  if (existingSheet) destSS.deleteSheet(existingSheet);

  // テンプレートをコピーして新しいシートを作成
  const newSheet = templateSheet.copyTo(destSS);
  newSheet.setName(destSheetName);

  // --- 4. ヘッダーの動的書き換え ---
  // テンプレートのヘッダーを取得
  const headerRange = newSheet.getRange(1, 1, 1, tmplMaxCols);
  const headerValues = [...tmplHeaders];

  // ★動的化：列番号での決め打ちをやめ、文字を探して書き換える
  for (let c = 0; c < headerValues.length; c++) {
    if (String(headerValues[c]).includes("記入例")) {
      headerValues[c] = `${currentYearNum}年度シフト（通期）記入例`;
    } else if (String(headerValues[c]).includes("提案シフト")) {
      headerValues[c] = `${nextYearNum}年度提案シフト`;
    }
  }

  // 書き戻し
  headerRange.setValues([headerValues]);

  // --- 5. データの書き込み ---
  
  // テンプレートに入っていた既存データ（2行目以降）をクリア
  // ※書式は残したいので clearContent() を使用
  const currentMaxRows = newSheet.getMaxRows();
  if (currentMaxRows > 1) {
    newSheet.getRange(2, 1, currentMaxRows - 1, tmplMaxCols).clearContent();
    // チェックボックスの状態などはクリアされないことがあるため、明示的にUncheckが必要ならここで行う
    // (今回は値の上書きで対応)
  }

  if (outputRows.length > 0) {
    newSheet.getRange(2, 1, outputRows.length, tmplMaxCols).setValues(outputRows);
  }

  // --- 6. 行数の最適化 (削除または追加) ---
  const requiredRows = outputRows.length + 1; // ヘッダー(1) + データ数

  if (currentMaxRows > requiredRows) {
    // 行が余っている場合 -> 削除
    newSheet.deleteRows(requiredRows + 1, currentMaxRows - requiredRows);
  } else if (currentMaxRows < requiredRows) {
    // 行が足りない場合 -> 追加
    newSheet.insertRowsAfter(currentMaxRows, requiredRows - currentMaxRows);
  }

  return outputRows.length;
}

/**
 * 上期固定の日付変換（4/1～9/30）
 */
function generateKamikiShiftText(text, nextYear) {
  if (!text) return "";
  const targetTerm = `${nextYear}/04/01～${nextYear}/09/30`;
  
  // 日付行があれば置換
  if (/\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}.*?[\n\r]/.test(text)) {
    return text.replace(/\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}.*?[\n\r]/, targetTerm + "\n");
  }
  // なければ先頭に追加
  return targetTerm + "\n" + text;
}