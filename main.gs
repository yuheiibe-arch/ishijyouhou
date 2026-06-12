// ==========================================
// 【Main】フォーム送信トリガー (エントリーポイント)
// ==========================================

/**
 * フォーム送信時に実行される関数
 * トリガー設定：スプレッドシートから > フォーム送信時
 * @param {Object} e - イベントオブジェクト
 */
function onFormSubmit(e) {
  // デバッグ用: イベントがない場合（エディタから直接実行した場合）のガード
  if (!e) {
    console.error("この関数はトリガーから実行してください。");
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = e.range.getSheet(); // 回答が書き込まれたシート
    
    // 1. ヘッダーと今回の回答データを配列として整理
    //    (e.namedValues は順序が保証されないため、シートのヘッダー順に並べ直す)
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    
    // e.range.getRow() は今回書き込まれた行番号
    // e.values はフォーマット前の値の可能性があるため、確実にシートから値を取得する
    const currentRow = e.range.getRow();
    const rowData = sheet.getRange(currentRow, 1, 1, lastCol).getValues()[0];

    console.log(`=== 新規回答処理開始 [行: ${currentRow}] ===`);

    // 2. 契約テキスト生成 (ContractGenerator.gs)
    const genResult = generateTextForRow(headers, rowData);
    const contractText = genResult.contract;
    const contractWithBreakText = genResult.contractWithBreak; // ★追加：休憩ありテキスト
    const targetYear = genResult.fy; // "2026年度" などが返る

    console.log(`対象年度: ${targetYear}`);

    // 3. マスタシートへの転記 (MasterSync.gs)
    // ★引数の最後に contractWithBreakText を追加して渡す
    syncAndExpandSchedule(headers, rowData, contractText, targetYear, contractWithBreakText);

    // 4. 勤怠シートへの列追加
    //    採用区分を取得して分岐
    const typeIdx = headers.findIndex(h => String(h).includes("採用区分"));
    const type = (typeIdx > -1) ? rowData[typeIdx] : "";

    // ★引数の最後に contractWithBreakText を追加して渡す
    if (type === "常勤") {
      appendFullTimeDoctor_Single(targetYear, headers, rowData, contractText, contractWithBreakText);
    } else if (type === "定期非常勤") {
      appendPartTimeDoctor_Single(targetYear, headers, rowData, contractText, contractWithBreakText);
    } else {
      console.warn(`未定義の採用区分です: ${type}`);
    }

    console.log("=== 処理完了 ===");

  } catch (err) {
    console.error(`メイン処理でエラーが発生しました: ${err.stack}`);
    // 必要であれば管理者にメール通知などをここに追加
  }
}