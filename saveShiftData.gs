function saveShiftData(rowIndex, text, mode, sendPermission, sheetNameFromHtml) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // 【デバッグ1】受け取ったデータをログと画面右下に出力
    console.log(`[Debug] saveShiftData 呼び出し: rowIndex=${rowIndex}, mode=${mode}, sheetName=${sheetNameFromHtml}`);
    ss.toast(`保存開始: 行番号=${rowIndex}`, "デバッグ1", 5);

    // 【原因調査】UIから行番号が送られてきているか？
    if (rowIndex === undefined || rowIndex === null || isNaN(rowIndex)) {
      throw new Error("行番号(rowIndex)が不明です。UI(HTML)側から正しくデータが送られていません！");
    }

    const sheetName = sheetNameFromHtml || ss.getActiveSheet().getName();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`シート「${sheetName}」が見つかりません。`);

    // 定期非常勤の場合はそのまま処理
    if (sheetName.includes("定期非常勤")) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const targetColIdx = headers.findIndex(h => h && h.toString().trim() === "勤務備考");
      if (targetColIdx === -1) throw new Error(`「勤務備考」列が見つかりません。`);
      
      sheet.getRange(rowIndex, targetColIdx + 1).setValue(text);
      ss.toast("定期非常勤：勤務備考を更新しました。", "保存完了"); 
      return "定期非常勤：更新完了";
    }

    // 常勤の場合はライブラリへ渡す
    ss.toast(`ライブラリへ送信中: ${sheetName}`, "デバッグ2", 5);
    const result = ShiftEditorLib.saveShiftData(rowIndex, text, mode, sendPermission, sheetName);
    
    ss.toast("保存成功！", "デバッグ完了", 5);
    return result;

  } catch (e) {
    // 【デバッグエラー】エラー内容をログと画面に出力
    console.error("[Debug Error] " + e.stack);
    ss.toast(`エラー停止: ${e.message}`, "🚨致命的エラー🚨", 15);
    
    // UI（画面）側へエラーを強制的に投げる（これで無限フリーズを解除します）
    throw new Error(e.message);
  }
}