/**
 * 新規利用者のための初期設定（認証）用スクリプト
 */
function forceAuthorization() {
  const ui = SpreadsheetApp.getUi();
  try {
    // 外部通信（UrlFetchApp）の権限をGoogleに要求するためだけのダミー処理
    UrlFetchApp.fetch("https://www.google.com");
    
    // 成功した場合はメッセージを表示
    ui.alert(
      "✅ 初期設定（承認）が完了しました！\n\n" +
      "これで「勤怠シート一括作成」などの機能が正常に使えるようになりました。\n" +
      "元の作業に戻ってください。"
    );
  } catch (e) {
    ui.alert("エラーが発生しました: " + e.message);
  }
}