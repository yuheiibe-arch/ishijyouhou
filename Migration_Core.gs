// ==========================================
// 【Migration_Core】移行機能のコアロジック
// ==========================================

const CALENDAR_SOURCE_ID = '1WlmirSDOPnIcV2cwY5ClXkMWBFMM-4zrAw4XFDNUPGw'; // カレンダーソース
const MIGRATION_DEST_ID = '13PlAElj8SODdBNME28TIDpifxG0oxS3wAnKIsFTa53k'; // 移行先スプシ

/**
 * メニューから呼ばれる関数：ダイアログ表示
 */
function openMigrationDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Migration_UI')
    .setWidth(400)
    .setHeight(380);
  SpreadsheetApp.getUi().showModalDialog(html, '次年度移行処理');
}

/**
 * クライアント側へ年度リストを返す（カレンダーソースから取得）
 */
function getAvailableYears() {
  try {
    const ss = SpreadsheetApp.openById(CALENDAR_SOURCE_ID);
    const sheets = ss.getSheets();
    const currentYear = new Date().getFullYear();
    const currentMonth = new Date().getMonth() + 1;
    
    // 現在が1~3月なら、今の年は「次年度」ではなく「今年度」扱いなので、currentYearはそのままでOK
    // 例: 2026年2月時点 -> 次年度は2026年度(4月〜)
    
    // シート名から "20xx年度" を抽出してフィルタ
    const years = [];
    sheets.forEach(s => {
      const match = s.getName().match(/(\d{4})年度?/);
      if (match) {
        const y = parseInt(match[1]);
        // 未来の年度のみ（または現在の年度以降）を表示
        if (y >= currentYear) {
          years.push(`${y}年度`);
        }
      }
    });
    
    return years.sort(); // 昇順
  } catch (e) {
    console.error(e);
    return [];
  }
}

/**
 * 実行メイン関数
 */
function executeMigration(targetYear, targetType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let msg = "";

  try {
    // 1. 常勤の処理
    if (targetType === 'all' || targetType === 'full') {
      const count = createFullTimeMigrationSheet(targetYear);
      msg += `【常勤】${count}件 作成完了\n`;
    }

    // 2. 定期非常勤の処理
    if (targetType === 'all' || targetType === 'part') {
      const count = createPartTimeMigrationSheet(targetYear);
      msg += `【定期非常勤】${count}件 作成完了\n`;
    }

    // 完了後のダイアログ表示
    const destSS = SpreadsheetApp.openById(MIGRATION_DEST_ID);
    const url = destSS.getUrl();
    
    const htmlOutput = HtmlService.createHtmlOutput(
      `<div style="font-family:sans-serif; padding:10px;">
        <h3 style="color:#0f9d58;">作成完了</h3>
        <p style="white-space:pre-wrap;">${msg}</p>
        <p>以下のリンクから移行シートを確認してください。</p>
        <p style="margin-top:15px; text-align:center;">
          <a href="${url}" target="_blank" style="display:inline-block; padding:10px 20px; background:#4285f4; color:#fff; text-decoration:none; border-radius:4px; font-weight:bold;">👉 移行シートを開く</a>
        </p>
       </div>`
    ).setWidth(400).setHeight(300);
    
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '処理完了');
    
  } catch (e) {
    throw new Error(e.message);
  }
}