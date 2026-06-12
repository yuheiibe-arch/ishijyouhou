// ==========================================
// 【Import Module】外部シートからの転記モジュール
// ==========================================

/**
 * カスタムメニューから呼ばれるダイアログ起動関数
 */
function openExternalImportDialog() {
  ExternalDataImporter.showDialog();
}

/**
 * HTML側から呼ばれる、既存シートの存在チェック関数
 */
function checkExistingSheets(selectedSheets, targetYear) {
  return ExternalDataImporter.checkExisting(selectedSheets, targetYear);
}

/**
 * HTML側から呼ばれる転記実行関数
 */
function runExternalImportBatch(selectedSheets, targetYear) {
  return ExternalDataImporter.processImport(selectedSheets, targetYear);
}

/**
 * 処理本体をカプセル化したオブジェクト
 */
const ExternalDataImporter = {
  
  EXTERNAL_SS_ID: '13PlAElj8SODdBNME28TIDpifxG0oxS3wAnKIsFTa53k',

  /**
   * 1. 外部シートを読み込み、ダイアログを表示する
   */
  showDialog: function() {
    const ui = SpreadsheetApp.getUi();
    try {
      const extSs = SpreadsheetApp.openById(this.EXTERNAL_SS_ID);
      const extSheets = extSs.getSheets();
      
      const targetSheets = extSheets
        .map(s => s.getName())
        .filter(name => name.includes("調整"));

      if (targetSheets.length === 0) {
        ui.alert("お知らせ", "外部ファイルに「調整」を含むシートが見つかりませんでした。", ui.ButtonSet.OK);
        return;
      }

      // 年度ロジック：1〜3月なら「今年」、4〜12月なら「来年」が次年度になる
      const now = new Date();
      const currentMonth = now.getMonth() + 1;
      const currentYear = now.getFullYear();
      const nextYear = (currentMonth <= 3) ? currentYear : currentYear + 1;

      // HTMLテンプレート
      const htmlTemplate = HtmlService.createTemplate(`
        <!DOCTYPE html>
        <html>
          <head>
            <base target="_top">
            <style>
              body { font-family: sans-serif; padding: 10px; color: #333; }
              .container { display: flex; flex-direction: column; gap: 15px; }
              .form-group { display: flex; flex-direction: column; gap: 5px; }
              label { font-weight: bold; font-size: 14px; }
              input[type="number"] { padding: 8px; font-size: 14px; border-radius: 4px; border: 1px solid #ccc; width: 100px; }
              .sheet-list { border: 1px solid #ccc; padding: 10px; border-radius: 4px; max-height: 120px; overflow-y: auto; background: #f9f9f9; }
              .sheet-item { display: flex; align-items: center; gap: 8px; margin-bottom: 5px; font-size: 13px; }
              .buttons { margin-top: 10px; display: flex; justify-content: flex-end; gap: 10px; }
              button { padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; }
              .btn-cancel { background-color: #f3f3f3; color: #333; }
              .btn-submit { background-color: #4285f4; color: white; font-weight: bold; }
              .btn-submit:hover { background-color: #357ae8; }
              .btn-submit:disabled { background-color: #a0c1f9; cursor: not-allowed; }
            </style>
          </head>
          <body>
            <div class="container">
              <div class="form-group">
                <label for="year">① 作成する年度</label>
                <div><input type="number" id="year" value="<?= nextYear ?>"> 年度</div>
              </div>

              <div class="form-group">
                <label>② 転記するシートを選択</label>
                <div class="sheet-list">
                  <? for (let i = 0; i < sheets.length; i++) { ?>
                    <label class="sheet-item">
                      <input type="checkbox" name="sheetCheck" value="<?= sheets[i] ?>" checked>
                      <?= sheets[i] ?>
                    </label>
                  <? } ?>
                </div>
              </div>

              <div class="buttons">
                <button class="btn-cancel" onclick="google.script.host.close()">キャンセル</button>
                <button class="btn-submit" onclick="runScript()">転 記</button>
              </div>
            </div>

            <script>
              function runScript() {
                const year = document.getElementById('year').value;
                const checkboxes = document.querySelectorAll('input[name="sheetCheck"]:checked');
                const selectedSheets = Array.from(checkboxes).map(cb => cb.value);

                if (selectedSheets.length === 0) {
                  alert("シートを1つ以上選択してください。");
                  return;
                }
                if (!year) {
                  alert("年度を入力してください。");
                  return;
                }

                const btn = document.querySelector('.btn-submit');
                btn.disabled = true;
                btn.innerText = '確認中...';

                google.script.run
                  .withSuccessHandler(function(existingSheets) {
                    if (existingSheets.length > 0) {
                      const msg = "以下のシートは既に存在します。上書き（既存のシートを削除して再作成）しますか？\\n\\n" + existingSheets.join("\\n");
                      if (!confirm(msg)) {
                        btn.disabled = false;
                        btn.innerText = '転 記';
                        return;
                      }
                    }

                    btn.innerText = '処理中...';
                    google.script.run
                      .withSuccessHandler(function(res) {
                        google.script.host.close();
                      })
                      .withFailureHandler(function(err) {
                        alert('エラーが発生しました: ' + err);
                        btn.disabled = false;
                        btn.innerText = '転 記';
                      })
                      .runExternalImportBatch(selectedSheets, year);
                  })
                  .withFailureHandler(function(err) {
                    alert('エラーが発生しました: ' + err);
                    btn.disabled = false;
                    btn.innerText = '転 記';
                  })
                  .checkExistingSheets(selectedSheets, year);
              }
            </script>
          </body>
        </html>
      `);

      htmlTemplate.nextYear = nextYear;
      htmlTemplate.sheets = targetSheets;
      
      const html = htmlTemplate.evaluate().setWidth(400).setHeight(350);
      ui.showModalDialog(html, '外部シートからの転記');

    } catch (e) {
      ui.alert("エラー", "外部ファイルへのアクセスに失敗しました。\n" + e.message, ui.ButtonSet.OK);
    }
  },

  /**
   * 作成予定のシートが既に存在するかをチェックしてリストを返す
   */
  checkExisting: function(selectedSheets, targetYear) {
    const activeSs = SpreadsheetApp.getActiveSpreadsheet();
    const existingNames = [];

    for (const sheetName of selectedSheets) {
      const isPartTime = sheetName.includes("非常勤");
      const isFullTime = !isPartTime && sheetName.includes("常勤");
      
      const targetName = (isFullTime ? "常勤" : isPartTime ? "定期非常勤" : "不明") + targetYear + "年度";
      
      if (activeSs.getSheetByName(targetName)) {
        existingNames.push(targetName);
      }
    }
    return [...new Set(existingNames)];
  },

  /**
   * 2. 実際の転記処理を行う
   */
  processImport: function(selectedSheets, targetYear) {
    const extSs = SpreadsheetApp.openById(this.EXTERNAL_SS_ID);
    const activeSs = SpreadsheetApp.getActiveSpreadsheet();
    let successCount = 0;

    for (const sheetName of selectedSheets) {
      const sourceSheet = extSs.getSheetByName(sheetName);
      if (!sourceSheet) continue;

      const isPartTime = sheetName.includes("非常勤");
      const isFullTime = !isPartTime && sheetName.includes("常勤");
      
      const targetName = (isFullTime ? "常勤" : isPartTime ? "定期非常勤" : "不明") + targetYear + "年度";

      // 上書き対応：同名のシートがすでに存在する場合は削除する
      const oldSheet = activeSs.getSheetByName(targetName);
      if (oldSheet) {
        activeSs.deleteSheet(oldSheet);
      }

      // Activeなスプレッドシートへ丸ごとコピー
      const newSheet = sourceSheet.copyTo(activeSs);

      if (isFullTime) {
        this.formatFullTimeSheet_(newSheet);
      } else if (isPartTime) {
        // 定期非常勤の場合は、選択された年度(targetYear)も渡す
        this.formatPartTimeSheet_(newSheet, targetYear);
      }

      newSheet.setName(targetName);
      successCount++;
    }

    activeSs.toast(`${successCount}件のシートを転記（完了）しました！`, "完了", 5);
    return true;
  },

  /**
   * 3. 常勤用のデータ成形 (完全動的化)
   */
  formatFullTimeSheet_: function(sheet) {
    const data = sheet.getDataRange().getValues();
    let headers = data[0];
    
    // 退職者の除外
    const retireIdx = headers.findIndex(h => String(h).includes("退職"));
    if (retireIdx > -1) {
      for (let i = data.length - 1; i > 0; i--) {
        const retirementDate = String(data[i][retireIdx]).trim();
        if (retirementDate !== "") {
          sheet.deleteRow(i + 1);
        }
      }
    }

    // 作業用列のクリア（名前で動的に検索）
    const clearKeywords = ["保留", "対応状況", "内容修正", "次年度用", "前年度からの変更"];
    clearKeywords.forEach(kw => {
      const idx = headers.findIndex(h => String(h).includes(kw));
      if (idx > -1 && sheet.getLastRow() > 1) {
        sheet.getRange(2, idx + 1, sheet.getLastRow() - 1, 1).clearContent();
      }
    });

    const newLastRow = sheet.getLastRow();
    if (newLastRow > 1) {
      sheet.getRange(2, 1, newLastRow - 1, sheet.getMaxColumns()).setBackground(null);
    }
  },

  /**
   * 4. 定期非常勤用のデータ成形 (完全動的化)
   */
  formatPartTimeSheet_: function(sheet, targetYear) {
    const data = sheet.getDataRange().getValues();
    let headers = data[0]; 
    
    // 各項目の列インデックスを動的に検索
    let retireIdx = headers.findIndex(h => String(h).includes("退職"));
    let notNeededIdx = headers.findIndex(h => String(h).includes("対応不要"));
    let hireIdx = headers.findIndex(h => String(h).includes("入職"));
    
    // 不要な行の除外（下から上へ）
    for (let i = data.length - 1; i > 0; i--) {
      const retirementDate = retireIdx > -1 ? String(data[i][retireIdx]).trim() : "";
      const isNotNeeded = notNeededIdx > -1 ? data[i][notNeededIdx] : false;
      const isSkip = (isNotNeeded === true || String(isNotNeeded).toUpperCase() === "TRUE");

      if (retirementDate !== "" || isSkip) {
        sheet.deleteRow(i + 1);
      }
    }

    const newLastRow = sheet.getLastRow();

    // 【追加ルール1】入職日を選択した年度の 4/1 に一括変更
    if (hireIdx !== -1 && newLastRow > 1) {
      const newHireDate = `${targetYear}/04/01`;
      sheet.getRange(2, hireIdx + 1, newLastRow - 1, 1).setValue(newHireDate);
    }

    // 【動的列整理】元の「J列〜M列削除」などの代わりに、不要な列を名前で検索してすべて削除する
    headers = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues()[0];
    for (let c = headers.length - 1; c >= 0; c--) {
      const h = String(headers[c]);
      // 調整用に使っていた「保留」「対応不要」「記入例」「ルール」「対応状況」「内容修正」列を消す
      if (h.includes("保留") || h.includes("対応不要") || h.includes("記入例") || h.includes("ルール") || h.includes("内容修正") || h.includes("対応状況")) {
        sheet.deleteColumn(c + 1);
      } else if (h.includes("提案シフト")) {
        // 次年度の「提案シフト」を本番の「勤務備考」に昇格（名前変更）させる
        sheet.getRange(1, c + 1).setValue("勤務備考");
      }
    }
    
    // 背景色をクリア（白に戻す）
    const finalMaxCols = sheet.getMaxColumns();
    if (newLastRow > 1) {
      sheet.getRange(2, 1, newLastRow - 1, finalMaxCols).setBackground(null);
    }
  }
};