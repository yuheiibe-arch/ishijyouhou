// ==========================================
// 1. 【Main Logic】メニュー・UI・自動化処理セクション
// ==========================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚡️管理メニュー')
    .addItem('① 医師を追加', 'openDoctorViewForm')
    .addItem('② フォーム編集画面', 'openDoctorEditForm')
    .addItem('③ 拠点を追加（マスタ登録＆更新）', 'addNewLocation')
    .addSeparator()
    .addItem('④ 契約内容を変更（エディタ起動）', 'launchContractEditor')
    .addSeparator()
    .addItem('★ 外部シートから転記', 'openExternalImportDialog')
    .addItem('★ 勤怠シート一括作成（手動）', 'openManualGenerationDialog')
    .addSeparator()
    .addSubMenu(ui.createMenu('🚀 次年度更新準備')
        .addItem('次年度移行シート作成', 'openMigrationDialog'))
    .addToUi();
}

function openDoctorViewForm() { openUrlDirectly('https://docs.google.com/forms/d/e/1FAIpQLSf2Gu90DPTzv06QwVHlOWnAHphyCaqhIRd_X64oIUs1lmG6hg/viewform?usp=dialog'); }
function openDoctorEditForm() {
  try { openUrlDirectly(FormApp.openById('1SDZBu-CWfLtrfou_fFFTmqDjumkPWnPbUhmTsItQN9Q').getEditUrl()); } 
  catch (e) { SpreadsheetApp.getUi().alert('エラー: ' + e.message); }
}

function openUrlDirectly(url) {
  const html = HtmlService.createHtmlOutput(`<html><body style="font-family: sans-serif; text-align: center; padding-top: 20px;"><p>ページを開いています...</p><p style="font-size: 12px; color: gray;">ポップアップがブロックされた場合は<br><a href="${url}" target="_blank">こちらをクリック</a>してください。</p><script>window.onload = function() { var win = window.open('${url}', '_blank'); if (win) google.script.host.close(); };</script></body></html>`).setWidth(250).setHeight(100);
  SpreadsheetApp.getUi().showModalDialog(html, '転送中...');
}

function addNewLocation() {
  const html = HtmlService.createHtmlOutputFromFile('dialog_location')
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '拠点の一括追加とマスター更新');
}

function processNewLocations(locationsToAdd) {
  const MASTER_SS_ID = '14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs';
  const FORM_ID = '1SDZBu-CWfLtrfou_fFFTmqDjumkPWnPbUhmTsItQN9Q';
  
  const ss = SpreadsheetApp.openById(MASTER_SS_ID);
  const sheet = ss.getSheetByName('拠点名');
  
  const lastRow = sheet.getLastRow();
  let masterData = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 6).getValues() : [];
  
  const existingSet = new Set(masterData.map(row => row[0] ? `【${row[5] || ""}】${row[0]}` : ""));

  let addedCount = 0;
  let skippedCount = 0;
  let newRows = [];

  locationsToAdd.forEach(loc => {
    const fullDisplayName = `【${loc.area}】${loc.name}`;
    if (existingSet.has(fullDisplayName)) {
      skippedCount++;
    } else {
      newRows.push([loc.name, "", "", "", "", loc.area]);
      addedCount++;
    }
  });

  if (addedCount > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 6).setValues(newRows);
  }

  const updatedData = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  updatedData.sort((a, b) => {
    const areaA = a[5] || "ZZ"; const areaB = b[5] || "ZZ";
    if (areaA < areaB) return -1; if (areaA > areaB) return 1;
    const nameA = a[0] || ""; const nameB = b[0] || "";
    if (nameA < nameB) return -1; if (nameA > nameB) return 1;
    return 0;
  });
  
  sheet.getRange(2, 1, updatedData.length, 6).setValues(updatedData);

  let choices = ["同拠点", "休日", "勤務なし"];
  updatedData.forEach(row => {
    if (row[0]) choices.push(row[5] ? `【${row[5]}】${row[0]}` : row[0]);
  });
  choices = [...new Set(choices)];

  let formUpdateCount = 0;
  const form = FormApp.openById(FORM_ID);
  form.getItems().forEach(item => {
    if (item.getTitle().includes('拠点')) {
      const type = item.getType();
      if (type === FormApp.ItemType.LIST) {
        item.asListItem().setChoiceValues(choices);
        formUpdateCount++;
      } else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
        item.asMultipleChoiceItem().setChoiceValues(choices);
        formUpdateCount++;
      }
    }
  });

  return `登録完了！\n\n・追加: ${addedCount}件\n・重複スキップ: ${skippedCount}件\n・フォーム更新: ${formUpdateCount}箇所`;
}

function launchContractEditor() { ShiftEditorLib.openSystemSelector(); }

function openManualGenerationDialog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const years = new Set();
  ss.getSheets().forEach(sheet => {
    const match = sheet.getName().match(/(?:常勤|定期非常勤)(\d{4})年度/);
    if (match) years.add(match[1]);
  });
  if (years.size === 0) { years.add(String(new Date().getFullYear())); years.add(String(new Date().getFullYear() + 1)); }

  const htmlTemplate = HtmlService.createTemplate(`
    <!DOCTYPE html><html><head><base target="_top"><style>body{font-family:sans-serif;padding:10px;color:#333}.container{display:flex;flex-direction:column;gap:15px}.form-group{display:flex;flex-direction:column;gap:5px}label{font-weight:bold;font-size:14px}select{padding:8px;font-size:14px;border-radius:4px;border:1px solid #ccc}.buttons{margin-top:10px;display:flex;justify-content:flex-end;gap:10px}button{padding:8px 16px;border:none;border-radius:4px;cursor:pointer;font-size:14px}.btn-cancel{background-color:#f3f3f3;color:#333}.btn-submit{background-color:#4285f4;color:white;font-weight:bold}.btn-submit:hover{background-color:#357ae8}.note{font-size:12px;color:#d93025;margin-top:5px}</style></head>
    <body><div class="container"><div class="form-group"><label>① 対象年度を選択</label><select id="year"><? for(var i=0;i<years.length;i++){ ?><option value="<?= years[i] ?>"><?= years[i] ?>年度</option><? } ?></select></div><div class="form-group"><label>② 対象の雇用形態</label><select id="type"><option value="all" selected>すべて (常勤 ＆ 定期非常勤)</option><option value="full">常勤のみ</option><option value="part">定期非常勤のみ</option></select></div><div class="note">※注意：実行すると内容が再計算され上書きされます。</div><div class="buttons"><button class="btn-cancel" onclick="google.script.host.close()">キャンセル</button><button class="btn-submit" onclick="runScript()">実 行</button></div></div>
    <script>function runScript(){const btn=document.querySelector('.btn-submit');btn.disabled=true;btn.innerText='処理中...';google.script.run.withSuccessHandler(function(){google.script.host.close();}).withFailureHandler(function(err){alert('エラー: '+err);btn.disabled=false;btn.innerText='実 行';}).runManualBatch(document.getElementById('year').value, document.getElementById('type').value);}</script></body></html>
  `);
  htmlTemplate.years = Array.from(years).sort().reverse(); 
  SpreadsheetApp.getUi().showModalDialog(htmlTemplate.evaluate().setWidth(350).setHeight(320), '勤怠シート一括作成');
}

// ==========================================
// ★ 差分抽出ロジック（システム記憶比較版）
// ==========================================
function getDoctorSnapshot(year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let snapshot = {};
  ['常勤', '定期非常勤'].forEach(type => {
    let sheet = ss.getSheetByName(`${type}${year}年度`);
    if (!sheet) return;
    let data = sheet.getDataRange().getDisplayValues();
    if (data.length < 2) return;
    
    let nameIdx = data[0].indexOf('医師名');
    let shiftIdx = data[0].findIndex(h => String(h).trim() === '勤務備考' || String(h).includes('シフト'));
    let enterIdx = data[0].indexOf('入職日');
    let exitIdx = data[0].indexOf('退職日');
    
    if (nameIdx === -1 || shiftIdx === -1) return;
    
    for (let i = 1; i < data.length; i++) {
      let docName = String(data[i][nameIdx]).trim();
      let shiftTxt = String(data[i][shiftIdx]).trim();
      let enterTxt = enterIdx !== -1 ? String(data[i][enterIdx]).trim() : "";
      let exitTxt = exitIdx !== -1 ? String(data[i][exitIdx]).trim() : "";
      
      if (docName) {
        snapshot[`${type}_${docName}`] = `SHIFT:${shiftTxt} | ENTER:${enterTxt} | EXIT:${exitTxt}`;
      }
    }
  });
  return snapshot;
}

function getChangedLocations(oldSnap, newSnap) {
  let allKeys = new Set([...Object.keys(oldSnap), ...Object.keys(newSnap)]);
  let affected = new Set();

  allKeys.forEach(k => {
    let o = oldSnap[k] || "";
    let n = newSnap[k] || "";

    if (o === n) return; 

    let oMeta = o.includes('| ENTER:') ? o.substring(o.indexOf('| ENTER:')) : "";
    let nMeta = n.includes('| ENTER:') ? n.substring(n.indexOf('| ENTER:')) : "";
    let oShift = o.includes('| ENTER:') ? o.substring(0, o.indexOf('| ENTER:')) : o;
    let nShift = n.includes('| ENTER:') ? n.substring(0, n.indexOf('| ENTER:')) : n;

    if (oMeta !== nMeta) {
      let m = (oShift + nShift).match(/【(.*?)】/g);
      if (m) m.forEach(x => affected.add(x.replace(/[【】]/g, '')));
      return;
    }

    let oCommon = oShift.split('\n').filter(l => !l.includes('【')).map(l => l.trim()).join('\n');
    let nCommon = nShift.split('\n').filter(l => !l.includes('【')).map(l => l.trim()).join('\n');

    if (oCommon !== nCommon) {
      let m = (oShift + nShift).match(/【(.*?)】/g);
      if (m) m.forEach(x => affected.add(x.replace(/[【】]/g, '')));
      return;
    }

    let oLines = oShift.split('\n').filter(l => l.includes('【')).map(l => l.trim());
    let nLines = nShift.split('\n').filter(l => l.includes('【')).map(l => l.trim());

    oLines.forEach(ol => {
      if (!nLines.includes(ol)) {
        let match = ol.match(/【(.*?)】/);
        if (match) affected.add(match[1]);
      }
    });
    nLines.forEach(nl => {
      if (!oLines.includes(nl)) {
        let match = nl.match(/【(.*?)】/);
        if (match) affected.add(match[1]);
      }
    });
  });

  return Array.from(affected);
}

/** * ★ 一括作成処理の実行 ＆ カレンダーへのピンポイント更新指示 */
function runManualBatch(year, type) {
  try { DriveApp.getRootFolder(); } catch(e) {}
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  
  try {
    const savedStr = props.getProperty(`DOCTOR_SNAPSHOT_${year}`);
    const oldSnapshot = savedStr ? JSON.parse(savedStr) : {};

    const processes = [];
    if (type === 'all' || type === 'part') {
      if (typeof processPartTime_AllSteps !== 'undefined') { processPartTime_AllSteps(year); processes.push("定期非常勤"); }
    }
    if (type === 'all' || type === 'full') {
      if (typeof processFullTime_AllSteps !== 'undefined') { processFullTime_AllSteps(year); processes.push("常勤"); }
    }
    if (processes.length === 0) throw new Error("実行する関数が見つかりません。");
    
    // ★★★ 勤怠作成直後に実績反映を自動実行 ★★★
    try {
      if (typeof updateAttendanceActuals === 'function') {
        updateAttendanceActuals();
      }
    } catch(e) {
      console.error("実績反映エラー: " + e.message);
    }
    // ★★★★★★★★★★★★★★★★★★★★★★★★★★★

    const newSnapshot = getDoctorSnapshot(year);
    const changedLocations = getChangedLocations(oldSnapshot, newSnapshot);

    if (changedLocations.length === 0) {
      ss.toast(`${processes.join("・")}の作成完了！\n※カレンダー側と契約の差分がなかったため、送信はスキップしました。`, "完了", 8);
      props.setProperty(`DOCTOR_SNAPSHOT_${year}`, JSON.stringify(newSnapshot));
      return;
    }

    ss.toast(`変更を検知: ${changedLocations.join(", ")}\nカレンダーのピンポイント更新を開始します...`, "処理中", 8);

    const apiUrl = "https://script.google.com/a/macros/mnys.jp/s/AKfycbzzMGqyKvNvc5daaRAxx0mKE5Ipn5mE2Ghuh6i0j-vOOws1xgFF1HBFC57od09B2-o5/exec";
    const targetsParam = encodeURIComponent(changedLocations.join(","));
    
    UrlFetchApp.fetch(apiUrl + "?year=" + year + "&targets=" + targetsParam, {
      method: "get", muteHttpExceptions: true, headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() }
    });
    
    props.setProperty(`DOCTOR_SNAPSHOT_${year}`, JSON.stringify(newSnapshot));

    ss.toast("カレンダーへの更新指示が完了しました！\n（指定された拠点だけが数秒で更新されます）", "連携成功", 8);

  } catch (e) {
    throw new Error(e.message); 
  }
}

// ==========================================
// 2. 【Bridge】ライブラリ中継（修正完了版）
// ==========================================
function getSettings() {
  try { return { locations: ShiftEditorLib.getGroupedLocations(SpreadsheetApp.openById('14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs')), locMap: ShiftEditorLib.fetchLocationMap(SpreadsheetApp.openById('14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs')) }; } 
  catch (e) { return { locations: {}, locMap: {} }; }
}
function openSystemSelector() { return ShiftEditorLib.openSystemSelector(); }
function getTargetSheetsList() { return ShiftEditorLib.getTargetSheetsList(); }
function launchEditorForSheet(sheetName) { return ShiftEditorLib.launchEditorForSheet(sheetName); }
function getDoctorList(sheetNameFromHtml) { return ShiftEditorLib.getDoctorList(sheetNameFromHtml || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()); }
function getDoctorDataByRow(rowIndex, sheetNameFromHtml) { return ShiftEditorLib.getDoctorDataByRow(rowIndex, sheetNameFromHtml || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()); }

// ■ 常勤/非常勤 共通シフト保存（メイン用）
function saveShiftData(rowIndex, text, mode, sendPermission, sheetNameFromHtml) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); 
    const sheetName = sheetNameFromHtml || ss.getActiveSheet().getName();
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) throw new Error(`シート「${sheetName}」が見つかりません。`);

    if (sheetName.includes("定期非常勤")) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const targetColIdx = headers.findIndex(h => h.toString().trim() === "勤務備考");
      if (targetColIdx === -1) throw new Error(`「${sheetName}」に「勤務備考」列が見つかりません。`);
      
      sheet.getRange(rowIndex, targetColIdx + 1).setValue(text);
      ss.toast("定期非常勤：勤務備考を更新しました。"); 
      return "定期非常勤：更新完了";
    }
    
    return ShiftEditorLib.saveShiftData(rowIndex, text, mode, sendPermission, sheetName);
  } catch (e) {
    console.error("saveShiftData Error: " + e.stack);
    throw new Error(e.message);
  }
}

// ■ 非常勤ウィザード用
function savePartTimeWizardData(rowIndex, shiftText, attributes) { 
  try {
    const sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
    return ShiftEditorLib.savePartTimeWizardData(rowIndex, shiftText, attributes, sheetName); 
  } catch (e) {
    throw new Error(e.message);
  }
}

// ■ 常勤ウィザード用（★これを追加しました！）
function saveJokinWizardData(rowIndex, fullText, attributes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 勤務備考（シフト）の保存
    const shiftColIdx = headers.findIndex(h => h.toString().trim() === "勤務備考" || h.toString().trim() === "提案シフト");
    if (shiftColIdx === -1) throw new Error("「勤務備考」列が見つかりません。");
    sheet.getRange(rowIndex, shiftColIdx + 1).setValue(fullText);

    // 属性（祝日・年末年始）の保存
    if (attributes) {
      const holIdx = headers.findIndex(h => h.toString().includes("祝日"));
      const yeIdx = headers.findIndex(h => h.toString().includes("年末年始"));
      if (holIdx > -1) sheet.getRange(rowIndex, holIdx + 1).setValue(attributes.holiday);
      if (yeIdx > -1) sheet.getRange(rowIndex, yeIdx + 1).setValue(attributes.yearend);
    }
    
    // ステータスの更新（もしあれば）
    const statusIdx = headers.findIndex(h => h.toString().includes("対応状況") || h.toString().includes("内容修正"));
    if (statusIdx > -1) sheet.getRange(rowIndex, statusIdx + 1).setValue("修正済");
    
    return "保存完了";
  } catch (e) {
    console.error("saveJokinWizardData Error: " + e.stack);
    throw new Error(e.message);
  }
}

// ■ 常勤属性のみ保存用（★これも追加しました！）
function saveJokinAttributes(rowIndex, attributes) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    const updateVal = (keyword, val) => {
      const idx = headers.findIndex(h => h.toString().includes(keyword));
      if (idx > -1) sheet.getRange(rowIndex, idx + 1).setValue(val);
    };

    updateVal("主務", attributes.duty);
    updateVal("管理", attributes.manager);
    updateVal("祝日", attributes.holiday);
    updateVal("年末年始", attributes.yearend);
    
    return "保存完了";
  } catch (e) {
    throw new Error(e.message);
  }
}

function getGroupedLocations() { return getSettings().locations; }
function fetchLocationMap() { return getSettings().locMap; }
function fetchRateData() { return ShiftEditorLib.fetchRateData(); }
function fetchRateData2025() { return ShiftEditorLib.fetchRateData2025(); }
function getJokinDoctorList(sheetNameFromHtml) { return ShiftEditorLib.getJokinDoctorList(sheetNameFromHtml || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()); }
function getJokinDoctorDataByRow(rowIndex, sheetNameFromHtml) { return ShiftEditorLib.getJokinDoctorDataByRow(rowIndex, sheetNameFromHtml || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()); }
function backToSelector() { return ShiftEditorLib.backToSelector(); }

// ==========================================
// ⏱️ 【自動実行用マスター】トリガーにはこれをセットする
// ==========================================
function autoRunScheduledBatch() {
  console.log("▶️ 定期自動実行バッチを開始します...");
  
  try {
    // ① まず、カレンダーの色塗りと「契約/確定」の併記を行う
    if (typeof updateAttendanceActuals === 'function') {
      updateAttendanceActuals(); 
    }
    
    // ② 次に、確定した欠勤などをお休み情報シートに抽出・同期する
    if (typeof syncActualsToExceptions === 'function') {
      syncActualsToExceptions(); 
    }
    
    console.log("✅ 定期自動実行バッチがすべて正常に完了しました。");
  } catch(e) {
    console.error("❌ 自動実行中にエラーが発生しました: " + e.message);
  }
}