function debugExtractKintaiParser() {
  console.log("=== 🔍 勤怠システム側のパーサー（抽出ロジック）を探します ===");
  const globalObj = this;
  let count = 0;

  for (let key in globalObj) {
    if (typeof globalObj[key] === 'function') {
      if (key.includes("debug")) continue;

      let funcStr = "";
      try { 
        funcStr = globalObj[key].toString(); 
      } catch(e) { 
        continue; 
      }

      // 備考欄から「拠点【】」や「第○週」「曜日」を抽出している関数を探す
      let isParser = false;
      
      // よくあるパーサーの特徴
      if (funcStr.includes("【") && (funcStr.includes("曜") || funcStr.includes("第") || funcStr.includes("weeks") || funcStr.includes("wNum"))) {
        isParser = true;
      }
      // 関数名に parse や extract が入っている場合
      if (key.toLowerCase().includes("parse") || key.toLowerCase().includes("extract")) {
        isParser = true;
      }
      // 先ほど除外した fetchKintaiData などの読み取り系はスキップ
      if (key === "fetchKintaiData" || key === "autoSyncManualExceptions" || key === "syncManualExceptions") {
        isParser = false; 
      }

      if (isParser) {
        console.log(`\n// =========================================`);
        console.log(`// 関数名: ${key}`);
        console.log(`// =========================================`);
        console.log(funcStr);
        count++;
      }
    }
  }

  if (count > 0) {
    console.log(`\n🚨 【合計 ${count} 件】の怪しい関数を抽出しました。コピーして教えてください！`);
  } else {
    console.log("⚠️ 該当する関数が見つかりませんでした。別のファイルに分かれている可能性があります。");
  }
}