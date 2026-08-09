function debugTestClinicNoConversion() {
  console.log("=== 🔍 クリニックNo変換（getDropdownFormalName）のテスト開始 ===");
  
  // 1. locMasterの取得（実際の関数を使用）
  const locMaster = typeof fetchLocationMaster === 'function' ? fetchLocationMaster() : {};
  console.log("✅ locMaster (fetchLocationMaster) の取得完了。要素数: " + Object.keys(locMaster).length);

  // 2. dictの構築（実際の処理を模倣）
  let dict = {};
  try {
    const mapUrl = "https://docs.google.com/spreadsheets/d/14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs/edit";
    const ssMap = SpreadsheetApp.openByUrl(mapUrl);
    const sheetMap = ssMap.getSheetByName("拠点名");
    const dataMap = sheetMap.getDataRange().getValues();
    for (let i = 1; i < dataMap.length; i++) {
      let canonical = dataMap[i][0] ? String(dataMap[i][0]).trim() : "";
      if (!canonical) continue;
      dict[canonical] = canonical;
      for (let j = 1; j <= 4; j++) {
        let variant = dataMap[i][j] ? String(dataMap[i][j]).replace(/[\s ]+/g, "") : "";
        if (variant) dict[variant] = canonical;
      }
    }
    console.log("✅ dict の構築完了。要素数: " + Object.keys(dict).length);
  } catch(e) { 
    console.error("❌ マスタ辞書構築エラー: " + e.message); 
  }

  // 3. テスト対象の関数（今回の修正箇所）
  const getDropdownFormalName = (rawCalendarLoc) => {
    let cleanRaw = String(rawCalendarLoc).replace(/[【】\(（]?(内科|小児科)[\)）]?/g, "").replace(/\/.*/, "").replace(/[\s ]+/g, "");
    
    // ★ 修正した核心部分
    let canonicalLoc = dict[cleanRaw] || locMaster[cleanRaw] || cleanRaw;
    
    return { original: rawCalendarLoc, cleanRaw: cleanRaw, result: canonicalLoc };
  };

  // 4. テスト実行
  // 実際に画像にあった数字や、よくある表記ブレを含めてテストします
  const testCases = [
    "41", 
    "49", 
    "柏の葉", 
    "【北葛西】内科",
    "1000", // 院外勤務などの特殊コード
    "8"     // 東品川の可能性
  ];

  console.log("\n=== 🧪 テスト結果 ===");
  testCases.forEach(testVal => {
    let res = getDropdownFormalName(testVal);
    let status = "";
    
    if (res.result !== res.original && isNaN(parseInt(res.result, 10))) {
      status = "⭕️ 変換成功";
    } else if (res.result === res.original && isNaN(parseInt(res.result, 10))) {
      status = "✅ 正常な拠点名";
    } else {
      status = "⚠️ 未変換(数字のまま)";
    }
    
    console.log(`[${status}] 入力: "${res.original}" -> クリーン: "${res.cleanRaw}" -> 出力: "${res.result}"`);
  });
  
  console.log("======================================================");
}