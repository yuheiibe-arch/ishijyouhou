function debugDoctor452062() {

  const sourceSheet = SpreadsheetApp
    .openById("1miSCaOe0vZhWFqFHXBj6rvspebRi-53L0SzawjKgWm0")
    .getSheetByName("2025年度：医師情報");

  const data = sourceSheet.getDataRange().getValues();

  Logger.log("=== 452062 検索開始 ===");

  data.forEach((row, i) => {
    const doctorNo = String(row[3]).trim(); // D列

    if (doctorNo === "452062") {
      Logger.log(
        `行:${i + 1} D=${row[3]} O=${row[14]}`
      );
    }
  });

  Logger.log("=== 検索終了 ===");
}