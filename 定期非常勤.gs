// ==========================================
// 【定期非常勤】一括処理セット（ラッパー）
// ファイル名：PartTimeAttendance.gs
// 依存：AttendanceCore.gs
// ==========================================

function processPartTime_AllSteps(targetYear) {
  generatePartTimeMaster_Dynamic(targetYear);
  generatePartTimeAttendance_Dynamic(targetYear);
}

function generatePartTimeMaster_Dynamic(targetYear) {
  generateMaster_Core(targetYear, "定期非常勤");
}

function generatePartTimeAttendance_Dynamic(targetYear) {
  generateAttendance_Core(targetYear, "定期非常勤");
}

// ★修正：contractText を引数として受け取り、Core関数へ渡すように変更
function appendPartTimeDoctor_Single(targetYear, headers, rowData, contractText) {
  appendDoctor_Single_Core(targetYear, headers, rowData, "定期非常勤", contractText);
}