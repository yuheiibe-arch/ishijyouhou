// ==========================================
// 【常勤】一括処理セット（ラッパー）
// ファイル名：FullTimeAttendance.gs
// 依存：AttendanceCore.gs
// ==========================================

function processFullTime_AllSteps(targetYear) {
  generateFullTimeMaster_Dynamic(targetYear);
  generateFullTimeAttendance_Dynamic(targetYear);
}

function generateFullTimeMaster_Dynamic(targetYear) {
  generateMaster_Core(targetYear, "常勤");
}

function generateFullTimeAttendance_Dynamic(targetYear) {
  generateAttendance_Core(targetYear, "常勤");
}

// ★修正：contractText を引数として受け取り、Core関数へ渡すように変更
function appendFullTimeDoctor_Single(targetYear, headers, rowData, contractText) {
  appendDoctor_Single_Core(targetYear, headers, rowData, "常勤", contractText);
}