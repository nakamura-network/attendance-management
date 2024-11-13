/**
 * ページを開いた時に最初に呼ばれるルートメソッド
 */
function doGet(e) {
  var page = e.parameter.page;
  var selectedEmpId = e.parameter.empId
  if (page == 'shift') {
    return HtmlService.createTemplateFromFile("view_shift")
      .evaluate().setTitle("ShiftTable");
  }
  if (selectedEmpId == undefined) {
    // empIdがセットされていない場合にはホーム画面を表示
    return HtmlService.createTemplateFromFile("view_home")
      .evaluate().setTitle("Home")
  }
  // 選択した従業員IDを後続の処理でも利用するためにPropertyに保存
    PropertiesService.getUserProperties().setProperty('selectedEmpId', selectedEmpId.toString())
  // 従業員の詳細画面を表示
  return HtmlService.createTemplateFromFile("view_detail")
      .evaluate().setTitle("Detail: " + selectedEmpId.toString())
}

/**
 * このアプリのURLを返す
 */
function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

function sendReasonEmail(targetDate, shiftTime, targetTime, reason) {
  var to = "kou011120@gmail.com";  // 送信先のメールアドレス(管理者のメールアドレス)
  var subject = "乖離時間の理由";
  var empName = getEmployeeName();
  var body = "従業員名: " + empName + "\n\n" +
             "シフトとの乖離時間が15分を超えました。理由は以下の通りです。\n\n" +
             "日付: " + targetDate + "\n" +
             "シフト予定時刻: " + shiftTime + "\n" +
             "打刻時刻: " + targetTime + "\n" +
             "理由: " + reason;

  // メールを送信
  GmailApp.sendEmail(to, subject, body);
}

/**
 * 従業員一覧
 */
function getEmployees() {
  var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1]// 「従業員名簿」のシート
  var last_row = empSheet.getLastRow()
  var empRange = empSheet.getRange(2, 1, last_row, 2);// シートの中のヘッダーを除く範囲を取得
  var employees = [];
  var i = 1;
  while (true) {
    var empId =empRange.getCell(i, 1).getValue();
    var empName =empRange.getCell(i, 2).getValue();
    if (empId === ""){ //　値を取得できなくなったら終了
      break;
    }
    employees.push({
      'id': empId,
      'name': empName
    })
    i++
  }
  return employees
}

function getShift() {
  var shiftSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[4]// 「シフト」のシート
  var last_row = shiftSheet.getLastRow()
  var last_column = shiftSheet.getLastColumn()
  var shiftRange = shiftSheet.getRange(1, 1, last_row, last_column);// シートの中のヘッダーを除く範囲を取得
  var shift = shiftRange.getValues();
  /*for(var i=0; i<last_column; i++){
    var cellValue = shift[0][i];
    cellValue = formatDate(cellValue);
    shift[0][i] = cellValue;
  }*/
  shift[0][0] = "従業員名";
  for(var j=0; j<last_row-1; j++){
    var num = shift[j+1][0];
    num = getShiftEmployeeName(num);
    shift[j+1][0] = num;
  }
  return shift
}

function formatDate(dateString) {
    var date = new Date(dateString);
    return date.toLocaleDateString('ja-JP'); // "年/月/日"形式に変換
  }

/**
 * 従業員情報の取得
 * ※ デバッグするときにはselectedEmpIdを存在するIDで書き換えてください
 */
function getEmployeeName() {
  var selectedEmpId =PropertiesService.getUserProperties().getProperty('selectedEmpId') // ※デバッグするにはこの変数を直接書き換える必要があります
  var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1]// 「従業員名簿」のシート
  var last_row = empSheet.getLastRow()
  var empRange = empSheet.getRange(2, 1, last_row, 2);// シートの中のヘッダーを除く範囲を取得
  var i = 1;
  var empName = ""
  while (true) {
    var id =empRange.getCell(i, 1).getValue();
    var name =empRange.getCell(i, 2).getValue();
    if (id === ""){ 
      break;
    }
    if(id == selectedEmpId){
      empName = name
    }
    i++
  }

  return empName
}

function getShiftEmployeeName(empId) {
  var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1]// 「従業員名簿」のシート
  var last_row = empSheet.getLastRow()
  var empRange = empSheet.getRange(2, 1, last_row, 2);// シートの中のヘッダーを除く範囲を取得
  var i = 1;
  var empName = ""
  while (true) {
    var id =empRange.getCell(i, 1).getValue();
    var name =empRange.getCell(i, 2).getValue();
    if (id === ""){ 
      break;
    }
    if(id == empId){
      empName = name
    }
    i++
  }
  return empName
}

/**
 * 勤怠情報の取得
 * 今月における今日までの勤怠情報が取得される
 */
function getTimeClocks() {
  var selectedEmpId =PropertiesService.getUserProperties().getProperty('selectedEmpId') // ※デバッグするにはこの変数を直接書き換える必要があります
  var timeClocksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[2]// 「打刻履歴」のシート
  var last_row = timeClocksSheet.getLastRow()
  var timeClocksRange = timeClocksSheet.getRange(2, 1, last_row, 3);// シートの中のヘッダーを除く範囲を取得
  var empTimeClocks = [];
  var i = 1;
  while (true) {
    var empId =timeClocksRange.getCell(i, 1).getValue();
    var type =timeClocksRange.getCell(i, 2).getValue();
    var datetime =timeClocksRange.getCell(i, 3).getValue();
    if (empId === ""){ //　値を取得できなくなったら終了
      break;
    }
    if (empId == selectedEmpId){
      empTimeClocks.push({
        'date': Utilities.formatDate(datetime, "Asia/Tokyo", "yyyy-MM-dd HH:mm"),
        'type': type
    })
    }
    i++
  }
  return empTimeClocks
}

/**
 * 勤怠情報登録
 */
function saveWorkRecord(form) {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId') // ※デバッグするにはこの変数を直接書き換える必要があります
  // inputタグのnameで取得
  var targetDate = form.target_date
  var targetTime = form.target_time
  var targetType = ''

  var currentStatus = getEmployeeStatus();

  // 既に「勤務中」の場合はエラーメッセージを返す
  if (form.target_type == 'clock_in' && currentStatus == '勤務中') {
    return 'エラー: 既に勤務中です。';
  }
  if (form.target_type == 'break_end' && currentStatus == '勤務中') {
    return 'エラー: 現在は休憩中ではありません。';
  }
  if (form.target_type == 'clock_out' && currentStatus == '退勤済') {
    return 'エラー: 既に退勤済です。';
  }
  if (form.target_type == 'break_begin' && currentStatus == '退勤済') {
    return 'エラー: 既に退勤済です。';
  }
  if (form.target_type == 'break_end' && currentStatus == '退勤済') {
    return 'エラー: 既に退勤済です。';
  }
  if (form.target_type == 'break_begin' && currentStatus == '休憩中') {
    return 'エラー: 既に休憩中です。';
  }
  if (form.target_type == 'clock_in' && currentStatus == '休憩中') {
    return 'エラー: 現在は休憩中です。';
  }
  if (form.target_type == 'clock_out' && currentStatus == '休憩中') {
    return 'エラー: 休憩を終了してから退勤してください。';
  }
  if (form.target_type == 'clock_out' && checkTime(3, form)) {
    return 'エラー: 退勤時刻が不適切です。';
  }
  if (form.target_type == 'break_end' && checkTime(5, form)) {
    return 'エラー: 休憩終了時刻が不適切です。';
  }

  switch (form.target_type) {
    case 'clock_in':
      saveClockInRecord(form)
      targetType = '出勤'
      break
    case 'break_begin':
      saveBreakBeginRecord(form)
      targetType = '休憩開始'
      break
    case 'break_end':
      saveBreakEndRecord(form)
      targetType = '休憩終了'
      break
    case 'clock_out':
      saveClockOutRecord(form)
      targetType = '退勤'
      break;
  }
  var timeClocksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[2]// 「打刻履歴」のシート
  var targetRow = timeClocksSheet.getLastRow() + 1
  timeClocksSheet.getRange(targetRow, 1).setValue(selectedEmpId)
  timeClocksSheet.getRange(targetRow, 2).setValue(targetType)
  timeClocksSheet.getRange(targetRow, 3).setValue(targetDate + ' ' + targetTime)
  return '登録しました'
}

/**
 * 選択している従業員のメモカラムの値をspread sheetから取得する
 */
function getEmpMemo() {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId') // ※デバッグするにはこの変数を直接書き換える必要があります
  var checkSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]// 「チェック結果」のシート
  var last_row = checkSheet.getLastRow()
  var timeClocksRange = checkSheet.getRange(2, 1, last_row, 2);// シートの中のヘッダーを除く範囲を取得
  var checkResult = "";
  var i = 1;
  while (true) {
    var empId =timeClocksRange.getCell(i, 1).getValue();
    var result =timeClocksRange.getCell(i, 2).getValue();
    if (empId === ""){ //　値を取得できなくなったら終了
      break;
    }
    if (empId == selectedEmpId){
        checkResult = result
        break;
    }
    i++
  }
  return checkResult
}

/**
 * メモの内容をSpreadSheetに保存する
 */
function saveMemo(form) {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId') // ※デバッグするにはこの変数を直接書き換える必要があります
  // inputタグのnameで取得
  var memo = form.memo

  var targetRowNumber = getTargetEmpRowNumber(selectedEmpId)
  var sheet = SpreadsheetApp.getActiveSheet()
  if (targetRowNumber == null) {
    // targetRowNumberがない場合には新規に行を追加する
    // 現在の最終行に+1した行番号
    targetRowNumber = sheet.getLastRow() + 1
    // 1列目にempIdをセットして保存
    sheet.getRange(targetRowNumber, 1).setValue(selectedEmpId)
  }
  // memoの内容を保存
  var values = sheet.getRange(targetRowNumber, 2).setValue(memo)

}

/**
 * spreadSheetに保存されている指定のemployee_idの行番号を返す
 */
function getTargetEmpRowNumber(empId) {
  // 開いているシートを取得
  var sheet = SpreadsheetApp.getActiveSheet()
  // 最終行取得
  var last_row = sheet.getLastRow()
  // 2行目から最終行までの1列目(emp_id)の範囲を取得
  var data_range = sheet.getRange(1, 1, last_row, 1);
  // 該当範囲のデータを取得
  var sheetRows = data_range.getValues();
  // ループ内で検索
  for (var i = 0; i <= sheetRows.length - 1; i++) {
    var row = sheetRows[i]
    if (row[0] == empId) {
      // spread sheetの行番号は1から始まるが配列のindexは0から始まるため + 1して行番号を返す
      return i + 1;
    }
  }
  // 見つからない場合にはnullを返す
  return null
}

function getEmployeeStatus() {
  var selectedEmpId =PropertiesService.getUserProperties().getProperty('selectedEmpId')
  var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[2]// 「打刻履歴」のシート
  var last_row = empSheet.getLastRow()
  var empRange = empSheet.getRange(2, 1, last_row, 3);
  var i = 1;
  var empStatus = ""
  while (true) {
    var id =empRange.getCell(i, 1).getValue();
    var status =empRange.getCell(i, 2).getValue();
    if (id === ""){ 
      break;
    }
    if(id == selectedEmpId){
      empStatus = status
    }
    i++
  }
  if(empStatus == "出勤"){
    empStatus = "勤務中"
  }
  if(empStatus == "退勤"){
    empStatus = "退勤済"
  }
  if(empStatus == "休憩開始"){
    empStatus = "休憩中"
  }
  if(empStatus == "休憩終了"){
    empStatus = "勤務中"
  }
  return empStatus
}

function homeGetEmployeeStatus(empId) {
  var selectedEmpId = empId
  var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[2]// 「打刻履歴」のシート
  var last_row = empSheet.getLastRow()
  var empRange = empSheet.getRange(2, 1, last_row, 3);
  var i = 1;
  var empStatus = ""
  while (true) {
    var id =empRange.getCell(i, 1).getValue();
    var status =empRange.getCell(i, 2).getValue();
    if (id === ""){ 
      break;
    }
    if(id == selectedEmpId){
      empStatus = status
    }
    i++
  }
  if(empStatus == "出勤"){
    empStatus = "勤務中"
  }
  if(empStatus == "退勤"){
    empStatus = "退勤済"
  }
  if(empStatus == "休憩開始"){
    empStatus = "休憩中"
  }
  if(empStatus == "休憩終了"){
    empStatus = "勤務中"
  }
  return empStatus
}

function saveClockInRecord(form){
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId')
  var recordSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[3]
  var targetDate = form.target_date
  var targetTime = form.target_time
  
  var targetRow = recordSheet.getLastRow() + 1
  recordSheet.getRange(targetRow, 1).setValue(selectedEmpId)
  recordSheet.getRange(targetRow, 2).setValue(targetDate)
  recordSheet.getRange(targetRow, 3).setValue(targetTime)
}

function saveClockOutRecord(form){
  var selectedEmpId =PropertiesService.getUserProperties().getProperty('selectedEmpId')
  var recordSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[3]
  var targetDate = form.target_date
  var targetTime = form.target_time
  var last_row = recordSheet.getLastRow()
  var empRange = recordSheet.getRange(2, 1, last_row, 9);
  var i = 1;
  while (true) {
    var id =empRange.getCell(i, 1).getValue();
    var workDate =empRange.getCell(i, 2).getDisplayValue();
    var clockInTime = empRange.getCell(i, 3).getDisplayValue();
    if (id === ""){ 
      break;
    }
    if(id == selectedEmpId){
      if(workDate == targetDate){
        empRange.getCell(i, 4).setValue(targetTime)
        var workTime = calculateTimeRange(empRange.getCell(i, 3).getValue(), empRange.getCell(i, 4).getValue())
        empRange.getCell(i, 7).setValue(workTime)
        var actualWorkTime = empRange.getCell(i, 7).getValue() - empRange.getCell(i, 8).getValue()
        empRange.getCell(i, 9).setValue(actualWorkTime)
      }
    }
    i++
  }
}

function saveBreakBeginRecord(form){
  var selectedEmpId =PropertiesService.getUserProperties().getProperty('selectedEmpId')
  var recordSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[3]
  var targetDate = form.target_date
  var targetTime = form.target_time
  var last_row = recordSheet.getLastRow()
  var empRange = recordSheet.getRange(2, 1, last_row, 7);
  var i = 1;
  while (true) {
    var id =empRange.getCell(i, 1).getValue();
    var workDate =empRange.getCell(i, 2).getDisplayValue();
    if (id === ""){ 
      break;
    }
    if(id == selectedEmpId){
      if(workDate == targetDate){
        empRange.getCell(i, 5).setValue(targetTime)
      }
    }
    i++
  }
}

function saveBreakEndRecord(form){
  var selectedEmpId =PropertiesService.getUserProperties().getProperty('selectedEmpId')
  var recordSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[3]
  var targetDate = form.target_date
  var targetTime = form.target_time
  var last_row = recordSheet.getLastRow()
  var empRange = recordSheet.getRange(2, 1, last_row, 9);
  var i = 1;
  while (true) {
    var id =empRange.getCell(i, 1).getValue();
    var workDate =empRange.getCell(i, 2).getDisplayValue();
    if (id === ""){ 
      break;
    }
    if(id == selectedEmpId){
      if(workDate == targetDate){
        empRange.getCell(i, 6).setValue(targetTime)
        var breakTime = calculateTimeRange(empRange.getCell(i, 5).getValue(), empRange.getCell(i, 6).getValue())
        empRange.getCell(i, 8).setValue(breakTime)
      }
    }
    i++
  }
}

function calculateTimeRange(startTime, endTime) {

  startTime = new Date(startTime).getTime();
  endTime = new Date(endTime).getTime();

  const diff = Math.abs(endTime - startTime);

  var diff_hours = diff / 1000 / 60 / 60;

  diff_hours = Math.round(diff_hours * 100) / 100
  console.log(diff_hours)
  return(diff_hours)
}

function getTimeRecord() {
  var selectedEmpId =PropertiesService.getUserProperties().getProperty('selectedEmpId') 
  var timeClocksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[3]// 「打刻履歴」のシート
  var last_row = timeClocksSheet.getLastRow()
  var timeClocksRange = timeClocksSheet.getRange(2, 1, last_row, 9);// シートの中のヘッダーを除く範囲を取得
  var empTimeRecord = [];
  var i = 1;
  while (true) {
    var empId =timeClocksRange.getCell(i, 1).getValue();
    var datetime =timeClocksRange.getCell(i, 2).getDisplayValue();
    var clockInTime =timeClocksRange.getCell(i, 3).getDisplayValue();
    var clockOutTime =timeClocksRange.getCell(i, 4).getDisplayValue();
    var breakBeginTime =timeClocksRange.getCell(i, 5).getDisplayValue();
    var breakEndTime =timeClocksRange.getCell(i, 6).getDisplayValue();
    var workTime =timeClocksRange.getCell(i, 9).getDisplayValue();
    if (empId === ""){ //　値を取得できなくなったら終了
      break;
    }
    if (empId == selectedEmpId){
      empTimeRecord.push({
        'date': datetime,
        'clockInTime': clockInTime,
        'clockOutTime': clockOutTime,
        'breakBeginTime': breakBeginTime,
        'breakEndTime': breakEndTime,
        'workTime' : workTime
    })
    }
    i++
  }
  return empTimeRecord
}

function validateTime(startTime, endTime) {
    var start = new Date("2024-01-01 "+startTime).getTime();
    var end = new Date("2024-01-01 "+endTime).getTime();
    console.log(start+"=start, "+end+"=end")
    if(start < end){
      return false
    }
    return true
}

function checkTime(column, form) {
  var selectedEmpId =PropertiesService.getUserProperties().getProperty('selectedEmpId') 
  var timeClocksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[3]// 「打刻履歴」のシート
  var last_row = timeClocksSheet.getLastRow()
  var targetDate = form.target_date
  var targetTime = form.target_time
  var timeClocksRange = timeClocksSheet.getRange(2, 1, last_row, 7);// シートの中のヘッダーを除く範囲を取得
  var i = 1;
  while (true) {
    var empId =timeClocksRange.getCell(i, 1).getValue();
    var datetime =timeClocksRange.getCell(i, 2).getDisplayValue();
    var clockInTime =timeClocksRange.getCell(i, column).getDisplayValue();
    if (empId === ""){ //　値を取得できなくなったら終了
      break;
    }
    if (empId == selectedEmpId){
      if(datetime == targetDate){
        console.log(clockInTime+", "+targetTime)
        return validateTime(clockInTime, targetTime)
      }
    }
    i++
  }
}

function checkShift(shiftDate){
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId') 
  var shiftSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[4];
  var targetDate = shiftDate
  var last_row = shiftSheet.getLastRow()
  var last_column = shiftSheet.getLastColumn()
  var shiftRange = shiftSheet.getRange(2, 1, last_row, last_column);
  var shiftRange2 = shiftSheet.getRange(1, 1, last_row, last_column);
  var i = 1;
  while (true) {
    var empId =shiftRange.getCell(i, 1).getValue();
    if (empId === ""){
      break;
    }
    if(empId == selectedEmpId){
      var j = 1;
      while(true){
      var date =shiftRange2.getCell(1, j).getDisplayValue();
      console.log("date = "+targetDate)
      if(date == ""){
        break;
      }
      if(date == targetDate){
        var shift = shiftRange2.getCell(i+1, j).getValue();
        console.log(shift)
        return shift
      }
      j++
    }
    }
    i++
  }
}

function clockInTimeSplit(timeRange){
  var result = timeRange.split("-");
  var shiftClockIn = result[0]
  return shiftClockIn
}

function clockOutTimeSplit(timeRange){
  var result = timeRange.split("-");
  var shiftClockOut = result[1]
  return shiftClockOut
}

function calculateShift(recordTime, time){
  var targetTime = recordTime
  console.log("form="+targetTime+"shiftTime="+time)
  var workTime = new Date("2024-01-01 "+targetTime).getTime();
  var shiftTime = new Date("2024-01-01 "+time).getTime();
  console.log("workTime="+workTime+"shiftTime="+shiftTime)
  const diff = Math.abs(workTime - shiftTime);
  var diff_minute = diff / 1000 / 60;
  console.log("minute= "+ diff_minute)
  return diff_minute
}

// 年を返す関数
function getYear(date) {
  return date.getFullYear();
}

// 月を返す関数
function getMonth(date) {
  return date.getMonth() + 1;
}

// 日を返す関数
function getDay(date) {
  return date.getDate(); 
}

function getDayOfWeek(date) {
  var daysOfWeek = ['日', '月', '火', '水', '木', '金', '土'];
  return daysOfWeek[date.getDay()]; 
}
