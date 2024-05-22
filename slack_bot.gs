/*
  メインの処理
  setTriggerDay() : 最終営業日(0時)に次の処理をする命令を出す。
  setTriggerHoursLast() : 最終営業日の18時に次の命令を出す。
  sendSlack() : slackにデータを投げる。

  サブ的な処理
  deleteTrigger() : 蓄積されている命令を削除する
  lastBusinessDay() : 今月の最終営業日を求める
  isHoliday() : 本日が[日本の祝日]かどうかチェック。土日は曜日で判定してるチェック
  .getDay()　: Dateオブジェクトから曜日を求めるメソッド(0:日, 6:土曜日)
*/

function setTriggerDay()
{  
  var last = lastBusinessDay();
  ScriptApp.newTrigger("setTriggerHoursLast")
    .timeBased()
    .atDate(last.getFullYear(), last.getMonth()+1, last.getDate())
    .create();
}

function setTriggerHoursLast()
{
  deleteTrigger("setTriggerHoursLast");
  ScriptApp.newTrigger("sendSlack")
    .timeBased()
    .after( 8 * 60 * 60 * 1000 )
    .create();
}

function sendSlack() 
{
  deleteTrigger("sendSlack");
  var options = 
  {
    "method" : "POST",
    "payload" : 
    {
      "token": "???",
      "channel": "入居者管理",
      "text": getinfofromspreadsheet()
    }
  }
  var url = "https://slack.com/api/chat.postMessage"
  UrlFetchApp.fetch(url, options);
}

function lastBusinessDay() 
{
  var today = new Date();

  var lastDayOfThisMonth = new Date(today.getFullYear(), today.getMonth()+1, -5);
  var day; // 0->日曜日

  for (var i = 0; i < 30; i++) {
    day = lastDayOfThisMonth.getDay();
    if (day == 0 || day == 6 || isHoliday(lastDayOfThisMonth)) {
      lastDayOfThisMonth = new Date(today.getFullYear(), today.getMonth()+1, -5 + i);
      continue;
    }
  }
  return lastDayOfThisMonth;
}

function deleteTrigger(name) 
{
  var triggers = ScriptApp.getProjectTriggers();
  for(var i=0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == name) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

function isHoliday(day) 
{
  var startDate = new Date(day.setHours(0, 0, 0, 0));
  var endDate = new Date(day.setHours(23, 59, 59));
  var cal = CalendarApp.getCalendarById("ja.japanese#holiday@group.v.calendar.google.com");
  var holidays =  cal.getEvents(startDate, endDate);
  return holidays.length != 0; // 祝日ならtrue
}

function testPost() {
  const app_auth_token = "???";
  const channel = "入居者管理";

  const result = postMessageToSlackChannel(
    app_auth_token, 
    channel, 
    getinfofromspreadsheet()
  );
  Logger.log(result);
}

/**
 * messageをSlackチャンネルにポストする関数
 */
function postMessageToSlackChannel(app_auth_token, channel, message){
  const payload = {
    "token" : app_auth_token,
    "channel" : channel,
    "text" : message
  };
  const options = {
    "method" : "post",
    'contentType': 'application/x-www-form-urlencoded',
    "payload" : payload
  };

  return UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", options);
}

function getinfofromspreadsheet(){
  const SHEET_ID = "1FvBdSHriglfTJ0cOGplPAUahwPU6xK-VnQTsNo5Fapk";
  const SHEET_NAME = "入居者情報";
  var text = "";
  var spreadSheet = SpreadsheetApp.openById(SHEET_ID);
  var sheet = spreadSheet.getSheetByName(SHEET_NAME);
  var date = new Date();
  var year = date.getFullYear();
  var month = date.getMonth() + 1;
  var d2 = date.getDate();

  var now = year+"/"+month+"/"+d2;

  var text = "@Yuki ISE" + "\n" +"通知者リスト (" + now + " 現在)";
  var checker = 0;

  for (var i = 2; i <= 7; i++){
    var name = sheet.getRange(i, 2).getValue();
    var end = sheet.getRange(i, 4).getValue();
    var y = end.getFullYear();
    if (sheet.getRange(i, 6).getValue() === "shorter than a year"){
      var six_before = "9999/99/99";
    }
    else{
      var six_before = sheet.getRange(i, 6).getValue()
      if (sheet.getRange(i, 8).getValue() === year && sheet.getRange(i, 9).getValue() === month){
        var ending = Utilities.formatDate(end, "JST", "MM/dd")
        var text = text + "\n" + name + "さんの契約終了六カ月前が迫っています。 (契約終了日: " + y + "/" + ending + ")";
        var checker = checker + 1;
      };
    };
    if (sheet.getRange(i, 10).getValue() === year && sheet.getRange(i, 11).getValue() === month){
      var ending = Utilities.formatDate(end, "JST", "MM/dd")
      var text = text + "\n" + name + "さんの契約終了二か月前が迫っています。 (契約終了日: " + y + "/" + ending + ")";
      var checker = checker + 1;
    };
  };

  if (checker === 0){
    text = text + "\n" + "直近で通知しなければいけない入居者はいません。";
  }
  return text;
}
