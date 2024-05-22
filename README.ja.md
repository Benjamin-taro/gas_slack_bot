# Slack通知スクリプト

このGoogle Apps Scriptは、毎月の最終営業日にSlackチャンネルに通知を送信します。契約終了が近づいている入居者にリマインダーを送信します。

## 目次

- [セットアップ](#セットアップ)
- [使用方法](#使用方法)
- [関数](#関数)
  - [setTriggerDay](#settriggerday)
  - [setTriggerHoursLast](#settriggerhourslast)
  - [sendSlack](#sendslack)
  - [deleteTrigger](#deletetrigger)
  - [lastBusinessDay](#lastbusinessday)
  - [isHoliday](#isholiday)
  - [testPost](#testpost)
  - [postMessageToSlackChannel](#postmessagetoslackchannel)
  - [getinfofromspreadsheet](#getinfofromspreadsheet)

## セットアップ

1. **Google Apps Scriptプロジェクト:**
   - 新しいGoogle Apps Scriptプロジェクトを作成します。
   - 提供されたスクリプトをプロジェクトにコピー＆ペーストします。

2. **GoogleカレンダーAPI:**
   - プロジェクトが祝日を確認するためにGoogleカレンダーAPIにアクセスできることを確認します。

3. **Slack APIトークン:**
   - `sendSlack` 関数内のプレースホルダー `token` をあなたのSlackボットトークンに置き換えます。
   - `channel` を適切なSlackチャンネルIDに置き換えます。

4. **Googleスプレッドシート:**
   - `getinfofromspreadsheet` 内のGoogleスプレッドシートIDとシート名があなたのスプレッドシート設定と一致していることを確認します。

## 使用方法

1. 毎月の初めに `setTriggerDay` を実行するように時間ベースのトリガーとしてスクリプトをデプロイします。
2. スクリプトは自動的に最終営業日とその日の18:00に通知を送信するトリガーを作成します。

## 関数

### setTriggerDay

月の最終営業日に `setTriggerHoursLast` 関数を実行するトリガーを設定します。

```javascript
function setTriggerDay() {
  var last = lastBusinessDay();
  ScriptApp.newTrigger("setTriggerHoursLast")
    .timeBased()
    .atDate(last.getFullYear(), last.getMonth() + 1, last.getDate())
    .create();
}
setTriggerHoursLast
既存の setTriggerHoursLast のトリガーを削除し、最終営業日の18:00に sendSlack 関数を実行する新しいトリガーを設定します。

javascript
Copy code
function setTriggerHoursLast() {
  deleteTrigger("setTriggerHoursLast");
  ScriptApp.newTrigger("sendSlack")
    .timeBased()
    .after(8 * 60 * 60 * 1000)
    .create();
}
sendSlack
既存の sendSlack のトリガーを削除し、スプレッドシートから情報を取得してSlackに通知を送信します。

javascript
Copy code
function sendSlack() {
  deleteTrigger("sendSlack");
  var options = {
    "method": "POST",
    "payload": {
      "token": "xoxb-XXXXXXXXXX-XXXXXXXXXX-XXXXXXXXXX",
      "channel": "your-channel-id",
      "text": getinfofromspreadsheet()
    }
  };
  var url = "https://slack.com/api/chat.postMessage";
  UrlFetchApp.fetch(url, options);
}
deleteTrigger
指定されたハンドラ関数名のトリガーをすべて削除します。

javascript
Copy code
function deleteTrigger(name) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == name) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
lastBusinessDay
今月の最終営業日を計算して返します。

javascript
Copy code
function lastBusinessDay() {
  var today = new Date();
  var lastDayOfThisMonth = new Date(today.getFullYear(), today.getMonth() + 1, -5);
  var day; // 0->Sunday

  for (var i = 0; i < 30; i++) {
    day = lastDayOfThisMonth.getDay();
    if (day == 0 || day == 6 || isHoliday(lastDayOfThisMonth)) {
      lastDayOfThisMonth = new Date(today.getFullYear(), today.getMonth() + 1, -5 + i);
      continue;
    }
  }
  return lastDayOfThisMonth;
}
isHoliday
指定された日が日本の祝日かどうかを確認します。

javascript
Copy code
function isHoliday(day) {
  var startDate = new Date(day.setHours(0, 0, 0, 0));
  var endDate = new Date(day.setHours(23, 59, 59));
  var cal = CalendarApp.getCalendarById("ja.japanese#holiday@group.v.calendar.google.com");
  var holidays = cal.getEvents(startDate, endDate);
  return holidays.length != 0; // 祝日ならtrue
}
testPost
Slackにサンプルメッセージを送信するテスト関数。

javascript
Copy code
function testPost() {
  const app_auth_token = "xoxb-XXXXXXXXXX-XXXXXXXXXX-XXXXXXXXXX";
  const channel = "your-channel-id";
  const result = postMessageToSlackChannel(app_auth_token, channel, getinfofromspreadsheet());
  Logger.log(result);
}
postMessageToSlackChannel
指定されたSlackチャンネルにメッセージを送信します。

javascript
Copy code
function postMessageToSlackChannel(app_auth_token, channel, message) {
  const payload = {
    "token": app_auth_token,
    "channel": channel,
    "text": message
  };
  const options = {
    "method": "post",
    'contentType': 'application/x-www-form-urlencoded',
    "payload": payload
  };
  return UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", options);
}
getinfofromspreadsheet
Googleスプレッドシートから情報を取得し、Slackメッセージ用にフォーマットします。

javascript
Copy code
function getinfofromspreadsheet() {
  const SHEET_ID = "your-spreadsheet-id";
  const SHEET_NAME = "your-sheet-name";
  var text = "";
  var spreadSheet = SpreadsheetApp.openById(SHEET_ID);
  var sheet = spreadSheet.getSheetByName(SHEET_NAME);
  var date = new Date();
  var year = date.getFullYear();
  var month = date.getMonth() + 1;
  var d2 = date.getDate();

  var now = year + "/" + month + "/" + d2;

  text = "@Yuki ISE" + "\n" + "通知者リスト (" + now + " 現在)";
  var checker = 0;

  for (var i = 2; i <= 7; i++) {
    var name = sheet.getRange(i, 2).getValue();
    var end = sheet.getRange(i, 4).getValue();
    var y = end.getFullYear();
    if (sheet.getRange(i, 6).getValue() === "shorter than a year") {
      var six_before = "9999/99/99";
    } else {
      var six_before = sheet.getRange(i, 6).getValue();
      if (sheet.getRange(i, 8).getValue() === year && sheet.getRange(i, 9).getValue() ===月) {
        var ending = Utilities.formatDate(end, "JST", "MM/dd");
        text = text + "\n" + name + "さんの契約終了六カ月前が迫っています。 (契約終了日: " + y + "/" + ending + ")";
        checker = checker + 1;
      }
    }
    if (sheet.getRange(i, 10).getValue() === year && sheet.getRange(i, 11).getValue() ===月) {
      var ending = Utilities.formatDate(end, "JST", "MM/dd");
      text = text + "\n" + name + "さんの契約終了二か月前が迫っています。 (契約終了日: " + y + "/" + ending + ")";
      checker = checker + 1;
    }
  }

  if (checker === 0) {
    text = text + "\n" + "直近で通知しなければいけない入居者はいません。";
  }
  return text;
}
