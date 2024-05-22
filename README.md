# Slack Notification Script

This Google Apps Script sends notifications to a Slack channel on the last business day of each month. It checks for upcoming lease expirations and sends reminders accordingly.

[日本語版はこちら](README.ja.md)

## Table of Contents

- [Setup](#setup)
- [Usage](#usage)
- [Functions](#functions)
  - [setTriggerDay](#settriggerday)
  - [setTriggerHoursLast](#settriggerhourslast)
  - [sendSlack](#sendslack)
  - [deleteTrigger](#deletetrigger)
  - [lastBusinessDay](#lastbusinessday)
  - [isHoliday](#isholiday)
  - [testPost](#testpost)
  - [postMessageToSlackChannel](#postmessagetoslackchannel)
  - [getinfofromspreadsheet](#getinfofromspreadsheet)

## Setup

1. **Google Apps Script Project:**
   - Create a new Google Apps Script project.
   - Copy and paste the provided script into the project.

2. **Google Calendar API:**
   - Ensure the project has access to the Google Calendar API to check for holidays.

3. **Slack API Token:**
   - Replace the placeholder `token` in the `sendSlack` function with your Slack bot token.
   - Replace the `channel` with the appropriate Slack channel ID.

4. **Google Sheets:**
   - Ensure the Google Sheets ID and sheet name in `getinfofromspreadsheet` match your spreadsheet setup.

## Usage

1. Deploy the script as a time-based trigger to run `setTriggerDay` at the beginning of each month.
2. The script will automatically create triggers for the last business day and at 18:00 on that day to send notifications.

## Functions

### setTriggerDay

Sets a trigger to run the `setTriggerHoursLast` function on the last business day of the month.

```javascript
function setTriggerDay() {
  var last = lastBusinessDay();
  ScriptApp.newTrigger("setTriggerHoursLast")
    .timeBased()
    .atDate(last.getFullYear(), last.getMonth()+1, last.getDate())
    .create();
}
```
### setTriggerHoursLast
Deletes any existing triggers for `setTriggerHoursLast` and sets a new trigger to run the sendSlack function at 18:00 on the last business day.

```javascript
function setTriggerHoursLast() {
  deleteTrigger("setTriggerHoursLast");
  ScriptApp.newTrigger("sendSlack")
    .timeBased()
    .after(8 * 60 * 60 * 1000)
    .create();
}
```

### sendSlack
Deletes any existing triggers for `sendSlack` and sends a notification to Slack with information from the spreadsheet.

```javascript
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
```

### deleteTrigger
Deletes all triggers with the specified handler function name.

```javascript
function deleteTrigger(name) {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == name) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
```

### lastBusinessDay
Calculates and returns the last business day of the current month.

```javascript
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
```
### isHoliday
Checks if the given date is a holiday in Japan.

```javascript
function isHoliday(day) {
  var startDate = new Date(day.setHours(0, 0, 0, 0));
  var endDate = new Date(day.setHours(23, 59, 59));
  var cal = CalendarApp.getCalendarById("ja.japanese#holiday@group.v.calendar.google.com");
  var holidays = cal.getEvents(startDate, endDate);
  return holidays.length != 0; // True if holiday
}
```
### testPost
A test function to send a sample message to Slack.

```javascript
function testPost() {
  const app_auth_token = "xoxb-XXXXXXXXXX-XXXXXXXXXX-XXXXXXXXXX";
  const channel = "your-channel-id";
  const result = postMessageToSlackChannel(app_auth_token, channel, getinfofromspreadsheet());
  Logger.log(result);
}
```
### postMessageToSlackChannel
Sends a message to a specified Slack channel.

```javascript
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
```

### getinfofromspreadsheet
Retrieves information from a Google Sheets spreadsheet and formats it for the Slack message.

```javascript
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
      if (sheet.getRange(i, 8).getValue() === year && sheet.getRange(i, 9).getValue() === month) {
        var ending = Utilities.formatDate(end, "JST", "MM/dd");
        text = text + "\n" + name + "さんの契約終了六カ月前が迫っています。 (契約終了日: " + y + "/" + ending + ")";
        checker = checker + 1;
      }
    }
    if (sheet.getRange(i, 10).getValue() === year && sheet.getRange(i, 11).getValue() === month) {
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
```
