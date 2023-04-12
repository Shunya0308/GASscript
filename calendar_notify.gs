//環境変数設定
var CHANNEL_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty("ACCESS_TOKEN");
var SHEET_ID = PropertiesService.getScriptProperties().getProperty("SHEET_ID");
var USER_ID = PropertiesService.getScriptProperties().getProperty("USER_ID");
var GOOGLE_ID = PropertiesService.getScriptProperties().getProperty("GOOGLE_ID");

//日付、時刻のフォーマット設定
var dateExp = /(\d{2})\/(\d{2})\s(\d{2}):(\d{2})/;
var dayExp = /(\d+)[\/月](\d+)/;
var hourMinExp = /(\d+)[:時](\d+)*/;

//テスト用関数
function test() {
  try {
  setTrigger();
  } catch(error) {
  Logger.log(error.message);
  }
}

function doPost(e) {
  try {
    handleMessage(e);
  } catch(error) {
    logging("ToCalendarFromLineBot");
    logging(JSON.stringify(e));
    logging(JSON.stringify(error));
    logging(error.message);
    var replyToken = JSON.parse(e.postData.contents).events[0].replyToken;
    reply(replyToken, error.message);
  }
}

function routineNotify() {
  setTrigger();
  var startDate = new Date();
  var today = new Date().getDay();
  if (today == 1){
    var period = 7;
  } else { 
  var period = 1;
  }
  var schedule = getSchedule(startDate, period);
  push(scheduleToMessage(schedule, startDate, period));
}

function setTrigger() {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName("TriggerTime");
  const date = new Date();
  const japanStandardTime = new Date().toLocaleString({ timeZone: 'Asia/Tokyo' });
  const nowHours = ('00' + new Date(japanStandardTime).getHours()).slice(-2);
  const nowMin = ('00' + new Date(japanStandardTime).getMinutes()).slice(-2);
  const nowTime = nowHours + ":" + nowMin;
  const setTime = (('00' + sh.getRange("B1").getValue()).slice(-2) + ":" + ('00' + sh.getRange("B2").getValue()).slice(-2));

  if (nowTime < setTime) {
    date.setDate(date.getDate())
  } else {
    date.setDate(date.getDate() + 1)
  }
  date.setHours(sh.getRange("B1").getValue());
  date.setMinutes(sh.getRange("B2").getValue());
  delTrigger();
  ScriptApp.newTrigger('routineNotify').timeBased().at(date).create();
}

function delTrigger(){
  //GASプロジェクトに設定したトリガーをすべて取得
  const triggers = ScriptApp.getProjectTriggers();
  //トリガーの登録数をログ出力
  console.log('トリガー数：' + triggers.length);
  //トリガー登録数のforループを実行
  for(let i=0;i<triggers.length;i++){
    //取得したトリガーをdeleteTriggerで削除
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

function logging(str) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("errorLog");
  var ts = new Date().toLocaleString("japanese", {timeZone: "Asia/Tokyo"});
  sheet.appendRow([ts, str]);
}

function handleMessage(e) {
  var replyToken = JSON.parse(e.postData.contents).events[0].replyToken;
  var lineType = JSON.parse(e.postData.contents).events[0].type
  if (typeof replyToken === "undefined" || lineType === "follow") {
    return;
  }
  var userMessage = JSON.parse(e.postData.contents).events[0].message.text;
  var cache = CacheService.getScriptCache();
  var type = cache.get("type");

  if (type === null) {
    if (userMessage === "予定追加") {
      cache.put("type", 1);
      reply(replyToken, "予定日を教えてください！\n「06/17, 6月17日」などの形式なら大丈夫です！");

    } else if (userMessage === "一週間の予定") {
      var startDate = new Date();
      var period = 7
      var schedule = getSchedule(startDate, period);
      reply(replyToken, scheduleToMessage(schedule, startDate, period));

    } else if (userMessage === "今日の予定") {
      var startDate = new Date();
      var period = 1
      var schedule = getSchedule(startDate, period);
      reply(replyToken, scheduleToMessage(schedule, startDate, period));

    } else if (userMessage === "通知時間設定") {
      cache.put("type", 6);
      reply(replyToken, "通知したい時間(◯時◯分)を教えてください。\n「13:00, 13時, 13:20, 13時20分」の形式なら大丈夫です！");

    } else {
      reply(replyToken, "リッチメニューの「予定追加」で予定追加を、「一週間の予定」で一週間の予定を、「今日の予定」で本日の予定を、「通知時間設定」で毎日通知する時間を設定できます^^");
    }
  } else {
    if (userMessage === "キャンセル") {
      cache.remove("type");
      reply(replyToken, "キャンセルしました！");
      return;
    }

    switch(type) {
      case "1":
        // 予定日
        var [matched, month, day] = userMessage.match(dayExp); //dayexpの型に合う文字列を取り出す。
        cache.put("type", 2);
        cache.put("month", month);
        cache.put("day", day);
        reply(replyToken, month + "/" + day + "ですね！　次に開始時刻を教えてください。「13:00, 13時, 13:20, 13時20分」などの形式なら大丈夫です！");
        break;

      case "2":
        // 開始時刻
        var [matched, startHour, startMin] = userMessage.match(hourMinExp);
        cache.put("type", 3);
        cache.put("start_hour", startHour);
        if (startMin == null) startMin = "00";
        cache.put("start_min", startMin);
        reply(replyToken, startHour + ":" + startMin + "ですね！　次に終了時刻を教えてください。");
        break;

      case "3":
        // 終了時刻
        var [matched, endHour, endMin] = userMessage.match(hourMinExp);
        cache.put("type", 4);
        cache.put("end_hour", endHour);
        if (endMin == null) endMin = "00";
        cache.put("end_min", endMin);
        reply(replyToken, endHour + ":" + endMin + "ですね！　最後に予定名を教えてください！");
        break;

      case "4":
        // 予定名
        cache.put("type", 5);
        cache.put("title", userMessage);
        var [title, startDate, endDate] = createEventData(cache);
        reply(replyToken, toEventFormat(title, startDate, endDate) + "\nで間違いないでしょうか？ よろしければ「はい」をやり直す場合は「いいえ」をお願いいたします！");
        break;

      case "5":
        // 確認の回答がはい or いいえ
        cache.remove("type");
        if (userMessage === "はい") {
          var [title, startDate, endDate] = createEventData(cache);
          CalendarApp.getDefaultCalendar().createEvent(title, startDate, endDate);
          reply(replyToken, "設定しました！");
        } else {
          reply(replyToken, "お手数ですがもう一度お願いいたします...");
        }
        break;
      
      case "6":
        //通知時間設定
        var [matched, notifyHour, notifyMin] = userMessage.match(hourMinExp);
        cache.put("type", 7);
        cache.put("notify_Hour", notifyHour);
          if (notifyMin == null) notifyMin = "00";
            cache.put("notify_Min", notifyMin);
            var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName("TriggerTime");
            sh.getRange("B1").setValue(notifyHour);
            sh.getRange("B2").setValue(notifyMin);
          reply(replyToken, notifyHour + "時" + notifyMin + "分で間違いないでしょうか？ よろしければ「はい」をやり直す場合は「いいえ」をお願いいたします！");
        break;

      case "7":
        // 確認の回答がはい or いいえ
        cache.remove("type");
        if (userMessage === "はい") {
          setTrigger();
          reply(replyToken, "追加しました！\nお疲れ様でした！");
        } else {
          reply(replyToken, "お手数ですがもう一度お願いいたします！");
        }
        break;
      
        default:
        reply(replyToken, "申し訳ありません。\n形式に誤りがないか確認してみて、なければ「キャンセル」で予定入力をキャンセルすることができるので、そちらを試していただけますか？");
        break;
    }
  }
}

// Googleカレンダーから予定取得
function getSchedule(startDate, period){
  var calendarIDs = [GOOGLE_ID]; //カレンダーIDの配列。カレンダーを共有している場合、複数指定可。

  var schedule = new Array(calendarIDs.length);
  for(var i=0; i<schedule.length; i++){schedule[i] = new Array(period);}

  for(var iCalendar=0; iCalendar < calendarIDs.length; iCalendar++){
    var calendar = CalendarApp.getCalendarById(calendarIDs[iCalendar]);

    var date = new Date(startDate);
    for(var iDate=0; iDate < period; iDate++){
      schedule[iCalendar][iDate] = getDayEvents(calendar, date);
      date.setDate(date.getDate() + 1);
    }
  }

  return schedule;
}  

//日付ごとの予定の格納処理
function getDayEvents(calendar, date){
  var dayEvents = "";
  var events = calendar.getEventsForDay(date);

　//予定の個数分ループ処理
  for(var iEvent = 0; iEvent < events.length; iEvent++){
    event = events[iEvent]
    var event = events[iEvent];
    var title = event.getTitle();
    var startTime = _HHmm(event.getStartTime());
    var endTime = _HHmm(event.getEndTime());
      
    dayEvents = dayEvents + '・' + startTime + '-' + endTime + ' ' + title + '\n';
  }

  return dayEvents;
}

function scheduleToMessage(schedule, startDate, period){
  var now = new Date();
  if (period == 7) {
    var body = '1週間の予定\n'
    + '(' +_Md(now) + ' ' +_HHmm(now) + '時点)\n'
    + '--------------------\n'; 
  } else {
    var body = '今日の予定\n'
    + '(' +_Md(now) + ' ' +_HHmm(now) + '時点)\n'
    + '--------------------\n';
  }

  var date = new Date(startDate);
  for(var iDay=0; iDay < period; iDay++){
    body = body + _Md(date) + '(' + _JPdayOfWeek(date) + ')\n';

    for(var iCalendar=0; iCalendar < schedule.length; iCalendar++){
      if (schedule[iCalendar][iDay] === "" && period === 1){
        body += '今日の予定はありません。';
      } else {
      body = body + schedule[iCalendar][iDay];
      }   
    }
    date.setDate(date.getDate() + 1);
    body = body + '\n';
  }

  return body;
}

function createEventData(cache) {
  var year = new Date().getFullYear();
  var title = cache.get("title");
  var startDate = new Date(year, cache.get("month") - 1, cache.get("day"), cache.get("start_hour"), cache.get("start_min"));
  var endDate = new Date(year, cache.get("month") - 1, cache.get("day"), cache.get("end_hour"), cache.get("end_min"));
  return [title, startDate, endDate];
}

function toEventFormat(title, startDate, endDate) {
  var start = Utilities.formatDate(startDate, "JST", "MM/dd HH:mm");
  var end = Utilities.formatDate(endDate, "JST", "MM/dd HH:mm");
  var str = title + ": " + start + " ~ " + end;
  return str;
}

function push(text) {

  var url = "https://api.line.me/v2/bot/message/push";
  var headers = {
    "Content-Type" : "application/json; charset=UTF-8",
    'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
  };

  var postData = {
    "to" : USER_ID,
    "messages" : [
      {
        'type':'text',
        'text':text,
      }
    ]
  };

  var options = {
    "method" : "post",
    "headers" : headers,
    "payload" : JSON.stringify(postData)
  };

  return UrlFetchApp.fetch(url, options);
}

function reply(replyToken, message) {
  var url = "https://api.line.me/v2/bot/message/reply";
  UrlFetchApp.fetch(url, {
    "headers": {
      "Content-Type": "application/json; charset=UTF-8",
      "Authorization": "Bearer " + CHANNEL_ACCESS_TOKEN,
    },
    "method": "post",
    "payload": JSON.stringify({
      "replyToken": replyToken,
      "messages": [{
        "type": "text",
        "text": message,
      }],
    }),
  });
  return ContentService.createTextOutput(JSON.stringify({"content": "post ok"})).setMimeType(ContentService.MimeType.JSON);
}

// 日付指定の関数
function _HHmm(date){
  return Utilities.formatDate(date, 'JST', 'HH:mm');
}

function _Md(date){
  return Utilities.formatDate(date, 'JST', 'M/d');
}

function _JPdayOfWeek(date){
  var date = new Date()
  var dayStr = ['日', '月', '火', '水', '木', '金', '土'];
  var dayOfWeek = date.getDay();

  return dayStr[dayOfWeek];
}
