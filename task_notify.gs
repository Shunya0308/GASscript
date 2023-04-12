var CHANNEL_ACCESS_TOKEN = PropertiesService.getScriptProperties().getProperty("CHANNEL_ACCESS_TOKEN");
var ERROR_SHEET_ID = PropertiesService.getScriptProperties().getProperty("ERROR_SHEET_ID");
var TRIGGER_SETTIME_SHEET_ID = PropertiesService.getScriptProperties().getProperty("TRIGGER_SETTIME_SHEET_ID");
var USER_ID = PropertiesService.getScriptProperties().getProperty("USER_ID");

//日付、時刻のフォーマット設定
var dateExp = /(\d{2})\/(\d{2})\s(\d{2}):(\d{2})/;
var dayExp = /(\d+)[\/月](\d+)/;
var hourMinExp = /(\d+)[:時](\d+)*/;

function doPost(e) {
  try {
    handleMessage(e);
  } catch(error) {
    logging("TaskNotifyForLINEbot");
    logging(JSON.stringify(e));
    logging(JSON.stringify(error));
    logging(error.message);
    var replyToken = JSON.parse(e.postData.contents).events[0].replyToken;
    reply(replyToken, error.message);
  }
}

function main() {
  setTrigger();
  push(getTask());
}

function setTrigger() {
  var sh = SpreadsheetApp.openById(TRIGGER_SETTIME_SHEET_ID).getSheetByName("TriggerTime");
  const date = new Date();
  const japanStandardTime = new Date().toLocaleString({ timeZone: 'Asia/Tokyo' });
  const nowHours = ('00' + new Date(japanStandardTime).getHours()).slice(-2);
  const nowMin = ('00' + new Date(japanStandardTime).getMinutes()).slice(-2);
  const nowTime = (nowHours + ":" + nowMin);
  const setTime = (('00' + sh.getRange("B1").getValue()).slice(-2) + ":" + ('00' + sh.getRange("B2").getValue()).slice(-2));

  if (nowTime < setTime) {
    date.setDate(date.getDate())
  } else {
    date.setDate(date.getDate() + 1)
  }
  date.setHours(sh.getRange("B1").getValue());
  date.setMinutes(sh.getRange("B2").getValue());
  delTrigger();
  ScriptApp.newTrigger('main').timeBased().at(date).create();
}

function logging(str) {
  var sheet = SpreadsheetApp.openById("1MePDOwolza3-mYFZU5o7vf2mwOZlMwn-9f5mxe6jRT8").getActiveSheet();
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
    if (userMessage === "通知時間設定") {
      cache.put("type", 1);
      reply(replyToken, "通知したい時間(◯時◯分)を教えてください。\n「13:00, 13時, 13:20, 13時20分」の形式なら大丈夫です^^");
    } else if (userMessage === "タスク通知") {
      reply(replyToken, getTask());
    } else {
      reply(replyToken, "「タスク通知」で、今現在Trelloにあるタスクを出力してくれます。\n「通知時間設定」で、自分の好きな時間に毎日タスクを通知してくれます。");
    }
  } else {
    if (userMessage === "キャンセル") {
      cache.remove("type");
      reply(replyToken, "キャンセルしました！");
      return;
    }

    switch(type) {
      case "1":
        // 通知時間設定
        var [matched, notifyHour, notifyMin] = userMessage.match(hourMinExp); //hourminexpの型に合う文字列を取り出す。
        cache.put("type", 2);
        cache.put("notify_hour", notifyHour);
        if (notifyMin == null) notifyMin = "00";
        cache.put("notify_min", notifyMin);
        var sh = SpreadsheetApp.openById(TRIGGER_SETTIME_SHEET_ID).getSheetByName("TriggerTime");
        sh.getRange("B1").setValue(notifyHour);
        sh.getRange("B2").setValue(notifyMin);
        reply(replyToken, notifyHour + ":" + notifyMin + "で間違いないですか？よろしければ「はい」をやり直す場合は「いいえ」をお願いします。");
        break;

      case "2":
        // 確認の回答がはい or いいえ
        cache.remove("type");
        if (userMessage === "はい") {
          setTrigger();
          reply(replyToken, "設定しました！");
        } else {
          reply(replyToken, "お手数ですがもう一度お願いいたします。");
        }
        break;
      default:
        reply(replyToken, "申し訳ありません。\n形式に誤りがないか確認してみて、なければ「キャンセル」で設定をキャンセルすることができるので、そちらを試していただけますか？");
        break;
    }
  }
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

function getTask() {
  var trelloKey   = "b8b272c1bc2e2b8ffd36bdced4327884";//keyを入力してください
  var trelloToken = "ATTAfa93725ec59658f74da7501aa072cd7f163084d53cf4ed6ad9aa1f5995863e139218567C";//tokenを入力してください
  var listId 　　 = "61b735243f7db71b2ecee590";//リストIDを入力してください

  var url = "https://trello.com/1/lists/" + listId + "/cards?key=" + trelloKey + "&token=" + trelloToken;
  var res = UrlFetchApp.fetch(url, {'method':'get'});
  var json = JSON.parse(res.getContentText());

  var cards =[]; //箱を作る
  var maxRows = json.length; //格納されているデータの行数を取得

  var now = new Date();
  var message = '現在のタスク\n'
    + '(' +_Md(now) + ' ' +_HHmm(now) + '時点)\n'
    + '--------------------\n';
  for(var i = 0; i < maxRows; i++){

     //必要なデータのKeyを指定して値を取得する
     var name = json[i].name;
     var due = json[i].due;
     var idMembers = json[i].idMembers;
     var dateLastActivity = json[i].dateLastActivity;
     var shortUrl = json[i].shortUrl;

     //取得したここのデータをまとめて、ひとつのカード情報としてまとめる
     var card = [name, due, idMembers, dateLastActivity, shortUrl];

     //LINEに出力するメッセージを作成する
     if (i === 0 && json[i].name === "") {
        message += "現在タスクはありません。"
      } else if (i === maxRows - 1) {
        message += '・' + name;
      } else {
        message += '・' + name + '\n';
      }

     //先程取得したcardsという箱に追加する
     cards.push(card);
  }

  Logger.log(message);
  return message;
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

// 日付指定の関数
function _HHmm(date){
  return Utilities.formatDate(date, 'JST', 'HH:mm');
}

function _Md(date){
  return Utilities.formatDate(date, 'JST', 'M/d');
}