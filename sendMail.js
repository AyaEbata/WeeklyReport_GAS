function onOpen(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var items = [
    {name: 'テスト送信', functionName: 'sendTestMail'}
  ];
  sheet.addMenu('メール送信メニュー', items);
}

/* テスト送信 */
function sendTestMail() {
  var book = SpreadsheetApp.getActive();
  var sheet = book.getActiveSheet();

  var sendTo = sheet.getRange("B2").getValue();

  var mailText = subject(sheet);
  
  MailApp.sendEmail(sendTo, mailText.subject, mailText.body);
}

/* 予約送信（曜日指定はトリガーで） */
function reservationMail() {
  var book = SpreadsheetApp.getActive();
  var sheet = book.getActiveSheet();

  var sendTo = sheet.getRange("B1").getValue();
  var sendBcc = sheet.getRange("B2").getValue();
  
  var mailText = subject(sheet);
    
  MailApp.sendEmail(sendTo, mailText.subject, mailText.body, {bcc: sendBcc});
}


function weekStart() {
  var between = {};
  var today = new Date();
  var week = today.getDay();
  var start;
  
  if (week == 0) {
    start = computeDate(today, -6);   
  } else {
    start = computeDate(today, 1-week);
  }
  return start;
}

function weekEnd() {
  var between = {};
  var today = new Date();
  var week = today.getDay();
  var end;
  
  if (week == 0) {
    end = computeDate(today, -2);   
  } else {
    end = computeDate(today, 5-week);
  }
  return end;
}

function addZero(num){
  return ('0' + num).slice(-2);
}

function computeDate(today, addDays) {
    var baseSec = today.getTime();
    var addSec = addDays * 86400000;
    var targetSec = baseSec + addSec;
    today.setTime(targetSec);
  
    return today;
}
    
function subject(sheet) {
  var sendToName = sheet.getRange("B3").getValue();
  var myName = sheet.getRange("B4").getValue();
  var renowned = sheet.getRange("B5").getValue();
  
  var start = weekStart();
  var end = weekEnd();
  
  var subject = '週報' + String(start.getFullYear()) + String(addZero(start.getMonth()+1)) + String(addZero(start.getDate())) + '-' + addZero(end.getDate());
  var body = 'to. ' + sendToName + '\n\n'
             + 'お疲れ様です。' + myName + 'です。\n\n'
             + '以下、' + String(start.getFullYear()) + '/' + String(addZero(start.getMonth()+1)) + '/' + String(addZero(start.getDate())) + '〜' + String(end.getFullYear()) + '/' + String(addZero(end.getMonth()+1)) + '/' + String(addZero(end.getDate())) + 'の週報を提出致します。\n'
             + 'お手数をおかけしますが、ご確認をお願い致します。\n\n'
             + '【勤務上の連絡】\n'
             + '　特になし\n\n'
             + '【営業への連絡】\n'
             + '　特になし\n\n'
             + '【その他の連絡】\n'
             + '　特になし\n\n'
             + '以上、よろしくお願い致します。\n'
             + renowned;
    
    return {'subject': subject, 'body':body};
}
    