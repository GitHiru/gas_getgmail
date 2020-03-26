/* ***********************************************************************************
 *　□ GA → [ Gsheet → Chatwork ]　アクセスデータ共有Bot 
 *
 *    Cf. GASでチャットワークにGoogleアナリティクスの前日レポートを自動送信
 *        https://tonari-it.com/gas-analytics-chatwork/
 *    Cf. GASを使って、ChatworkにGoogle Analyticsのレポートを送信する
 *        https://aimstogeek.hatenablog.com/entry/2019/04/18/160000
 *    Cf. 【初心者向けGAS】面倒なことはライブラリに任せよう！その概要と追加の方法
 *        https://tonari-it.com/gas-library/
 * ******************************************************************************** */

//▼ statement plactice
function myfunction(){
  console.log("Helllo");
}


//▼ ChatWorkAPI連携テスト（個人チャットに書き込み）
function testMessage(){
  var cw = ChatWorkClient.factory({token:'c193e0b11fd0c4e5281859a73e1fd795'});
  var body = "テストだよ！🐧";
  cw.sendMessageToMyChat(body);
}


//▼ ChatWorkAPI連携テスト（グループチャットに書き込み）
function sendMessageTest() { 
  var client = ChatWorkClient.factory({token: 'c193e0b11fd0c4e5281859a73e1fd795'});　//チャットワークAPI
  client.sendMessage({
    room_id:148285064, //ルームID #!rid148285064
    body: 'チャットワークにメッセージを表示するテスト🐧'});
}


//▼ 【GJK】ChatWorkBot
function cwFromGA() {
 
//  var mySS=SpreadsheetApp.getActiveSpreadsheet(); //スプレッドシートを取得
  var mySS = SpreadsheetApp.openById("1EfZB3sVkutkdToMXzD8jJluUVrMaoYhJRB5IwVwPhrc"); //スプレッドシートを取得
  var sheetDaily=mySS.getSheetByName("ga_cwbot_all"); //個別シートを取得

  var rowDaily=sheetDaily.getDataRange().getLastRow(); //Dailyシートの使用範囲のうち最終行を取得 

  var yDate = sheetDaily.getRange(rowDaily,1).getValue();
 
  // チャットワークに送る文字列を生成
  var strBody = "[toall]" + "\n" + "[info][title]アクセス報告　"
      + Utilities.formatDate(yDate, 'JST', 'yyyy/MM/dd') + "(昨日)　のアクセス数でーす！ by 🐧" + "[/title]" +  //ga:date
        "ユーザー        : "
      + sheetDaily.getRange(rowDaily,2).getValue() + //ga:users
        "  (新規ユーザー : "
      + sheetDaily.getRange(rowDaily,3).getValue() + ")" + "\n" + //ga:newUsers
        "[hr]" +　"ページビュー : "
      + sheetDaily.getRange(rowDaily,4).getValue() + "\n" + //ga:pageviews
        "[hr]" +　"セッション     : "
      + sheetDaily.getRange(rowDaily,5).getValue() + "[/info]"; //ga:sessions
  
  
  // チャットワークにメッセージを送る
  var cwClient = ChatWorkClient.factory({token: 'c193e0b11fd0c4e5281859a73e1fd795'}); //チャットワークAPI
  cwClient.sendMessage({
    room_id: 148285064, //ルームID
    body: strBody
  });
//  cwClient.sendMessageToMyChat(strBody);//（テスト）個人チャットに送信
}


//自動実行するトリガー作成
function setTrigger(){
  var setTime = new Date();
  setTime.setHours(10);
  setTime.setMinutes(00); 
  ScriptApp.newTrigger('cwFromGA').timeBased().at(setTime).create();
}