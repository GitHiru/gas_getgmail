
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

////▼ statement plactice
//function myfunction(){
//  console.log("Helllo");
//}
//
//
////▼ ChatWorkAPI連携テスト（個人チャットに書き込み）
//function testMessage(){
//  var cw = ChatWorkClient.factory({token:'c193e0b11fd0c4e5281859a73e1fd795'});
//  var body = "テストだよ！🐧";
//  cw.sendMessageToMyChat(body);
//}
//
//
////▼ ChatWorkAPI連携テスト（グループチャットに書き込み）
//function sendMessageTest() { 
//  var client = ChatWorkClient.factory({token: 'c193e0b11fd0c4e5281859a73e1fd795'});　//チャットワークAPI
//  client.sendMessage({
//    room_id:148285064, //ルームID #!rid148285064
//    body: 'チャットワークにメッセージを表示するテスト🐧'});
//}


//▼ 自動実行するトリガー作成
function setTrigger(){
  var setTime = new Date();
  setTime.setHours(10);
  setTime.setMinutes(00); 
  ScriptApp.newTrigger('{関数名}').timeBased().at(setTime).create();
}


//▼ [GJK]ChatWorkBot all
function cwFromGA() {
  var mySS       = SpreadsheetApp.openById("1EfZB3sVkutkdToMXzD8jJluUVrMaoYhJRB5IwVwPhrc"); //スプレッドシートを取得
  var sheetDaily = mySS.getSheetByName("ga_cwbot_all"); //個別シートを取得
  var rowDaily   = sheetDaily.getDataRange().getLastRow(); //使用範囲のうち最終行を取得 
  var yDate      = sheetDaily.getRange(rowDaily,1).getValue();
 
  // チャットワークに送る文字列を生成
  // 前日差分を視認条件分岐（PV/SS）
  var pvDif = sheetDaily.getRange(rowDaily-1,4).getValue();
  var ssDif = sheetDaily.getRange(rowDaily-1,5).getValue();
  var pv    = sheetDaily.getRange(rowDaily,4).getValue();
  var ss    = sheetDaily.getRange(rowDaily,5).getValue();
  
  var pvVal = pv - pvDif;
  if ( pvVal >= 0 ){
    pvVal = "➚";
  } else{
    pvVal = "➘";
  }
  
  var ssVal = ss - ssDif;
  if ( ssVal >= 0 ){
    ssVal = "➚";
  } else{
    ssVal = "➘";
  }
  
  var pvGrow = Math.ceil(((pv-pvDif)/pv * 100)*10)/10 + "%";
  var ssGrow = Math.ceil(((ss-ssDif)/pv * 100)*10)/10 + "%";  
  
  var strBody = "[toall]" + "\n" 
      + "[info][title]【ごじょクル】アクセス全体報告（前日）　"
      + Utilities.formatDate(yDate, 'JST', 'yyyy/MM/dd')  + "[/title]"
      + "　ユーザー        : "
      + sheetDaily.getRange(rowDaily,2).getValue()
      + "  (新規 : "
      + sheetDaily.getRange(rowDaily,3).getValue() + ")" + "\n"
      + "[hr]" 
      + "　ページビュー : " + pv + " " + pvVal + " (前日比：" + pvGrow + ")" + "\n" 
      + "[hr]" 
      + "　セッション     : " + ss + " " + ssVal + " (前日比：" + ssGrow + ")";
  
  strBody = strBody + "[/info]";
  
  strBody = strBody + "\n"
  + "[info][title]　デイリー獲得（📃=資料請求 👥=見積もり、仮入会）[/title]"
  + "互助会　　:" + "📃[" + sheetDaily.getRange(rowDaily,6).getValue() +"] " + "👥[" + sheetDaily.getRange(rowDaily,7).getValue() +"]" + "\n"
  + "[hr]"
  + "葬儀場　　:" + "📃[" + sheetDaily.getRange(rowDaily,8).getValue() +"] " + "👥[" + sheetDaily.getRange(rowDaily,9).getValue() +"]" + "\n"
  + "[hr]"
  + "結婚式場　:" + "📃[" + sheetDaily.getRange(rowDaily,10).getValue() +"] " + "👥[" + sheetDaily.getRange(rowDaily,11).getValue() +"]" ;
  
  strBody = strBody + "[/info]" + "※　こちらの報告はBotによる投稿です。";
  
  
  // チャットワークにメッセージを送る
  var cwClient = ChatWorkClient.factory({token: 'c193e0b11fd0c4e5281859a73e1fd795'}); //チャットワークAPI
//  cwClient.sendMessage(
//    {room_id: 148285064, body: strBody} //「ごじょくる」
//  );
  cwClient.sendMessageToMyChat(strBody);


//  // (ADrim) チャットワークに送る文字列を生成
//  //https://docs.google.com/spreadsheets/d/15qzgLaaN7XWHgudWVyLkyl9QSQl_UqRy07iVZJJUFpI/edit#gid=0
//  var adSS     = SpreadsheetApp.openById("15qzgLaaN7XWHgudWVyLkyl9QSQl_UqRy07iVZJJUFpI");
//  var adDaily  = adSS.getSheetByName("2020年4月");
//  var rowDaily2 = adDaily.getDataRange(2).getLastRow();
//  var strBody2 = strBody + "\n\n" +
//    "[info][title] 運用型広告報告 [/title]" + 
//      " 消化予算 ：" + adDaily.getRange(rowDaily2,3).getValue() +
//      " 獲得件数 :" + adDaily.getRange(rowDaily2,4).getValue() + "[/info]"+
//    "[info][title] アフィリエイト広告報告 [/title]" +
//      " 成果報酬金額予測 ：" + adDaily.getRange(rowDaily2,5).getValue() +
//      " 獲得件数 :" + adDaily.getRange(rowDaily2,6).getValue() + "[/info]" + "※　こちらの報告はBotによる投稿です。";

  
//  //チャットワークにメッセージを送る
//  var cwClient = ChatWorkClient.factory({token: 'c193e0b11fd0c4e5281859a73e1fd795'}); //チャットワークAPI
//  cwClient.sendMessage(
//    {room_id: 74269449, body: strBody} //「Hershe ＆ ADrim Gr」
//  );
}


//▼ [GJK]ChatworkBot Owend
function cwFromGAOwend(){
  var mySS = SpreadsheetApp.openById("1EfZB3sVkutkdToMXzD8jJluUVrMaoYhJRB5IwVwPhrc");
  
  var sheetDaily = mySS.getSheetByName("ga_cwbot_owned");
  var rowDaily = sheetDaily.getDataRange().getLastRow();
  var yDate = sheetDaily.getRange(rowDaily,1).getValue();
  
  //　get data you wont
  var strBody = "[toall]" + "\n"
      + "[info][title] 【終活スタイル】前日アクセス報告　"
      + Utilities.formatDate(yDate, 'JST', 'yyyy/MM/dd') + "[/title]" + "\n"
      + " ユーザー        : " + sheetDaily.getRange(rowDaily,2).getValue() + "  (新規ユーザー : " + sheetDaily.getRange(rowDaily,3).getValue() + ")" + "\n"
      + "[hr]"
      + " ページビュー : " + sheetDaily.getRange(rowDaily,4).getValue() + "\n"
      + "[hr]" 
      + " セッション     : " + sheetDaily.getRange(rowDaily,5).getValue() + "[/info]" + "\n";

  //　get data you wont -vol2(owend page ranking)
  strBody = strBody + "[info][title] 記事別PV数ランキングTOP10👑（※PV数当月累積） [/title]" + "\n";
  var sheetPost = mySS.getSheetByName("ga_cwbot_owned_ranking");
  
  for(var i=1;i<=10;i++){
    // [1]{{記事タイトル}}：{{PV数}}
    strBody = strBody + "[" + i + "] " + sheetPost.getRange(i+15,1).getValue() + "：" + sheetPost.getRange(i+15,3).getValue() + "PV" + "\n";
  }
  
  strBody = strBody + "[/info]" + "※　こちらの報告はBotによる投稿です。";
  
  //　send message to Chatwork
  var cwClient = ChatWorkClient.factory({token: 'c193e0b11fd0c4e5281859a73e1fd795'});
  cwClient.sendMessage({
    room_id: 182163803,
    body: strBody
  });
}








//▼ [OTC]ChatworkBot LP
function otcBot(){
  var mySS = SpreadsheetApp.openById("16nWrrg6alfmq2168KrGgJkwAJ9wrCKd5PGLlhkmSskI");

  
  var lpDaily_st = mySS.getSheetByName("cw_lp_standard");   
  var rowDaily = lpDaily_st.getDataRange().getLastRow();
  var yDate = lpDaily_st.getRange(rowDaily,1).getValue();
  var strBody = "[toall]" + "\n\n"
      + "💎【お得婚LP施策】前日アクセス報告" + "(" +Utilities.formatDate(yDate, 'JST', 'yyyy/MM/dd') + ")" + "\n\n"
      + "[info][title] 通常 LP [/title]" + "\n"
      + "　　ページビュー　　　　: " + lpDaily_st.getRange(rowDaily,2).getValue() + "\n"
      + "[hr]"
      + "　　セッション　　　　　: " + lpDaily_st.getRange(rowDaily,3).getValue() + "\n"
      + "[hr]" 
      + "　　コンバージョン　　　: " + lpDaily_st.getRange(rowDaily,4).getValue() + "[/info]" + "\n";

 
  var lpDaily_ho = mySS.getSheetByName("cw_lp_hokkaido"); 
  var rowDaily = lpDaily_ho.getDataRange().getLastRow();
  var yDate = lpDaily_ho.getRange(rowDaily,1).getValue();
  strBody = strBody
      + "[info][title] 北海道 LP  [/title]" + "\n"
      + "　　ページビュー　　　　: " + lpDaily_ho.getRange(rowDaily,2).getValue() + "\n"
      + "[hr]"
      + "　　セッション　　　　　: " + lpDaily_ho.getRange(rowDaily,3).getValue() + "\n"
      + "[hr]" 
      + "　　コンバージョン　　　: " + lpDaily_ho.getRange(rowDaily,4).getValue() + "[/info]" + "\n";
  
  
  var lpDaily_fa = mySS.getSheetByName("cw_lp_family"); 
  var rowDaily = lpDaily_fa.getDataRange().getLastRow();
  var yDate = lpDaily_fa.getRange(rowDaily,1).getValue();
  strBody = strBody
      + "[info][title] パパママ LP [/title]" + "\n"
      + "　　ページビュー　　　　: " + lpDaily_fa.getRange(rowDaily,2).getValue() + "\n"
      + "[hr]"
      + "　　セッション　　　　　: " + lpDaily_fa.getRange(rowDaily,3).getValue() + "\n"
      + "[hr]" 
      + "　　コンバージョン　　　: " + lpDaily_fa.getRange(rowDaily,4).getValue() + "[/info]" + "※　こちらの報告はBotによる投稿です。";
  
  
  var cwClient = ChatWorkClient.factory({token: 'c193e0b11fd0c4e5281859a73e1fd795'});
  cwClient.sendMessage({
    room_id: 65572941,
    body: strBody
  });
  
//  cwClient.sendMessageToMyChat(strBody); //test
  
  
  
  
  /*
   *　｛｛開発中｝｝
   */
//  var mySS = SpreadsheetApp.openById("16nWrrg6alfmq2168KrGgJkwAJ9wrCKd5PGLlhkmSskI");
//  var ssName = ["cw_lp_standard","cw_lp_hokkaido","cw_lp_family"];
//  var len = ssName.length;
//  var res ="";
//  
//  for(var i=0; i<=len; i++){
//    
//    var ssData = mySS.getSheetByName(ssName[i]);
//    
//    var rowDaily = ssData.getDataRange().getLastRow();
//    var yDate = ssData.getRange(rowDaily,1).getValue();
//      
//    var strBody =
////          if(j===0){
////              "[toall]" + "\n"
////            + "【お得婚LP施策】前日アクセス報告" + "\n\n"
////          }
//        "[info][title]"+ ssData.getName() + Utilities.formatDate(yDate, 'JST', 'yyyy/MM/dd') + "[/title]" + "\n"
//      + "　　ページビュー　　　　: " + ssData.getRange(rowDaily,2).getValue() + "\n"
//      + "[hr]"
//      + "　　セッション　　　　　: " + ssData.getRange(rowDaily,3).getValue() + "\n"
//      + "[hr]" 
//      + "　　コンバージョン　　　: " + ssData.getRange(rowDaily,4).getValue() + "[/info]" + "\n";
//    
//    return(res);
//  }
//  
//  console.log(res);
//  
//    
//  var cwClient = ChatWorkClient.factory({token: 'c193e0b11fd0c4e5281859a73e1fd795'});
//  cwClient.sendMessageToMyChat(strBody); //test
    
}