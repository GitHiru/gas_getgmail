/* ***********************************************************************************
 *　□ Gmail →　Gsheet　データ抽出用GAS 
 *
 * Cf.) https://www.casleyconsulting.co.jp/blog/engineer/2693/
 * Cf.) https://qiita.com/kazinoue/items/1e8ed4aebfb5c3c886db
 * 　
 * 既存のコード
 * Cf.) GAS	https://script.google.com/a/hershe.jp/d/1Pjy9wyh77JJ0Mvis9OmjfS5LXDeZxz-yuU2t3IhiQd5wqv-JDPx9jkA8/edit
 * Cf.) 応募sheet	https://docs.google.com/spreadsheets/d/1J7ltLekcOiI9pVyGng_kiDrbA_TPy4RjlUVP7Pe1Dv0/edit#gid=0
 * ********************************************************************************** */


var myFile = SpreadsheetApp.openById('16nWrrg6alfmq2168KrGgJkwAJ9wrCKd5PGLlhkmSskI'); //ファイル名
var mySheet = myFile.getSheetByName('mail'); //シート名
var strTerms = 'subject:お得婚.net☆結婚式プレゼント☆お申込みがありました';　//検索文字列（条件）
var searchCount = 30; //最大500通
var myThreads = GmailApp.search(strTerms,0,searchCount);　//条件にマッチしたスレッドを検索して取得


function otc_getMail() {
  var myMsgs = GmailApp.getMessagesForThreads(myThreads);　//二次元配列
  var valMsgs = [];
  //Logger.log(myFile.getName());

  /* 各メールから日時、送信元、件名、内容を取り出す */
  for(var i=0;i<myMsgs.length;i++){
    
    for(var j=0;j<myMsgs[i].length;j++){
      var msid = myMsgs[i][j].getId();//メッセージIDを取得
　　　 //もしメッセージIDがスプレッドシートに存在しなければ
　　　 if(!hasId(msid)){
　　　　 var date = myMsgs[i][j].getDate();
　　　　 var from = myMsgs[i][j].getFrom();
　　　　 var subj = myMsgs[i][j].getSubject();
     　 var perm = myThreads[i].getPermalink();
　　　　 var body = myMsgs[i][j].getPlainBody(); //.slice(0,200);
      　//var line_body = body.replace(/\r?\n/g, '');
      
      　valMsgs.push([date,from,subj,body,msid,perm]);
      
　　　} //if
　　} //for
　} //for
 
　/* スプレッドシートに出力 */
　if(valMsgs.length>0){//新規メールがある場合、末尾に追加する
　　var lastRow = mySheet.getDataRange().getLastRow();
　　mySheet.getRange(lastRow+1, 1, valMsgs.length, 6).setValues(valMsgs);
　}
}
 
function hasId(id){
 
  var data = mySheet.getRange(1, 5, mySheet.getLastRow(),1).getValues();//E列(メッセージID)を検索範囲とする
  var hasId = data.some(function(value,index,data){//コールバック関数
    return (value[0] === id);
  });
  return hasId;
}















var fileid = "1J7ltLekcOiI9pVyGng_kiDrbA_TPy4RjlUVP7Pe1Dv0"; //16nWrrg6alfmq2168KrGgJkwAJ9wrCKd5PGLlhkmSskI
var SearchString = "subject:お得婚.net☆結婚式プレゼント☆お申込みがありました";
var sheetfile = SpreadsheetApp.openById(fileid);
var sheet    = sheetfile.getSheetByName('応募者情報');
var row = sheet.getRange(2, 16, sheet.getLastRow()).getValues();

function getMail() {
    var threads = GmailApp.search(SearchString);
    for (var i = 0; i < threads.length; i++) {
        var thread = threads[i];
        var messages = thread.getMessages();
        
        for (var j = 0; j < messages.length; j++) {
            var message = messages[j];
            var body = message.getPlainBody();
            var id = message.getId();
            var date = message.getDate();
            var line_body = body.replace(/\r?\n/g, '');
            var groom = fetchData2(line_body, '新郎様', '新婦様');
            var bride = fetchData2(line_body, '新婦様', '連絡先');
            if(!hasId(id)){
              sheet.appendRow([
                date,
                fetchData(body, 'ご選択式場：', '\r'),
                fetchData(groom, 'お名前：', 'フリガナ'),
                fetchData(groom, 'フリガナ：', 'ご年齢'),
                fetchData(groom, 'ご年齢：', '新婦様'),
                fetchData(bride, 'お名前：', 'フリガナ'),
                fetchData(bride, 'フリガナ：', 'ご年齢'),
                fetchData(bride, 'ご年齢：', '連絡先'),
                fetchData(body, '連絡先：', '\r'),
                fetchData(body, '連絡先電話番号：', '\r'),
                fetchData(body, '連絡先メールアドレス：', '\r'),
                fetchData(body, 'SNS：', '\r'),
                fetchData(body, '連絡の取りやすい時間帯：', '\r'),
                fetchData(body, 'ご応募のきっかけ：', '\r'),
                fetchData3(line_body, '応募動機：', '-'),
                id
              ]);
            }
            sheet.sort(1, false);
        }
    }
}

function hasId(id) {
  var hasId = row.some(function(array, i, row) {
    return (array[0] === id);
  });
  return hasId;
}
 
 
function fetchData(str, pre, suf) {
  var reg = new RegExp(pre + '.*' + suf);
  var data = str.match(reg)[0].replace(pre, '')
    .replace(suf, '');
    return data;
}

function fetchData3(str, pre, suf) {
  var reg = new RegExp(pre + '.*?' + suf);
  var data = str.match(reg)[0].replace(pre, '').replace(suf, '');
    return data;
}



////Cf.https://www.casleyconsulting.co.jp/blog/engineer/2693/
//function getmail() {
//  //現在開いているシートを指定します。
//  var sheet = SpreadsheetApp.getActiveSheet();
//  //Gmailのメールを検索して配列で取得します。(threads[0], threads[1], …)
//  var threads = GmailApp.search('from:sample@test.co.jp subject:"【サンプル】"')
//  
//  //検索にヒットしたスレッドの数だけ繰り返します。
//  for(var i=0; i<threads.length; i++){
//    //スレッド1つを取得します。
//    var thread = threads[i];
//    //スレッドの中の複数のメールを配列で取得します。(mails[0], mails[1], …)
//    var mails = thread.getMessages();
//    
//    //メールの数だけ繰り返します。
//    for(var j=0; j<mails.length; j++){
//      //メール1つを取得します。
//      var mail = mails[j];
//      //最下行に出力します。
//      sheet.appendRow([mail.getDate(), mail.getSubject(), mail.getBody()]);
//    }//for  
//  }//for
//}//getmail
