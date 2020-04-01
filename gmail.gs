/* ***********************************************************************************
 *　□ Gmail →　Gsheet　データ抽出用GAS 
 *
 *　　 Cf.https://www.casleyconsulting.co.jp/blog/engineer/2693/
 *　　 Cf.https://qiita.com/kazinoue/items/1e8ed4aebfb5c3c886db
 * ********************************************************************************** */

//Cf. 佐々木さん記述コード
var fileid = "15VepwALaweBw9JXID1YGolVVrGBXP70Ru_WPfDlHAbY";
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

function fetchData２(str, pre, suf) {
  var reg = new RegExp(pre + '.*?' + suf);
  var data = str.match(reg)[0].replace(pre, '');
    return data;
}

function fetchData3(str, pre, suf) {
  var reg = new RegExp(pre + '.*?' + suf);
  var data = str.match(reg)[0].replace(pre, '').replace(suf, '');
    return data;
}



//Cf.https://www.casleyconsulting.co.jp/blog/engineer/2693/
function getmail() {
  //現在開いているシートを指定します。
  var sheet = SpreadsheetApp.getActiveSheet();
  //Gmailのメールを検索して配列で取得します。(threads[0], threads[1], …)
  var threads = GmailApp.search('from:sample@test.co.jp subject:"【サンプル】"')
  
  //検索にヒットしたスレッドの数だけ繰り返します。
  for(var i=0; i<threads.length; i++){
    //スレッド1つを取得します。
    var thread = threads[i];
    //スレッドの中の複数のメールを配列で取得します。(mails[0], mails[1], …)
    var mails = thread.getMessages();
    
    //メールの数だけ繰り返します。
    for(var j=0; j<mails.length; j++){
      //メール1つを取得します。
      var mail = mails[j];
      //最下行に出力します。
      sheet.appendRow([mail.getDate(), mail.getSubject(), mail.getBody()]);
    }//for  
  }//for
}//getmail



//geLptMail（自作）
function getLpMail() {
  //書き出し用シート指定
  var ss = SpreadsheetApp.getOpenById(''); //GoogleSheetID
  var sheet = ss.getSheetByName(''); //シート名
  
  //Gmailのメールを検索して配列で取得　(threads[0], threads[1], …)
  var threads = GmailApp.search('from:sample@test.co.jp subject:"【サンプル】"'); 
  
  //検索にヒットした複数スレッドをスレッド毎に配列
  for( var i = 0; i < thresds.length; i++ ) {
    var thread = threads[i];
    //スレッドに分けた複数メールをメール毎に配列
    var mails = thread.getMessages();
    for(var j = 0; j < mails.lengs; j++) {
      var mail = mails[j];
      
      //メール内容吐出し
      sheet.appendRow([mail.getDate(),mail.getSubject(),mail.getBody()]);
      
    }
  }
}





//Cf.https://qiita.com/kazinoue/items/1e8ed4aebfb5c3c886db
var SearchString = "subject:お得婚.net☆結婚式プレゼント☆お申込みがありました";


function myFunction() {
  var myThreads = GmailApp.search(SearchString, 0, 1);
  var myMsgs = GmailApp.getMessagesForThreads(myThreads);

  for ( var threadIndex = 0 ; threadIndex < myThreads.length ; threadIndex++ ) {
    // メールから data と Start, End を抜き出す
    var mailBody = myMsgs[threadIndex][0].getPlainBody();

    // 正規表現マッチにより、メール本文から情報を抽出する。
    var data      = mailBody.match(/[何かの正規表現](.+)[何かの正規表現]/);

    // 今回の事例では、上記情報の有効期限がメール本文に差し込まれてきているため、それも抽出しておく。
    var startTime = mailBody.match(/Start Time: ([0-9]+)\/([0-9]+)\/([0-9]+) /);
    var endTime   = mailBody.match(  /End Time: ([0-9]+)\/([0-9]+)\/([0-9]+) /);

    // シートへ値を追加する：最初に作ったコードで、Spreadsheet への依存性があるコード
    // 単独の Google Apps Script から Spreadsheet を操作する場合はこう書く
    // var objSpreadSheet = SpreadsheetApp.openByUrl("[データの差し込みを行うスプレッドシートのURLを書く]");
    // var objSheet = objSpreadSheet.getSheetByName("[データを差し込むシート名を書く]");

    // シートへ値を追加する：getActive()を使って Spreadsheet への依存性を下げたコード
    // Spreadsheet の Apps Script として書く場合はこれでよい
    var objSheet = SpreadsheetApp.getActive().getSheetByName("otc_cv_mail");

    // 指定されたシートの1行目は見出し行とし、最新のデータを常に２行目に表示させたいので、１行目の後ろ（２行目）に空行を差し込む。
    objSheet.insertRowAfter(1);

    // 抽出した情報は A2 セルに書き込む。
    objSheet.getRange("A2").setValue(data[1]);

    // メール本文から抽出した有効期限情報は B2 = 開始日、C2 = 終了日 に書き込む。
    // ここでは mm/dd/yyyy なフォーマットで日付が表記されている前提で処理を行っている。   
    objSheet.getRange("B2").setValue(Utilities.formatString("%d/%d/%d", startTime[3], startTime[1], startTime[2]));
    objSheet.getRange("C2").setValue(Utilities.formatString("%d/%d/%d",   endTime[3],   endTime[1],   endTime[2]));
  }
}
