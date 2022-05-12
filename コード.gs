function createInvoice_all() {
  const ExportBilling = SpreadsheetApp.getActive().getSheetByName("一覧抽出");
  const LastRow = ExportBilling.getLastRow();

  const Invoice = SpreadsheetApp.getActive().getSheetByName("請求書");
  

  //const Name = ExportBilling.getRange(3,3).getValue();
  const StartDate = ExportBilling.getRange(4,3).getValue();
  const EndDate = ExportBilling.getRange(5,3).getValue();
  const Query = [];
  
  const Plist = SpreadsheetApp.getActive().getSheetByName("生徒マスタ");
  const finalRow = Plist.getLastRow();

  //Logger.log(Name);
  Logger.log(StartDate);
  Logger.log(EndDate);


  //ExportBillingに入っているデータ（D7J26）をクリアする
  ExportBilling.getRange(7,4,20,9).clearContent();


    for (let i = 2; i <= finalRow; i++){
      const Name = Plist.getRange(i, 2).getValue();//生徒マスタの名前一覧からとってくる
      Logger.log(Name);
      ExportBilling.getRange(3,3).setValue(Name);

      //priceSetting();
      //listCount();
      //sum();
      //paste();
      savepdf(Name);







}




}

function createInvoice_single() {
  const ExportBilling = SpreadsheetApp.getActive().getSheetByName("一覧抽出");
  const LastRow = ExportBilling.getLastRow();

  const Invoice = SpreadsheetApp.getActive().getSheetByName("請求書");
  

  //const Name = ExportBilling.getRange(3,3).getValue();
  const StartDate = ExportBilling.getRange(4,3).getValue();
  const EndDate = ExportBilling.getRange(5,3).getValue();
  const Query = [];
  
  const Plist = SpreadsheetApp.getActive().getSheetByName("生徒マスタ");
  const finalRow = Plist.getLastRow();

  //Logger.log(Name);
  Logger.log(StartDate);
  Logger.log(EndDate);


  //ExportBillingに入っているデータ（D7J26）をクリアする
  ExportBilling.getRange(7,4,20,9).clearContent();
  const Name = ExportBilling.getRange(3, 3).getValue();//一覧抽出の名前


      //const Name = Plist.getRange(i, 2).getValue();
      //Logger.log(Name);
      //ExportBilling.getRange(3,3).setValue(Name);

      priceSetting();
      listCount();
      sum();
      paste();
      savepdf(Name);












}





function updateStudentName(){
  //
  //const id = "1___GfeJI65M2v5NV5lNIadp9p0Rd8bv9m77epuylhQA";
  const st = "生徒マスタ";
  const book = SpreadsheetApp.getActive().getSheetByName(st);
  
  //
  //const value = SpreadsheetApp.getActive().getSheetByName("test").getRange("B2").getValue();
  
  const lastRow = book.getLastRow();
  const range = book.getRange(2,2,lastRow-1);
  //const range = book.getRange(2,2);
  const value = range.getValues();

  Logger.log(value);

  //const value = range.getvalue();


  SpreadsheetApp.getActive().getSheetByName("settings").getRange(2,3,value.length).setValues(value);

}

function priceSetting(){
  const List = SpreadsheetApp.getActive().getSheetByName("単価項目");
  const lastRow = List.getLastRow();
  const range = List.getRange(2,1,lastRow-1,2);//A2～B末まで

  const value = range.getValues();//単価項目から入塾金などを抽出
  const Key = [];
  const Data = [];

  const Plist = SpreadsheetApp.getActive().getSheetByName("生徒マスタ");
  const finalRow = Plist.getLastRow();
  const price = Plist.getRange(2,1,finalRow-1,13).getValues();//生徒マスタA2～Mまで 
  const Keyname = [];
  const Kakaku = [];
  const SKakaku = [];

  const Billing = SpreadsheetApp.getActive().getSheetByName("一覧抽出")
  const LastRow = Billing.getLastRow();
  const SearchKey = Billing.getRange(7,3,LastRow - 6,2).getValues();//C7～E末まで
  const Name = Billing.getRange(3,3).getValue();

Logger.log(value);
Logger.log(price);

  // キーと返したいデータを配列に格納
  for (let i = 0; i < value.length; i++){
    Key.push(value[i][0]);//内容項目
    Data.push(value[i][1]); //単価項目。配列の末尾を追加→配列の要素を増やす。



  }

  Logger.log(Key);
  Logger.log(Data);
  
      for (let j = 0; j < price.length; j++){
        Keyname.push(price[j][1]);//氏名をキーとする
        Kakaku.push(price[j][12]);//通常授業価格
        //SKakaku.push(price[j][13]);//季節講習授業価格
  }
    if (Kakaku[Keyname.indexOf(Name)] !== ''){ //なんかしら値があれば価格を入れる、なければ単価項目の価格を反映する仕様。
    Data[0] = Kakaku[Keyname.indexOf(Name)];
    }

    /*if (SKakaku[Keyname.indexOf(Name)] !== ''){ //なんかしら値があれば価格を入れる、なければ単価項目の価格を反映する仕様。
    Data[0] = SKakaku[Keyname.indexOf(Name)];
    }  
*/

  Logger.log(Keyname);
  Logger.log(Kakaku);
  //Logger.log(SKakaku);

  // 検索キーを取得し、対応したデータを返す

  
  for (let i = 0; i < SearchKey.length; i++){
  
    //その他で入力されている金額が上書きされないように条件を分岐
    if (SearchKey[i][0] !== "その他"){
    SearchKey[i][1] = Data[Key.indexOf(SearchKey[i][0])];
    }

    

  }

  // 対応データの入った配列「SearchKey」をシートに書き込む
  Billing.getRange(7,3,LastRow - 6,2).setValues(SearchKey);
  //6．その他が入っている場合はスキップする処理を入れたい。

Logger.log(SearchKey);

}

function listCount(){

  const Billing = SpreadsheetApp.getActive().getSheetByName("一覧抽出")
  const LastRow = Billing.getLastRow();
  const Values = Billing.getRange(7,2,LastRow - 6,4).getValues();//b2～f末まで
  const ClassLogID = "1APbAjNplIpW--AnuhRNb60mltaXH3gyRkMDXUxebjKY";
  const KeyName = Billing.getRange(3,3).getValue();

  const Normal_count = SpreadsheetApp.getActive().getSheetByName("通常カウント");
  const NCount = Normal_count.getRange(2,1,Normal_count.getLastRow()-1,3).getValues();
  const Season_count = SpreadsheetApp.getActive().getSheetByName("講習カウント");
  const SCount = Season_count.getRange(2,1,Season_count.getLastRow()-1,3).getValues();
  const MVP_count = SpreadsheetApp.getActive().getSheetByName("MVPカウント");
  const MCount = MVP_count.getRange(2,1,MVP_count.getLastRow()-1,3).getValues();

  const Plist = SpreadsheetApp.getActive().getSheetByName("生徒マスタ");
  const finalRow = Plist.getLastRow();
  const price = Plist.getRange(2,1,finalRow-1,13).getValues();//生徒マスタA2～Mまで 

Logger.log(Values);
Logger.log(NCount);
Logger.log(SCount);
Logger.log(price);

  for (let i = 0; i < Values.length; i++){
        //if (Values[i][0] === "4．設備費" || "5．製本代" || "6．その他"){
    if (Values[i][1] === "設備費"){
      Values[i][3] = 1;
    } 
    else if(Values[i][1] === "製本代（初月のみ請求、今後は発生いたしません）"){
      
      
        const Keyname = [];
        const Syokai = [];
       

        for (let j = 0; j < price.length; j++){
        Keyname.push(price[j][1]);//氏名をキーとする
        Syokai.push(price[j][11]);//初回請求
  }
      if(Syokai[Keyname.indexOf(KeyName)] === "未"){//未が入っていたら１を立てて請求する。
        Values[i][3] = 1;
        }
      else if(Syokai[Keyname.indexOf(KeyName)] === "済"){//済が入っていたら0を立てる。
        Values[i][3] = 0;
        }

    }

    else if (Values[i][1] === "入塾金（初月のみ請求、今後は発生いたしません）"){
        const Keyname = [];
        const Syokai = [];

        for (let j = 0; j < price.length; j++){
        Keyname.push(price[j][1]);//氏名をキーとする
        Syokai.push(price[j][11]);//初回請求
  }
      if(Syokai[Keyname.indexOf(KeyName)] === "未"){//未が入っていたら１を立てて請求する。
      Values[i][3] = 1;
      //price[j][10] = "済" 入力したら済に変更する。
      


      }

      else if (Syokai[Keyname.indexOf(KeyName)] === "済"){//済が入っていたら0を立てる。
      Values[i][3] = 0;
      }
    }

    else if (Values[i][1] === "その他"){
      Values[i][3] = 1;
    }

      else if(Values[i][1] === "通常授業料"){
          const Key = [];
          const Data = [];

            for (let i = 0; i < NCount.length; i++){
              Data.push(NCount[i][1]); //生徒名でのカウント。配列の末尾を追加→配列の要素を増やす。
              Key.push(NCount[i][0]);//生徒名
            }

              Logger.log(Key);
              Logger.log(Data);

        for (let i = 0; i < Values.length; i++){

          //Values[i][0]（生徒名）をキーとしてNormal_countから数値を引っ張ってくる
          if(Values[i][1] === "通常授業料"){

          //Values[i][3] = Data[Key.indexOf(Values[i][0])];
          Values[i][3] = Data[Key.indexOf(KeyName)];

          }
        }
 
      }
     else if(Values[i][1] === "季節講習授業料"){
          const Key = [];
          const Data = [];

            for (let i = 0; i < SCount.length; i++){
              Data.push(SCount[i][1]); //生徒名でのカウント。配列の末尾を追加→配列の要素を増やす。
              Key.push(SCount[i][0]);//生徒名
            }

              Logger.log(Key);
              Logger.log(Data);

        for (let i = 0; i < Values.length; i++){

          //Values[i][0]（生徒名）をキーとしてNormal_countから数値を引っ張ってくる
          if(Values[i][1] === "季節講習授業料"){

          //Values[i][3] = Data[Key.indexOf(Values[i][0])];
          Values[i][3] = Data[Key.indexOf(KeyName)];

          }
        }
 
      }

       else if(Values[i][1] === "特別講義授業料"){
          const Key = [];
          const Data = [];

            for (let i = 0; i < MCount.length; i++){
              Data.push(MCount[i][1]); //生徒名でのカウント。配列の末尾を追加→配列の要素を増やす。
              Key.push(MCount[i][0]);//生徒名
            }

              Logger.log(Key);
              Logger.log(Data);

        for (let i = 0; i < Values.length; i++){

          //Values[i][0]（生徒名）をキーとしてNormal_countから数値を引っ張ってくる
          if(Values[i][1] === "特別講義授業料"){

          //Values[i][3] = Data[Key.indexOf(Values[i][0])];
          Values[i][3] = Data[Key.indexOf(KeyName)];

          }
        }
 
      }
 


  }
     
  

Logger.log(Values);

// 対応データの入った配列「Values」をシートに書き込む
Billing.getRange(7,2,LastRow - 6,4).setValues(Values);





}

function sum() {
  const Billing = SpreadsheetApp.getActive().getSheetByName("一覧抽出");
  const lastRow = Billing.getLastRow();
  
  for (let i = 7; i <= lastRow; i++){
    Logger.log(i);
  const tannka= Billing.getRange(i, 4).getValue();
  const suuryou= Billing.getRange(i, 5).getValue();
  Billing.getRange(i, 6).setValue(tannka * suuryou);

  }


}

function paste(){
  const ExportBilling = SpreadsheetApp.getActive().getSheetByName("一覧抽出");
  const Invoice = SpreadsheetApp.getActive().getSheetByName("請求書");
  //const Plist = SpreadsheetApp.getActive().getSheetByName("生徒マスタ");
  //const finalRow = Plist.getLastRow();
  const LastRow = ExportBilling.getLastRow();
  const Query = [];
  const SearchKey = ExportBilling.getRange(7,3,LastRow - 6,4).getValues();//C7～E末まで

//一覧から請求書に移す。仕様としては0を含む列を除外して請求書に書き込んでいく。

  for (let i = 0; i < SearchKey.length; i++){
  if (SearchKey[i][2]!== ""){
    if (SearchKey[i][2]!== 0){
    Query.push(SearchKey[i]);
    }

  }

}

Logger.log(SearchKey);
Logger.log(Query);

Invoice.getRange(13,2,20,4).clearContent();//B14E33
Invoice.getRange(13,2,Query.length,4).setValues(Query);

}

function sendpdf(){

  const ss = SpreadsheetApp.getActive().getSheetByName("生徒マスタ");
  const finalRow = ss.getLastRow();
  const StuMaster = ss.getRange(2,1,finalRow-1,13).getValues();//生徒マスタA2～Mまで 

  
  const Keyname = [];
  const MailAdd = [];


  for (let i = 0; i < StuMaster.length; i++){
        Keyname.push(StuMaster[i][1]);//氏名をつっこむ
        MailAdd.push(StuMaster[i][9]);//メールアドレスをつっこむ

  }

Logger.log(Keyname);
Logger.log(MailAdd);
  


}

function savepdf(Name){

  const ss = SpreadsheetApp.getActive().getSheetByName("生徒マスタ");
  const finalRow = ss.getLastRow();
  const StuMaster = ss.getRange(2,1,finalRow-1,15).getValues();//生徒マスタA2～Oまで 

  //const Name2 = ss.getRange("B2").getValue(); //テスト用取得キー
  
  const Keyname = [];//リンク取得用キー配列
  const Keyname2 = [];//メアド取得用キー配列
  
  const URLAdd = [];
  const MailAdd = [];
  
  const SearchKey = [];
  const MailKey = [];


  for (let i = 0; i < StuMaster.length; i++){
        Keyname.push(StuMaster[i][1]);//氏名をつっこむ
        URLAdd.push(StuMaster[i][14]);//O列フォルダ保存先リンクをつっこむ

        //SearchKey[0] = URLAdd[Keyname.indexOf(StuMaster[i][1])];

        Logger.log(Keyname);
        Logger.log(URLAdd);

        
  }


  for (let i = 0; i < StuMaster.length; i++){
        Keyname2.push(StuMaster[i][1]);//氏名をつっこむ
        MailAdd.push(StuMaster[i][9]);//メールアドレスをつっこむ
        Logger.log(Keyname2);
        Logger.log(MailAdd);      

  }



SearchKey[0] = URLAdd[Keyname.indexOf(Name)];//名前から保存フォルダ置き場を引っ張ってくる
Logger.log(SearchKey);

MailKey[0] = MailAdd[Keyname2.indexOf(Name)];//名前からメアドを引っ張ってくる
Logger.log(MailKey);

makepdf(SearchKey, MailKey);

}
function makepdf(SearchKey, MailKey){
  const ExportBilling = SpreadsheetApp.getActive().getSheetByName("一覧抽出");
  const Invoice = SpreadsheetApp.getActive().getSheetByName("請求書");
  const Title = ExportBilling.getRange(2,3).getValue(); //一覧抽出のC2セルの中身



//とりあえず請求書フォルダにじゃかじゃか作成していく。
//抽出した請求書を指定のフォルダにmoveで移動する。
//

// PDFの保存先となるフォルダID 確認方法は後述
  //const FolderId = "1ef7e300FKgJU7k67sweShh7xMggxeo3k";//請求書フォルダ直下
  const FolderId = SearchKey;
  
  // マイドライブ直下に保存したい場合は以下
  // var root= DriveApp.getRootFolder();
  // var folderid = root.getId();
  
  /////////////////////////////////////////////  
  // 現在開いているスプレッドシートをPDF化したい場合//
  ////////////////////////////////////////////
  // 現在開いているスプレッドシートを取得
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 現在開いているスプレッドシートのIDを取得
  const ssid = ss.getId();
  
  // 現在開いているスプレッドシートのシートIDを取得
  const Invoiceid = Invoice.getSheetId();
  // getActiveSheetの後の()を忘れると、TypeError: オブジェクト function getActiveSheet() {/* */} で関数 getSheetId が見つかりません。

  // ファイル名に使用する名前を取得
  const Name = ExportBilling.getRange(3,3).getValue();
  // スプレッドシートのC３の生徒名をファイル名用に取得。
  
  Logger.log(Invoiceid);

// PDFファイルの保存先となるフォルダをフォルダIDで指定
const Folder = DriveApp.getFolderById(FolderId);

// スプレッドシートをPDFにエクスポートするためのURL。このURLに色々なオプションを付けてPDFを作成
const url = "https://docs.google.com/spreadsheets/d/SSID/export?".replace("SSID", ssid)

// PDF作成のオプションを指定
  const opts = {
    exportFormat: "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
    format:       "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
    size:         "A4",     // 用紙サイズの指定 legal / letter / A4
    portrait:     "true",   // true → 縦向き、false → 横向き
    fitw:         "true",   // 幅を用紙に合わせるか
    sheetnames:   "false",  // シート名をPDF上部に表示するか
    printtitle:   "false",  // スプレッドシート名をPDF上部に表示するか
    pagenumbers:  "false",  // ページ番号の有無
    gridlines:    "false",  // グリッドラインの表示有無
    fzr:          "false",  // 固定行の表示有無
    gid:          Invoiceid   // シートIDを指定 sheetidは引数で取得
  };
  
  const url_ext = [];
  
  // 上記のoptsのオプション名と値を「=」で繋げて配列url_extに格納
  for( optName in opts ){
    url_ext.push( optName + "=" + opts[optName] );
  }

  // url_extの各要素を「&」で繋げる
  const options = url_ext.join("&");

  // optionsは以下のように作成しても同じです。
  // var ptions = 'exportFormat=pdf&format=pdf'
  // + '&size=A4'                       
  // + '&portrait=true'                    
  // + '&sheetnames=false&printtitle=false' 
  // + '&pagenumbers=false&gridlines=false' 
  // + '&fzr=false'                         
  // + '&gid=' + sheetid;

  // API使用のためのOAuth認証
  const token = ScriptApp.getOAuthToken();

    // PDF作成
    const response = UrlFetchApp.fetch(url + options, {
      headers: {
        'Authorization': 'Bearer ' +  token
      }
    });

    // 
    const blob = response.getBlob().setName(Name + '様_' + Title + '_ご請求書.pdf');

  //}

  //　PDFを指定したフォルダに保存
  Folder.createFile(blob);
  
const To = MailKey;
const Subject =Title + "ご請求書";
const Body =　ExportBilling.getRange(4,15).getValues();;//一覧抽出のO4にメール本文内容を記載している。

  //GmailApp.sendEmail(To,Subject,Body,{attachments: blob}) //メールを添付して送信
  GmailApp.createDraft(To,Subject,Body,{attachments: blob}) //メールの下書きを作成



}



