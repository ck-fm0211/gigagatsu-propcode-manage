/**
 * ギガ活のメールからコードを抽出し、スプレッドシートに書き出す
 */

// スプレッドシート
const sheet = SpreadsheetApp.getActive().getSheetByName('コード一覧');
const configSheet = SpreadsheetApp.getActive().getSheetByName('設定');

/**
 * 本体
 * 1. ギガ活のラベルがついたメールを抽出
 * 2. メール本文からプロモコードと利用期限を抽出
 * 3. スプレッドシートに記入
 *    - レコード挿入
 *    - 重複レコード削除
 *    - 条件付き書式を設定
 * 4. 処理したメールに処理済みラベルを付与
 */
function getPromotionCodes() {

  console.log("[START]メール取得");
  // 検索条件に該当するスレッド一覧を取得
  let threads = GmailApp.search('label:00_ギガ活  -label:00_ギガ活/処理済み');

  console.log("[START]スレッド処理");
  // スレッドを一つずつ取り出す
  threads.forEach(function(thread) {
    // スレッド内のメール一覧を取得
    let messages = thread.getMessages();

    // メールを一つずつ取り出す
    messages.forEach(function(message) {
      // メールからコード値と利用期限を抽出する
      let values = extractCode(message)
      // メールを処理しているのにプロモコードが取れなかったらアラート
      if(values["codes"].length === 0){
        console.log("[ERROR]データ抽出");
        throw new Error("プロモコードの抽出に失敗しました")
      }
      // スプレッドシートに書き出す
      console.log("[START]スプレッドシート書き込み");
      writeSheet(values)
    });

    // 重複を排除する
    console.log("[START]重複レコード削除");
    removeDupilicatedRecords();

    // スプレッドシートに条件付き書式を設定
    console.log("[START]条件付き書式更新");
    upsertConditionalFormatRule();

    // スレッドに処理済みラベルを付ける
    console.log("[START]ラベル追加");
    let label = GmailApp.getUserLabelByName('00_ギガ活/処理済み');
    thread.addLabel(label);
  });
}

/**
 * メール本文からプロモコードと利用期限、メール受信日を抽出
 * 抽出する際、プロモコードは<strong>タグの中にあるため、html形式のbodyから正規表現で抽出する
 * 入力期限は「コードの入力期限」というメッセージのあとに記載されているため、プレーンテキストのメールbodyから抽出する
 */
function extractCode(message){
  let body = message.getBody();
  let plainBody = message.getPlainBody();

  let ret = {};

  // 形式:300MBXXXXXXXXX(300MBコード) or U24H10TXXXXXXXXX（24H使い放題10回分）
  let codes = body.match(/([1-9]{1}[A-Z0-9]{13}|U24H[0-9]{1,2}T[A-Z0-9]{9})(?=.*<\/strong>)/g);
  ret["codes"] = codes;
  // 入力期限
  let limitDate = plainBody.match(/(?<=コードの入力期限\n*)20[0-9]{2}\/[0-9]{2}\/[0-9]{2}/g);
  ret["limitDate"] = limitDate;
  //使用回数制限
  let limitNumberOfTimes = plainBody.match(/(?<=コードの利用回数\n*)[0-9]{1,9}(?=回)/g);
  // nullのときは1、そうでないときはmatchした値
  ret["limitNumberOfTimes"] = !limitNumberOfTimes ? ["1"] : limitNumberOfTimes;
  // 受信日
  let receiveDate = message.getDate();
  ret["receiveDate"] = receiveDate;
  return ret;
}

/**
 * スプレッドシートに書き込む
 */
function writeSheet(values){
  // 処理日
  const today = new Date();

  // 最終行を取得
  let lastRow = sheet.getLastRow() + 1;

  // セルを取得して値を転記
  // code["limitNumberOfTimes"]の回数だけ繰り返す。
  for(let i = 1; i <= Number(values["limitNumberOfTimes"]); i++){
    let index = 0
    for(let code of values["codes"]){
      sheet.getRange(lastRow + index, 1).setValue(today);                  // 記入日
      sheet.getRange(lastRow + index, 2).setValue(values["receiveDate"]);  // メール受信日
      sheet.getRange(lastRow + index, 3).setValue(code);                   // プロモコード
      // sheet.getRange(lastRow + index, 4).setValue(email[1]);            // 容量 ※コードの仕様がわからんと入れられない
      sheet.getRange(lastRow + index, 5).setValue(i + "回目");              // 使用回数
      sheet.getRange(lastRow + index, 6).setValue(values["limitDate"]);    // 使用期限
      sheet.getRange(lastRow + index, 7).insertCheckboxes();               // 使用済みチェックボックス
      index += 1;
    }
    lastRow += 1;
  }
}

/**
 * スプレッドシートに条件付き書式を設定する
 * 条件：使用済み列（F列）がTRUE（チェックON） 書式: グレーアウト
 */
function upsertConditionalFormatRule(){
  let conditionalFormatRule = SpreadsheetApp.newConditionalFormatRule()
  .whenFormulaSatisfied('=$G1=TRUE')
  .setBackground('#B7B7B7')
  .setRanges([sheet.getRange(1,1,sheet.getLastRow() + 1, 7)])
  .build()
  let conditionalFormatRules = sheet.getConditionalFormatRules();
  if(conditionalFormatRules.length > 0){
    // 既存の条件付き書式と差し替え
    conditionalFormatRules.splice(conditionalFormatRules.length - 1, 1, conditionalFormatRule);
  }else{
    // 初回起動時には条件付き書式がないので直接push
    conditionalFormatRules.push(conditionalFormatRule)
  }
  // 条件付き書式を設定
  sheet.setConditionalFormatRules(conditionalFormatRules);
}

/**
 * 重複レコードを削除する
 * 受信メールがスレッドにまとめられてしまった場合にデータ重複することがある。
 * 記入日列（A列）を除いて重複を排除する。
 */
function removeDupilicatedRecords(){
  let range = sheet.getRange(1,1,sheet.getLastRow() + 1, 7);

  range.removeDuplicates([2,3,4,5,7]);
}
