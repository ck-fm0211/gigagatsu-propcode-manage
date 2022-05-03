/**
 * ギガ活のメールからコードを抽出し、スプレッドシートに書き出す
 */

// スプレッドシート
const MAIN_SHEET_NAME = "コード一覧";
const CONFIG_SHEET_NAME = "設定";
const USER_SHEET_NAME = "ユーザーリスト";
const TEMPORARY_SHEET_NAME = "work";
const SHEET = SpreadsheetApp.getActive().getSheetByName(MAIN_SHEET_NAME);
const CONFIG_SHEET =
  SpreadsheetApp.getActive().getSheetByName(CONFIG_SHEET_NAME);
const USER_SHEET = SpreadsheetApp.getActive().getSheetByName(USER_SHEET_NAME);
const TEMPORARY_SHEET =
  SpreadsheetApp.getActive().getSheetByName(TEMPORARY_SHEET_NAME);

// 各列番号
const INSERT_DATE_COLUMN_IDX = 1; // 記入日
const MAIL_RECEIVED_DATE_COLUMN_IDX = 2; // メール受信日
const PROMO_CODE_COLUMN_IDX = 3; // プロモコード
const PROMO_CODE_AMOUNT_COLUMN_IDX = 4; // 容量
const NUMBER_OF_TIMES_COLUMN_IDX = 5; // 使用回数
const LIMIT_DATE_COLUMN_IDX = 6; // 利用期限
const USED_FLAG_COLUMN_IDX = 7; // 使用済みフラグ

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
  let threads = GmailApp.search("label:00_ギガ活  -label:00_ギガ活/処理済み");

  console.log("[START]スレッド処理");
  // スレッドを一つずつ取り出す
  threads.forEach(function (thread) {
    // スレッド内のメール一覧を取得
    let messages = thread.getMessages();

    // メールを一つずつ取り出す
    messages.forEach(function (message) {
      // メールからコード値と利用期限を抽出する
      let values = extractCode(message);
      // メールを処理しているのにプロモコードが取れなかったらアラート
      if (values["codes"].length === 0) {
        console.log("[ERROR]データ抽出");
        throw new Error("プロモコードの抽出に失敗しました");
      }
      // スプレッドシートに書き出す
      console.log("[START]スプレッドシート書き込み");
      writeSheet(values);
    });

    // 重複を排除する
    console.log("[START]重複レコード削除");
    removeDupilicatedRecords();

    // スプレッドシートに条件付き書式を設定
    console.log("[START]条件付き書式更新");
    upsertConditionalFormatRule();

    // スレッドに処理済みラベルを付ける
    console.log("[START]ラベル追加");
    let label = GmailApp.getUserLabelByName("00_ギガ活/処理済み");
    thread.addLabel(label);
  });
}

/**
 * メール本文からプロモコードと利用期限、メール受信日を抽出
 * 抽出する際、プロモコードは<strong>タグの中にあるため、html形式のbodyから正規表現で抽出する
 * 入力期限は「コードの入力期限」というメッセージのあとに記載されているため、プレーンテキストのメールbodyから抽出する
 */
function extractCode(message) {
  let body = message.getBody();
  let plainBody = message.getPlainBody();

  let ret = {};

  // 形式:300MBXXXXXXXXX(300MBコード) or U24H10TXXXXXXXXX（24H使い放題10回分）
  let codes = body.match(
    /([1-9]{1}[A-Z0-9]{13}|U24H[0-9]{1,2}T[A-Z0-9]{9})(?=.*<\/strong>)/g
  );
  ret["codes"] = codes;
  // 入力期限
  let limitDate = plainBody.match(
    /(?<=コードの入力期限\n*)20[0-9]{2}\/[0-9]{2}\/[0-9]{2}/g
  );
  ret["limitDate"] = limitDate;
  //使用回数制限
  let limitNumberOfTimes = plainBody.match(
    /(?<=コードの利用回数\n*)[0-9]{1,9}(?=回)/g
  );
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
function writeSheet(values) {
  // 処理日
  const TODAY = new Date();

  // 最終行を取得
  let lastRow = SHEET.getLastRow() + 1;

  // セルを取得して値を転記
  // code["limitNumberOfTimes"]の回数だけ繰り返す。
  for (let i = 1; i <= Number(values["limitNumberOfTimes"]); i++) {
    let index = 0;
    for (let code of values["codes"]) {
      SHEET.getRange(lastRow + index, INSERT_DATE_COLUMN_IDX).setValue(TODAY); // 記入日
      SHEET.getRange(lastRow + index, MAIL_RECEIVED_DATE_COLUMN_IDX).setValue(
        values["receiveDate"]
      ); // メール受信日
      SHEET.getRange(lastRow + index, PROMO_CODE_COLUMN_IDX).setValue(code); // プロモコード
      SHEET.getRange(lastRow + index, PROMO_CODE_AMOUNT_COLUMN_IDX).setValue(
        getCodeAmount(code)
      ); // 容量
      SHEET.getRange(lastRow + index, NUMBER_OF_TIMES_COLUMN_IDX).setValue(
        i + "回目"
      ); // 使用回数
      SHEET.getRange(lastRow + index, LIMIT_DATE_COLUMN_IDX).setValue(
        values["limitDate"]
      ); // 使用期限
      SHEET.getRange(lastRow + index, USED_FLAG_COLUMN_IDX).insertCheckboxes(); // 使用済みチェックボックス
      index += 1;
    }
    lastRow += 1;
  }
}

/**
 * スプレッドシートに条件付き書式を設定する
 * 条件：使用済み列（F列）がTRUE（チェックON） 書式: グレーアウト
 */
function upsertConditionalFormatRule() {
  let conditionalFormatRuleUsedCodes = getConditionalFormatRuleUsedCodes();
  let conditionalFormatRuleUnlimitedPromoCodes =
    getConditionalFormatRuleUnlimitedPromoCodes();
  let conditionalFormatRules = [
    conditionalFormatRuleUsedCodes,
    conditionalFormatRuleUnlimitedPromoCodes,
  ];
  // 条件付き書式を設定
  SHEET.setConditionalFormatRules(conditionalFormatRules);
}

/**
 * 条件付き書式オブジェクトを返却する
 * 条件：使用済み列（F列）がTRUE（チェックON） 書式: グレーアウト
 */
function getConditionalFormatRuleUsedCodes() {
  let conditionalFormatRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$G1=TRUE")
    .setBackground("#474A4D")
    .setRanges([
      SHEET.getRange(1, 1, SHEET.getLastRow() + 1, SHEET.getLastColumn()),
    ])
    .build();
  return conditionalFormatRule;
}

/**
 * 条件付き書式オブジェクトを返却する
 * 条件：容量列（D列）に「無制限」の文字がある 書式: 背景色を薄いグレー
 */
function getConditionalFormatRuleUnlimitedPromoCodes() {
  let conditionalFormatRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=find("無制限",$D1)')
    .setBackground("#DCDDDD")
    .setRanges([
      SHEET.getRange(1, 1, SHEET.getLastRow() + 1, SHEET.getLastColumn()),
    ])
    .build();
  return conditionalFormatRule;
}

/**
 * 重複レコードを削除する
 * 受信メールがスレッドにまとめられてしまった場合にデータ重複することがある。
 * 記入日列（A列）を除いて重複を排除する。
 */
function removeDupilicatedRecords() {
  let range = SHEET.getRange(
    1,
    1,
    SHEET.getLastRow() + 1,
    SHEET.getLastColumn() + 1
  );

  range.removeDuplicates([
    MAIL_RECEIVED_DATE_COLUMN_IDX,
    PROMO_CODE_COLUMN_IDX,
    PROMO_CODE_AMOUNT_COLUMN_IDX,
    NUMBER_OF_TIMES_COLUMN_IDX,
    USED_FLAG_COLUMN_IDX,
  ]);
}

/**
 * プロモコードからコード容量を返却する。
 * プロモコードの先頭文字列に対して以下の文字列を返却
 *  300MB -> 300MB/3Days
 *  U24   -> 無制限/24H
 *  1GB   -> 1GB/7Days
 *  3GB   -> 3GB/30Days
 *  20MB  -> 20GB/30Days
 */
function getCodeAmount(code) {
  if (code.slice(0, 5) == "300MB") {
    return "300MB/3Days";
  }

  if (code.slice(0, 3) == "U24") {
    return "無制限/24H";
  }

  if (code.slice(0, 3) == "1GB") {
    return "1GB/7Days";
  }

  if (code.slice(0, 3) == "3GB") {
    return "3GB/30Days";
  }

  if (code.slice(0, 4) == "20GB") {
    return "20GB/30Days";
  }

  return "!!!!要確認!!!!";
}
