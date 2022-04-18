/**
 * スプレッドシートからpovoプロモコードを取得するLINE BOT
 */

// 定数
const LINE_BOT_ACCESS_TOKEN = CONFIG_SHEET.getRange(2, 2).getValue();
const LINE_REPLY_ENDPOINT = "https://api.line.me/v2/bot/message/reply"; // 応答メッセージ用のAPI URL

/**
 * doPost
 * ユーザーがLINEにメッセージ送信した時の処理
 **/
function doPost(e) {
  let contents = JSON.parse(e.postData.contents);
  let events = contents.events;
  console.log(contents);
  for (let i = 0; i < events.length; i++) {
    let event = events[i];
    let messages = [];
    if (isKnownUser(event.source.userId)) {
      messages = messages.concat(generateReplyMessagesToEvent(event));
    } else {
      // 事前設定済みのユーザー以外からの投稿には定型文で返す
      messages = messages.concat(
        generateReplyMessagesToUnknownUserEvent(event)
      );
    }

    // メッセージ返信
    if (messages.length) {
      console.log(messages);
      replyMessages(messages, event.replyToken);
    }
  }

  // https://mofumofupower.hatenablog.com/entry/2021/05/24/151430
  // webhookのテストで200を返すためにいるらしい
  return ContentService.createTextOutput(
    JSON.stringify({ content: "ok" })
  ).setMimeType(ContentService.MimeType.JSON);
}

/**
 * doGet
 * GETリクエストのハンドリング
 */
function doGet(e) {
  return ContentService.createTextOutput("SUCCESS");
}

/**
 * ユーザーからのメッセージへの応答メッセージを生成
 * event.type=message/postbackの場合で場合分け
 * @param {json} event: ユーザーからの投稿メッセージ
 * @returns {list} messages: 返信メッセージ
 */
function generateReplyMessagesToEvent(event) {
  let messages = [];
  if (event.type == "message") {
    messages = messages.concat(generateQuickReplyTopMessage(event));
  } else if (event.type == "postback") {
    // postback: https://developers.line.biz/ja/reference/messaging-api/#postback-event
    messages = messages.concat(generateMessagesToPostbackEvent(event));
  }
  return messages;
}

/**
 * 事前設定なしのユーザーへの応答
 * @returns {list}: 応答メッセージ
 */
function generateReplyMessagesToUnknownUserEvent() {
  return [
    {
      type: "text",
      text: "あなたは誰？",
    },
  ];
}

/**
 * ユーザーへの最初の応答メッセージを返却する
 *
 * @param {json} event
 * @returns {array}: 返信メッセージ
 */
function generateQuickReplyTopMessage(event) {
  return {
    type: "text",
    text: "メニューを選んでください",
    quickReply: {
      items: [
        {
          type: "action",
          action: {
            type: "postback",
            label: "取得",
            displayText: "プロモコード取得",
            data: JSON.stringify({
              state: "PROMOCODE",
            }),
          },
        },
        {
          type: "action",
          action: {
            type: "postback",
            label: "300MB",
            displayText: "300MBコード取得",
            data: JSON.stringify({
              state: "300MB",
            }),
          },
        },
        {
          type: "action",
          action: {
            type: "postback",
            label: "リスト取得",
            displayText: "リスト取得",
            data: JSON.stringify({
              state: "COUNT",
            }),
          },
        },
        {
          type: "action",
          action: {
            type: "postback",
            label: "キャンセル",
            displayText: "キャンセル",
            data: JSON.stringify({
              state: "CANCEL",
            }),
          },
        },
      ],
    },
  };
}

/**
 * postbackでユーザーが投稿したメッセージへの応答メッセージを生成する。
 * postbackメッセージに含まれるdata.stateの値によって返信内容を設定する。
 * @param {json} event
 * @returns {list}: 返信メッセージ
 */
function generateMessagesToPostbackEvent(event) {
  let data = JSON.parse(event.postback.data);
  let messages = [];

  if (data.state === "ROOT") {
    messages.push(generateQuickReplyTopMessage());
  } else if (data.state === "PROMOCODE") {
    messages.push(generateQuickReplyForGetPromoCode());
  } else if (data.state === "300MB") {
    messages = messages.concat(
      generateQuickReplyForGetSpecifyPromoCode("300MB")
    );
  } else if (data.state === "1GB") {
    messages = messages.concat(generateQuickReplyForGetSpecifyPromoCode("1GB"));
  } else if (data.state === "3GB") {
    messages = messages.concat(generateQuickReplyForGetSpecifyPromoCode("3GB"));
  } else if (data.state === "7GB") {
    messages = messages.concat(generateQuickReplyForGetSpecifyPromoCode("7GB"));
  } else if (data.state === "20GB") {
    messages = messages.concat(
      generateQuickReplyForGetSpecifyPromoCode("20GB")
    );
  } else if (data.state === "UNLIMITED") {
    messages = messages.concat(
      generateQuickReplyForGetSpecifyPromoCode("無制限")
    );
  } else if (data.state === "UNLIMITED_CHECK") {
    messages.push(generateQuickReplyForGetUnlimitedPromoCode());
  } else if (data.state === "COUNT") {
    messages.push(generateMessageForCountPromoCode());
  } else if (data.state === "USED_FLAG") {
    messages.push(generateMessageForUpdateUsedFlag(data.code));
  } else if (data.state === "CANCEL") {
    messages.push(generateMessageForCancel());
  }
  return messages;
}

/**
 * プロモコード取得時の返信メッセージを生成する
 * @returns {list}: 返信メッセージ
 */
function generateQuickReplyForGetPromoCode() {
  return {
    type: "text",
    text: "どのコードを取得しますか？",
    quickReply: {
      items: [
        {
          type: "action",
          action: {
            type: "postback",
            label: "300MB",
            displayText: "300MB",
            data: JSON.stringify({
              state: "300MB",
            }),
          },
        },
        {
          type: "action",
          action: {
            type: "postback",
            label: "1GB",
            displayText: "1GB",
            data: JSON.stringify({
              state: "1GB",
            }),
          },
        },
        {
          type: "action",
          action: {
            type: "postback",
            label: "3GB",
            displayText: "3GB",
            data: JSON.stringify({
              state: "3GB",
            }),
          },
        },
        {
          type: "action",
          action: {
            type: "postback",
            label: "7GB",
            displayText: "7GB",
            data: JSON.stringify({
              state: "7GB",
            }),
          },
        },
        {
          type: "action",
          action: {
            type: "postback",
            label: "20GB",
            displayText: "20GB",
            data: JSON.stringify({
              state: "20GB",
            }),
          },
        },
        {
          type: "action",
          action: {
            type: "postback",
            label: "無制限",
            displayText: "無制限",
            data: JSON.stringify({
              state: "UNLIMITED_CHECK",
            }),
          },
        },
        {
          type: "action",
          action: {
            type: "postback",
            label: "キャンセル",
            displayText: "キャンセル",
            data: JSON.stringify({
              state: "CANCEL",
            }),
          },
        },
      ],
    },
  };
}

/**
 * キャンセル操作時の応答メッセージ
 * @returns {array}: 返信メッセージ
 */
function generateMessageForCancel() {
  return {
    type: "text",
    text: "操作をキャンセルしました",
  };
}

/**
 * 無制限コードを利用しようとした場合の返信メッセージ
 * @returns {array}: 返信メッセージ
 */
function generateQuickReplyForGetUnlimitedPromoCode() {
  return {
    type: "text",
    text: "無制限のコードは24H使い放題ですが、貴重です。本当に使いますか？",
    quickReply: {
      items: [
        {
          type: "action",
          action: {
            type: "postback",
            label: "使いたい",
            displayText: "使いたい",
            data: JSON.stringify({
              state: "UNLIMITED",
            }),
          },
        },
        {
          type: "action",
          action: {
            type: "postback",
            label: "やめとく",
            displayText: "やめとく",
            data: JSON.stringify({
              state: "CANCEL",
            }),
          },
        },
      ],
    },
  };
}

/**
 * 指定されたプロモコード容量に対応するプロモコードと、使用済みにするかどうかの確認メッセージを生成する
 * @param {str} promoCodeAmount : プロモコード容量を表す文字列
 * @returns {list}: 返信メッセージリスト
 */
function generateQuickReplyForGetSpecifyPromoCode(promoCodeAmount) {
  let messages = [];
  promoCode = getReturnPromoCode(promoCodeAmount);
  messages.push({
    type: "text",
    text: promoCode || "対象のコードがありませんでした",
  });
  if (promoCode) {
    messages.push({
      type: "text",
      text: "使用済みにしますか？",
      quickReply: {
        items: [
          {
            type: "action",
            action: {
              type: "postback",
              label: "使用済みにする",
              displayText: "使用済みにする",
              data: JSON.stringify({
                state: "USED_FLAG",
                code: promoCode,
              }),
            },
          },
          {
            type: "action",
            action: {
              type: "postback",
              label: "キャンセル",
              displayText: "キャンセル",
              data: JSON.stringify({
                state: "CANCEL",
              }),
            },
          },
        ],
      },
    });
  }
  return messages;
}

/**
 * 指定されたプロモコードを使用済みにし、完了メッセージを返却する
 * @param {str} promoCode : プロモコード
 * @returns {array}: 返信メッセージ
 */
function generateMessageForUpdateUsedFlag(promoCode) {
  updateUsedFlagTrue(promoCode);
  return {
    type: "text",
    text: promoCode + "を使用済みにしました",
  };
}

/**
 * 利用可能なプロモコードの集計結果を取得し、返却する
 * @returns {array}: 返信メッセージ
 */
function generateMessageForCountPromoCode() {
  message = getPromoCodeList();
  return {
    type: "text",
    text: message,
  };
}

/**
 * LINEトークへメッセージを送信する
 * @param {*} messages : 返信メッセージ
 * @param {*} replyToken : 返信用トークン
 */
function replyMessages(messages, replyToken) {
  UrlFetchApp.fetch(LINE_REPLY_ENDPOINT, {
    headers: {
      "Content-Type": "application/json; charset=UTF-8",
      Authorization: "Bearer " + LINE_BOT_ACCESS_TOKEN,
    },
    method: "post",
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: messages || [
        {
          type: "text",
          text: "Error Occured",
        },
      ],
    }),
  });
}

/**
 * プロモコード容量を条件にスプレッドシートから「未使用かつ利用期限が最も近い」プロモコードを取得する
 * スプレッドシートのQUERY関数を用いている
 *   生成される数式例：=iferror(QUERY('コード一覧'!A2:G,"select C where G = FALSE and D like '300MB%' order by F asc limit 1"),"")
 * @param {str} promoCodeAmount : プロモコード容量
 * @returns {str}: プロモコード
 */
function getReturnPromoCode(promoCodeAmount) {
  let dummyCell = TEMPORARY_SHEET.getRange("A1"); // 使っていないセルを取得
  let startCol = getColName(INSERT_DATE_COLUMN_IDX);
  let endCol = getColName(USED_FLAG_COLUMN_IDX);
  let queryString =
    "=iferror(QUERY('" +
    MAIN_SHEET_NAME +
    "'!" +
    startCol +
    "2:" +
    endCol +
    ',"select ' +
    getColName(PROMO_CODE_COLUMN_IDX) +
    " where " +
    getColName(USED_FLAG_COLUMN_IDX) +
    " = FALSE and " +
    getColName(PROMO_CODE_AMOUNT_COLUMN_IDX) +
    " like '" +
    promoCodeAmount +
    "%' order by " +
    getColName(LIMIT_DATE_COLUMN_IDX) +
    ' asc limit 1"),"")';
  console.log("execute QUERY: " + queryString);
  dummyCell.setFormula(queryString); // 関数を設定して演算
  let code = dummyCell.getValue(); // 演算結果を取り出し
  dummyCell.clear(); // 演算で利用したしたセルを初期状態に戻す
  console.log("return: " + code);
  return code;
}

/**
 * 未使用のプロモコード容量とその数の集計結果を取得する。
 * @returns {atr}: プロモコード集計結果
 * スプレッドシートのQUERY関数を用いている
 *   生成される数式例：=iferror(QUERY('コード一覧'!A2:G,"select D, count(D) where G = FALSE group by D LABEL count(D) ''"))
 */
function getPromoCodeList() {
  let dummyCell = TEMPORARY_SHEET.getRange("A1"); // 使っていないセルを取得
  let startCol = getColName(INSERT_DATE_COLUMN_IDX);
  let endCol = getColName(USED_FLAG_COLUMN_IDX);
  let message = "";
  let tmpMessage = [];
  let queryString =
    "=iferror(QUERY('" +
    MAIN_SHEET_NAME +
    "'!" +
    startCol +
    "2:" +
    endCol +
    ',"select ' +
    getColName(PROMO_CODE_AMOUNT_COLUMN_IDX) +
    ", count(" +
    getColName(PROMO_CODE_AMOUNT_COLUMN_IDX) +
    ")" +
    " where " +
    getColName(USED_FLAG_COLUMN_IDX) +
    " = FALSE group by " +
    getColName(PROMO_CODE_AMOUNT_COLUMN_IDX) +
    " LABEL count(" +
    getColName(PROMO_CODE_AMOUNT_COLUMN_IDX) +
    ") ''" +
    '"))';
  console.log("execute QUERY: " + queryString);
  dummyCell.setFormula(queryString); // 関数を設定して演算
  let values = TEMPORARY_SHEET.getRange(
    1,
    1,
    TEMPORARY_SHEET.getLastRow(),
    TEMPORARY_SHEET.getLastColumn()
  ).getValues();
  for (let row in values) {
    tmpMessage.push(values[row].join(": ") + "個");
  }
  message = tmpMessage.join("\n");
  dummyCell.clear(); // 演算で利用したしたセルを初期状態に戻す
  console.log("return: " + message);
  return message;
}

/**
 * 引数のプロモコードが入力されているスプレッドシートの行番号を返却する
 * 複数ある場合には最も若い番号が返却される
 * @param {str} promoCode : プロモコード
 * @returns {int}: 行番号
 */
function findUpdateTargetRow(promoCode) {
  let lastRow = SHEET.getDataRange().getLastRow(); //対象となるシートの最終行を取得
  for (let i = 1; i <= lastRow; i++) {
    if (
      SHEET.getRange(i, PROMO_CODE_COLUMN_IDX).getValue() === promoCode &&
      SHEET.getRange(i, USED_FLAG_COLUMN_IDX).getValue() === false
    ) {
      return i;
    }
  }
  return 0;
}

/**
 * 引数に渡した数値に対応するカラム列番号（英字）を返却する
 * @param {int} num : カラム番号（数字）
 * @returns {str} : カラム番号（英字）
 */
function getColName(num) {
  let result = SHEET.getRange(1, num);
  colName = result.getA1Notation().replace(/\d/, "");

  return colName;
}

/**
 * 引数に渡されたプロモコードを使用済みに変更する
 * @param {str} promoCode : プロモコード
 */
function updateUsedFlagTrue(promoCode) {
  targetRow = findUpdateTargetRow(promoCode);
  SHEET.getRange(targetRow, USED_FLAG_COLUMN_IDX).setValue(true); // 使用期限
}

/**
 * 引数に渡されたuserIdが事前設定済みかどうかを判断する
 * @param {str} userId : LINEユーザーID
 * @returns {boolean}
 */
function isKnownUser(userId) {
  let lastRow = USER_SHEET.getDataRange().getLastRow(); //対象となるシートの最終行を取得
  console.log("userId: " + userId);
  for (let i = 1; i <= lastRow; i++) {
    if (USER_SHEET.getRange(i, 1).getValue() === userId) {
      return true;
    }
  }
  return false;
}
