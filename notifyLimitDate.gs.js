/**
 * 指定の日付以内で期限が切れる未使用のプロモーションコードがあった場合に通知する
 */

// 処理日
let dt = new Date();
const limitDateCount = 7 // 7日以内
dt.setDate(dt.getDate()+limitDateCount)

/**
 * スプレッドシートから指定の日付以内で期限が切れる未使用のプロモーションコードを抽出し、LINEに通知する
 */
function notifyLimitDate() {
  console.log("[START]データ抽出");
  let lastRow = sheet.getLastRow();
  let lastColumn = sheet.getLastColumn();
  let data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

  let limitList = [];
  data.some(function(value) {
    // 期限が指定の日付以内 かつ チェックボックスにチェックが付いていない
    if((value[5].getTime()<=dt.getTime())&&(!value[6])){
      limitList.push([Utilities.formatDate(value[5], "JST", "yyyy/MM/dd"), value[2]])
    }
  });
  // 期限が近いものがある場合はLINEに通知する
  if(limitList.length > 0){
    console.log("[START]LINE通知");
    sendLINE(limitList);
  }else{
    console.log("対象のプロモコードはありません");
  };
}

/**
 * 受け取ったリストをもとにLINEに通知する
 */
function sendLINE(limitList) {

  let limitMessageTemp = [];
  for(let i = 0; i < limitList.length; i++) {
    limitMessageTemp.push(limitList[i].join("："));
  }

  let messageText = "\n" + String.fromCodePoint("0x26A0") + "有効期限が" + limitDateCount + "日以内のものがあります"
  messageText += "\n\n"
  messageText += "■期限日：対象コード"
  messageText += limitMessageTemp.join("\n")

  // LINEから取得したトークン
  let token = configSheet.getRange(1, 2).getValue()
  let options = {
    "method" : "post",
    "headers" : {
      "Authorization" : "Bearer "+ token
    },
    "payload" : {
      "message" : messageText
    }
  }

  let url  = "https://notify-api.line.me/api/notify"
  UrlFetchApp.fetch(url, options)
}