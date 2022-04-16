# ギガ活プロモコード管理
povo2.0のギガ活で得たプロモコードを管理するためのツール。  
Google Apps Script（GAS）で稼動する。

## Script
|ファイル|処理概要|
|:--|:--|
|[getPromoCode.gs.js](./getPromoCode.gs.js)|Gmailで受信したプロモコードをスプレッドシートへ転記する。|
|[notifyLimitDate.gs.js](./notifyLimitDate.gs.js)|スプレッドシートに記載されたプロモコードの利用期限の数日前になったらLINEに通知する|

## 使い方
以下のようなスプレッドシートを用意する。
- [コード一覧]シートに以下のカラムを用意しておく。
  - 記入日
  - メール受信日
  - コード
  - 容量
  - 使用回数（同じプロモコードを複数回使用できる場合がある）
  - 利用期限
  - 使用済み
- [設定]シートのB1セルに[LINE Notify](https://notify-bot.line.me/ja/)のトークンを記載する。

![\[コード一覧\]シート](https://github.com/ck-fm0211/gigakatsu-promocode-manage/blob/images/sheet1.png)  
![\[設定\]シート](https://github.com/ck-fm0211/gigakatsu-promocode-manage/blob/images/sheet2.png)

### `容量`の記載内容
プロモコードから容量を判断している。  
※povo2.0公式には以下のような記載がある。
```
■利用開始後のデータ消費期限
300MB：3日間、1GB：7日間、3GB,20GB：30日間
※各コードに含まれる冒頭文字がデータ容量を示しています
```

|コード値|スプレッドシートへの記載内容|
|:--|:--|
|`300MBxxxxxxxxx`|`300MB/3Days`|
|`U24xxxxxxxxx`|`無制限/24H`|
|`1GBxxxxxxxxx`|`1GB/7Days`|
|`3GBxxxxxxxxx`|`3GB/30Days`|
|`20GBxxxxxxxxx`|`20GB/30Days`|
|上記以外|`!!!!要確認!!!!`|

## コード管理
claspでGASからローカルにcloneし、それをGitHub（本リポジトリ）へpushしている。

