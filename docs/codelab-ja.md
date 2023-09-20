authors: Yuki Tanabe
summary: Apps Script ハンズオン
id: ja
categories: apps-script,gsuite,javascript
environments: Web
status: Published
url: https://tanabee.dev/gas-codelab/ja/
analytics account: UA-141719032-2
feedback link: https://github.com/tanabee/gas-codelab/issues

# Apps Script ハンズオン

## 概要
Duration: 0:01:00

この Codelab では、 Google Apps Script を用いて Gmail のメッセージ一覧を Google Spreadsheet に抽出するアプリケーションを作成します。
![Output](img/en/output.png)

### ソースコード

この Codelab の最終的なソースコードは GitHub の [tanabee/gas-codelab](https://github.com/tanabee/gas-codelab) から取得できます。

### 対象者

- Google Apps Script 初心者
- JavaScript 経験者

### ユースケース

この Codelab は以下のようなケースを想定しています。

- 特定のメーリングリスト、もしくはメールアドレスへのユーザーからの問い合わせの集計、分析 ( 参考: [G Suite x Zendesk API で問い合わせの分析・可視化ツールを作ってみた](https://medium.com/google-cloud-jp/9aa831c1bf0e) )
- REST API は提供していないがメール通知機能を提供しているアプリケーションと Google Apps のシステム連携

## 準備
Duration: 0:01:00

Gmail もしくは G Suite のアカウントが必要です。もしまだアカウントがない場合には [Gmail アカウントを作成](https://accounts.google.com/signup)してください。

Negative
: Gmail アカウントを作成した場合には、受信済みのメールを用意するために以下の手順に沿ってください。

### 1. mail-to-me アプリケーションを開く

[mail-to-me](https://script.google.com/d/1p2B0NwXg3APMc-nh-fRWuMX6P_W5eQijsi4CVEHLOOq7_zJmQbWyZAWl/edit?usp=sharing) アプリケーションをブラウザで開く。

### 2. コピーを作成

**ファイル** > **コピーを作成** を選択。

### 3. mail-to-me の実行

**実行** ボタンをクリック。Gmail を送信する権限をスクリプトに与えるため、ポップアップに従って許可を与えます。

### 4. Gmail を確認

![mail-to-me application](img/ja/mail-to-me.png)
[Gmail](https://mail.google.com) をブラウザで開き、メールが受信されていることを確認する。以上で準備は完了です。

## プロジェクトの作成
Duration: 0:02:00

Apps Script のプロジェクトを作成します。まず [drive.google.com](https://drive.google.com) にアクセスし、スプレッドシートファイルを作成します。

![Create a new drive file](img/ja/create-drive-file.png)
**新規** をクリック

![Create a new Spreadsheet file](img/ja/create-spreadsheet-file.png)
**Google スプレッドシート** を選択。スプレッドシートファイルが作成されることを確認します。

![Rename Spreadsheet file](img/ja/rename-spreadsheet-file.png)
ファイル名を選択し、名前（gmail-to-spreadsheet など）を入力します。

![Select Script editor menu](img/ja/select-script-editor.png)
**ツール** > **スクリプトエディタ** を選択すると Google Apps Script プロジェクトが作成されます。

![Rename Google Apps Script project](img/ja/rename-gas-project.png)
プロジェクト名を選択し、名前（gmail-to-spreadsheet など）を入力します。

![Input Google Apps Script project name](img/ja/input-project-name.png)
プロジェクト名を入力して **OK** ボタンをクリックするとプロジェクトが保存されます。トーストが表示されてから、非表示になったところで保存が完了します。

これでスクリプトを実行できるようになりました。次のセクションでスクリプトを実行していきます。

## スクリプトの実行
Duration: 0:02:00

既存のコードを削除して、下記のコードを実装します。

```JavaScript
function main() {
  Logger.log('Hello Google Apps Script!');
}
```

`Logger.log()` 関数は Google Apps Script のログ出力関数です。

![Run script](img/ja/run-script.png)

**実行** ボタンをクリックします。

![View the script logs](img/ja/view-logs.png)
**表示** > **ログ** を選択すると出力されたログを確認できます。

![log viewer](img/ja/log-viewer.png)
`Hello Google Apps Script!` と出力されているのが確認できました。

通常の JavaScript と同様に `console.log()` 関数を使うこともできますが、`console.log()` 関数を利用するためには Google Cloud Platform のプロジェクトを作成する手続きを要するため、この Codelab では簡単のために `Logger.log()` 関数を利用します。

## GmailApp クラス
Duration: 0:03:00

次に、 GmailApp クラスを見てみます。GmailApp クラスとメソッドについては[公式リファレンス](https://developers.google.com/apps-script/reference/) から確認できます。[GmailApp](https://developers.google.com/apps-script/reference/gmail/gmail-app) のドキュメントを見てみます。Gmail に関連するクラス (e.g. [GmailMessage](https://developers.google.com/apps-script/reference/gmail/gmail-message), [GmailThread](https://developers.google.com/apps-script/reference/gmail/gmail-thread) ) やメソッド (e.g. [search](https://developers.google.com/apps-script/reference/gmail/gmail-app#searchquery,-start,-max), [sendEmail](https://developers.google.com/apps-script/reference/gmail/gmail-app#sendemailrecipient,-subject,-body,-options)) が確認できます。

![GmailApp reference](img/en/reference-gmailapp.png)

今回は [Gmail.search](https://developers.google.com/apps-script/reference/gmail/gmail-app#searchquery,-start,-max) を使ってメール一覧を取得します。

## GmailThead の取得
Duration: 0:03:00

ではスクリプトを実装していきます。[Gmail.search](https://developers.google.com/apps-script/reference/gmail/gmail-app#searchquery,-start,-max) メソッドを使って Gmail のスレッド一覧を取得します。

```JavaScript
function main() {
  var searchText = '';// You can set value.
  var threads = GmailApp.search(searchText, 0, 5);
  Logger.log(threads);
}
```

送信元や件名、日時など、検索条件を設定したい場合には searchText に値をセットしてください。公式サポートページで[検索演算子](https://support.google.com/mail/answer/7190?hl=ja)についてまとめられています。

**実行** ボタンをクリックして実行します。

![Authorization popup](img/ja/authorization-required.png)
このスクリプトに Gmail の操作を認可するため、ポップアップが表示されます。Gmail リソースへのアクセスを認可する必要があります。 **許可を確認** をクリックします。

![Choose an account](img/ja/choose-account.png)
この Codelab で使っているアカウントを選択します。

![Application verification](img/ja/verify-app.png)
![Application verification advanced](img/ja/verify-app-advanced.png)
このアプリケーションを確認するために、 **詳細** を選択し **... (安全ではないページ) に移動** と書かれたリンクをクリックします。もし、この画面が表示されない場合には、この手順はスキップしてください。

![Allow authentication](img/ja/allow-auth.png)
このアプリケーションに与える必要のあるスコープが表示されます。 **許可** ボタンをクリック後、スクリプトエディタに戻ってスクリプトが実行されます。

![Allow authentication](img/ja/gmail-threads.png)
ログを見てみましょう。GmailThread の配列が出力されています。たった 5 行のコードでメールの一覧を取得して出力することができました！この認可のフローのために、通常実装するのが大変な認証周りのコードを書く必要がなくなるため、非常に簡単にアプリケーション連携ができるのです。

## GmailMessage のパース
Duration: 0:03:00

直前のセクションで Gmail のスレッド一覧を取得できました。今回はメッセージの件名、本文、送信元、送信先、日時を取得します。そのため、GmailThread からそれに紐づくメッセージを取得します。

![GmailThread.getMessages reference](img/en/reference-getmessages.png)
Apps Script のリファレンスを見てみましょう。 GmailThread クラスには [getMessages()](https://developers.google.com/apps-script/reference/gmail/gmail-thread#getmessages) メソッドが用意されており GmailMessage の配列を返します。 **GmailMessage** のリンクをクリックしてそのメソッドを確認しましょう。

![GmailMessage reference](img/en/reference-gmailmessage.png)
GmailMessage クラスには `getSubject`, `getBody`, `getFrom`, `getTo`, `getDate` など多数の取得系メソッドが用意されています。 `GmailThread.getMessages()` の結果を用いて求める値が取得できそうです。

今回はスレッド内の最初のメッセージを使います。以下のコードを実装してメッセージを取得してログを見てみましょう。

```JavaScript
function main() {
  var searchText = '';
  var threads = GmailApp.search(searchText, 0, 5);
  threads.forEach(function (thread) {
    var message = thread.getMessages()[0];
    Logger.log(message);
  });
}
```

![GmailMessage log](img/ja/gmail-messages.png)
**GmailMessage** というテキストが表示されます。

最後に、取得系のメソッドを叩いて値を取得します。

```JavaScript
function main() {
  var searchText = '';
  var threads = GmailApp.search(searchText, 0, 5);
  threads.forEach(function (thread) {
    var message = thread.getMessages()[0];
    Logger.log(message.getSubject());
    Logger.log(message.getFrom());
    Logger.log(message.getTo());
    Logger.log(message.getPlainBody());
    Logger.log(message.getDate());
  });
}
```

![Get Gmail message values](img/ja/get-message-values.png)
スクリプトを実行してログを見てみます。Gmail メッセージの値を取得できました。次のセッションからは、これらの値を Spreadsheet に保存していきます。

## SpreadsheetApp クラス
Duration: 0:03:00

次に、 SpreadsheetApp クラスについて理解します。[SpreadsheetApp reference](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app) にアクセスし、[getActiveSheet](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app#getactivesheet) のセクションを確認します。このメソッドを使うことで作成したスプレッドシートにアクセスすることができます。

![Spreadsheet classes](img/ja/spreadsheet-classes.png)
SpreadsheetApp のクラス群は Spreadsheet > Sheet > Range という順に階層構造になっています。`SpreadsheetApp.getActiveSheet()` の返り値は Sheet クラスなので、スプレッドシートに値を挿入する場合には Range クラスまで掘り下げてアクセスする必要があります。

下記の関数を実行します。Gmail の時と同様に認証の許可が必要です。

```JavaScript
function saveMessages() {
  Logger.log(SpreadsheetApp.getActiveSheet().getName());
}
```

![Change function](img/ja/change-function.png)
実行する関数を変更する際には上の画像のように選択します。

![Logging Spreadsheet tab name](img/ja/log-getactivesheet.png)
![Spreadsheet tab name](img/ja/spreadsheet-tabname.png)
ログビュアーでスプレッドシートのタブ名が表示されます。スプレッドシートにデータを保存するためには [Range](https://developers.google.com/apps-script/reference/spreadsheet/range) クラスにアクセスし[Range.setValues()](https://developers.google.com/apps-script/reference/spreadsheet/range#setvaluesvalues) をたたく必要があります。

以下を実装して `saveMessages` 関数を実行します。

```JavaScript
function saveMessages() {
  var data = [
    ['a', 'b', 'c'],
    ['d', 'e', 'f'],
  ];
  SpreadsheetApp
    .getActiveSheet()
    .getRange("A1:C2")
    .setValues(data);
}
```

![Range.setValues](img/ja/range-setvalues.png)
スプレッドシートを見てみると値が挿入されていることが確認できます。ここで引数に指定したデータが 2 次元配列であることに注意してください。

さらに、データを引数で渡せるように関数を編集します。今回は 5 つのタイプの値を保存するため、列 "E" は固定とします。

```JavaScript
function saveMessages(data) {
  SpreadsheetApp
    .getActiveSheet()
    .getRange("A1:E" + data.length)
    .setValues(data);
}
```

外から叩くためのテストとして、 `test` 関数を実装し、実行してみます。

```JavaScript
function test() {
  var data = [
    ['a', 'b', 'c', 'd', 'e'],
    ['f', 'g', 'h', 'i', 'j'],
    ['A', 'B', 'C', 'D', 'E'],
    ['F', 'G', 'H', 'I', 'J'],
  ];
  saveMessages(data);
}
```

![set data as argument](img/ja/data-argument.png)
`saveMessages` 関数を叩けてスプレッドシートに値が挿入されました。テストできたので test 関数を削除します。次のセクションでは `main` 関数から `saveMessages` 関数を叩いて Gmail のメッセージ一覧を保存します。

## メッセージを Spreadsheet に保存
Duration: 0:03:00

`saveMessages` 関数を叩けるように `main` 関数を更新します。新たに `messages` 変数を定義し 2 次元配列になるようにメッセージデータを格納します。すべてのスレッドのメッセージを抽出したら `saveMessages` 関数をコールして、スプレッドシートに保存します。

```JavaScript
function main() {
  var searchText = '';
  var threads = GmailApp.search(searchText, 0, 5);
  var messages = [];
  threads.forEach(function (thread) {
    var message = thread.getMessages()[0];
    // single cell characters limit
    if (message.getPlainBody().length > 50000) {
      return;
    }
    messages.push([
      message.getSubject(),
      message.getFrom(),
      message.getTo(),
      message.getPlainBody(),
      message.getDate(),
    ]);
  });
  saveMessages(messages);
}

function saveMessages(data) {
  SpreadsheetApp
    .getActiveSheet()
    .getRange("A1:E" + data.length)
    .setValues(data);
}
```

`main` 関数を実行し、スプレッドシートを確認します。

![save gmail messages to Spreadsheet](img/ja/save-gmail-messages.png)

スプレッドシートで Gmail メッセージのデータを確認できました。それぞれの列が何を示すのか理解しやすくするために、最初の列に列名を指定します。コードの 4 行目で messages を定義している行を以下に書き換えて実行します。

```JavaScript
  var messages = [['Subject', 'From', 'To', 'Body', 'Date']];
```

![add column names](img/ja/add-column-names.png)

列の意味がわかりやすくなりました。30 行に満たないコードで Gmail と Spreadsheet を連携するアプリケーションを作成することができました！

## Spreadsheet のカスタムメニュー
Duration: 0:03:00

これまで作ったアプリケーションで十分要件を満たしますが、これをより使いやすく改善することができます。 スプレッドシートにカスタムメニューを追加し、それを選択することで Gmail からメッセージ一覧を取得できるようにします。下記の `onOpen` 関数を追加し実行します。 `onOpen` という関数名は予約されており、スプレッドシートが開かれるタイミングで呼ばれます。詳しくは [公式ドキュメント](https://developers.google.com/apps-script/guides/triggers#onopene) で確認できます。

```JavaScript
function onOpen() {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('Gmail', [
      {name: 'Fetch', functionName: 'main'},
    ]);
}
```

[SpreadsheetApp.getActiveSpreadsheet()](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app#getactivespreadsheet) と [Spreadsheet.addMenu](https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#addmenuname,-submenus) を用いて実装します。

![Spreadsheet custom menu](img/ja/spreadsheet-custom-menu.png)

スプレッドシートを見てみましょう。カスタムメニューが表示されます。 **Gmail** > **Fetch** を選択するとメッセージが反映されます。これでスプレッドシートの UI から Apps Script の関数を実行することができるようになりました。

スプレッドシートの値をリセットできるようにします。 `clearSheet` 関数を追加してメニューに追加します。

```JavaScript
function onOpen() {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('Gmail', [
      {name: 'Fetch', functionName: 'main'},
      {name: 'Clear sheet', functionName: 'clearSheet'},
    ]);
}

function clearSheet() {
  SpreadsheetApp
    .getActiveSheet()
    .clear();
}
```

![clear menu](img/ja/clear-menu.png)
**Clear sheet** サブメニューが追加されるので実行してみましょう。スプレッドシートに挿入されたデータがなくなれば OK です。

## 自動化
Duration: 0:05:00

定期的に実行したい場合など、[トリガー](https://developers.google.com/apps-script/guides/triggers/installable)を使って自動化の設定をすることも可能です。 [Time-driven trigger](https://developers.google.com/apps-script/guides/triggers/installable#time-driven_triggers) を設定して `main` 関数を 1 分おきに実行してみましょう。

以下に沿って進めます。

![Current project's triggers](img/ja/current-project-triggers.png)
**編集** > **現在のプロジェクトのトリガー** を選択します

![Add triggers](img/ja/add-trigger.png)
**トリガーを追加** を選択します

![Trigger settings](img/ja/trigger-settings.png)
ポップアップが表示されるのでトリガーの設定を行います。上の画像の通り設定して **保存** ボタンを選択します。

![Trigger is created](img/ja/trigger-is-created.png)
トリガーが作成されました！ドットアイコンを選択します。

![Select executions menu](img/ja/select-executions.png)
**実行数** を選択します。

![Executions](img/ja/executions.png)

`main` 関数が 1 分おきに実行されているのが確認できます。新しいメッセージを受信した際でも最新の状態を保つことができます。

![Delete trigger](img/ja/delete-trigger.png)
**ドットアイコン** > **トリガーの削除** を選択することでトリガーを削除できます。

![Triggers](img/en/triggers.png)
Google Apps Script には[様々な種類のトリガー](https://developers.google.com/apps-script/guides/triggers)が用意されています。トリガーを使うことでプロジェクトをより便利にすることができます。

## Congrats!
Duration: 0:01:00

おめでとうございます！この Codelab は以上で終了です。最終的なコードは以下で確認できます。また [tanabee/gas-codelab](https://github.com/tanabee/gas-codelab) の GitHub にも上がっています。

```JavaScript
function onOpen() {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('Gmail', [
      {name: 'Fetch', functionName: 'main'},
      {name: 'Clear sheet', functionName: 'clearSheet'},
    ]);
}

function main() {
  var searchText = '';
  var threads = GmailApp.search(searchText, 0, 500);
  var messages = [['Subject', 'From', 'To', 'Body', 'Date']];
  threads.forEach(function (thread) {
    var message = thread.getMessages()[0];
    // single cell characters limit
    if (message.getPlainBody().length > 50000) {
      return;
    }
    messages.push([
      message.getSubject(),
      message.getFrom(),
      message.getTo(),
      message.getPlainBody(),
      message.getDate(),
    ]);
  });
  saveMessages(messages);
}

function saveMessages(data) {
  clearSheet();
  SpreadsheetApp
    .getActiveSheet()
    .getRange("A1:E" + data.length)
    .setValues(data);
}

function clearSheet() {
  SpreadsheetApp
    .getActiveSheet()
    .clear();
}
```

### Next Action

このプロジェクトをより便利にするアイデアを載せておきます。

- 500 件以上の Gmail メッセージの取得 (for 文を使って)
- スプレッドシートへの保存を上書きでなく追加する形式にする (重複の考慮)
- フォームの送信結果など固定フォーマットの本文をパースして保存する
- [Data Studio](https://datastudio.google.com) を使った見える化
