author: Yuki Tanabe
summary: Apps Script Hands-on
id: en
categories: apps-script,gsuite,javascript
environments: Web
status: Draft
analytics account: UA-141719032-2
feedback link: https://github.com/tanabee/gas-codelab/issues

# Apps Script Hands-on

## Intro

In this codelab, You can make a application that exports Gmail messages to Google Spreadsheet.

## Preparation

You need own Gmail or G Suite account.

Negative
: If you use Gmail for the first time, you need to do below

### 1. Open mail-to-me application

Open the application in your browser.
[mail-to-me](https://script.google.com/d/1p2B0NwXg3APMc-nh-fRWuMX6P_W5eQijsi4CVEHLOOq7_zJmQbWyZAWl/edit?usp=sharing)

### 2. Make a copy

Select **File** > **Make a Copy**.

### 3. Run mail-to-me

Select **Run** button.

### 4. Confirm your mail

Open [Gmail](https://mail.google.com) in your browser and confirm to receiving emails.

## Create Apps Script project

First, create a Apps Script project. Visit [script.google.com](https://script.google.com/home) to open the script editor. Click **+ New Project** to proceed to the script editor. Click **Untitled Project** and input your project name. Then it appears toast and it complete to save when the toast disappears.

## Run script

Add a row below and Click **Run** button. `Logger.log()` is logging function in Google Apps Script.

```JavaScript
function myFunction() {
  Logger.log('Hello Apps Script');
}
```

You can confirm logs when you select **View** > **Logs** in the menu.

You can alse use `console.log()`, but we use `Logger.log()` in this Hands-on because you need additional process to use `console.log()`. 

## GmailApp class

Next, understand GmailApp class. Confirm Google apps classes in [official reference](https://developers.google.com/apps-script/reference/). See [GmailApp document](https://developers.google.com/apps-script/reference/gmail/gmail-app). You can confirm GmailApp classes (ex. [GmailMessage](https://developers.google.com/apps-script/reference/gmail/gmail-message), [GmailThread](https://developers.google.com/apps-script/reference/gmail/gmail-thread) ) and methods (ex. [search](https://developers.google.com/apps-script/reference/gmail/gmail-app#searchquery,-start,-max), [sendEmail](https://developers.google.com/apps-script/reference/gmail/gmail-app#sendemailrecipient,-subject,-body,-options)) 

We use [getInboxThreads()](https://developers.google.com/apps-script/reference/gmail/gmail-app#getinboxthreadsstart,-max) method for Retrieving emails in this time.

## Fetch emails

```JavaScript
function main() {
  var threads = GmailApp.getInboxThreads(0, 100);
  var messages = [];
  threads.forEach(function (thread) {
    var message = thread.getMessages()[0];
    messages.push([
      message.getSubject(),
      message.getFrom(),
      message.getTo(),
      message.getPlainBody(),
      message.getDate(),
    ]);
  });
  Logger.log(messages);
}
```

## SpreadsheetApp class

```JavaScript
function insertSpreadSheet() {
  var data = [
    ['a', 'b', 'c'],
    ['d', 'e', 'f'],
  ];
  SpreadsheetApp
    .getActiveSheet()
    .getRange("A1:B2")
    .setValues(data);
}
```

## Save emails to Spreadsheet

```JavaScript
function main() {
  var threads = GmailApp.getInboxThreads(0, 200);
  var messages = [];
  threads.forEach(function (thread) {
    var message = thread.getMessages()[0];
    if (message.getPlainBody().length > 10000) {
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
  insertSpreadSheet(messages);
}

function insertSpreadSheet(data) {
  SpreadsheetApp
    .getActiveSheet()
    .getRange("A1:E" + data.length)
    .setValues(data);
}
```

## Spreadsheet Custom Menu

```JavaScript
function onOpen() {
  SpreadsheetApp
    .getActiveSpreadsheet()
    .addMenu('Gmail', [
      {name: 'Fetch', functionName: 'main'},
    ]);
}
```

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

## Congrats

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
  var threads = GmailApp.getInboxThreads(0, 200);
  var messages = [['Subject', 'From', 'To', 'Body', 'Date']];
  threads.forEach(function (thread) {
    var message = thread.getMessages()[0];
    if (message.getPlainBody().length > 10000) {
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
  insertSpreadSheet(messages);
}

function insertSpreadSheet(data) {
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
