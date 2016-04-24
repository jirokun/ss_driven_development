/** GET受付. GETのパラメータでssIdを受けつける。SpreadSheetのIDを指定する */
function doGet(e) {
  var ssId = e.parameter.ssId;
  var t = HtmlService.createTemplateFromFile('Index');
  t.ssId = ssId;
  return t.evaluate()
      .setTitle('Form Sample')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

/** HTMLから呼ばれるメソッド. Spreadsheetの基本情報シートのデータをオブジェクトにして返す */
function getBasicInformation(ssId) {
  var ss = SpreadsheetApp.openById(ssId);
  var fi = ss.getSheetByName('基本情報');
  var range = fi.getRange(1, 1, fi.getLastRow(), 2);
  var values = {};
  range.getValues().forEach(function(row) {
    values[row[0]] = row[1];
  });
  return values;
}

/** HTMLから呼ばれるメソッド. Spreadsheetのフォーム定義のデータを配列にして返す */
function getFormDefinition(ssId) {
  var ss = SpreadsheetApp.openById(ssId);
  var fi = ss.getSheetByName('フォーム定義');
  var range = fi.getRange(1, 1, fi.getLastRow(), fi.getLastColumn());
  var header = range.getValues()[0].map(function(cellValue) { return cellValue; });
  var values = [];
  var rows = range.getValues();
  rows.shift();
  rows.forEach(function(row) {
    var value = {};
    values.push(value);
    row.forEach(function(cell, i) {
      value[header[i]] = row[i];
    });
  });
  return values;
}

/** HTMLから呼ばれるメソッド. フォームのsubmit処理 */
function submitForm(form) {
  Logger.log(form);
  addData(form.ssId, form);
  mailData(form.ssId, form);
}

/** HTMLから呼ばれるメソッド. Spreadsheetのデータシートにデータを追加する */
function addData(ssId, form) {
  var data = normalizeData(form);
  var ss = SpreadsheetApp.openById(ssId);
  var sheetData = ss.getSheetByName('データ');
  var dataHeader = sheetData.getRange(1, 1, 1, sheetData.getLastColumn()).getValues()[0];
  if (!data['作成日時']) {
    data['作成日時'] = new Date();
  }
  var rowData = dataHeader.map(function(header) { return data[header] || null; });
  sheetData.appendRow(rowData);
}

/** HTMLから呼ばれるメソッド. フォームのデータをメールする */
function mailData(ssId, form) {
  var data = normalizeData(form);
  var def = getBasicInformation(ssId);
  var mailTo = def['メール送付先'];
  // 設定されていなかったら何もしない
  if (!mailTo || mailTo === '') return;
  var mailTitle = def['メールタイトル'];
  var headers = getFormDefinition(ssId).map(function(row) { return row['項目名']; });
  var body = "<table>" + headers.map(function(column) {
    return '<tr><th style="text-align: right;">' + column + '</th><td>' + data[column] + '</td></tr>';
  }).join('\n') + "</table>";
  var attachments = [];
  for (var prop in data) {
    var d = data[prop];
    if (d.copyBlob) {
      // blobと判定する
      attachments.push(d);
    }
  }
  MailApp.sendEmail({
    to: mailTo,
    subject: mailTitle,
    htmlBody: body,
    attachments: attachments
  });
}

/** formの情報からセクション名を取り除いたキーのデータを作成する */
function normalizeData(form) {
  var data = {};
  for (var prop in form) {
    var itemName = prop.split(/_/)[1];
    data[itemName] = form[prop];
  }
  return data;
}

/** テスト用のメソッド.Apps Scriptのエディタから簡単に実行できるように作成した */
function test() {
  var ssId = '1PyVzzwhHqp4XDWLq5fXce61DiUDT78Go_f3W_q18Q84';
  getFormBasicInformation(ssId);
  getDefinition(ssId);
}

function test2() {
  var ssId = '1PyVzzwhHqp4XDWLq5fXce61DiUDT78Go_f3W_q18Q84';
  addData(ssId, {"アンケートタイトル": "あいうえお"});
}
