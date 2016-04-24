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
  Logger.log(range.getValues());
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

/** テスト用のメソッド.Apps Scriptのエディタから簡単に実行できるように作成した */
function test() {
  var ssId = '1PyVzzwhHqp4XDWLq5fXce61DiUDT78Go_f3W_q18Q84';
  getFormBasicInformation(ssId);
  getDefinition(ssId);
}
