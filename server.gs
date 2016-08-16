var mapping = {
  '項目名': 'name',
  'タイプ': 'type',
  '必須': 'required',
  '選択肢': 'choices',
  '表示条件': 'showCondition',
  'プレースホルダ': 'placeholder',
  '下部に表示するテキスト': 'bottomHTML',
  'ポップアップで表示するテキスト': 'tooltip'
};
var global = {};

/** GET受付. GETのパラメータでssIdを受けつける。SpreadSheetのIDを指定する */
function doGet(e) {
  var ssId = e.parameter.ssId;
  var uuid = e.parameter.uuid || '';
  var t = HtmlService.createTemplateFromFile('Index');
  t.ssId = ssId;
  t.uuid = uuid;
  var bi = getBasicInformation(ssId);
  return t.evaluate()
      .setTitle(bi['タイトル'])
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
  var extInfo = getExtInfo(ssId);

  var rows = range.getValues();
  rows.shift();
  rows.forEach(function(row) {
    var value = {};
    values.push(value);
    row.forEach(function(cell, i) {
      value[mapping[header[i]]] = row[i];
    });
    value.required = value.required == null || value.required != '';
    value.choices = value.choices ? replaceTemplate(value.choices, extInfo).split(',') : null;
  });
  return values;
}

/** Spreadsheetのシートのデータを配列にして返す */
function getSheetDataAsArray(ssId, sheetName) {
  var ss = SpreadsheetApp.openById(ssId);
  var sheet = ss.getSheetByName(sheetName);
  var range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());
  var header = range.getValues()[0].map(function(cellValue) { return cellValue; });
  var values = [];
  var rows = range.getValues();
  rows.shift();
  rows.forEach(function(row) {
    var value = {};
    row.forEach(function(cell, i) {
      value[header[i]] = row[i];
    });
    values.push(value);
  });
  return values;
}

/** 入力データを返す */
function getFormData(ssId, uuid) {
  var values = getSheetDataAsArray(ssId, 'データ');
  for (var i = 0, len = values.length; i < len; i++) {
    var element = values[i];
    if (element['uuid'] === uuid) {
      return dataConvertor(element);
    }
  }
  return null;
}

/** ajaxで返せる値に変換する */
function dataConvertor(data) {
  var newData = {};
  for (var prop in data) {
    if (data[prop].getTime) {
      newData[prop] = data[prop].getTime();
    } else {
      newData[prop] = data[prop];
    }
  }
  return newData;
}

/** HTMLから呼ばれるメソッド. フォームのsubmit処理 */
function submitForm(form) {
  saveData(form.ssId, form);
  mailData(form.ssId, form);
}
/** 文字列置換をおこなう */
function replaceTemplate(template, extInfo) {
  if (template == undefined) {
    return undefined;
  }
  for (var prop in extInfo) {
    var target = '${' + prop + '}';
    while (template.indexOf(target, 0) !== -1 ) {
      template = template.replace(target, extInfo[prop]);
    }
  }
  return template;
}

/**
 * HTMLから呼ばれるメソッド. 
 * uuidがなければSpreadsheetのデータシートにデータを追加
 * uuidがあればSpreadsheetのデータシートにデータを更新
 */
function saveData(ssId, form) {
  var data = form;
  var ss = SpreadsheetApp.openById(ssId);
  var sheetData = ss.getSheetByName('データ');
  var headerDefs = getFormDefinition(ssId);
  var dataHeader = sheetData.getRange(1, 1, 1, sheetData.getLastColumn()).getValues()[0];
  if (!data['作成日時']) {
    data['作成日時'] = new Date();
  }
  if (!data.uuid) {
    data.uuid = Utilities.getUuid();
    var rowData = dataHeader.map(function(header) { return data[header] || null; });
    sheetData.appendRow(rowData);
  } else {
    // TODO update row
  }
}

/** formの入力値を取得する. browserでの実行と互換性を持たせるためにメソッド化している */
function getValue(name) {
  return global.form[name];
}

/** HTMLから呼ばれるメソッド. フォームのデータをメールする */
function mailData(ssId, form) {
  global.form = form;
  var mailDefs = getSheetDataAsArray(ssId, 'メール定義');
  var enabledMail = mailDefs.filter(function(def) {
    var condition = def['送信条件'];
    return eval(condition);
  });
  var extInfo = getExtInfo(ssId);
  enabledMail.forEach(function(def) {
    var title = def['タイトル'];
    var to = def['宛先'];
    var body = def['本文'];
    var attachmentVariables = def['添付ファイル'].split(/,/);
    // 本文のテンプレート文字を置き換える
    body = replaceTemplate(body, form);
    body = replaceTemplate(body, extInfo);
    var headerDefs = getFormDefinition(ssId);
    var attachments = headerDefs.filter(function(def) {
      return def.type == 'file' && attachmentVariables.indexOf(def.name) != -1;
    }).map(function(def) {
      var d = form[def.name];
      Logger.log(form);
      Logger.log(def);
      Logger.log(d);
      var files = d.split(";").map(function(str) {
        var token = str.split(':');
        return {fname: token[0], mimetype: token[1], data: Utilities.base64Decode(token[2])};
      });
      return files.map(function(file) {
        return Utilities.newBlob(file.data, file.mimetype, file.fname);
      });
    });
    MailApp.sendEmail({
      to: to,
      subject: title,
      htmlBody: body,
      attachments: Array.prototype.concat.apply([], attachments) // flatten
    });
  });
}

/** 拡張情報から値を取得する */
function getExtInfo(ssId) {
  var rows = getSheetDataAsArray(ssId, '拡張情報');
  var values = {};
  rows.forEach(function(row) {
    if (row['関数'] !== '') {
      values[row['項目']] = eval(row['値'])(ssId);
    } else {
      var value = row['値'];
      if (value.getTime) {
        values[row['項目']] = value.getTime();
      } else {
        values[row['項目']] = row['値'];
      }
    }
  });
  Logger.log(values);
  return values;
}

/** テスト用のメソッド.Apps Scriptのエディタから簡単に実行できるように作成した */
function test() {
  var ssId = '1wmKMViejLkpqOFOY3seY0UPfbYbHFmX70s2MNd40jg0';
  getFormBasicInformation(ssId);
  getDefinition(ssId);
}

function test2() {
  var ssId = '1wmKMViejLkpqOFOY3seY0UPfbYbHFmX70s2MNd40jg0';
  var values = getExtInfo(ssId);
  Logger.log(values);
  //mailData(ssId, {"アンケートタイトル": "あいうえお"});
}
