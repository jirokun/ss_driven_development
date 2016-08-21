/**
 * required librarey
 * moment.js as Momemnt
 */
var mapping = {
  '項目名': 'name',
  'タイプ': 'type',
  '必須': 'required',
  '選択肢': 'choices',
  '表示条件': 'showCondition',
  'スタイル': 'style',
  'プレースホルダ': 'placeholder',
  '下部に表示するテキスト': 'bottomHTML',
  'ポップアップで表示するテキスト': 'tooltip'
};
var global = {};

/** GET受付. GETのパラメータでssIdを受けつける。SpreadSheetのIDを指定する */
function doGet(e) {
  var ssId = e.parameter.ssId;
  var uuid = e.parameter.uuid || '';
  var mode = e.parameter.mode || 'new';
  var t = HtmlService.createTemplateFromFile('Index');
  t.ssId = ssId;
  t.uuid = uuid;
  t.mode = mode;
  t.basicInformation = getBasicInformation(ssId);
  t.formDefinition = getFormDefinition(ssId);
  t.basicInformation = getBasicInformation(ssId);
  t.extInfo = getExtInfo(ssId);
  t.formData = uuid === '' ? null : getFormData(ssId, uuid);
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
    value.required = value.required === null || value.required !== '';
    value.choices = value.choices ? replaceTemplate(value.choices, [extInfo]).split(',') : null;
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
    if (element.uuid === uuid) {
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
  var currentData, diff;
  var isNew = !form.uuid;
  if (!isNew) {
    currentData = getFormData(form.ssId, form.uuid);
    var headerDefs = getFormDefinition(form.ssId);
    diff = diffRow(headerDefs, currentData, form);
  }
  saveData(form.ssId, form, currentData, isNew);
  mailData(form.ssId, form, diff, isNew);
}
/** 文字列置換をおこなう。 */
function replaceTemplate(template, values) {
  if (template === undefined) {
    return undefined;
  }
  values.forEach(function(obj) {
    for (var prop in obj) {
      var target = '${' + prop + '}';
      while (template.indexOf(target, 0) !== -1 ) {
        template = template.replace(target, obj[prop]);
      }
    }
  });
  return template;
}
/** HtmlServiceで文字列置換をおこなう。 */
function evaluateMailTemplate(template, values) {
  if (template === undefined) {
    return undefined;
  }
  var t = HtmlService.createTemplate(template);
  
  values.forEach(function(obj) {
    for (var prop in obj) {
      t[prop] = obj[prop];
    }
  });
  return t.evaluate().getContent();
}
function toDateStr(d) {
  var month = d.getMonth() + 1;
  var date = d.getDate();
  if (month < 10) {
    month = '0' + month.toString();
  }
  if (date < 10) {
    date = '0' + date.toString();
  }
  return [d.getFullYear(), month, date].join('-');
}
function toMonthStr(d) {
  var month = d.getMonth() + 1;
  if (month < 10) {
    month = '0' + month.toString();
  }
  return [d.getFullYear(), month].join('-');
}
function nextSequenceNo(dataArray) {
  return dataArray.length + 1;
}
function countMonthNo(ssId, dataArray) {
  var currentDate = getCurrentDate(ssId);
  var startMonth = new Date(currentDate.getFullYear(), currentDate.getMonth(), 1);
  var counter = 1;  
  dataArray.forEach(function(data) {
    if (data['作成日時'] >= startMonth.getTime()) {
      counter++;
    }
  });
  return Utilities.formatString('%3d', counter);
}
function getCurrentDate(ssId) {
  var basicInformation = getBasicInformation(ssId);
  var currentDate = basicInformation['現在時刻'];
  if (currentDate) {
    return currentDate;
  }
  return new Date();
}

/**
 * HTMLから呼ばれるメソッド. 
 * uuidがなければSpreadsheetのデータシートにデータを追加
 * uuidがあればSpreadsheetのデータシートにデータを更新
 */
function saveData(ssId, form, currentData, isNew) {
  var rowData;
  form['更新日時'] = getCurrentDate(ssId);
  form['最終更新者'] = Session.getActiveUser().getEmail();
  var ss = SpreadsheetApp.openById(ssId);
  var sheetData = ss.getSheetByName('データ');
  var dataArray = getSheetDataAsArray(ssId, 'データ');
  if (isNew) {
    form['作成日時'] = form['更新日時'];
    form.uuid = Utilities.getUuid();
    form.NO = nextSequenceNo(dataArray);
    form.MONTH_NO = countMonthNo(ssId, dataArray);
    rowData = convertData2Row(sheetData, ssId, form);
    sheetData.appendRow(rowData);
  } else {
    form['作成日時'] = new Date(currentData['作成日時']); // 作成日時は変更しない
    var index = findIndex(dataArray, function(row) { return row.uuid === form.uuid; });
    var range = sheetData.getRange(index + 2, 1, 1, sheetData.getLastColumn());
    var row = range.getValues()[0];
    rowData = convertData2Row(sheetData, ssId, form);
    // rangeを取得した時とsetValuesしたときでは行が変わる可能性があるので、データシートではappendRowのみ許し、insertは禁止する
    range.setValues([rowData]);
  }
}

/** formのデータをSpreadSheetのRowに変換する */
function convertData2Row(sheetData, ssId, form) {
  var dataHeader = sheetData.getRange(1, 1, 1, sheetData.getLastColumn()).getValues()[0];
  var headerDefs = getFormDefinition(ssId);

  return dataHeader.map(function(header) {
    var def = getHeaderDef(headerDefs, header);
    if (!!def && def.type === 'file') {
      var files = parseAttachments(form[header]);
      var fileNames = files.map(function(file) { return file.fname; }).join('\n');
      form[header + '_filenames'] = fileNames;
      return fileNames;
    }
    if (Array.isArray(form[header])) {
      return form[header].join(',');
    } else {
      if (form[header] === undefined) {
        return null;
      } else {
        return form[header];
      }
    }
  });
}

/** 名前からカラムの定義を取得する */
function getHeaderDef(headerDefs, name) {
  var index = findIndex(headerDefs, function(def) {
    return def.name === name;
  });
  return headerDefs[index];
}

function findIndex(arr, cb) {
  for (var i = 0, len = arr.length; i < len; i++) {
    if (cb(arr[i])) {
      return i;
    }
  }
  return -1;
}

/** formの入力値を取得する. browserでの実行と互換性を持たせるためにメソッド化している */
function getValue(name) {
  return global.form[name];
}

/** HTMLから呼ばれるメソッド. フォームのデータをメールする */
function mailData(ssId, form, diff, isNew) {
  global.form = form;
  var extInfo = getExtInfo(ssId);
  var translateArr = [form, extInfo, {
    isNew: isNew,
    url: ScriptApp.getService().getUrl(),
    diffHTML: diffHTML(diff)
  }];
  var mailDefs = getSheetDataAsArray(ssId, 'メール定義');
  var enabledMail = mailDefs.filter(function(def) {
    var condition = replaceTemplate(def['送信条件'], translateArr);
    return eval(condition);
  });
  var headerDefs = getFormDefinition(ssId);
  enabledMail.forEach(function(def) {
    var title = replaceTemplate(def['タイトル'], translateArr);
    var to = replaceTemplate(def['宛先'], translateArr);
    var cc = replaceTemplate(def['CC'], translateArr);
    var body = evaluateMailTemplate(replaceTemplate(def['本文'], translateArr), translateArr);
    var attachmentVariables = def['添付ファイル'].split(/,/);
    var attachments = headerDefs.filter(function(def) {
      return def.type == 'file' && attachmentVariables.indexOf(def.name) != -1 && form[def.name];
    }).map(function(def) {
      var files = parseAttachments(form[def.name]);
      return files.map(function(file) {
        return Utilities.newBlob(file.data, file.mimetype, file.fname);
      });
    });
    if (def['タイプ'] === 'HTML') {
      MailApp.sendEmail({
        to: to,
        cc: cc,
        subject: title,
        htmlBody: body,
        attachments: Array.prototype.concat.apply([], attachments) // flatten
      });
    } else {
      MailApp.sendEmail({
        to: to,
        cc: cc,
        subject: title,
        body: body,
        attachments: Array.prototype.concat.apply([], attachments) // flatten
      });
    }      
  });
}

function parseAttachments(value) {
  Logger.log(value);
  if (!value) return [];
  return files = value.split(";").map(function(str) {
    var token = str.split(':');
    return {fname: token[0], mimetype: token[1], data: Utilities.base64Decode(token[2])};
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
  return values;
}

/** 変更点だけを抽出する */
function diffRow(headerDefs, oldRow, newRow) {
  var diff = [], props = [], prop;
  for (prop in oldRow) {
    if (props.indexOf(prop) !== -1) continue;
    props.push(prop);
  }
  for (prop in newRow) {
    if (props.indexOf(prop) !== -1) continue;
    props.push(prop);
  }
  props.forEach(function(prop) {
    if (prop === '更新日時' || prop === '作成日時' || prop === 'mode' || prop === 'ssId') return; // 管理項目はスキップ
    // undefinedを潰す
    var oldValue = oldRow[prop] === undefined ? '': oldRow[prop];
    var newValue = newRow[prop] === undefined ? '': newRow[prop];
    // 数値をintegerに変換できるものは変換する
    oldValue = isInt(oldValue) ? parseInt(oldValue) : oldValue;
    newValue = isInt(newValue) ? parseInt(newValue) : newValue;
    // 日付は文字列に変換する
    var def = getHeaderDef(headerDefs, prop);
    if (def && def.type === 'date' && !!oldValue) {
      oldValue = toDateStr(new Date(oldValue));
    }
    if (oldValue == newValue) {
      return;
    }
    diff.push({key: prop, oldValue: oldValue, newValue: newValue});
  });
  return diff;
}

function isInt(val) {
  return parseInt(val, 10) === val;
}

/** diffの結果をHTMLのTableに変換する */
function diffHTML(diff) {
  if (!diff) return null;
  var html = '<table style="border-collapse: collapse;"><tr><th style="border: 1px solid #ccc;">項目</th><th style="border: 1px solid #ccc;">変更前</th><th style="border: 1px solid #ccc;">変更後</th></tr>';
  html += diff.map(function(d) { return '<tr><td style="border: 1px solid #ccc;">' + d.key + '</td><td style="border: 1px solid #ccc;">' + d.oldValue + '</td><td style="border: 1px solid #ccc;">' + d.newValue + '</td></tr>'; }).join('');
  html += '</table>';
  return html;
}

/** テスト用のメソッド.Apps Scriptのエディタから簡単に実行できるように作成した */
function test() {
  var moment = Moment.moment;
  Logger.log(moment().format());
}

