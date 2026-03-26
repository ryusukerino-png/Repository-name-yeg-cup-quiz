const SHEET_NAME = 'entries';

function doGet(e) {
  const action = e.parameter.action;
  
  if (action === 'getAll') {
    return getAll();
  }
  
  if (action === 'register') {
    const payload = {
      action: 'register',
      name: e.parameter.name,
      bets: JSON.parse(decodeURIComponent(e.parameter.bets || '[]')),
      time: e.parameter.time || '',
      code: e.parameter.code || ''
    };
    return register(payload);
  }
  
  if (action === 'delete') {
    return deleteEntry({ name: e.parameter.name });
  }
  
  return ContentService
    .createTextOutput(JSON.stringify({status: 'error', message: 'unknown action'}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    let payload;
    
    try {
      payload = JSON.parse(e.postData.contents);
    } catch(err) {
      payload = e.parameter;
      if (payload.bets) {
        payload.bets = JSON.parse(decodeURIComponent(payload.bets));
      }
    }
    
    const action = payload.action;
    
    if (action === 'register') {
      return register(payload);
    } else if (action === 'delete') {
      return deleteEntry(payload);
    } else if (action === 'reset') {
      return resetAll();
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: 'unknown action'}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // ヘッダー: name, bets, time, code
    sheet.appendRow(['name', 'bets', 'time', 'code']);
  }
  return sheet;
}

function getAll() {
  try {
    const sheet = getSheet();
    const data = sheet.getDataRange().getValues();
    const entries = [];
    
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      entries.push({
        name: data[i][0],
        bets: JSON.parse(data[i][1] || '[]'),
        time: data[i][2] || '',
        code: data[i][3] || ''   // ← コードを返す
      });
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({status: 'ok', data: entries}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function register(payload) {
  try {
    const sheet = getSheet();
    const name = payload.name;
    const bets = payload.bets;
    const time = payload.time || '';
    const code = payload.code || String(Math.floor(1000 + Math.random() * 9000));
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === name) {
        // 上書き時はコードも更新
        sheet.getRange(i + 1, 1, 1, 4).setValues([[name, JSON.stringify(bets), time, code]]);
        return ContentService
          .createTextOutput(JSON.stringify({status: 'ok', message: '上書き完了', code: code}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    sheet.appendRow([name, JSON.stringify(bets), time, code]);
    
    return ContentService
      .createTextOutput(JSON.stringify({status: 'ok', message: '登録完了', code: code}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function deleteEntry(payload) {
  try {
    const sheet = getSheet();
    const name = payload.name;
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === name) {
        sheet.deleteRow(i + 1);
        return ContentService
          .createTextOutput(JSON.stringify({status: 'ok', message: '削除完了'}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: '該当データなし'}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function resetAll() {
  try {
    const sheet = getSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    return ContentService
      .createTextOutput(JSON.stringify({status: 'ok', message: 'リセット完了'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({status: 'error', message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
