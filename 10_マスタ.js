// =========================================
// 10_マスタ.gs
// =========================================

function doGet(e) {
  // Webアプリとしての画面描画
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('標準工数マスタ管理')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 共通：スプレッドシートを開く（IDの未設定・誤りをその場で切り分ける）
function openSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) {
    throw new Error('スクリプトプロパティ「SPREADSHEET_ID」が未設定です。');
  }
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error('SPREADSHEET_ID「' + id + '」でスプレッドシートを開けません: ' + e.message);
  }
}

// 共通：シートデータをオブジェクト配列として取得
// ※ Dateなど直列化できない値は文字列へ寄せる
function getTableData_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data.shift();
  return data.map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      const v = row[i];
      obj[h] = (v instanceof Date)
        ? Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss')
        : v;
    });
    return obj;
  });
}

// クライアントへ初期データを送る
// google.script.run のオブジェクト直列化で null になる事故を避けるため JSON文字列で返す
function getAppData() {
  try {
    const ss = openSpreadsheet_();
    const product = getTableData_(ss, SHEET_NAMES.PRODUCT)[0]; // 今回は単一製品前提
    if (!product) {
      throw new Error('「' + SHEET_NAMES.PRODUCT + '」シートにデータ行がありません。initSpreadsheet() を実行してください。');
    }
    const data = {
      product: product,
      params: getTableData_(ss, SHEET_NAMES.PARAM),
      processes: getTableData_(ss, SHEET_NAMES.PROCESS),
      groups: getTableData_(ss, SHEET_NAMES.GROUP),
      works: getTableData_(ss, SHEET_NAMES.WORK),
      details: getTableData_(ss, SHEET_NAMES.WORK_DETAIL),
      choices: getTableData_(ss, SHEET_NAMES.CHOICE)
    };
    return JSON.stringify({ ok: true, data: data });
  } catch (e) {
    return JSON.stringify({ ok: false, message: e.message || String(e) });
  }
}

// 切り分け用：各シートの行数だけを返す（GASエディタから直接実行しても可）
function getDiag() {
  const out = { scriptUser: Session.getActiveUser().getEmail() };
  try {
    const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
    out.spreadsheetId = id || '(未設定)';
    const ss = openSpreadsheet_();
    out.spreadsheetName = ss.getName();
    out.sheets = {};
    for (const key in SHEET_NAMES) {
      const sh = ss.getSheetByName(SHEET_NAMES[key]);
      out.sheets[SHEET_NAMES[key]] = sh ? (sh.getLastRow() - 1) + '行' : '(シートなし)';
    }
  } catch (e) {
    out.error = e.message || String(e);
  }
  Logger.log(JSON.stringify(out, null, 2));
  return JSON.stringify(out);
}

// 作業内訳の1行保存（排他制御＋競合検知付き）
function saveWorkDetail(newRow, baseRow) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // 20秒待ち
    const ss = openSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_NAMES.WORK_DETAIL);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf('内訳ID');
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex] === newRow['内訳ID']) {
        rowIndex = i;
        break;
      }
    }

    if (rowIndex === -1) {
      // 新規追加
      const appendData = headers.map(h => newRow[h] !== undefined ? newRow[h] : '');
      sheet.appendRow(appendData);
      return JSON.stringify({ ok: true, message: '追加しました', row: newRow });
    }

    // 既存行の更新（競合チェック）
    const currentRow = data[rowIndex];
    let conflict = false;
    let conflictDetails = {};

    headers.forEach((h, colIndex) => {
      // 監査項目は無視
      if (['作成日時', '作成者', '更新日時', '更新者'].includes(h)) return;
      
      const currentVal = String(currentRow[colIndex] || '');
      const baseVal = String(baseRow[h] || '');
      const newVal = String(newRow[h] || '');

      if (currentVal !== baseVal && currentVal !== newVal) {
        // 基準値からシートの値が変わっており、かつ自分が送信した値とも違う場合＝競合
        conflict = true;
        conflictDetails[h] = { current: currentVal, yours: newVal };
      } else {
        // 競合なし -> 新しい値を適用
        currentRow[colIndex] = newRow[h] !== undefined ? newRow[h] : currentRow[colIndex];
      }
    });

    if (conflict) {
      return JSON.stringify({ ok: false, conflict: true, message: '他のユーザによる変更と競合しました。', details: conflictDetails });
    }

    // 更新日時等の反映
    currentRow[headers.indexOf('更新日時')] = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    currentRow[headers.indexOf('更新者')] = Session.getActiveUser().getEmail();
    
    sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([currentRow]);
    
    // (※本来はここで40_履歴.gsを呼んで履歴を記録しますが、段階的に後ほど実装します)

    return JSON.stringify({ ok: true, message: '保存しました' });

  } catch (e) {
    return JSON.stringify({ ok: false, message: e.message || String(e) });
  } finally {
    lock.releaseLock();
  }
}

// 作業内訳の新規追加・複製
function addWorkDetail(workId, copySourceDetail) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const ss = openSpreadsheet_();
    const sheet = ss.getSheetByName(SHEET_NAMES.WORK_DETAIL);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf('内訳ID');
    const orderIndex = headers.indexOf('順序');
    const workIdIndex = headers.indexOf('作業ID');
    
    let maxNum = 0;
    let maxOrder = 0;
    const productCode = workId.split(CONFIG.ID_PREFIX_WORK)[0];

    for (let i = 1; i < data.length; i++) {
      const currentId = data[i][idIndex];
      if (currentId) {
        const numPart = String(currentId).split(CONFIG.ID_PREFIX_DETAIL)[1];
        if (numPart) {
          const num = parseInt(numPart, 10);
          if (num > maxNum) maxNum = num;
        }
      }
      if (data[i][workIdIndex] === workId) {
        const order = parseInt(data[i][orderIndex], 10);
        if (!isNaN(order) && order > maxOrder) maxOrder = order;
      }
    }

    const newIdNum = maxNum + 1;
    const newId = productCode + CONFIG.ID_PREFIX_DETAIL + newIdNum;
    const newOrder = maxOrder + CONFIG.ORDER_INTERVAL;
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const user = Session.getActiveUser().getEmail();

    let newRowData = {};
    if (copySourceDetail) {
      newRowData = Object.assign({}, copySourceDetail);
      newRowData['内訳ID'] = newId;
      newRowData['順序'] = newOrder;
      newRowData['作成日時'] = now;
      newRowData['作成者'] = user;
      newRowData['更新日時'] = now;
      newRowData['更新者'] = user;
    } else {
      headers.forEach(h => newRowData[h] = '');
      newRowData['内訳ID'] = newId;
      newRowData['作業ID'] = workId;
      newRowData['順序'] = newOrder;
      newRowData['区分'] = '実作業';
      newRowData['単位'] = '台';
      newRowData['時間方式'] = '分担作業';
      newRowData['数量方式'] = '固定';
      newRowData['固定数量'] = 1;
      newRowData['状態'] = '有効';
      newRowData['作成日時'] = now;
      newRowData['作成者'] = user;
      newRowData['更新日時'] = now;
      newRowData['更新者'] = user;
    }

    const appendData = headers.map(h => newRowData[h] !== undefined ? newRowData[h] : '');
    sheet.appendRow(appendData);

    return JSON.stringify({ ok: true, message: '追加しました', row: newRowData });
  } catch (e) {
    return JSON.stringify({ ok: false, message: e.message || String(e) });
  } finally {
    lock.releaseLock();
  }
}
