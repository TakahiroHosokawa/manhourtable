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

// 共通：シートデータをオブジェクト配列として取得
function getTableData_(sheetName) {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data.shift();
  return data.map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// クライアントへ初期データを送る
function getAppData() {
  try {
    const data = {
      product: getTableData_(SHEET_NAMES.PRODUCT)[0], // 今回は単一製品前提
      params: getTableData_(SHEET_NAMES.PARAM),
      processes: getTableData_(SHEET_NAMES.PROCESS),
      groups: getTableData_(SHEET_NAMES.GROUP),
      works: getTableData_(SHEET_NAMES.WORK),
      details: getTableData_(SHEET_NAMES.WORK_DETAIL),
      choices: getTableData_(SHEET_NAMES.CHOICE)
    };
    return { ok: true, data: data };
  } catch (e) {
    return { ok: false, message: e.toString() };
  }
}

// 作業内訳の1行保存（排他制御＋競合検知付き）
function saveWorkDetail(newRow, baseRow) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // 20秒待ち
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
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
      return { ok: true, message: '追加しました', row: newRow };
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
      return { ok: false, conflict: true, message: '他のユーザによる変更と競合しました。', details: conflictDetails };
    }

    // 更新日時等の反映
    currentRow[headers.indexOf('更新日時')] = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    currentRow[headers.indexOf('更新者')] = Session.getActiveUser().getEmail();
    
    sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([currentRow]);
    
    // (※本来はここで40_履歴.gsを呼んで履歴を記録しますが、段階的に後ほど実装します)

    return { ok: true, message: '保存しました' };

  } catch (e) {
    return { ok: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}