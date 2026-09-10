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
    headers.forEach((h, i) => {
      // Date型が含まれると google.script.run の戻り値が null になるため文字列化する
      obj[h] = (row[i] instanceof Date) ? row[i].toISOString() : row[i];
    });
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
      equipment: getTableData_(SHEET_NAMES.EQUIPMENT), // v0.7追加
      works: getTableData_(SHEET_NAMES.WORK),
      details: getTableData_(SHEET_NAMES.WORK_DETAIL),
      equipment_usage: getTableData_(SHEET_NAMES.EQUIPMENT_USAGE), // v0.7追加
      choices: getTableData_(SHEET_NAMES.CHOICE),
      records: getTableData_(SHEET_NAMES.RECORD) // 観測記録
    };
    return JSON.stringify({ ok: true, data: data });
  } catch (e) {
    return JSON.stringify({ ok: false, message: e.toString() });
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

    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const user = Session.getActiveUser().getEmail();

    if (rowIndex === -1) {
      // 新規追加（監査項目を必ず埋める）
      const appendData = headers.map(h => {
        if (h === '作成日時' || h === '更新日時') return now;
        if (h === '作成者' || h === '更新者') return user;
        return newRow[h] !== undefined ? newRow[h] : '';
      });
      sheet.appendRow(appendData);
      appendHistory_(newRow['内訳ID'], '追加', [], now, user);
      return { ok: true, message: '追加しました', row: newRow };
    }

    // 既存行の更新（競合チェック）
    const currentRow = data[rowIndex];
    let conflict = false;
    let conflictDetails = {};
    const changes = []; // 5.17 変更履歴に残す項目単位の差分

    headers.forEach((h, colIndex) => {
      // 監査項目は無視
      if (['作成日時', '作成者', '更新日時', '更新者'].includes(h)) return;

      const currentVal = normalizeValue_(currentRow[colIndex]);
      const baseVal = normalizeValue_(baseRow[h]);
      const newVal = normalizeValue_(newRow[h]);

      if (currentVal !== baseVal && currentVal !== newVal) {
        // 基準値からシートの値が変わっており、かつ自分が送信した値とも違う場合＝競合
        conflict = true;
        conflictDetails[h] = { current: currentVal, yours: newVal };
      } else {
        // 競合なし -> 新しい値を適用
        if (newRow[h] !== undefined) {
          if (currentVal !== newVal) changes.push({ field: h, oldValue: currentVal, newValue: newVal });
          currentRow[colIndex] = newRow[h];
        }
      }
    });

    if (conflict) {
      return { ok: false, conflict: true, message: '他のユーザによる変更と競合しました。', details: conflictDetails };
    }

    if (changes.length === 0) {
      return { ok: true, message: '変更はありません' };
    }

    // 更新日時等の反映
    currentRow[headers.indexOf('更新日時')] = now;
    currentRow[headers.indexOf('更新者')] = user;

    sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([currentRow]);

    // F10: 項目単位の変更履歴。無効化・復活は操作名を分けて記録する（5.17）
    const stateChange = changes.filter(c => c.field === '状態')[0];
    let operation = '変更';
    if (stateChange) {
      if (stateChange.newValue === '無効') operation = '無効化';
      else if (stateChange.oldValue === '無効') operation = '復活';
    }
    appendHistory_(newRow['内訳ID'], operation, changes, now, user);

    return { ok: true, message: '保存しました' };

  } catch (e) {
    return { ok: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// 値の比較・履歴用の正規化。0 と空欄、数値と文字列の「6」と「6.0」を取り違えないようにする
function normalizeValue_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'number') return String(v);
  if (typeof v === 'boolean') return v ? 'true' : 'false';
  const s = String(v).trim();
  if (s !== '' && !isNaN(Number(s))) return String(Number(s));
  return s;
}

/**
 * F10・5.17: 変更履歴を項目単位で追記する
 * @param {string} targetId - 対象ID（内訳ID など）
 * @param {string} operation - 追加／変更／無効化／復活／順序変更／複製／版確定
 * @param {Array} changes - [{ field, oldValue, newValue }]
 */
function appendHistory_(targetId, operation, changes, now, user) {
  try {
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
    const sheet = ss.getSheetByName(SHEET_NAMES.HISTORY);
    if (!sheet) return;

    const productId = productIdOfDetail_(targetId);
    const seq = Math.max(sheet.getLastRow() - 1, 0);
    const rows = (changes && changes.length)
      ? changes.map((c, i) => ['H-' + (seq + i + 1), now, user, productId, '作業内訳', targetId, operation, c.field, c.oldValue, c.newValue])
      : [['H-' + (seq + 1), now, user, productId, '作業内訳', targetId, operation, '', '', '']];

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.HISTORY.length).setValues(rows);
  } catch (e) {
    // 履歴の失敗で本体の保存を巻き戻さない。ログにだけ残す
    console.error('変更履歴の記録に失敗: ' + e);
  }
}

// 内訳ID（製品コード-E-連番）から製品IDを引く
function productIdOfDetail_(detailId) {
  const code = String(detailId || '').split(CONFIG.ID_PREFIX_DETAIL)[0];
  if (!code) return '';
  const products = getTableData_(SHEET_NAMES.PRODUCT);
  for (let i = 0; i < products.length; i++) {
    if (products[i]['製品コード'] === code) return products[i]['製品ID'];
  }
  return '';
}
