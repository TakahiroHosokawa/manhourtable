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
function getTableData_(sheetName, ss) {
  if (!ss) {
    ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
  }
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
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
    const currentProductId = PropertiesService.getUserProperties().getProperty('CURRENT_PRODUCT_ID') || 'P-001';
    
    const allProducts = getTableData_(SHEET_NAMES.PRODUCT, ss);
    let currentProduct = allProducts.find(p => p['製品ID'] === currentProductId);
    
    if (!currentProduct && allProducts.length > 0) {
      currentProduct = allProducts[0];
      PropertiesService.getUserProperties().setProperty('CURRENT_PRODUCT_ID', currentProduct['製品ID']);
    }

    // 製品に紐づくデータだけフィルタ
    const works = getTableData_(SHEET_NAMES.WORK, ss).filter(w => w['製品ID'] === currentProduct['製品ID']);
    const workIds = works.map(w => w['作業ID']);
    const details = getTableData_(SHEET_NAMES.WORK_DETAIL, ss).filter(d => workIds.includes(d['作業ID']));
    const detailIds = details.map(d => d['内訳ID']);

    const data = {
      product: currentProduct,
      allProducts: allProducts, // プルダウン用
      params: getTableData_(SHEET_NAMES.PARAM, ss).filter(p => p['製品ID'] === currentProduct['製品ID']),
      processes: getTableData_(SHEET_NAMES.PROCESS, ss), // マスタは全製品共通
      groups: getTableData_(SHEET_NAMES.GROUP, ss),      // マスタは全製品共通
      equipment: getTableData_(SHEET_NAMES.EQUIPMENT, ss), // マスタは全製品共通
      works: works,
      details: details,
      choices: getTableData_(SHEET_NAMES.CHOICE, ss),
      parts: getTableData_(SHEET_NAMES.PART, ss),
      productParts: getTableData_(SHEET_NAMES.PRODUCT_PART, ss).filter(p => p['製品ID'] === currentProduct['製品ID']),
      partUsages: getTableData_(SHEET_NAMES.PART_USAGE, ss),
      versions: getTableData_(SHEET_NAMES.VERSION, ss).filter(v => v['製品ID'] === currentProduct['製品ID'])
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
      // ID自動採番 (製品コード-E-連番)
      let newId = newRow['内訳ID'];
      if (!newId) {
         // 簡単のため現在時刻ミリ秒を連番代わりにするか、本来はシートから最大値を取る
         const maxNum = data.length; 
         newId = 'MP2500-E-' + (maxNum + 1000).toString() + Date.now().toString().slice(-4);
         newRow['内訳ID'] = newId;
      }
      
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

// 作業内訳の一括保存（順序変更や一括編集用）
function saveMultipleWorkDetails(rowsToUpdate) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
    const sheet = ss.getSheetByName(SHEET_NAMES.WORK_DETAIL);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf('内訳ID');
    
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const user = Session.getActiveUser().getEmail();
    let updatedCount = 0;
    const historyEntries = [];

    // IDごとに現在の行インデックスとデータをマップ化
    const rowMap = {};
    for (let i = 1; i < data.length; i++) {
      rowMap[data[i][idIndex]] = { rowIndex: i, rowData: data[i] };
    }

    rowsToUpdate.forEach(({ newRow, baseRow }) => {
      const existing = rowMap[newRow['内訳ID']];
      if (!existing) return; // 新規追加は個別保存機能で行う前提
      
      const currentRow = existing.rowData;
      let conflict = false;
      const changes = [];

      headers.forEach((h, colIndex) => {
        if (['作成日時', '作成者', '更新日時', '更新者'].includes(h)) return;
        const currentVal = normalizeValue_(currentRow[colIndex]);
        const baseVal = normalizeValue_(baseRow[h]);
        const newVal = normalizeValue_(newRow[h]);

        if (currentVal !== baseVal && currentVal !== newVal) {
          conflict = true;
        } else if (newRow[h] !== undefined && currentVal !== newVal) {
          changes.push({ field: h, oldValue: currentVal, newValue: newVal });
          currentRow[colIndex] = newRow[h];
        }
      });

      if (!conflict && changes.length > 0) {
        currentRow[headers.indexOf('更新日時')] = now;
        currentRow[headers.indexOf('更新者')] = user;
        
        sheet.getRange(existing.rowIndex + 1, 1, 1, headers.length).setValues([currentRow]);
        updatedCount++;
        
        const operation = changes.some(c => c.field === '順序' || c.field === '作業ID') ? '順序・移動' : '変更';
        changes.forEach(c => {
          historyEntries.push([now, user, productIdOfDetail_(newRow['内訳ID']), '作業内訳', newRow['内訳ID'], operation, c.field, c.oldValue, c.newValue]);
        });
      }
    });

    if (historyEntries.length > 0) {
      const hSheet = ss.getSheetByName(SHEET_NAMES.HISTORY);
      const seq = Math.max(hSheet.getLastRow() - 1, 0);
      const hRows = historyEntries.map((c, i) => ['H-' + (seq + i + 1), c[0], c[1], c[2], c[3], c[4], c[5], c[6], c[7], c[8]]);
      hSheet.getRange(hSheet.getLastRow() + 1, 1, hRows.length, HEADERS.HISTORY.length).setValues(hRows);
    }

    return { ok: true, message: `${updatedCount}件の作業内訳を保存しました` };
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

// ----------------------------------------------------
// 部品・設備の紐付けデータ取得・保存 (F26, F60)
// ----------------------------------------------------

function getBindingData(type, detailId) {
  try {
    if (type === 'parts') {
      const partsAll = getTableData_(SHEET_NAMES.PART) || [];
      const usedAll = getTableData_(SHEET_NAMES.PART_USAGE) || [];
      const used = usedAll.filter(row => row['内訳ID'] === detailId);
      
      const available = partsAll.filter(p => p['状態'] !== '無効').map(p => ({
        '品番': p['品番'],
        '品名': p['品名'] + ' (' + p['品番'] + ')'
      }));
      
      return { ok: true, data: { used, available } };
      
    } else if (type === 'equipment') {
      const equipAll = getTableData_(SHEET_NAMES.EQUIPMENT) || [];
      const usedAll = getTableData_(SHEET_NAMES.EQUIPMENT_USAGE) || [];
      const used = usedAll.filter(row => row['内訳ID'] === detailId);
      
      const available = equipAll.filter(e => e['状態'] !== '無効').map(e => ({
        '設備ID': e['設備ID'],
        '設備名': e['設備名'] + ' (保有: ' + e['保有台数'] + ')',
        '保有台数': e['保有台数']
      }));
      
      return { ok: true, data: { used, available } };
    }
    
    return { ok: false, message: '不明な紐付け種別です' };
  } catch (e) {
    return { ok: false, message: e.toString() };
  }
}

function saveBindingData(type, detailId, newDataList) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
    const sheetName = type === 'parts' ? SHEET_NAMES.PART_USAGE : SHEET_NAMES.EQUIPMENT_USAGE;
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) return { ok: false, message: 'シートが見つかりません' };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf('内訳ID');
    
    // C36: 設備必要台数の検証
    if (type === 'equipment') {
       const equipAll = getTableData_(SHEET_NAMES.EQUIPMENT) || [];
       for (const row of newDataList) {
         const eq = equipAll.find(e => e['設備ID'] === row['設備ID']);
         if (eq && Number(row['必要台数']) > Number(eq['保有台数'])) {
           return { ok: false, message: `エラー: ${eq['設備名']} の必要台数(${row['必要台数']})が保有台数(${eq['保有台数']})を超えています。` };
         }
       }
    }
    
    // 現在の対象IDの行を全て削除し、洗い替える
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][idIndex] === detailId) {
        sheet.deleteRow(i + 1);
      }
    }
    
    // 新しいデータを追加
    if (newDataList && newDataList.length > 0) {
      const appendRows = newDataList.map(row => {
        return headers.map(h => {
          if (h === '内訳ID') return detailId;
          return row[h] !== undefined ? row[h] : '';
        });
      });
      sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, headers.length).setValues(appendRows);
    }
    
    // 履歴を残す (簡易)
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const user = Session.getActiveUser().getEmail();
    appendHistory_(detailId, '変更', [{ field: type === 'parts' ? '使用部品' : '使用設備', oldValue: '', newValue: '洗い替え更新' }], now, user);
    
    return { ok: true };
    
  } catch (e) {
    return { ok: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ----------------------------------------------------
// マスタ設定（S07）の保存処理
// ----------------------------------------------------

function saveMasterData(type, newRow) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
    
    let sheetName = '';
    let idKey = '';
    let prefix = '';
    
    if (type === 'equipment') {
      sheetName = SHEET_NAMES.EQUIPMENT;
      idKey = '設備ID';
      prefix = 'EQ-';
    } else if (type === 'process') {
      sheetName = SHEET_NAMES.PROCESS;
      idKey = '工程ID';
      prefix = 'S-';
    } else if (type === 'group') {
      sheetName = SHEET_NAMES.GROUP;
      idKey = '班ID';
      prefix = 'G-';
    } else {
      return { ok: false, message: '不明なマスタ種別です' };
    }
    
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { ok: false, message: '対象シートが見つかりません' };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf(idKey);
    
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const user = Session.getActiveUser().getEmail();
    
    // 既存チェック
    let rowIndex = -1;
    if (newRow[idKey]) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][idIndex] === newRow[idKey]) {
          rowIndex = i;
          break;
        }
      }
    }
    
    if (rowIndex === -1) {
      // 新規追加
      // ID自動採番 (プレフィックス + 連番)
      if (!newRow[idKey]) {
         const maxNum = data.length;
         newRow[idKey] = prefix + (maxNum + 100).toString() + Date.now().toString().slice(-3);
      }
      
      const appendData = headers.map(h => {
        if (h === '作成日時' || h === '更新日時') return now;
        if (h === '作成者' || h === '更新者') return user;
        return newRow[h] !== undefined ? newRow[h] : '';
      });
      sheet.appendRow(appendData);
      return { ok: true, message: '追加しました' };
    } else {
      // 既存行の更新
      const currentRow = data[rowIndex];
      headers.forEach((h, colIndex) => {
        if (['作成日時', '作成者', '更新日時', '更新者'].includes(h)) return;
        if (newRow[h] !== undefined) {
          currentRow[colIndex] = newRow[h];
        }
      });
      currentRow[headers.indexOf('更新日時')] = now;
      currentRow[headers.indexOf('更新者')] = user;
      
      sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([currentRow]);
      return { ok: true, message: '保存しました' };
    }
    
  } catch (e) {
    return { ok: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ----------------------------------------------------
// 製品（型式）の管理 (S01 / F47)
// ----------------------------------------------------

function switchProduct(productId) {
  // サーバー側のユーザープロパティ（セッション）に選択中の製品IDを保存する
  PropertiesService.getUserProperties().setProperty('CURRENT_PRODUCT_ID', productId);
  return { ok: true, message: '製品を切り替えました' };
}

function cloneProduct(newProductCode, newProductName) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // 製品丸ごとコピーは時間がかかるため30秒
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
    const pSheet = ss.getSheetByName(SHEET_NAMES.PRODUCT);
    
    // 1. 現在の製品IDを取得
    const currentProductId = PropertiesService.getUserProperties().getProperty('CURRENT_PRODUCT_ID') || 'P-001';
    
    // 2. 新しい製品IDの採番
    const pData = pSheet.getDataRange().getValues();
    const newProductId = 'P-' + (pData.length + 100).toString() + Date.now().toString().slice(-3);
    
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const user = Session.getActiveUser().getEmail();
    
    // 3. 製品シートに新規行を追加
    const pHeaders = pData[0];
    let sourceProductRow = pData.find((r, i) => i > 0 && r[pHeaders.indexOf('製品ID')] === currentProductId);
    
    if (!sourceProductRow) {
       // 見つからなければデフォルト値で作成
       sourceProductRow = pHeaders.map(() => '');
    }
    
    const newProductRow = [...sourceProductRow];
    newProductRow[pHeaders.indexOf('製品ID')] = newProductId;
    newProductRow[pHeaders.indexOf('製品コード')] = newProductCode;
    newProductRow[pHeaders.indexOf('型式名')] = newProductName;
    newProductRow[pHeaders.indexOf('派生元製品ID')] = currentProductId;
    newProductRow[pHeaders.indexOf('作成日時')] = now;
    newProductRow[pHeaders.indexOf('更新日時')] = now;
    newProductRow[pHeaders.indexOf('作成者')] = user;
    newProductRow[pHeaders.indexOf('更新者')] = user;
    pSheet.appendRow(newProductRow);
    
    // 4. 作業(ブロック) と  作業内訳 のコピー
    // ※本来はパラメータや使用部品、使用設備もコピーする（F47）が、プロトタイプとして枠組みをコピー
    const wSheet = ss.getSheetByName(SHEET_NAMES.WORK);
    const dSheet = ss.getSheetByName(SHEET_NAMES.WORK_DETAIL);
    
    const wData = wSheet.getDataRange().getValues();
    const wHeaders = wData[0];
    const dData = dSheet.getDataRange().getValues();
    const dHeaders = dData[0];
    
    const newWorks = [];
    const newDetails = [];
    
    for (let i = 1; i < wData.length; i++) {
      if (wData[i][wHeaders.indexOf('製品ID')] === currentProductId && wData[i][wHeaders.indexOf('状態')] !== '無効') {
        const oldWorkId = wData[i][wHeaders.indexOf('作業ID')];
        // S-連番 または 型式-W-連番 の形。新規に採番
        const newWorkId = newProductCode + '-W-' + Date.now().toString().slice(-6) + i;
        
        const wRow = [...wData[i]];
        wRow[wHeaders.indexOf('作業ID')] = newWorkId;
        wRow[wHeaders.indexOf('製品ID')] = newProductId;
        wRow[wHeaders.indexOf('派生元作業ID')] = oldWorkId;
        wRow[wHeaders.indexOf('作成日時')] = now;
        wRow[wHeaders.indexOf('更新日時')] = now;
        newWorks.push(wRow);
        
        // 子（内訳）のコピー
        for (let j = 1; j < dData.length; j++) {
          if (dData[j][dHeaders.indexOf('作業ID')] === oldWorkId && dData[j][dHeaders.indexOf('状態')] !== '無効') {
            const oldDetailId = dData[j][dHeaders.indexOf('内訳ID')];
            const newDetailId = newProductCode + '-E-' + Date.now().toString().slice(-6) + j;
            
            const dRow = [...dData[j]];
            dRow[dHeaders.indexOf('内訳ID')] = newDetailId;
            dRow[dHeaders.indexOf('作業ID')] = newWorkId;
            dRow[dHeaders.indexOf('派生元内訳ID')] = oldDetailId;
            dRow[dHeaders.indexOf('確度')] = '暫定'; // F47: コピー後は暫定
            dRow[dHeaders.indexOf('作成日時')] = now;
            dRow[dHeaders.indexOf('更新日時')] = now;
            newDetails.push(dRow);
          }
        }
      }
    }
    
    if (newWorks.length > 0) wSheet.getRange(wSheet.getLastRow() + 1, 1, newWorks.length, wHeaders.length).setValues(newWorks);
    if (newDetails.length > 0) dSheet.getRange(dSheet.getLastRow() + 1, 1, newDetails.length, dHeaders.length).setValues(newDetails);
    
    // 切り替え
    PropertiesService.getUserProperties().setProperty('CURRENT_PRODUCT_ID', newProductId);
    
    return { ok: true, message: '製品を複製しました', newProductId: newProductId };
  } catch(e) {
    return { ok: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ----------------------------------------------------
// 組立ブロック（作業）の管理 (S01 / F70, F02)
// ----------------------------------------------------

function saveWorkBlock(newRow) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
    const sheet = ss.getSheetByName(SHEET_NAMES.WORK);
    if (!sheet) return { ok: false, message: 'シートが見つかりません' };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf('作業ID');
    
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const user = Session.getActiveUser().getEmail();
    
    let rowIndex = -1;
    if (newRow['作業ID']) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][idIndex] === newRow['作業ID']) {
          rowIndex = i;
          break;
        }
      }
    }
    
    if (rowIndex === -1) {
      // 新規追加
      if (!newRow['作業ID']) {
         newRow['作業ID'] = 'MP2500-W-' + (data.length + 1000).toString() + Date.now().toString().slice(-4);
      }
      const appendData = headers.map(h => {
        if (h === '作成日時' || h === '更新日時') return now;
        if (h === '作成者' || h === '更新者') return user;
        return newRow[h] !== undefined ? newRow[h] : '';
      });
      sheet.appendRow(appendData);
      return { ok: true, message: '追加しました' };
    } else {
      // 更新
      const currentRow = data[rowIndex];
      headers.forEach((h, colIndex) => {
        if (['作成日時', '作成者', '更新日時', '更新者', '作業ID', '製品ID'].includes(h)) return;
        if (newRow[h] !== undefined) {
          currentRow[colIndex] = newRow[h];
        }
      });
      currentRow[headers.indexOf('更新日時')] = now;
      currentRow[headers.indexOf('更新者')] = user;
      sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([currentRow]);
      return { ok: true, message: '保存しました' };
    }
  } catch(e) {
    return { ok: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function cloneWorkBlock(workId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
    const workSheet = ss.getSheetByName(SHEET_NAMES.WORK);
    const detailSheet = ss.getSheetByName(SHEET_NAMES.WORK_DETAIL);
    
    const wData = workSheet.getDataRange().getValues();
    const wHeaders = wData[0];
    
    let targetWorkRow = null;
    for (let i=1; i<wData.length; i++) {
      if (wData[i][wHeaders.indexOf('作業ID')] === workId) {
        targetWorkRow = wData[i];
        break;
      }
    }
    if (!targetWorkRow) return { ok:false, message: 'コピー元のブロックが見つかりません' };
    
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const user = Session.getActiveUser().getEmail();
    
    // 1. 親ブロックのコピー
    const newWorkId = 'MP2500-W-' + Date.now();
    const newWorkRow = [...targetWorkRow];
    newWorkRow[wHeaders.indexOf('作業ID')] = newWorkId;
    newWorkRow[wHeaders.indexOf('作業名')] = targetWorkRow[wHeaders.indexOf('作業名')] + ' (コピー)';
    newWorkRow[wHeaders.indexOf('派生元作業ID')] = workId; // F48 派生元の記録
    newWorkRow[wHeaders.indexOf('作成日時')] = now;
    newWorkRow[wHeaders.indexOf('更新日時')] = now;
    newWorkRow[wHeaders.indexOf('作成者')] = user;
    newWorkRow[wHeaders.indexOf('更新者')] = user;
    workSheet.appendRow(newWorkRow);
    
    // 2. 紐づく子内訳のコピー
    const dData = detailSheet.getDataRange().getValues();
    const dHeaders = dData[0];
    const appendRows = [];
    
    for (let i=1; i<dData.length; i++) {
      if (dData[i][dHeaders.indexOf('作業ID')] === workId && dData[i][dHeaders.indexOf('状態')] !== '無効') {
        const row = [...dData[i]];
        row[dHeaders.indexOf('内訳ID')] = 'MP2500-E-' + Date.now() + i;
        row[dHeaders.indexOf('作業ID')] = newWorkId;
        row[dHeaders.indexOf('派生元内訳ID')] = dData[i][dHeaders.indexOf('内訳ID')];
        row[dHeaders.indexOf('作成日時')] = now;
        row[dHeaders.indexOf('更新日時')] = now;
        row[dHeaders.indexOf('作成者')] = user;
        row[dHeaders.indexOf('更新者')] = user;
        appendRows.push(row);
      }
    }
    
    if (appendRows.length > 0) {
      detailSheet.getRange(detailSheet.getLastRow() + 1, 1, appendRows.length, dHeaders.length).setValues(appendRows);
    }
    
    return { ok: true, message: '複製しました' };
  } catch(e) {
    return { ok: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function reorderAllBlocks() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
    const sheet = ss.getSheetByName(SHEET_NAMES.WORK);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const orderIndex = headers.indexOf('順序');
    
    // データ行だけ抽出し、現在の順序でソート
    let rows = data.slice(1).map((row, idx) => ({ row, idx: idx + 1 }));
    rows.sort((a, b) => Number(a.row[orderIndex] || 999) - Number(b.row[orderIndex] || 999));
    
    // 10刻みで振り直し
    let order = 10;
    rows.forEach(r => {
      sheet.getRange(r.idx + 1, orderIndex + 1).setValue(order);
      order += 10;
    });
    
    return { ok: true };
  } catch(e) {
    return { ok: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function getRecords(detailId) {
  try {
    const records = getTableData_(SHEET_NAMES.RECORD).filter(r => r['内訳ID'] === detailId);
    return { ok: true, data: records };
  } catch (e) {
    return { ok: false, message: e.toString() };
  }
}

function saveRecordData(newRow) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
    const sheet = ss.getSheetByName(SHEET_NAMES.RECORD);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idIndex = headers.indexOf('観測ID');
    
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const user = Session.getActiveUser().getEmail();
    
    let rowIndex = -1;
    if (newRow['観測ID']) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][idIndex] === newRow['観測ID']) {
          rowIndex = i;
          break;
        }
      }
    }

    if (rowIndex === -1) {
      // 追加
      newRow['観測ID'] = 'O-' + (data.length + 1000).toString() + Date.now().toString().slice(-4);
      const appendData = headers.map(h => {
        if (h === '作成日時' || h === '更新日時') return now;
        if (h === '作成者' || h === '更新者') return user;
        return newRow[h] !== undefined ? newRow[h] : '';
      });
      sheet.appendRow(appendData);
      return { ok: true, message: '追加しました', data: { row: newRow } };
    } else {
      // 更新
      const currentRow = data[rowIndex];
      headers.forEach((h, colIndex) => {
        if (['作成日時', '作成者', '更新日時', '更新者', '観測ID', '内訳ID'].includes(h)) return;
        if (newRow[h] !== undefined) {
          currentRow[colIndex] = newRow[h];
        }
      });
      currentRow[headers.indexOf('更新日時')] = now;
      currentRow[headers.indexOf('更新者')] = user;
      sheet.getRange(rowIndex + 1, 1, 1, headers.length).setValues([currentRow]);
      return { ok: true, message: '保存しました', data: { row: newRow } };
    }
  } catch (e) {
    return { ok: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}


/**
 * 品番の正規化（03_データ構造 4.12）：前後空白除去・全角英数記号→半角・英字大文字化
 */
/** SHEET_NAMES/HEADERS のキーでシートを取得し、無ければヘッダ行付きで作成する */
function getOrCreateSheet_(ss, key) {
  const name = SHEET_NAMES[key];
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headers = HEADERS[key];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function normalizePartNo_(v) {
  if (v === null || v === undefined) return '';
  let s = String(v).replace(/[\u3000\s]+/g, ' ').trim();
  // 全角英数・記号（U+FF01〜U+FF5E）→ 半角
  s = s.replace(/[\uFF01-\uFF5E]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0));
  return s.toUpperCase();
}

/**
 * 部品表（BOM）CSV の取込。
 * 形式：ヘッダあり／A:製品型番 B:部品品番 C:部品名。引用符内のカンマ・改行に対応。
 * 現在の製品の CSV型番 と一致する行だけを製品部品に登録する（CSV型番 が未設定なら全行）。
 */
function importPartsCSV(csvText) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // 30秒
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
    // 部品系シートは後から追加された定義のため、無ければヘッダ付きで自動作成する
    const partSheet = getOrCreateSheet_(ss, 'PART');
    const prodPartSheet = getOrCreateSheet_(ss, 'PRODUCT_PART');
    getOrCreateSheet_(ss, 'PART_USAGE');
    const productSheet = ss.getSheetByName(SHEET_NAMES.PRODUCT);

    // BOM（U+FEFF）除去のうえ、引用符を考慮して解析する
    const text = String(csvText || '').replace(/^\uFEFF/, '');
    if (text.trim() === '') return { ok: false, message: 'データがありません' };

    let rows;
    try {
      rows = Utilities.parseCsv(text);
    } catch (e) {
      return { ok: false, message: 'CSVの解析に失敗しました（引用符の対応が崩れている可能性があります）: ' + e.message };
    }
    rows = rows.filter(r => r.join('').trim() !== '');
    if (rows.length === 0) return { ok: false, message: 'データがありません' };

    // 1行目がヘッダなら飛ばす（品番らしい値なら data として扱う）
    const firstCell = String(rows[0][0] || '');
    const startRow = /製品型番|型番|製品/.test(firstCell) ? 1 : 0;
    if (rows.length <= startRow) return { ok: false, message: 'データがありません' };

    const currentProductId = PropertiesService.getUserProperties().getProperty('CURRENT_PRODUCT_ID') || 'P-001';
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const user = Session.getActiveUser().getEmail();

    // 現在の製品の CSV型番（未設定なら型番での絞り込みをしない）
    let csvTypeCode = '';
    if (productSheet) {
      const prodData = productSheet.getDataRange().getValues();
      const prodHeaders = prodData[0];
      const idIdx = prodHeaders.indexOf('製品ID');
      const csvIdx = prodHeaders.indexOf('CSV型番');
      for (let i = 1; i < prodData.length; i++) {
        if (prodData[i][idIdx] === currentProductId) {
          csvTypeCode = normalizePartNo_(csvIdx >= 0 ? prodData[i][csvIdx] : '');
          break;
        }
      }
    }

    // 既存部品の取得
    const partData = partSheet.getDataRange().getValues();
    const partHeaders = partData[0];
    const pnIdx = partHeaders.indexOf('品番');
    const pnameIdx = partHeaders.indexOf('品名');
    const partMap = {}; // { 品番: rowIndex }
    for (let i = 1; i < partData.length; i++) {
      const pn = normalizePartNo_(partData[i][pnIdx]);
      if (pn) partMap[pn] = i;
    }

    // 既存製品部品の取得（現在の製品に紐づくもののみ）
    const prodPartData = prodPartSheet.getDataRange().getValues();
    const ppHeaders = prodPartData[0];
    const ppSeen = {};
    for (let i = 1; i < prodPartData.length; i++) {
      if (prodPartData[i][ppHeaders.indexOf('製品ID')] === currentProductId) {
        ppSeen[normalizePartNo_(prodPartData[i][ppHeaders.indexOf('品番')])] = true;
      }
    }

    let addedCount = 0;     // 部品マスタへの新規追加
    let updatedCount = 0;   // 品名の更新
    let linkedCount = 0;    // 製品部品への新規紐付け
    let skippedCount = 0;   // 型番違いで対象外
    let dupCount = 0;       // CSV内の重複行
    const appendParts = [];
    const appendProdParts = [];
    const nameUpdates = []; // { row, name }
    const seenInCsv = {};

    for (let i = startRow; i < rows.length; i++) {
      const cols = rows[i];
      if (!cols || cols.length < 2) continue;

      const typeCode = normalizePartNo_(cols[0]);
      const normPartNo = normalizePartNo_(cols[1]);
      const partName = String(cols[2] === undefined ? '' : cols[2]).trim();
      if (!normPartNo) continue;

      // 現在の製品の CSV型番 と一致する行のみ取り込む
      if (csvTypeCode && typeCode && typeCode !== csvTypeCode) {
        skippedCount++;
        continue;
      }

      if (seenInCsv[normPartNo]) { dupCount++; continue; }
      seenInCsv[normPartNo] = true;

      // 部品マスタの更新・追加
      if (partMap[normPartNo] !== undefined) {
        const rIdx = partMap[normPartNo];
        if (partName && String(partData[rIdx][pnameIdx]) !== partName) {
          partData[rIdx][pnameIdx] = partName;
          nameUpdates.push({ row: rIdx + 1, name: partName });
          updatedCount++;
        }
      } else {
        const newRow = Array(partHeaders.length).fill('');
        newRow[pnIdx] = normPartNo;
        newRow[pnameIdx] = partName;
        newRow[partHeaders.indexOf('状態')] = '有効';
        newRow[partHeaders.indexOf('作成日時')] = now;
        newRow[partHeaders.indexOf('作成者')] = user;
        appendParts.push(newRow);
        addedCount++;
      }

      // 製品部品（簡易BOM）の追加
      if (!ppSeen[normPartNo]) {
        const newRow = Array(ppHeaders.length).fill('');
        newRow[ppHeaders.indexOf('製品ID')] = currentProductId;
        newRow[ppHeaders.indexOf('品番')] = normPartNo;
        newRow[ppHeaders.indexOf('状態')] = '有効';
        newRow[ppHeaders.indexOf('取込日')] = now;
        newRow[ppHeaders.indexOf('作成日時')] = now;
        newRow[ppHeaders.indexOf('作成者')] = user;
        appendProdParts.push(newRow);
        ppSeen[normPartNo] = true;
        linkedCount++;
      }
    }

    if (addedCount === 0 && updatedCount === 0 && linkedCount === 0 && skippedCount > 0) {
      return { ok: false, message: 'CSVの製品型番が現在の製品のCSV型番（' + csvTypeCode + '）と一致しません。製品を切り替えるか、製品マスタのCSV型番を確認してください。' };
    }

    nameUpdates.forEach(u => {
      partSheet.getRange(u.row, pnameIdx + 1).setValue(u.name);
      if (partHeaders.indexOf('更新日時') >= 0) partSheet.getRange(u.row, partHeaders.indexOf('更新日時') + 1).setValue(now);
      if (partHeaders.indexOf('更新者') >= 0) partSheet.getRange(u.row, partHeaders.indexOf('更新者') + 1).setValue(user);
    });
    if (appendParts.length > 0) {
      partSheet.getRange(partSheet.getLastRow() + 1, 1, appendParts.length, partHeaders.length).setValues(appendParts);
    }
    if (appendProdParts.length > 0) {
      prodPartSheet.getRange(prodPartSheet.getLastRow() + 1, 1, appendProdParts.length, ppHeaders.length).setValues(appendProdParts);
    }

    // CSVになかった製品部品は自動無効化しない（履歴が壊れるため）

    return { ok: true, data: { added: addedCount, updated: updatedCount, linked: linkedCount, skipped: skippedCount, duplicated: dupCount } };
  } catch (e) {
    return { ok: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function savePartMaster(newRow) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
    const sheet = ss.getSheetByName(SHEET_NAMES.PART);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const user = Session.getActiveUser().getEmail();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][headers.indexOf('品番')] === newRow['品番']) {
        const currentRow = data[i];
        ['部品区分', '入数', '荷姿', '置場', '搬入条件'].forEach(col => {
          currentRow[headers.indexOf(col)] = newRow[col];
        });
        currentRow[headers.indexOf('更新日時')] = now;
        currentRow[headers.indexOf('更新者')] = user;
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([currentRow]);
        return { ok: true, message: '保存しました' };
      }
    }
    return { ok: false, message: '部品が見つかりません' };
  } catch (e) {
    return { ok: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function generateKitData(targetQty, axis) {
  try {
    const currentProductId = PropertiesService.getUserProperties().getProperty('CURRENT_PRODUCT_ID') || 'P-001';
    
    // 必要なデータを取得
    const details = getTableData_(SHEET_NAMES.WORK_DETAIL).filter(d => d['状態'] !== '無効');
    const works = getTableData_(SHEET_NAMES.WORK).filter(w => w['製品ID'] === currentProductId && w['状態'] !== '無効');
    const parts = getTableData_(SHEET_NAMES.PART);
    const usages = getTableData_(SHEET_NAMES.PART_USAGE);
    const processes = getTableData_(SHEET_NAMES.PROCESS);
    
    // 有効な小工程IDリスト
    const workIds = works.map(w => w['作業ID']);
    const activeDetails = details.filter(d => workIds.includes(d['作業ID']));
    const activeDetailIds = activeDetails.map(d => d['内訳ID']);
    
    // 対象の部品使用リスト
    const activeUsages = usages.filter(u => activeDetailIds.includes(u['内訳ID']));
    
    // マッピング
    const partMap = {};
    parts.forEach(p => partMap[p['品番']] = p);
    
    const workMap = {};
    works.forEach(w => workMap[w['作業ID']] = w);
    
    const detailMap = {};
    activeDetails.forEach(d => detailMap[d['内訳ID']] = d);
    
    const processMap = {};
    processes.forEach(p => processMap[p['工程ID']] = p['工程名']);

    // 集計
    const results = [];
    
    activeUsages.forEach(usage => {
      const detail = detailMap[usage['内訳ID']];
      const work = workMap[detail['作業ID']];
      const part = partMap[usage['品番']];
      
      if (!part || !detail || !work) return;
      
      // 軸に応じたグループ名
      let groupName = '未分類';
      if (axis === '大工程別') {
         groupName = processMap[work['工程ID']] || work['大工程'] || '未分類';
      } else if (axis === '組立ブロック別') {
         groupName = work['作業名'] || '未分類';
      } else if (axis === '中工程別') {
         groupName = work['作業名'] || '未分類'; // 今回は2階層なのでブロック名と同義
      }
      
      const unitQty = Number(usage['使用数']) || 0;
      if (unitQty <= 0) return;
      
      const totalQty = unitQty * targetQty;
      const packageQty = Number(part['入数']) || 0;
      
      let boxes = 0;
      let remainder = totalQty;
      if (packageQty > 0) {
         boxes = Math.floor(totalQty / packageQty);
         remainder = totalQty % packageQty;
      }
      
      // 同じグループ・同じ品番があれば合算（中工程別などの要件次第だが、基本は合算）
      const existing = results.find(r => r.groupName === groupName && r['品番'] === part['品番']);
      if (existing) {
         existing.unitQty += unitQty;
         existing.totalQty += totalQty;
         if (packageQty > 0) {
            existing.boxes = Math.floor(existing.totalQty / packageQty);
            existing.remainder = existing.totalQty % packageQty;
         } else {
            existing.remainder = existing.totalQty;
         }
      } else {
         results.push({
           groupName: groupName,
           '品番': part['品番'],
           '品名': part['品名'],
           '部品区分': part['部品区分'],
           unitQty: unitQty,
           totalQty: totalQty,
           '入数': packageQty,
           boxes: boxes,
           remainder: remainder,
           '荷姿': part['荷姿'] || '',
           '置場': part['置場'] || ''
         });
      }
    });
    
    // 並び替え (グループ名 -> 品番)
    results.sort((a, b) => {
      if (a.groupName < b.groupName) return -1;
      if (a.groupName > b.groupName) return 1;
      if (a['品番'] < b['品番']) return -1;
      if (a['品番'] > b['品番']) return 1;
      return 0;
    });
    
    return { ok: true, data: results };
    
  } catch(e) {
    return { ok: false, message: e.toString() };
  }
}
