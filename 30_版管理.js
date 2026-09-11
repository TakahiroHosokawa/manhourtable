// =========================================
// 30_版管理.gs
// =========================================

function commitVersionData(reason) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // 30秒待ち
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
    const versionSheet = ss.getSheetByName(SHEET_NAMES.VERSION);
    const detailSheet = ss.getSheetByName(SHEET_NAMES.VERSION_DETAIL);
    const prodSheet = ss.getSheetByName(SHEET_NAMES.PRODUCT);
    
    if (!versionSheet || !detailSheet || !prodSheet) return { ok: false, message: '必要なシートが存在しません' };

    const currentProductId = PropertiesService.getUserProperties().getProperty('CURRENT_PRODUCT_ID') || 'P-001';
    
    // 現在の製品の情報を取得し、版番号を採番
    const prodData = prodSheet.getDataRange().getValues();
    const prodHeaders = prodData[0];
    let prodRowIndex = -1;
    let nextVersionNumber = 1;
    
    for (let i = 1; i < prodData.length; i++) {
      if (prodData[i][prodHeaders.indexOf('製品ID')] === currentProductId) {
        prodRowIndex = i;
        const curVer = Number(prodData[i][prodHeaders.indexOf('現行確定版番号')]) || 0;
        nextVersionNumber = curVer + 1;
        break;
      }
    }
    
    if (prodRowIndex === -1) return { ok: false, message: '製品情報が見つかりません' };

    const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    const user = Session.getActiveUser().getEmail();
    const versionId = 'V-' + Date.now() + '-' + nextVersionNumber;

    // 1. 版 (VERSION) シートへの追加
    const vHeaders = versionSheet.getDataRange().getValues()[0];
    const newVerRow = Array(vHeaders.length).fill('');
    newVerRow[vHeaders.indexOf('版ID')] = versionId;
    newVerRow[vHeaders.indexOf('製品ID')] = currentProductId;
    newVerRow[vHeaders.indexOf('版番号')] = nextVersionNumber;
    newVerRow[vHeaders.indexOf('確定日時')] = now;
    newVerRow[vHeaders.indexOf('確定者')] = user;
    newVerRow[vHeaders.indexOf('改訂理由')] = reason;
    newVerRow[vHeaders.indexOf('作成日時')] = now;
    newVerRow[vHeaders.indexOf('作成者')] = user;
    newVerRow[vHeaders.indexOf('更新日時')] = now;
    newVerRow[vHeaders.indexOf('更新者')] = user;
    versionSheet.appendRow(newVerRow);

    // 2. 製品シートの「現行確定版番号」を更新
    prodSheet.getRange(prodRowIndex + 1, prodHeaders.indexOf('現行確定版番号') + 1).setValue(nextVersionNumber);

    // 3. 現在の有効な工数データ（スナップショット）を取得して計算
    const works = getTableData_(SHEET_NAMES.WORK, ss).filter(w => w['製品ID'] === currentProductId && w['状態'] !== '無効');
    const workIds = works.map(w => w['作業ID']);
    const details = getTableData_(SHEET_NAMES.WORK_DETAIL, ss).filter(d => workIds.includes(d['作業ID']) && d['状態'] !== '無効');
    const paramsMap = {};
    getTableData_(SHEET_NAMES.PARAM, ss).filter(p => p['製品ID'] === currentProductId).forEach(p => {
      paramsMap[p['パラメータID']] = p['値'];
    });
    const processes = getTableData_(SHEET_NAMES.PROCESS, ss);
    const groups = getTableData_(SHEET_NAMES.GROUP, ss);
    
    const processMap = {}; processes.forEach(p => processMap[p['工程ID']] = p['工程名']);
    const groupMap = {}; groups.forEach(g => groupMap[g['班ID']] = g);
    const workMap = {}; works.forEach(w => workMap[w['作業ID']] = w);
    
    const prodObj = getTableData_(SHEET_NAMES.PRODUCT, ss).find(p => p['製品ID'] === currentProductId);
    const allowanceRate = prodObj['余裕率'] || 15;

    // 版明細 (VERSION_DETAIL) への追加準備
    const vdHeaders = detailSheet.getDataRange().getValues()[0];
    const appendRows = [];

    details.forEach(d => {
      const w = workMap[d['作業ID']];
      if (!w) return;
      const g = groupMap[w['班ID']] || {};
      
      const calc = calculateDetail(d, allowanceRate, paramsMap);
      
      const vdRow = Array(vdHeaders.length).fill('');
      vdRow[vdHeaders.indexOf('版ID')] = versionId;
      vdRow[vdHeaders.indexOf('内訳ID')] = d['内訳ID'];
      vdRow[vdHeaders.indexOf('工程名')] = processMap[w['工程ID']] || '';
      vdRow[vdHeaders.indexOf('班名')] = g['班名'] || '';
      vdRow[vdHeaders.indexOf('担当職種')] = g['担当職種'] || '';
      vdRow[vdHeaders.indexOf('担当会社')] = g['担当会社'] || '';
      vdRow[vdHeaders.indexOf('大工程')] = w['大工程'] || '';
      vdRow[vdHeaders.indexOf('作業名')] = w['作業名'] || '';
      vdRow[vdHeaders.indexOf('作業順序')] = w['順序'];
      vdRow[vdHeaders.indexOf('内訳順序')] = d['順序'];
      vdRow[vdHeaders.indexOf('作業内訳名')] = d['作業内訳名'];
      vdRow[vdHeaders.indexOf('区分')] = d['区分'];
      vdRow[vdHeaders.indexOf('作業方式')] = d['作業方式'];
      vdRow[vdHeaders.indexOf('単位')] = d['単位'];
      vdRow[vdHeaders.indexOf('時間方式')] = d['時間方式'];
      vdRow[vdHeaders.indexOf('正味時間_現状')] = d['正味時間_現状'];
      vdRow[vdHeaders.indexOf('数量方式')] = d['数量方式'];
      vdRow[vdHeaders.indexOf('固定数量')] = d['固定数量'];
      vdRow[vdHeaders.indexOf('数量パラメータID')] = d['数量パラメータID'];
      vdRow[vdHeaders.indexOf('係数')] = d['係数'];
      vdRow[vdHeaders.indexOf('人数_現状')] = d['人数_現状'];
      vdRow[vdHeaders.indexOf('上限人数')] = d['上限人数'];
      vdRow[vdHeaders.indexOf('正味時間_目標')] = d['正味時間_目標'];
      vdRow[vdHeaders.indexOf('人数_目標')] = d['人数_目標'];
      vdRow[vdHeaders.indexOf('目標メモ')] = d['目標メモ'];
      vdRow[vdHeaders.indexOf('中断可否')] = d['中断可否'];
      vdRow[vdHeaders.indexOf('先行作業')] = d['先行作業'];
      // 計算結果
      vdRow[vdHeaders.indexOf('工数_現状')] = calc.current.workSec;
      vdRow[vdHeaders.indexOf('所要_現状')] = calc.current.durationSec;
      vdRow[vdHeaders.indexOf('工数_目標')] = calc.target.workSec;
      vdRow[vdHeaders.indexOf('所要_目標')] = calc.target.durationSec;
      
      appendRows.push(vdRow);
    });

    if (appendRows.length > 0) {
      detailSheet.getRange(detailSheet.getLastRow() + 1, 1, appendRows.length, vdHeaders.length).setValues(appendRows);
    }

    // 履歴にも「版確定」として記録
    appendHistory_(currentProductId, '版確定', [{ field: '版番号', oldValue: '', newValue: nextVersionNumber }], now, user);

    return { ok: true, data: { versionNumber: nextVersionNumber } };

  } catch (e) {
    return { ok: false, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function exportVersionToSpreadsheet(versionId) {
  try {
    const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'));
    const currentProductId = PropertiesService.getUserProperties().getProperty('CURRENT_PRODUCT_ID') || 'P-001';
    const prodObj = getTableData_(SHEET_NAMES.PRODUCT, ss).find(p => p['製品ID'] === currentProductId);
    const prodName = prodObj ? prodObj['型式名'] : currentProductId;
    
    // 出力データ構築用
    let outputRows = [];
    let headers = [];
    let ssName = `工数表_${prodName}`;

    if (versionId === 'current') {
      ssName += `_未確定(現行)`;
      // 現在の有効な工数データを取得してフラット化
      const works = getTableData_(SHEET_NAMES.WORK, ss).filter(w => w['製品ID'] === currentProductId && w['状態'] !== '無効');
      const workIds = works.map(w => w['作業ID']);
      const details = getTableData_(SHEET_NAMES.WORK_DETAIL, ss).filter(d => workIds.includes(d['作業ID']) && d['状態'] !== '無効');
      const paramsMap = {};
      getTableData_(SHEET_NAMES.PARAM, ss).filter(p => p['製品ID'] === currentProductId).forEach(p => { paramsMap[p['パラメータID']] = p['値']; });
      
      const processes = getTableData_(SHEET_NAMES.PROCESS, ss);
      const groups = getTableData_(SHEET_NAMES.GROUP, ss);
      const processMap = {}; processes.forEach(p => processMap[p['工程ID']] = p['工程名']);
      const groupMap = {}; groups.forEach(g => groupMap[g['班ID']] = g);
      const workMap = {}; works.forEach(w => workMap[w['作業ID']] = w);
      const allowanceRate = prodObj['余裕率'] || 15;

      headers = ['工程名', '班名', '作業順序', '作業名', '内訳順序', '作業内訳名', '区分', '作業方式', '単位', '正味時間(秒)', '人数', '工数(分)', '所要(分)'];
      
      // ブロック順 -> 内訳順
      details.sort((a, b) => {
        const wA = workMap[a['作業ID']];
        const wB = workMap[b['作業ID']];
        if (wA['順序'] !== wB['順序']) return wA['順序'] - wB['順序'];
        return a['順序'] - b['順序'];
      });

      details.forEach(d => {
        const w = workMap[d['作業ID']];
        const g = groupMap[w['班ID']] || {};
        const calc = calculateDetail(d, allowanceRate, paramsMap);
        
        outputRows.push([
          processMap[w['工程ID']] || '',
          g['班名'] || '',
          w['順序'],
          w['作業名'],
          d['順序'],
          d['作業内訳名'],
          d['区分'],
          d['作業方式'],
          d['単位'],
          d['正味時間_現状'],
          d['人数_現状'],
          Math.round((calc.current.workSec / 60) * 10) / 10,
          Math.round((calc.current.durationSec / 60) * 10) / 10
        ]);
      });

    } else {
      // 指定された版のデータを取得
      const versions = getTableData_(SHEET_NAMES.VERSION, ss);
      const targetVer = versions.find(v => v['版ID'] === versionId);
      if (!targetVer) return { ok: false, message: '指定された版が見つかりません' };
      
      ssName += `_版${targetVer['版番号']}`;
      
      const vDetails = getTableData_(SHEET_NAMES.VERSION_DETAIL, ss).filter(d => d['版ID'] === versionId);
      headers = ['工程名', '班名', '作業順序', '作業名', '内訳順序', '作業内訳名', '区分', '作業方式', '単位', '正味時間(秒)', '人数', '工数(分)', '所要(分)'];
      
      vDetails.sort((a, b) => {
        if (a['作業順序'] !== b['作業順序']) return Number(a['作業順序']) - Number(b['作業順序']);
        return Number(a['内訳順序']) - Number(b['内訳順序']);
      });

      vDetails.forEach(d => {
        outputRows.push([
          d['工程名'],
          d['班名'],
          d['作業順序'],
          d['作業名'],
          d['内訳順序'],
          d['作業内訳名'],
          d['区分'],
          d['作業方式'],
          d['単位'],
          d['正味時間_現状'],
          d['人数_現状'],
          Math.round((Number(d['工数_現状']) / 60) * 10) / 10,
          Math.round((Number(d['所要_現状']) / 60) * 10) / 10
        ]);
      });
    }

    // 新しいスプレッドシートを作成
    const newSs = SpreadsheetApp.create(ssName);
    const sheet = newSs.getSheets()[0];
    sheet.setName('工数表');
    
    if (outputRows.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f0f0f0');
      sheet.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);
      
      // 罫線など簡単な見た目調整
      sheet.getDataRange().setBorder(true, true, true, true, true, true);
    } else {
      sheet.getRange(1, 1).setValue('対象データがありません');
    }
    
    // スプレッドシートのURLを返す
    return { ok: true, data: { url: newSs.getUrl() } };
    
  } catch(e) {
    return { ok: false, message: e.toString() };
  }
}
