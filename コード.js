// =========================================
// 00_定数.gs
// =========================================
const SHEET_NAMES = {
  PRODUCT: '製品',
  PARAM: '製品構成パラメータ',
  PROCESS: '工程',
  GROUP: '作業班',
  CHOICE: '選択肢',
  WORK: '作業',
  WORK_DETAIL: '作業内訳',
  VERSION: '版',
  VERSION_DETAIL: '版明細',
  HISTORY: '変更履歴',
  PLAN_PARAM: '計画パラメータ',
  RECORD: '観測記録'
};

const CONFIG = {
  DEFAULT_ALLOWANCE_RATE: 15, // 余裕率(%) 既定値（疲労・用達・職場余裕のみ）
  COEFFICIENT_DEFAULT: 1, // 係数の既定値
  ID_PREFIX_WORK: '-W-', // 作業IDの区切り
  ID_PREFIX_DETAIL: '-E-', // 内訳IDの区切り
  ORDER_INTERVAL: 10 // 順序の採番間隔
};

const HEADERS = {
  PRODUCT: ['製品ID', '製品コード', '型式名', '余裕率', '現行確定版番号', '状態', '備考', '作成日時', '作成者', '更新日時', '更新者'],
  PARAM: ['パラメータID', '製品ID', 'パラメータ名', '値', '単位', '状態', '作成日時', '作成者', '更新日時', '更新者'],
  PROCESS: ['工程ID', '工程名', '表示順', '状態', '作成日時', '作成者', '更新日時', '更新者'],
  GROUP: ['班ID', '班名', '担当職種', '担当会社', '表示順', '状態', '作成日時', '作成者', '更新日時', '更新者'],
  CHOICE: ['種別', '値ID', '表示名', '表示順', '状態', '作成日時', '作成者', '更新日時', '更新者'],
  WORK: ['作業ID', '製品ID', '工程ID', '班ID', '大工程', '作業名', '順序', '状態', '作成日時', '作成者', '更新日時', '更新者'],
  WORK_DETAIL: ['内訳ID', '作業ID', '順序', '作業内訳名', '区分', '単位', '時間方式', '正味時間_現状', '数量方式', '固定数量', '数量パラメータID', '係数', '人数_現状', '上限人数', '正味時間_目標', '人数_目標', '目標メモ', '先行作業', 'ムリムラムダ', '改善提案', '改善状態', 'issue番号', '根拠', '状態', '作成日時', '作成者', '更新日時', '更新者'],
  VERSION: ['版ID', '製品ID', '版番号', '確定日時', '確定者', '改訂理由', '委託先提示日', '提示先', '合意状況', '備考', '作成日時', '作成者', '更新日時', '更新者'],
  VERSION_DETAIL: ['版ID', '内訳ID', '工程名', '班名', '担当職種', '担当会社', '大工程', '作業名', '作業順序', '内訳順序', '作業内訳名', '区分', '単位', '時間方式', '正味時間_現状', '数量方式', '固定数量', '数量パラメータID', '係数', '人数_現状', '上限人数', '正味時間_目標', '人数_目標', '目標メモ', '先行作業', '数量_確定値', '工数_現状', '所要_現状', '工数_目標', '所要_目標'],
  HISTORY: ['履歴ID', '日時', '操作者', '製品ID', '対象種別', '対象ID', '操作', '項目', '旧値', '新値'],
  PLAN_PARAM: ['計画ID', '製品ID', '名称', '稼働日数', '定時稼働時間', '生産計画数', 'リードタイム', '作業人員', '日産台数', '便数', '備考', '作成日時', '作成者', '更新日時', '更新者'],
  RECORD: ['観測ID', '内訳ID', '観測日', '方法', '観測者', '観測値', 'サンプル番号', '採用', '備考', '作成日時', '作成者', '更新日時', '更新者']
};

// =========================================
// 01_初期化.gs
// =========================================
function initSpreadsheet() {
  const spreadSheetId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!spreadSheetId) {
    throw new Error('スクリプトプロパティに「SPREADSHEET_ID」が設定されていません。設定画面から追加してください。');
  }
  const ss = SpreadsheetApp.openById(spreadSheetId);
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  const user = Session.getActiveUser().getEmail();

  for (const key in SHEET_NAMES) {
    const sheetName = SHEET_NAMES[key];
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    sheet.clear();
    const headers = HEADERS[key];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  insertSampleData(ss, now, user);
}

function insertSampleData(ss, now, user) {
  // 製品
  ss.getSheetByName(SHEET_NAMES.PRODUCT).appendRow(['P-1', 'MP2500', 'MegaPower 2500', CONFIG.DEFAULT_ALLOWANCE_RATE, 0, '有効', 'サンプル', now, user, now, user]);

  // パラメータ
  ss.getSheetByName(SHEET_NAMES.PARAM).appendRow(['MP2500-K-1', 'P-1', 'MD数', 153, '個', '有効', now, user, now, user]);

  // 工程
  ss.getSheetByName(SHEET_NAMES.PROCESS).appendRow(['S-1', 'St.1 補器類-1', 10, '有効', now, user, now, user]);
  ss.getSheetByName(SHEET_NAMES.PROCESS).appendRow(['S-2', '出荷エリア', 20, '有効', now, user, now, user]);

  // 作業班
  ss.getSheetByName(SHEET_NAMES.GROUP).appendRow(['G-1', 'MD挿入班', '職種A', '会社A', 10, '有効', now, user, now, user]);

  // 選択肢
  ss.getSheetByName(SHEET_NAMES.CHOICE).appendRow(['大工程', 'C-1', 'MD挿入', 10, '有効', now, user, now, user]);

  // 作業
  ss.getSheetByName(SHEET_NAMES.WORK).appendRow(['MP2500-W-1', 'P-1', 'S-1', 'G-1', 'C-1', 'MD挿入作業', 10, '有効', now, user, now, user]);

  // 作業内訳 (テストケースを網羅するサンプル)
  const detailSheet = ss.getSheetByName(SHEET_NAMES.WORK_DETAIL);
  // 1. 分担・正味6秒・数量153・2人 (目標あり)
  detailSheet.appendRow(['MP2500-E-1', 'MP2500-W-1', 10, 'MD仮置き', '実作業', '台', '分担作業', 6.0, 'パラメータ', '', 'MP2500-K-1', 1, 2, 2, 5.0, 2, '改善', '', '', '', '未着手', '', '観測者A', '有効', now, user, now, user]);
  // 2. 同時・正味1200秒・数量1・2人
  detailSheet.appendRow(['MP2500-E-2', 'MP2500-W-1', 20, '位置決め', '実作業', '台', '同時作業', 1200.0, '固定', 1, '', '', 2, '', '', '', '', '', '', '', '', '', '観測者A', '有効', now, user, now, user]);
  // 3. 機械・自動（人数0）
  detailSheet.appendRow(['MP2500-E-3', 'MP2500-W-1', 30, '自動圧入', '実作業', '台', '同時作業', 9000.0, '固定', 1, '', '', 0, '', '', '', '', '', '', '', '', '', '観測者A', '有効', now, user, now, user]);
  // 4. 手待ち
  detailSheet.appendRow(['MP2500-E-4', 'MP2500-W-1', 40, '手待ち', '手待ち', '台', '同時作業', 17.0, 'パラメータ', '', 'MP2500-K-1', 1, 2, '', '', '', '', '', '', '', '', '', '観測者A', '有効', now, user, now, user]);
  // 5. 段取り (台あたり)
  detailSheet.appendRow(['MP2500-E-5', 'MP2500-W-1', 50, '治具準備', '段取り', '台', '同時作業', 300.0, '固定', 1, '', '', 1, '', '', '', '', '', '', '', '', '', '観測者A', '有効', now, user, now, user]);
  // 6. 間接作業 (日あたり)
  detailSheet.appendRow(['MP2500-E-6', 'MP2500-W-1', 60, '朝礼', '間接作業', '日', '同時作業', 600.0, '固定', 1, '', '', 10, '', '', '', '', '', '', '', '', '', '規定', '有効', now, user, now, user]);
}