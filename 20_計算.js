// =========================================
// 20_計算.gs
// =========================================

/**
 * 1行（作業内訳）の計算を行う（C01〜C06, C18対応）
 * @param {Object} row - 作業内訳のデータオブジェクト（キーは列名）
 * @param {number} allowanceRate - 余裕率（％、既定15）
 * @param {Object} parameters - { "パラメータID": 値 } のマップ
 * @returns {Object} 計算結果を含むオブジェクト
 */
function calculateDetail(row, allowanceRate, parameters) {
  const calcSingle = (isTarget) => {
    const netTimePrefix = isTarget ? '正味時間_目標' : '正味時間_現状';
    const personPrefix = isTarget ? '人数_目標' : '人数_現状';

    // C06: 目標未入力の場合は現状と同値として扱う
    let netTime = row[netTimePrefix];
    if (isTarget && (netTime === '' || netTime === null || netTime === undefined)) {
      netTime = row['正味時間_現状'];
    }
    netTime = Number(netTime);

    let persons = row[personPrefix];
    if (isTarget && (persons === '' || persons === null || persons === undefined)) {
      persons = row['人数_現状'];
    }
    persons = Number(persons);

    // C06: 目標正味時間＝0 の行は廃止扱い
    if (isTarget && netTime === 0) {
      return { durationSec: 0, workSec: 0, stdDurationSec: 0, stdWorkSec: 0 };
    }

    // C01: 数量計算
    let qty = 1;
    if (row['数量方式'] === '固定') {
      qty = Number(row['固定数量'] || 1);
    } else if (row['数量方式'] === 'パラメータ') {
      const paramValue = Number(parameters[row['数量パラメータID']] || 0);
      const coef = Number(row['係数'] || CONFIG.COEFFICIENT_DEFAULT);
      qty = paramValue * coef;
    }

    let durationSec = 0;
    let workSec = 0;

    // C02, C03, C04: 分担作業と同時作業（手待ちは同時作業扱い）の計算
    if (row['時間方式'] === '分担作業') {
      if (row['上限人数'] !== '' && row['上限人数'] !== null && row['上限人数'] !== undefined) {
        persons = Math.min(persons, Number(row['上限人数']));
      }
      if (persons < 1) persons = 1; // 0割防止
      workSec = netTime * qty;
      durationSec = workSec / persons;
    } else { // 同時作業・手待ち
      durationSec = netTime * qty;
      workSec = durationSec * persons; // 人数0(機械等)なら工数0
    }

    // C05: 標準時間（余裕率加算）
    const rate = 1 + (Number(allowanceRate) / 100);
    const stdDurationSec = durationSec * rate;
    const stdWorkSec = workSec * rate;

    return { durationSec, workSec, stdDurationSec, stdWorkSec };
  };

  const current = calcSingle(false);
  const target = calcSingle(true);

  // C18: 台あたり以外は別枠集計するためのフラグ
  const unit = row['単位'] || '台';
  const isPerUnit = (unit === '台');

  // C38: 作業方式と人数の整合検証
  const warnings = [];
  const method = row['作業方式'] || '';
  const currentPersons = Number(row['人数_現状'] || 0);
  if (['機械自動（人は離れられる）', '自動搬送'].includes(method)) {
    if (currentPersons > 0) warnings.push('【警告】自動化方式では人数は0である必要があります。');
  } else if (method) {
    if (currentPersons < 1) warnings.push('【警告】手作業または操作・監視を伴う方式では人数は1以上である必要があります。');
  }
  // C19: 数量方式「部品の使用数」は使用部品シート（5.13）が未実装のため、数量1として計算している
  if (row['数量方式'] === '部品の使用数') {
    warnings.push('【警告】数量方式「部品の使用数」は未対応です。数量1として計算しています。');
  }
  const warningMessage = warnings.join(' ');

  return {
    unit: unit,
    isPerUnit: isPerUnit, // true なら「台あたり集計」の対象
    current: current,
    target: target,
    warning: warningMessage,
    // C07: 改善効果 (現状工数 - 目標工数)
    effectWorkSec: current.workSec - target.workSec,
    effectStdWorkSec: current.stdWorkSec - target.stdWorkSec
  };
}

/**
 * C14: 表示用の丸め処理（秒を「分」または「人・分」に変換して小数第1桁）
 */
function formatMin(sec) {
  return Math.round((sec / 60) * 10) / 10;
}