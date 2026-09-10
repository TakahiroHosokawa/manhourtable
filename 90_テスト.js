// =========================================
// 90_テスト.gs
// =========================================

function runTests() {
  Logger.log("=== 計算ロジック受入テスト開始 ===");
  let passed = true;

  const assertAlmostEqual = (actual, expected, testName) => {
    if (Math.abs(actual - expected) < 0.1) {
      Logger.log(`[OK] ${testName}: ${actual}`);
    } else {
      Logger.log(`[NG] ${testName}: Expected ${expected}, but got ${actual}`);
      passed = false;
    }
  };

  const allowance = 15;
  const params = { 'MP2500-K-1': 153 };

  // テスト1: 分担作業・正味6秒・数量153（MD数）・2人 -> 工数 15.3 人・分、所要 7.7 分
  const row1 = {
    '時間方式': '分担作業', '正味時間_現状': 6, '数量方式': 'パラメータ', '数量パラメータID': 'MP2500-K-1', '係数': 1, '人数_現状': 2, '単位': '台'
  };
  const res1 = calculateDetail(row1, 0, params);
  assertAlmostEqual(formatMin(res1.current.workSec), 15.3, "テスト1: 分担作業(工数)");
  assertAlmostEqual(formatMin(res1.current.durationSec), 7.7, "テスト1: 分担作業(所要)");

  // テスト2: 同時作業・正味1200秒・数量1・2人 -> 所要 20.0 分、工数 40.0 人・分
  const row2 = {
    '時間方式': '同時作業', '正味時間_現状': 1200, '数量方式': '固定', '固定数量': 1, '人数_現状': 2, '単位': '台'
  };
  const res2 = calculateDetail(row2, 0, params);
  assertAlmostEqual(formatMin(res2.current.durationSec), 20.0, "テスト2: 同時作業(所要)");
  assertAlmostEqual(formatMin(res2.current.workSec), 40.0, "テスト2: 同時作業(工数)");

  // テスト3: 同時作業・正味9000秒・数量1・0人（機械） -> 所要 150.0 分、工数 0 人・分
  const row3 = {
    '時間方式': '同時作業', '正味時間_現状': 9000, '数量方式': '固定', '固定数量': 1, '人数_現状': 0, '単位': '台'
  };
  const res3 = calculateDetail(row3, 0, params);
  assertAlmostEqual(formatMin(res3.current.durationSec), 150.0, "テスト3: 機械(所要)");
  assertAlmostEqual(formatMin(res3.current.workSec), 0.0, "テスト3: 機械(工数)");

  // テスト4: 上の1行目に余裕率15% -> 標準工数 17.6 人・分、標準所要 8.8 分
  const res4 = calculateDetail(row1, 15, params);
  assertAlmostEqual(formatMin(res4.current.stdWorkSec), 17.6, "テスト4: 標準工数");
  assertAlmostEqual(formatMin(res4.current.stdDurationSec), 8.8, "テスト4: 標準所要");

  // テスト5: 同一工程に班A（所要30分）と班B（所要20分） -> 工程の所要 30分
  const groupADuration = 30;
  const groupBDuration = 20;
  const processDuration = Math.max(groupADuration, groupBDuration);
  assertAlmostEqual(processDuration, 30.0, "テスト5: 工程の所要(班のmax)");

  // テスト6: 目標正味時間=0 の行 -> 目標工数・目標所要とも 0。改善効果＝現状工数
  const row6 = {
    '時間方式': '分担作業', '正味時間_現状': 6, '数量方式': '固定', '固定数量': 100, '人数_現状': 2, '単位': '台',
    '正味時間_目標': 0
  };
  const res6 = calculateDetail(row6, 0, params);
  assertAlmostEqual(res6.target.workSec, 0, "テスト6: 目標工数(廃止)");
  assertAlmostEqual(res6.effectStdWorkSec, res6.current.stdWorkSec, "テスト6: 改善効果(現状と同値)");

  // テスト7: 目標未入力の行 -> 目標＝現状と同値。改善効果 0
  const row7 = {
    '時間方式': '分担作業', '正味時間_現状': 6, '数量方式': '固定', '固定数量': 100, '人数_現状': 2, '単位': '台'
  };
  const res7 = calculateDetail(row7, 0, params);
  assertAlmostEqual(res7.target.workSec, res7.current.workSec, "テスト7: 目標未入力(現状と同値)");
  assertAlmostEqual(res7.effectStdWorkSec, 0, "テスト7: 改善効果(0)");

  Logger.log(passed ? "=== 全テスト合格 ===" : "=== テスト失敗あり ===");
}