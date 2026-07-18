/**
 * 「売上データ」シートを月ごとに集計し、「月次サマリー」シートへ書き込んだ上で、
 * 月次推移を表す棒グラフを作成するメイン処理。
 */
function runDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName('売上データ');
  if (!dataSheet) {
    throw new Error('「売上データ」シートが見つかりません。');
  }

  var timeZone = ss.getSpreadsheetTimeZone();
  var lastRow = dataSheet.getLastRow();

  // ヘッダー行のみ、またはデータが無い場合は何もせず終了
  if (lastRow < 2) {
    Logger.log('集計対象データがありません。');
    return;
  }

  // A〜D列（日付・担当者名・商品名・金額）を一括取得
  var values = dataSheet.getRange(2, 1, lastRow - 1, 4).getValues();

  // 月ごとの集計結果を保持するマップ（キー：yyyy-MM、値：{label, total, count}）
  var summaryMap = {};

  values.forEach(function (row) {
    var rawDate = row[0];
    var amount = row[3];

    // 日付が空の行、または金額が数値でない行は集計対象から除外
    if (!rawDate || typeof amount !== 'number') {
      return;
    }

    var dateValue = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
    if (isNaN(dateValue.getTime())) {
      return;
    }

    var monthKey = Utilities.formatDate(dateValue, timeZone, 'yyyy-MM');
    var monthLabel = Utilities.formatDate(dateValue, timeZone, 'yyyy年M月');

    if (!summaryMap[monthKey]) {
      summaryMap[monthKey] = { label: monthLabel, total: 0, count: 0 };
    }
    summaryMap[monthKey].total += amount;
    summaryMap[monthKey].count += 1;
  });

  // 月の昇順に並び替えて出力用の配列を作成
  var sortedKeys = Object.keys(summaryMap).sort();
  var outputRows = sortedKeys.map(function (key) {
    var entry = summaryMap[key];
    return [entry.label, entry.total, entry.count];
  });

  // 出力先シートを取得（無ければ新規作成）し、内容とグラフを一旦クリア
  var summarySheet = ss.getSheetByName('月次サマリー');
  if (!summarySheet) {
    summarySheet = ss.insertSheet('月次サマリー');
  }
  summarySheet.clear();
  summarySheet.getCharts().forEach(function (chart) {
    summarySheet.removeChart(chart);
  });

  // ヘッダーと集計結果を書き込む
  summarySheet.getRange(1, 1, 1, 3).setValues([['月', '合計売上', '件数']]);
  if (outputRows.length > 0) {
    summarySheet.getRange(2, 1, outputRows.length, 3).setValues(outputRows);
  }

  // 月次推移を可視化する棒グラフを作成
  createMonthlyChart(summarySheet, outputRows.length);
}

/**
 * 月次サマリーの「月」列・「合計売上」列を対象に棒グラフを作成する。
 */
function createMonthlyChart(summarySheet, dataRowCount) {
  if (dataRowCount === 0) {
    return;
  }

  // ヘッダー行を含む「月」列（A列）と「合計売上」列（B列）をグラフ化
  var chartRange = summarySheet.getRange(1, 1, dataRowCount + 1, 2);

  var chart = summarySheet.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(chartRange)
    .setPosition(2, 5, 0, 0)
    .setOption('title', '月次売上推移')
    .setOption('legend', { position: 'none' })
    .setOption('hAxis', { title: '月' })
    .setOption('vAxis', { title: '合計売上' })
    .build();

  summarySheet.insertChart(chart);
}

/**
 * 毎朝9時に runDashboard を自動実行するトリガーを設定する。
 * 実行のたびに既存の runDashboard 用トリガーを削除してから登録するため、
 * 複数回実行してもトリガーが重複することはない。
 */
function setDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'runDashboard') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('runDashboard')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
}
