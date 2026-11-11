/**
 * excel-to-gas-transfer
 *
 * 公開用のサンプルであり、受注成果物ではありません。
 * 「入力」シート（Excel / CSV と同じ列）の行を「取込」へ追記します。
 * 日付＋商品名が同じ行は重複としてスキップし、件数を「実行メモ」に残します。
 *
 * VBA マクロの全機能移植ではありません。いま手でコピーしている1フローの例です。
 */

var CONFIG = Object.freeze({
  SPREADSHEET_ID: '',
  INPUT_SHEET: '入力',
  DEST_SHEET: '取込',
  LOG_SHEET: '実行メモ',
  INPUT_HEADERS: ['日付', '商品名（ダミー）', '数量', 'メモ'],
  DEST_HEADERS: ['日付', '商品名（ダミー）', '数量', 'メモ', '取込日時', '取込元'],
  LOG_HEADERS: ['実行日時', '追加', 'スキップ（重複）', 'エラー', '結果メモ'],
  KEY_DATE: '日付',
  KEY_NAME: '商品名（ダミー）',
  EMAIL_SUBJECT: '【サンプル】Excel転記の取り込みが終わりました'
});

/**
 * スプレッドシートを開いたときにメニューを追加する。
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('転記サンプル')
    .addItem('初期セットアップ', 'setupOnce')
    .addItem('入力シートから取り込む', 'importFromInputSheet')
    .addToUi();
}

/**
 * 初期セットアップ。「入力」「取込」「実行メモ」を用意し、ダミー行を入れる。
 * 既存の同名シートがある場合は、中身を上書きしない。
 */
function setupOnce() {
  var ss = getSpreadsheet_();
  var input = ensureSheet_(ss, CONFIG.INPUT_SHEET, CONFIG.INPUT_HEADERS);
  ensureSheet_(ss, CONFIG.DEST_SHEET, CONFIG.DEST_HEADERS);
  ensureSheet_(ss, CONFIG.LOG_SHEET, CONFIG.LOG_HEADERS);

  if (input.getLastRow() < 2) {
    input.getRange(2, 1, 4, CONFIG.INPUT_HEADERS.length).setValues([
      ['2026-11-01', 'サンプル紅茶（ダミー）', 3, '初回分'],
      ['2026-11-02', 'サンプルクッキー（ダミー）', 10, ''],
      ['2026-11-01', 'サンプル紅茶（ダミー）', 3, '同じキーの重複行（スキップされる）'],
      ['2026-11-03', 'サンプル緑茶（ダミー）', 2, '新規']
    ]);
  }

  Logger.log('setupOnce が完了しました。メニュー「転記サンプル」>「入力シートから取り込む」で確認できます。');
}

/**
 * 「入力」シートの行を「取込」へ追記する（このサンプルの標準操作）。
 */
function importFromInputSheet() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('同時実行のロックを取得できませんでした。しばらくしてから再実行してください。');
  }

  var added = 0;
  var skipped = 0;
  var errors = 0;
  var memo = '';

  try {
    var ss = getSpreadsheet_();
    var input = getSheetOrThrow_(ss, CONFIG.INPUT_SHEET);
    var dest = getSheetOrThrow_(ss, CONFIG.DEST_SHEET);
    var inputMap = getHeaderMap_(input, CONFIG.INPUT_HEADERS);
    getHeaderMap_(dest, CONFIG.DEST_HEADERS);

    var last = input.getLastRow();
    if (last < 2) {
      throw new Error('「入力」シートにデータ行がありません。先に setupOnce を実行するか、samples/input.csv と同じ列を貼り付けてください。');
    }

    var width = input.getLastColumn();
    var values = input.getRange(2, 1, last - 1, width).getDisplayValues();
    var existing = loadExistingKeys_(dest);
    var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
    var toAppend = [];
    var i;

    for (i = 0; i < values.length; i++) {
      var row = values[i];
      var dateVal = String(row[inputMap[CONFIG.KEY_DATE] - 1] || '').trim();
      var nameVal = String(row[inputMap[CONFIG.KEY_NAME] - 1] || '').trim();
      if (!dateVal && !nameVal) {
        continue;
      }
      if (!dateVal || !nameVal) {
        errors += 1;
        memo += '行' + (i + 2) + 'は日付または商品名が空のため取り込みません。 ';
        continue;
      }
      var key = makeKey_(dateVal, nameVal);
      if (existing[key]) {
        skipped += 1;
        continue;
      }
      toAppend.push([
        dateVal,
        nameVal,
        String(row[inputMap['数量'] - 1] || '').trim(),
        String(row[inputMap['メモ'] - 1] || '').trim(),
        now,
        '入力シート'
      ]);
      existing[key] = true;
      added += 1;
    }

    if (toAppend.length) {
      dest.getRange(dest.getLastRow() + 1, 1, toAppend.length, CONFIG.DEST_HEADERS.length).setValues(toAppend);
    }

    if (!memo) {
      memo = added + '件を追加し、重複' + skipped + '件をスキップしました。';
    }
    writeLog_(ss, added, skipped, errors, memo);
    maybeNotify_(added, skipped, errors, memo);
    SpreadsheetApp.getActive().toast(memo, '転記サンプル', 8);
    Logger.log(memo);
  } catch (err) {
    errors += 1;
    memo = String(err && err.message ? err.message : err);
    try {
      writeLog_(getSpreadsheet_(), added, skipped, errors, memo);
    } catch (logErr) {
      Logger.log('実行メモへの記入にも失敗しました: ' + logErr);
    }
    logError_('importFromInputSheet', err);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function makeKey_(dateVal, nameVal) {
  return dateVal + '\t' + nameVal;
}

function loadExistingKeys_(dest) {
  var map = getHeaderMap_(dest, CONFIG.DEST_HEADERS);
  var last = dest.getLastRow();
  var keys = {};
  if (last < 2) {
    return keys;
  }
  var dates = dest.getRange(2, map[CONFIG.KEY_DATE], last - 1, 1).getDisplayValues();
  var names = dest.getRange(2, map[CONFIG.KEY_NAME], last - 1, 1).getDisplayValues();
  var i;
  for (i = 0; i < dates.length; i++) {
    var d = String(dates[i][0] || '').trim();
    var n = String(names[i][0] || '').trim();
    if (d && n) {
      keys[makeKey_(d, n)] = true;
    }
  }
  return keys;
}

function writeLog_(ss, added, skipped, errors, memo) {
  var logSheet = getSheetOrThrow_(ss, CONFIG.LOG_SHEET);
  getHeaderMap_(logSheet, CONFIG.LOG_HEADERS);
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  logSheet.appendRow([now, added, skipped, errors, memo]);
}

function maybeNotify_(added, skipped, errors, memo) {
  var to = PropertiesService.getScriptProperties().getProperty('NOTIFY_TO');
  if (!to) {
    return;
  }
  var body = [
    'excel-to-gas-transfer のサンプル通知です（公開用サンプルであり、受注成果物ではありません）。',
    '',
    '追加: ' + added,
    'スキップ（重複）: ' + skipped,
    'エラー: ' + errors,
    'メモ: ' + memo
  ].join('\n');
  MailApp.sendEmail({
    to: to,
    subject: CONFIG.EMAIL_SUBJECT,
    body: body
  });
}

function getSpreadsheet_() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  var ss = SpreadsheetApp.getActive();
  if (!ss) {
    throw new Error('スプレッドシートにバインドして実行するか、CONFIG.SPREADSHEET_ID を設定してください。');
  }
  return ss;
}

function ensureSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getSheetOrThrow_(ss, name) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error('シート「' + name + '」が見つかりません。先に setupOnce を実行してください。');
  }
  return sheet;
}

function getHeaderMap_(sheet, required) {
  if (sheet.getLastRow() < 1) {
    throw new Error('シート「' + sheet.getName() + '」の1行目にヘッダーがありません。');
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var map = {};
  var i;
  for (i = 0; i < headers.length; i++) {
    var name = String(headers[i] || '').trim();
    if (name && map[name] == null) {
      map[name] = i + 1;
    }
  }
  for (i = 0; i < required.length; i++) {
    if (map[required[i]] == null) {
      throw new Error('シート「' + sheet.getName() + '」に列「' + required[i] + '」がありません。1行目の見出しを確認してください。');
    }
  }
  return map;
}

function logError_(where, err) {
  var message = err && err.stack ? err.stack : String(err);
  Logger.log('[' + where + '] ' + message);
}
