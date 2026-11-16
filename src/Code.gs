/**
 * Googleフォームの回答が届いたら、スプレッドシートの「集計」シートを作り直し、担当者にメールで知らせる。
 * スプレッドシートに紐づいたスクリプト（拡張機能 → Apps Script で開くもの）として使う。
 * 動かすには、トリガー「フォーム送信時」に onFormSubmit を登録する（README の手順を参照）。
 */

// フォームの質問名やシート名を変えたときは、ここだけを直す
const CONFIG = {
  ITEM_COLUMN: '品目',
  QTY_COLUMN: '数量',
  SUMMARY_SHEET: '集計',
  // 通知先のメールアドレスは、コードではなくスクリプト プロパティに入れる（リポジトリに残さないため）
  RECIPIENT_PROPERTY: 'NOTIFY_TO',
};

/**
 * トリガー「フォーム送信時」から呼ばれる入口。
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e
 */
function onFormSubmit(e) {
  // 回答シートは、届いた回答の場所から取る（シート名は言語設定で変わるため決め打ちしない）
  const sheet = e && e.range ? e.range.getSheet() : null;
  if (!sheet) {
    throw new Error('フォームの回答の情報がありません。この関数はエディタから直接ではなく、トリガー「フォーム送信時」から動かしてください。');
  }

  const result = updateSummary_(sheet);

  // 今回の回答の品目と数量は、見出しの位置から取り出す
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = e.range.getValues()[0];
  const item = values[header.indexOf(CONFIG.ITEM_COLUMN)];
  const qty = values[header.indexOf(CONFIG.QTY_COLUMN)];

  notify_(item, qty, result);
}

/**
 * 回答シートを全部読み直して、「集計」シートに書き出す。
 * 毎回作り直すので、回答を手で消したり直したりしても集計がずれない。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} responseSheet
 * @return {ReturnType<typeof aggregateResponses>}
 */
function updateSummary_(responseSheet) {
  const all = responseSheet.getDataRange().getValues();
  const header = all[0];
  const rows = all.slice(1);
  const result = aggregateResponses(header, rows, CONFIG.ITEM_COLUMN, CONFIG.QTY_COLUMN);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName(CONFIG.SUMMARY_SHEET) || ss.insertSheet(CONFIG.SUMMARY_SHEET);
  const updatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  const table = buildSummaryTable(result, updatedAt);

  // 「最終更新」の日時をそのまま書くと、スプレッドシートが日付に変換してセルに日付の書式を付ける。
  // 品目が増えて行の位置がずれると、その書式が残ったセルの件数が「1900-01-02」のような日付で
  // 表示されてしまう（2026-11-16 に実機で確認。clear() だけでは直らなかった）。
  // そのため毎回、数字のセルは数値の書式、日時のセルは文字列の書式をはっきり指定してから書く。
  summary.clear();
  const range = summary.getRange(1, 1, table.length, table[0].length);
  range.setNumberFormat('0');
  summary.getRange(table.length, 2).setNumberFormat('@');
  range.setValues(table);
  return result;
}

/**
 * 集計結果をメールで知らせる。通知先が未設定なら送らずにログだけ残す（集計は終わっているので止めない）。
 * @param {*} item
 * @param {*} qty
 * @param {ReturnType<typeof aggregateResponses>} result
 */
function notify_(item, qty, result) {
  const to = PropertiesService.getScriptProperties().getProperty(CONFIG.RECIPIENT_PROPERTY);
  if (!to) {
    console.warn('通知先が未設定のため、メールを送りませんでした。スクリプト プロパティに ' + CONFIG.RECIPIENT_PROPERTY + ' を追加してください。');
    return;
  }
  const mail = buildNotification(item, qty, result);
  MailApp.sendEmail(to, mail.subject, mail.body);
}
