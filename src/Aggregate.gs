/**
 * 集計の計算だけをまとめたファイル。
 * スプレッドシートやメールには触らないので、Node.js のテストでもそのまま動かせる。
 */

/**
 * 回答の行を品目ごとにまとめ、件数と数量の合計を出す。
 * @param {string[]} header 1行目（見出し）
 * @param {Array<Array<*>>} rows 2行目以降（回答）
 * @param {string} itemColumn 品目の列の見出し
 * @param {string} qtyColumn 数量の列の見出し
 * @return {{items: Array<{item: string, count: number, qty: number}>, totalCount: number, totalQty: number, invalidQty: number}}
 */
function aggregateResponses(header, rows, itemColumn, qtyColumn) {
  const itemIndex = header.indexOf(itemColumn);
  const qtyIndex = header.indexOf(qtyColumn);
  // 見出しの名前が変わっていると集計できないので、どの列が見つからないかをはっきり伝える
  if (itemIndex === -1 || qtyIndex === -1) {
    throw new Error('見出しに「' + itemColumn + '」または「' + qtyColumn + '」の列が見つかりません。フォームの質問名と CONFIG を確認してください。');
  }

  const byItem = {};
  const order = [];
  let totalQty = 0;
  let invalidQty = 0;

  rows.forEach(function (row) {
    const item = String(row[itemIndex] === undefined || row[itemIndex] === null ? '' : row[itemIndex]).trim();
    // 空行（品目が空）は数えない
    if (item === '') return;

    if (!byItem[item]) {
      byItem[item] = { item: item, count: 0, qty: 0 };
      order.push(item);
    }
    byItem[item].count += 1;

    // 数量は0以上の整数だけ足す。読めない値は件数にだけ入れ、別に数えておく
    const qty = Number(row[qtyIndex]);
    if (Number.isInteger(qty) && qty >= 0) {
      byItem[item].qty += qty;
      totalQty += qty;
    } else {
      invalidQty += 1;
    }
  });

  const items = order.map(function (key) { return byItem[key]; });
  const totalCount = items.reduce(function (sum, x) { return sum + x.count; }, 0);
  return { items: items, totalCount: totalCount, totalQty: totalQty, invalidQty: invalidQty };
}

/**
 * 集計結果を、「集計」シートにそのまま書ける2次元配列にする。
 * @param {ReturnType<typeof aggregateResponses>} result
 * @param {string} updatedAt 更新日時の文字列
 * @return {Array<Array<*>>}
 */
function buildSummaryTable(result, updatedAt) {
  const table = [['品目', '件数', '数量の合計']];
  result.items.forEach(function (x) {
    table.push([x.item, x.count, x.qty]);
  });
  table.push(['合計', result.totalCount, result.totalQty]);
  table.push(['数量が読めなかった件数', result.invalidQty, '']);
  table.push(['最終更新', updatedAt, '']);
  return table;
}

/**
 * 通知メールの件名と本文を作る。
 * 個人情報が入る列は本文に入れず、品目と数量と集計だけを載せる。
 * @param {string} item 今回の回答の品目
 * @param {*} qty 今回の回答の数量
 * @param {ReturnType<typeof aggregateResponses>} result
 * @return {{subject: string, body: string}}
 */
function buildNotification(item, qty, result) {
  const lines = [
    '新しい申請がありました。',
    '',
    '品目：' + item,
    '数量：' + qty,
    '',
    '現在の集計（品目：件数／数量の合計）',
  ];
  result.items.forEach(function (x) {
    lines.push('・' + x.item + '：' + x.count + '件／' + x.qty);
  });
  lines.push('合計：' + result.totalCount + '件／' + result.totalQty);
  lines.push('');
  lines.push('詳しくは、スプレッドシートの「集計」シートを確認してください。');
  return {
    subject: '[備品申請] ' + item + ' ×' + qty + '（累計' + result.totalCount + '件）',
    body: lines.join('\n'),
  };
}
