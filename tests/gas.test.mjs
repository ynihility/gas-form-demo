// GAS のコードを Node.js で確かめるテスト（node --test で実行。依存ライブラリなし）
// SpreadsheetApp・MailApp などの Google のサービスは、メモリ上の偽物（モック）に置き換えている。
// ここで分かるのは「計算と処理の流れが正しいか」まで。Google 上での実際の動作は確かめられない。

import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import vm from 'node:vm';

const SRC = join(dirname(fileURLToPath(import.meta.url)), '..', 'src');
const HEADER = ['タイムスタンプ', '部署', '品目', '数量'];

// メモリ上で値を持つだけのシートの偽物
function fakeSheet(name, values) {
  const sheet = {
    name,
    values: values.map((r) => r.slice()),
    getDataRange: () => ({ getValues: () => sheet.values.map((r) => r.slice()) }),
    getLastColumn: () => (sheet.values[0] || []).length,
    getRange: (row, col, numRows, numCols) => ({
      getValues: () => sheet.values.slice(row - 1, row - 1 + numRows).map((r) => r.slice(col - 1, col - 1 + numCols)),
      setValues: (v) => { v.forEach((r, i) => { sheet.values[row - 1 + i] = r.slice(); }); },
      getSheet: () => sheet,
    }),
    clearContents: () => { sheet.values = []; },
  };
  return sheet;
}

// src の .gs を、偽物のサービスを入れた環境で読み込む
function loadGas({ responses, recipient }) {
  const responseSheet = fakeSheet('フォームの回答 1', [HEADER, ...responses]);
  const sheets = { 'フォームの回答 1': responseSheet };
  const sent = [];
  const warnings = [];
  const context = vm.createContext({
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: (n) => sheets[n] || null,
        insertSheet: (n) => (sheets[n] = fakeSheet(n, [])),
      }),
    },
    MailApp: { sendEmail: (to, subject, body) => sent.push({ to, subject, body }) },
    PropertiesService: {
      getScriptProperties: () => ({ getProperty: (k) => (k === 'NOTIFY_TO' ? recipient : null) }),
    },
    Utilities: { formatDate: () => '2026-11-16 12:00' },
    Session: { getScriptTimeZone: () => 'Asia/Tokyo' },
    console: { warn: (m) => warnings.push(m), log: () => {} },
  });
  for (const f of ['Aggregate.gs', 'Code.gs']) {
    vm.runInContext(readFileSync(join(SRC, f), 'utf8'), context, { filename: f });
  }
  // 最後の回答行を「今届いた回答」として、トリガーの引数 e を作る
  const lastRow = responseSheet.values.length;
  const e = { range: responseSheet.getRange(lastRow, 1, 1, HEADER.length) };
  return { context, sheets, sent, warnings, e };
}

test('品目ごとの件数と数量を集計する（正常系）', () => {
  const { context } = loadGas({ responses: [] });
  const r = context.aggregateResponses(HEADER, [
    ['t1', '総務', 'ボールペン', 10],
    ['t2', '営業', 'コピー用紙', 2],
    ['t3', '総務', 'ボールペン', '5'],
  ], '品目', '数量');
  assert.equal(r.totalCount, 3);
  assert.equal(r.totalQty, 17);
  assert.equal(r.invalidQty, 0);
  assert.deepEqual(JSON.parse(JSON.stringify(r.items)), [
    { item: 'ボールペン', count: 2, qty: 15 },
    { item: 'コピー用紙', count: 1, qty: 2 },
  ]);
});

test('読めない数量と空行の扱い（異常系）', () => {
  const { context } = loadGas({ responses: [] });
  const r = context.aggregateResponses(HEADER, [
    ['t1', '総務', 'ボールペン', 'たくさん'],
    ['t2', '総務', 'ボールペン', -3],
    ['t3', '総務', 'ボールペン', 1.5],
    ['', '', '', ''],
  ], '品目', '数量');
  assert.equal(r.totalCount, 3);
  assert.equal(r.totalQty, 0);
  assert.equal(r.invalidQty, 3);
});

test('見出しの列が見つからないときは、どの列かを伝えて止める（異常系）', () => {
  const { context } = loadGas({ responses: [] });
  assert.throws(() => context.aggregateResponses(['タイムスタンプ', '品名', '数量'], [], '品目', '数量'), /品目/);
});

test('フォーム送信時：集計シートを作り、通知メールを1通送る', () => {
  const g = loadGas({
    recipient: 'demo@example.com',
    responses: [
      ['t1', '総務', 'ボールペン', 10],
      ['t2', '営業', 'コピー用紙', 2],
    ],
  });
  g.context.onFormSubmit(g.e);
  // vm の中で作られた配列は別の環境のものなので、JSON を通して普通の配列にそろえてから比べる
  const summary = JSON.parse(JSON.stringify(g.sheets['集計'].values));
  assert.deepEqual(summary[0], ['品目', '件数', '数量の合計']);
  assert.deepEqual(summary[1], ['ボールペン', 1, 10]);
  assert.deepEqual(summary[3], ['合計', 2, 12]);
  assert.equal(g.sent.length, 1);
  assert.equal(g.sent[0].to, 'demo@example.com');
  assert.match(g.sent[0].subject, /コピー用紙 ×2/);
  // 個人情報になりうる「部署」の値は本文に入れない
  assert.doesNotMatch(g.sent[0].body, /営業|総務/);
});

test('通知先が未設定なら、集計だけ行ってメールは送らない（異常系）', () => {
  const g = loadGas({ recipient: null, responses: [['t1', '総務', 'ボールペン', 1]] });
  g.context.onFormSubmit(g.e);
  assert.ok(g.sheets['集計']);
  assert.equal(g.sent.length, 0);
  assert.equal(g.warnings.length, 1);
});

test('エディタから直接実行した（e がない）ときは、分かる言葉で止める（異常系）', () => {
  const g = loadGas({ recipient: 'demo@example.com', responses: [] });
  assert.throws(() => g.context.onFormSubmit(undefined), /トリガー/);
});
