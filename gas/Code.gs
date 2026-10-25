/**
 * 包装指図書 タブレット入力 — Google Apps Script(サーバー側)
 *
 * 使い方(初回だけ)
 *   1. 新しいGoogleスプレッドシートを作り、「拡張機能 → Apps Script」を開く
 *   2. このファイルの中身を Code.gs に丸ごと貼り付けて保存
 *   3. 関数「setup」を選んで実行(シートが自動で作られます)
 *   4. 「デプロイ → 新しいデプロイ → ウェブアプリ」
 *        次のユーザーとして実行: 自分 / アクセスできるユーザー: 全員
 *      → 表示されたURL(…/exec)を index.html の GAS_URL に貼る
 *
 * シート
 *   M_商品      … 品目マスタ(事務所が管理)
 *   M_担当      … 担当者の名前一覧
 *   オーダー登録 … 前日に事務所が入力(1行=1日×1品目)
 *   包装実績    … タブレットからの記録(1行=1記録。手で直さない)
 *   集計_日別   … 品目別・生地別の自動集計(見る用)
 */

const SHEETS = {
  item: 'M_商品',
  staff: 'M_担当',
  order: 'オーダー登録',
  result: '包装実績',
  summary: '集計_日別',
};

const HEAD = {
  item: ['商品ID', '包装ライン', '表示順', '商品コード', '品名', '箱入数', '消費期限日数', '集計先商品ID', '集計先生地ID', '状態'],
  staff: ['担当者名', '包装ライン', '状態'],
  order: ['日付', '商品ID', 'オーダー数', '品名(自動)', '包装ライン(自動)', '登録者', '備考'],
  result: ['記録ID', '日付', '年月', '包装ライン', '商品ID', '商品コード', '品名', '製造年月日', '消費期限',
           'オーダー数', '箱入数', '出来高', '出来上時刻', '余剰廃棄', '予備', '包装機カウント', '金検OKカウント',
           '担当', '集計先商品ID', '集計先生地ID', '入力日時', '端末', '状態', '取消した記録ID'],
};

const LINES = ['食パン', 'ロール', 'サンド', '菓子パン'];
const TZ = 'Asia/Tokyo';

/* ───────── 初期設定 ───────── */

function setup() {
  const ss = SpreadsheetApp.getActive();
  makeSheet_(ss, SHEETS.item, HEAD.item, [
    ['食パン-S001', '食パン', 10, 111, '3斤棒1', 5, 4, '食パン-S001', '食パン-001', '有効'],
    ['食パン-S002', '食パン', 20, 630, '3斤棒1.1スラ', 5, 4, '食パン-S002', '食パン-001', '有効'],
    ['ロール-S001', 'ロール', 10, 369, '超熟ロール6', 7, 4, 'ロール-S001', 'ロール-001', '有効'],
  ]);
  makeSheet_(ss, SHEETS.staff, HEAD.staff, [['担当A', '食パン', '有効'], ['担当B', 'ロール', '有効']]);
  const order = makeSheet_(ss, SHEETS.order, HEAD.order, []);
  makeSheet_(ss, SHEETS.result, HEAD.result, []);

  // オーダー登録: 商品IDはマスタから選ぶ。品名・ラインは自動表示
  const itemSheet = ss.getSheetByName(SHEETS.item);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(itemSheet.getRange('A2:A'), true).setAllowInvalid(false).build();
  order.getRange('B2:B').setDataValidation(rule);
  order.getRange('A2:A').setNumberFormat('yyyy/mm/dd');
  order.getRange('D2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,{M_商品!A2:A,M_商品!E2:E},2,FALSE),"マスタにない商品ID")))');
  order.getRange('E2').setFormula(
    '=ARRAYFORMULA(IF(B2:B="","",IFERROR(VLOOKUP(B2:B,{M_商品!A2:A,M_商品!B2:B},2,FALSE),"")))');

  // 集計_日別: B1 の日付を変えると切り替わる
  let sum = ss.getSheetByName(SHEETS.summary) || ss.insertSheet(SHEETS.summary);
  sum.clear();
  sum.getRange('A1').setValue('日付');
  sum.getRange('B1').setValue(new Date()).setNumberFormat('yyyy/mm/dd').setBackground('#fff2cc');
  sum.getRange('C1').setValue('← 日付を変えると切り替わります');
  sum.getRange('A3').setValue('品目別(品目別シートの出来高個数にあたる数)').setFontWeight('bold');
  sum.getRange('A4').setFormula(
    '=IFERROR(QUERY(包装実績!B:W,"select S, sum(L), sum(N) where B = date \'"&TEXT(B1,"yyyy-mm-dd")&"\' and W = \'有効\' group by S label S \'集計先商品ID\', sum(L) \'出来高\', sum(N) \'余剰廃棄\'",1),"記録なし")');
  sum.getRange('F3').setValue('生地別(部門シートの製品出来数・余剰生産ロスにあたる数)').setFontWeight('bold');
  sum.getRange('F4').setFormula(
    '=IFERROR(QUERY(包装実績!B:W,"select T, sum(L), sum(N) where B = date \'"&TEXT(B1,"yyyy-mm-dd")&"\' and W = \'有効\' group by T label T \'集計先生地ID\', sum(L) \'製品出来数\', sum(N) \'余剰生産ロス\'",1),"記録なし")');
  SpreadsheetApp.getUi().alert('シートを作成しました。M_商品 と M_担当 を本番の内容に書き換えてください。');
}

function makeSheet_(ss, name, head, rows) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
    if (rows.length) sh.getRange(2, 1, rows.length, head.length).setValues(rows);
    sh.setFrozenRows(1);
  }
  return sh;
}

/* ───────── タブレットから呼ばれる口 ───────── */

// 読み込み: ?action=load&date=2026-10-08&line=食パン
function doGet(e) {
  try {
    const p = e.parameter || {};
    if (p.action === 'load') return json_(load_(p.date, p.line));
    if (p.action === 'lines') return json_({ ok: true, lines: LINES });
    return json_({ ok: false, error: 'action が不明です' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

// 保存: body = {records:[{...}], device:"包装1"}
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const body = JSON.parse(e.postData.contents);
    return json_(save_(body));
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function load_(dateStr, line) {
  if (!dateStr || !line) throw new Error('日付とラインを指定してください');
  const ss = SpreadsheetApp.getActive();
  const items = rows_(ss.getSheetByName(SHEETS.item)).filter(r => r['状態'] !== '停止');
  const itemById = {};
  items.forEach(r => (itemById[r['商品ID']] = r));

  const orders = rows_(ss.getSheetByName(SHEETS.order))
    .filter(r => ymd_(r['日付']) === dateStr && itemById[r['商品ID']] && itemById[r['商品ID']]['包装ライン'] === line);

  // 同じ品目のオーダーが複数行あれば合計
  const qty = {};
  orders.forEach(r => (qty[r['商品ID']] = (qty[r['商品ID']] || 0) + Number(r['オーダー数'] || 0)));

  // その日すでに入力済みの値(有効な最新の記録)
  const done = {};
  rows_(ss.getSheetByName(SHEETS.result))
    .filter(r => ymd_(r['日付']) === dateStr && r['包装ライン'] === line && r['状態'] === '有効')
    .forEach(r => (done[r['商品ID']] = {
      id: r['記録ID'], out: num_(r['出来高']), time: r['出来上時刻'], waste: num_(r['余剰廃棄']),
      spare: num_(r['予備']), count: num_(r['包装機カウント']), kenpin: num_(r['金検OKカウント']), staff: r['担当'],
    }));

  const list = Object.keys(qty).map(id => {
    const it = itemById[id];
    return {
      itemId: id, code: it['商品コード'], name: it['品名'], order: qty[id],
      box: num_(it['箱入数']), sort: Number(it['表示順'] || 9999), saved: done[id] || null,
    };
  }).sort((a, b) => a.sort - b.sort);

  const staff = rows_(ss.getSheetByName(SHEETS.staff))
    .filter(r => r['状態'] !== '停止' && (!r['包装ライン'] || r['包装ライン'] === line))
    .map(r => r['担当者名']);

  return { ok: true, date: dateStr, line: line, items: list, staff: staff };
}

function save_(body) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEETS.result);
  const itemById = {};
  rows_(ss.getSheetByName(SHEETS.item)).forEach(r => (itemById[r['商品ID']] = r));

  // 二重送信よけ: 同じ送信IDはもう書かない
  const props = PropertiesService.getScriptProperties();
  const seenKey = 'sent_' + body.sendId;
  if (body.sendId && props.getProperty(seenKey)) return { ok: true, saved: 0, duplicate: true };

  const now = new Date();
  const out = [];
  let seq = Number(props.getProperty('seq') || 0);

  (body.records || []).forEach(rec => {
    const it = itemById[rec.itemId];
    if (!it) throw new Error('マスタにない商品IDです: ' + rec.itemId);
    if (rec.out === null || rec.out === undefined || rec.out === '') throw new Error(it['品名'] + ' の出来高が空です');

    // 修正の場合は前の記録を「取消」にする
    if (rec.replaceId) markCancelled_(sh, rec.replaceId);

    const d = parseYmd_(rec.date);
    const limitDays = Number(it['消費期限日数'] || 0);
    const limit = limitDays ? new Date(d.getTime() + limitDays * 86400000) : '';
    seq += 1;
    out.push([
      'H' + Utilities.formatDate(now, TZ, 'yyMMdd') + '-' + ('00000' + seq).slice(-5),
      d, Utilities.formatDate(d, TZ, 'yyyy-MM'), it['包装ライン'], rec.itemId, it['商品コード'], it['品名'],
      d, limit, num_(rec.order), num_(it['箱入数']), num_(rec.out), rec.time || '', num_(rec.waste), num_(rec.spare),
      num_(rec.count), num_(rec.kenpin), rec.staff || '', it['集計先商品ID'] || rec.itemId, it['集計先生地ID'] || '',
      now, body.device || '', '有効', rec.replaceId || '',
    ]);
  });

  if (out.length) {
    sh.getRange(sh.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
    sh.getRange(sh.getLastRow() - out.length + 1, 2, out.length, 1).setNumberFormat('yyyy/mm/dd');
  }
  props.setProperty('seq', String(seq));
  if (body.sendId) props.setProperty(seenKey, '1');
  return { ok: true, saved: out.length, ids: out.map(r => r[0]) };
}

function markCancelled_(sh, recordId) {
  const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  const col = HEAD.result.indexOf('状態') + 1;
  for (let i = ids.length - 1; i >= 0; i--) {
    if (ids[i][0] === recordId) { sh.getRange(i + 2, col).setValue('取消'); return; }
  }
}

/* ───────── 小物 ───────── */

function rows_(sh) {
  const v = sh.getDataRange().getValues();
  const head = v.shift();
  return v.filter(r => r.join('') !== '').map(r => {
    const o = {};
    head.forEach((h, i) => (o[h] = r[i]));
    return o;
  });
}
function ymd_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v || '').replace(/\//g, '-');
}
function parseYmd_(s) {
  const [y, m, d] = String(s).split('-').map(Number);
  return new Date(y, m - 1, d);
}
function num_(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(v);
  return isNaN(n) ? '' : n;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
