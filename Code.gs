// JAN棚卸 試作 v4（スプレッドシートの「拡張機能 > Apps Script」）
// スマホの画面は GitHub Pages に置き、このスクリプトは記録と集計だけを受け持つ。
// シート：商品マスタ / スキャン記録 / 集計_商品別 / 集計_場所別 / 進み具合 / 読み忘れ / 設定（合言葉。共有しない）
// v4：同じJANに複数の商品（熟成年・ヴィンテージ違い）をつなげられる。記録と集計は商品の No 単位。

const SHEET_MASTER = '商品マスタ';
const SHEET_LOG = 'スキャン記録';
const SHEET_BY_ITEM = '集計_商品別';
const SHEET_BY_LOC = '集計_場所別';
const SHEET_PROGRESS = '進み具合';
const SHEET_MISSING = '読み忘れ';
const SHEET_CONFIG = '設定';

// 商品マスタの列（A〜Cは集計の式が参照するので動かさない）
const MASTER_HEAD = ['JAN', '商品名', '仕入値', '容量', '種類', '楽一の商品名', '楽一コード', '楽一照合',
  '前回の場所', '前回の本数', 'No', '読み忘れから外す'];
const COL = { jan: 1, name: 2, cap: 4, kind: 5, rname: 6, mark: 8, area: 9, last: 10, no: 11, skip: 12 };
const AREAS = ['店内', '冷蔵庫', '小屋', '倉庫', 'その他'];

// 最初に1回（列を変えたときも）実行：シートと見出し・集計式・合言葉を作る（記録とマスタの中身はそのまま）
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const master = getOrCreate_(ss, SHEET_MASTER);
  master.getRange(1, 1, 1, MASTER_HEAD.length).setValues([MASTER_HEAD]);
  // 仕入値・本数・No 以外は文字として扱う（JANや楽一コードの先頭の0、容量が日付や数に化けないように）
  ['A:B', 'D:I', 'L:L'].forEach(function (r) { master.getRange(r).setNumberFormat('@'); });
  master.setFrozenRows(1);

  const log = getOrCreate_(ss, SHEET_LOG);
  if (log.getLastRow() === 0) {
    log.appendRow(['日時', '担当', '場所', 'JAN', '商品名', '数量', 'No']);
    log.getRange('D:D').setNumberFormat('@');
    log.getRange('A2:A').setNumberFormat('yyyy/mm/dd hh:mm:ss');
    log.setFrozenRows(1);
  }
  log.getRange('G1').setValue('No');
  fillLogNo_(log);

  const byItem = getOrCreate_(ss, SHEET_BY_ITEM);
  byItem.clear();
  // A:JAN B:数量 C:No（ここまで QUERY）  D:商品名 E:仕入値 F:仕入金額
  const M = "'" + SHEET_MASTER + "'!";
  byItem.getRange('A1').setFormula(
    "=QUERY('" + SHEET_LOG + "'!A:G,\"select D, sum(F), G where D <> '' group by D, G order by D label D 'JAN', sum(F) '数量', G 'No'\",1)");
  byItem.getRange('D1:F1').setValues([['商品名', '仕入値', '仕入金額']]);
  byItem.getRange('D2').setFormula(nameFormula_('A', 'C'));
  byItem.getRange('E2').setFormula(
    '=ARRAYFORMULA(IF(A2:A="","",IFERROR(1/(1/IF(C2:C="",VLOOKUP(A2:A,' + M + 'A:C,3,FALSE),' +
    'INDEX(' + M + 'C:C,MATCH(C2:C,' + M + 'K:K,0)))),"")))');
  byItem.getRange('F2').setFormula(
    '=ARRAYFORMULA(IF((A2:A="")+(E2:E=""),"",B2:B*E2:E))');
  byItem.getRange('H1:H2').setValues([['仕入金額 合計'], ['']]);
  byItem.getRange('H2').setFormula('=SUM(F2:F)');
  byItem.getRange('I1:I2').setValues([['仕入値 未入力の品目数'], ['']]);
  byItem.getRange('I2').setFormula('=COUNTIFS(A2:A,"<>",E2:E,"")');
  byItem.setFrozenRows(1);

  const byLoc = getOrCreate_(ss, SHEET_BY_LOC);
  byLoc.clear();
  // A:場所 B:JAN C:数量 D:No（ここまで QUERY）  E:商品名
  byLoc.getRange('A1').setFormula(
    "=QUERY('" + SHEET_LOG + "'!A:G,\"select C, D, sum(F), G where D <> '' group by C, D, G order by C label C '場所', D 'JAN', sum(F) '数量', G 'No'\",1)");
  byLoc.getRange('E1').setValue('商品名');
  byLoc.getRange('E2').setFormula(nameFormula_('B', 'D'));
  byLoc.setFrozenRows(1);

  // 棚ごとの進み具合（A列に棚を先に書いておけば「未着手」として並ぶ）
  const prog = getOrCreate_(ss, SHEET_PROGRESS);
  // （本数は「集計_場所別」で見る。ここに配列の式を置くと行の末尾が分からなくなるので置かない）
  prog.getRange('A1:D1').setValues([['場所・棚', '状態', '担当', '更新日時']]);
  prog.getRange('D2:D').setNumberFormat('yyyy/mm/dd hh:mm');
  prog.setFrozenRows(1);

  // 前回の棚卸にあったのに、今回まだどこでも読んでいない商品（JAN未紐づけを含む）
  const miss = getOrCreate_(ss, SHEET_MISSING);
  miss.clear();
  miss.getRange('A1:F1').setValues([['前回の場所', '種類', '商品名', '容量', '前回の本数', 'No']]);
  const m = "'" + SHEET_MASTER + "'!", lg = "'" + SHEET_LOG + "'!";
  // No で読んだか見る（No の無い古い記録は JAN で見る）
  miss.getRange('A2').setFormula(
    '=IFERROR(SORT(FILTER({' + m + 'I2:I,' + m + 'E2:E,' + m + 'B2:B,' + m + 'D2:D,' + m + 'J2:J,' + m + 'K2:K},' +
    m + 'J2:J>0,' + m + 'L2:L="",' +
    'COUNTIF(' + lg + 'G2:G,' + m + 'K2:K)=0,' +
    '(' + m + 'A2:A="")+(COUNTIFS(' + lg + 'D2:D,' + m + 'A2:A,' + lg + 'G2:G,"")=0)),1,TRUE,2,TRUE,3,TRUE),"（読み忘れはありません）")');
  miss.setFrozenRows(1);

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PASS')) props.setProperty('PASS', Utilities.getUuid().replace(/-/g, '').slice(0, 12));
  const conf = getOrCreate_(ss, SHEET_CONFIG);
  conf.getRange('A1:B1').setValues([['合言葉（スタッフ用の接続QRに入る。外に出さない）', props.getProperty('PASS')]]);
}

// 集計シートの商品名：No があれば No で、無ければ（古い記録）JAN で商品マスタを引く
function nameFormula_(janCol, noCol) {
  const M = "'" + SHEET_MASTER + "'!";
  const j = janCol + '2:' + janCol, n = noCol + '2:' + noCol;
  return '=ARRAYFORMULA(IF(' + j + '="","",IFERROR(IF(' + n + '="",' +
    'VLOOKUP(' + j + ',' + M + 'A:D,2,FALSE)&"　"&VLOOKUP(' + j + ',' + M + 'A:D,4,FALSE),' +
    'INDEX(' + M + 'B:B,MATCH(' + n + ',' + M + 'K:K,0))&"　"&INDEX(' + M + 'D:D,MATCH(' + n + ',' + M + 'K:K,0))),"（未登録）")))';
}

// No の空いている記録に、JAN から商品の No を入れる（そのJANの商品が1つだけのとき）
function fillLogNo_(log) {
  const last = log.getLastRow();
  if (last < 2) return;
  const one = noByJan_();
  const vals = log.getRange(2, 4, last - 1, 4).getValues();   // D〜G
  const out = vals.map(function (r) {
    if (String(r[3]).trim()) return [r[3]];
    const list = one[String(r[0]).trim()];
    return [list && list.length === 1 ? list[0] : ''];
  });
  log.getRange(2, 7, out.length, 1).setValues(out);
}

// JAN -> その JAN につながった商品の No の一覧
function noByJan_() {
  const map = {};
  masterRows_().rows.forEach(function (r) {
    const jan = String(r[COL.jan - 1]).trim(), no = Number(r[COL.no - 1]) || 0;
    if (jan && no) (map[jan] = map[jan] || []).push(no);
  });
  return map;
}

// 合言葉を作り直す（漏れたときに実行。古い接続QRは使えなくなる）
function resetPass() {
  PropertiesService.getScriptProperties().deleteProperty('PASS');
  setup();
}

// 画面とのやりとり（合言葉をURLに載せないよう、すべてPOST。本文はJSON。仕入値は返さない）
//   { action: 'master' }                          → { items: [[No, JAN, 商品名, 容量, 種類, 楽一の商品名, 前回の場所]], master: {JAN: 商品名} }
//   { action: 'lookup', jan }                      → { items: [そのJANの商品…], item: 先頭（無ければ null） }
//   { action: 'scans', rows: [{staff, loc, jan, no, name, qty}] }
//   { action: 'link', jan, no, extra }             リストの商品にJANを紐づける（extra=true で同じJANの別商品として足す）
//   { action: 'register', jan, name, cap, kind, extra }   リストに無い商品を新しく足す
//   { action: 'shelf', loc }                       → この棚で読んだ商品・担当・状態と、全棚の進み具合
//   { action: 'done', loc, staff, done }           棚を完了にする（done=false で作業中に戻す）
//   { action: 'missing', area }                    → 前回その場所にあって、まだ読んでいない商品
function doGet() {
  return json_({ ok: true, app: 'JAN棚卸' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '送信内容が読めません' }); }
  if (!checkPass_(body.pass)) return json_({ ok: false, error: '合言葉が違います' });
  try {
    if (body.action === 'master') return json_(Object.assign({ ok: true }, getMaster_()));
    if (body.action === 'lookup') { const found = findByJan_(body.jan); return json_({ ok: true, items: found, item: found[0] || null }); }
    if (body.action === 'scans') return json_({ ok: true, count: addScans_(body.rows || []) });
    if (body.action === 'link') return json_({ ok: true, item: linkItem_(body.jan, body.no, body.extra === true) });
    if (body.action === 'register') return json_({ ok: true, item: registerItem_(body.jan, body.name, body.cap, body.kind, body.extra === true) });
    if (body.action === 'shelf') return json_(Object.assign({ ok: true }, getShelf_(body.loc)));
    if (body.action === 'done') return json_({ ok: true, status: setDone_(body.loc, body.staff, body.done !== false) });
    if (body.action === 'missing') return json_({ ok: true, items: getMissing_(body.area) });
    return json_({ ok: false, error: '不明な操作です' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function masterRows_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MASTER);
  const last = sh.getLastRow();
  return { sh: sh, rows: last < 2 ? [] : sh.getRange(2, 1, last - 1, MASTER_HEAD.length).getValues() };
}

function itemOf_(r) {
  return [Number(r[COL.no - 1]) || 0, String(r[COL.jan - 1]).trim(), String(r[COL.name - 1]),
    String(r[COL.cap - 1]), String(r[COL.kind - 1]), String(r[COL.rname - 1]), String(r[COL.area - 1])];
}

// そのJANにつながった商品をすべて返す（ふつうは1つ。年違いがあれば複数）
function findByJan_(jan) {
  jan = String(jan || '').trim();
  if (!jan) return [];
  return masterRows_().rows.filter(function (r) { return String(r[COL.jan - 1]).trim() === jan; }).map(itemOf_);
}

function getMaster_() {
  const items = [], map = {};
  masterRows_().rows.forEach(function (r) {
    if (!String(r[COL.name - 1]).trim()) return;
    const it = itemOf_(r);
    items.push(it);
    if (it[1]) map[it[1]] = it[2];
  });
  return { items: items, master: map };
}

// 数量はマイナスも可（取り消し）
function addScans_(rows) {
  const values = [];
  rows.slice(0, 200).forEach(function (r) {
    const jan = String(r.jan || '').trim();
    const qty = Number(r.qty);
    if (!/^\d{8,14}$/.test(jan) || !qty || Math.abs(qty) > 9999) return;
    values.push([new Date(), cut_(r.staff), cut_(r.loc), "'" + jan, cut_(r.name), qty, Number(r.no) || '']);
  });
  if (!values.length) return 0;
  // No の無い送信（古い画面）は、そのJANの商品が1つだけなら補う
  if (values.some(function (v) { return !v[6]; })) {
    const one = noByJan_();
    values.forEach(function (v) { const l = one[String(v[3]).slice(1)]; if (!v[6] && l && l.length === 1) v[6] = l[0]; });
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
    sh.getRange(sh.getLastRow() + 1, 1, values.length, 7).setValues(values);
    // 読み始めた棚は「作業中」にする（完了済みはそのまま）
    const seen = {};
    values.forEach(function (v) { if (v[2]) (seen[v[2]] = seen[v[2]] || []).push(v[1]); });
    Object.keys(seen).forEach(function (loc) { touchProgress_(loc, seen[loc], '作業中', false); });
  } finally {
    lock.releaseLock();
  }
  return values.length;
}

// リストの商品（No）に、読んだJANを紐づける。
// extra=true のときは、すでに同じJANの商品があっても別の商品（年違い）として足す
function linkItem_(jan, no, extra) {
  jan = String(jan || '').trim();
  no = Number(no);
  if (!/^\d{8,14}$/.test(jan) || !no) throw new Error('JANか商品が選ばれていません');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const m = masterRows_();
    let at = -1;
    for (let i = 0; i < m.rows.length; i++) {
      const r = m.rows[i];
      if (Number(r[COL.no - 1]) === no) at = i;
      else if (!extra && String(r[COL.jan - 1]).trim() === jan) return itemOf_(r);   // ほかの人が先に紐づけ済み
    }
    if (at < 0) throw new Error('選んだ商品がリストに見つかりません');
    const cur = String(m.rows[at][COL.jan - 1]).trim();
    if (cur === jan) return itemOf_(m.rows[at]);
    if (cur) throw new Error('この商品には別のJAN（' + cur + '）が付いています。スプシの商品マスタで確認してください');
    m.sh.getRange(at + 2, COL.jan).setValue("'" + jan);
    m.rows[at][COL.jan - 1] = jan;
    return itemOf_(m.rows[at]);
  } finally {
    lock.releaseLock();
  }
}

// リストに無い商品を足す（仕入値は空欄のまま。あとでシートで入れる）
function registerItem_(jan, name, cap, kind, extra) {
  jan = String(jan || '').trim();
  name = cut_(name);
  if (!/^\d{8,14}$/.test(jan) || !name) throw new Error('JANか商品名が空です');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const m = masterRows_();
    let maxNo = 0;
    for (let i = 0; i < m.rows.length; i++) {
      const same = String(m.rows[i][COL.jan - 1]).trim() === jan;
      if (same && (!extra || String(m.rows[i][COL.name - 1]).trim() === name)) return itemOf_(m.rows[i]);
      maxNo = Math.max(maxNo, Number(m.rows[i][COL.no - 1]) || 0);
    }
    const row = MASTER_HEAD.map(function () { return ''; });
    row[COL.jan - 1] = "'" + jan;
    row[COL.name - 1] = name;
    row[COL.cap - 1] = cut_(cap);
    row[COL.kind - 1] = cut_(kind);
    row[COL.mark - 1] = 'スキャン時に追加';
    row[COL.no - 1] = maxNo + 1;
    m.sh.appendRow(row);
    row[COL.jan - 1] = jan;
    return itemOf_(row);
  } finally {
    lock.releaseLock();
  }
}

// この棚で読んだ商品（全員分）と、全棚の進み具合
function getShelf_(loc) {
  loc = String(loc || '').trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(SHEET_LOG);
  const sums = {}, staff = {};
  if (loc && log.getLastRow() >= 2) {
    log.getRange(2, 2, log.getLastRow() - 1, 6).getValues().forEach(function (r) {
      if (String(r[1]).trim() !== loc) return;
      const jan = String(r[2]).trim(), no = Number(r[5]) || 0;
      const k = no ? 'n' + no : 'j' + jan;
      sums[k] = sums[k] || { jan: jan, no: no, name: String(r[3]), qty: 0 };
      sums[k].qty += Number(r[4]) || 0;
      if (r[0]) staff[String(r[0])] = true;
    });
  }
  const items = Object.keys(sums).map(function (k) { return sums[k]; }).filter(function (s) { return s.qty !== 0; });
  const shelves = progressRows_().map(function (r) { return { loc: r[0], status: r[1], staff: r[2] }; });
  const mine = shelves.filter(function (s) { return s.loc === loc; })[0];
  return { items: items, staff: Object.keys(staff), status: mine ? mine.status : '', shelves: shelves };
}

function progressRows_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PROGRESS);
  const last = lastRowA_(sh);
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 3).getValues()
    .map(function (r) { return [String(r[0]).trim(), String(r[1]), String(r[2])]; })
    .filter(function (r) { return r[0]; });
}

// 進み具合の行を更新する。staff は名前か名前の配列。force=false なら「完了」は上書きしない
function touchProgress_(loc, staff, status, force) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_PROGRESS);
  const last = lastRowA_(sh);
  const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, 3).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim() !== loc) continue;
    if (!force && String(rows[i][1]) === '完了') return '完了';
    sh.getRange(i + 2, 2, 1, 3).setValues([[status, addName_(rows[i][2], staff), new Date()]]);
    return status;
  }
  sh.getRange(Math.max(last, 1) + 1, 1, 1, 4).setValues([[loc, status, addName_('', staff), new Date()]]);
  return status;
}

// A列に値がある最後の行（シートの書式や式に左右されない）
function lastRowA_(sh) {
  const n = sh.getLastRow();
  if (n < 1) return 0;
  const a = sh.getRange(1, 1, n, 1).getValues();
  for (let i = n - 1; i >= 0; i--) if (String(a[i][0]).trim()) return i + 1;
  return 0;
}

function addName_(names, staff) {
  const list = String(names || '').split('・').filter(String);
  [].concat(staff).forEach(function (s) { if (s && list.indexOf(s) < 0) list.push(s); });
  return list.join('・');
}

function setDone_(loc, staff, done) {
  loc = String(loc || '').trim();
  if (!loc) throw new Error('場所が空です');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return touchProgress_(cut_(loc), cut_(staff), done ? '完了' : '作業中', true);
  } finally {
    lock.releaseLock();
  }
}

// 前回 area にあった商品のうち、今回まだどこでも読んでいないもの
function getMissing_(area) {
  area = String(area || '').trim();
  if (AREAS.indexOf(area) < 0) throw new Error('場所の区分を選んでください');
  const log = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_LOG);
  const read = {};   // 'n'+No（No の無い古い記録は 'j'+JAN）
  if (log.getLastRow() >= 2) {
    log.getRange(2, 4, log.getLastRow() - 1, 4).getValues().forEach(function (r) {
      const no = Number(r[3]) || 0;
      read[no ? 'n' + no : 'j' + String(r[0]).trim()] = true;
    });
  }
  const out = [];
  masterRows_().rows.forEach(function (r) {
    if (!(Number(r[COL.last - 1]) > 0) || String(r[COL.skip - 1]).trim()) return;
    if (String(r[COL.area - 1]).split('・').indexOf(area) < 0) return;
    const jan = String(r[COL.jan - 1]).trim();
    if (read['n' + Number(r[COL.no - 1])] || (jan && read['j' + jan])) return;
    out.push([String(r[COL.kind - 1]), String(r[COL.name - 1]), String(r[COL.cap - 1]), Number(r[COL.last - 1])]);
  });
  out.sort(function (a, b) { return a[0] < b[0] ? -1 : a[0] > b[0] ? 1 : a[1] < b[1] ? -1 : 1; });
  return out;
}

function checkPass_(pass) {
  const real = PropertiesService.getScriptProperties().getProperty('PASS');
  return !!real && String(pass || '') === real;
}

// 先頭が = + - @ の文字はスプシで式にならないように ' をつける
function cut_(v) {
  const s = String(v == null ? '' : v).trim().slice(0, 100);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function getOrCreate_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
