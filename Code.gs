/**
 * 鮮魚の予約注文（洋食だいにんぐなかむら）— Google Apps Script
 *
 * はじめにやること（くわしくは「設定手順.md」）
 *  1) 下の SPREADSHEET_ID を自分のスプレッドシートのIDに書き換える
 *  2) 関数「setup」を1回だけ実行する（シートが2枚できる）
 *  3) スクリプト プロパティに ADMIN_PASS / LINE_TOKEN / LINE_TO を登録する
 *  4) ウェブアプリとしてデプロイ（アクセスできるユーザー：全員）
 *
 * パスワードやLINEのトークンは、このコードには書かずに
 * 「スクリプト プロパティ」に保存します。
 */

const SPREADSHEET_ID = '14TeXUIa9000q1OuRk7LlyHNWq7uaqnm7Uv8azjqbI8c';
const TZ = 'Asia/Tokyo';

const ITEM_HEAD = ['id', '魚名', '100g単価(税込)', '目安重量(g)', '入荷数', '残り', '1人上限',
  '三枚おろし手間賃(税込)', '受け取り日', '受け取り時間', '締切', '産地', 'メモ', '状態', '登録日時'];
const ORDER_HEAD = ['受付番号', '受付日時', '入荷id', '魚名', '数', '渡し方', 'お名前', '電話番号',
  '受け取り日', '受け取り時間', '目安金額(税込)', 'ご要望', '状態'];

/* ---------- 共通の道具 ---------- */
function ss_() { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function sheet_(name) { return ss_().getSheetByName(name); }
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function str_(v, fmt) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, fmt);
  return v === null || v === undefined ? '' : String(v);
}
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }
// 先頭の = + - @ を外す（シートで式として動かないように）
function clean_(s, max) {
  return String(s || '').trim().replace(/^[=+\-@]+/, '').slice(0, max || 100);
}
function rows_(sh) {
  const v = sh.getDataRange().getValues();
  const head = v.shift() || [];
  return v.map(function (r, i) {
    const o = { _row: i + 2 };
    head.forEach(function (h, j) { o[h] = r[j]; });
    return o;
  });
}
function setCell_(sh, row, head, col, val) {
  sh.getRange(row, head.indexOf(col) + 1).setValue(String(val));
}
function deadlinePassed_(it) {
  const s = str_(it['締切'], "yyyy-MM-dd'T'HH:mm").replace(' ', 'T').slice(0, 16);
  if (!s) return false;
  try {
    return new Date() > Utilities.parseDate(s, TZ, "yyyy-MM-dd'T'HH:mm");
  } catch (e) {
    return false;
  }
}

/* ---------- 最初に1回だけ実行 ---------- */
function setup() {
  const ss = ss_();
  [['入荷', ITEM_HEAD], ['注文', ORDER_HEAD]].forEach(function (pair) {
    const sh = ss.getSheetByName(pair[0]) || ss.insertSheet(pair[0]);
    sh.getRange('A:Z').setNumberFormat('@'); // 電話番号の先頭0や日付が勝手に変わらないよう文字として保存
    sh.getRange(1, 1, 1, pair[1].length).setValues([pair[1]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
}

/* ---------- 入口（すべてGETで受ける） ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  try {
    const data = p.data ? JSON.parse(p.data) : {};
    if (p.action === 'list') return out_({ ok: true, items: openItems_() });
    if (p.action === 'order') return out_(placeOrder_(data));
    if (p.action === 'admin') return out_(admin_(data));
    return out_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

/* ---------- お客さん向け ---------- */
function publicItem_(it) {
  return {
    id: String(it.id),
    name: String(it['魚名']),
    price: Number(it['100g単価(税込)']) || 0,
    weight: Number(it['目安重量(g)']) || 0,
    total: Number(it['入荷数']) || 0,
    remain: Number(it['残り']) || 0,
    limit: Number(it['1人上限']) || 1,
    fee: Number(it['三枚おろし手間賃(税込)']) || 0,
    pickupDate: str_(it['受け取り日'], 'yyyy-MM-dd'),
    pickupTime: String(it['受け取り時間']),
    deadline: str_(it['締切'], "yyyy-MM-dd'T'HH:mm"),
    origin: String(it['産地']),
    memo: String(it['メモ']),
    state: String(it['状態'])
  };
}

function openItems_() {
  return rows_(sheet_('入荷'))
    .filter(function (it) { return it['状態'] === '受付中' && !deadlinePassed_(it); })
    .map(publicItem_);
}

function estimate_(it, qty, style) {
  const one = Math.round((Number(it['目安重量(g)']) || 0) / 100 * (Number(it['100g単価(税込)']) || 0));
  const fee = style === '三枚おろし' ? (Number(it['三枚おろし手間賃(税込)']) || 0) : 0;
  return (one + fee) * qty;
}

function placeOrder_(d) {
  const qty = parseInt(d.qty, 10);
  const name = clean_(d.name, 40);
  const tel = String(d.tel || '').replace(/[^\d]/g, '');
  const style = d.style === '三枚おろし' ? '三枚おろし' : '丸のまま';
  if (!d.id || !qty || qty < 1) return { ok: false, error: '数を選んでください' };
  if (!name) return { ok: false, error: 'お名前を入れてください' };
  if (tel.length < 10 || tel.length > 11) return { ok: false, error: '電話番号を確認してください' };

  const lock = LockService.getScriptLock();
  lock.waitLock(15000); // 同時に注文が来ても1件ずつ処理する
  let res;
  let itemCopy;
  try {
    const sh = sheet_('入荷');
    const it = rows_(sh).filter(function (r) { return String(r.id) === String(d.id); })[0];
    if (!it || it['状態'] !== '受付中' || deadlinePassed_(it)) {
      res = { ok: false, code: 'closed', error: 'この魚の受付は終了しました' };
    } else {
      const remain = Number(it['残り']) || 0;
      const limit = Number(it['1人上限']) || 1;
      if (qty > limit) {
        res = { ok: false, error: 'お一人様' + limit + '尾までです' };
      } else if (qty > remain) {
        res = {
          ok: false, code: 'soldout', remain: remain,
          error: remain > 0 ? '残りが' + remain + '尾になりました。数を選び直してください' : '売り切れました'
        };
      } else {
        setCell_(sh, it._row, ITEM_HEAD, '残り', remain - qty);
        const osh = sheet_('注文');
        const today = Utilities.formatDate(new Date(), TZ, 'MMdd');
        const count = rows_(osh).filter(function (o) {
          return String(o['受付番号']).indexOf('F' + today) === 0;
        }).length;
        const no = 'F' + today + '-' + ('00' + (count + 1)).slice(-3);
        const est = estimate_(it, qty, style);
        const pDate = str_(it['受け取り日'], 'yyyy-MM-dd');
        const pTime = String(it['受け取り時間']);
        osh.appendRow([no, now_(), String(it.id), String(it['魚名']), String(qty), style, name, tel,
          pDate, pTime, String(est), clean_(d.note, 200), '受付']);
        SpreadsheetApp.flush();
        itemCopy = it;
        res = {
          ok: true, no: no, name: String(it['魚名']), qty: qty, style: style, estimate: est,
          pickupDate: pDate, pickupTime: pTime, remainAfter: remain - qty, note: clean_(d.note, 200)
        };
      }
    }
  } finally {
    lock.releaseLock();
  }

  if (res.ok) {
    notify_([
      '【鮮魚の注文】' + res.no,
      res.name + ' ×' + res.qty + '（' + res.style + '）',
      name + ' さん　' + tel,
      '受け取り ' + res.pickupDate + ' ' + res.pickupTime,
      '目安 税込' + res.estimate.toLocaleString() + '円',
      '残り ' + res.remainAfter + '尾',
      res.note ? 'ご要望：' + res.note : ''
    ].filter(String).join('\n'));
  }
  return res;
}

/* ---------- 管理画面向け ---------- */
function admin_(d) {
  if (!prop_('ADMIN_PASS') || d.pass !== prop_('ADMIN_PASS')) {
    return { ok: false, code: 'auth', error: 'パスワードが違います' };
  }

  if (d.op === 'list') {
    const items = rows_(sheet_('入荷')).reverse().slice(0, 30).map(function (it) {
      const o = publicItem_(it);
      o.expired = deadlinePassed_(it);
      return o;
    });
    const orders = rows_(sheet_('注文')).reverse().slice(0, 150).map(function (o) {
      return {
        no: String(o['受付番号']), at: str_(o['受付日時'], 'yyyy-MM-dd HH:mm'), itemId: String(o['入荷id']),
        name: String(o['魚名']), qty: Number(o['数']) || 0, style: String(o['渡し方']),
        customer: String(o['お名前']), tel: String(o['電話番号']),
        pickupDate: str_(o['受け取り日'], 'yyyy-MM-dd'), pickupTime: String(o['受け取り時間']),
        estimate: Number(o['目安金額(税込)']) || 0, note: String(o['ご要望']), state: String(o['状態'])
      };
    });
    return { ok: true, items: items, orders: orders };
  }

  if (d.op === 'addItem') {
    const f = d.item || {};
    const name = clean_(f.name, 30);
    const price = Math.round(Number(f.price));
    const total = parseInt(f.total, 10);
    if (!name) return { ok: false, error: '魚の名前を入れてください' };
    if (!(price > 0)) return { ok: false, error: '100gあたりの値段を入れてください' };
    if (!(total > 0)) return { ok: false, error: '入荷数を入れてください' };
    if (!f.deadline) return { ok: false, error: '締切を入れてください' };
    const id = Utilities.getUuid().slice(0, 8);
    sheet_('入荷').appendRow([id, name, String(price), String(Math.round(Number(f.weight)) || 0),
      String(total), String(total), String(parseInt(f.limit, 10) || 1),
      String(Math.round(Number(f.fee)) || 0), clean_(f.pickupDate, 10), clean_(f.pickupTime, 30),
      clean_(f.deadline, 16), clean_(f.origin, 40), clean_(f.memo, 200), '受付中', now_()]);
    return { ok: true, id: id };
  }

  if (d.op === 'closeItem') {
    const sh = sheet_('入荷');
    const it = rows_(sh).filter(function (r) { return String(r.id) === String(d.id); })[0];
    if (!it) return { ok: false, error: '見つかりません' };
    setCell_(sh, it._row, ITEM_HEAD, '状態', '締切');
    return { ok: true };
  }

  if (d.op === 'orderState') {
    const next = d.state;
    if (['受け取り済み', '取消', '受付'].indexOf(next) < 0) return { ok: false, error: '状態が不正です' };
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const osh = sheet_('注文');
      const o = rows_(osh).filter(function (r) { return String(r['受付番号']) === String(d.no); })[0];
      if (!o) return { ok: false, error: '注文が見つかりません' };
      const prev = String(o['状態']);
      if (prev === '取消') return { ok: false, error: '取り消した注文は戻せません' };
      setCell_(osh, o._row, ORDER_HEAD, '状態', next);
      if (next === '取消') { // 取り消した分を残りに戻す
        const sh = sheet_('入荷');
        const it = rows_(sh).filter(function (r) { return String(r.id) === String(o['入荷id']); })[0];
        if (it) setCell_(sh, it._row, ITEM_HEAD, '残り', (Number(it['残り']) || 0) + (Number(o['数']) || 0));
      }
      SpreadsheetApp.flush();
      return { ok: true };
    } finally {
      lock.releaseLock();
    }
  }

  if (d.op === 'testLine') {
    return notify_('鮮魚の注文フォーム：テスト通知です。これが届けばLINEの設定は完了です。');
  }

  return { ok: false, error: 'unknown op' };
}

/* ---------- LINE通知 ---------- */
function notify_(text) {
  const token = prop_('LINE_TOKEN');
  const to = prop_('LINE_TO');
  if (!token || !to) return { ok: false, error: 'LINE_TOKEN か LINE_TO が未設定です' };
  let err = '';
  to.split(',').map(function (s) { return s.trim(); }).filter(String).forEach(function (id) {
    const r = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ to: id, messages: [{ type: 'text', text: text }] }),
      muteHttpExceptions: true
    });
    if (r.getResponseCode() !== 200) err = 'LINEエラー ' + r.getResponseCode() + '：' + r.getContentText();
  });
  return { ok: !err, error: err };
}
