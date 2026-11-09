/**
 * 鮮魚の予約注文（洋食だいにんぐなかむら）— Google Apps Script
 *
 * はじめにやること（くわしくは「設定手順.md」）
 *  1) 下の SPREADSHEET_ID を自分のスプレッドシートのIDに書き換える
 *  2) 関数「setup」を実行する（シートが2枚と、写真を入れるドライブのフォルダができる）
 *  3) スクリプト プロパティに ADMIN_PASS / LINE_TOKEN / LINE_TO を登録する
 *  4) ウェブアプリとしてデプロイ（アクセスできるユーザー：全員）
 *
 * パスワードやLINEのトークンは、このコードには書かずに
 * 「スクリプト プロパティ」に保存します。
 *
 * 魚は1尾ずつ量りに乗せて写真を撮り、重さを入れて登録する（1行＝1尾）。
 * 写真はドライブの「鮮魚の写真」フォルダに保存し、リンクを知っている人が見られる設定にする。
 */

const SPREADSHEET_ID = '14TeXUIa9000q1OuRk7LlyHNWq7uaqnm7Uv8azjqbI8c';
const TZ = 'Asia/Tokyo';

const ITEM_HEAD = ['id', '魚名', '100g単価(税込)', '重さ(g)', '入荷数', '残り', '1人上限',
  '三枚おろし手間賃(税込)', '受け取り日', '受け取り時間', '締切', '産地', 'メモ', '状態', '登録日時', '写真'];
const ORDER_HEAD = ['受付番号', '受付日時', '入荷id', '魚名', '数', '渡し方', 'お名前', '電話番号',
  '受け取り日', '受け取り時間', '金額(税込)', 'ご要望', '状態', '重さ(g)'];

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
// 1行を文字のまま足す（appendRow だと締切が日付に、電話番号が数字に変わってしまうため）
function appendText_(sh, row) {
  sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
}
// スプレッドシートの時間帯を東京にそろえる（clasp で作ると東京以外になっていて、日時が9時間ずれた）
function fixTimeZone_(ss) {
  if (ss.getSpreadsheetTimeZone() !== TZ) ss.setSpreadsheetTimeZone(TZ);
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

/* ---------- 最初に実行（もう一度実行しても、入っているデータはそのまま） ---------- */
function setup() {
  const ss = ss_();
  fixTimeZone_(ss);
  [['入荷', ITEM_HEAD], ['注文', ORDER_HEAD]].forEach(function (pair) {
    const sh = ss.getSheetByName(pair[0]) || ss.insertSheet(pair[0]);
    sh.getRange('A:Z').setNumberFormat('@'); // 電話番号の先頭0や日付が勝手に変わらないよう文字として保存
    sh.getRange(1, 1, 1, pair[1].length).setValues([pair[1]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  Logger.log('準備できました。写真の保存先：' + photoFolder_().getUrl());
}

/* ---------- 入口（写真つきの登録だけPOST、ほかはすべてGET） ---------- */
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

// 写真はURLに入らない大きさなので、登録はPOSTで受ける
function doPost(e) {
  let d = {};
  let res;
  try {
    d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    res = d.op === 'addFish' ? admin_(d) : { ok: false, error: 'unknown op' };
  } catch (err) {
    res = { ok: false, error: String(err) };
  }
  // no-cors で送られるので返事はブラウザに届かない。失敗の理由は hasItem で読めるよう10分だけ残す
  const id = d.item && d.item.id;
  if (!res.ok && id) CacheService.getScriptCache().put('post_' + String(id).slice(0, 20), String(res.error), 600);
  return out_(res);
}

/* ---------- お客さん向け ---------- */
function publicItem_(it) {
  return {
    id: String(it.id),
    name: String(it['魚名']),
    price: Number(it['100g単価(税込)']) || 0,
    weight: Number(it['重さ(g)']) || 0,
    total: Number(it['入荷数']) || 0,
    remain: Number(it['残り']) || 0,
    limit: Number(it['1人上限']) || 1,
    fee: Number(it['三枚おろし手間賃(税込)']) || 0,
    pickupDate: str_(it['受け取り日'], 'yyyy-MM-dd'),
    pickupTime: String(it['受け取り時間']),
    deadline: str_(it['締切'], "yyyy-MM-dd'T'HH:mm"),
    origin: String(it['産地']),
    memo: String(it['メモ']),
    state: String(it['状態']),
    photo: str_(it['写真'])
  };
}

function openItems_() {
  return rows_(sheet_('入荷'))
    .filter(function (it) { return it['状態'] === '受付中' && !deadlinePassed_(it); })
    .map(publicItem_);
}

function amount_(it, qty, style) {
  const one = Math.round((Number(it['重さ(g)']) || 0) / 100 * (Number(it['100g単価(税込)']) || 0));
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
  try {
    const sh = sheet_('入荷');
    const all = rows_(sh);
    const it = all.filter(function (r) { return String(r.id) === String(d.id); })[0];
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
          error: remain > 0 ? '残りが' + remain + '尾になりました。数を選び直してください' : '先に注文が入り、売り切れました'
        };
      } else {
        setCell_(sh, it._row, ITEM_HEAD, '残り', remain - qty);
        const osh = sheet_('注文');
        const today = Utilities.formatDate(new Date(), TZ, 'MMdd');
        const count = rows_(osh).filter(function (o) {
          return String(o['受付番号']).indexOf('F' + today) === 0;
        }).length;
        const no = 'F' + today + '-' + ('00' + (count + 1)).slice(-3);
        const amount = amount_(it, qty, style);
        const weight = Number(it['重さ(g)']) || 0;
        const pDate = str_(it['受け取り日'], 'yyyy-MM-dd');
        const pTime = String(it['受け取り時間']);
        appendText_(osh, [no, now_(), String(it.id), String(it['魚名']), String(qty), style, name, tel,
          pDate, pTime, String(amount), clean_(d.note, 200), '受付', String(weight)]);
        SpreadsheetApp.flush();
        // 同じ魚・同じ受け取り日で、まだ注文できる尾数（LINE通知用）
        const left = all.filter(function (r) {
          return String(r['魚名']) === String(it['魚名']) && str_(r['受け取り日'], 'yyyy-MM-dd') === pDate &&
            r['状態'] === '受付中' && !deadlinePassed_(r);
        }).reduce(function (n, r) { return n + (Number(r['残り']) || 0); }, 0) - qty;
        res = {
          ok: true, no: no, name: String(it['魚名']), weight: weight, qty: qty, style: style, amount: amount,
          pickupDate: pDate, pickupTime: pTime, left: left, note: clean_(d.note, 200)
        };
      }
    }
  } finally {
    lock.releaseLock();
  }

  if (res.ok) {
    notify_([
      '【鮮魚の注文】' + res.no,
      res.name + (res.weight ? ' ' + res.weight.toLocaleString() + 'g' : '') + ' ×' + res.qty + '（' + res.style + '）',
      name + ' さん　' + tel,
      '受け取り ' + res.pickupDate + ' ' + res.pickupTime,
      '税込' + res.amount.toLocaleString() + '円',
      res.name + 'の残り ' + res.left + '尾',
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
    fixTimeZone_(ss_()); // 管理画面を開いたときに一度そろえる（店主がsetupをやり直さなくて済むように）
    const itemRows = rows_(sheet_('入荷'));
    const photos = {};
    itemRows.forEach(function (it) { photos[String(it.id)] = str_(it['写真']); });
    const items = itemRows.reverse().slice(0, 60).map(function (it) {
      const o = publicItem_(it);
      o.expired = deadlinePassed_(it);
      return o;
    });
    const orders = rows_(sheet_('注文')).reverse().slice(0, 150).map(function (o) {
      return {
        no: String(o['受付番号']), at: str_(o['受付日時'], 'yyyy-MM-dd HH:mm'), itemId: String(o['入荷id']),
        name: String(o['魚名']), weight: Number(o['重さ(g)']) || 0, qty: Number(o['数']) || 0,
        style: String(o['渡し方']), customer: String(o['お名前']), tel: String(o['電話番号']),
        pickupDate: str_(o['受け取り日'], 'yyyy-MM-dd'), pickupTime: String(o['受け取り時間']),
        amount: Number(o['金額(税込)']) || 0, note: String(o['ご要望']), state: String(o['状態']),
        photo: photos[String(o['入荷id'])] || ''
      };
    });
    return { ok: true, items: items, orders: orders };
  }

  if (d.op === 'addFish') {
    const f = d.item || {};
    const id = String(f.id || '');
    const name = clean_(f.name, 30);
    const price = Math.round(Number(f.price));
    const weight = Math.round(Number(f.weight));
    if (!/^[0-9a-f]{8}$/.test(id)) return { ok: false, error: '登録番号が正しくありません' };
    if (!name) return { ok: false, error: '魚の名前を入れてください' };
    if (!(price > 0)) return { ok: false, error: '100gあたりの値段を入れてください' };
    if (!(weight > 0)) return { ok: false, error: '量りの重さを入れてください' };
    if (!f.deadline) return { ok: false, error: '締切を入れてください' };
    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      const sh = sheet_('入荷');
      // 電波が悪くて同じ1尾を2回送ったときは、2回目を無視する
      if (rows_(sh).some(function (r) { return String(r.id) === id; })) return { ok: true, id: id };
      const photo = savePhoto_(d.photo, id, name);
      appendText_(sh, [id, name, String(price), String(weight), '1', '1', '1',
        String(Math.round(Number(f.fee)) || 0), clean_(f.pickupDate, 10), clean_(f.pickupTime, 30),
        clean_(f.deadline, 16), clean_(f.origin, 40), clean_(f.memo, 200), '受付中', now_(), photo]);
      SpreadsheetApp.flush();
      return { ok: true, id: id };
    } finally {
      lock.releaseLock();
    }
  }

  // POSTの返事は読めないので、登録できたかをここで確かめる
  if (d.op === 'hasItem') {
    const found = rows_(sheet_('入荷')).some(function (r) { return String(r.id) === String(d.id); });
    const error = found ? '' : (CacheService.getScriptCache().get('post_' + String(d.id).slice(0, 20)) || '');
    return { ok: true, found: found, error: error };
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

/* ---------- 写真 ---------- */
// data:image/jpeg;base64,… を「鮮魚の写真」フォルダに保存して、ファイルIDを返す
function savePhoto_(dataUrl, id, name) {
  if (!dataUrl) return '';
  const m = String(dataUrl).match(/^data:(image\/(?:jpeg|png));base64,([A-Za-z0-9+\/=]+)$/);
  if (!m) throw new Error('写真の形式が違います');
  if (m[2].length > 8000000) throw new Error('写真が大きすぎます');
  const ext = m[1] === 'image/png' ? '.png' : '.jpg';
  const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], id + '_' + name + ext);
  const file = photoFolder_().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // 注文フォームに出すため
  return file.getId();
}

function photoFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('PHOTO_FOLDER');
  if (id) {
    try {
      return DriveApp.getFolderById(id);
    } catch (e) {
      // フォルダを消してしまったときは作り直す
    }
  }
  const folder = DriveApp.createFolder('鮮魚の写真');
  props.setProperty('PHOTO_FOLDER', folder.getId());
  return folder;
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
