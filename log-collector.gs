/**
 * 言葉の星座 ― 読みの記録を集める Google Apps Script
 *
 * 使い方（先生用・約5分）
 *  1. Google スプレッドシートを新しく作る（名前は「言葉の星座 記録」など）。
 *  2. 拡張機能 → Apps Script を開き、このファイルの中身をすべて貼りつける。
 *  3. 下の KEY を、先生だけが知る合言葉に書きかえて保存する。
 *  4. デプロイ → 新しいデプロイ → 種類「ウェブアプリ」
 *       次のユーザーとして実行：自分 ／ アクセスできるユーザー：全員
 *     → 表示された「ウェブアプリのURL」（…/exec）をコピーする。
 *  5. 生徒に配るアプリのリンクの末尾に  &log=（そのURL）  を付ける。
 *     例： https://kokugo-lab.github.io/yume-seiza/?feed=…&cls=1&log=https://script.google.com/macros/s/…/exec
 *  6. 先生の端末では、取りこみ画面の「読みの記録（先生用）」で URL と合言葉を入れると、クラス全体を集計できる。
 *
 * 記録は「log」シートに一行ずつたまる（時刻・セッション・ID・組・番号・名前・画面・場所・問・種類・内容）。
 */
const KEY = 'ここに合言葉';
const SHEET = 'log';
const HEAD = ['時刻','セッション','順','ID','組','番号','名前','画面','場所','問','種類','内容'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const evs = body.events || [];
    if (!evs.length) return ContentService.createTextOutput('ok');
    const sh = sheet_();
    const rows = evs.map(v => [v.ts, v.sid, v.seq, v.pid, v.cls, v.no, v.nick, v.mode, v.place, v.q, v.type,
                               JSON.stringify(v.d || {}).slice(0, 45000)]);
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEAD.length).setNumberFormat('@').setValues(rows);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  if (!e.parameter.all) return json_({ ok: true });
  if (String(e.parameter.key || '') !== KEY) return json_({ error: '合言葉がちがいます' });
  const sh = sheet_(), n = sh.getLastRow();
  if (n < 2) return json_({ events: [] });
  const vals = sh.getRange(2, 1, n - 1, HEAD.length).getValues();
  const events = vals.map(r => ({
    ts: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), sid: String(r[1]), seq: r[2], pid: String(r[3]),
    cls: String(r[4]), no: String(r[5]), nick: String(r[6]), mode: String(r[7]), place: r[8], q: String(r[9]),
    type: String(r[10]), d: String(r[11])
  }));
  return json_({ events: events });
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  return sh;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
