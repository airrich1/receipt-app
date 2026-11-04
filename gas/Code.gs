/**
 * パシャ領収書 — 受け取り側（Google スプレッドシート + Apps Script）
 *
 * 使い方（税理士さん or 自分が1回だけ）:
 *  1. Google スプレッドシートを新規作成
 *  2. メニュー「拡張機能」→「Apps Script」を開き、このファイルの中身を貼り付けて保存
 *  3. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *       次のユーザーとして実行: 自分 / アクセスできるユーザー: 全員
 *  4. 表示された「ウェブアプリURL（…/exec）」をアプリに登録
 *     例: https://<アプリの場所>/index.html?to=<ウェブアプリURL>  をお客さんに送るだけ
 *
 * 写真はスプレッドシートと同じ場所の「領収書写真」フォルダに保存され、
 * シート「領収書」に 1 行ずつ追記されます。
 */

const SHEET_NAME = '領収書';
const FOLDER_NAME = '領収書写真';
const HEADER = ['受信日時', '名前', '金額', 'メモ（音声）', '写真', '仕訳済'];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (!data.image || String(data.image).indexOf('data:image/') !== 0) {
      return json_({ ok: false, error: 'image required' });
    }

    const at = data.at ? new Date(data.at) : new Date();
    const stamp = Utilities.formatDate(at, 'Asia/Tokyo', 'yyyyMMdd_HHmmss');

    // 画像を Drive に保存
    const b64 = String(data.image).split(',')[1];
    const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg',
      stamp + '_' + (data.amount || '') + '.jpg');
    const file = getFolder_().createFile(blob);

    // シートに追記
    const sh = getSheet_();
    sh.appendRow([
      at,
      String(data.who || ''),
      data.amount ? Number(data.amount) : '',
      String(data.memo || ''),
      file.getUrl(),
      false
    ]);
    const row = sh.getLastRow();
    sh.getRange(row, 6).insertCheckboxes();

    return json_({ ok: true, row: row });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json_({ ok: true, message: 'パシャ領収書 受信口は動いています' });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(HEADER);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm');
    sh.getRange('C:C').setNumberFormat('#,##0"円"');
  }
  return sh;
}

function getFolder_() {
  const ssFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const parents = ssFile.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const it = parent.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : parent.createFolder(FOLDER_NAME);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
