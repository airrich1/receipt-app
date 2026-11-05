/**
 * パシャ領収書 — 受け取り側（Google スプレッドシート + Apps Script）
 *
 * 使い方（税理士さん or 自分が1回だけ）:
 *  1. Google スプレッドシートを新規作成
 *  2. メニュー「拡張機能」→「Apps Script」を開き、このファイルの中身を貼り付けて保存
 *  3. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *       次のユーザーとして実行: 自分 / アクセスできるユーザー: 全員
 *  4. 表示された「ウェブアプリURL（…/exec）」をアプリに登録
 *
 * コードを更新したとき: 貼り替えて保存 →「デプロイ」→「デプロイを管理」→ ✏️ →
 *   バージョン「新バージョン」→「デプロイ」（URLは変わりません）
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

    // 画像を Drive に保存（保存先フォルダのIDは覚えておき、毎回探さない）
    const b64 = String(data.image).split(',')[1];
    const blob = Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', stamp + '_' + String(data.id || '').slice(-6) + '.jpg');
    const file = getFolder_().createFile(blob);

    // シートに追記（仕訳済の列は、シート作成時にまとめてチェックボックスにしてある）
    getSheet_().appendRow([
      at,
      String(data.who || ''),
      data.amount ? Number(data.amount) : '',
      String(data.memo || ''),
      file.getUrl(),
      false
    ]);

    return json_({ ok: true });
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
  // 「仕訳済」列を一度だけまとめてチェックボックスにする（1件ごとの設定をやめて速くする）
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('checkboxColumn') !== '1') {
    const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sh.getRange(2, 6, sh.getMaxRows() - 1, 1).setDataValidation(rule);
    props.setProperty('checkboxColumn', '1');
  }
  return sh;
}

function getFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('folderId');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* 消された場合は作り直す */ }
  }
  const ssFile = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  const parents = ssFile.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const it = parent.getFoldersByName(FOLDER_NAME);
  const folder = it.hasNext() ? it.next() : parent.createFolder(FOLDER_NAME);
  props.setProperty('folderId', folder.getId());
  return folder;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
