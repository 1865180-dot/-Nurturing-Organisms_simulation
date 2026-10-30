/**
 * 生物育成シミュレーター 提出記録用 Apps Script
 * 使い方：
 *  1. 記録用のGoogleスプレッドシートを新しく作る
 *  2. メニュー「拡張機能」→「Apps Script」を開き、このコードを貼り付けて保存
 *  3. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *     ・次のユーザーとして実行：自分
 *     ・アクセスできるユーザー：全員（学校ドメイン内なら「○○内の全員」でも可）
 *  4. 表示された「ウェブアプリのURL」を index.html の GAS_URL に貼り付ける
 * © Irie Kanji
 */
const SHEET_NAME = '記録';
const HEADERS = ['送信日時','種類','組','番号','氏名','題材','挑戦回数','ゲームオーバー回数',
  '設定（品種・方法など）','育成の目標','計画の理由','到達ターン','結果の概要',
  '評価1','評価2','評価3','評価4','合計（20点満点）',
  'ゲームオーバーの原因','ゲームオーバーの振り返り（30〜50字）',
  '振り返り1 工夫した管理と理由','振り返り2 立場ごとの評価','振り返り3 品種についての考え','振り返り4 次の改善'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#e6efe9');
    }
    const ax = d.axes || [];
    sh.appendRow([new Date(), d.kind, d.kumi, d.ban, d.name, d.sim, d.play, d.overCount,
      d.setting, d.goal, d.why, d.turn, d.summary,
      ax[0] || '', ax[1] || '', ax[2] || '', ax[3] || '', d.total || '',
      d.overCause, d.overReflect, d.q1 || '', d.q2 || '', d.q3 || '', d.q4 || '']);
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return ContentService.createTextOutput('生物育成シミュレーターの提出先は動作しています。');
}
