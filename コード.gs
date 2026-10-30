/**
 * 生物育成シミュレーター 提出記録用 Apps Script
 * 使い方：
 *  1. 記録用のGoogleスプレッドシートを新しく作る
 *  2. メニュー「拡張機能」→「Apps Script」を開き、このコードを貼り付けて保存
 *  3. 「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」
 *     ・次のユーザーとして実行：自分
 *     ・アクセスできるユーザー：（学校のドメイン）内の全員 でも 全員 でも可
 *  4. 表示された「ウェブアプリのURL」を index.html の GAS_URL に貼り付ける
 *  ※ コードを書き換えたら「デプロイを管理」→ 編集 →「新バージョン」で更新（URLは変わらない）
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
  let count = 0;
  try {
    // アプリからはフォーム送信（payload）で届く。JSONの直接送信にも対応。
    const raw = (e.parameter && e.parameter.payload) ? e.parameter.payload : e.postData.contents;
    const data = JSON.parse(raw);
    const list = Array.isArray(data) ? data : [data];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#e6efe9');
    }
    list.forEach(d => {
      const ax = d.axes || [];
      sh.appendRow([new Date(), d.kind, d.kumi, d.ban, d.name, d.sim, d.play, d.overCount,
        d.setting, d.goal, d.why, d.turn, d.summary,
        ax[0] || '', ax[1] || '', ax[2] || '', ax[3] || '', d.total || '',
        d.overCause, d.overReflect, d.q1 || '', d.q2 || '', d.q3 || '', d.q4 || '']);
      count++;
    });
  } catch (err) {
    return page('⚠️ 提出できませんでした', 'アプリにもどって「もう一度送る」を押してください。<br><small>' + err + '</small>');
  } finally {
    lock.releaseLock();
  }
  return page('✅ 提出できました', (count > 1 ? count + '件の記録を受け付けました。<br>' : '') + 'このタブを閉じて、アプリにもどってください。');
}

function doGet() {
  return page('生物育成シミュレーター', '提出先は動作しています。');
}

function page(title, msg) {
  return HtmlService.createHtmlOutput(
    '<div style="font-family:sans-serif;text-align:center;padding:48px 16px">' +
    '<h1 style="font-size:28px">' + title + '</h1><p style="font-size:18px;line-height:1.8">' + msg + '</p></div>'
  ).setTitle(title).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
