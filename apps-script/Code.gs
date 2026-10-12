/**
 * 町のエアコンお掃除専門店＋α 予約フォーム用 Google Apps Script（受信専用）
 *
 * フォーム本体は GitHub Pages に置き、送信内容だけをこのスクリプトで受け取ります。
 *  - 予約内容をスプレッドシート「予約一覧」に記録
 *  - 店舗（NOTIFY_EMAIL）へ通知メール、お客様へ受付確認メールを送信
 *
 * 設置手順：
 *  1. NOTIFY_EMAIL を通知先に変更（複数はカンマ区切り）
 *  2. setup を一度実行して権限を承認（記録用スプレッドシートも作成されます）
 *  3. 「デプロイ」→「新しいデプロイ」→ ウェブアプリ（実行：自分 ／ アクセス：全員）
 *  4. 発行された /exec のURLを、フォーム（index.html）の ENDPOINT_URL に設定
 *  ※ コード変更後は「デプロイを管理 → 新しいバージョン」で再デプロイ（URLは変わりません）
 */

const NOTIFY_EMAIL = 'ここに通知先メールアドレスを入力';   // 複数はカンマ区切り
const SHEET_NAME   = '予約一覧';
const SHOP_NAME    = '町のエアコンお掃除専門店＋α';

// 初回に一度だけ実行：権限の承認と、記録用スプレッドシートの作成
function setup() {
  const ss = getSpreadsheet_();
  Logger.log('記録用スプレッドシート: ' + ss.getUrl());
  Logger.log('本日のメール送信可能数: ' + MailApp.getRemainingDailyQuota());
}

// 動作確認用（ブラウザで /exec を開いたとき）
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({status: 'ok', service: SHOP_NAME + ' 予約受付'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// フォームからの送信を受け取る（Content-Type: text/plain で JSON 文字列が届く）
function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (!d || !d.name || !d.tel || !d.email) throw new Error('必須項目が不足しています');
    processBooking(d);
    return ContentService.createTextOutput(JSON.stringify({status: 'ok'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: 'error', message: String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function processBooking(d) {
  saveToSheet_(d);
  sendOwnerMail_(d);
  if (d.email) sendCustomerMail_(d);
  return {status: 'ok'};
}

// 記録用スプレッドシート（初回に作成し、IDをスクリプトプロパティに保存して使い回す）
function getSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create(SHOP_NAME + ' 予約一覧');
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function saveToSheet_(d) {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['受付日時','送信方法','お名前','電話','メール','郵便番号','住所','建物','対応者','対応者名','対応者電話','追加作業確認先',
      '第1希望','第2希望','第3希望','空き枠連絡','エアコン詳細','オプション','ハウス','支払い','駐車','ペット','利用回数','備考','概算金額']);
  }
  const ac = (d.acUnits || []).map(u =>
    u.no + '台目[' + u.type + '] ' + u.location + ' 床:' + u.floorMaterial + ' 壁:' + u.wallMaterial +
    ' 作業スペース:' + u.workSpace + (u.moveOk ? '(' + u.moveOk + ')' : '') +
    (u.model ? ' 型番:' + u.model : '') + (u.year ? ' ' + u.year + '年' : '') + (u.kosho ? ' 高所' : '')).join('\n');
  sh.appendRow([d.submitted, d.sendVia, d.name, d.tel, d.email, d.zip, d.address, d.building, d.attendeeType, d.attendeeName,
    d.attendeeTel, d.upsellConfirm, d.date1 + ' ' + d.time1, d.date2 + ' ' + d.time2, d.date3 + ' ' + d.time3, d.cancelSlot,
    ac, d.acOptions, d.houseItems, d.payment, d.parking, d.pet, d.visitCount, d.note, d.estimate]);
}

function bodyText_(d) {
  const ac = (d.acUnits || []).map(u =>
    '  ・' + u.no + '台目[' + u.type + '] ' + u.location +
    '\n     床材:' + u.floorMaterial + ' / 壁材:' + (u.wallMaterial || '-') +
    '\n     作業スペース:' + u.workSpace + (u.moveOk ? '（' + u.moveOk + '）' : '') +
    (u.model ? '\n     型番:' + u.model : '') + (u.year ? ' 製造年:' + u.year : '') + (u.kosho ? '\n     ★高所作業（2.3m以上）' : '')
  ).join('\n') || '  なし';
  return [
    '■ お客様情報',
    'お名前：' + d.name + ' 様', '電話番号：' + d.tel, 'メール：' + d.email,
    'ご住所：〒' + d.zip + ' ' + d.address + (d.building ? ' ' + d.building : ''),
    '当日の対応者：' + d.attendeeType + (d.attendeeName ? '（' + d.attendeeName + ' / ' + d.attendeeTel + '）' : ''),
    d.upsellConfirm ? '追加作業の確認先：' + d.upsellConfirm : '',
    '', '■ ご依頼内容', '【エアコン】', ac,
    'オプション：' + d.acOptions, 'ハウス：' + d.houseItems,
    '概算金額：¥' + Number(d.estimate).toLocaleString() + '（税込）',
    '', '■ ご希望日程',
    '第1希望：' + d.date1 + ' ' + d.time1, '第2希望：' + d.date2 + ' ' + d.time2, '第3希望：' + d.date3 + ' ' + d.time3,
    '空き枠連絡：' + d.cancelSlot,
    '', '■ その他',
    '駐車スペース：' + d.parking, 'お支払い：' + d.payment, 'ペット：' + d.pet, 'ご利用：' + d.visitCount,
    d.note ? '備考：' + d.note : '', '', '受付日時：' + d.submitted + '（' + d.sendVia + '）'
  ].filter(function (l, i, a) { return !(l === '' && a[i - 1] === ''); }).join('\n');
}

function sendOwnerMail_(d) {
  if (!NOTIFY_EMAIL || NOTIFY_EMAIL.indexOf('@') < 0) return;
  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    replyTo: d.email || '',
    subject: '【予約リクエスト】' + d.name + ' 様',
    body: bodyText_(d)
  });
}

function sendCustomerMail_(d) {
  MailApp.sendEmail({
    to: d.email,
    name: SHOP_NAME,
    subject: '【' + SHOP_NAME + '】ご予約リクエストを受け付けました',
    body: d.name + ' 様\n\nこの度はご予約リクエストをいただき、ありがとうございます。\n' +
      '以下の内容で受け付けました。担当スタッフより改めてご連絡いたしますので、今しばらくお待ちください。\n' +
      '※ この時点では予約は確定しておりません。\n\n' + bodyText_(d) +
      '\n\n--------------------\n' + SHOP_NAME + '\n〒243-0204 神奈川県厚木市鳶尾5-26-41 2階\nTEL：070-5080-7830\nhttps://plus-alpha.org/\n'
  });
}
