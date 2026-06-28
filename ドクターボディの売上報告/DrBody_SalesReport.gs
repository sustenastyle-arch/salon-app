// ================================================
// Dr.Body, Inc. 売上自動報告スクリプト v2
// Square → Googleスプレッドシート自動入力 ＋ メール自動送信
// ================================================

// ▼▼▼ ここを自分の情報に書き換えてください ▼▼▼

const SQUARE_ACCESS_TOKEN = 'ここにProductionアクセストークンを貼り付け'; // EAAAで始まる文字列
const LOCATION_ID = 'ここにLocation IDを貼り付け';

const EMAIL_RECIPIENTS = [
  'dr.bodyinc@gmail.com',
  'hiroki@pt-w.jp',
  'sustenastyle@gmail.com'
].join(',');

// ハワイ時間 (UTC-10)
const HAWAII_OFFSET_HOURS = -10;

// 物販として認識するキーワード
const PRODUCT_KEYWORDS = ['procell', 'muscle rub', 'epicutis', 'serum', 'cream', 'lotion'];

// ▲▲▲ ここまで書き換え ▲▲▲


// ---------------------------------------------------
// メイン：前日の売上をシートに記入 ＆ 全員にメール送信
// ---------------------------------------------------
function sendDailySalesReport() {
  const reportDate = getYesterdayHawaii();
  runReport(reportDate, EMAIL_RECIPIENTS, false);
}


// ---------------------------------------------------
// テスト用：今日のデータ（Mamiさんのみに送信）
// ---------------------------------------------------
function sendTodayReport() {
  const now = new Date();
  const hawaiiNow = new Date(now.getTime() + HAWAII_OFFSET_HOURS * 3600000);
  const reportDate = new Date(hawaiiNow.getFullYear(), hawaiiNow.getMonth(), hawaiiNow.getDate());
  runReport(reportDate, 'sustenastyle@gmail.com', true);
}


// ---------------------------------------------------
// 共通処理：データ取得 → シート記入 → メール送信
// ---------------------------------------------------
function runReport(reportDate, recipients, isTest) {
  const payments = fetchSquarePayments(reportDate);
  if (payments === null) {
    Logger.log('Square APIの取得に失敗しました');
    return;
  }

  const orders = fetchOrdersBatch(payments);
  const result = processPayments(payments, orders);

  // Googleスプレッドシートに自動記入
  writeToSheet(reportDate, result);

  // メール送信
  const prefix = isTest ? '[テスト] ' : '';
  const subject = prefix + formatSubject(reportDate);
  const body = formatEmail(result, reportDate);
  GmailApp.sendEmail(recipients, subject, body);
  Logger.log('完了: ' + subject);
}


// ---------------------------------------------------
// Googleスプレッドシートに書き込む（Excelと同じ形式）
// ---------------------------------------------------
function writeToSheet(date, r) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = getMonthSheetName(date);

  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = createMonthSheet(ss, sheetName, date);
  }

  const day = date.getDate();
  const row = day + 4; // 行1〜4はヘッダー、5行目から日付データ

  const totalCash  = r.cashTreatment + r.cashProduct;
  const totalCard  = r.cardTreatment + r.cardProduct;
  const totalSales = totalCash + totalCard;
  const totalTips  = r.cashTip + r.cardTip;
  const grandTotal = totalSales + totalTips;

  sheet.getRange(row, 2).setValue(totalSales);       // B: Total Sales
  sheet.getRange(row, 3).setValue(r.clientCount);    // C: 客数
  sheet.getRange(row, 4).setValue(r.cashTreatment);  // D: Cash Treatment
  sheet.getRange(row, 5).setValue(r.cashProduct);    // E: Cash Product
  sheet.getRange(row, 6).setValue(totalCash);         // F: Total Cash
  sheet.getRange(row, 7).setValue(r.cashTip);        // G: Cash Tip
  sheet.getRange(row, 8).setValue(r.cardTreatment);  // H: Card Treatment
  sheet.getRange(row, 9).setValue(r.cardProduct);    // I: Card Product
  sheet.getRange(row, 10).setValue(totalCard);        // J: Total Card
  sheet.getRange(row, 11).setValue(r.cardTip);       // K: Card Tip
  sheet.getRange(row, 12).setValue(totalTips);        // L: Total Tip
  sheet.getRange(row, 13).setValue(grandTotal);       // M: Total Sales + Tip

  Logger.log('シート記入完了: ' + sheetName + ' ' + day + '日');
}


// ---------------------------------------------------
// 月別シートを自動で新規作成する
// ---------------------------------------------------
function createMonthSheet(ss, sheetName, date) {
  const sheet = ss.insertSheet(sheetName);
  const year  = date.getFullYear();
  const month = date.getMonth() + 1;
  const daysInMonth = new Date(year, month, 0).getDate();

  // タイトル
  sheet.getRange('A1').setValue('Dr.Body,Inc. Sales Report in ' + getMonthName(date.getMonth()) + ' ' + year);
  sheet.getRange('A1:M1').merge().setFontSize(14).setFontWeight('bold');

  // ヘッダー（3行目）
  const headers3 = ['Date', 'Total Sales', '客数', 'Cash', '', '', '', 'Card', '', '', '', 'Total tip', 'Total sales ans Tip'];
  headers3.forEach((h, i) => sheet.getRange(3, i + 1).setValue(h));
  sheet.getRange('D3:G3').merge().setHorizontalAlignment('center');
  sheet.getRange('H3:K3').merge().setHorizontalAlignment('center');
  sheet.getRange('L3:L4').merge();
  sheet.getRange('M3:M4').merge();

  // ヘッダー（4行目）
  const headers4 = ['', '', '', 'Treatment', 'Product', 'Total Cash', 'Tip', 'Treatment', 'Product', 'Total Card', 'Tip', '', ''];
  headers4.forEach((h, i) => sheet.getRange(4, i + 1).setValue(h));

  // ヘッダーの書式
  sheet.getRange('A3:M4').setFontWeight('bold').setBackground('#d0e4f7').setHorizontalAlignment('center');

  // 日付データ行（初期値0）
  for (let day = 1; day <= daysInMonth; day++) {
    const row = day + 4;
    sheet.getRange(row, 1).setValue(day);
    for (let col = 2; col <= 13; col++) {
      sheet.getRange(row, col).setValue(0).setNumberFormat('$#,##0.00');
    }
    sheet.getRange(row, 3).setNumberFormat('0'); // 客数は整数
  }

  // Total行
  const totalRow = daysInMonth + 5;
  sheet.getRange(totalRow, 1).setValue('Total').setFontWeight('bold');
  for (let col = 2; col <= 13; col++) {
    const letter = String.fromCharCode(64 + col);
    sheet.getRange(totalRow, col)
      .setFormula('=SUM(' + letter + '5:' + letter + (totalRow - 1) + ')')
      .setNumberFormat('$#,##0.00')
      .setFontWeight('bold');
  }
  sheet.getRange(totalRow, 3).setNumberFormat('0'); // 客数合計は整数

  // 列幅調整
  sheet.setColumnWidth(1, 50);
  for (let col = 2; col <= 13; col++) {
    sheet.setColumnWidth(col, 100);
  }

  return sheet;
}


// ---------------------------------------------------
// Square API：支払いを取得
// ---------------------------------------------------
function fetchSquarePayments(date) {
  const startUTC = new Date(date.getTime() - HAWAII_OFFSET_HOURS * 3600000);
  const endDate  = new Date(date.getTime() + 24 * 3600000);
  const endUTC   = new Date(endDate.getTime() - HAWAII_OFFSET_HOURS * 3600000);

  const url = 'https://connect.squareup.com/v2/payments'
    + '?location_id=' + LOCATION_ID
    + '&begin_time=' + encodeURIComponent(startUTC.toISOString())
    + '&end_time='   + encodeURIComponent(endUTC.toISOString())
    + '&limit=200&sort_order=ASC';

  const res = UrlFetchApp.fetch(url, {
    method: 'GET',
    headers: {
      'Authorization': 'Bearer ' + SQUARE_ACCESS_TOKEN,
      'Square-Version': '2025-01-23'
    },
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    Logger.log('API Error: ' + res.getContentText());
    return null;
  }

  const data = JSON.parse(res.getContentText());
  return (data.payments || []).filter(p => p.status === 'COMPLETED');
}


// ---------------------------------------------------
// Square API：注文を一括取得
// ---------------------------------------------------
function fetchOrdersBatch(payments) {
  const orderIds = payments.filter(p => p.order_id).map(p => p.order_id);
  if (orderIds.length === 0) return {};

  const res = UrlFetchApp.fetch('https://connect.squareup.com/v2/orders/batch-retrieve', {
    method: 'POST',
    headers: {
      'Authorization': 'Bearer ' + SQUARE_ACCESS_TOKEN,
      'Content-Type': 'application/json',
      'Square-Version': '2025-01-23'
    },
    payload: JSON.stringify({ location_id: LOCATION_ID, order_ids: orderIds }),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) return {};

  const data = JSON.parse(res.getContentText());
  const map = {};
  (data.orders || []).forEach(o => { map[o.id] = o; });
  return map;
}


// ---------------------------------------------------
// 支払いデータを集計する
// ---------------------------------------------------
function processPayments(payments, orders) {
  let cashTreatment = 0, cashProduct = 0, cashTip = 0;
  let cardTreatment = 0, cardProduct = 0, cardTip = 0;
  const clientList = [];
  const customerCache = {};

  for (const p of payments) {
    const total = (p.amount_money ? p.amount_money.amount : 0) / 100;
    const tip   = (p.tip_money   ? p.tip_money.amount   : 0) / 100;
    const sale  = total - tip;
    const isCash = p.source_type === 'CASH';

    if (isCash) cashTip += tip;
    else        cardTip += tip;

    const order = p.order_id ? orders[p.order_id] : null;
    let treatAmt = 0, prodAmt = 0;
    const itemNames = [];

    if (order && order.line_items) {
      for (const item of order.line_items) {
        const itemTotal = item.total_money ? item.total_money.amount / 100 : 0;
        const name = item.name || '';
        itemNames.push(name);
        const isProduct = PRODUCT_KEYWORDS.some(k => name.toLowerCase().includes(k));
        if (isProduct) prodAmt += itemTotal;
        else           treatAmt += itemTotal;
      }
    } else {
      treatAmt = sale;
    }

    if (isCash) { cashTreatment += treatAmt; cashProduct += prodAmt; }
    else        { cardTreatment += treatAmt; cardProduct += prodAmt; }

    // 顧客名を取得
    let customerName = '不明';
    if (p.customer_id) {
      if (customerCache[p.customer_id]) {
        customerName = customerCache[p.customer_id];
      } else {
        try {
          const cr = UrlFetchApp.fetch('https://connect.squareup.com/v2/customers/' + p.customer_id, {
            headers: {
              'Authorization': 'Bearer ' + SQUARE_ACCESS_TOKEN,
              'Square-Version': '2025-01-23'
            },
            muteHttpExceptions: true
          });
          if (cr.getResponseCode() === 200) {
            const c = JSON.parse(cr.getContentText()).customer;
            customerName = ((c.given_name || '') + ' ' + (c.family_name || '')).trim() || '不明';
            customerCache[p.customer_id] = customerName;
          }
        } catch(e) {}
      }
    }

    clientList.push({ name: customerName, items: itemNames.join(' / ') || '不明', treatment: treatAmt, product: prodAmt, tip, paymentType: isCash ? '現金' : 'カード' });
  }

  return { clientCount: payments.length, cashTreatment, cashProduct, cashTip, cardTreatment, cardProduct, cardTip, clientList };
}


// ---------------------------------------------------
// メール本文を作成する
// ---------------------------------------------------
function formatEmail(r, date) {
  const totalCash  = r.cashTreatment + r.cashProduct;
  const totalCard  = r.cardTreatment + r.cardProduct;
  const totalSales = totalCash + totalCard;
  const totalTips  = r.cashTip + r.cardTip;
  const grandTotal = totalSales + totalTips;
  const line = '─'.repeat(38);

  let body = 'Dr.Body, Inc. 売上報告\n';
  body += '日付: ' + formatDateJP(date) + '\n';
  body += line + '\n\n';
  body += '来客数: ' + r.clientCount + ' 人\n\n';
  body += '【現金】\n';
  body += '  施術売上:   $' + fmt(r.cashTreatment) + '\n';
  body += '  物販売上:   $' + fmt(r.cashProduct)   + '\n';
  body += '  現金合計:   $' + fmt(totalCash)        + '\n';
  body += '  チップ:     $' + fmt(r.cashTip)        + '\n\n';
  body += '【クレジットカード】\n';
  body += '  施術売上:   $' + fmt(r.cardTreatment) + '\n';
  body += '  物販売上:   $' + fmt(r.cardProduct)   + '\n';
  body += '  カード合計: $' + fmt(totalCard)        + '\n';
  body += '  チップ:     $' + fmt(r.cardTip)        + '\n\n';
  body += line + '\n';
  body += 'チップ合計:         $' + fmt(totalTips)  + '\n';
  body += '総売上 (施術+物販): $' + fmt(totalSales) + '\n';
  body += '総売上 + チップ:    $' + fmt(grandTotal) + '\n';
  body += line + '\n\n';
  body += '【来店リスト】\n';
  r.clientList.forEach(function(c, i) {
    body += '\n' + (i + 1) + '. ' + c.name + '\n';
    body += '   コース: ' + c.items + '\n';
    body += '   施術: $' + fmt(c.treatment);
    if (c.product > 0) body += '  物販: $' + fmt(c.product);
    body += '  チップ: $' + fmt(c.tip) + '  [' + c.paymentType + ']\n';
  });
  body += '\n' + line + '\n';
  body += 'このメールはDr.Body売上管理システムから自動送信されています。';
  return body;
}


// ---------------------------------------------------
// 毎朝9時の自動トリガーを設定（一度だけ実行）
// ---------------------------------------------------
function setDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sendDailySalesReport')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  Logger.log('✅ トリガー設定完了！毎朝9時に自動送信されます。');
}


// ---------------------------------------------------
// ユーティリティ関数
// ---------------------------------------------------
function getMonthSheetName(date) {
  return getMonthName(date.getMonth()) + ' ' + date.getFullYear();
}

function getMonthName(i) {
  return ['January','February','March','April','May','June',
          'July','August','September','October','November','December'][i];
}

function getYesterdayHawaii() {
  const now = new Date();
  const hawaiiNow = new Date(now.getTime() + HAWAII_OFFSET_HOURS * 3600000);
  return new Date(hawaiiNow.getFullYear(), hawaiiNow.getMonth(), hawaiiNow.getDate() - 1);
}

function formatSubject(date) {
  const days = ['日','月','火','水','木','金','土'];
  return 'Dr.Body 売上報告 ' + (date.getMonth()+1) + '/' + date.getDate() + '/' + date.getFullYear() + ' (' + days[date.getDay()] + ')';
}

function formatDateJP(date) {
  const days = ['日','月','火','水','木','金','土'];
  return (date.getMonth()+1) + '/' + date.getDate() + '/' + date.getFullYear() + ' (' + days[date.getDay()] + '曜日)';
}

function fmt(n) {
  return (Math.round(n * 100) / 100).toFixed(2);
}
