// ============================================================
// TELEGRAM BOT FRONT-END (depends on Code.gs + Receipts.gs)
// Setup: run "💰 Spending → 🤖 Configure Telegram Bot" for steps.
// ============================================================

function doPost(e) {
  try {
    const update = JSON.parse(e.postData.contents);
    const msg = update.message || update.edited_message;
    if (msg) handleTelegramMessage(getSettings(), msg);
  } catch (err) {
    console.error('doPost error:', err);
  }
  return ContentService.createTextOutput('ok');
}

function handleTelegramMessage(settings, msg) {
  const chatId = String(msg.chat.id);

  // /chatid bypasses auth so the user can discover their own chat id.
  if (msg.text === '/chatid') {
    tgSend(settings, chatId,
      `Your chat id: ${chatId}\nPaste this into Settings → "Telegram Chat ID".`);
    return;
  }

  if (settings.telegramChatId && chatId !== settings.telegramChatId) {
    tgSend(settings, chatId, '🚫 This bot is locked to another chat.');
    return;
  }

  if (msg.photo) { handleReceiptPhoto(settings, msg); return; }
  if (!msg.text) return;

  const text = msg.text.trim();
  if (text === '/start' || text === '/help') return tgSend(settings, chatId, helpText());
  if (text === '/today') return tgSend(settings, chatId, todayLine());
  if (text === '/week')  return tgSend(settings, chatId, weekLine());
  if (text === '/undo')  return undoLastManual(settings, chatId);

  handleTextExpense(settings, chatId, text);
}

function helpText() {
  return [
    '💰 Spending bot',
    '',
    'Log an expense: "Lunch $8 hawker"',
    'Or send a receipt photo.',
    '',
    'Commands:',
    '/today - today\'s spending',
    '/week  - last 7 days',
    '/undo  - remove last manual entry',
    '/chatid - show this chat\'s id',
    '/help  - this message'
  ].join('\n');
}

function handleTextExpense(settings, chatId, text) {
  const m = text.match(/\$?\s*(\d+(?:\.\d{1,2})?)/);
  if (!m) return tgSend(settings, chatId, '❓ No amount found. Try: "Lunch $8 hawker"');
  const amount = +m[1];
  const merchant = text.replace(m[0], '').replace(/\s+/g, ' ').trim() || 'Cash';
  const category = categorize(merchant, 'Cash');
  appendManualTx({
    date: new Date(), bank: 'Manual', type: 'Cash',
    amount, sgd: amount, merchant, category, source: 'telegram-text'
  });
  tgSend(settings, chatId,
    `✅ ${merchant}: SGD ${amount.toFixed(2)} → ${category}\n` +
    `Today: SGD ${getTodayTotal().toFixed(2)}\n\n` +
    `Wrong? Reply /undo`);
}

function handleReceiptPhoto(settings, msg) {
  const chatId = String(msg.chat.id);
  tgSend(settings, chatId, '📸 Reading receipt...');
  try {
    const photos = msg.photo;
    const fileId = photos[photos.length - 1].file_id; // largest size
    const info = JSON.parse(UrlFetchApp.fetch(
      `https://api.telegram.org/bot${settings.telegramToken}/getFile?file_id=${fileId}`
    ).getContentText());
    if (!info.ok) return tgSend(settings, chatId, '❌ Could not fetch photo from Telegram.');
    const blob = UrlFetchApp.fetch(
      `https://api.telegram.org/file/bot${settings.telegramToken}/${info.result.file_path}`
    ).getBlob();
    const text = ocrImage(blob);
    const parsed = parseReceipt(text);
    if (!parsed.amount) {
      tgSend(settings, chatId,
        '❌ Could not find a total on the receipt.\n\n' +
        'First 300 chars of OCR:\n' + text.substring(0, 300) +
        '\n\nLog manually instead: "<merchant> $<amount>"');
      return;
    }
    const category = categorize(parsed.merchant, 'Cash');
    appendManualTx({
      date: parsed.date || new Date(), bank: 'Manual', type: 'Receipt',
      amount: parsed.amount, sgd: parsed.amount, merchant: parsed.merchant,
      category, source: 'telegram-receipt'
    });
    tgSend(settings, chatId,
      `✅ ${parsed.merchant}: SGD ${parsed.amount.toFixed(2)} → ${category}\n` +
      `Date: ${Utilities.formatDate(parsed.date || new Date(), 'GMT+8', 'dd MMM yyyy')}\n` +
      `Today: SGD ${getTodayTotal().toFixed(2)}\n\n` +
      `Wrong? Reply /undo`);
  } catch (err) {
    tgSend(settings, chatId, '❌ Receipt error: ' + err.message);
  }
}

function undoLastManual(settings, chatId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  const last = sheet.getLastRow();
  if (last < 2) return tgSend(settings, chatId, 'Nothing to undo.');
  const row = sheet.getRange(last, 1, 1, 11).getValues()[0];
  if (String(row[10] || '').startsWith('manual_')) {
    sheet.deleteRow(last);
    tgSend(settings, chatId, `↩️ Removed: ${row[6]} SGD ${row[5]}`);
  } else {
    tgSend(settings, chatId, 'Last entry was bank-imported — won\'t touch it. Edit the sheet directly if needed.');
  }
}

function appendManualTx(t) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  const msgId = 'manual_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
  sheet.appendRow([
    t.date, t.bank, t.type, 'SGD',
    t.amount, t.sgd, t.merchant, t.category,
    '[manual] via ' + t.source, 'success', msgId
  ]);
}

function getTodayTotal() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  const data = sheet.getDataRange().getValues();
  const todayStr = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd');
  let total = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === 'FAST') continue;
    const d = Utilities.formatDate(new Date(data[i][0]), 'GMT+8', 'yyyy-MM-dd');
    if (d === todayStr) total += +data[i][5];
  }
  return total;
}

function todayLine() { return `📅 Today: SGD ${getTodayTotal().toFixed(2)}`; }

function weekLine() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  const data = sheet.getDataRange().getValues();
  const weekAgo = new Date(Date.now() - 7 * 86400000);
  let total = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === 'FAST') continue;
    const d = new Date(data[i][0]);
    if (d >= weekAgo) total += +data[i][5];
  }
  return `📆 Last 7 days: SGD ${total.toFixed(2)}`;
}

function tgSend(settings, chatId, text) {
  if (!settings.telegramToken) return;
  try {
    UrlFetchApp.fetch(`https://api.telegram.org/bot${settings.telegramToken}/sendMessage`, {
      method: 'post',
      payload: { chat_id: String(chatId), text },
      muteHttpExceptions: true
    });
  } catch (e) { console.error('tgSend:', e); }
}

// ===== Setup helpers (called from the menu) =====
function configureTelegramBot() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Configure Telegram Bot',
    [
      'One-time setup:',
      '',
      '1. In Telegram, message @BotFather → /newbot. Copy the HTTP API token.',
      '2. In the Settings sheet, paste it into "Telegram Bot Token".',
      '3. Apps Script editor → Deploy → New deployment → type "Web app".',
      '   Execute as: Me. Who has access: Anyone. Copy the /exec URL.',
      '4. Back here: 💰 Spending → 🔗 Set Telegram Webhook, paste the URL.',
      '5. In Telegram, send /chatid to your bot. It will reply with your',
      '   chat id. Paste that into Settings → "Telegram Chat ID".',
      '6. Done. Try: "Lunch $8 hawker" or send a receipt photo.'
    ].join('\n'),
    ui.ButtonSet.OK
  );
}

function setTelegramWebhook() {
  const ui = SpreadsheetApp.getUi();
  const settings = getSettings();
  if (!settings.telegramToken) {
    ui.alert('Set "Telegram Bot Token" in the Settings sheet first.');
    return;
  }
  const r = ui.prompt('Set Webhook', 'Paste your Apps Script Web App /exec URL:', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const url = r.getResponseText().trim();
  if (!/^https:\/\/script\.google\.com\/macros\/s\/.+\/exec$/.test(url)) {
    ui.alert('Doesn\'t look like an Apps Script Web App URL.\nExpected: https://script.google.com/macros/s/.../exec');
    return;
  }
  const resp = UrlFetchApp.fetch(
    `https://api.telegram.org/bot${settings.telegramToken}/setWebhook?url=${encodeURIComponent(url)}`,
    { muteHttpExceptions: true }
  );
  ui.alert('Webhook result:\n\n' + resp.getContentText());
}
