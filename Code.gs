// ============================================================
// SPENDING TRACKER v3
// Changes from v2:
// 1. Settings sheet (configure without touching code)
// 2. Parse Failures sheet (catches silent fails)
// 3. Pagination on backfill (handles >500 emails)
// 4. Spending trend (this month vs last month)
// 5. Manual category overrides ([manual] in Notes)
// 6. Idempotent setup (preserves existing data)
// 7. Live FX rates via GOOGLEFINANCE
// ============================================================

const DEFAULT_FX_RATES = { SGD: 1.0, MYR: 0.31, CNY: 0.19, USD: 1.35 };
const DEFAULT_DAILY_TARGET = 20;
const DEFAULT_LOOKBACK = 30;
const DEFAULT_BUDGETS = {
  'Fast Food': 10, 'Restaurants': 25, 'Coffee/Drinks': 6,
  'Online Shopping': 15, 'Grab Food': 15, 'Grab Rides': 10,
  'TNG/Malaysia': 20, 'Public Transport': 8,
  'Subscriptions': 35, 'Hawker/Canteen': 30
};

const CATEGORY_RULES = [
  { name: 'Fast Food', regex: /popeyes|mcdonald|kfc|burger king|subway|pizza hut|domino|texas chicken/i },
  { name: 'Restaurants', regex: /takagi|chihanbao|ramen|sushi|din tai|saizeriya|crystal jade|paradise|astons/i },
  { name: 'Coffee/Drinks', regex: /luckin|chagee|starbucks|coffee bean|liho|gong cha|koi the|tealive/i },
  { name: 'Online Shopping', regex: /shopee|shein|taobao|lazada|amazon|qoo10|aliexpress|carousell|zalora/i },
  { name: 'Subscriptions', regex: /claude|anthropic|apple\.com|spotify|netflix|youtube premium|disney|adobe|chatgpt/i },
  { name: 'Grab Food', regex: /grab-ec|grabfood|grab\*\s*[a-z0-9]/i },
  { name: 'Grab Rides', regex: /grab rides|grab.transport/i },
  { name: 'TNG/Malaysia', regex: /tng|touch.?n.?go|causeway/i },
  { name: 'Public Transport', regex: /bus\/mrt|smrt|sbs transit|kovan mrt|transit\b|ezlink|ez-link|transitlink|comfort.delgro|tada\b|gojek/i },
  { name: 'Bike/Scooter', regex: /anywheel|helloride/i },
  { name: 'Hawker/Canteen', regex: /koufu|hawker|food court|kopitiam|nyp/i },
  { name: 'Vending', regex: /apac vending|vending/i },
  { name: 'Healthcare', regex: /healthcare|clinic|hospital|pharmacy|guardian|watsons|raffles medical/i },
  { name: 'Bills/Utilities', regex: /spl-eservices|sp services|singtel|starhub|m1\b|circles\.life/i },
];

function categorize(merchant, type) {
  if (type === 'FAST') return 'Bank Transfer';
  const m = (merchant || '').toString();
  for (const rule of CATEGORY_RULES) if (rule.regex.test(m)) return rule.name;
  if (type === 'PayNow') return 'PayNow';
  return 'Other';
}

// ===== SETTINGS =====
function getSettings() {
  initSettings();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  const s = { dailyTarget: DEFAULT_DAILY_TARGET, lookback: DEFAULT_LOOKBACK, budgets: {...DEFAULT_BUDGETS}, fx: {...DEFAULT_FX_RATES} };
  let section = 'config';
  for (let i = 1; i < data.length; i++) {
    const k = data[i][0], v = data[i][1];
    if (typeof k === 'string' && k.startsWith('---')) {
      if (k.includes('Budget')) section = 'budget';
      else if (k.includes('FX')) section = 'fx';
      else if (k.includes('Telegram')) section = 'telegram';
      continue;
    }
    if (!k || v === '' || v === null) continue;
    if (section === 'config') {
      if (k === 'Daily Target') s.dailyTarget = +v;
      if (k === 'Lookback Days') s.lookback = +v;
    } else if (section === 'budget') s.budgets[k] = +v;
    else if (section === 'fx') s.fx[k] = +v;
    else if (section === 'telegram') {
      if (k === 'Telegram Bot Token') s.telegramToken = String(v);
      else if (k === 'Telegram Chat ID') s.telegramChatId = String(v);
    }
  }
  return s;
}

function initSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
    const rows = [
      ['Setting', 'Value', 'Notes'],
      ['Daily Target', DEFAULT_DAILY_TARGET, 'Max SGD per day target'],
      ['Lookback Days', DEFAULT_LOOKBACK, 'Days the dashboard shows'],
      ['--- Weekly Budgets ---', '', 'Edit values to change limits'],
    ];
    Object.entries(DEFAULT_BUDGETS).forEach(([k,v]) => rows.push([k, v, '']));
    rows.push(['--- FX Rates ---', '', 'Live rates from Google Finance']);
    rows.push(['MYR', '=IFERROR(GOOGLEFINANCE("CURRENCY:MYRSGD"),0.31)', 'Auto']);
    rows.push(['CNY', '=IFERROR(GOOGLEFINANCE("CURRENCY:CNYSGD"),0.19)', 'Auto']);
    rows.push(['USD', '=IFERROR(GOOGLEFINANCE("CURRENCY:USDSGD"),1.35)', 'Auto']);
    rows.push(['SGD', 1.0, 'Always 1']);
    rows.push(['--- Telegram Bot ---', '', 'See "Configure Telegram Bot" menu']);
    rows.push(['Telegram Bot Token', '', 'From @BotFather in Telegram']);
    rows.push(['Telegram Chat ID', '', 'Send /chatid to your bot to get this']);
    sheet.getRange(1, 1, rows.length, 3).setValues(rows);
    sheet.getRange('A1:C1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#fff');
    sheet.setColumnWidth(1, 200); sheet.setColumnWidth(2, 140); sheet.setColumnWidth(3, 280);
    sheet.setFrozenRows(1);
    return;
  }
  // Migrate: append any missing rows so older installs pick up new settings.
  const existing = new Set(sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().map(r => r[0]));
  const additions = [];
  if (!existing.has('--- Telegram Bot ---')) additions.push(['--- Telegram Bot ---', '', 'See "Configure Telegram Bot" menu']);
  if (!existing.has('Telegram Bot Token')) additions.push(['Telegram Bot Token', '', 'From @BotFather in Telegram']);
  if (!existing.has('Telegram Chat ID')) additions.push(['Telegram Chat ID', '', 'Send /chatid to your bot to get this']);
  if (additions.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, additions.length, 3).setValues(additions);
  }
}

// ===== MENU =====
function onOpen() {
  SpreadsheetApp.getUi().createMenu('💰 Spending')
    .addItem('🔄 Refresh Dashboard', 'buildDashboard')
    .addItem('📥 Sync Now', 'syncTransactions')
    .addSeparator()
    .addItem('🔧 Set Lookback Period', 'promptLookback')
    .addItem('📅 Backfill 30 days', 'backfill30')
    .addItem('📅 Backfill 90 days', 'backfill90')
    .addItem('📅 Backfill 180 days', 'backfill180')
    .addItem('📅 Backfill 365 days', 'backfill365')
    .addSeparator()
    .addItem('📧 Send Daily Email Now', 'dailyCheckIn')
    .addItem('📧 Send Weekly Email Now', 'weeklyDeepDive')
    .addItem('📧 Send Parse Failure Report', 'monthlyParseFailureReport')
    .addSeparator()
    .addItem('🔄 Re-categorize All', 'recategorizeAll')
    .addItem('🔃 Setup / Migrate', 'setup')
    .addSeparator()
    .addItem('🤖 Configure Telegram Bot', 'configureTelegramBot')
    .addItem('🔗 Set Telegram Webhook', 'setTelegramWebhook')
    .addToUi();
}
function backfill30() { backfill(30); buildDashboard(); }
function backfill90() { backfill(90); buildDashboard(); }
function backfill180() { backfill(180); buildDashboard(); }
function backfill365() { backfill(365); buildDashboard(); }

function promptLookback() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Set Lookback Period', 'Days to display (30/60/90/180/365):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() === ui.Button.OK) {
    const days = parseInt(r.getResponseText());
    if (days > 0) {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === 'Lookback Days') { sheet.getRange(i+1, 2).setValue(days); break; }
      }
      buildDashboard();
      ui.alert(`✅ Lookback set to ${days} days.`);
    }
  }
}

// ===== SETUP (idempotent) =====
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  initSettings();

  let sheet = ss.getSheetByName('Transactions');
  if (!sheet) {
    sheet = ss.insertSheet('Transactions');
    sheet.appendRow(['Date','Bank','Type','Currency','Amount','SGD','Merchant','Category','Notes','Status','MsgID']);
    sheet.getRange('A1:K1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#fff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(7, 250); sheet.setColumnWidth(8, 130); sheet.setColumnWidth(9, 200);
  } else {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.includes('Category')) {
      sheet.insertColumnAfter(7); sheet.getRange(1, 8).setValue('Category');
      sheet.insertColumnAfter(8); sheet.getRange(1, 9).setValue('Notes');
      sheet.getRange('A1:K1').setFontWeight('bold').setBackground('#4285f4').setFontColor('#fff');
      if (sheet.getLastRow() > 1) {
        const ex = sheet.getRange(2, 1, sheet.getLastRow()-1, 11).getValues();
        const cats = ex.map(r => [categorize(r[6], r[2])]);
        sheet.getRange(2, 8, cats.length, 1).setValues(cats);
      }
    }
  }

  if (!ss.getSheetByName('Parse Failures')) {
    const pf = ss.insertSheet('Parse Failures');
    pf.appendRow(['Date', 'Sender', 'Subject', 'Snippet', 'Reason', 'MsgID']);
    pf.getRange('A1:F1').setFontWeight('bold').setBackground('#ea4335').setFontColor('#fff');
    pf.setFrozenRows(1);
  }

  // Only backfill if empty; otherwise just sync recent
  if (sheet.getLastRow() < 2) backfill(90);
  else syncTransactions();
  setupTriggers();
  buildDashboard();
  SpreadsheetApp.getUi().alert('✅ v3 ready! New: Settings, Parse Failures, Trend, Manual locks, Pagination, Live FX.');
}

const OUR_TRIGGERS = ['syncTransactions', 'buildDashboard', 'dailyCheckIn', 'weeklyDeepDive', 'monthlyParseFailureReport'];
function setupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (OUR_TRIGGERS.includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncTransactions').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('buildDashboard').timeBased().everyHours(1).create();
  ScriptApp.newTrigger('dailyCheckIn').timeBased().atHour(21).everyDays(1).create();
  ScriptApp.newTrigger('weeklyDeepDive').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(21).create();
  ScriptApp.newTrigger('monthlyParseFailureReport').timeBased().onMonthDay(1).atHour(9).create();
}

// ===== PAGINATION =====
function searchAllThreads(query) {
  const all = []; let start = 0; const batch = 100;
  while (true) {
    const res = GmailApp.search(query, start, batch);
    if (res.length === 0) break;
    all.push(...res);
    if (res.length < batch) break;
    start += batch;
    if (all.length >= 5000) break;
  }
  return all;
}

// ===== PARSING (with failure logging) =====
function parseTrust(msg) {
  const body = msg.getPlainBody();
  const subject = msg.getSubject();
  if (/protect|scam|rash|cash|tip|invest|enhanc|safeguard/i.test(subject)) return null;
  if (/cancelled|declined|refunded/i.test(subject)) return null;

  let m = body.match(/spent SGD ([\d.]+) at (.+?) on \d+ \w+ \d+/);
  if (m) return { bank:'Trust', type:'Card', currency:'SGD', amount:+m[1], merchant:m[2].trim(), status:'success' };
  m = body.match(/spent (\w{3}) ([\d.]+) (?:using Trust Link card )?at (.+?) on \d+ \w+ \d+/);
  if (m) return { bank:'Trust', type:'Card (Overseas)', currency:m[1], amount:+m[2], merchant:m[3].trim(), status:'success' };
  m = body.match(/PayNow transfer of SGD ([\d.]+) from A\/C ending \d+ to (.+?) on/);
  if (m) return { bank:'Trust', type:'PayNow', currency:'SGD', amount:+m[1], merchant:m[2].trim(), status:'success' };

  if (/transaction|spent|paynow/i.test(subject)) {
    return { _failed: true, reason: 'Trust: regex did not match body' };
  }
  return null;
}

function parseDBS(msg) {
  const body = (msg.getPlainBody() + ' ' + msg.getBody().replace(/<[^>]+>/g, ' ')).replace(/\s+/g, ' ');
  const subject = msg.getSubject();
  if (/eDocument|Limit|received a transfer|OTP|protect|scam/i.test(subject)) return null;

  const amtMatch = body.match(/Amount:\s*(?:SGD|S\$)\s*([\d,]+\.?\d{0,2})/i);
  if (!amtMatch) {
    if (/transaction|paynow|nets|fast|alert/i.test(subject)) {
      return { _failed: true, reason: 'DBS: amount field not found' };
    }
    return null;
  }
  const amount = parseFloat(amtMatch[1].replace(/,/g, ''));

  if (subject.includes('NETS')) {
    const merch = body.match(/(?:Merchant|To|Payee|At):\s*([A-Z0-9][^\n<]{2,80}?)(?:\s+If\s|\s+Thank|\s+Yours|\s+Date|\s+Reference|\s+Transaction|\s{3,})/i);
    if (!merch) logParseFailure(msg, 'DBS NETS: merchant not extracted (recorded as Unknown)');
    return { bank:'DBS', type:'NETS', currency:'SGD', amount, merchant: merch ? merch[1].trim() : 'NETS (Unknown Merchant)', status:'success' };
  }
  if (/PAYNOW/i.test(body)) {
    const toMatch = body.match(/To:\s*([A-Z0-9][^\n<]{2,80}?)(?:\s+If\s|\s+Thank|\s+Yours|\s+Date|\s+Reference|\s+Transaction|\s{3,})/i);
    return { bank:'DBS', type:'PayNow', currency:'SGD', amount, merchant: toMatch ? toMatch[1].trim() : 'PayNow (recipient unknown)', status:'success' };
  }
  if (/FAST/i.test(body)) {
    const toMatch = body.match(/To:\s*([A-Z0-9][^\n<]{2,80}?)(?:\s+If\s|\s+Thank|\s+Yours|\s+Date|\s+Reference|\s{3,})/i);
    return { bank:'DBS', type:'FAST', currency:'SGD', amount, merchant: toMatch ? toMatch[1].trim() : 'FAST Transfer', status:'success' };
  }
  if (subject.includes('Card Transaction')) {
    return { bank:'DBS', type:'Card', currency:'SGD', amount, merchant:'DBS Card Txn', status:'success' };
  }
  return { _failed: true, reason: 'DBS: amount found but no transaction type' };
}

function logParseFailure(msg, reason) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Parse Failures');
  if (!sheet) {
    sheet = ss.insertSheet('Parse Failures');
    sheet.appendRow(['Date', 'Sender', 'Subject', 'Snippet', 'Reason', 'MsgID']);
    sheet.getRange('A1:F1').setFontWeight('bold').setBackground('#ea4335').setFontColor('#fff');
    sheet.setFrozenRows(1);
  }
  const ids = new Set();
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 6, sheet.getLastRow()-1, 1).getValues().forEach(r => ids.add(r[0]));
  }
  if (ids.has(msg.getId())) return;
  sheet.appendRow([msg.getDate(), msg.getFrom(), msg.getSubject(), (msg.getPlainBody() || '').substring(0, 200), reason, msg.getId()]);
}

// ===== SYNC =====
function syncTransactions() { _sync(2); }
function backfill(days) { _sync(days); }

function _sync(days) {
  const settings = getSettings();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  const existing = new Set();
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 11, sheet.getLastRow()-1, 1).getValues().forEach(r => existing.add(r[0]));
  }
  const queries = [`from:from_us@trustbank.sg newer_than:${days}d`, `from:ibanking.alert@dbs.com newer_than:${days}d`];
  const rows = [];
  queries.forEach(q => {
    searchAllThreads(q).forEach(thread => {
      thread.getMessages().forEach(m => {
        if (existing.has(m.getId())) return;
        const tx = m.getFrom().includes('trustbank') ? parseTrust(m) : parseDBS(m);
        if (!tx) return;
        if (tx._failed) { logParseFailure(m, tx.reason); return; }
        if (tx.status === 'success') {
          if (!settings.fx[tx.currency]) {
            logParseFailure(m, `Unknown currency: ${tx.currency} - skipped to avoid distortion`);
            return;
          }
          const sgd = +(tx.amount * settings.fx[tx.currency]).toFixed(2);
          const cat = categorize(tx.merchant, tx.type);
          rows.push([m.getDate(), tx.bank, tx.type, tx.currency, tx.amount, sgd, tx.merchant, cat, '', tx.status, m.getId()]);
        }
      });
    });
  });
  if (rows.length) sheet.getRange(sheet.getLastRow()+1, 1, rows.length, 11).setValues(rows);
}

function recategorizeAll() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const data = sheet.getRange(2, 1, lastRow-1, 11).getValues();
  let preserved = 0;
  const newCats = data.map(row => {
    if (row[8] && row[8].toString().toLowerCase().includes('[manual]')) {
      preserved++;
      return [row[7]];
    }
    return [categorize(row[6], row[2])];
  });
  sheet.getRange(2, 8, newCats.length, 1).setValues(newCats);
  buildDashboard();
  SpreadsheetApp.getUi().alert(`✅ Re-categorized ${data.length - preserved} rows. Preserved ${preserved} manual locks.`);
}

// ===== EMAILS =====
function dailyCheckIn() {
  syncTransactions();
  const settings = getSettings();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  const data = sheet.getDataRange().getValues();
  const todayStr = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd');
  let total = 0;
  const txs = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === 'FAST') continue;
    const d = Utilities.formatDate(new Date(data[i][0]), 'GMT+8', 'yyyy-MM-dd');
    if (d === todayStr) { total += +data[i][5]; txs.push(data[i]); }
  }
  const verdict = total === 0 ? '🎉 SGD 0 today!' : total < settings.dailyTarget*0.7 ? '🟢 Under budget' : total < settings.dailyTarget ? '🟡 Watch it' : '🔴 OVER BUDGET';
  let html = `<h2>💰 Today: SGD ${total.toFixed(2)}</h2><h3>${verdict}</h3><p>Target: SGD ${settings.dailyTarget}</p>`;
  if (txs.length) {
    html += `<table border=1 cellpadding=5 style="border-collapse:collapse"><tr><th>Bank</th><th>Type</th><th>Merchant</th><th>Category</th><th>Amount</th></tr>`;
    txs.forEach(t => html += `<tr><td>${t[1]}</td><td>${t[2]}</td><td>${t[6]}</td><td>${t[7]}</td><td>${t[3]} ${t[4]} (≈SGD${t[5]})</td></tr>`);
    html += `</table>`;
  }
  GmailApp.sendEmail(Session.getActiveUser().getEmail(), `💰 Daily: SGD ${total.toFixed(2)} - ${todayStr}`, '', { htmlBody: html });
}

function weeklyDeepDive() {
  syncTransactions();
  const settings = getSettings();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Transactions');
  const data = sheet.getDataRange().getValues();
  const weekAgo = new Date(Date.now() - 7*86400000);
  const prevWeekStart = new Date(Date.now() - 14*86400000);
  let total = 0, prevTotal = 0;
  const daily = {}, cats = {};
  let earliestTx = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === 'FAST') continue;
    const d = new Date(data[i][0]);
    const sgd = +data[i][5];
    if (!earliestTx || d < earliestTx) earliestTx = d;
    if (d >= weekAgo) {
      const k = Utilities.formatDate(d, 'GMT+8', 'EEE dd MMM');
      total += sgd;
      daily[k] = (daily[k] || 0) + sgd;
      cats[data[i][7] || 'Other'] = (cats[data[i][7] || 'Other'] || 0) + sgd;
    } else if (d >= prevWeekStart) {
      prevTotal += sgd;
    }
  }
  const trend = prevTotal > 0 ? ((total - prevTotal) / prevTotal * 100) : 0;
  const trendStr = prevTotal > 0 ? `${trend >= 0 ? '↑' : '↓'} ${Math.abs(trend).toFixed(1)}% vs last week` : 'No prior week data';
  let biggest = '', max = 0;
  Object.entries(daily).forEach(([k,v]) => { if (v > max) { max = v; biggest = k; } });
  const daysOfData = earliestTx
    ? Math.min(7, Math.max(1, Math.ceil((Date.now() - earliestTx.getTime()) / 86400000) + 1))
    : 7;
  let html = `<h1>📊 Weekly Deep Dive</h1><h2>Total: SGD ${total.toFixed(2)} <small>(${trendStr})</small></h2>`;
  html += `<p>Daily avg: SGD ${(total/daysOfData).toFixed(2)} (over ${daysOfData} day${daysOfData===1?'':'s'}) | 🔴 Biggest: ${biggest} (SGD ${max.toFixed(2)})</p>`;
  html += `<h3>Daily breakdown</h3><ul>`;
  Object.entries(daily).forEach(([k,v]) => html += `<li>${v===max?'🔴':v>30?'🟡':'🟢'} ${k}: SGD ${v.toFixed(2)}</li>`);
  html += `</ul><h3>Category breakdown</h3><table border=1 cellpadding=5 style="border-collapse:collapse"><tr><th>Category</th><th>Spent</th><th>%</th></tr>`;
  Object.entries(cats).sort((a,b) => b[1]-a[1]).forEach(([c,v]) => html += `<tr><td>${c}</td><td>SGD ${v.toFixed(2)}</td><td>${(v/total*100).toFixed(1)}%</td></tr>`);
  html += `</table><h3>Budget check</h3><ul>`;
  Object.entries(settings.budgets).forEach(([c,b]) => { const s = cats[c]||0; html += `<li>${s<=b?'✅':'❌'} ${c}: SGD ${s.toFixed(2)} / ${b}</li>`; });
  html += `</ul>`;
  GmailApp.sendEmail(Session.getActiveUser().getEmail(), `📈 Weekly: SGD ${total.toFixed(2)} (${trendStr})`, '', { htmlBody: html });
}

function monthlyParseFailureReport() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Parse Failures');
  if (!sheet || sheet.getLastRow() < 2) return;
  const monthAgo = new Date(Date.now() - 30*86400000);
  const data = sheet.getRange(2, 1, sheet.getLastRow()-1, 6).getValues();
  const recent = data.filter(r => new Date(r[0]) >= monthAgo);
  if (recent.length === 0) return;
  let html = `<h2>⚠️ ${recent.length} bank emails couldn't be parsed in last 30 days</h2>`;
  html += `<p>This usually means your bank changed an email format. Check the <b>Parse Failures</b> sheet.</p>`;
  html += `<table border=1 cellpadding=5 style="border-collapse:collapse"><tr><th>Date</th><th>Subject</th><th>Reason</th></tr>`;
  recent.slice(0, 20).forEach(r => html += `<tr><td>${r[0]}</td><td>${r[2]}</td><td>${r[4]}</td></tr>`);
  html += `</table>`;
  GmailApp.sendEmail(Session.getActiveUser().getEmail(), `⚠️ ${recent.length} unparsed bank emails`, '', { htmlBody: html });
}

// ===== DASHBOARD =====
function buildDashboard() {
  syncTransactions();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings();
  const txSheet = ss.getSheetByName('Transactions');
  let dash = ss.getSheetByName('Dashboard');
  if (dash) ss.deleteSheet(dash);
  dash = ss.insertSheet('Dashboard', 0);

  const data = txSheet.getDataRange().getValues();
  if (data.length < 2) return;

  const lookback = settings.lookback;
  const now = new Date();
  const today = Utilities.formatDate(now, 'GMT+8', 'yyyy-MM-dd');
  const weekAgo = new Date(now.getTime() - 7 * 86400000);
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);
  const lastMonthStart = new Date(now.getFullYear(), now.getMonth()-1, 1);
  const daysInPrevMonth = new Date(now.getFullYear(), now.getMonth(), 0).getDate();
  const lastMonthSamePeriod = new Date(now.getFullYear(), now.getMonth()-1, Math.min(now.getDate(), daysInPrevMonth));
  const lookbackStart = new Date(now.getTime() - lookback * 86400000);

  let todayTotal = 0, weekTotal = 0, monthTotal = 0, lookbackTotal = 0;
  let lastMonthSameDay = 0, biggestDay = '', biggestAmt = 0, transferTotal = 0;
  const daily = {}, cats = {}, merchants = {};
  const allTxs = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const d = new Date(row[0]);
    if (isNaN(d.getTime())) continue;
    const sgd = +row[5] || 0;
    const dStr = Utilities.formatDate(d, 'GMT+8', 'yyyy-MM-dd');

    if (row[2] === 'FAST') {
      if (d >= lookbackStart) transferTotal += sgd;
      continue;
    }

    if (dStr === today) todayTotal += sgd;
    if (d >= weekAgo) weekTotal += sgd;
    if (d >= monthStart) monthTotal += sgd;
    if (d >= lastMonthStart && d <= lastMonthSamePeriod) lastMonthSameDay += sgd;
    if (d >= lookbackStart) {
      lookbackTotal += sgd;
      daily[dStr] = (daily[dStr] || 0) + sgd;
      const c = row[7] || categorize(row[6], row[2]);
      cats[c] = (cats[c] || 0) + sgd;
      merchants[row[6]] = (merchants[row[6]] || 0) + sgd;
      allTxs.push({ date: dStr, sgd, merchant: row[6], cat: c });
    }
  }
  Object.entries(daily).forEach(([k,v]) => { if (v > biggestAmt) { biggestAmt = v; biggestDay = k; } });

  const monthTrend = lastMonthSameDay > 0 ? ((monthTotal - lastMonthSameDay) / lastMonthSameDay * 100) : null;
  const trendIcon = monthTrend === null ? '' : monthTrend > 10 ? '🔴' : monthTrend < -10 ? '🟢' : '🟡';
  const trendStr = monthTrend === null ? '(no prior data)' : `${monthTrend >= 0 ? '↑' : '↓'}${Math.abs(monthTrend).toFixed(0)}% vs last month`;

  dash.getRange('A1').setValue('💰 Spending Dashboard').setFontSize(22).setFontWeight('bold').setFontColor('#1a73e8');
  dash.getRange('A2').setValue(`Lookback: ${lookback} days | Updated: ${Utilities.formatDate(now, 'GMT+8', 'dd MMM yyyy HH:mm')}`).setFontColor('#5f6368').setFontStyle('italic');
  dash.getRange('A3').setValue('💰 Spending menu → Set Lookback | Edit Settings sheet for budgets/FX').setFontColor('#5f6368').setFontSize(10);

  dash.getRange('A5:E5').setValues([['📅 TODAY', '📆 THIS WEEK', '🗓 THIS MONTH', `📊 ${lookback}d AVG`, '🔴 BIGGEST DAY']])
    .setFontWeight('bold').setBackground('#1a73e8').setFontColor('#fff').setHorizontalAlignment('center');
  dash.getRange('A6:E6').setValues([[
    'SGD ' + todayTotal.toFixed(2),
    'SGD ' + weekTotal.toFixed(2),
    `SGD ${monthTotal.toFixed(2)}\n${trendIcon} ${trendStr}`,
    'SGD ' + (lookbackTotal / Math.max(1, Math.min(lookback, Object.keys(daily).length || 1))).toFixed(2),
    biggestDay ? `${biggestDay}\nSGD ${biggestAmt.toFixed(2)}` : '-'
  ]]).setFontSize(12).setFontWeight('bold').setHorizontalAlignment('center').setBackground('#e8f0fe').setVerticalAlignment('middle').setWrap(true);
  dash.setRowHeight(6, 60);

  if (transferTotal > 0) {
    dash.getRange('A7').setValue(`ℹ️ Excluded SGD ${transferTotal.toFixed(2)} in FAST bank transfers (not personal spending)`).setFontColor('#5f6368').setFontSize(10).setFontStyle('italic');
  }

  let row = 9;
  dash.getRange(row, 1).setValue(`📅 Daily Spending`).setFontWeight('bold').setFontSize(13);
  dash.getRange(row+1, 1, 1, 2).setValues([['Date', 'SGD']]).setFontWeight('bold').setBackground('#e8f0fe');
  const sortedDaily = Object.entries(daily).sort();
  if (sortedDaily.length) dash.getRange(row+2, 1, sortedDaily.length, 2).setValues(sortedDaily);

  dash.getRange(row, 4).setValue(`📂 By Category`).setFontWeight('bold').setFontSize(13);
  dash.getRange(row+1, 4, 1, 4).setValues([['Category', 'SGD', '%', 'vs Budget']]).setFontWeight('bold').setBackground('#e8f0fe');
  const sortedCats = Object.entries(cats).sort((a,b) => b[1]-a[1]);
  if (sortedCats.length) {
    const catRows = sortedCats.map(([c,v]) => {
      const wkBudget = settings.budgets[c];
      const monthBudget = wkBudget ? wkBudget * 4.33 : null;
      const status = monthBudget ? `${v <= monthBudget ? '✅' : '❌'} ${monthBudget.toFixed(0)}` : '-';
      return [c, v.toFixed(2), (v/lookbackTotal*100).toFixed(1)+'%', status];
    });
    dash.getRange(row+2, 4, catRows.length, 4).setValues(catRows);
  }

  dash.getRange(row, 9).setValue(`🏪 Merchants (all)`).setFontWeight('bold').setFontSize(13);
  dash.getRange(row+1, 9, 1, 2).setValues([['Merchant', 'SGD']]).setFontWeight('bold').setBackground('#e8f0fe');
  const topMerch = Object.entries(merchants).sort((a,b) => b[1]-a[1]);
  if (topMerch.length) dash.getRange(row+2, 9, topMerch.length, 2).setValues(topMerch.map(([m,v]) => [m, v.toFixed(2)]));

  dash.getRange(row, 12).setValue(`💸 All Purchases (by size)`).setFontWeight('bold').setFontSize(13);
  dash.getRange(row+1, 12, 1, 3).setValues([['Date', 'Merchant', 'SGD']]).setFontWeight('bold').setBackground('#e8f0fe');
  const largestTxs = allTxs.sort((a,b) => b.sgd - a.sgd);
  if (largestTxs.length) dash.getRange(row+2, 12, largestTxs.length, 3).setValues(largestTxs.map(t => [t.date, t.merchant, t.sgd.toFixed(2)]));

  const chartTopRow = row + Math.max(sortedDaily.length, sortedCats.length, topMerch.length, largestTxs.length) + 4;
  if (sortedDaily.length) {
    dash.insertChart(dash.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(dash.getRange(row+1, 1, sortedDaily.length + 1, 2))
      .setPosition(chartTopRow, 1, 0, 0)
      .setOption('title', `Daily Spending (Last ${lookback} Days)`)
      .setOption('width', 800).setOption('height', 350)
      .setOption('legend.position', 'none').setOption('colors', ['#1a73e8']).build());
  }
  if (sortedCats.length) {
    dash.insertChart(dash.newChart().setChartType(Charts.ChartType.PIE)
      .addRange(dash.getRange(row+1, 4, sortedCats.length + 1, 2))
      .setPosition(chartTopRow + 18, 1, 0, 0)
      .setOption('title', 'Spending by Category')
      .setOption('width', 600).setOption('height', 400).setOption('pieHole', 0.4).build());
  }

  dash.setColumnWidth(1, 110); dash.setColumnWidth(2, 80);
  dash.setColumnWidth(4, 150); dash.setColumnWidth(5, 80); dash.setColumnWidth(6, 60); dash.setColumnWidth(7, 100);
  dash.setColumnWidth(9, 220); dash.setColumnWidth(10, 80);
  dash.setColumnWidth(12, 100); dash.setColumnWidth(13, 200); dash.setColumnWidth(14, 80);
}
