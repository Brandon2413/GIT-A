// ============================================================
// RECEIPT OCR — uses Drive API's built-in OCR.
// Free, no GCP project, no API key. Accuracy is good for clear photos.
//
// Required: in the Apps Script editor, click Services (+) and add
// "Drive API". Without it, Drive.Files.create is undefined.
// ============================================================

function ocrImage(blob) {
  const resource = {
    name: 'ocr_temp_' + Date.now(),
    mimeType: 'application/vnd.google-apps.document'
  };
  const file = Drive.Files.create(resource, blob, { ocrLanguage: 'en' });
  let text = '';
  try {
    text = DocumentApp.openById(file.id).getBody().getText();
  } finally {
    try { DriveApp.getFileById(file.id).setTrashed(true); } catch (e) {}
  }
  return text;
}

function parseReceipt(text) {
  const lines = text.split(/\n/).map(l => l.trim()).filter(Boolean);

  // 1. Find total amount: prefer labeled lines (TOTAL, AMOUNT, NETT...).
  let amount = null;
  const totalRegex = /(?:grand\s*total|amount\s*due|total\s*amount|nett|sub.?total|balance|total|amount)[^\d]{0,12}(\d+\.\d{2})/i;
  for (const line of lines) {
    const m = line.match(totalRegex);
    if (m) { amount = parseFloat(m[1]); break; }
  }
  // Fallback: largest 2dp number in the receipt (usually the total).
  if (!amount) {
    const nums = (text.match(/\d+\.\d{2}/g) || [])
      .map(parseFloat).filter(n => n > 0 && n < 100000);
    if (nums.length) amount = Math.max(...nums);
  }

  // 2. Merchant: first reasonable text line near the top.
  let merchant = 'Receipt';
  for (const line of lines.slice(0, 6)) {
    if (line.length >= 3 && line.length <= 60 && /[a-z]/i.test(line) && !/receipt|invoice|tax/i.test(line)) {
      merchant = line.replace(/\s{2,}/g, ' ');
      break;
    }
  }

  // 3. Date: try common SG/intl formats. Falls back to today if not found.
  const monthRe = /Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec/i;
  let date = null;
  for (const line of lines) {
    // DD/MM/YYYY or DD-MM-YYYY (SG convention: day first)
    let m = line.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](20\d{2}|\d{2})/);
    if (m) {
      const yr = m[3].length === 2 ? 2000 + (+m[3]) : +m[3];
      const d = new Date(yr, +m[2] - 1, +m[1]);
      if (validDate(d)) { date = d; break; }
    }
    // DD MMM YYYY
    m = line.match(new RegExp('(\\d{1,2})\\s+(' + monthRe.source + ')[a-z]*\\s+(20\\d{2})', 'i'));
    if (m) {
      const d = new Date(`${m[2]} ${m[1]}, ${m[3]}`);
      if (validDate(d)) { date = d; break; }
    }
    // YYYY-MM-DD
    m = line.match(/(20\d{2})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (m) {
      const d = new Date(+m[1], +m[2] - 1, +m[3]);
      if (validDate(d)) { date = d; break; }
    }
  }

  return { amount, merchant, date };
}

function validDate(d) {
  return d instanceof Date && !isNaN(d.getTime()) && d.getFullYear() >= 2000 && d.getFullYear() <= 2100;
}
