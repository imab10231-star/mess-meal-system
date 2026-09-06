// ===================== মেস মিল ম্যানেজমেন্ট সিস্টেম — PHASE 1 =====================
// Backend: Google Apps Script + Google Sheets
// এই ফাইলটি Google Sheet-এর সাথে যুক্ত Apps Script প্রজেক্টে (Code.gs) পেস্ট করুন।

const SS = SpreadsheetApp.getActiveSpreadsheet();

function doGet() {
  ensureSheets();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('মেস মিল বোর্ড')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------------- Sheet setup (Section 20: Database Structure) ----------------
function ensureSheets() {
  createSheetIfMissing('Members', ['ID', 'Name', 'Mobile', 'PIN', 'Active', 'CreatedAt']);
  createSheetIfMissing('MealRecords', ['ID', 'Date', 'MemberID', 'Lunch', 'Dinner', 'CreatedAt', 'UpdatedAt']);
  createSheetIfMissing('Expenses', ['ID', 'Date', 'Category', 'Amount', 'Description', 'AddedBy', 'CreatedAt']);
  createSheetIfMissing('Payments', ['ID', 'MemberID', 'Date', 'Amount', 'Method', 'Note', 'CreatedAt']);
  createSheetIfMissing('MonthlySummaries', ['Month', 'Closed', 'ClosedAt']);
  createSheetIfMissing('Settings', ['Key', 'Value']);
}

function createSheetIfMissing(name, headers) {
  let sheet = SS.getSheetByName(name);
  if (!sheet) {
    sheet = SS.insertSheet(name);
    if (headers.length) {
      sheet.appendRow(headers);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }
  return sheet;
}

function getSheet(name) {
  return SS.getSheetByName(name) || createSheetIfMissing(name, []);
}

// ---------------- Helpers ----------------
function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function todayStr() {
  const tz = Session.getScriptTimeZone() || 'Asia/Dhaka';
  return Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
}

function newId(prefix) {
  return prefix + '_' + new Date().getTime() + '_' + Math.floor(Math.random() * 1000);
}

// ---------------- Members (Section 3) ----------------
function getActiveMembers() {
  ensureSheets();
  const members = sheetToObjects(getSheet('Members'));
  return members
    .filter(m => String(m.Active).toUpperCase() !== 'N' && String(m.Active).toUpperCase() !== 'FALSE')
    .map(m => ({ id: String(m.ID), name: m.Name, hasPin: !!(m.PIN && String(m.PIN).length) }));
}

function verifyPin(memberId, pin) {
  const members = sheetToObjects(getSheet('Members'));
  const m = members.find(x => String(x.ID) === String(memberId));
  if (!m) return false;
  if (!m.PIN || String(m.PIN).length === 0) return true; // পিন সেট না থাকলে খোলা
  return String(m.PIN) === String(pin);
}

// ---------------- Home Meal Board (Section 2, 26, 27) ----------------
function getBoardData(dateStr) {
  ensureSheets();
  dateStr = dateStr || todayStr();
  const members = getActiveMembers();
  const records = sheetToObjects(getSheet('MealRecords')).filter(r => String(r.Date) === dateStr);
  const byMember = {};
  records.forEach(r => { byMember[String(r.MemberID)] = r; });

  let lunchCount = 0, dinnerCount = 0;
  const board = members.map(m => {
    const r = byMember[m.id];
    const lunch = r ? Number(r.Lunch) : null;   // null = এখনো সিদ্ধান্ত নেওয়া হয়নি
    const dinner = r ? Number(r.Dinner) : null;
    if (lunch === 1) lunchCount++;
    if (dinner === 1) dinnerCount++;
    return { id: m.id, name: m.name, hasPin: m.hasPin, lunch: lunch, dinner: dinner };
  });

  return {
    date: dateStr,
    members: board,
    lunchCount: lunchCount,
    dinnerCount: dinnerCount,
    totalMeals: lunchCount + dinnerCount
  };
}

// ---------------- Update meal (Section 24, 25: validation + independence) ----------------
function updateMeal(memberId, dateStr, mealType, value, pin) {
  ensureSheets();
  dateStr = dateStr || todayStr();

  if (mealType !== 'Lunch' && mealType !== 'Dinner') {
    throw new Error('অজানা মিল ধরণ।');
  }
  if (value !== 0 && value !== 1) {
    throw new Error('ভুল মিল মান।');
  }
  if (!verifyPin(memberId, pin)) {
    throw new Error('ভুল পিন। আবার চেষ্টা করুন।');
  }

  const sheet = getSheet('MealRecords');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = {};
  headers.forEach((h, i) => col[h] = i);

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][col['Date']]) === dateStr && String(data[i][col['MemberID']]) === String(memberId)) {
      rowIndex = i;
      break;
    }
  }

  const now = new Date();

  if (rowIndex === -1) {
    // নতুন রেকর্ড — অন্য মিল ০ দিয়ে শুরু হবে, একটাই পরিবর্তন হবে
    const newRow = new Array(headers.length).fill('');
    newRow[col['ID']] = newId('MR');
    newRow[col['Date']] = dateStr;
    newRow[col['MemberID']] = memberId;
    newRow[col['Lunch']] = mealType === 'Lunch' ? value : 0;
    newRow[col['Dinner']] = mealType === 'Dinner' ? value : 0;
    newRow[col['CreatedAt']] = now;
    newRow[col['UpdatedAt']] = now;
    sheet.appendRow(newRow);
  } else {
    // বিদ্যমান রেকর্ড — শুধু এই মিলটাই আপডেট হবে, অন্যটা অপরিবর্তিত থাকবে
    sheet.getRange(rowIndex + 1, col[mealType] + 1).setValue(value);
    sheet.getRange(rowIndex + 1, col['UpdatedAt'] + 1).setValue(now);
  }

  return getBoardData(dateStr);
}
