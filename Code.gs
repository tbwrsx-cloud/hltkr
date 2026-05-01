// ============================================================
// HealthTrack — Google Apps Script Backend
// File: Code.gs
// Paste this entire file into script.google.com
// Then: Deploy → New Deployment → Web App
//       Execute as: Me | Who has access: Anyone
//       Copy the Web App URL into index.html
// ============================================================

const FILE_ID = "15T-6RNCGzIQVpEmnPnRa7eYGtrxZ4HG9PsTtLJ1wZvA";
const SHEET_DASHBOARD = "Dashboard";
const SHEET_PLAN      = "Plan&Rec";
const SHEET_KCAL_DB   = "kCal-DB";

// ── CORS wrapper ─────────────────────────────────────────────
function doGet(e)  { return handle(e); }
function doPost(e) { return handle(e); }

function handle(e) {
  try {
    const result = route(e);
    return respond(result);
  } catch(err) {
    return respond({ error: err.message }, 500);
  }
}

function respond(data, code) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// Handle preflight OPTIONS request for CORS
function doOptions(e) {
  return ContentService
    .createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}

// ── Router ────────────────────────────────────────────────────
function route(e) {
  const action = (e.parameter && e.parameter.action) || "";
  const method = e.postData ? "POST" : "GET";

  if (method === "GET") {
    if (action === "dashboard") return getDashboard();
    if (action === "today")     return getToday();
    if (action === "foods")     return getFoods();
    if (action === "history")   return getHistory();
    if (action === "plan")      return getPlan();
    if (action === "plan")      return getPlan();
    if (action === "addFood")   return addFood(body);
    return { status: "HealthTrack API running ✓" };
  }

  if (method === "POST") {
    const body = JSON.parse(e.postData.contents);
    if (action === "logMeal")    return logMeal(body);
    if (action === "logMetrics") return logMetrics(body);
  }

  return { error: "Unknown action: " + action };
}

// ── GET /dashboard ────────────────────────────────────────────
function getDashboard() {
  const wb = SpreadsheetApp.openById(FILE_ID);
  const sh = wb.getSheetByName(SHEET_DASHBOARD);
  const rows = sh.getDataRange().getValues();

  const actuals = [];
  const goals   = [];
  let current   = { weight: null, hba1c: null, ldl: null, trig: null };
  let foundFirst = false;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];

    // Parse date from col B (index 1)
    const dateVal = r[1];
    if (!dateVal || dateVal === 'Date' || dateVal === '') continue;

    const dateStr = formatSheetDate(dateVal);
    if (!dateStr) continue;

    const kcalOut = Number(r[2]) || 0;
    const kcalIn  = Number(r[3]) || 0;
    const protein = Number(r[4]) || 0;

    // Only include rows that have calorie data logged
    if (kcalIn > 0 || protein > 0) {
      actuals.push({ date: dateStr, kcalOut, kcalIn, protein });
    }

    // Biomarkers — cols H,I,J,K (indices 7,8,9,10)
    const w = r[7], h = r[8], l = r[9], t = r[10];
    if (w || h || l || t) {
      if (!foundFirst) {
        current = {
          weight: Number(w) || null,
          hba1c:  Number(h) || null,
          ldl:    Number(l) || null,
          trig:   Number(t) || null
        };
        foundFirst = true;
      }
    }

    // Goals — cols M,N,O,P,Q (indices 12,13,14,15,16)
    if (r[12]) {
      goals.push({
        year:   String(r[12]),
        weight: Number(r[13]) || null,
        hba1c:  Number(r[14]) || null,
        ldl:    Number(r[15]) || null,
        trig:   Number(r[16]) || null,
      });
    }
  }

  // Sort actuals by date ascending
  actuals.sort((a, b) => {
    const pa = a.date.split('/').reverse().join('');
    const pb = b.date.split('/').reverse().join('');
    return pa.localeCompare(pb);
  });

  return { actuals, goals, current };
}

// ── GET /today ────────────────────────────────────────────────
function getToday() {
  const wb   = SpreadsheetApp.openById(FILE_ID);
  const sh   = wb.getSheetByName(SHEET_PLAN);
  const rows = sh.getDataRange().getValues();
  const todayStr = formatSheetDate(new Date());
  const logged = {};

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0] || formatSheetDate(r[0]) !== todayStr) continue;
    const mealType = String(r[2] || '').trim();
    if (!mealType) continue;
    const mealCal = Number(r[11]) || 0;
    if (mealCal <= 0) continue; // skip unlogged planned rows
    logged[mealType] = {
      mealType,
      mealCal,
      mealProt: Number(r[13]) || 0,
      dailyCal: Number(r[12]) || null,
      dailyProt: Number(r[14]) || null,
    };
  }

  const meals = Object.values(logged);
  const dailyKcal = meals.reduce((s, m) => s + m.mealCal, 0);
  const dailyProt = meals.reduce((s, m) => s + m.mealProt, 0);
  return { date: todayStr, meals, dailyKcal, dailyProt };
}

// ── GET /history ──────────────────────────────────────────────
function getHistory() {
  const wb   = SpreadsheetApp.openById(FILE_ID);
  const sh   = wb.getSheetByName(SHEET_PLAN);
  const rows = sh.getDataRange().getValues();
  const byDate = {};

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    const d = formatSheetDate(r[0]);
    if (!byDate[d]) byDate[d] = { date: d, meals: [], kcal: 0, protein: 0 };
    const mealCal  = r[11] || 0;
    const mealProt = r[13] || 0;
    byDate[d].meals.push({ type: r[2], cal: mealCal, prot: mealProt });
    if (r[12]) byDate[d].kcal    = r[12];
    if (r[14]) byDate[d].protein = r[14];
  }

  return { history: Object.values(byDate).slice(-30) };
}

// ── GET /plan ─────────────────────────────────────────────────
// Returns ALL Plan&Rec rows - no date restriction
function getPlan() {
  const wb   = SpreadsheetApp.openById(FILE_ID);
  const sh   = wb.getSheetByName(SHEET_PLAN);
  const rows = sh.getDataRange().getValues();
  const byDate = {};

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    const dp = parseSheetDate(r[0]);
    if (!dp) continue;
    const dateStr = formatSheetDate(r[0]);
    if (!dateStr) continue;
    const mealType = String(r[2] || '').trim();
    if (!mealType || mealType === 'Meal Type') continue;

    if (!byDate[dateStr]) {
      byDate[dateStr] = {
        date: dateStr,
        dayName: String(r[1] || '').trim(),
        ts: new Date(dp.getFullYear(), dp.getMonth(), dp.getDate()).getTime(),
        meals: {}
      };
    }

    const items = [];
    for (let c = 3; c <= 9; c += 2) {
      const food = r[c], qty = r[c + 1];
      const fs = food ? String(food).trim() : '';
      if (fs && fs !== 'Food Item' && fs !== 'food item') {
        items.push({ food: fs, qty: (qty && !isNaN(Number(qty)) && Number(qty) > 0) ? Number(qty) : 1 });
      }
    }
    if (items.length > 0 || Number(r[11]) > 0) {
      byDate[dateStr].meals[mealType] = {
        mealType, items,
        mealCal: Number(r[11]) || 0,
        mealProt: Number(r[13]) || 0
      };
    }
  }

  const sorted = Object.values(byDate)
    .sort((a, b) => a.ts - b.ts)
    .map(day => ({
      date: day.date, dayName: day.dayName,
      meals: ['Breakfast','Lunch','Dinner','Snack']
        .map(mt => day.meals[mt] || { mealType: mt, items: [], mealCal: 0, mealProt: 0 })
    }));

  return { plan: sorted };
}

// ── GET /foods ────────────────────────────────────────────────
function getFoods() {
  const wb   = SpreadsheetApp.openById(FILE_ID);
  const sh   = wb.getSheetByName(SHEET_KCAL_DB);
  const rows = sh.getDataRange().getValues();
  const foods = [];

  for (let i = 1; i < rows.length; i++) {  // skip header row
    const r = rows[i];
    if (!r[1]) continue;
    foods.push({
      category: r[0],
      name:     r[1],
      portion:  r[2],
      cal:      r[3],
      protein:  r[4],
      glycemic: r[5] || "",
      remark:   r[6] || "",
    });
  }

  return { foods };
}

// ── POST /logMeal ─────────────────────────────────────────────
function logMeal(body) {
  const wb     = SpreadsheetApp.openById(FILE_ID);
  const planSh = wb.getSheetByName(SHEET_PLAN);
  const dashSh = wb.getSheetByName(SHEET_DASHBOARD);
  const { date, dayName, mealType, items, totalCal, totalProtein } = body;

  // Convert DD/MM/YYYY to Date object
  const dp = date.split('/');
  const dateObj = new Date(parseInt(dp[2]), parseInt(dp[1]) - 1, parseInt(dp[0]));

  // Append logged row
  planSh.appendRow([
    dateObj, dayName, mealType,
    items[0] ? items[0].food : '', items[0] ? items[0].qty : '',
    items[1] ? items[1].food : '', items[1] ? items[1].qty : '',
    items[2] ? items[2].food : '', items[2] ? items[2].qty : '',
    items[3] ? items[3].food : '', items[3] ? items[3].qty : '',
    totalCal, '', totalProtein, ''
  ]);

  // Recalculate day totals - deduplicate by mealType (keep highest cal entry)
  const allRows = planSh.getDataRange().getValues();
  const byMeal = {};
  allRows.forEach(r => {
    if (formatSheetDate(r[0]) === date && Number(r[11]) > 0) {
      byMeal[String(r[2]).trim()] = r;
    }
  });
  const dayKcal = Object.values(byMeal).reduce((s, r) => s + (Number(r[11]) || 0), 0);
  const dayProt = Object.values(byMeal).reduce((s, r) => s + (Number(r[13]) || 0), 0);

  // Write daily totals to last row
  const lastRow = planSh.getLastRow();
  planSh.getRange(lastRow, 13).setValue(dayKcal);
  planSh.getRange(lastRow, 15).setValue(dayProt);

  // Update Dashboard
  const dashRows = dashSh.getDataRange().getValues();
  for (let i = 1; i < dashRows.length; i++) {
    if (formatSheetDate(dashRows[i][1]) === date) {
      dashSh.getRange(i + 1, 4).setValue(dayKcal);
      dashSh.getRange(i + 1, 5).setValue(dayProt);
      break;
    }
  }

  return { success: true, date, mealType, totalCal, totalProtein, dayKcal, dayProt };
}

function logMetrics(body) {
  const wb     = SpreadsheetApp.openById(FILE_ID);
  const dashSh = wb.getSheetByName(SHEET_DASHBOARD);
  const rows   = dashSh.getDataRange().getValues();
  const { date, weight, hba1c, ldl, trig } = body;

  for (let i = 2; i < rows.length; i++) {
    const cellDate = formatSheetDate(rows[i][1]);
    if (cellDate === date) {
      if (weight != null) dashSh.getRange(i+1, 8).setValue(weight);  // col H
      if (hba1c  != null) dashSh.getRange(i+1, 9).setValue(hba1c);   // col I
      if (ldl    != null) dashSh.getRange(i+1, 10).setValue(ldl);    // col J
      if (trig   != null) dashSh.getRange(i+1, 11).setValue(trig);   // col K
      break;
    }
  }

  return { success: true, date, weight, hba1c, ldl, trig };
}


// ── POST /addFood ─────────────────────────────────────────────
// Body: { category, name, portion, cal, protein, glycemic, remark }
function addFood(body) {
  const wb = SpreadsheetApp.openById(FILE_ID);
  const sh = wb.getSheetByName(SHEET_KCAL_DB);

  // Check for duplicate name (case-insensitive)
  const rows = sh.getDataRange().getValues();
  const exists = rows.some((r, i) => i > 0 && 
    String(r[1]).trim().toLowerCase() === String(body.name).trim().toLowerCase()
  );
  if (exists) {
    return { success: false, error: "Food item already exists: " + body.name };
  }

  // Append new row in same format as existing kCal-DB rows
  // Cols: Category | Food Item | Portion | Calories(kcal) | Protein(g) | Glycemic | Remark
  sh.appendRow([
    body.category  || "Snack",
    body.name      || "",
    body.portion   || "1 serving",
    Number(body.cal)     || 0,
    Number(body.protein) || 0,
    body.glycemic  || "",
    body.remark    || "",
  ]);

  return {
    success: true,
    message: body.name + " added to kCal-DB",
    food: {
      category: body.category,
      name:     body.name,
      portion:  body.portion,
      cal:      Number(body.cal),
      protein:  Number(body.protein),
      glycemic: body.glycemic || "",
      remark:   body.remark  || "",
    }
  };
}

// ── Helpers ───────────────────────────────────────────────────
// Google Sheets always returns proper Date objects - clean parser
function parseSheetDate(val) {
  if (!val || val === '' || val === 0) return null;
  let d;
  if (val instanceof Date) {
    d = val;
  } else if (typeof val === 'number') {
    // Excel serial fallback (shouldn't happen with Sheets)
    d = new Date(Math.round((val - 25569) * 86400 * 1000));
  } else {
    d = new Date(String(val));
  }
  if (isNaN(d.getTime())) return null;
  return d;
}

function formatSheetDate(val) {
  const d = parseSheetDate(val);
  if (!d) return '';
  const dd   = String(d.getDate()).padStart(2,'0');
  const mm   = String(d.getMonth()+1).padStart(2,'0');
  const yyyy = d.getFullYear();
  return dd + '/' + mm + '/' + yyyy;
}

// Match a sheet date value against a DD/MM/YYYY string
function dateMatches(sheetVal, ddmmyyyy) {
  return formatSheetDate(sheetVal) === ddmmyyyy;
}

