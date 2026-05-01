// ============================================================
// HealthTrack — Google Apps Script Backend
// File: Code.gs
// Paste this entire file into script.google.com
// Then: Deploy → New Deployment → Web App
//       Execute as: Me | Who has access: Anyone
//       Copy the Web App URL into index.html
// ============================================================

const FILE_ID = "1FQUkfpoA2cQxeCgL0GCtquks3aRHaPh6";
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

// ── Router ────────────────────────────────────────────────────
function route(e) {
  const action = (e.parameter && e.parameter.action) || "";
  const method = e.postData ? "POST" : "GET";

  if (method === "GET") {
    if (action === "dashboard") return getDashboard();
    if (action === "today")     return getToday();
    if (action === "foods")     return getFoods();
    if (action === "history")   return getHistory();
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

  // Row 2 = headers, Row 3+ = data (0-indexed: row[1]=headers, row[2+]=data)
  const actuals = [];
  const goals   = [];

  for (let i = 2; i < rows.length; i++) {
    const r = rows[i];
    const dateVal = r[1];
    if (!dateVal) continue;

    const dateStr = formatDate(dateVal);
    const kcalOut = r[2] || 0;
    const kcalIn  = r[3] || 0;
    const protein = r[4] || 0;

    if (kcalIn > 0 || protein > 0) {
      actuals.push({ date: dateStr, kcalOut, kcalIn, protein });
    }

    // Goals columns: Year(12), Weight(13), HbA1c(14), LDL(15), Trig(16)
    if (r[12]) {
      goals.push({
        year:   r[12],
        weight: r[13],
        hba1c:  r[14],
        ldl:    r[15],
        trig:   r[16],
      });
    }
  }

  // Current biomarkers from row 2 (first data row with metrics)
  const current = {
    weight: rows[2][7] || null,
    hba1c:  rows[2][8] || null,
    ldl:    rows[2][9] || null,
    trig:   rows[2][10] || null,
  };

  return { actuals, goals, current };
}

// ── GET /today ────────────────────────────────────────────────
function getToday() {
  const wb   = SpreadsheetApp.openById(FILE_ID);
  const sh   = wb.getSheetByName(SHEET_PLAN);
  const rows = sh.getDataRange().getValues();
  const todayStr = formatDate(new Date());
  const meals = [];

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const rowDate = formatDate(r[0]);
    if (rowDate !== todayStr) continue;

    // Col: 0=date,1=day,2=mealType,3=food1,4=qty1,5=food2,6=qty2,7=food3,8=qty3,9=food4,10=qty4
    // 11=mealCal, 12=dailyCal, 13=mealProt, 14=dailyProt
    const items = [];
    for (let c = 3; c <= 9; c += 2) {
      if (r[c]) items.push({ food: r[c], qty: r[c+1] || 1 });
    }
    meals.push({
      mealType: r[2],
      items,
      mealCal:  r[11] || 0,
      dailyCal: r[12] || null,
      mealProt: r[13] || 0,
      dailyProt:r[14] || null,
    });
  }

  const totalKcal = meals.reduce((s, m) => s + (m.dailyCal || m.mealCal), 0);
  // Use the last dailyCal found as total
  const dailyKcal  = meals.filter(m => m.dailyCal).slice(-1)[0]?.dailyCal || meals.reduce((s,m)=>s+m.mealCal,0);
  const dailyProt  = meals.filter(m => m.dailyProt).slice(-1)[0]?.dailyProt || meals.reduce((s,m)=>s+m.mealProt,0);

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
    const d = formatDate(r[0]);
    if (!byDate[d]) byDate[d] = { date: d, meals: [], kcal: 0, protein: 0 };
    const mealCal  = r[11] || 0;
    const mealProt = r[13] || 0;
    byDate[d].meals.push({ type: r[2], cal: mealCal, prot: mealProt });
    if (r[12]) byDate[d].kcal    = r[12];
    if (r[14]) byDate[d].protein = r[14];
  }

  return { history: Object.values(byDate).slice(-30) };
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
// Body: { date, dayName, mealType, items:[{food,qty,cal,protein}], totalCal, totalProtein }
function logMeal(body) {
  const wb      = SpreadsheetApp.openById(FILE_ID);
  const planSh  = wb.getSheetByName(SHEET_PLAN);
  const dashSh  = wb.getSheetByName(SHEET_DASHBOARD);

  const { date, dayName, mealType, items, totalCal, totalProtein } = body;

  // ── 1. Append to Plan&Rec ──────────────────────────────────
  // Format: date | day | mealType | food1 | qty1 | food2 | qty2 | food3 | qty3 | food4 | qty4 | mealCal | dailyCal | mealProt | dailyProt
  const row = [date, dayName, mealType,
    items[0]?.food||"", items[0]?.qty||"",
    items[1]?.food||"", items[1]?.qty||"",
    items[2]?.food||"", items[2]?.qty||"",
    items[3]?.food||"", items[3]?.qty||"",
    totalCal, "", totalProtein, ""
  ];
  planSh.appendRow(row);

  // ── 2. Recalculate today's daily totals ───────────────────
  const allRows   = planSh.getDataRange().getValues();
  const todayRows = allRows.filter(r => formatDate(r[0]) === date);
  const dayKcal   = todayRows.reduce((s, r) => s + (Number(r[11]) || 0), 0);
  const dayProt   = todayRows.reduce((s, r) => s + (Number(r[13]) || 0), 0);

  // Update the last snack row's dailyCal and dailyProt columns (cols 12,14 = 1-indexed 13,15)
  const lastPlanRow = planSh.getLastRow();
  planSh.getRange(lastPlanRow, 13).setValue(dayKcal);
  planSh.getRange(lastPlanRow, 15).setValue(dayProt);

  // ── 3. Update Dashboard ───────────────────────────────────
  const dashRows = dashSh.getDataRange().getValues();
  for (let i = 2; i < dashRows.length; i++) {
    const cellDate = formatDate(dashRows[i][1]);
    if (cellDate === date) {
      dashSh.getRange(i + 1, 4).setValue(dayKcal);    // col D = kCal-In
      dashSh.getRange(i + 1, 5).setValue(dayProt);    // col E = Daily Protein
      break;
    }
  }

  return {
    success: true,
    date,
    mealType,
    totalCal,
    totalProtein,
    dayKcal,
    dayProt,
    message: `${mealType} logged: ${totalCal} kcal, ${totalProtein}g protein. Day total: ${dayKcal} kcal, ${dayProt}g protein.`
  };
}

// ── POST /logMetrics ──────────────────────────────────────────
// Body: { date, weight, hba1c, ldl, trig }
function logMetrics(body) {
  const wb     = SpreadsheetApp.openById(FILE_ID);
  const dashSh = wb.getSheetByName(SHEET_DASHBOARD);
  const rows   = dashSh.getDataRange().getValues();
  const { date, weight, hba1c, ldl, trig } = body;

  for (let i = 2; i < rows.length; i++) {
    const cellDate = formatDate(rows[i][1]);
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

// ── Helpers ───────────────────────────────────────────────────
function formatDate(val) {
  if (!val) return "";
  let d;
  if (val instanceof Date) {
    d = val;
  } else if (typeof val === "number") {
    // Excel serial date
    d = new Date(Math.round((val - 25569) * 86400 * 1000));
  } else {
    d = new Date(val);
  }
  if (isNaN(d)) return String(val);
  return d.toLocaleDateString("en-GB", { day:"2-digit", month:"2-digit", year:"numeric" });
}
