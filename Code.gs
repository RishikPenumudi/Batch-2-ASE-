const SHEETS = {
  USERS: "Users",
  FOOD: "FoodData",
  FEEDBACK: "Feedback",
  SETTINGS: "Settings",
  ALERTS: "Alerts"
};
const SESSION_TTL_SECONDS = 21600;
const BOOTSTRAP_ADMIN = { username: "ASE", password: "2026" }; // Demo only; change after setup.

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Hostel Food Waste Monitoring System")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const definitions = {};
  definitions[SHEETS.USERS] = ["Username", "PasswordHash", "Role", "Active", "CreatedAt"];
  definitions[SHEETS.FOOD] = ["Date", "Meal", "Prepared_KG", "Consumed_KG", "Waste_KG", "Students", "Cost_Per_KG", "Waste_Cost", "Recorded_By", "CreatedAt"];
  definitions[SHEETS.FEEDBACK] = ["Timestamp", "Meal", "Rating", "Comments", "Username"];
  definitions[SHEETS.SETTINGS] = ["Setting", "Value"];
  definitions[SHEETS.ALERTS] = ["Timestamp", "Date", "Meal", "Waste_KG", "Waste_Percent", "Message", "Status"];
  Object.keys(definitions).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(definitions[name]);
    } else {
      // Safely extend existing project sheets with any new columns.
      const existingHeaders = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getDisplayValues()[0];
      definitions[name].forEach(header => {
        if (!existingHeaders.includes(header)) {
          sh.getRange(1, sh.getLastColumn() + 1).setValue(header);
          existingHeaders.push(header);
        }
      });
    }
    sh.setFrozenRows(1);
  });

  const settings = getSheet_(SHEETS.SETTINGS);
  const existing = settings.getDataRange().getValues().slice(1).map(r => r[0]);
  [
    ["Hostel_Name", "Smart Hostel"],
    ["Waste_Target_Percent", "5"],
    ["Currency", "INR"],
    ["Alert_Email", ""]
  ].forEach(row => { if (!existing.includes(row[0])) settings.appendRow(row); });

  const users = getSheet_(SHEETS.USERS);
  const names = users.getLastRow() > 1
    ? users.getRange(2, 1, users.getLastRow() - 1, 1).getDisplayValues().flat()
    : [];
  if (!names.includes(BOOTSTRAP_ADMIN.username)) {
    users.appendRow([BOOTSTRAP_ADMIN.username, hashPassword_(BOOTSTRAP_ADMIN.password), "Admin", true, new Date()]);
  }
  return {success:true, message:"Setup complete. Demo administrator: ASE / 2026. Change this password after signing in."};
}

function registerUser(payload) {
  setupGuard_();
  payload = payload || {};
  const username = String(payload.username || "").trim();
  const password = String(payload.password || "");
  if (!/^[A-Za-z0-9_.-]{3,30}$/.test(username)) throw new Error("Username must be 3–30 letters, numbers, dots, underscores, or hyphens.");
  if (password.length < 8) throw new Error("Choose a password with at least 8 characters.");
  const sh = getSheet_(SHEETS.USERS);
  const users = sh.getDataRange().getDisplayValues();
  if (users.slice(1).some(r => String(r[0]).toLowerCase() === username.toLowerCase())) throw new Error("That username already exists.");
  sh.appendRow([username, hashPassword_(password), "Student", true, new Date()]);
  return {success:true, message:"Registration successful. You can now sign in as a Student."};
}

function login(payload) {
  setupGuard_();
  payload = payload || {};
  const username = String(payload.username || "").trim();
  const password = String(payload.password || "");
  const rows = getSheet_(SHEETS.USERS).getDataRange().getValues();
  const user = rows.slice(1).find(r => String(r[0]).toLowerCase() === username.toLowerCase());
  if (!user || user[1] !== hashPassword_(password) || user[3] !== true) {
    throw new Error("Invalid username/password or inactive account.");
  }
  const token = Utilities.getUuid() + Utilities.getUuid();
  const session = {username:String(user[0]), role:String(user[2])};
  CacheService.getScriptCache().put("session_" + token, JSON.stringify(session), SESSION_TTL_SECONDS);
  return {success:true, token:token, username:session.username, role:session.role, expiresIn:SESSION_TTL_SECONDS};
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove("session_" + token);
  return {success:true};
}

function changePassword(token, oldPassword, newPassword) {
  const session = requireSession_(token);
  if (String(newPassword || "").length < 8) throw new Error("New password must have at least 8 characters.");
  const sh = getSheet_(SHEETS.USERS);
  const data = sh.getDataRange().getValues();
  const idx = data.findIndex((r, i) => i > 0 && String(r[0]) === session.username);
  if (idx < 1 || data[idx][1] !== hashPassword_(String(oldPassword || ""))) throw new Error("Current password is incorrect.");
  sh.getRange(idx + 1, 2).setValue(hashPassword_(String(newPassword)));
  return {success:true, message:"Password changed."};
}

function getAppData(token) {
  const session = requireSession_(token);
  const food = readFood_();
  const feedback = readFeedback_();
  const target = Number(getSetting_("Waste_Target_Percent", 5));
  const total = summarize_(food);
  const byMeal = {};
  const byDate = {};
  food.forEach(r => {
    if (!byMeal[r.meal]) byMeal[r.meal] = {meal:r.meal, prepared:0, consumed:0, waste:0, cost:0, records:0};
    const m = byMeal[r.meal]; m.prepared += r.prepared; m.consumed += r.consumed; m.waste += r.waste; m.cost += r.wasteCost; m.records++;
    if (!byDate[r.date]) byDate[r.date] = {date:r.date, prepared:0, consumed:0, waste:0, cost:0};
    const d = byDate[r.date]; d.prepared += r.prepared; d.consumed += r.consumed; d.waste += r.waste; d.cost += r.wasteCost;
  });
  const meals = Object.values(byMeal).map(roundObject_);
  const days = Object.values(byDate).sort((a,b) => a.date.localeCompare(b.date)).map(roundObject_);
  const alerts = food.filter(r => r.prepared > 0 && r.waste / r.prepared * 100 > target)
    .sort((a,b) => b.date.localeCompare(a.date)).slice(0,20)
    .map(r => ({date:r.date, meal:r.meal, waste:r.waste, percent:round2_(r.waste/r.prepared*100), message:r.meal+" on "+r.date+" wasted "+round2_(r.waste/r.prepared*100)+"% (target "+target+"%)."}));
  const sortedMeals = meals.slice().sort((a,b) => b.waste-a.waste);
  return {
    user:session,
    settings:{hostelName:getSetting_("Hostel_Name","Smart Hostel"), wasteTargetPercent:target, currency:getSetting_("Currency","INR")},
    totals:total,
    mealTotals:meals,
    dailyTotals:days,
    highestWasteMeal:sortedMeals[0] || null,
    alerts:alerts,
    recentRecords:food.slice(-10).reverse(),
    feedback:{averageRating:feedback.length ? round2_(feedback.reduce((s,r)=>s+r.rating,0)/feedback.length) : 0, total:feedback.length, recent:feedback.slice(-10).reverse()},
    users: session.role === "Admin" ? listUsers_() : []
  };
}

function addFoodRecord(token, payload) {
  const session = requireRole_(token, ["Admin","Staff"]);
  payload = payload || {};
  const date = normalizeDate_(payload.date);
  const meal = String(payload.meal || "").trim();
  const prepared = Number(payload.prepared);
  const consumed = Number(payload.consumed);
  const students = Number(payload.students);
  const costPerKg = Number(payload.costPerKg);
  if (!["Breakfast","Lunch","Dinner","Snacks"].includes(meal)) throw new Error("Choose a valid meal.");
  if (![prepared,consumed,students,costPerKg].every(Number.isFinite) || prepared <= 0 || consumed < 0 || consumed > prepared || students < 0 || costPerKg < 0) {
    throw new Error("Check the values: prepared must be above 0, consumed cannot exceed prepared, and other values cannot be negative.");
  }
  const waste = round2_(prepared-consumed);
  const wasteCost = round2_(waste*costPerKg);
  getSheet_(SHEETS.FOOD).appendRow([date,meal,prepared,consumed,waste,students,costPerKg,wasteCost,session.username,new Date()]);
  const percent = prepared ? waste/prepared*100 : 0;
  const target = Number(getSetting_("Waste_Target_Percent",5));
  if (percent > target) {
    const message = meal+" waste is "+round2_(percent)+"%, above the "+target+"% target.";
    getSheet_(SHEETS.ALERTS).appendRow([new Date(),date,meal,waste,round2_(percent),message,"Open"]);
    sendWasteAlert_(date, meal, waste, percent, target);
  }
  return {success:true, message:"Food record saved.", waste:waste, wasteCost:wasteCost};
}

function addFeedback(token, payload) {
  const session = requireSession_(token);
  payload = payload || {};
  const meal = String(payload.meal || "").trim();
  const rating = Number(payload.rating);
  const comments = String(payload.comments || "").trim().slice(0,1000);
  if (!["Breakfast","Lunch","Dinner","Snacks"].includes(meal)) throw new Error("Choose a valid meal.");
  if (!Number.isInteger(rating) || rating < 1 || rating > 5) throw new Error("Rating must be from 1 to 5.");
  getSheet_(SHEETS.FEEDBACK).appendRow([new Date(),meal,rating,comments,session.username]);
  return {success:true,message:"Thank you for your feedback."};
}

function predictDemand(token, meal) {
  requireSession_(token);
  meal = String(meal || "Breakfast");
  const rows = readFood_().filter(r => r.meal === meal);
  if (!rows.length) return {meal:meal, predictedKg:0, confidence:"Not enough data", sampleDays:0, note:"Add meal records to enable a historical estimate."};
  const recent = rows.slice(-14);
  const avg = recent.reduce((s,r) => s+r.prepared,0)/recent.length;
  const avgConsumed = recent.reduce((s,r) => s+r.consumed,0)/recent.length;
  const suggested = Math.max(0, Math.round(avgConsumed*10)/10);
  return {meal:meal, predictedKg:suggested, averagePreparedKg:round2_(avg), confidence:recent.length>=7?"Moderate (7+ recent records)":"Low (limited history)", sampleDays:recent.length, note:"Estimate based on the mean consumed quantity in the most recent records for this meal. Check attendance and events before preparing food."};
}

function getCSVReport(token, period, dateOrMonth) {
  requireRole_(token, ["Admin","Staff"]);
  const periodValue = String(period || "");
  let start, end, filename;
  if (periodValue === "daily") {
    const date = normalizeDate_(dateOrMonth);
    start = date; end = date; filename = "hostel_daily_" + date + ".csv";
  } else if (periodValue === "weekly") {
    const date = normalizeDate_(dateOrMonth);
    const d = new Date(date+"T12:00:00");
    const day = (d.getDay()+6)%7;
    d.setDate(d.getDate()-day);
    start = dateString_(d);
    d.setDate(d.getDate()+6);
    end = dateString_(d);
    filename = "hostel_weekly_" + start + "_to_" + end + ".csv";
  } else if (periodValue === "monthly") {
    if (!/^\d{4}-\d{2}$/.test(String(dateOrMonth))) throw new Error("Choose a valid month.");
    start = String(dateOrMonth)+"-01";
    const d = new Date(start+"T12:00:00"); d.setMonth(d.getMonth()+1); d.setDate(d.getDate()-1);
    end = dateString_(d);
    filename = "hostel_monthly_" + dateOrMonth + ".csv";
  } else throw new Error("Report period must be daily, weekly, or monthly.");
  const rows = readFood_().filter(r => r.date >= start && r.date <= end);
  const header = ["Date","Meal","Prepared_KG","Consumed_KG","Waste_KG","Students","Cost_Per_KG","Waste_Cost","Recorded_By"];
  const csvRows = [header].concat(rows.map(r => [r.date,r.meal,r.prepared,r.consumed,r.waste,r.students,r.costPerKg,r.wasteCost,r.recordedBy]));
  return {filename:filename, csv:"\uFEFF"+csvRows.map(row=>row.map(csvEscape_).join(",")).join("\r\n"), count:rows.length};
}

function updateSettings(token, payload) {
  requireRole_(token, ["Admin"]);
  payload = payload || {};
  const allowed = ["Hostel_Name","Waste_Target_Percent","Currency","Alert_Email"];
  const sh = getSheet_(SHEETS.SETTINGS);
  const values = sh.getDataRange().getValues();
  Object.keys(payload).forEach(key => {
    if (!allowed.includes(key)) return;
    if (key === "Waste_Target_Percent") {
      const n = Number(payload[key]);
      if (!Number.isFinite(n) || n < 0 || n > 100) throw new Error("Waste target must be 0–100.");
    }
    const index = values.findIndex((r,i) => i > 0 && r[0] === key);
    if (index > 0) sh.getRange(index+1,2).setValue(payload[key]);
    else sh.appendRow([key,payload[key]]);
  });
  return {success:true,message:"Settings updated."};
}

function adminUpdateUser(token, username, role, active) {
  requireRole_(token, ["Admin"]);
  if (!["Admin","Staff","Student"].includes(role)) throw new Error("Invalid role.");
  const sh = getSheet_(SHEETS.USERS);
  const rows = sh.getDataRange().getValues();
  const index = rows.findIndex((r,i) => i > 0 && String(r[0]) === String(username));
  if (index < 1) throw new Error("User not found.");
  if (String(username) === requireSession_(token).username && (role !== "Admin" || active === false)) throw new Error("You cannot deactivate or demote your own administrator account.");
  sh.getRange(index+1,3).setValue(role);
  sh.getRange(index+1,4).setValue(Boolean(active));
  return {success:true,message:"User updated."};
}

function getAlertLog(token) {
  requireRole_(token, ["Admin","Staff"]);
  const rows = getSheet_(SHEETS.ALERTS).getDataRange().getValues().slice(1);
  return rows.slice(-100).reverse().map(r=>({timestamp:formatDateTime_(r[0]),date:formatDate_(r[1]),meal:r[2],waste:Number(r[3])||0,percent:Number(r[4])||0,message:r[5],status:r[6]}));
}

function markAlertResolved(token, timestampText) {
  requireRole_(token, ["Admin","Staff"]);
  const sh = getSheet_(SHEETS.ALERTS);
  const rows = sh.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (formatDateTime_(rows[i][0]) === timestampText) {
      sh.getRange(i+1,7).setValue("Resolved");
      return {success:true};
    }
  }
  throw new Error("Alert not found.");
}

function getSheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error("Missing sheet "+name+". Run setupSystem() from Apps Script first.");
  return sh;
}
function setupGuard_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEETS.USERS)) setupSystem();
}
function requireSession_(token) {
  const value = token ? CacheService.getScriptCache().get("session_"+token) : null;
  if (!value) throw new Error("Your session expired. Please sign in again.");
  const session = JSON.parse(value);
  const user = getSheet_(SHEETS.USERS).getDataRange().getValues().slice(1).find(r=>String(r[0])===session.username);
  if (!user || user[3] !== true || String(user[2]) !== session.role) throw new Error("Your account is inactive or your permissions changed. Sign in again.");
  return session;
}
function requireRole_(token, roles) {
  const s = requireSession_(token);
  if (!roles.includes(s.role)) throw new Error("Your role ("+s.role+") does not have permission for this action.");
  return s;
}
function hashPassword_(password) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return bytes.map(b => (b<0?b+256:b).toString(16).padStart(2,"0")).join("");
}
function readFood_() {
  const rows = getSheet_(SHEETS.FOOD).getDataRange().getValues().slice(1);
  return rows.filter(r=>r[0]).map(r=>({
    date:normalizeDate_(r[0]), meal:String(r[1]||""), prepared:Number(r[2])||0,
    consumed:Number(r[3])||0, waste:Number(r[4])||0, students:Number(r[5])||0,
    costPerKg:Number(r[6])||0, wasteCost:Number(r[7])||0, recordedBy:String(r[8]||"")
  }));
}
function readFeedback_() {
  return getSheet_(SHEETS.FEEDBACK).getDataRange().getValues().slice(1).filter(r=>r[0]).map(r=>({
    timestamp:formatDateTime_(r[0]), meal:String(r[1]||""), rating:Number(r[2])||0,
    comments:String(r[3]||""), username:String(r[4]||"")
  })).filter(r=>r.rating>=1&&r.rating<=5);
}
function summarize_(rows) {
  const t=rows.reduce((a,r)=>{a.prepared+=r.prepared;a.consumed+=r.consumed;a.waste+=r.waste;a.wasteCost+=r.wasteCost;a.records++;return a;},{prepared:0,consumed:0,waste:0,wasteCost:0,records:0});
  t.prepared=round2_(t.prepared);t.consumed=round2_(t.consumed);t.waste=round2_(t.waste);t.wasteCost=round2_(t.wasteCost);
  t.wastePercent=t.prepared?round2_(t.waste/t.prepared*100):0;return t;
}
function listUsers_() {
  return getSheet_(SHEETS.USERS).getDataRange().getValues().slice(1).map(r=>({username:r[0],role:r[2],active:r[3]===true,createdAt:formatDateTime_(r[4])}));
}
function getSetting_(key, fallback) {
  const rows=getSheet_(SHEETS.SETTINGS).getDataRange().getValues();
  const row=rows.slice(1).find(r=>r[0]===key);
  return row ? row[1] : fallback;
}
function sendWasteAlert_(date, meal, waste, percent, target) {
  const email=String(getSetting_("Alert_Email","")||"").trim();
  if (!email || email.indexOf("@")<1) return;
  try {
    MailApp.sendEmail(email, "Hostel food waste alert: "+meal, "Date: "+date+"\nMeal: "+meal+"\nWaste: "+waste+" kg\nWaste percentage: "+round2_(percent)+"%\nTarget: "+target+"%\nPlease review preparation quantity.");
  } catch (err) { console.error("Alert email failed: "+err.message); }
}
function normalizeDate_(value) {
  if (value instanceof Date) return Utilities.formatDate(value,Session.getScriptTimeZone(),"yyyy-MM-dd");
  const s=String(value||"").trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const m=s.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if(m) return m[3]+"-"+m[2]+"-"+m[1];
  const d=new Date(s); if(isNaN(d.getTime())) throw new Error("Invalid date.");
  return Utilities.formatDate(d,Session.getScriptTimeZone(),"yyyy-MM-dd");
}
function dateString_(d) { return Utilities.formatDate(d,Session.getScriptTimeZone(),"yyyy-MM-dd"); }
function formatDate_(value) { try { return normalizeDate_(value); } catch(e) { return String(value||""); } }
function formatDateTime_(value) {
  if(value instanceof Date) return Utilities.formatDate(value,Session.getScriptTimeZone(),"yyyy-MM-dd HH:mm:ss");
  return String(value||"");
}
function round2_(n) { return Math.round((Number(n)||0)*100)/100; }
function roundObject_(o) { Object.keys(o).forEach(k=>{if(typeof o[k]==="number")o[k]=round2_(o[k]);});return o; }
function csvEscape_(v) {
  let s=String(v==null?"":v);
  if (/^[=+\-@]/.test(s)) s="'"+s;
  return '"'+s.replace(/"/g,'""')+'"';
}
