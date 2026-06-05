/**
 * ═══════════════════════════════════════════════════════════
 *  APNA MART — Warehouse Gate Management (Gate In / Gate Out)
 * ═══════════════════════════════════════════════════════════
 */

const SHEET_ID = "1zhXMu3qjcpYG68WVmy0XKYXKf1CRQofqn2K8NL8tvyc";

const TABS = {
  VENDOR:     "Vendor_List_Ex",
  VENDOR_PO:  "Extract_Vendor_PO_List", // daily auto-populated PO list (vendor source)
  TOKEN_LOG:  "Token_Log",
  DASHBOARD:  "Dashboard",
  WH_CONFIG:  "WH_Config",
  COUNTERS:   "WH_Counters",
};

/* ═══════════════════════════════════════════════════════════
   WAREHOUSE ID MAP  ←  EDIT THIS
   The daily PO list (Extract_Vendor_PO_List) identifies warehouses
   by a NUMERIC wh_id (e.g. 1, 2, 10, 52). The Gate In/Out app uses
   LETTER codes: KOL (Kolkata), RNC (Ranchi), RPR (Raipur), PAT (Patna).
   Map every numeric wh_id you use to its app code below.
   ⚠️ The guesses below are based on vendor locations — VERIFY each one.
   Any wh_id NOT listed here is ignored (its vendors won't appear in the app).
═══════════════════════════════════════════════════════════ */
const WH_ID_MAP = {
  "10": "KOL",   // Kolkata  (West Bengal vendors) — verify
  "2":  "RPR",   // Raipur   (Chhattisgarh vendors) — verify
  "1":  "RNC",   // Ranchi (confirmed)
  "52": "PAT"    // Patna (confirmed) — Patna's second wh_id
};

/* ═══════════════════════════════════════════════════════════
   doGet — handles all READ requests
═══════════════════════════════════════════════════════════ */
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";

    if (!action) {
      return HtmlService.createHtmlOutputFromFile("index")
        .setTitle("WH Gate — Apna Mart")
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no");
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    let result;

    switch (action) {
      case "ping":
        result = { success: true, message: "Apps Script is connected!", timestamp: new Date().toISOString() };
        break;
      case "getVendors":
        result = getVendors(ss);
        break;
      case "getTokens":
        result = getTokens(ss, e.parameter.wh || "");
        break;
      case "getAllTokens":
        result = getTokens(ss, "");
        break;
      case "setup":
        setupAllSheets();
        result = { success: true, message: "All sheets created!" };
        break;
      default:
        result = { success: false, error: "Unknown action: '" + action + "'. Valid: ping, getVendors, getTokens, getAllTokens" };
    }

    return buildResponse(result);

  } catch (err) {
    return buildResponse({ success: false, error: err.toString(), stack: err.stack });
  }
}

/* ═══════════════════════════════════════════════════════════
   doPost — handles all WRITE requests
   Body: { "action": "gateIn", ... }  OR  { "action": "gateOut", ... }
═══════════════════════════════════════════════════════════ */
function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const raw = e.postData ? e.postData.contents : "{}";
    const data = JSON.parse(raw);
    const action = data.action || "";

    let result;
    switch (action) {
      case "gateIn":
        result = saveGateIn(ss, data);
        break;
      case "gateOut":
        result = saveGateOut(ss, data);
        break;
      default:
        result = { success: false, error: "Unknown POST action: " + action + ". Valid: gateIn, gateOut" };
    }

    return buildResponse(result);

  } catch (err) {
    return buildResponse({ success: false, error: err.toString() });
  }
}

/* ═══════════════════════════════════════════════════════════
   RESPONSE BUILDER
═══════════════════════════════════════════════════════════ */
function buildResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

/* ═══════════════════════════════════════════════════════════
   READ: Get Vendors  (auto-built from Extract_Vendor_PO_List)
   PO list columns: A purchase_order_id | B wh_id | C full_name | D vendor_name
   Builds { CODE: [unique vendor names, sorted] } using WH_ID_MAP above,
   so the app's vendor dropdown reflects the daily list automatically.
═══════════════════════════════════════════════════════════ */
function getVendors(ss) {
  const sh = ss.getSheetByName(TABS.VENDOR_PO);
  if (!sh) return { success: false, error: "Extract_Vendor_PO_List sheet not found." };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { success: true, vendorMap: {}, warehouseList: [] };

  const rows = sh.getRange(2, 1, lastRow - 1, 4).getValues(); // 4 cols: PO id, wh_id, full_name, vendor_name
  const vendorMap = {};

  rows.forEach(function(row) {
    const whId   = (row[1] || "").toString().trim();   // Col B - wh_id (numeric)
    const vendor = (row[3] || "").toString().trim();   // Col D - vendor_name
    const code   = WH_ID_MAP[whId];                    // numeric id -> app code
    if (code && vendor) {
      if (!vendorMap[code]) vendorMap[code] = [];
      if (vendorMap[code].indexOf(vendor) === -1) vendorMap[code].push(vendor);
    }
  });

  // Sort each warehouse's vendor list alphabetically for the dropdown
  Object.keys(vendorMap).forEach(function(code) { vendorMap[code].sort(); });

  return { success: true, vendorMap: vendorMap, warehouseList: Object.keys(vendorMap).sort() };
}

/* ═══════════════════════════════════════════════════════════
   READ: Get Tokens
   Sheet columns: Token | Vehicle | Vendor | Mobile |  WH |
                  Gate In | Gate Out | Status | TAT | Date
═══════════════════════════════════════════════════════════ */
function getTokens(ss, filterWH) {
  const sh = ss.getSheetByName(TABS.TOKEN_LOG);
  if (!sh) return { success: false, error: "Token_Log sheet not found. Run setupAllSheets() first." };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { success: true, tokens: [] };

  const rows = sh.getRange(2, 1, lastRow - 1, 10).getValues(); // 10 cols, no dock
  var tokens = rows
    .filter(function(r) { return r[0] && r[0].toString().trim() !== ""; })
    .map(function(r) {
      return {
        token:   r[0].toString(),
        vehicle: r[1].toString(),
        vendor:  r[2].toString(),
        mobile:  r[3].toString(),
        wh:      r[4].toString(),
        gateIn:  r[5] ? r[5].toString() : null,
        gateOut: r[6] ? r[6].toString() : null,
        status:  r[7].toString(),
        tat:     r[8] ? r[8].toString() : "",
        date:    r[9] ? r[9].toString() : "",
      };
    });

  if (filterWH) {
    tokens = tokens.filter(function(t) { return t.wh === filterWH; });
  }

  return { success: true, tokens: tokens };
}

/* ═══════════════════════════════════════════════════════════
   FAST TOKEN COUNTER
   WH_Counters sheet: Col A = WH code, Col B = last count
   Avoids scanning all Token_Log rows — single cell read + write
═══════════════════════════════════════════════════════════ */
function getNextTokenNumber(ss, whCode) {
  let sh = ss.getSheetByName(TABS.COUNTERS);
  if (!sh) {  // self-heal: create the counter tab on the fly if it's missing
    sh = ss.insertSheet(TABS.COUNTERS);
    sh.getRange(1, 1, 1, 2).setValues([["WH Code", "Last Token #"]]);
  }

  const lastRow = sh.getLastRow();
  const wh = whCode.toUpperCase();

  if (lastRow >= 2) {
    const rows = sh.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][0].toString().trim().toUpperCase() === wh) {
        const nextNum = (parseInt(rows[i][1]) || 0) + 1;
        sh.getRange(i + 2, 2).setValue(nextNum); // single cell write — fast
        return nextNum;
      }
    }
  }

  // WH not yet in counters — add it
  sh.appendRow([wh, 1]);
  return 1;
}

/* ═══════════════════════════════════════════════════════════
   WRITE: Gate In
═══════════════════════════════════════════════════════════ */
function saveGateIn(ss, data) {
  const sh = ss.getSheetByName(TABS.TOKEN_LOG);
  if (!sh) return { success: false, error: "Token_Log sheet not found" };

  const wh = (data.wh || "").toString().trim().toUpperCase();
  if (!wh) return { success: false, error: "Warehouse (wh) is required" };

  const now      = new Date();
  const ist      = Utilities.formatDate(now, "Asia/Kolkata", "dd-MMM-yyyy HH:mm:ss");
  const dateOnly = Utilities.formatDate(now, "Asia/Kolkata", "dd-MMM-yyyy");

  const tokenNum = getNextTokenNumber(ss, wh); // fast counter
  const token    = wh + "-" + String(tokenNum).padStart(4, "0");

  sh.appendRow([
    token,               // A - Token
    data.vehicle || "",  // B - Vehicle No
    data.vendor  || "",  // C - Vendor Name
    data.mobile  || "",  // D - Mobile Number
    wh,                  // E - Warehouse
    ist,                 // F - Gate In Time
    "",                  // G - Gate Out Time
    "ACTIVE",            // H - Status
    "",                  // I - TAT (mins)
    dateOnly             // J - Date
  ]);

  return { success: true, token: token, gateIn: ist, wh: wh };
}

/* ═══════════════════════════════════════════════════════════
   WRITE: Gate Out
═══════════════════════════════════════════════════════════ */
function saveGateOut(ss, data) {
  const sh = ss.getSheetByName(TABS.TOKEN_LOG);
  if (!sh) return { success: false, error: "Token_Log sheet not found" };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { success: false, error: "No tokens found" };

  // Single batch read of token column
  const tokenCol = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  let rowIndex = -1;
  const searchToken = (data.token || "").toString().trim();

  for (var i = 0; i < tokenCol.length; i++) {
    if (tokenCol[i][0].toString().trim() === searchToken) {
      rowIndex = i + 2;
      break;
    }
  }

  if (rowIndex === -1) return { success: false, error: "Token not found: " + searchToken };

  const gateInVal = sh.getRange(rowIndex, 6).getValue();
  const now       = new Date();
  const ist       = Utilities.formatDate(now, "Asia/Kolkata", "dd-MMM-yyyy HH:mm:ss");
  let tatMins     = "";

  if (gateInVal) {
    try { tatMins = Math.round((now - new Date(gateInVal)) / 60000); } catch(e) {}
  }

  // Write Gate Out + Status + TAT in one call (cols G, H, I)
  sh.getRange(rowIndex, 7, 1, 3).setValues([[ist, "CLOSED", tatMins]]);

  return { success: true, token: searchToken, gateOut: ist, tatMins: tatMins };
}

/* ═══════════════════════════════════════════════════════════
   SETUP — Run once to create all sheets
═══════════════════════════════════════════════════════════ */
function setupAllSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ss.rename("WH Gate — Apna Mart");

  _setupVendorSheet(ss);
  _setupTokenLogSheet(ss);
  _setupCountersSheet(ss);
  _setupDashboardSheet(ss);
  _setupWHConfigSheet(ss);

  ["Sheet1", "Sheet 1"].forEach(function(name) {
    const s = ss.getSheetByName(name);
    if (s && ss.getSheets().length > 1) {
      try { ss.deleteSheet(s); } catch(e) {}
    }
  });

  SpreadsheetApp.flush();
  Logger.log("✅ All sheets created. Now deploy as Web App.");
}

function _setupVendorSheet(ss) {
  let sh = ss.getSheetByName(TABS.VENDOR);
  if (!sh) sh = ss.insertSheet(TABS.VENDOR);
  else sh.clear();
  sh.setTabColor("#3B82F6");

  const headers = ["Sr No","Vendor Code","Vendor Type","Category","Sub Category","Region","Zone","State","City","Address","Pin Code","Contact Name","Contact Phone","Warehouse","Vendor Name","WH ID"];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground("#1D4ED8").setFontColor("#FFFFFF").setFontWeight("bold");

  const data = [
    [1,"VND001","Regular","FMCG","Beverages","East","WB","West Bengal","Kolkata","12 Park Street","700016","Ramesh Kumar","9800012345","Kolkata Warehouse","Bengal Beverages Pvt Ltd","KOL"],
    [2,"VND002","Regular","FMCG","Food","East","WB","West Bengal","Kolkata","45 Salt Lake","700064","Suresh Das","9800023456","Kolkata Warehouse","Maa Tara Traders","KOL"],
    [3,"VND003","Premium","Dairy","Milk","East","WB","West Bengal","Kolkata","8 Gariahat Rd","700029","Priya Roy","9800034567","Kolkata Warehouse","Eastern Supplies Ltd","KOL"],
    [4,"VND004","Regular","FMCG","Snacks","East","JH","Jharkhand","Ranchi","22 Main Rd","834001","Amit Singh","9800045678","Ranchi Warehouse","Aaditya Enterprises","RNC"],
    [5,"VND005","Regular","FMCG","Beverages","East","JH","Jharkhand","Ranchi","5 Station Rd","834001","Vikas Jha","9800056789","Ranchi Warehouse","Jharkhand Traders","RNC"],
    [6,"VND006","Regular","Grocery","Rice","Central","CG","Chhattisgarh","Raipur","10 Civil Lines","492001","Deepak Sahu","9800067890","Raipur Warehouse","Raipur Traders Ltd","RPR"],
    [7,"VND007","Premium","Grocery","Pulses","Central","CG","Chhattisgarh","Raipur","33 Pandri","492004","Sunita Verma","9800078901","Raipur Warehouse","Chhattisgarh Foods","RPR"],
    [8,"VND008","Regular","FMCG","Beverages","South","KA","Karnataka","Bengaluru","15 MG Road","560001","Kiran Rao","9800089012","Bengaluru Warehouse","Aaditya Enterprises","BLR"],
    [9,"VND009","Regular","FMCG","Food","South","KA","Karnataka","Bengaluru","88 Indiranagar","560038","Meena Reddy","9800090123","Bengaluru Warehouse","Bangalore Agro Foods","BLR"],
    [10,"VND010","Premium","Dairy","Curd","South","KA","Karnataka","Bengaluru","3 Koramangala","560034","Rajesh Nair","9800001234","Bengaluru Warehouse","Deccan Distributors","BLR"],
    [11,"VND011","Regular","FMCG","Snacks","South","TS","Telangana","Hyderabad","45 Banjara Hills","500034","Arjun Kumar","9811112222","Hyderabad Warehouse","Hyderabad Foods Ltd","HYD"],
    [12,"VND012","Regular","Grocery","Rice","South","TS","Telangana","Hyderabad","12 Jubilee Hills","500033","Sita Devi","9811123333","Hyderabad Warehouse","Pearl City Traders","HYD"],
    [13,"VND013","Regular","FMCG","Beverages","North","DL","Delhi","New Delhi","8 Connaught Place","110001","Rohit Sharma","9811134444","Delhi Warehouse","Delhi Traders Co","DEL"],
    [14,"VND014","Premium","FMCG","Food","North","DL","Delhi","New Delhi","22 Lajpat Nagar","110024","Anita Gupta","9811145555","Delhi Warehouse","Capital Foods Ltd","DEL"],
  ];
  if (data.length > 0) sh.getRange(2, 1, data.length, headers.length).setValues(data);
  sh.setFrozenRows(1);
  sh.setColumnWidth(15, 220);
  sh.setColumnWidth(16, 70);
  Logger.log("✅ Vendor_List_Ex ready");
}

/* ── Token Log: 10 columns, Dock columns removed ── */
function _setupTokenLogSheet(ss) {
  let sh = ss.getSheetByName(TABS.TOKEN_LOG);
  if (!sh) sh = ss.insertSheet(TABS.TOKEN_LOG);
  else sh.clear();
  sh.setTabColor("#10B981");

  const headers = [
    "Token",       // A
    "Vehicle No",  // B
    "Vendor Name", // C
    "Mobile Number", // D
    "Warehouse",   // E
    "Gate In Time",// F
    "Gate Out Time",// G
    "Status",      // H
    "TAT (mins)",  // I
    "Date"         // J
  ];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground("#065F46").setFontColor("#FFFFFF").setFontWeight("bold");

  sh.setColumnWidths(1, 10, [110, 120, 200, 140, 80, 160, 160, 80, 85, 110]);
  sh.setFrozenRows(1);

  // Status conditional formatting (col H)
  const statusRange = sh.getRange("H2:H1000");
  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("ACTIVE").setBackground("#D1FAE5").setFontColor("#065F46")
      .setRanges([statusRange]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("CLOSED").setBackground("#F1F5F9").setFontColor("#475569")
      .setRanges([statusRange]).build(),
  ];
  sh.setConditionalFormatRules(rules);
  Logger.log("✅ Token_Log ready (Gate In / Gate Out only)");
}

/* ── WH_Counters: fast token number generation ── */
function _setupCountersSheet(ss) {
  let sh = ss.getSheetByName(TABS.COUNTERS);
  if (!sh) sh = ss.insertSheet(TABS.COUNTERS);
  else sh.clear();
  sh.setTabColor("#F59E0B");

  sh.getRange(1, 1, 1, 2).setValues([["WH Code", "Last Token #"]])
    .setBackground("#92400E").setFontColor("#fff").setFontWeight("bold");

  // Pre-seed all warehouses at 0
  const whs = [["KOL",0],["RNC",0],["RPR",0],["PAT",0]];
  sh.getRange(2, 1, whs.length, 2).setValues(whs);
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, 2, [100, 120]);
  Logger.log("✅ WH_Counters ready");
}

/* ── Dashboard: Gate In / Gate Out only ── */
function _setupDashboardSheet(ss) {
  let sh = ss.getSheetByName(TABS.DASHBOARD);
  if (!sh) sh = ss.insertSheet(TABS.DASHBOARD);
  else sh.clear();
  sh.setTabColor("#F59E0B");

  sh.getRange("A1").setValue("📊 APNA MART — WH GATE LIVE DASHBOARD");
  sh.getRange("A1:H1").merge()
    .setFontSize(14).setFontWeight("bold")
    .setBackground("#1A2236").setFontColor("#F1F5F9");

  // Summary
  sh.getRange("A3").setValue("SUMMARY").setFontWeight("bold");
  const summaryLabels = [["Total Tokens"],["Active (Gate In)"],["Closed (Gate Out)"],["Today's Tokens"]];
  sh.getRange("A4:A7").setValues(summaryLabels).setFontWeight("bold");
  sh.getRange("B4").setFormula('=COUNTA(Token_Log!A:A)-1');
  sh.getRange("B5").setFormula('=COUNTIF(Token_Log!H:H,"ACTIVE")');
  sh.getRange("B6").setFormula('=COUNTIF(Token_Log!H:H,"CLOSED")');
  sh.getRange("B7").setFormula('=COUNTIFS(Token_Log!J:J,TEXT(TODAY(),"dd-mmm-yyyy"))');

  // Per WH
  sh.getRange("A9").setValue("PER WAREHOUSE").setFontWeight("bold");
  sh.getRange("A10:F10")
    .setValues([["WH", "Total", "Active", "Closed", "Avg TAT (mins)", "Today"]])
    .setFontWeight("bold").setBackground("#1D4ED8").setFontColor("#fff");

  const whs = ["KOL","RNC","RPR","PAT"];
  whs.forEach(function(wh, i) {
    const r = 11 + i;
    sh.getRange(r, 1).setValue(wh);
    sh.getRange(r, 2).setFormula('=COUNTIF(Token_Log!E:E,"' + wh + '")');
    sh.getRange(r, 3).setFormula('=COUNTIFS(Token_Log!E:E,"' + wh + '",Token_Log!H:H,"ACTIVE")');
    sh.getRange(r, 4).setFormula('=COUNTIFS(Token_Log!E:E,"' + wh + '",Token_Log!H:H,"CLOSED")');
    sh.getRange(r, 5).setFormula('=IFERROR(AVERAGEIFS(Token_Log!I:I,Token_Log!E:E,"' + wh + '"),"-")');
    sh.getRange(r, 6).setFormula('=COUNTIFS(Token_Log!E:E,"' + wh + '",Token_Log!J:J,TEXT(TODAY(),"dd-mmm-yyyy"))');
  });

  sh.setFrozenRows(1);
  Logger.log("✅ Dashboard ready");
}

function _setupWHConfigSheet(ss) {
  let sh = ss.getSheetByName(TABS.WH_CONFIG);
  if (!sh) sh = ss.insertSheet(TABS.WH_CONFIG);
  else sh.clear();
  sh.setTabColor("#8B5CF6");

  const headers = ["WH Code","WH Name","City","State","Address","Capacity","Manager","Phone","Active"];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground("#6D28D9").setFontColor("#fff").setFontWeight("bold");

  const data = [
    ["KOL","Kolkata Warehouse","Kolkata","West Bengal","12 Industrial Area, Kolkata",500,"Ramesh Ghosh","9800011111","YES"],
    ["RNC","Ranchi Warehouse","Ranchi","Jharkhand","45 Ring Road, Ranchi",300,"Suresh Sinha","9800022222","YES"],
    ["RPR","Raipur Warehouse","Raipur","Chhattisgarh","8 NH-6, Raipur",400,"Deepak Tiwari","9800033333","YES"],
    ["PAT","Patna Warehouse","Patna","Bihar","Patna",400,"","","YES"],
  ];
  sh.getRange(2, 1, data.length, headers.length).setValues(data);
  sh.setFrozenRows(1);
  Logger.log("✅ WH_Config ready");
}

/* ═══════════════════════════════════════════════════════════
   TEST FUNCTION — Run in editor, check Logs (Ctrl+Enter)
═══════════════════════════════════════════════════════════ */
function testScript() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  Logger.log("=== TESTING SCRIPT ===");

  const v = getVendors(ss);
  Logger.log("getVendors: " + v.success + " | WHs: " + JSON.stringify(Object.keys(v.vendorMap || {})));

  const t = getTokens(ss, "");
  Logger.log("getTokens: " + t.success + " | count: " + (t.tokens ? t.tokens.length : 0));

  const gi = saveGateIn(ss, { vehicle: "KA01AB9999", vendor: "Test Vendor", mobile: "9876543210", wh: "PAT" });
  Logger.log("gateIn: " + gi.success + " | token: " + gi.token);

  if (gi.success) {
    const go = saveGateOut(ss, { token: gi.token });
    Logger.log("gateOut: " + go.success + " | TAT: " + go.tatMins + " mins");
  }

  Logger.log("=== ALL TESTS DONE ===");
}