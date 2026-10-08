// ══════════════════════════════════════════════════════
//  Project Nazuna — Leave Application Command Center
//  Code.gs  v8.0 — Indexed, Stateless, Rate-Limited
// ══════════════════════════════════════════════════════

var SHEET_NAME  = "Leave Results V2.0";
var CONFIG_NAME = "FormConfig";
var SUP_NAME    = "Supervisors";
var LOG_NAME    = "System Log";

// ── INCREMENTAL SYNC CACHE ────────────────────────────
// Stores the last sheet modification time so getEventsOnly()
// can skip a full read when nothing has changed.
var _cachedEvents   = null;
var _cacheTimestamp = 0;   // epoch ms of the last sheet edit we processed

// ── STRUCTURED LOGGER ────────────────────────────────
function _log(level, fn, msg, extra) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(LOG_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(LOG_NAME);
      sheet.getRange("A1:F1").setValues([["Timestamp","Level","Function","Actor","Message","Extra"]])
        .setBackground("#202124").setFontColor("#ffffff").setFontWeight("bold");
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 160); sheet.setColumnWidth(3, 140);
      sheet.setColumnWidth(4, 180); sheet.setColumnWidth(5, 300); sheet.setColumnWidth(6, 200);
    }
    var actor = "";
    try { actor = Session.getActiveUser().getEmail(); } catch(e) {}
    var tz = Session.getScriptTimeZone();
    sheet.appendRow([
      Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss"),
      level, fn, actor, msg,
      extra ? JSON.stringify(extra) : ""
    ]);
    Logger.log("[" + level + "] " + fn + " — " + msg);
  } catch(e) { Logger.log("LOG_FAIL: " + e.message); }
}

// ── SAFE CACHE WRITE ──────────────────────────────────
// CacheService caps values at 100KB. If the payload is too big, cache.put
// throws — and if that happens silently inside a bigger try/catch, caching
// effectively stops working forever with no visible error. This guards
// against that and just skips caching (rather than crashing) when oversized.
function _safeCachePut(key, dataArray, ttlSeconds) {
  try {
    var json = JSON.stringify(dataArray);
    if (json.length < 95000) {
      CacheService.getScriptCache().put(key, json, ttlSeconds || 15);
    } else {
      _log("WARN", "_safeCachePut", "Payload too large to cache", { key: key, size: json.length });
    }
  } catch(e) {
    _log("WARN", "_safeCachePut", "Cache put failed", { key: key, err: e.message });
  }
}

// ── RATE LIMITER (uses CacheService) ─────────────────
// key: string identifier, maxCalls: per window, windowSec: window in seconds
function _checkRateLimit(key, maxCalls, windowSec) {
  var cache   = CacheService.getScriptCache();
  var cacheKey = "rl_" + key;
  var raw     = cache.get(cacheKey);
  var count   = raw ? parseInt(raw) : 0;
  if (count >= maxCalls) {
    _log("WARN", "_checkRateLimit", "Rate limit hit", { key: key, count: count });
    return false;
  }
  cache.put(cacheKey, String(count + 1), windowSec);
  return true;
}

// ── ROW INDEX: stable UUID-based row lookup ───────────
// Ensures column "Row UUID" exists and backfills missing UUIDs.
// Returns a map of { uuid -> 1-based row number } for O(1) lookup.
function _getRowIndex(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var uuidColIdx = headers.findIndex(function(h) {
    return h.toString().replace(/[^a-z0-9]/gi,'').toLowerCase() === "rowuuid";
  });

  // Create the column if missing
  if (uuidColIdx === -1) {
    uuidColIdx = headers.length;
    sheet.getRange(1, uuidColIdx + 1).setValue("Row UUID")
      .setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold");
    _log("INFO", "_getRowIndex", "Created Row UUID column", { col: uuidColIdx + 1 });
  }

  var lastRow = sheet.getLastRow();
  var index   = {};
  if (lastRow < 2) return { index: index, uuidCol: uuidColIdx + 1 };

  var uuidRange  = sheet.getRange(2, uuidColIdx + 1, lastRow - 1, 1);
  var uuidValues = uuidRange.getValues();
  var toWrite    = [];
  var hasGaps    = false;

  for (var i = 0; i < uuidValues.length; i++) {
    var existing = (uuidValues[i][0] || "").toString().trim();
    if (!existing) {
      existing = Utilities.getUuid();
      uuidValues[i][0] = existing;
      hasGaps = true;
    }
    index[existing] = i + 2; // 1-based row number
  }

  if (hasGaps) {
    uuidRange.setValues(uuidValues);
    _log("INFO", "_getRowIndex", "Backfilled missing UUIDs", { count: Object.keys(index).length });
  }

  return { index: index, uuidCol: uuidColIdx + 1 };
}

// ── SHEET HELPERS ─────────────────────────────────────
function doGet() {
  return HtmlService.createHtmlOutputFromFile('VLIndex')
    .setTitle('Google Play VL Calendar')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getSessionEmail() {
  try { return Session.getActiveUser().getEmail(); } catch(e) { return ""; }
}

function getOrCreateConfigSheet() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName(CONFIG_NAME);
  if (!cfg) {
    cfg = ss.insertSheet(CONFIG_NAME);
    cfg.getRange("A1:B1").setValues([["Setting","Value"]])
      .setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold");
    cfg.getRange("A2:B4").setValues([
      ["IsOpen",       "TRUE"],
      ["ActiveMonths", "May 2026, June 2026"],
      ["ClosedMsg",    "The VL filing period is currently closed."]
    ]);
    cfg.setColumnWidth(1, 160); cfg.setColumnWidth(2, 420); cfg.setFrozenRows(1);
    _log("INFO", "getOrCreateConfigSheet", "Created FormConfig sheet");
  }
  return cfg;
}

function getOrCreateSupervisorSheet() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var sup = ss.getSheetByName(SUP_NAME);
  if (!sup) {
    sup = ss.insertSheet(SUP_NAME);
    sup.getRange("A1:C1").setValues([["Supervisor Email","Team Name","Role"]])
      .setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold");
    sup.getRange("A2:C2").setValues([["stevenjosephc@google.com","Team Steven","Admin"]]);
    sup.setColumnWidth(1, 250); sup.setColumnWidth(2, 150); sup.setColumnWidth(3, 100);
    sup.setFrozenRows(1);
    _log("INFO", "getOrCreateSupervisorSheet", "Created Supervisors sheet");
  }
  return sup;
}

function getFormConfig() {
  var res = { isOpen: true, activeMonths: ["May 2026", "June 2026"], closedMsg: "Form is closed.", sunsetMode: false, isSupervisor: false, isAdmin: false, teamName: "" };

  // Config/Supervisor sheets rarely change — cache the non-personalized
  // part (isOpen/activeMonths/closedMsg) for 30s so this doesn't do 2
  // sheet reads on every single dashboard load from every user.
  var cache      = CacheService.getScriptCache();
  var baseCached = cache.get("vl_config_base");
  if (baseCached) {
    var base = JSON.parse(baseCached);
    res.isOpen = base.isOpen; res.activeMonths = base.activeMonths; res.closedMsg = base.closedMsg; res.sunsetMode = base.sunsetMode;
  } else {
  try {
    var data = getOrCreateConfigSheet().getRange("A2:B10").getValues();
    data.forEach(function(r) {
      var k = (r[0]||"").toString().trim();
      var v = (r[1]||"").toString().trim();
      if (!k) return;
      if (k === "IsOpen")       res.isOpen       = (v.toUpperCase() === "TRUE");
      if (k === "ActiveMonths") res.activeMonths = v.split(",").map(function(m){ return m.trim(); }).filter(Boolean);
      if (k === "ClosedMsg")    res.closedMsg    = v;
      if (k === "SunsetMode")   res.sunsetMode   = (v.toUpperCase() === "TRUE");
    });
    cache.put("vl_config_base", JSON.stringify({ isOpen: res.isOpen, activeMonths: res.activeMonths, closedMsg: res.closedMsg, sunsetMode: res.sunsetMode }), 30);
  } catch(e) { _log("ERROR", "getFormConfig", "Config read failed", { err: e.message }); }
  }

  try {
    var email = getSessionEmail().toLowerCase();
    if (email) {
      var supCacheKey = "vl_sup_" + email;
      var supCached   = cache.get(supCacheKey);
      if (supCached) {
        var supRole = JSON.parse(supCached);
        res.isSupervisor = supRole.isSupervisor;
        res.teamName     = supRole.teamName;
        res.isAdmin      = supRole.isAdmin;
      } else {
        var supData = getOrCreateSupervisorSheet().getDataRange().getValues();
        for (var i = 1; i < supData.length; i++) {
          if ((supData[i][0]||"").toString().trim().toLowerCase() === email) {
            res.isSupervisor = true;
            res.teamName     = (supData[i][1]||"").toString().trim();
            res.isAdmin      = (supData[i][2]||"").toString().trim().toLowerCase() === "admin";
            break;
          }
        }
        _safeCachePut(supCacheKey, { isSupervisor: res.isSupervisor, teamName: res.teamName, isAdmin: res.isAdmin }, 30);
      }
    }
  } catch(e) { _log("ERROR", "getFormConfig", "Supervisor lookup failed", { err: e.message }); }

  return res;
}

function getVLSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    var h = ["Timestamp","Email Address","LDAP","Channel","VL Date","Team Lead",
             "Reason for VL","Work Group","Site","Accruals Snip-it","Accruals",
             "Month","Date of Birthday","Proof / Artifacts","Status","Comments",
             "Attendance","Confirmation on Status","Email Sent","Row UUID"];
    sheet.getRange(1,1,1,h.length).setValues([h])
      .setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
    _log("INFO", "getVLSheet", "Created " + SHEET_NAME + " sheet");
  }
  return sheet;
}

function ensureColumn(sheet, headerName) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var idx = headers.findIndex(function(h) {
    return h.toString().toLowerCase().replace(/[^a-z0-9]/g,'') === headerName.toLowerCase().replace(/[^a-z0-9]/g,'');
  });
  if (idx === -1) {
    sheet.getRange(1, headers.length + 1).setValue(headerName)
      .setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold");
    _log("INFO", "ensureColumn", "Added column: " + headerName);
    return headers.length;
  }
  return idx;
}

// ── CALENDAR DATA (fixed closure bug + UUID IDs) ──────
function getCalendarData() {
  try {
    var sheet   = getVLSheet();
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return [];

    // Ensure UUID index exists and is backfilled
    var rowIdx  = _getRowIndex(sheet);
    lastCol     = sheet.getLastColumn(); // refresh after possible UUID column add

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colMap  = {};
    headers.forEach(function(h, idx) {
      colMap[h.toString().replace(/[^a-zA-Z0-9]/g,'').toLowerCase()] = idx;
    });

    // Re-read lastRow in case UUID backfill changed row count
    lastRow = sheet.getLastRow();
    var data        = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var tz          = Session.getScriptTimeZone();
    var events      = [];

    // Pre-resolve column indices for better performance
    var cTs        = colMap["timestamp"];
    var cEmail     = colMap["emailaddress"];
    var cLdap      = colMap["ldap"];
    var cChannel   = colMap["channel"];
    var cVlDate    = colMap["vldate"];
    var cTeamLead  = colMap["teamlead"];
    var cReason    = colMap["reasonforvl"];
    var cWorkGroup = colMap["workgroup"];
    var cSite      = colMap["site"];
    var cAccProof  = colMap["accrualssnipit"];
    var cAccruals  = colMap["accruals"];
    var cMonth     = colMap["month"];
    var cStatus    = colMap["status"];
    var cComments  = colMap["comments"];
    var cRemarks   = colMap["remarks"]; // fallback for comments
    var cAttendance= colMap["attendance"];
    var cConf      = colMap["confirmationonstatus"];
    var cEmailSent = colMap["emailsent"];
    var cProof     = colMap["proofartifacts"];
    var cUuid      = colMap["rowuuid"];

    for (var i = 0; i < data.length; i++) {
      var row = data[i];

      var rawTs   = cTs !== undefined && row[cTs] != null ? row[cTs] : "";
      var rawLdap = cLdap !== undefined && row[cLdap] != null ? row[cLdap].toString().trim() : "";
      if (!rawTs || !rawLdap) continue;

      // Format VL date
      var formattedDate = "";
      var rawDate = cVlDate !== undefined && row[cVlDate] != null ? row[cVlDate] : "";
      if (Object.prototype.toString.call(rawDate) === '[object Date]' && !isNaN(rawDate.getTime())) {
        formattedDate = Utilities.formatDate(rawDate, tz, "yyyy-MM-dd");
      } else if (rawDate) {
        var rdStr = rawDate.toString().trim();
        if (/^\d{4}-\d{2}-\d{2}/.test(rdStr)) {
          // Already a plain yyyy-MM-dd string (submitted via the web form).
          // Do NOT round-trip through new Date() — it parses date-only
          // strings as UTC and can shift the day by -1 depending on the
          // script's project timezone. This was the cause of dates
          // appearing on the wrong day on the calendar.
          formattedDate = rdStr.substring(0, 10);
        } else {
          var parsed = new Date(rdStr + "T00:00:00"); // force LOCAL parsing
          formattedDate = !isNaN(parsed.getTime())
            ? Utilities.formatDate(parsed, tz, "yyyy-MM-dd")
            : rdStr;
        }
      }

      // Format timestamp
      var tsStr = "Unknown";
      if (Object.prototype.toString.call(rawTs) === '[object Date]' && !isNaN(rawTs.getTime())) {
        tsStr = Utilities.formatDate(rawTs, tz, "MMM d, yyyy h:mm a");
      } else {
        tsStr = rawTs.toString();
      }
      var tsEpoch = 0;
      if (Object.prototype.toString.call(rawTs) === '[object Date]' && !isNaN(rawTs.getTime())) {
        tsEpoch = rawTs.getTime();
      } else {
        var parsedTs = new Date(rawTs);
        tsEpoch = !isNaN(parsedTs.getTime()) ? parsedTs.getTime() : 0;
      }

      // Use stable UUID as event ID
      var uuid = cUuid !== undefined && row[cUuid] != null ? row[cUuid].toString().trim() : "";
      if (!uuid) uuid = "ev_" + i; // fallback for rows not yet indexed

      // Manually format attendance and accruals
      var rawAccruals = cAccruals !== undefined && row[cAccruals] != null ? row[cAccruals] : "";
      var accrualsStr = rawAccruals.toString().trim();

      var rawAttendance = cAttendance !== undefined && row[cAttendance] != null ? row[cAttendance] : "";
      var attendanceStr = rawAttendance.toString().trim();
      if (typeof rawAttendance === 'number') {
        attendanceStr = (rawAttendance * 100).toFixed(2) + "%";
      }

      var comments = cComments !== undefined && row[cComments] != null ? row[cComments].toString().trim() : "";
      var remarks  = cRemarks !== undefined && row[cRemarks] != null ? row[cRemarks].toString().trim() : "";

      events.push({
        id           : uuid,
        rowNum       : i + 2,
        timestamp    : tsStr,
        tsEpoch      : tsEpoch,
        email        : cEmail !== undefined && row[cEmail] != null ? row[cEmail].toString().trim().toLowerCase() : "",
        ldap         : rawLdap,
        channel      : cChannel !== undefined && row[cChannel] != null ? row[cChannel].toString().trim() : "",
        date         : formattedDate,
        teamLead     : cTeamLead !== undefined && row[cTeamLead] != null ? row[cTeamLead].toString().trim() : "",
        reason       : cReason !== undefined && row[cReason] != null ? row[cReason].toString().trim() : "",
        workGroup    : cWorkGroup !== undefined && row[cWorkGroup] != null ? row[cWorkGroup].toString().trim() : "",
        site         : cSite !== undefined && row[cSite] != null ? row[cSite].toString().trim() : "",
        accruals     : accrualsStr,
        accrualsProof: cAccProof !== undefined && row[cAccProof] != null ? row[cAccProof].toString().trim() : "",
        month        : cMonth !== undefined && row[cMonth] != null ? row[cMonth].toString().trim() : "",
        status       : cStatus !== undefined && row[cStatus] != null && row[cStatus].toString().trim() ? row[cStatus].toString().trim() : "Pending",
        remarks      : comments || remarks,
        attendance   : attendanceStr,
        confirmation : cConf !== undefined && row[cConf] != null ? row[cConf].toString().trim() : "",
        emailSent    : cEmailSent !== undefined && row[cEmailSent] != null ? row[cEmailSent].toString().trim() : "",
        proof        : cProof !== undefined && row[cProof] != null ? row[cProof].toString().trim() : ""
      });
    }

    // NOTE: removed the per-call INFO log here — this function runs on
    // every dashboard load AND every cache-miss poll, and _log() does an
    // appendRow() to a Log sheet each time. Under load that turned every
    // read into an extra sheet write, compounding the concurrency problem.
    return events;

  } catch(e) {
    _log("ERROR", "getCalendarData", e.message);
    return [];
  }
}

// ── STATUS UPDATE (UUID-indexed, rate-limited, logged) ─
function updateRequestStatus(eventUuid, oldStatus, newStatus, adminComment) {
  var actor = getSessionEmail();

  // Rate limit: max 30 status changes per minute per script instance
  if (!_checkRateLimit("statusUpdate_" + actor, 30, 60)) {
    return { success: false, message: "Rate limit exceeded. Please wait before making more changes." };
  }

  try {
    var sheet   = getVLSheet();
    var rowIdx  = _getRowIndex(sheet);
    var rowNum  = rowIdx.index[eventUuid];

    if (!rowNum) {
      _log("WARN", "updateRequestStatus", "UUID not found in index", { uuid: eventUuid });
      return { success: false, message: "Row not found. The calendar may be out of sync — please refresh." };
    }

    var statusColIdx = ensureColumn(sheet, "Status") + 1;
    var confColIdx   = ensureColumn(sheet, "Confirmation on Status") + 1;
    var commentColIdx= ensureColumn(sheet, "Comments") + 1;

    // Concurrency check: read current status from the actual row
    var currentStatus    = sheet.getRange(rowNum, statusColIdx).getValue().toString().trim();
    var normalizedCurrent = currentStatus === "" ? "Pending" : currentStatus;
    var normalizedOld     = oldStatus === "" ? "Pending" : oldStatus;

    if (normalizedCurrent.toLowerCase() !== normalizedOld.toLowerCase()) {
      _log("WARN", "updateRequestStatus", "Concurrency conflict", {
        uuid: eventUuid, expected: normalizedOld, found: normalizedCurrent
      });
      return {
        success: false,
        message: "Another admin already updated this to '" + currentStatus + "'.",
        newStatus: currentStatus
      };
    }

    var tz    = Session.getScriptTimeZone();
    var stamp = actor.split('@')[0] + " — " + Utilities.formatDate(new Date(), tz, "MM/dd/yy HH:mm");

    sheet.getRange(rowNum, statusColIdx).setValue(newStatus);
    sheet.getRange(rowNum, confColIdx).setValue(stamp);
    if (adminComment && adminComment.trim()) {
      sheet.getRange(rowNum, commentColIdx).setValue(adminComment.trim());
    }

    _log("INFO", "updateRequestStatus", "Status updated", {
      uuid: eventUuid, row: rowNum, from: oldStatus, to: newStatus,
      comment: adminComment || ""
    });

    return { success: true, newStatus: newStatus };

  } catch(e) {
    _log("ERROR", "updateRequestStatus", e.message, { uuid: eventUuid });
    return { success: false, message: e.message };
  }
}

// ── BATCH STATUS UPDATE (optimized for multiple rows) ──
function batchUpdateStatuses(eventUuuids, newStatus, adminComment) {
  var actor = getSessionEmail();
  if (!Array.isArray(eventUuuids) || eventUuuids.length === 0) {
    return { success: false, message: "No requests selected." };
  }

  // Rate limit: batch counts as one action for basic quota, but we log the size
  if (!_checkRateLimit("batchUpdate_" + actor, 10, 60)) {
    return { success: false, message: "Rate limit exceeded. Please wait a moment." };
  }

  try {
    var sheet   = getVLSheet();
    var rowIdx  = _getRowIndex(sheet);
    var tz      = Session.getScriptTimeZone();
    var stamp   = actor.split('@')[0] + " — " + Utilities.formatDate(new Date(), tz, "MM/dd/yy HH:mm") + " (Batch)";

    var statusColIdx  = ensureColumn(sheet, "Status") + 1;
    var confColIdx    = ensureColumn(sheet, "Confirmation on Status") + 1;
    var commentColIdx = ensureColumn(sheet, "Comments") + 1;

    var updatedCount = 0;
    var errors = [];

    // Optimized: group updates to minimize sheet API calls
    var minRow = 999999, maxRow = 0;
    eventUuuids.forEach(function(uuid){
      var r = rowIdx.index[uuid];
      if(r){ if(r < minRow) minRow = r; if(r > maxRow) maxRow = r; }
    });

    if (maxRow > 0) {
      var numRows = maxRow - minRow + 1;
      var lastCol = sheet.getLastColumn();
      var range = sheet.getRange(minRow, 1, numRows, lastCol);
      var values = range.getValues();

      eventUuuids.forEach(function(uuid) {
        var rowNum = rowIdx.index[uuid];
        if (rowNum) {
          var localIdx = rowNum - minRow;
          values[localIdx][statusColIdx - 1] = newStatus;
          values[localIdx][confColIdx - 1]   = stamp;
          if (adminComment && adminComment.trim()) {
            values[localIdx][commentColIdx - 1] = adminComment.trim();
          }
          updatedCount++;
        } else {
          errors.push(uuid + " not found");
        }
      });
      range.setValues(values);
    }

    _log("INFO", "batchUpdateStatuses", "Batch update complete", {
      count: updatedCount, to: newStatus, errorCount: errors.length
    });

    return { success: true, count: updatedCount, errors: errors };

  } catch(e) {
    _log("ERROR", "batchUpdateStatuses", e.message);
    return { success: false, message: e.message };
  }
}

// ── FORM SUBMISSION (rate-limited, logged) ────────────
function processVLForm(data) {
  var actor = getSessionEmail();

  // Rate limit: max 3 submissions per 5 minutes per user
  if (!_checkRateLimit("submit_" + actor, 3, 300)) {
    _log("WARN", "processVLForm", "Submission rate limit hit", { actor: actor });
    return { status: "error", message: "Too many submissions. Please wait a few minutes before trying again." };
  }

  var config = getFormConfig();
  if (!config.isOpen) return { status: "closed", message: config.closedMsg };

  var submittedMonth   = normalizeMonthFull(data.month || "");
  var normalizedActive = config.activeMonths.map(function(m){ return normalizeMonthFull(m); });

  // Validation (no sheet access yet — do this before taking the lock)
  if (!data.ldap || !data.ldap.trim()) return { status:"error", message:"LDAP is required." };
  if (!data.channel)  return { status:"error", message:"Channel is required." };
  if (!data.vlDate)   return { status:"error", message:"VL Date is required." };
  if (data.channel === "Phone") {
    var day = new Date(data.vlDate + "T00:00:00").getDay();
    if (day === 0 || day === 6) return { status:"error", message:"Phone agents cannot file leave on weekends." };
  }
  if (normalizedActive.indexOf(submittedMonth) < 0) {
    return { status:"error", message:"Selected month (" + submittedMonth + ") is no longer active." };
  }
  if (data.reasonForVL === "Birthday" && !data.bdayDate) {
    return { status:"error", message:"Birthday Leave requires a Date of Birthday." };
  }
  var vlDateParsed = new Date(data.vlDate + "T00:00:00");
  var vlMonthFull  = MO_NAMES_GS[vlDateParsed.getMonth()] + " " + vlDateParsed.getFullYear();
  if (vlMonthFull !== submittedMonth) {
    _log("WARN", "processVLForm", "Date/month mismatch", { vlDate: data.vlDate, declaredMonth: submittedMonth });
    return { status:"error", message:"Your VL Date (" + data.vlDate + ") does not fall within the declared month (" + submittedMonth + ")." };
  }

  // ── LOCK: only the duplicate-check + write needs to be serialized ──
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // wait up to 30s for the lock, then give up

    var sheet = getVLSheet();

    if (checkVLDuplicate(sheet, data.ldap.trim(), data.channel.trim(), data.vlDate)) {
      _log("WARN", "processVLForm", "Duplicate blocked", { ldap: data.ldap, date: data.vlDate });
      return { status:"error", message:"Duplicate Request Blocked." };
    }

    if (data.reasonForVL === "Birthday" && checkBirthdayLeaveYearlyDuplicate(sheet, data.ldap.trim(), data.vlDate)) {
      _log("WARN", "processVLForm", "Yearly birthday duplicate blocked", { ldap: data.ldap, date: data.vlDate });
      return { status:"error", message:"You already have an approved Birthday Leave for this calendar year." };
    }

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colMap  = {};
    headers.forEach(function(h, idx) {
      colMap[h.toString().replace(/[^a-zA-Z0-9]/g,'').toLowerCase()] = idx;
    });

    var newRow = new Array(headers.length).fill("");
    function setVal(colName, val) {
      var idx = colMap[colName.replace(/[^a-zA-Z0-9]/g,'').toLowerCase()];
      if (idx !== undefined) newRow[idx] = val;
    }

    setVal("Timestamp",              new Date());
    setVal("Email Address",          actor);
    setVal("LDAP",                   data.ldap.trim());
    setVal("Channel",                data.channel);
    setVal("VL Date",                data.vlDate);
    setVal("Team Lead",              data.teamLead);
    setVal("Reason for VL",          data.reasonForVL);
    setVal("Work Group",             data.workGroup);
    setVal("Site",                   data.site);
    setVal("Accruals Snip-it",       data.accrualUrl.trim());
    setVal("Accruals",               Math.round(parseFloat(data.accrualNum)));
    setVal("Month",                  "'" + submittedMonth);
    setVal("Date of Birthday",       data.bdayDate || "N/A");
    setVal("Proof / Artifacts",      data.proofUrl || "N/A");
    setVal("Status",                 "");
    setVal("Confirmation on Status", "");
    setVal("Email Sent",             "");
    setVal("Row UUID",               Utilities.getUuid());

    // Single appendRow call — no manual "find true last row" scan needed;
    // appendRow always writes after the sheet's real last row.
    sheet.appendRow(newRow);

    CacheService.getScriptCache().remove("vl_events_cache");
    _log("INFO", "processVLForm", "Submission accepted", {
      ldap: data.ldap, date: data.vlDate, channel: data.channel
    });
    return { status: "success" };

  } catch(e) {
    if (e.message && e.message.indexOf("timeout") >= 0) {
      _log("WARN", "processVLForm", "Lock timeout", { ldap: data.ldap || "unknown" });
      return { status:"error", message:"Server is busy — please try submitting again in a few seconds." };
    }
    _log("ERROR", "processVLForm", e.message, { ldap: data.ldap || "unknown" });
    return { status:"error", message:"Server error: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── DASHBOARD PAYLOAD (rate-limited) ─────────────────
function getDashboardPayload() {
  var actor = getSessionEmail();

  // Rate limit: max 20 full reloads per minute per user
  if (!_checkRateLimit("dashboard_" + actor, 20, 60)) {
    _log("WARN", "getDashboardPayload", "Dashboard rate limit hit", { actor: actor });
    return { error: "Too many requests. Please wait before reloading." };
  }

  var payload = {
    email: '', photoUrl: null,
    config: { isOpen: true, activeMonths: ["May 2026", "June 2026"], closedMsg: "", sunsetMode: false,
              isSupervisor: false, isAdmin: false, teamName: "" },
    events: []
  };

  try { payload.email = actor; } catch(e) {}
  try { var cfg = getFormConfig(); if (cfg) payload.config = cfg; } catch(e) {}
  try {
    var scriptCache = CacheService.getScriptCache();
    var cached = scriptCache.get("vl_events_cache");
    if (cached) {
      payload.events = JSON.parse(cached);
    } else {
      var lock2    = LockService.getScriptLock();
      var gotLock2 = lock2.tryLock(5000);

      if (!gotLock2) {
        Utilities.sleep(1500);
        cached = scriptCache.get("vl_events_cache");
      }

      try {
        cached = scriptCache.get("vl_events_cache");
        if (cached) {
          payload.events = JSON.parse(cached);
        } else {
          var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
          if (sheet && sheet.getLastRow() >= 2) {
            var evts = getCalendarData();
            var safe = Array.isArray(evts) ? evts : [];
            payload.events = safe;
            _safeCachePut("vl_events_cache", safe, 15);
          }
        }
      } finally {
        if (gotLock2) lock2.releaseLock();
      }
    }
  } catch(e) {
    _log("ERROR", "getDashboardPayload", e.message);
  }

  return payload;
}

// ── BIRTHDAY LEAVE YEARLY CHECK ───────────────────────
function checkBirthdayLeaveYearlyDuplicate(sheet, ldap, vlDate) {
  try {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return false;

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colMap  = {};
    headers.forEach(function(h, idx) {
      colMap[h.toString().replace(/[^a-zA-Z0-9]/g,'').toLowerCase()] = idx;
    });
    if (colMap["ldap"] === undefined || colMap["vldate"] === undefined || colMap["reasonforvl"] === undefined || colMap["status"] === undefined) return false;

    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var tz   = Session.getScriptTimeZone();

    var reqYear = null;
    var parsedReq = new Date(vlDate + "T00:00:00");
    if (!isNaN(parsedReq.getTime())) {
      reqYear = parsedReq.getFullYear();
    } else {
      reqYear = new Date(vlDate).getFullYear();
    }

    for (var i = 0; i < data.length; i++) {
      var sheetLdap = (data[i][colMap["ldap"]]||"").toString().trim().toLowerCase();
      var reason    = (data[i][colMap["reasonforvl"]]||"").toString().trim().toLowerCase();
      var status    = (data[i][colMap["status"]]||"").toString().trim().toLowerCase();
      var rawDate   = data[i][colMap["vldate"]];

      if (!sheetLdap || !rawDate) continue;

      if (sheetLdap === ldap.toLowerCase() && reason.includes("birthday")) {
        if (status.includes("approved") || status.includes("birthday")) {
          var sheetYear = null;
          if (Object.prototype.toString.call(rawDate) === '[object Date]' && !isNaN(rawDate.getTime())) {
            sheetYear = parseInt(Utilities.formatDate(rawDate, tz, "yyyy"), 10);
          } else {
            var rdStr2 = rawDate.toString().trim();
            if (/^\d{4}-\d{2}-\d{2}/.test(rdStr2)) {
              sheetYear = parseInt(rdStr2.substring(0, 4), 10);
            } else {
              var p = new Date(rdStr2 + "T00:00:00");
              sheetYear = !isNaN(p.getTime()) ? parseInt(Utilities.formatDate(p, tz, "yyyy"), 10) : new Date(rdStr2).getFullYear();
            }
          }
          if (sheetYear === reqYear) return true;
        }
      }
    }
    return false;
  } catch(e) { return false; }
}

// ── DUPLICATE CHECK ───────────────────────────────────
function checkVLDuplicate(sheet, ldap, channel, vlDate) {
  try {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return false;

    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colMap  = {};
    headers.forEach(function(h, idx) {
      colMap[h.toString().replace(/[^a-zA-Z0-9]/g,'').toLowerCase()] = idx;
    });
    if (colMap["ldap"] === undefined || colMap["channel"] === undefined || colMap["vldate"] === undefined) return false;

    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var tz   = Session.getScriptTimeZone();

    for (var i = 0; i < data.length; i++) {
      var sheetLdap    = (data[i][colMap["ldap"]]||"").toString().trim().toLowerCase();
      var sheetChannel = (data[i][colMap["channel"]]||"").toString().trim().toLowerCase();
      var rawDate      = data[i][colMap["vldate"]];
      if (!sheetLdap || !rawDate) continue;
      var sheetDateStr = "";
      if (Object.prototype.toString.call(rawDate) === '[object Date]' && !isNaN(rawDate.getTime())) {
        sheetDateStr = Utilities.formatDate(rawDate, tz, "yyyy-MM-dd");
      } else {
        var rdStr2 = rawDate.toString().trim();
        if (/^\d{4}-\d{2}-\d{2}/.test(rdStr2)) {
          sheetDateStr = rdStr2.substring(0, 10);
        } else {
          var p = new Date(rdStr2 + "T00:00:00");
          sheetDateStr = !isNaN(p.getTime()) ? Utilities.formatDate(p, tz, "yyyy-MM-dd") : rdStr2;
        }
      }
      if (sheetLdap === ldap.toLowerCase() && sheetChannel === channel.toLowerCase() && sheetDateStr === vlDate) return true;
    }
    return false;
  } catch(e) { return false; }
}

// ── MONTH HELPERS ─────────────────────────────────────
var MO_NAMES_GS = ["January","February","March","April","May","June","July","August","September","October","November","December"];
var MO_SHORT_GS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

function normalizeMonthFull(raw) {
  raw = (raw || "").toString().trim();
  var d = new Date(raw);
  if (!isNaN(d.getTime()) && raw.length > 8) return MO_NAMES_GS[d.getMonth()] + " " + d.getFullYear();
  var parts = raw.split(" ");
  if (parts.length >= 2) {
    if (MO_NAMES_GS.indexOf(parts[0]) >= 0) return parts[0] + " " + parts[1];
    var mi = MO_SHORT_GS.indexOf(parts[0]);
    if (mi >= 0) return MO_NAMES_GS[mi] + " " + parts[1];
  }
  return raw;
}

// ── FEEDBACK ─────────────────────────────────────────
function submitFeedback(feedbackText) {
  var actor = getSessionEmail() || "Unknown";
  if (!_checkRateLimit("feedback_" + actor, 5, 300)) {
    return { status: 'error', message: 'Too many feedback submissions. Please wait.' };
  }
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Feedback");
    if (!sheet) {
      sheet = ss.insertSheet("Feedback");
      sheet.getRange("A1:C1").setValues([["Timestamp","Email","Feedback"]])
        .setBackground("#f9ab00").setFontColor("#3e2723").setFontWeight("bold");
    }
    sheet.appendRow([new Date(), actor, feedbackText]);
    _log("INFO", "submitFeedback", "Feedback received", { actor: actor });
    return { status: 'success' };
  } catch(e) {
    _log("ERROR", "submitFeedback", e.message);
    return { status: 'error', message: e.message };
  }
}

// ── EMAIL FUNCTIONS (unchanged — kept for compatibility) ─
function getUnsentSummary() {
  try {
    var sheet = getVLSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { count: 0, html: "No data available." };
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colMap = {};
    headers.forEach(function(h, idx) { colMap[h.toString().replace(/[^a-zA-Z0-9]/g,'').toLowerCase()] = idx; });
    var statIdx = colMap["status"];
    var confIdx = colMap["confirmationonstatus"];
    var sentIdx = colMap["emailsent"];
    if (statIdx === undefined || confIdx === undefined) return { count: 0, html: "Required columns missing." };
    var data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    var appr = 0, den = 0, noal = 0;
    for (var i = 0; i < data.length; i++) {
      var status = (data[i][statIdx] || "").toString().trim();
      var conf   = (data[i][confIdx] || "").toString().trim();
      var sent   = sentIdx !== undefined ? (data[i][sentIdx] || "").toString().trim() : "";
      var lStat  = status.toLowerCase();
      if (status && !lStat.includes("pending") && conf && !sent) {
        if (lStat.includes("approved") || lStat.includes("birthday leave")) appr++;
        else if (lStat.includes("denied")) den++;
        else if (lStat.includes("no alloc")) noal++;
      }
    }
    var total = appr + den + noal;
    var html = total === 0
      ? "<div style='text-align:center;padding:20px;color:var(--text-muted);'>No confirmed, unsent requests.</div>"
      : "<div style='display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;text-align:center;margin-bottom:16px;'>" +
        "<div style='background:var(--green-bg);border:1px solid var(--green-border);border-radius:8px;padding:12px;'><div style='font-size:20px;font-weight:700;color:var(--green-main);'>" + appr + "</div><div style='font-size:10px;font-weight:700;color:var(--green-main);text-transform:uppercase;'>Approved</div></div>" +
        "<div style='background:var(--red-bg);border:1px solid var(--red-border);border-radius:8px;padding:12px;'><div style='font-size:20px;font-weight:700;color:var(--red-main);'>" + den + "</div><div style='font-size:10px;font-weight:700;color:var(--red-main);text-transform:uppercase;'>Denied</div></div>" +
        "<div style='background:var(--pink-bg);border:1px solid var(--pink-border);border-radius:8px;padding:12px;'><div style='font-size:20px;font-weight:700;color:var(--pink-main);'>" + noal + "</div><div style='font-size:10px;font-weight:700;color:var(--pink-main);text-transform:uppercase;'>No Alloc</div></div>" +
        "</div>";
    return { count: total, html: html };
  } catch(e) { return { count: 0, html: "Error." }; }
}

function sendNazunaNotifications() {
  var actor = getSessionEmail();
  if (!_checkRateLimit("sendNotif_" + actor, 3, 300)) {
    return { success: false, message: "Rate limit: wait before sending again." };
  }
  try {
    var sheet = getVLSheet();
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: "No data." };
    var statusColIdx = ensureColumn(sheet, "Status");
    var confColIdx   = ensureColumn(sheet, "Confirmation on Status");
    var sentColIdx   = ensureColumn(sheet, "Email Sent");
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colMap = {};
    headers.forEach(function(h, idx) { colMap[h.toString().replace(/[^a-zA-Z0-9]/g,'').toLowerCase()] = idx; });
    var data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    var tz = Session.getScriptTimeZone();
    var sentCount = 0;

    // Resolve columns ahead of loop
    var cEmail    = colMap["emailaddress"];
    var cLdap     = colMap["ldap"];
    var cVlDate   = colMap["vldate"];
    var cTeamLead = colMap["teamlead"];
    var cReason   = colMap["reasonforvl"];
    var cWorkGrp  = colMap["workgroup"];
    var cComments = colMap["comments"];
    var cTs       = colMap["timestamp"];
    var cAccruals = colMap["accruals"];
    var cAttend   = colMap["attendance"];
    var cSite     = colMap["site"];

    for (var i = 0; i < data.length; i++) {
      var row    = data[i];
      var status = (row[statusColIdx] || "").toString().trim();
      var conf   = (row[confColIdx]   || "").toString().trim();
      var sent   = (row[sentColIdx]   || "").toString().trim();
      var lStat  = status.toLowerCase();
      if (!status || lStat.includes("pending") || !conf || sent) continue;

      var emailAddress = cEmail !== undefined && row[cEmail] != null ? row[cEmail].toString().trim() : "";
      if (!emailAddress) continue;

      var ldap = cLdap !== undefined && row[cLdap] != null ? row[cLdap].toString().trim() : "";
      var vlDateRaw = cVlDate !== undefined && row[cVlDate] != null ? row[cVlDate] : "";
      var formattedDate = Object.prototype.toString.call(vlDateRaw) === '[object Date]'
        ? Utilities.formatDate(vlDateRaw, tz, "MMMM dd, yyyy") : vlDateRaw.toString();
      var emoji = "⚠️";
      if (lStat.includes("birthday")) emoji = "🎂";
      else if (lStat.includes("approved")) emoji = "✅";
      else if (lStat.includes("denied"))   emoji = "⛔";
      else if (lStat.includes("no alloc")) emoji = "🚫";

      var teamLead = cTeamLead !== undefined && row[cTeamLead] != null ? row[cTeamLead].toString().trim() : "";
      var reason   = cReason !== undefined && row[cReason] != null ? row[cReason].toString().trim() : "";
      var workGrp  = cWorkGrp !== undefined && row[cWorkGrp] != null ? row[cWorkGrp].toString().trim() : "";
      var comments = cComments !== undefined && row[cComments] != null ? row[cComments].toString().trim() : "";
      var ts       = cTs !== undefined && row[cTs] != null ? row[cTs].toString().trim() : "";
      var site     = cSite !== undefined && row[cSite] != null ? row[cSite].toString().trim() : "";

      var accruals = cAccruals !== undefined && row[cAccruals] != null ? row[cAccruals].toString().trim() : "";
      var rawAttend= cAttend !== undefined && row[cAttend] != null ? row[cAttend] : "";
      var attend   = rawAttend.toString().trim();
      if (typeof rawAttend === 'number') {
        attend = (rawAttend * 100).toFixed(2) + "%";
      }

      var htmlBody = createEmailTemplate(ldap, formattedDate, status, "-",
        teamLead, reason, workGrp, comments, ts, emoji, accruals, attend, site, conf);

      try {
        MailApp.sendEmail({ to: emailAddress, subject: emoji + " Google Play VL Calendar Update - " + status, htmlBody: htmlBody, name: "Google Play VL Calendar" });
        var stamp = Utilities.formatDate(new Date(), tz, "MM/dd/yy HH:mm") + " by " + actor.split('@')[0];
        sheet.getRange(i + 2, sentColIdx + 1).setValue(stamp);
        sentCount++;
        _log("INFO", "sendNazunaNotifications", "Email sent", { ldap: ldap, status: status });
      } catch(mailErr) {
        _log("ERROR", "sendNazunaNotifications", "Mail failed for " + ldap, { err: mailErr.message });
      }
    }
    return { success: true, count: sentCount };
  } catch(e) {
    _log("ERROR", "sendNazunaNotifications", e.message);
    return { success: false, message: e.message };
  }
}

// ── LIGHTWEIGHT EVENTS-ONLY ENDPOINT for live sync polling ──
// Returns only the events array — no config, no supervisor lookup.
// Much cheaper on quota than a full getDashboardPayload call.
function getEventsOnly() {
  var actor = getSessionEmail();

  if (!_checkRateLimit("eventsOnly_" + actor, 60, 60)) {
    return { events: null };
  }

  try {
    var cache  = CacheService.getScriptCache();
    var cached = cache.get("vl_events_cache");
    if (cached) {
      return { events: JSON.parse(cached), cached: true };
    }

    // ── STAMPEDE GUARD ──
    var lock    = LockService.getScriptLock();
    var gotLock = lock.tryLock(5000);

    if (!gotLock) {
      // Someone else is already refreshing — wait briefly rather than
      // also hitting the sheet ourselves.
      Utilities.sleep(1500);
      cached = cache.get("vl_events_cache");
      if (cached) return { events: JSON.parse(cached), cached: true };
      // Still nothing after waiting — fall through as a last resort.
    }

    try {
      cached = cache.get("vl_events_cache");
      if (cached) {
        return { events: JSON.parse(cached), cached: true };
      }

      var ss    = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName(SHEET_NAME);
      if (!sheet || sheet.getLastRow() < 2) {
        _safeCachePut("vl_events_cache", [], 15);
        return { events: [] };
      }

      var evts = getCalendarData();
      var safe = Array.isArray(evts) ? evts : [];
      _safeCachePut("vl_events_cache", safe, 15);
      _cachedEvents = safe;
      return { events: safe };
    } finally {
      if (gotLock) lock.releaseLock();
    }

  } catch(e) {
    _log("ERROR", "getEventsOnly", e.message);
    if (_cachedEvents) return { events: _cachedEvents, cached: true };
    return { events: null };
  }
}

// ── DELETE / ARCHIVE REQUEST ──────────────────────────
function deleteVLRequest(eventUuid) {
  var actor = getSessionEmail();
  if (!actor) return { success: false, message: "Not authenticated." };

  if (!_checkRateLimit("delete_" + actor, 10, 60)) {
    return { success: false, message: "Rate limit exceeded. Please wait before trying again." };
  }

  try {
    var ss        = SpreadsheetApp.getActiveSpreadsheet();
    var sheet     = getVLSheet();
    var rowIdx    = _getRowIndex(sheet);
    var rowNum    = rowIdx.index[eventUuid];

    if (!rowNum) {
      return { success: false, message: "Request not found. Please refresh the calendar." };
    }

    var lastCol  = sheet.getLastColumn();
    var headers  = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colMap   = {};
    headers.forEach(function(h, idx) {
      colMap[h.toString().replace(/[^a-zA-Z0-9]/g,'').toLowerCase()] = idx;
    });

    var rowData  = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
    var rowEmail = (rowData[colMap["emailaddress"]] || "").toString().trim().toLowerCase();
    var status   = (rowData[colMap["status"]] || "").toString().trim().toLowerCase();

    var config = getFormConfig();
    var isActorAdmin = config.isAdmin;
    var isActorSup   = config.isSupervisor;

    // Only owner, supervisor, or admin can delete
    var isOwner = (rowEmail === actor.toLowerCase());
    if (!isOwner && !isActorAdmin && !isActorSup) {
      _log("WARN", "deleteVLRequest", "Unauthorized delete attempt", { actor: actor, owner: rowEmail, uuid: eventUuid });
      return { success: false, message: "You can only remove your own requests unless you are a supervisor or admin." };
    }

    // Admins/Supervisors can delete anything; regular users cannot delete approved requests
    if (!isActorAdmin && !isActorSup && (status === "approved" || status.includes("birthday leave"))) {
      return { success: false, message: "Approved requests cannot be removed. Please contact your supervisor." };
    }

    // ── Get or create the Deleted Requests archive sheet ──
    var archiveName  = "Deleted Requests";
    var archiveSheet = ss.getSheetByName(archiveName);
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet(archiveName);
      var archiveHeaders = headers.concat(["Deleted By", "Deleted At"]);
      archiveSheet.getRange(1, 1, 1, archiveHeaders.length).setValues([archiveHeaders])
        .setBackground("#ea4335").setFontColor("#ffffff").setFontWeight("bold");
      archiveSheet.setFrozenRows(1);
      _log("INFO", "deleteVLRequest", "Created Deleted Requests archive sheet");
    }

    // ── Copy row to archive with metadata ──
    var tz         = Session.getScriptTimeZone();
    var deletedAt  = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    var archiveRow = rowData.concat([actor, deletedAt]);
    archiveSheet.appendRow(archiveRow);

    // ── Delete row from main sheet ──
    sheet.deleteRow(rowNum);

    _log("INFO", "deleteVLRequest", "Request archived and deleted", { uuid: eventUuid, actor: actor, row: rowNum });
    return { success: true };

  } catch(e) {
    _log("ERROR", "deleteVLRequest", e.message, { uuid: eventUuid });
    return { success: false, message: "Server error: " + e.message };
  }
}

// ══════════════════════════════════════════════════════
//  MWL/MVL DATA READER
// ══════════════════════════════════════════════════════

var MWL_SHEET_ID   = "1cp7CdA89e9qzIdcddeCRA12t5XUW2OBrCOwp4IVnh6g";
var MWL_TAB_NAME   = "Form Responses 1";

// Column indices (0-based, row 1 = headers so data starts row 2)
var MWL_COL = {
  TIMESTAMP:       4,   // E
  EMAIL:           5,   // F
  LDAP:            6,   // G
  LEAVE_DATE:      7,   // H
  TYPE:            8,   // I
  REGION:          11,  // L
  MGR_APPROVAL:    13,  // N
  WFM_REMARKS:     14,  // O
  REASON:          15,  // P
  REASON_COMMENTS: 16,  // Q
  SITE:            18,  // S
  TEAM_CAPTAIN:    19,  // T
  EMP_ID:          2,   // C
  ROLE:            3,   // D
};

function getMWLData() {
  var actor = getSessionEmail();
  if (!_checkRateLimit("mwl_" + actor, 20, 60)) {
    return { error: "Too many requests. Please wait before reloading." };
  }

  try {
    var ss    = SpreadsheetApp.openById(MWL_SHEET_ID);
    var sheet = ss.getSheetByName(MWL_TAB_NAME);
    if (!sheet) return { error: "Sheet not found: " + MWL_TAB_NAME };

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return { events: [] };

    var data    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var tz      = Session.getScriptTimeZone();

    // ── Dedup: key = LDAP|LeaveDate|Type, keep latest timestamp ──
    var dedupMap = {};

    for (var i = 0; i < data.length; i++) {
      var row  = data[i];

      var ldap      = (row[MWL_COL.LDAP]      || "").toString().trim();
      var typeOfLv  = (row[MWL_COL.TYPE]       || "").toString().trim();
      var rawDate   = row[MWL_COL.LEAVE_DATE];
      var rawTs     = row[MWL_COL.TIMESTAMP];

      if (!ldap || !rawDate || !typeOfLv) continue;

      // Format leave date
      var leaveDateStr = "";
      if (Object.prototype.toString.call(rawDate) === '[object Date]' && !isNaN(rawDate.getTime())) {
        leaveDateStr = Utilities.formatDate(rawDate, tz, "yyyy-MM-dd");
      } else {
        var parsed = new Date(rawDate);
        leaveDateStr = !isNaN(parsed.getTime())
          ? Utilities.formatDate(parsed, tz, "yyyy-MM-dd")
          : rawDate.toString().trim();
      }

      // Format timestamp for comparison
      var tsVal = 0;
      if (Object.prototype.toString.call(rawTs) === '[object Date]' && !isNaN(rawTs.getTime())) {
        tsVal = rawTs.getTime();
      } else if (rawTs) {
        tsVal = new Date(rawTs).getTime() || 0;
      }

      var dedupKey = ldap.toLowerCase() + "|" + leaveDateStr + "|" + typeOfLv.toLowerCase();

      if (!dedupMap[dedupKey] || tsVal > dedupMap[dedupKey].tsVal) {
        // Resolve status: N first, fall back to O
        var mgrApproval = (row[MWL_COL.MGR_APPROVAL] || "").toString().trim();
        var wfmRemarks  = (row[MWL_COL.WFM_REMARKS]  || "").toString().trim();
        var status      = "";

        if (mgrApproval) {
          status = mgrApproval;
        } else if (wfmRemarks) {
          status = wfmRemarks;
        } else {
          status = "Pending";
        }

        // Format timestamp display string
        var tsStr = "Unknown";
        if (Object.prototype.toString.call(rawTs) === '[object Date]' && !isNaN(rawTs.getTime())) {
          tsStr = Utilities.formatDate(rawTs, tz, "MMM d, yyyy h:mm a");
        } else if (rawTs) {
          tsStr = rawTs.toString();
        }

        dedupMap[dedupKey] = {
          tsVal:          tsVal,
          id:             dedupKey,
          ldap:           ldap,
          email:          (row[MWL_COL.EMAIL]           || "").toString().trim().toLowerCase(),
          empId:          (row[MWL_COL.EMP_ID]          || "").toString().trim(),
          role:           (row[MWL_COL.ROLE]             || "").toString().trim(),
          leaveDate:      leaveDateStr,
          type:           typeOfLv,
          region:         (row[MWL_COL.REGION]          || "").toString().trim(),
          status:         status,
          mgrApproval:    mgrApproval,
          wfmRemarks:     wfmRemarks,
          reason:         (row[MWL_COL.REASON]          || "").toString().trim(),
          reasonComments: (row[MWL_COL.REASON_COMMENTS] || "").toString().trim(),
          site:           (row[MWL_COL.SITE]            || "").toString().trim(),
          teamCaptain:    (row[MWL_COL.TEAM_CAPTAIN]    || "").toString().trim(),
          timestamp:      tsStr,
        };
      }
    }

    var events = Object.values(dedupMap).map(function(e) {
      delete e.tsVal;
      return e;
    });

    _log("INFO", "getMWLData", "Loaded MWL events", { count: events.length });
    return { events: events };

  } catch(e) {
    _log("ERROR", "getMWLData", e.message);
    return { error: e.message };
  }
}

// ── VL ALLOCATIONS CONFIG ─────────────────────────────
var ALLOC_NAME = "VL Allocations";

function getOrCreateAllocationsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ALLOC_NAME);
  if (!sh) {
    sh = ss.insertSheet(ALLOC_NAME);
    sh.getRange("A1:F1").setValues([["Pool Name","Channel","Workgroups","Region","Daily Cap","Weekly Cap"]])
      .setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold");
    var defaults = [
      ["Chat (NA+GTV)",       "Chat",  "Play NA, Play GTV",           "MNL", 1, 6],
      ["Chat (NA+GTV)",       "Chat",  "Play NA, Play GTV",           "CEB", 1, 4],
      ["Phone (NA+HVU+DVIP)", "Phone", "Play NA, Play HVU, Play VIP", "MNL", 1, 2],
      ["Phone (NA+HVU+DVIP)", "Phone", "Play NA, Play HVU, Play VIP", "CEB", 1, 3],
      ["Email (NA+GTV+HVU)",  "Email", "Play NA, Play GTV, Play HVU", "MNL", 1, 3],
      ["Email (NA+GTV+HVU)",  "Email", "Play NA, Play GTV, Play HVU", "CEB", 1, 1],
      ["HVU+VIP (Chat)",      "Chat",  "Play HVU, Play VIP",          "MNL", 1, 2],
      ["HVU+VIP (Chat)",      "Chat",  "Play HVU, Play VIP",          "CEB", 1, 2]
    ];
    sh.getRange(2, 1, defaults.length, 6).setValues(defaults);
    sh.setColumnWidth(1, 170); sh.setColumnWidth(3, 220); sh.setFrozenRows(1);
    _log("INFO", "getOrCreateAllocationsSheet", "Created VL Allocations sheet with defaults");
  }
  return sh;
}

// ── ONE-TIME: Push updated allocation caps to the existing sheet ──
// Run this once from the Apps Script editor (select this function in the
// dropdown, click Run) to overwrite "VL Allocations" with the latest caps.
// Safe to leave in place afterward — getVLAllocations() always reads
// whatever is currently in the sheet, so this only matters when you need
// to push a new set of numbers.
function syncVLAllocationDefaults() {
  var sh = getOrCreateAllocationsSheet();
  var defaults = [
    ["Chat (NA+GTV)",       "Chat",  "Play NA, Play GTV",           "MNL", 1, 6],
    ["Chat (NA+GTV)",       "Chat",  "Play NA, Play GTV",           "CEB", 1, 4],
    ["Phone (NA+HVU+DVIP)", "Phone", "Play NA, Play HVU, Play VIP", "MNL", 1, 2],
    ["Phone (NA+HVU+DVIP)", "Phone", "Play NA, Play HVU, Play VIP", "CEB", 1, 3],
    ["Email (NA+GTV+HVU)",  "Email", "Play NA, Play GTV, Play HVU", "MNL", 1, 3],
    ["Email (NA+GTV+HVU)",  "Email", "Play NA, Play GTV, Play HVU", "CEB", 1, 1],
    ["HVU+VIP (Chat)",      "Chat",  "Play HVU, Play VIP",          "MNL", 1, 2],
    ["HVU+VIP (Chat)",      "Chat",  "Play HVU, Play VIP",          "CEB", 1, 2]
  ];
  var lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, 6).clearContent();
  sh.getRange(2, 1, defaults.length, 6).setValues(defaults);
  _log("INFO", "syncVLAllocationDefaults", "Overwrote VL Allocations with updated caps", { rows: defaults.length });
  SpreadsheetApp.getActiveSpreadsheet().toast("VL Allocations updated: " + defaults.length + " rows written.");
}

function getVLAllocations() {
  try {
    var sh = getOrCreateAllocationsSheet();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return [];
    var data = sh.getRange(2, 1, lastRow - 1, 6).getValues();
    var pools = [];
    data.forEach(function(r) {
      var poolName = (r[0]||"").toString().trim();
      if (!poolName) return;
      pools.push({
        pool:       poolName,
        channel:    (r[1]||"").toString().trim(),
        workgroups: (r[2]||"").toString().split(",").map(function(w){return w.trim();}).filter(Boolean),
        region:     (r[3]||"").toString().trim().toUpperCase(),
        dailyCap:   Number(r[4]) || 0,
        weeklyCap:  Number(r[5]) || 0
      });
    });
    return pools;
  } catch(e) {
    _log("ERROR", "getVLAllocations", e.message);
    return [];
  }
}
