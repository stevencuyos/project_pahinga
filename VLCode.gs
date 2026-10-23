// ══════════════════════════════════════════════════════
//  Project Pahinga — Support VL Calendar & Schedule Management
//  Code.gs  v9.0 — Indexed, Stateless, Rate-Limited
// ══════════════════════════════════════════════════════

var SHEET_NAME  = "Support VL Requests";
var CONFIG_NAME = "FormConfig";
var SUP_NAME    = "Roles";
var LOG_NAME    = "System Log";

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
      var wk = "warn_big_" + key;
      if (!CacheService.getScriptCache().get(wk)) {
        CacheService.getScriptCache().put(wk, "1", 600);
        _log("WARN", "_safeCachePut", "Payload too large to cache", { key: key, size: json.length });
      }
    }
  } catch(e) {
    _log("WARN", "_safeCachePut", "Cache put failed", { key: key, err: e.message });
  }
}

// ── RATE LIMITER (uses CacheService) ─────────────────
// key: string identifier, maxCalls: per window, windowSec: window in seconds
function _checkRateLimit(key, maxCalls, windowSec) {
  var cache    = CacheService.getScriptCache();
  var cacheKey = "rl_" + key;
  var now      = Date.now();
  var st = null;
  try { st = JSON.parse(cache.get(cacheKey) || "null"); } catch (e) {}
  if (!st || now - st.t >= windowSec * 1000) st = { c: 0, t: now };
  if (st.c >= maxCalls) {
    _log("WARN", "_checkRateLimit", "Rate limit hit", { key: key, count: st.c });
    return false;
  }
  st.c++;
  cache.put(cacheKey, JSON.stringify(st), windowSec);
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
    .setTitle('Pahinga - Support VL Calendar')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://www.gstatic.com/android/market_images/web/favicon_v2.ico');
}

function getSessionEmail() {
  try { return Session.getActiveUser().getEmail(); } catch(e) { return ""; }
}

function _redactEvents(events, email) {
  var me = (email || "").toLowerCase();
  return (events || []).map(function(e) {
    if (e.email === me) return e;
    var c = {};
    Object.keys(e).forEach(function(k) { c[k] = e[k]; });
    c.reasonText = ""; c.docsUrl = ""; c.creditsUrl = ""; c.creditsSource = ""; c.creditsAccruals = ""; c.bcpSubject = "";
    c.remarks = ""; c.confirmation = ""; c.emailSent = ""; c.cancelReq = ""; c.cancelDays = "";
    return c;
  });
}

function getOrCreateConfigSheet() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var cfg = ss.getSheetByName(CONFIG_NAME);
  if (!cfg) {
    cfg = ss.insertSheet(CONFIG_NAME);
    cfg.getRange("A1:B1").setValues([["Setting","Value"]])
      .setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold");
    cfg.getRange("A2:B3").setValues([
      ["IsOpen",       "TRUE"],
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
    sup.getRange("A1:G1").setValues([["Nickname","Email","Site","LDAP","Role","Supervisor","Access"]])
      .setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold");
    sup.getRange("A2:G2").setValues([["Steven","stevenjosephc@google.com","CEB","stevenjosephc","Supervisor","","Admin"]]);
    sup.setColumnWidth(1, 120); sup.setColumnWidth(2, 250); sup.setColumnWidth(3, 80); sup.setColumnWidth(4, 140); sup.setColumnWidth(5, 110); sup.setColumnWidth(6, 140); sup.setColumnWidth(7, 90);
    sup.setFrozenRows(1);
    _log("INFO", "getOrCreateSupervisorSheet", "Created Supervisors sheet");
  }
  return sup;
}

function getFormConfig() {
  var res = { isOpen: true, closedMsg: "Form is closed.", isSupervisor: false, isAdmin: false, teamName: "", role: "", site: "" };

  // Config/Supervisor sheets rarely change — cache the non-personalized
  // part (isOpen/activeMonths/closedMsg) for 30s so this doesn't do 2
  // sheet reads on every single dashboard load from every user.
  var cache      = CacheService.getScriptCache();
  var baseCached = cache.get("vl_config_base");
  if (baseCached) {
    var base = JSON.parse(baseCached);
    res.isOpen = base.isOpen; res.closedMsg = base.closedMsg;
  } else {
  try {
    var data = getOrCreateConfigSheet().getRange("A2:B10").getValues();
    data.forEach(function(r) {
      var k = (r[0]||"").toString().trim();
      var v = (r[1]||"").toString().trim();
      if (!k) return;
      if (k === "IsOpen")       res.isOpen       = (v.toUpperCase() === "TRUE");
      if (k === "ClosedMsg")    res.closedMsg    = v;
    });
    cache.put("vl_config_base", JSON.stringify({ isOpen: res.isOpen, closedMsg: res.closedMsg }), 30);
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
        res.role         = supRole.role || "";
        res.site         = supRole.site || "";
      } else {
        var supData = getOrCreateSupervisorSheet().getDataRange().getValues();
        for (var i = 1; i < supData.length; i++) {
          if ((supData[i][1]||"").toString().trim().toLowerCase() === email) {
            var access       = (supData[i][6]||"").toString().trim().toLowerCase();
            res.role         = (supData[i][4]||"").toString().trim();
            res.site         = (supData[i][2]||"").toString().trim().toUpperCase();
            res.teamName     = (supData[i][0]||"").toString().trim();
            res.isAdmin      = access === "admin";
            res.isSupervisor = res.isAdmin || access === "manager";
            break;
          }
        }
        _safeCachePut(supCacheKey, { isSupervisor: res.isSupervisor, teamName: res.teamName, isAdmin: res.isAdmin, role: res.role, site: res.site }, 30);
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
    var h = ["Timestamp","Email Address","LDAP","Role","Site","VL Start Date","VL End Date","Days",
             "VL Type","Reason","Supporting Docs","Support Coverage","VL Credits Snip",
             "BCP Email Subject","Month","Status","Comments","Confirmation on Status",
             "Email Sent","Row UUID"];
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

// ── DATE HELPERS ──────────────────────────────────────
function _toYmd(raw, tz) {
  if (Object.prototype.toString.call(raw) === '[object Date]' && !isNaN(raw.getTime())) {
    return Utilities.formatDate(raw, tz, "yyyy-MM-dd");
  }
  var s = (raw == null ? "" : raw).toString().trim();
  if (!s) return "";
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  var p = new Date(s + "T00:00:00");
  return !isNaN(p.getTime()) ? Utilities.formatDate(p, tz, "yyyy-MM-dd") : s;
}

// Returns every yyyy-MM-dd from start to end inclusive (hard cap 62 days)
function _expandDates(startYmd, endYmd) {
  var out = [];
  if (!startYmd) return out;
  var end = endYmd || startYmd;
  if (end < startYmd) end = startYmd;
  var p = startYmd.split('-');
  var d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2]));
  for (var n = 0; n < 62; n++) {
    var ymd = Utilities.formatDate(d, "GMT", "yyyy-MM-dd");
    out.push(ymd);
    if (ymd >= end) break;
    d.setUTCDate(d.getUTCDate() + 1);
  }
  return out;
}

// ── ROSTER (from Roles tab: A Nickname, B Email, C Site, D LDAP, E Role) ──
function getRoster() {
  var cache = CacheService.getScriptCache();
  var hit   = cache.get("vl_roster");
  if (hit) { try { return JSON.parse(hit); } catch(e) {} }
  var out = [];
  try {
    var vals  = getOrCreateSupervisorSheet().getDataRange().getValues();
    var clean = function(s) { return (s || "").toString().replace(/[<>&"']/g, "").trim(); };
    for (var i = 1; i < vals.length; i++) {
      var email = (vals[i][1] || "").toString().trim().toLowerCase();
      var role  = clean(vals[i][4]);
      if (!email || !role) continue;
      var ldap = ((vals[i][3] || "").toString().trim().toLowerCase() || email.split("@")[0]).replace(/[^a-z0-9._-]/g, "");
      out.push({ ldap: ldap, name: clean(vals[i][0]), site: clean(vals[i][2]).toUpperCase(), role: role });
    }
    _safeCachePut("vl_roster", out, 300);
  } catch(e) { _log("ERROR", "getRoster", e.message); }
  return out;
}

function getCalendarData() {
  try {
    var sheet = getVLSheet();
    if (sheet.getLastRow() < 2 || sheet.getLastColumn() < 1) return [];
    _getRowIndex(sheet); // ensures + backfills Row UUIDs
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colMap  = {};
    headers.forEach(function(h, idx) {
      colMap[h.toString().replace(/[^a-zA-Z0-9]/g,'').toLowerCase()] = idx;
    });
    var data   = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var tz     = Session.getScriptTimeZone();
    var events = [];
    var cutoff = Utilities.formatDate(new Date(Date.now() - 120 * 86400000), tz, "yyyy-MM-dd");

    function g(row, key) {
      var ix = colMap[key];
      return ix !== undefined && row[ix] != null ? row[ix].toString().trim() : "";
    }

    for (var i = 0; i < data.length; i++) {
      var row   = data[i];
      var rawTs = colMap["timestamp"] !== undefined ? row[colMap["timestamp"]] : "";
      var ldap  = g(row, "ldap");
      if (!rawTs || !ldap) continue;

      var startYmd = colMap["vlstartdate"] !== undefined ? _toYmd(row[colMap["vlstartdate"]], tz) : "";
      if (!startYmd) continue;
      var endYmd = colMap["vlenddate"] !== undefined ? _toYmd(row[colMap["vlenddate"]], tz) : "";
      if (!endYmd || endYmd < startYmd) endYmd = startYmd;
      if (endYmd < cutoff) continue;

      var isDate  = Object.prototype.toString.call(rawTs) === '[object Date]' && !isNaN(rawTs.getTime());
      var tsStr   = isDate ? Utilities.formatDate(rawTs, tz, "MMM d, yyyy h:mm a") : rawTs.toString();
      var tsEpoch = 0;
      if (isDate) tsEpoch = rawTs.getTime();
      else { var pt = new Date(rawTs); tsEpoch = !isNaN(pt.getTime()) ? pt.getTime() : 0; }

      var uuid      = g(row, "rowuuid") || ("ev_" + i);
      var roleVal   = g(row, "role");
      var vlTypeVal = g(row, "vltype");
      var dayList   = _expandDates(startYmd, endYmd);

      dayList.forEach(function(ymd) {
        events.push({
          id: uuid, rowNum: i + 2, timestamp: tsStr, tsEpoch: tsEpoch,
          email: g(row, "emailaddress").toLowerCase(),
          ldap: ldap, role: roleVal, site: g(row, "site"),
          date: ymd, startDate: startYmd, endDate: endYmd, totalDays: dayList.length,
          vlType: vlTypeVal, reasonText: g(row, "reason"), docsUrl: g(row, "supportingdocs"),
          coverage: g(row, "supportcoverage"), creditsUrl: g(row, "vlcreditssnip"), creditsSource: g(row, "vlcreditssource"), creditsAccruals: g(row, "vlcreditsaccruals"),
          bcpSubject: g(row, "bcpemailsubject"),
          bcpNotes: g(row, "bcpnotes"), bcpSentFor: g(row, "bcpsentfor"), bcpOpened: g(row, "bcpopened"),
          month: g(row, "month"),
          status: g(row, "status") || "Pending",
          remarks: g(row, "comments"),
          confirmation: g(row, "confirmationonstatus"),
          emailSent: g(row, "emailsent"),
          cancelReq: g(row, "cancelrequest"),
          cancelDays: g(row, "canceldays"),
        });
      });
    }
    return events;
  } catch(e) {
    _log("ERROR", "getCalendarData", e.message);
    return [];
  }
}

// ── STATUS UPDATE (Supports Full or Partial/Split Multi-Day Approvals) ─
function updateRequestStatus(eventUuid, oldStatus, newStatus, adminComment, selectedDates, unselectedStatus) {
  var actor = getSessionEmail();

  if (!getFormConfig().isSupervisor) return { success: false, message: "Managers only." };

  if (!_checkRateLimit("statusUpdate_" + actor, 30, 60)) {
    return { success: false, message: "Rate limit exceeded. Please wait before making more changes." };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var sheet  = getVLSheet();
    var rowIdx = _getRowIndex(sheet);
    var rowNum = rowIdx.index[eventUuid];

    if (!rowNum) {
      _log("WARN", "updateRequestStatus", "UUID not found in index", { uuid: eventUuid });
      return { success: false, message: "Row not found. The calendar may be out of sync — please refresh." };
    }

    var statusColIdx  = ensureColumn(sheet, "Status") + 1;
    var confColIdx    = ensureColumn(sheet, "Confirmation on Status") + 1;
    var commentColIdx = ensureColumn(sheet, "Comments") + 1;

    var currentStatus     = sheet.getRange(rowNum, statusColIdx).getValue().toString().trim();
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
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colMap  = {};
    headers.forEach(function(h, idx) {
      colMap[h.toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase()] = idx;
    });

    var rowVals  = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
    var startYmd = _toYmd(rowVals[colMap["vlstartdate"]], tz);
    var endYmd   = _toYmd(rowVals[colMap["vlenddate"]], tz) || startYmd;
    var allDates = _expandDates(startYmd, endYmd);

    // Check if this is a partial-day update on a multi-day request
    var selSet = {};
    if (Array.isArray(selectedDates) && selectedDates.length > 0) {
      selectedDates.forEach(function(d) { selSet[d] = true; });
    }
    var pickedCount = allDates.filter(function(d) { return !!selSet[d]; }).length;
    var isPartial   = allDates.length > 1 && pickedCount > 0 && pickedCount < allDates.length;

    if (!isPartial) {
      // Standard full-request update
      sheet.getRange(rowNum, statusColIdx).setValue(newStatus);
      sheet.getRange(rowNum, confColIdx).setValue(stamp);
      sheet.getRange(rowNum, ensureColumn(sheet, "Email Sent") + 1).clearContent();
      if (adminComment && adminComment.trim()) {
        sheet.getRange(rowNum, commentColIdx).setValue(adminComment.trim());
      }
      var sched = _syncScheduleForStatus(sheet, rowNum, eventUuid, oldStatus, newStatus);
      CacheService.getScriptCache().remove("vl_events_cache");
      _log("INFO", "updateRequestStatus", "Status updated", {
        uuid: eventUuid, row: rowNum, from: oldStatus, to: newStatus
      });
      return { success: true, newStatus: newStatus, sched: sched };
    }

    // ── PARTIAL APPROVAL / SPLIT LOGIC ──
    var restStatus = (unselectedStatus === "Keep" || !unselectedStatus) ? currentStatus : unselectedStatus;
    var segments = [];
    allDates.forEach(function(ymd) {
      var st = selSet[ymd] ? newStatus : restStatus;
      var last = segments[segments.length - 1];
      if (last && last.status === st) {
        last.end = ymd;
        last.days++;
      } else {
        segments.push({ start: ymd, end: ymd, days: 1, status: st });
      }
    });

    // If previously approved, revert old schedule sync entries before splitting
    if (VL_APPROVED_RX.test(normalizedCurrent)) {
      try { _revertVLJobs([eventUuid]); } catch (re) {}
    }

    var monthLabel = function(s) { var p = s.split("-"); return MO_NAMES_GS[parseInt(p[1], 10) - 1] + " " + p[0]; };
    var applyJobs  = [];
    var emailVal   = (rowVals[colMap["emailaddress"]] || "").toString().trim().toLowerCase();
    var ldapVal    = (rowVals[colMap["ldap"]] || "").toString().trim().toLowerCase() || emailVal.split("@")[0];

    segments.forEach(function(seg, idx) {
      var segRow  = rowVals.slice();
      var segUuid = idx === 0 ? eventUuid : Utilities.getUuid();
      var segChanged = (seg.status || "Pending").toLowerCase() !== normalizedCurrent.toLowerCase();

      segRow[colMap["vlstartdate"]] = seg.start;
      segRow[colMap["vlenddate"]]   = seg.end;
      if (colMap["days"] !== undefined)  segRow[colMap["days"]]  = seg.days;
      if (colMap["month"] !== undefined) segRow[colMap["month"]] = "'" + monthLabel(seg.start);
      segRow[statusColIdx - 1] = seg.status;
      segRow[colMap["rowuuid"]] = segUuid;

      if (selSet[seg.start] || segChanged) {
        segRow[confColIdx - 1] = stamp + " (Split)";
        if (colMap["emailsent"] !== undefined) segRow[colMap["emailsent"]] = "";
        if (adminComment && adminComment.trim()) {
          segRow[commentColIdx - 1] = adminComment.trim();
        }
      }

      if (idx === 0) {
        sheet.getRange(rowNum, 1, 1, lastCol).setValues([segRow]);
      } else {
        sheet.appendRow(segRow);
      }

      if (VL_APPROVED_RX.test(seg.status || "")) {
        applyJobs.push({ uuid: segUuid, ldap: ldapVal, email: emailVal, start: seg.start, end: seg.end });
      }
    });

    SpreadsheetApp.flush();
    var schedRes = applyJobs.length ? _applyVLJobs(applyJobs) : null;
    var freshEvents = getCalendarData();
    _safeCachePut("vl_events_cache", freshEvents, 15);

    _log("INFO", "updateRequestStatus", "Split multi-day request", {
      uuid: eventUuid, segments: segments.length, picked: pickedCount, to: newStatus, rest: restStatus
    });

    return {
      success: true,
      split: true,
      newStatus: newStatus,
      events: freshEvents,
      sched: schedRes ? { action: "applied", applied: schedRes.applied, skipped: schedRes.skipped, noColumn: schedRes.noColumn, noRow: schedRes.noRow } : null
    };

  } catch(e) {
    _log("ERROR", "updateRequestStatus", e.message, { uuid: eventUuid });
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── BATCH STATUS UPDATE (optimized for multiple rows) ──
function batchUpdateStatuses(eventUuuids, newStatus, adminComment) {
  var actor = getSessionEmail();
  if (!Array.isArray(eventUuuids) || eventUuuids.length === 0) {
    return { success: false, message: "No requests selected." };
  }

  if (!getFormConfig().isSupervisor) return { success: false, message: "Managers only." };

  // Rate limit: batch counts as one action for basic quota, but we log the size
  if (!_checkRateLimit("batchUpdate_" + actor, 10, 60)) {
    return { success: false, message: "Rate limit exceeded. Please wait a moment." };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var sheet   = getVLSheet();
    var rowIdx  = _getRowIndex(sheet);
    var tz      = Session.getScriptTimeZone();
    var stamp   = actor.split('@')[0] + " — " + Utilities.formatDate(new Date(), tz, "MM/dd/yy HH:mm") + " (Batch)";

    var statusColIdx  = ensureColumn(sheet, "Status") + 1;
    var confColIdx    = ensureColumn(sheet, "Confirmation on Status") + 1;
    var commentColIdx = ensureColumn(sheet, "Comments") + 1;
    var sentColIdx    = ensureColumn(sheet, "Email Sent") + 1;

    var updatedCount = 0;
    var sched = null;
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
          values[localIdx][sentColIdx - 1]   = "";
          if (adminComment && adminComment.trim()) {
            values[localIdx][commentColIdx - 1] = adminComment.trim();
          }
          updatedCount++;
        } else {
          errors.push(uuid + " not found");
        }
      });
      range.setValues(values);
      sched = _syncScheduleBatch(sheet, lastCol, values, minRow, rowIdx, eventUuuids, newStatus);
      CacheService.getScriptCache().remove("vl_events_cache");
    }

    _log("INFO", "batchUpdateStatuses", "Batch update complete", {
      count: updatedCount, to: newStatus, errorCount: errors.length
    });

    return { success: true, count: updatedCount, errors: errors, sched: sched };

  } catch(e) {
    _log("ERROR", "batchUpdateStatuses", e.message);
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── FORM SUBMISSION (rate-limited, logged) ────────────
function _safeCell(s, max) {
  s = (s == null ? "" : String(s)).trim().slice(0, max || 500);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function processVLForm(data) {
  var actor = getSessionEmail();
  if (!_checkRateLimit("submit_" + actor, 3, 300)) {
    _log("WARN", "processVLForm", "Submission rate limit hit", { actor: actor });
    return { status: "error", message: "Too many submissions. Please wait a few minutes before trying again." };
  }

  var config = getFormConfig();
  if (!config.isOpen) return { status: "closed", message: config.closedMsg };
  if (!actor) return { status: "error", message: "Could not identify your account." };
  if (!config.role || !config.site) return { status: "error", message: "You are not registered in the Roles list. Please contact your manager." };

  var ymdRx = /^\d{4}-\d{2}-\d{2}$/;
  var start = (data.startDate || "").toString().trim();
  var end   = (data.endDate   || "").toString().trim() || start;
  if (!ymdRx.test(start) || !ymdRx.test(end)) return { status: "error", message: "A valid start and end date are required." };
  if (end < start) return { status: "error", message: "End date cannot be before the start date." };
  var numDays = _expandDates(start, end).length;
  if (numDays > 31) return { status: "error", message: "A single request can cover at most 31 days." };

  var vlType     = (data.vlType || "").toString().trim();
  var reason     = (data.reason || "").toString().trim();
  var docsUrl    = (data.docsUrl || "").toString().trim();
  var creditsUrl = (data.creditsUrl || "").toString().trim();
  var bcpNotes = (data.bcpNotes || "").toString().trim();
  var bcpSubject = (data.bcpSubject || "").toString().trim();
  var creditsSource = /supervisor|team\s*lead/i.test(config.role || "") ? "Workday" : "eStart";
  var creditsAccruals = parseFloat(data.creditsAccruals);
  if (!vlType) return { status: "error", message: "VL Type is required." };
  if (["Out of Town / Vacation", "Personal", "Family Gathering", "Emergency", "Mental Wellness", "Birthday", "Medical", "Bereavement", "Government / Legal", "Moving House"].indexOf(vlType) < 0) return { status: "error", message: "Invalid VL Type." };
  if (reason.length < 3) return { status: "error", message: "Reason for VL is required." };
  if (isNaN(creditsAccruals) || creditsAccruals < 0 || creditsAccruals > 999) return { status: "error", message: "Enter your VL accruals as a number." };
  if (!/^https?:\/\//i.test(creditsUrl)) return { status: "error", message: "A valid VL credits snip link is required." };
  if (docsUrl && !/^https?:\/\//i.test(docsUrl)) return { status: "error", message: "Supporting documents must be a valid link." };
  if (!BCP_AUTO_COMPOSE && !bcpSubject) return { status: "error", message: "Email subject of your BCP is required." };

  var myLdap   = actor.split("@")[0].toLowerCase();
  var coverage = (Array.isArray(data.coverage) ? data.coverage : []).map(function(c) {
    return c.toString().trim().toLowerCase().replace(/[^a-z0-9._-]/g, "");
  }).filter(function(c, i, arr) { return c && arr.indexOf(c) === i && c !== myLdap; });
  if (coverage.length === 0) return { status: "error", message: "Please add at least one person covering for you." };
  var known = getRoster().map(function(p) { return p.ldap; });
  var unknown = coverage.filter(function(c) { return known.indexOf(c) < 0; });
  if (unknown.length) return { status: "error", message: "Not found in the roster: " + unknown.join(", ") };

  var monthLabel = function(s) { var p = s.split("-"); return MO_NAMES_GS[parseInt(p[1], 10) - 1] + " " + p[0]; };
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  if (start < todayStr) return { status: "error", message: "VL start date cannot be in the past." };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var sheet = getVLSheet();

    if (checkVLDuplicate(sheet, actor, start, end)) {
      _log("WARN", "processVLForm", "Overlap blocked", { actor: actor, start: start, end: end });
      return { status: "error", message: "You already have a request overlapping these dates." };
    }

    // Make sure every column we write to exists (adds any that are missing)
    var REQUIRED_HEADERS = ["Timestamp","Email Address","LDAP","Role","Site","VL Start Date",
      "VL End Date","Days","VL Type","Reason","Supporting Docs","Support Coverage",
      "VL Credits Snip","BCP Email Subject","Month","Status","Comments",
      "Confirmation on Status","Email Sent","Row UUID","VL Credits Source","VL Credits Accruals",
      "BCP Notes","BCP Sent For","BCP Opened"];
    REQUIRED_HEADERS.forEach(function(h) { ensureColumn(sheet, h); });

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var colMap  = {};
    headers.forEach(function(h, idx) {
      colMap[h.toString().replace(/[^a-zA-Z0-9]/g,'').toLowerCase()] = idx;
    });
    var newRow = new Array(headers.length).fill("");
    function setVal(colName, val) {
      var idx = colMap[colName.replace(/[^a-zA-Z0-9]/g,'').toLowerCase()];
      if (idx === undefined) throw new Error("Missing column: " + colName);
      newRow[idx] = val;
    }

    setVal("Timestamp",              new Date());
    setVal("Email Address",          actor);
    setVal("LDAP",                   myLdap);
    setVal("Role",                   config.role);
    setVal("Site",                   config.site);
    setVal("VL Start Date",          start);
    setVal("VL End Date",            end);
    setVal("Days",                   numDays);
    setVal("VL Type",                vlType);
    setVal("Reason",                 _safeCell(reason, 1000));
    setVal("Supporting Docs",        docsUrl);
    setVal("Support Coverage",       coverage.join(", "));
    setVal("VL Credits Snip",        creditsUrl);
    setVal("VL Credits Source",      creditsSource);
    setVal("VL Credits Accruals",    creditsAccruals);
    var bcpPlan = null;
    try { bcpPlan = _bcpPlan(actor, "new", start, end, coverage); } catch (pe) { _log("WARN", "processVLForm", "BCP plan failed: " + pe.message); }
    setVal("BCP Email Subject",      BCP_AUTO_COMPOSE ? (bcpPlan ? bcpPlan.subject : "") : _safeCell(bcpSubject, 300));
    setVal("BCP Notes",              _safeCell(bcpNotes, 1000));
    setVal("Month",                  "'" + monthLabel(start));
    setVal("Status",                 "");
    setVal("Comments",               "");
    setVal("Confirmation on Status", "");
    setVal("Email Sent",             "");
    var newUuid = Utilities.getUuid();
    setVal("Row UUID",               newUuid);

    sheet.appendRow(newRow);
    _notifyNewVL({ email: actor.toLowerCase(), ldap: myLdap, site: config.site, role: config.role, start: start, end: end, days: numDays, vlType: vlType, reason: reason, coverage: coverage.join(", ") });

    CacheService.getScriptCache().remove("vl_events_cache");
    _log("INFO", "processVLForm", "Submission accepted", { ldap: myLdap, start: start, end: end, days: numDays });
    return { status: "success", uuid: newUuid, bcp: bcpPlan };

  } catch(e) {
    if (e.message && e.message.indexOf("timeout") >= 0) {
      _log("WARN", "processVLForm", "Lock timeout", { ldap: myLdap });
      return { status: "error", message: "Server is busy — please try submitting again in a few seconds." };
    }
    _log("ERROR", "processVLForm", e.message, { ldap: myLdap });
    return { status: "error", message: "Server error: " + e.message };
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
    email: '',
    config: { isOpen: true, closedMsg: "",
              isSupervisor: false, isAdmin: false, teamName: "" },
    events: []
  };

  try { payload.email = actor; } catch(e) {}
  try { getVLSheet(); } catch(e) {}
  try { payload.roster = getRoster(); } catch(e) { payload.roster = []; }
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

  if (!payload.config.isSupervisor) payload.events = _redactEvents(payload.events, actor);
  return payload;
}

// ── DUPLICATE CHECK ───────────────────────────────────
function checkVLDuplicate(sheet, email, startYmd, endYmd) {
  try {
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2) return false;
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colMap  = {};
    headers.forEach(function(h, idx) {
      colMap[h.toString().replace(/[^a-zA-Z0-9]/g,'').toLowerCase()] = idx;
    });
    var cE = colMap["emailaddress"], cS = colMap["vlstartdate"], cEn = colMap["vlenddate"], cSt = colMap["status"];
    if (cE === undefined || cS === undefined) return false;

    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var tz   = Session.getScriptTimeZone();
    for (var i = 0; i < data.length; i++) {
      if ((data[i][cE] || "").toString().trim().toLowerCase() !== email.toLowerCase()) continue;
      var st = cSt !== undefined ? (data[i][cSt] || "").toString().toLowerCase() : "";
      if (st.indexOf("denied") >= 0 || st.indexOf("no alloc") >= 0 || st.indexOf("cancel") >= 0) continue;
      var s = _toYmd(data[i][cS], tz);
      if (!s) continue;
      var e = cEn !== undefined ? (_toYmd(data[i][cEn], tz) || s) : s;
      if (startYmd <= e && endYmd >= s) return true;
    }
    return false;
  } catch(err) { return false; }
}

// ── MONTH HELPERS ─────────────────────────────────────
var MO_NAMES_GS = ["January","February","March","April","May","June","July","August","September","October","November","December"];

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
    sheet.appendRow([new Date(), actor, _safeCell(feedbackText, 2000)]);
    _log("INFO", "submitFeedback", "Feedback received", { actor: actor });
    return { status: 'success' };
  } catch(e) {
    _log("ERROR", "submitFeedback", e.message);
    return { status: 'error', message: e.message };
  }
}

function _mailEsc(s) {
  return (s == null ? "" : String(s)).replace(/[&<>"']/g, function(c) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
  });
}

function _fmtMailRange(start, end) {
  end = end || start;
  var p = function(y) { var q = y.split("-"); return new Date(Date.UTC(+q[0], +q[1] - 1, +q[2])); };
  var a = p(start), b = p(end);
  var f = function(dt, fmt) { return Utilities.formatDate(dt, "GMT", fmt); };
  if (start === end) return f(a, "MMMM d, yyyy");
  if (f(a, "yyyy-MM") === f(b, "yyyy-MM")) return f(a, "MMMM d") + " – " + f(b, "d, yyyy");
  if (f(a, "yyyy") === f(b, "yyyy")) return f(a, "MMMM d") + " – " + f(b, "MMMM d, yyyy");
  return f(a, "MMMM d, yyyy") + " – " + f(b, "MMMM d, yyyy");
}

// ── EMAIL DESIGN SYSTEM ───────────────────────────────
var MAIL_FONT   = "-apple-system,BlinkMacSystemFont,'SF Pro Text','Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif";
var WORKDAY_LINK = "https://workday.concentrix.com";

function _mailRows(rows) {
  return rows.map(function(r) {
    return '<tr>' +
      '<td style="padding:15px 0;border-top:1px solid #e8e8ed;font-size:13px;line-height:1.4;color:#86868b;width:36%;vertical-align:top;">' + r[0] + '</td>' +
      '<td style="padding:15px 0;border-top:1px solid #e8e8ed;font-size:15px;line-height:1.4;color:#1d1d1f;font-weight:500;text-align:right;vertical-align:top;">' + r[1] + '</td></tr>';
  }).join("");
}

// A soft grey instruction panel with numbered steps. steps/lead are pre-escaped HTML.
function _mailSteps(title, lead, steps, btnUrl, btnLabel) {
  var li = steps.map(function(s, i) {
    return '<tr><td style="width:30px;vertical-align:top;padding:6px 0;"><div style="width:22px;height:22px;line-height:22px;border-radius:11px;background:#1d1d1f;color:#ffffff;font-size:12px;font-weight:600;text-align:center;">' + (i + 1) + '</div></td>' +
           '<td style="padding:6px 0 6px 6px;font-size:14px;line-height:1.55;color:#1d1d1f;">' + s + '</td></tr>';
  }).join("");
  var btn = btnUrl
    ? '<table role="presentation" style="margin-top:14px;"><tr><td style="background:#ffffff;border:1px solid #d2d2d7;border-radius:980px;"><a href="' + _mailEsc(btnUrl) + '" style="display:inline-block;padding:10px 22px;font-size:13px;font-weight:600;color:#1d1d1f;text-decoration:none;">' + _mailEsc(btnLabel) + '</a></td></tr></table>'
    : "";
  return '<div style="margin-top:20px;background:#f5f5f7;border-radius:18px;padding:24px 26px;">' +
    '<div style="font-size:16px;font-weight:600;letter-spacing:-.01em;color:#1d1d1f;">' + title + '</div>' +
    '<div style="font-size:14px;line-height:1.55;color:#6e6e73;margin:6px 0 14px;">' + lead + '</div>' +
    '<table role="presentation" style="width:100%;border-collapse:collapse;">' + li + '</table>' + btn + '</div>';
}

// o: { accent, eyebrow, title, intro(html), rows[[label,html]], blocks[html], link, linkLabel, footer(html) }
function _mailShell(o) {
  var accent = o.accent || "#1d1d1f";
  var cta = o.link
    ? '<table role="presentation" align="center" style="margin:34px auto 0;"><tr><td style="background:#1d1d1f;border-radius:980px;"><a href="' + _mailEsc(o.link) + '" style="display:inline-block;padding:15px 34px;font-size:15px;font-weight:500;color:#ffffff;text-decoration:none;">' + _mailEsc(o.linkLabel || "Open Pahinga") + '</a></td></tr></table>'
    : "";
  return '<div style="background:#f5f5f7;padding:32px 12px;font-family:' + MAIL_FONT + ';-webkit-font-smoothing:antialiased;">' +
    '<div style="max-width:560px;margin:0 auto;background:#ffffff;border-radius:28px;overflow:hidden;">' +
      '<div style="padding:30px 44px 0;font-size:11px;font-weight:700;letter-spacing:.24em;text-transform:uppercase;color:#1d1d1f;">Pahinga<span style="color:#a1a1a6;font-weight:500;letter-spacing:.14em;">&nbsp;&nbsp;/&nbsp;&nbsp;Play Support</span></div>' +
      '<div style="padding:52px 44px 10px;">' +
        '<div style="font-size:13px;font-weight:600;letter-spacing:.01em;color:' + accent + ';"><span style="display:inline-block;width:8px;height:8px;border-radius:4px;background:' + accent + ';margin-right:9px;vertical-align:1px;"></span>' + _mailEsc(o.eyebrow) + '</div>' +
        '<h1 style="margin:16px 0 0;font-size:40px;line-height:1.08;font-weight:600;letter-spacing:-.03em;color:#1d1d1f;">' + _mailEsc(o.title) + '</h1>' +
        '<p style="margin:16px 0 0;font-size:17px;line-height:1.5;color:#6e6e73;">' + o.intro + '</p>' +
      '</div>' +
      '<div style="padding:26px 44px 44px;">' +
        '<table role="presentation" style="width:100%;border-collapse:collapse;border-bottom:1px solid #e8e8ed;">' + _mailRows(o.rows || []) + '</table>' +
        cta + (o.blocks || []).join("") +
      '</div>' +
      '<div style="padding:22px 44px 30px;background:#fbfbfd;border-top:1px solid #f0f0f3;font-size:12px;line-height:1.6;color:#86868b;">' + (o.footer || "Automated notification from Pahinga · gUP Play Ops") + '</div>' +
    '</div></div>';
}

// m: { ldap, status, start, end, days, vlType, reason, coverage, comments, filed, confirmation, site, role }
function createVLEmailTemplate(m, dashLink) {
  var ls = (m.status || "").toLowerCase();
  var approved = ls.indexOf("approved") >= 0;
  var accent = "#6e6e73", title = "Your leave was updated.";
  if (approved)                         { accent = "#1e8e3e"; title = "Your leave is approved."; }
  else if (ls.indexOf("denied") >= 0)   { accent = "#d93025"; title = "Your leave wasn’t approved."; }
  else if (ls.indexOf("no alloc") >= 0) { accent = "#b06000"; title = "No allocation available."; }
  var range = _fmtMailRange(m.start, m.end) + (m.days > 1 ? " (" + m.days + " days)" : "");

  var blocks = [];
  if (approved) {
    blocks.push(_mailSteps(
      "File this in Workday",
      "Your approval is recorded in Pahinga only. Once it’s approved, file it in Workday to make it official.",
      ["Open <strong>Workday</strong>.",
       "Go to <strong>Absence</strong> &rsaquo; <strong>Request Absence</strong>.",
       "Select <strong>" + _mailEsc(range) + "</strong> and submit."],
      WORKDAY_LINK, "Open Workday"));
    blocks.push(_mailSteps(
      "Worked a holiday?",
      "If you work on a holiday, don’t forget to file CTO in Workday.",
      ["In <strong>Workday</strong>, search <strong>Enter My Time</strong>.",
       "File only the <strong>holiday dates you actually worked</strong>."],
      "", ""));
  }

  return _mailShell({
    accent: accent,
    eyebrow: "VL request · " + (m.status || "Updated"),
    title: title,
    intro: "Hi " + _mailEsc(m.ldap) + ", your leave for <strong style=\"color:#1d1d1f;\">" + _mailEsc(range) + "</strong> has been reviewed.",
    rows: [
      ["Dates", _mailEsc(range)],
      ["Type", _mailEsc(m.vlType || "—")],
      ["Site &amp; role", _mailEsc((m.site || "N/A") + " · " + (m.role || "—"))],
      ["Covering", _mailEsc(m.coverage || "—")],
      ["Manager note", m.comments ? "<span style=\"font-style:italic;font-weight:400;\">“" + _mailEsc(m.comments) + "”</span>" : "<span style=\"color:#86868b;font-weight:400;\">None</span>"]
    ],
    link: dashLink, linkLabel: "Open Pahinga",
    blocks: blocks,
    footer: "Submitted " + _mailEsc(m.filed) + " · Reason: " + _mailEsc(m.reason || "—") + "<br>Ref " + _mailEsc(m.confirmation || "N/A") + " · Automated notification from Pahinga · gUP Play Ops"
  });
}

function sendVLNotifications() {
  var actor = getSessionEmail();
  if (!getFormConfig().isAdmin) return { success: false, message: "Admin access required." };
  if (!_checkRateLimit("sendNotif_" + actor, 3, 300)) {
    return { success: false, message: "Rate limit: wait before sending again." };
  }
  try {
    var sheet = getVLSheet();
    if (sheet.getLastRow() < 2) return { success: false, message: "No data." };
    ensureColumn(sheet, "Status");
    ensureColumn(sheet, "Confirmation on Status");
    var sentColIdx = ensureColumn(sheet, "Email Sent");

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colMap  = {};
    headers.forEach(function(h, i) { colMap[h.toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase()] = i; });
    var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var tz   = Session.getScriptTimeZone();
    var link = "";
    try { link = ScriptApp.getService().getUrl(); } catch (e) {}
    var g = function(row, key) {
      var ix = colMap[key];
      return ix !== undefined && row[ix] != null ? row[ix].toString().trim() : "";
    };

    var sent = 0, failed = 0;
    for (var i = 0; i < data.length; i++) {
      if (MailApp.getRemainingDailyQuota() < 1) break;
      var row    = data[i];
      var status = g(row, "status");
      var ls     = status.toLowerCase();
      var conf   = g(row, "confirmationonstatus");
      if (!status || ls.indexOf("pending") >= 0 || !conf || g(row, "emailsent")) continue;
      var to = g(row, "emailaddress");
      if (!to) continue;
      var start = _toYmd(row[colMap["vlstartdate"]], tz);
      var end   = _toYmd(row[colMap["vlenddate"]], tz) || start;
      if (!start) continue;

      var tsRaw = row[colMap["timestamp"]];
      var filed = Object.prototype.toString.call(tsRaw) === '[object Date]'
        ? Utilities.formatDate(tsRaw, tz, "MMM d, yyyy h:mm a") : (tsRaw == null ? "" : tsRaw.toString());
      var emoji = ls.indexOf("approved") >= 0 ? "✅" : ls.indexOf("denied") >= 0 ? "⛔" : ls.indexOf("no alloc") >= 0 ? "🚫" : "⚠️";

      var html = createVLEmailTemplate({
        ldap: g(row, "ldap") || to.split("@")[0], status: status, start: start, end: end,
        days: _expandDates(start, end).length, vlType: g(row, "vltype"), reason: g(row, "reason"),
        coverage: g(row, "supportcoverage"), comments: g(row, "comments"), filed: filed,
        confirmation: conf, site: g(row, "site"), role: g(row, "role")
      }, link);

      try {
        MailApp.sendEmail({ to: to, subject: emoji + " Google Play VL Calendar Update - " + status, htmlBody: html, name: "Play VL Calendar · Google Play", noReply: true });
        sheet.getRange(i + 2, sentColIdx + 1).setValue(Utilities.formatDate(new Date(), tz, "MM/dd/yy HH:mm") + " by " + actor.split("@")[0]);
        sent++;
        _log("INFO", "sendVLNotifications", "Email sent", { ldap: g(row, "ldap"), status: status });
      } catch (mailErr) {
        failed++;
        _log("ERROR", "sendVLNotifications", "Mail failed for " + g(row, "ldap"), { err: mailErr.message });
      }
    }
    return { success: true, count: sent, failed: failed };
  } catch (e) {
    _log("ERROR", "sendVLNotifications", e.message);
    return { success: false, message: e.message };
  }
}

// ── EMAIL FUNCTIONS (unchanged — kept for compatibility) ─
function getUnsentSummary() {
  if (!getFormConfig().isAdmin) return { count: 0, html: "Admin access required." };
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
        if (lStat.includes("approved")) appr++;
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

// ── LIGHTWEIGHT EVENTS-ONLY ENDPOINT for live sync polling ──
// Returns only the events array — no config, no supervisor lookup.
// Much cheaper on quota than a full getDashboardPayload call.
function getEventsOnly() {
  var res = _getEventsOnlyRaw();
  if (res && res.events && !getFormConfig().isSupervisor) res.events = _redactEvents(res.events, getSessionEmail());
  return res;
}

function _getEventsOnlyRaw() {
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
      return { events: safe };
    } finally {
      if (gotLock) lock.releaseLock();
    }

  } catch(e) {
    _log("ERROR", "getEventsOnly", e.message);
    return { events: null };
  }
}

// ── SUPPORT SCHEDULES (read-only) ─────────────────────
var SCHED_SHEET_NAME  = "Support Schedules";
var SCHED_HEADER_ROW  = 2;   // row with real dates
var SCHED_FIRST_ROW   = 3;   // first support row
var SCHED_LAST_ROW    = 40;  // last support row (same range as heatmap formula)
var SCHED_FIRST_COL   = 4;   // first date column (D)
var SCHED_SHIFT_HOURS = 9;

function _schedNorm(s) {
  s = (s || "").toString().trim();
  if (!s) return "";
  var m = s.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(am|pm)?$/i);
  if (!m) return s.toUpperCase();
  var h = +m[1], mi = +m[2], ap = (m[3] || "").toLowerCase();
  if (ap === "pm" && h < 12) h += 12;
  if (ap === "am" && h === 12) h = 0;
  if (h > 23 || mi > 59) return s.toUpperCase();
  return ("0" + h).slice(-2) + ":" + ("0" + mi).slice(-2);
}

function getScheduleData(startYmd, endYmd) {
  var actor = getSessionEmail();
  var cfg   = getFormConfig();
  if (!cfg.isAdmin) return { success: false, message: "Admin access required." };
  if (!_checkRateLimit("sched_" + actor, 60, 60)) return { success: false, message: "Rate limit exceeded. Please wait." };

  try {
    var rx = /^\d{4}-\d{2}-\d{2}$/;
    startYmd = (startYmd || "").toString();
    endYmd   = (endYmd || startYmd).toString();
    if (!rx.test(startYmd) || !rx.test(endYmd)) return { success: false, message: "Invalid date range." };

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SCHED_SHEET_NAME);
    if (!sheet) return { success: false, message: "Sheet '" + SCHED_SHEET_NAME + "' not found." };
    var tz      = ss.getSpreadsheetTimeZone();
    var lastCol = sheet.getLastColumn();

    var dates = _expandDates(startYmd, endYmd);
    var p = startYmd.split("-");
    var prevDate = Utilities.formatDate(new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] - 1)), "GMT", "yyyy-MM-dd");

    var colOf = {};
    if (lastCol >= SCHED_FIRST_COL) {
      var dateRow = sheet.getRange(SCHED_HEADER_ROW, SCHED_FIRST_COL, 1, lastCol - SCHED_FIRST_COL + 1).getValues()[0];
      dateRow.forEach(function(v, i) {
        if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
          colOf[Utilities.formatDate(v, tz, "yyyy-MM-dd")] = i;
        }
      });
    }

    var hasData = {};
    dates.forEach(function(y) { hasData[y] = colOf[y] !== undefined; });
    var wanted = [prevDate].concat(dates).filter(function(y) { return colOf[y] !== undefined; });

    var nRows = SCHED_LAST_ROW - SCHED_FIRST_ROW + 1;
    var info  = sheet.getRange(SCHED_FIRST_ROW, 1, nRows, 3).getDisplayValues();
    var grid = [], base = 0;
    if (wanted.length) {
      var idxs = wanted.map(function(y) { return colOf[y]; });
      var lo = Math.min.apply(null, idxs), hi = Math.max.apply(null, idxs);
      base = lo;
      grid = sheet.getRange(SCHED_FIRST_ROW, SCHED_FIRST_COL + lo, nRows, hi - lo + 1).getDisplayValues();
    }

    var people = [];
    for (var r = 0; r < nRows; r++) {
      var email = (info[r][1] || "").toString().trim().toLowerCase();
      if (!email) continue;
      var cells = {};
      for (var k = 0; k < wanted.length; k++) {
        cells[wanted[k]] = _schedNorm(grid[r][colOf[wanted[k]] - base]);
      }
      people.push({
        row: SCHED_FIRST_ROW + r,
        nick: (info[r][0] || "").toString().trim(),
        email: email,
        ldap: email.split("@")[0],
        site: (info[r][2] || "").toString().trim().toUpperCase(),
        cells: cells
      });
    }

    var weeks = Object.keys(colOf).filter(function(y) {
      var q = y.split("-");
      return new Date(Date.UTC(+q[0], +q[1] - 1, +q[2])).getUTCDay() === 0;
    }).sort();
    return { success: true, dates: dates, prevDate: prevDate, hasData: hasData, weeks: weeks, people: people, shiftHours: SCHED_SHIFT_HOURS };
  } catch (e) {
    _log("ERROR", "getScheduleData", e.message);
    return { success: false, message: "Server error: " + e.message };
  }
}

// ── SUPPORT SCHEDULES (write) ─────────────────────────
// edits: [{ row, email, ymd, value, old }]  value: "HH:mm" | OFF | LOA | VL | ""
function saveScheduleEdits(edits) {
  var actor = getSessionEmail();
  var cfg   = getFormConfig();
  if (!cfg.isAdmin) return { success: false, message: "Admin access required." };
  if (!Array.isArray(edits) || edits.length === 0) return { success: false, message: "No changes to save." };
  if (edits.length > 800) return { success: false, message: "Too many changes at once (max 800)." };
  if (!_checkRateLimit("schedSave_" + actor, 20, 60)) return { success: false, message: "Rate limit exceeded. Please wait." };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SCHED_SHEET_NAME);
    if (!sheet) return { success: false, message: "Sheet '" + SCHED_SHEET_NAME + "' not found." };
    var tz      = ss.getSpreadsheetTimeZone();
    var lastCol = sheet.getLastColumn();
    var nRows   = SCHED_LAST_ROW - SCHED_FIRST_ROW + 1;

    var colOf = {};
    var dateRow = sheet.getRange(SCHED_HEADER_ROW, SCHED_FIRST_COL, 1, lastCol - SCHED_FIRST_COL + 1).getValues()[0];
    dateRow.forEach(function(v, i) {
      if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
        colOf[Utilities.formatDate(v, tz, "yyyy-MM-dd")] = i;
      }
    });
    var emails = sheet.getRange(SCHED_FIRST_ROW, 2, nRows, 1).getValues();

    var timeRx   = /^([01]\d|2[0-3]):[0-5]\d$/;
    var statuses = { "OFF": 1, "LOA": 1, "VL": 1 };
    var clean = [], rejected = [];
    edits.forEach(function(e) {
      var row   = parseInt(e.row, 10);
      var ymd   = (e.ymd || "").toString();
      var val   = (e.value || "").toString().trim().toUpperCase();
      var old   = (e.old || "").toString().trim().toUpperCase();
      var email = (e.email || "").toString().trim().toLowerCase();
      if (!(row >= SCHED_FIRST_ROW && row <= SCHED_LAST_ROW) || colOf[ymd] === undefined) {
        rejected.push({ row: row, ymd: ymd, reason: "Out of range" }); return;
      }
      if ((emails[row - SCHED_FIRST_ROW][0] || "").toString().trim().toLowerCase() !== email) {
        rejected.push({ row: row, ymd: ymd, reason: "Row changed — please refresh" }); return;
      }
      if (val !== "" && !timeRx.test(val) && !Object.prototype.hasOwnProperty.call(statuses, val)) {
        rejected.push({ row: row, ymd: ymd, reason: "Invalid value" }); return;
      }
      clean.push({ row: row, ymd: ymd, value: val, old: old, email: email });
    });
    if (clean.length === 0) return { success: true, saved: 0, conflicts: [], rejected: rejected };

    var rs = clean.map(function(e) { return e.row; });
    var cs = clean.map(function(e) { return colOf[e.ymd]; });
    var minR = Math.min.apply(null, rs), maxR = Math.max.apply(null, rs);
    var minC = Math.min.apply(null, cs), maxC = Math.max.apply(null, cs);
    var blockRange = sheet.getRange(minR, SCHED_FIRST_COL + minC, maxR - minR + 1, maxC - minC + 1);
    var blockDisp  = blockRange.getDisplayValues();
    var blockVals  = blockRange.getValues();
    var blockFmts  = blockRange.getNumberFormats();

    var saved = 0, conflicts = [], applied = [];
    clean.forEach(function(e) {
      var rIdx = e.row - minR;
      var cIdx = colOf[e.ymd] - minC;
      var cur  = _schedNorm(blockDisp[rIdx][cIdx]);
      if (cur !== e.old) { conflicts.push({ row: e.row, ymd: e.ymd, current: cur }); return; }
      if (e.value === "") {
        blockVals[rIdx][cIdx] = "";
      } else if (timeRx.test(e.value)) {
        blockFmts[rIdx][cIdx] = "H:mm";
        blockVals[rIdx][cIdx] = (parseInt(e.value.slice(0, 2), 10) * 60 + parseInt(e.value.slice(3), 10)) / 1440;
      } else {
        blockFmts[rIdx][cIdx] = "@";
        blockVals[rIdx][cIdx] = e.value;
      }
      saved++;
      applied.push({ r: e.row, y: e.ymd, from: e.old, to: e.value, email: e.email });
    });
    if (saved > 0) {
      blockRange.setNumberFormats(blockFmts);
      blockRange.setValues(blockVals);
    }
    SpreadsheetApp.flush();
    CacheService.getScriptCache().put("sched_ver", String(Date.now()), 21600);

    _notifyScheduleChanges(applied, actor);
    _log("INFO", "saveScheduleEdits", "Schedule edited", { saved: saved, conflicts: conflicts.length, rejected: rejected.length, edits: applied.slice(0, 40) });
    return { success: true, saved: saved, conflicts: conflicts, rejected: rejected };
  } catch (e) {
    _log("ERROR", "saveScheduleEdits", e.message);
    return { success: false, message: "Server error: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── SCHEDULE STARTS (read-only: calendar badges + Insights) ──
// Returns each person's shift start (minutes after midnight) per date.
function getScheduleStarts(startYmd, endYmd) {
  var actor = getSessionEmail();
  var cfg   = getFormConfig();
  if (!cfg.isSupervisor) return { success: false, message: "Managers only." };
  if (!_checkRateLimit("schedStarts_" + actor, 60, 60)) return { success: false, message: "Rate limit exceeded. Please wait." };
  try {
    var rx = /^\d{4}-\d{2}-\d{2}$/;
    startYmd = (startYmd || "").toString();
    endYmd   = (endYmd || startYmd).toString();
    if (!rx.test(startYmd) || !rx.test(endYmd)) return { success: false, message: "Invalid date range." };

    var cache = CacheService.getScriptCache();
    var key   = "sched_st_" + (cache.get("sched_ver") || "0") + "_" + startYmd + "_" + endYmd;
    var hit   = cache.get(key);
    if (hit) { try { return JSON.parse(hit); } catch (e) {} }

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SCHED_SHEET_NAME);
    if (!sheet) return { success: false, message: "Sheet '" + SCHED_SHEET_NAME + "' not found." };
    var tz      = ss.getSpreadsheetTimeZone();
    var lastCol = sheet.getLastColumn();
    var dates   = _expandDates(startYmd, endYmd);
    var p = startYmd.split("-");
    var prevDate = Utilities.formatDate(new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] - 1)), "GMT", "yyyy-MM-dd");

    var colOf = {};
    if (lastCol >= SCHED_FIRST_COL) {
      sheet.getRange(SCHED_HEADER_ROW, SCHED_FIRST_COL, 1, lastCol - SCHED_FIRST_COL + 1).getValues()[0].forEach(function(v, i) {
        if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
          colOf[Utilities.formatDate(v, tz, "yyyy-MM-dd")] = i;
        }
      });
    }
    var hasData = {};
    dates.forEach(function(y) { hasData[y] = colOf[y] !== undefined; });
    var wanted = [prevDate].concat(dates).filter(function(y) { return colOf[y] !== undefined; });

    var starts = {}, site = {};
    if (wanted.length) {
      var nRows = SCHED_LAST_ROW - SCHED_FIRST_ROW + 1;
      var info  = sheet.getRange(SCHED_FIRST_ROW, 2, nRows, 2).getDisplayValues(); // B email, C site
      var idxs  = wanted.map(function(y) { return colOf[y]; });
      var lo = Math.min.apply(null, idxs), hi = Math.max.apply(null, idxs);
      var grid  = sheet.getRange(SCHED_FIRST_ROW, SCHED_FIRST_COL + lo, nRows, hi - lo + 1).getDisplayValues();
      wanted.forEach(function(y) { starts[y] = {}; });
      // Map original shift times for cells already overwritten with "VL"
      var origMap = {};
      var syncSh  = ss.getSheetByName(VL_SYNC_SHEET);
      if (syncSh && syncSh.getLastRow() >= 2) {
        syncSh.getRange(2, 1, syncSh.getLastRow() - 1, 7).getValues().forEach(function(sr) {
          if ((sr[6] || "").toString() === "Applied") {
            var sl = (sr[1] || "").toString().trim().toLowerCase();
            var sy = _toYmd(sr[3], tz);
            var so = _schedNorm(sr[4]);
            if (sl && sy && /^\d\d:\d\d$/.test(so)) origMap[sl + "|" + sy] = so;
          }
        });
      }

      for (var r = 0; r < nRows; r++) {
        var email = (info[r][0] || "").toString().trim().toLowerCase();
        if (!email) continue;
        var ldap = email.split("@")[0];
        site[ldap] = (info[r][1] || "").toString().trim().toUpperCase();
        for (var k = 0; k < wanted.length; k++) {
          var ymd = wanted[k];
          var v   = _schedNorm(grid[r][colOf[ymd] - lo]);
          if (v === "VL" && origMap[ldap + "|" + ymd]) v = origMap[ldap + "|" + ymd];
          if (/^\d\d:\d\d$/.test(v)) starts[ymd][ldap] = parseInt(v.slice(0, 2), 10) * 60 + parseInt(v.slice(3), 10);
        }
      }
    }
    var out = { success: true, dates: dates, prevDate: prevDate, hasData: hasData, starts: starts, site: site, shiftHours: SCHED_SHIFT_HOURS };
    _safeCachePut(key, out, 120);
    return out;
  } catch (e) {
    _log("ERROR", "getScheduleStarts", e.message);
    return { success: false, message: "Server error: " + e.message };
  }
}

// ── APPROVED VL → SUPPORT SCHEDULES ───────────────────
var VL_SYNC_SHEET = "VL Schedule Sync";
var VL_APPROVED_RX = /approved/i;

function _getVLSyncSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(VL_SYNC_SHEET);
  if (!sh) {
    sh = ss.insertSheet(VL_SYNC_SHEET);
    sh.getRange("A1:H1").setValues([["Request UUID","LDAP","Email","Date","Original","Applied","Status","Updated"]])
      .setBackground("#202124").setFontColor("#ffffff").setFontWeight("bold");
    sh.setFrozenRows(1);
    sh.getRange("D:E").setNumberFormat("@");   // keep dates and times as plain text
    _log("INFO", "_getVLSyncSheet", "Created " + VL_SYNC_SHEET + " sheet");
  }
  return sh;
}

function _schedContext() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SCHED_SHEET_NAME);
  if (!sheet) return null;
  var tz      = ss.getSpreadsheetTimeZone();
  var lastCol = sheet.getLastColumn();
  var colOf   = {};
  if (lastCol >= SCHED_FIRST_COL) {
    sheet.getRange(SCHED_HEADER_ROW, SCHED_FIRST_COL, 1, lastCol - SCHED_FIRST_COL + 1).getValues()[0].forEach(function(v, i) {
      if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
        colOf[Utilities.formatDate(v, tz, "yyyy-MM-dd")] = i;
      }
    });
  }
  var nRows  = SCHED_LAST_ROW - SCHED_FIRST_ROW + 1;
  var emails = sheet.getRange(SCHED_FIRST_ROW, 2, nRows, 1).getValues().map(function(r) {
    return (r[0] || "").toString().trim().toLowerCase();
  });
  return { sheet: sheet, colOf: colOf, emails: emails };
}

function _vlInfoFromRow(row, colMap, uuid) {
  var tz    = Session.getScriptTimeZone();
  var email = (row[colMap["emailaddress"]] || "").toString().trim().toLowerCase();
  var start = _toYmd(row[colMap["vlstartdate"]], tz);
  var end   = _toYmd(row[colMap["vlenddate"]], tz) || start;
  if (!email || !start) return null;
  var ldap  = ((row[colMap["ldap"]] || "").toString().trim().toLowerCase()) || email.split("@")[0];
  return { uuid: uuid, ldap: ldap, email: email, start: start, end: end };
}

function _vlRowInfo(sheet, rowNum, uuid) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var colMap  = {};
  headers.forEach(function(h, i) { colMap[h.toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase()] = i; });
  return _vlInfoFromRow(sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0], colMap, uuid);
}

// jobs: [{ uuid, ldap, email, start, end }]  →  writes "VL" over shift times / blanks (today onward)
function _applyVLJobs(jobs) {
  var res = { applied: 0, skipped: 0, noColumn: 0, noRow: 0 };
  if (!jobs || !jobs.length) return res;
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ctx = _schedContext();
    if (!ctx) return res;
    var tz    = Session.getScriptTimeZone();
    var today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    var logSh = _getVLSyncSheet();
    var done  = {};
    var lr    = logSh.getLastRow();
    if (lr >= 2) {
      logSh.getRange(2, 1, lr - 1, 8).getValues().forEach(function(r) {
        if ((r[6] || "").toString() === "Applied") done[r[0] + "|" + _toYmd(r[3], tz)] = 1;
      });
    }
    var width = ctx.sheet.getLastColumn() - SCHED_FIRST_COL + 1;
    var stamp = Utilities.formatDate(new Date(), tz, "MM/dd/yy HH:mm") + " by " + (getSessionEmail() || "system").split("@")[0];
    var logRows = [];

    jobs.forEach(function(job) {
      var ri = ctx.emails.indexOf(job.email);
      if (ri < 0) { res.noRow++; return; }
      var row  = SCHED_FIRST_ROW + ri;
      var vals = ctx.sheet.getRange(row, SCHED_FIRST_COL, 1, width).getDisplayValues()[0];
      _expandDates(job.start, job.end).forEach(function(ymd) {
        if (ymd < today) return;                       // never rewrite the past
        var c = ctx.colOf[ymd];
        if (c === undefined) { res.noColumn++; return; }
        if (done[job.uuid + "|" + ymd]) return;        // already applied for this request/day
        var cur = _schedNorm(vals[c]);
        if (cur === "VL") return;
        if (cur !== "" && !/^\d\d:\d\d$/.test(cur)) { res.skipped++; return; }   // OFF / LOA etc. stay as they are
        ctx.sheet.getRange(row, SCHED_FIRST_COL + c).setValue("VL");
        logRows.push([job.uuid, job.ldap, job.email, ymd, cur, "VL", "Applied", stamp]);
        done[job.uuid + "|" + ymd] = 1;
        res.applied++;
      });
    });

    if (logRows.length) {
      var rng = logSh.getRange(logSh.getLastRow() + 1, 1, logRows.length, 8);
      rng.setNumberFormat("@");
      rng.setValues(logRows);
    }
    SpreadsheetApp.flush();
    if (res.applied) CacheService.getScriptCache().put("sched_ver", String(Date.now()), 21600);
    _log("INFO", "_applyVLJobs", "Approved VL written to schedule", {
      jobs: jobs.length, applied: res.applied, skipped: res.skipped, noColumn: res.noColumn, noRow: res.noRow
    });
    return res;
  } finally {
    lock.releaseLock();
  }
}

// Restores the original shift for every logged "Applied" day of these requests
function _revertVLJobs(uuids) {
  var res = { restored: 0, kept: 0 };
  if (!uuids || !uuids.length) return res;
  var logSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VL_SYNC_SHEET);
  if (!logSh || logSh.getLastRow() < 2) return res;
  var want = {};
  uuids.forEach(function(u) { want[u] = 1; });
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var tz   = Session.getScriptTimeZone();
    var rows = logSh.getRange(2, 1, logSh.getLastRow() - 1, 8).getValues();
    var hit  = [];
    rows.forEach(function(r, i) {
      if (want[r[0]] && (r[6] || "").toString() === "Applied") hit.push(i);
    });
    if (!hit.length) return res;
    var ctx = _schedContext();
    if (!ctx) return res;
    var stamp = Utilities.formatDate(new Date(), tz, "MM/dd/yy HH:mm") + " by " + (getSessionEmail() || "system").split("@")[0];

    hit.forEach(function(i) {
      var r      = rows[i];
      var email  = (r[2] || "").toString().trim().toLowerCase();
      var ymd    = _toYmd(r[3], tz);
      var orig   = _schedNorm((r[4] || "").toString());
      var ri     = ctx.emails.indexOf(email);
      var c      = ctx.colOf[ymd];
      var status = "Reverted";
      if (ri < 0 || c === undefined) {
        status = "Skipped (cell not found)";
      } else {
        var cell = ctx.sheet.getRange(SCHED_FIRST_ROW + ri, SCHED_FIRST_COL + c);
        if (_schedNorm(cell.getDisplayValue()) !== "VL") {
          status = "Kept (cell was changed manually)";
        } else {
          if (orig === "") cell.clearContent();
          else if (/^\d\d:\d\d$/.test(orig)) cell.setNumberFormat("H:mm").setValue((parseInt(orig.slice(0, 2), 10) * 60 + parseInt(orig.slice(3), 10)) / 1440);
          else cell.setValue(orig);
          res.restored++;
        }
      }
      if (status !== "Reverted") res.kept++;
      logSh.getRange(i + 2, 7, 1, 2).setValues([[status, stamp]]);
    });
    SpreadsheetApp.flush();
    CacheService.getScriptCache().put("sched_ver", String(Date.now()), 21600);
    _log("INFO", "_revertVLJobs", "VL removed from schedule", { restored: res.restored, kept: res.kept });
    return res;
  } finally {
    lock.releaseLock();
  }
}

// Called after a single status change. Never throws: the approval itself has already been saved.
function _syncScheduleForStatus(sheet, rowNum, uuid, oldStatus, newStatus) {
  try {
    var wasAppr = VL_APPROVED_RX.test(oldStatus || "");
    var isAppr  = VL_APPROVED_RX.test(newStatus || "");
    if (isAppr) {
      var info = _vlRowInfo(sheet, rowNum, uuid);
      if (!info) return null;
      var r = _applyVLJobs([info]);
      return { action: "applied", applied: r.applied, skipped: r.skipped, noColumn: r.noColumn, noRow: r.noRow };
    }
    if (wasAppr) {
      var rv = _revertVLJobs([uuid]);
      return { action: "reverted", restored: rv.restored, kept: rv.kept };
    }
  } catch (e) {
    _log("ERROR", "_syncScheduleForStatus", e.message, { uuid: uuid });
    return { error: e.message };
  }
  return null;
}

// Called after a batch status change.
function _syncScheduleBatch(sheet, lastCol, values, minRow, rowIdx, uuids, newStatus) {
  try {
    if (VL_APPROVED_RX.test(newStatus || "")) {
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0], cm = {};
      headers.forEach(function(h, i) { cm[h.toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase()] = i; });
      var jobs = [];
      uuids.forEach(function(u) {
        var rn = rowIdx.index[u];
        if (!rn) return;
        var info = _vlInfoFromRow(values[rn - minRow], cm, u);
        if (info) jobs.push(info);
      });
      var r = _applyVLJobs(jobs);
      return { action: "applied", applied: r.applied, skipped: r.skipped, noColumn: r.noColumn, noRow: r.noRow };
    }
    var rv = _revertVLJobs(uuids.filter(function(u) { return rowIdx.index[u]; }));
    return { action: "reverted", restored: rv.restored, kept: rv.kept };
  } catch (e) {
    _log("ERROR", "_syncScheduleBatch", e.message);
    return { error: e.message };
  }
}

// Admin button: write every already-approved request (today onward) into the schedule
function syncApprovedVLToSchedule() {
  var actor = getSessionEmail();
  var cfg   = getFormConfig();
  if (!cfg.isAdmin) return { success: false, message: "Admin access required." };
  if (!_checkRateLimit("vlSync_" + actor, 5, 60)) return { success: false, message: "Rate limit exceeded. Please wait." };
  try {
    var sheet = getVLSheet();
    if (sheet.getLastRow() < 2) return { success: true, requests: 0, applied: 0, skipped: 0, noColumn: 0, noRow: 0 };
    _getRowIndex(sheet);                                   // makes sure every row has a UUID
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colMap  = {};
    headers.forEach(function(h, i) { colMap[h.toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase()] = i; });
    var data  = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    var tz    = Session.getScriptTimeZone();
    var today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    var jobs  = [], approvedAll = {};
    for (var i = 0; i < data.length; i++) {
      var st = colMap["status"] !== undefined ? (data[i][colMap["status"]] || "").toString() : "";
      if (!VL_APPROVED_RX.test(st)) continue;
      var uuid = (data[i][colMap["rowuuid"]] || "").toString().trim();
      if (!uuid) continue;
      approvedAll[uuid] = 1;
      var info = _vlInfoFromRow(data[i], colMap, uuid);
      if (!info || info.end < today) continue;
      jobs.push(info);
    }
    // Anything previously written as "VL" whose request is no longer approved (denied / pending / removed) gets restored
    var stale = [], seenU = {};
    var logSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VL_SYNC_SHEET);
    if (logSh && logSh.getLastRow() >= 2) {
      logSh.getRange(2, 1, logSh.getLastRow() - 1, 7).getValues().forEach(function(lr) {
        var u = (lr[0] || "").toString();
        if (u && (lr[6] || "").toString() === "Applied" && !approvedAll[u] && !seenU[u]) { seenU[u] = 1; stale.push(u); }
      });
    }
    var rv = _revertVLJobs(stale);
    var r = _applyVLJobs(jobs);
    _log("INFO", "syncApprovedVLToSchedule", "Manual sync", { requests: jobs.length, applied: r.applied, restored: rv.restored });
    return { success: true, requests: jobs.length, applied: r.applied, restored: rv.restored, skipped: r.skipped, noColumn: r.noColumn, noRow: r.noRow };
  } catch (e) {
    _log("ERROR", "syncApprovedVLToSchedule", e.message);
    return { success: false, message: "Server error: " + e.message };
  }
}

// ── DELETE / ARCHIVE REQUEST ──────────────────────────
function deleteVLRequest(eventUuid) {
  var actor = getSessionEmail();
  if (!actor) return { success: false, message: "Not authenticated." };

  if (!_checkRateLimit("delete_" + actor, 10, 60)) {
    return { success: false, message: "Rate limit exceeded. Please wait before trying again." };
  }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
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
    if (!isActorAdmin && !isActorSup && VL_APPROVED_RX.test(status)) {
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
    var aHeaders = archiveSheet.getRange(1, 1, 1, archiveSheet.getLastColumn()).getValues()[0];
    var normH = function(s) { return s.toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase(); };
    var aMap = {};
    headers.forEach(function(h, i) { aMap[normH(h)] = rowData[i]; });
    aMap["deletedby"] = actor;
    aMap["deletedat"] = deletedAt;
    archiveSheet.appendRow(aHeaders.map(function(h) { var v = aMap[normH(h)]; return v == null ? "" : v; }));

    // ── Delete row from main sheet ──
    try { _revertVLJobs([eventUuid]); } catch (revErr) { _log("ERROR", "deleteVLRequest", "Schedule revert failed: " + revErr.message, { uuid: eventUuid }); }
    sheet.deleteRow(rowNum);
    CacheService.getScriptCache().remove("vl_events_cache");

    _log("INFO", "deleteVLRequest", "Request archived and deleted", { uuid: eventUuid, actor: actor, row: rowNum });
    return { success: true };

  } catch(e) {
    _log("ERROR", "deleteVLRequest", e.message, { uuid: eventUuid });
    return { success: false, message: "Server error: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── SCHEDULE VIEW (read-only, any registered user) ────
function getMyScheduleView(startYmd, endYmd) {
  var actor = getSessionEmail();
  var cfg   = getFormConfig();
  if (!cfg.role) return { success: false, message: "You are not registered in the Roles list." };
  if (!_checkRateLimit("mySched_" + actor, 40, 60)) return { success: false, message: "Rate limit exceeded. Please wait." };
  try {
    var rx = /^\d{4}-\d{2}-\d{2}$/;
    startYmd = (startYmd || "").toString();
    endYmd   = (endYmd || startYmd).toString();
    if (!rx.test(startYmd) || !rx.test(endYmd)) return { success: false, message: "Invalid date range." };

    var cache = CacheService.getScriptCache();
    var key   = "sched_my_" + (cache.get("sched_ver") || "0") + "_" + startYmd + "_" + endYmd;
    var hit   = cache.get(key);
    if (hit) { try { return JSON.parse(hit); } catch (e) {} }

    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SCHED_SHEET_NAME);
    if (!sheet) return { success: false, message: "Schedules have not been published yet." };
    var tz      = ss.getSpreadsheetTimeZone();
    var lastCol = sheet.getLastColumn();
    var dates   = _expandDates(startYmd, endYmd).slice(0, 35);
    var p = startYmd.split("-");
    var prevDate = Utilities.formatDate(new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] - 1)), "GMT", "yyyy-MM-dd");

    var colOf = {};
    if (lastCol >= SCHED_FIRST_COL) {
      sheet.getRange(SCHED_HEADER_ROW, SCHED_FIRST_COL, 1, lastCol - SCHED_FIRST_COL + 1).getValues()[0].forEach(function(v, i) {
        if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
          colOf[Utilities.formatDate(v, tz, "yyyy-MM-dd")] = i;
        }
      });
    }
    var hasData = {};
    dates.forEach(function(y) { hasData[y] = colOf[y] !== undefined; });
    var wanted = [prevDate].concat(dates).filter(function(y) { return colOf[y] !== undefined; });

    var nRows = SCHED_LAST_ROW - SCHED_FIRST_ROW + 1;
    var info  = sheet.getRange(SCHED_FIRST_ROW, 1, nRows, 3).getDisplayValues();
    var grid = [], base = 0;
    if (wanted.length) {
      var idxs = wanted.map(function(y) { return colOf[y]; });
      var lo = Math.min.apply(null, idxs), hi = Math.max.apply(null, idxs);
      base = lo;
      grid = sheet.getRange(SCHED_FIRST_ROW, SCHED_FIRST_COL + lo, nRows, hi - lo + 1).getDisplayValues();
    }

    var people = [];
    for (var r = 0; r < nRows; r++) {
      var email = (info[r][1] || "").toString().trim().toLowerCase();
      if (!email) continue;
      var cells = {};
      for (var k = 0; k < wanted.length; k++) {
        cells[wanted[k]] = _schedNorm(grid[r][colOf[wanted[k]] - base]);
      }
      people.push({
        nick:  (info[r][0] || "").toString().trim(),
        ldap:  email.split("@")[0],
        site:  (info[r][2] || "").toString().trim().toUpperCase(),
        cells: cells
      });
    }
    var out = { success: true, dates: dates, prevDate: prevDate, hasData: hasData, people: people, shiftHours: SCHED_SHIFT_HOURS };
    _safeCachePut(key, out, 120);
    return out;
  } catch (e) {
    _log("ERROR", "getMyScheduleView", e.message);
    return { success: false, message: "Server error: " + e.message };
  }
}

// ── ALERT EMAILS / CANCELLATION REQUESTS / AUDIT LOG ──
// Optional switches in the FormConfig tab (Setting | Value): NotifyNewVL, NotifyScheduleChanges  (TRUE / FALSE, default TRUE)
function _cfgFlag(name, def) {
  try {
    var vals = getOrCreateConfigSheet().getRange("A2:B20").getValues();
    for (var i = 0; i < vals.length; i++) {
      if ((vals[i][0] || "").toString().trim() === name) {
        var v = String(vals[i][1] == null ? "" : vals[i][1]).trim().toUpperCase();
        if (v === "TRUE") return true;
        if (v === "FALSE") return false;
      }
    }
  } catch (e) {}
  return def;
}

function _managerEmails(site) {
  var out = [];
  try {
    var vals = getOrCreateSupervisorSheet().getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      var email  = (vals[i][1] || "").toString().trim().toLowerCase();
      var access = (vals[i][6] || "").toString().trim().toLowerCase();
      var s      = (vals[i][2] || "").toString().trim().toUpperCase();
      if (!email) continue;
      if (access === "admin" || (access === "manager" && (!site || s === site))) {
        if (out.indexOf(email) < 0) out.push(email);
      }
    }
  } catch (e) {}
  return out;
}

function _sendMgrMail(to, subject, html) {
  if (!to.length || MailApp.getRemainingDailyQuota() < to.length) return;
  MailApp.sendEmail({ to: to.join(","), subject: subject, htmlBody: html, name: "Play VL Calendar · Google Play", noReply: true });
}

// rows: [[label, htmlValue]] (values must already be escaped). color = accent dot colour; bg is ignored (kept for older callers).
function _infoMail(badge, color, bg, title, intro, rows, link, linkLabel) {
  return _mailShell({
    accent: color, eyebrow: badge, title: title, intro: intro,
    rows: rows, link: link, linkLabel: linkLabel || "Open Pahinga"
  });
}

function _notifyNewVL(d) {
  try {
    if (!_cfgFlag("NotifyNewVL", true)) return;
    var to = _managerEmails(d.site).filter(function(e) { return e !== d.email; });
    if (!to.length) return;
    var link = ""; try { link = ScriptApp.getService().getUrl(); } catch (e) {}
    var range = _fmtMailRange(d.start, d.end) + (d.days > 1 ? " (" + d.days + " days)" : "");
    var html = _infoMail("Pending review", "#1a73e8", "#e8f0fe", "New VL request",
      '<strong>' + _mailEsc(d.ldap) + '</strong> filed leave for <strong>' + _mailEsc(range) + '</strong>.',
      [["Dates", _mailEsc(range)], ["VL type", _mailEsc(d.vlType)], ["Site / role", _mailEsc(d.site + " • " + d.role)],
       ["Covering", _mailEsc(d.coverage)], ["Reason", _mailEsc(d.reason)]],
      link, "Open Pending Queue");
    _sendMgrMail(to, "🗓️ New VL request — " + d.ldap + " (" + range + ")", html);
  } catch (e) { _log("ERROR", "_notifyNewVL", e.message); }
}

function _notifyScheduleChanges(applied, actor) {
  try {
    if (!applied || !applied.length || !_cfgFlag("NotifyScheduleChanges", true)) return;
    var tz    = Session.getScriptTimeZone();
    var today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    var me    = (actor || "").toLowerCase();
    var link = ""; try { link = ScriptApp.getService().getUrl(); } catch (e) {}
    var by = {};
    applied.forEach(function(a) {
      if (!a.email || a.y < today || a.email === me) return;
      (by[a.email] = by[a.email] || []).push(a);
    });
    var fmt = function(y) { var p = y.split("-"); return Utilities.formatDate(new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])), "GMT", "EEE, MMM d"); };
    Object.keys(by).forEach(function(email) {
      if (MailApp.getRemainingDailyQuota() < 1) return;
      var list = by[email].sort(function(a, b) { return a.y < b.y ? -1 : a.y > b.y ? 1 : 0; });
      var rows = list.slice(0, 25).map(function(a) {
        return [_mailEsc(fmt(a.y)), '<span style="color:#80868b;text-decoration:line-through;">' + _mailEsc(a.from || "blank") + '</span> → <strong>' + _mailEsc(a.to || "cleared") + '</strong>'];
      });
      if (list.length > 25) rows.push(["…", "+" + (list.length - 25) + " more change(s)"]);
      var html = _infoMail("Schedule update", "#b06000", "#fef7e0", "Your schedule changed",
        list.length + ' day' + (list.length > 1 ? 's' : '') + ' on your schedule ' + (list.length > 1 ? 'were' : 'was') + ' updated by <strong>' + _mailEsc(me.split("@")[0]) + '</strong>.',
        rows, link, "View my schedule");
      try {
        MailApp.sendEmail({ to: email, subject: "🗓️ Your schedule was updated", htmlBody: html, name: "Play VL Calendar · Google Play", noReply: true });
      } catch (m) { _log("ERROR", "_notifyScheduleChanges", "Mail failed for " + email, { err: m.message }); }
    });
  } catch (e) { _log("ERROR", "_notifyScheduleChanges", e.message); }
}

// Owner asks to cancel an already-approved VL
function requestVLCancellation(eventUuid, note, days) {
  var actor = getSessionEmail();
  if (!actor) return { success: false, message: "Not authenticated." };
  if (!_checkRateLimit("cancelReq_" + actor, 5, 300)) return { success: false, message: "Too many requests. Please wait a few minutes." };
  note = (note || "").toString().trim().slice(0, 300);
  if (note.length < 3) return { success: false, message: "Please give a short reason." };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var sheet  = getVLSheet();
    var cIdx   = ensureColumn(sheet, "Cancel Request") + 1;
    var cdIdx  = ensureColumn(sheet, "Cancel Days") + 1;
    var rowNum = _getRowIndex(sheet).index[eventUuid];
    if (!rowNum) return { success: false, message: "Request not found. Please refresh the calendar." };
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0], colMap = {};
    headers.forEach(function(h, i) { colMap[h.toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase()] = i; });
    var rv = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
    var tz = Session.getScriptTimeZone();
    if ((rv[colMap["emailaddress"]] || "").toString().trim().toLowerCase() !== actor.toLowerCase())
      return { success: false, message: "You can only request cancellation of your own VL." };
    if (!VL_APPROVED_RX.test((rv[colMap["status"]] || "").toString()))
      return { success: false, message: "Only approved VL needs a cancellation request. Pending requests can be removed directly." };
    if ((rv[colMap["cancelrequest"]] || "").toString().trim())
      return { success: false, message: "A cancellation request is already waiting for your supervisor." };
    var start = _toYmd(rv[colMap["vlstartdate"]], tz), end = _toYmd(rv[colMap["vlenddate"]], tz) || start;
    var today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    if (end < today) return { success: false, message: "This leave has already ended." };

    // Optional: only some days. Empty / all upcoming days = the whole request.
    var all = _expandDates(start, end), pick = {};
    (Array.isArray(days) ? days : []).forEach(function(d) { pick[d] = true; });
    var cd = all.filter(function(d) { return pick[d] && d >= today; });
    if (Array.isArray(days) && days.length && !cd.length) return { success: false, message: "Pick at least one upcoming day." };
    var partial = cd.length > 0 && cd.length < all.length;

    var safe  = /^[=+\-@]/.test(note) ? "'" + note : note;
    var value = safe + " (" + Utilities.formatDate(new Date(), tz, "MM/dd/yy HH:mm") + ")";
    sheet.getRange(rowNum, cIdx).setValue(value);
    var cdCell = sheet.getRange(rowNum, cdIdx);
    cdCell.setNumberFormat("@");
    cdCell.setValue(partial ? cd.join(",") : "");
    CacheService.getScriptCache().remove("vl_events_cache");
    _log("INFO", "requestVLCancellation", "Cancellation requested", { uuid: eventUuid, actor: actor, days: partial ? cd.length : "all" });

    try {
      var site = (rv[colMap["site"]] || "").toString().trim().toUpperCase();
      var ldap = (rv[colMap["ldap"]] || "").toString().trim() || actor.split("@")[0];
      var link = ""; try { link = ScriptApp.getService().getUrl(); } catch (e) {}
      var range = _fmtMailRange(start, end);
      var rows = [["Dates", _mailEsc(range)]];
      if (partial) rows.push(["Days to cancel", _mailEsc(cd.map(function(d) { return _fmtMailRange(d, d); }).join(" · "))]);
      rows.push(["Type", _mailEsc(rv[colMap["vltype"]])], ["Site &amp; role", _mailEsc(site + " · " + rv[colMap["role"]])], ["Reason", _mailEsc(note)]);
      var html = _infoMail("Cancellation request", "#b06000", "", "VL cancellation requested.",
        '<strong style="color:#1d1d1f;">' + _mailEsc(ldap) + '</strong> asked to cancel ' + (partial ? "part of" : "") + ' an approved leave.',
        rows, link, "Review in Pahinga");
      _sendMgrMail(_managerEmails(site).filter(function(e) { return e !== actor.toLowerCase(); }), "VL cancellation request — " + ldap + " (" + range + ")", html);
    } catch (m) { _log("ERROR", "requestVLCancellation", "Mail failed: " + m.message); }
    return { success: true, value: value, days: partial ? cd : [] };
  } catch (e) {
    _log("ERROR", "requestVLCancellation", e.message);
    return { success: false, message: "Server error: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

// Manager keeps the VL (clears the cancellation request)
function dismissCancelRequest(eventUuid) {
  var actor = getSessionEmail();
  if (!getFormConfig().isSupervisor) return { success: false, message: "Managers only." };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var sheet  = getVLSheet();
    var cIdx   = ensureColumn(sheet, "Cancel Request") + 1;
    var rowNum = _getRowIndex(sheet).index[eventUuid];
    if (!rowNum) return { success: false, message: "Request not found. Please refresh the calendar." };
    sheet.getRange(rowNum, cIdx).clearContent();
    sheet.getRange(rowNum, ensureColumn(sheet, "Cancel Days") + 1).clearContent();
    CacheService.getScriptCache().remove("vl_events_cache");
    _log("INFO", "dismissCancelRequest", "Cancellation request dismissed", { uuid: eventUuid, actor: actor });
    return { success: true };
  } catch (e) {
    _log("ERROR", "dismissCancelRequest", e.message);
    return { success: false, message: "Server error: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

// ── REMOVE SPECIFIC DAYS FROM A MULTI-DAY REQUEST ─────
// Owner: pending requests only. Supervisors/admins: any status (this is how a day-level cancellation is approved).
function _getArchiveSheet(ss, headers) {
  var sh = ss.getSheetByName("Deleted Requests");
  if (!sh) {
    sh = ss.insertSheet("Deleted Requests");
    var ah = headers.concat(["Deleted By", "Deleted At"]);
    sh.getRange(1, 1, 1, ah.length).setValues([ah]).setBackground("#ea4335").setFontColor("#ffffff").setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function removeVLDays(eventUuid, days, note) {
  var actor = getSessionEmail();
  if (!actor) return { success: false, message: "Not authenticated." };
  if (!_checkRateLimit("rmDays_" + actor, 10, 60)) return { success: false, message: "Rate limit exceeded. Please wait before trying again." };
  var rx = /^\d{4}-\d{2}-\d{2}$/, pick = {}, n = 0;
  (Array.isArray(days) ? days : []).forEach(function(d) {
    d = (d == null ? "" : d).toString();
    if (rx.test(d) && !pick[d]) { pick[d] = true; n++; }
  });
  if (!n) return { success: false, message: "Pick at least one day." };

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var ss     = SpreadsheetApp.getActiveSpreadsheet();
    var sheet  = getVLSheet();
    var rowNum = _getRowIndex(sheet).index[eventUuid];
    if (!rowNum) return { success: false, message: "Request not found. Please refresh the calendar." };

    var cIdx  = ensureColumn(sheet, "Cancel Request") + 1;
    var cdIdx = ensureColumn(sheet, "Cancel Days") + 1;
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0], colMap = {};
    var normH = function(s) { return s.toString().replace(/[^a-zA-Z0-9]/g, '').toLowerCase(); };
    headers.forEach(function(h, i) { colMap[normH(h)] = i; });
    var rowVals = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
    var tz      = Session.getScriptTimeZone();

    var cfg     = getFormConfig();
    var isMgr   = cfg.isSupervisor;
    var owner   = (rowVals[colMap["emailaddress"]] || "").toString().trim().toLowerCase();
    var isOwner = owner === actor.toLowerCase();
    if (!isOwner && !isMgr) {
      _log("WARN", "removeVLDays", "Unauthorized attempt", { actor: actor, owner: owner, uuid: eventUuid });
      return { success: false, message: "You can only change your own requests unless you are a supervisor or admin." };
    }
    var status   = (rowVals[colMap["status"]] || "").toString().trim();
    var approved = VL_APPROVED_RX.test(status);
    if (approved && !isMgr) return { success: false, message: "Approved leave can’t be changed directly. Send a cancellation request instead." };

    var start = _toYmd(rowVals[colMap["vlstartdate"]], tz);
    var end   = _toYmd(rowVals[colMap["vlenddate"]], tz) || start;
    var all   = _expandDates(start, end);
    var today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    var rm    = all.filter(function(d) { return pick[d]; });
    if (!rm.length) return { success: false, message: "Those days are not part of this request." };
    if (!isMgr && rm.some(function(d) { return d < today; })) return { success: false, message: "Days that have already passed can’t be removed." };
    if (rm.length === all.length) return { success: false, message: "That would remove every day. Use Remove Request instead." };

    // Contiguous runs of removed / kept days
    var segs = [];
    all.forEach(function(ymd) {
      var gone = !!pick[ymd], last = segs[segs.length - 1];
      if (last && last.gone === gone) { last.end = ymd; last.days++; }
      else segs.push({ start: ymd, end: ymd, days: 1, gone: gone });
    });
    var keepSegs = segs.filter(function(s) { return !s.gone; });
    var goneSegs = segs.filter(function(s) { return s.gone; });

    // 1) Archive the removed days
    var archive  = _getArchiveSheet(ss, headers);
    var aHeaders = archive.getRange(1, 1, 1, archive.getLastColumn()).getValues()[0];
    var stampNow = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
    goneSegs.forEach(function(s) {
      var aMap = {};
      headers.forEach(function(h, i) { aMap[normH(h)] = rowVals[i]; });
      aMap["vlstartdate"] = s.start; aMap["vlenddate"] = s.end; aMap["days"] = s.days;
      aMap["deletedby"] = actor; aMap["deletedat"] = stampNow;
      archive.appendRow(aHeaders.map(function(h) { var v = aMap[normH(h)]; return v == null ? "" : v; }));
    });

    // 2) Put the original shifts back (approved only); re-applied below for the days that stay
    if (approved) {
      try { _revertVLJobs([eventUuid]); } catch (re) { _log("ERROR", "removeVLDays", "Schedule revert failed: " + re.message, { uuid: eventUuid }); }
    }

    // 3) Rewrite the request as the kept segments
    var monthLabel = function(s) { var p = s.split("-"); return MO_NAMES_GS[parseInt(p[1], 10) - 1] + " " + p[0]; };
    var ldapVal  = (rowVals[colMap["ldap"]] || "").toString().trim().toLowerCase() || owner.split("@")[0];
    var applyJobs = [];
    keepSegs.forEach(function(seg, idx) {
      var segRow  = rowVals.slice();
      var segUuid = idx === 0 ? eventUuid : Utilities.getUuid();
      segRow[colMap["vlstartdate"]] = seg.start;
      segRow[colMap["vlenddate"]]   = seg.end;
      if (colMap["days"]  !== undefined) segRow[colMap["days"]]  = seg.days;
      if (colMap["month"] !== undefined) segRow[colMap["month"]] = "'" + monthLabel(seg.start);
      segRow[colMap["rowuuid"]] = segUuid;
      segRow[cIdx - 1]  = "";
      segRow[cdIdx - 1] = "";
      if (idx === 0) sheet.getRange(rowNum, 1, 1, lastCol).setValues([segRow]);
      else           sheet.appendRow(segRow);
      if (approved) applyJobs.push({ uuid: segUuid, ldap: ldapVal, email: owner, start: seg.start, end: seg.end });
    });
    SpreadsheetApp.flush();
    var schedRes = applyJobs.length ? _applyVLJobs(applyJobs) : null;
    CacheService.getScriptCache().remove("vl_events_cache");
    var fresh = getCalendarData();
    _safeCachePut("vl_events_cache", fresh, 15);
    _log("INFO", "removeVLDays", "Days removed", { uuid: eventUuid, actor: actor, removed: rm.length, kept: all.length - rm.length, status: status });

    // 4) Let the other side know
    try {
      var link = ""; try { link = ScriptApp.getService().getUrl(); } catch (e) {}
      var site = (rowVals[colMap["site"]] || "").toString().trim().toUpperCase();
      var pretty = function(list) { return list.map(function(s) { return _fmtMailRange(s.start, s.end); }).join(" · "); };
      var rows = [["Removed", _mailEsc(pretty(goneSegs))], ["Still on", _mailEsc(pretty(keepSegs))], ["Type", _mailEsc(rowVals[colMap["vltype"]] || "—")]];
      if (!isOwner) {
        if (MailApp.getRemainingDailyQuota() > 0) {
          MailApp.sendEmail({ to: owner, subject: "Google Play VL Calendar Update - Days cancelled", name: "Play VL Calendar · Google Play", noReply: true,
            htmlBody: _infoMail("VL request · Updated", "#b06000", "", "Some days were cancelled.",
              "A supervisor removed <strong style=\"color:#1d1d1f;\">" + rm.length + " day" + (rm.length > 1 ? "s" : "") + "</strong> from your leave.", rows, link, "Open Pahinga") });
        }
      } else if (_cfgFlag("NotifyNewVL", true)) {
        var to = _managerEmails(site).filter(function(e) { return e !== owner; });
        _sendMgrMail(to, "Pending VL changed — " + ldapVal,
          _infoMail("Pending review", "#1a73e8", "", "A pending request changed.",
            "<strong style=\"color:#1d1d1f;\">" + _mailEsc(ldapVal) + "</strong> removed " + rm.length + " day" + (rm.length > 1 ? "s" : "") + " from a pending request.", rows, link, "Open Pending Queue"));
      }
    } catch (m) { _log("ERROR", "removeVLDays", "Mail failed: " + m.message); }

    return {
      success: true, removed: rm.length, events: isMgr ? fresh : _redactEvents(fresh, actor),
      sched: schedRes ? { action: "applied", applied: schedRes.applied, skipped: schedRes.skipped, noColumn: schedRes.noColumn, noRow: schedRes.noRow } : null
    };
  } catch (e) {
    _log("ERROR", "removeVLDays", e.message, { uuid: eventUuid });
    return { success: false, message: "Server error: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

// Admin Console → Audit Log
function getSystemLog(limit) {
  var actor = getSessionEmail();
  if (!getFormConfig().isAdmin) return { success: false, message: "Admin access required." };
  if (!_checkRateLimit("syslog_" + actor, 20, 60)) return { success: false, message: "Rate limit exceeded. Please wait." };
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_NAME);
    if (!sheet || sheet.getLastRow() < 2) return { success: true, rows: [], total: 0 };
    var n = Math.max(1, Math.min(parseInt(limit, 10) || 300, 1000));
    var last = sheet.getLastRow(), start = Math.max(2, last - n + 1);
    var vals = sheet.getRange(start, 1, last - start + 1, 6).getDisplayValues();
    var rows = vals.reverse().map(function(r) { return { ts: r[0], level: r[1], fn: r[2], actor: r[3], msg: r[4], extra: r[5] }; });
    return { success: true, rows: rows, total: last - 1 };
  } catch (e) {
    return { success: false, message: "Server error: " + e.message };
  }
}

// ══ ONE-TIME MIGRATION: old VL form -> Support VL Requests ══
var MIG_URL     = "https://docs.google.com/spreadsheets/d/1woHaccV8wkrNZATwrzSDIrolOoWOi0dPQblTE442Kxs/edit?resourcekey=0-0Zb_uSKqqjE963DcVD7D-A";
var MIG_TAB     = "Form Responses 1";
var MIG_MIN_END = "";   // "" = only leave ending today or later. Or a date like "2026-10-01" to include older rows.

function migrateOldVL_DryRun() { _migrateOldVL(true); }
function migrateOldVL()        { _migrateOldVL(false); }

function _migrateOldVL(dry) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var tz     = Session.getScriptTimeZone();
    var today  = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
    var minEnd = MIG_MIN_END || today;

    var srcSheet = SpreadsheetApp.openByUrl(MIG_URL).getSheetByName(MIG_TAB);
    if (!srcSheet) throw new Error("Tab not found: " + MIG_TAB);
    var src = srcSheet.getDataRange().getValues();

    // Roles lookup (email -> site / role) and known LDAPs
    var roles = {}, known = {};
    getOrCreateSupervisorSheet().getDataRange().getValues().forEach(function(r, i) {
      if (i === 0) return;
      var em = (r[1] || "").toString().trim().toLowerCase();
      if (!em) return;
      var ld = ((r[3] || "").toString().trim().toLowerCase() || em.split("@")[0]);
      roles[em] = { site: (r[2] || "").toString().trim().toUpperCase(), role: (r[4] || "").toString().trim(), ldap: ld };
      known[ld] = 1;
    });

    // Target sheet
    var sheet = getVLSheet();
    ["Timestamp","Email Address","LDAP","Role","Site","VL Start Date","VL End Date","Days","VL Type","Reason",
     "Supporting Docs","Support Coverage","VL Credits Snip","BCP Email Subject","Month","Status","Comments",
     "Confirmation on Status","Email Sent","Row UUID"].forEach(function(h) { ensureColumn(sheet, h); });
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var cm = {};
    headers.forEach(function(h, i) { cm[h.toString().replace(/[^a-zA-Z0-9]/g, "").toLowerCase()] = i; });

    // Existing rows -> duplicate guard (email|start|end)
    var seen = {};
    if (sheet.getLastRow() >= 2) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues().forEach(function(r) {
        var e = (r[cm["emailaddress"]] || "").toString().trim().toLowerCase();
        var s = _toYmd(r[cm["vlstartdate"]], tz);
        var n = _toYmd(r[cm["vlenddate"]], tz) || s;
        if (e && s) seen[e + "|" + s + "|" + n] = 1;
      });
    }

    var safe = function(s, max) {
      s = (s == null ? "" : String(s)).trim().slice(0, max || 1000);
      return /^[=+\-@]/.test(s) ? "'" + s : s;
    };
    var guessType = function(t) {
      t = (t || "").toLowerCase();
      if (/emergency|hospital|accident|passed away|wake|funeral/.test(t)) return "Emergency";
      if (/out of town|vacation|travel|trip|province|flight|beach/.test(t)) return "Out of Town / Vacation";
      if (/family|birthday|reunion|gathering|anniversary|christmas|holiday/.test(t)) return "Family Gathering";
      return "Personal";
    };
    var roleMap = function(r) { return /team lead|supervisor/i.test(r) ? "Supervisor" : r; };

    var out = [], stats = { added: 0, old: 0, dup: 0, noRoster: [], noDates: 0 };

    for (var i = 1; i < src.length; i++) {
      var r = src[i];
      var email = (r[1] || "").toString().trim().toLowerCase();
      var start = _toYmd(r[3], tz);
      if (!email || !start) { if (email || r[3]) stats.noDates++; continue; }
      var end = _toYmd(r[4], tz) || start;
      if (end < start) end = start;
      if (end < minEnd) { stats.old++; continue; }
      if (seen[email + "|" + start + "|" + end]) { stats.dup++; continue; }

      var info = roles[email];
      if (!info || !info.site) { stats.noRoster.push(email); continue; }

      var reason   = (r[5] || "").toString();
      var covRaw   = (r[6] || "").toString().toLowerCase();
      var covList  = covRaw.split(/[^a-z0-9._-]+/).filter(function(t) { return t && known[t] && t !== info.ldap; });
      covList = covList.filter(function(t, k, a) { return a.indexOf(t) === k; });
      var coverage = covList.length ? covList.join(", ") : safe(covRaw, 300);
      var status   = (r[9] || "").toString().trim();
      var days     = _expandDates(start, end);
      var p        = start.split("-");

      var row = new Array(headers.length).fill("");
      var set = function(name, v) { var ix = cm[name.replace(/[^a-zA-Z0-9]/g, "").toLowerCase()]; if (ix !== undefined) row[ix] = v; };
      set("Timestamp", r[0] || new Date());
      set("Email Address", email);
      set("LDAP", info.ldap);
      set("Role", info.role ? roleMap(info.role) : roleMap((r[2] || "").toString().trim()));
      set("Site", info.site);
      set("VL Start Date", start);
      set("VL End Date", end);
      set("Days", days.length);
      set("VL Type", guessType(reason));
      set("Reason", safe(reason, 1000));
      set("Support Coverage", coverage);
      set("VL Credits Snip", safe(r[7], 500));
      set("BCP Email Subject", safe(r[8], 300));
      set("Month", "'" + MO_NAMES_GS[parseInt(p[1], 10) - 1] + " " + p[0]);
      set("Status", status);
      set("Comments", safe(r[10], 500));
      if (status) {
        set("Confirmation on Status", "Migrated from old sheet");
        set("Email Sent", "Migrated (no email)");   // keeps Send Notifications from emailing old decisions
      }
      set("Row UUID", Utilities.getUuid());
      out.push(row);
      seen[email + "|" + start + "|" + end] = 1;
      stats.added++;
    }

    var msg = (dry ? "DRY RUN — nothing written\n\n" : "MIGRATION DONE\n\n") +
      "To add: " + stats.added + "\nSkipped (ended before " + minEnd + "): " + stats.old +
      "\nSkipped (already in Pahinga): " + stats.dup + "\nSkipped (no dates): " + stats.noDates +
      "\nSkipped (not in Roles tab): " + stats.noRoster.length +
      (stats.noRoster.length ? "\n  " + stats.noRoster.filter(function(e, k, a) { return a.indexOf(e) === k; }).join("\n  ") : "");

    if (!dry && out.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, out.length, headers.length).setValues(out);
      SpreadsheetApp.flush();
      CacheService.getScriptCache().remove("vl_events_cache");
      _log("INFO", "migrateOldVL", "Migrated from old VL sheet", { added: out.length, minEnd: minEnd });
    }
    Logger.log(msg);
    try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
  } finally {
    lock.releaseLock();
  }
}

function onEdit(e) {
  try {
    if (e && e.range && e.range.getSheet().getName() === SCHED_SHEET_NAME) {
      CacheService.getScriptCache().put("sched_ver", String(Date.now()), 21600);
    }
  } catch (err) {}
}

// ── BCP EMAIL PLAN ────────────────────────────────────
// Keep BCP_TEST_MODE = true while testing: recipients become ONLY the sender.
var BCP_TEST_MODE = true;

// Master switch for the auto-compose BCP flow. Keep false until stakeholders approve.
// (Also flip BCP_AUTO_COMPOSE in VLIndex.html.)
var BCP_AUTO_COMPOSE = false;

function _bcpPlan(actor, kind, start, end, coverage) {
  var supData = getOrCreateSupervisorSheet().getDataRange().getValues();
  var mySite = "", myRole = "", myNick = "", myLdap = actor.split("@")[0];
  var nickByLdap = {};
  for (var i = 1; i < supData.length; i++) {
    var rNick  = (supData[i][0] || "").toString().trim();
    var rEmail = (supData[i][1] || "").toString().trim().toLowerCase();
    var rSite  = (supData[i][2] || "").toString().trim().toUpperCase();
    var rLdap  = ((supData[i][3] || "").toString().trim().toLowerCase()) || rEmail.split("@")[0];
    var rRole  = (supData[i][4] || "").toString().trim();
    if (rLdap) nickByLdap[rLdap] = rNick ? (rNick + " (" + rLdap + ")") : rLdap;
    if (rEmail === actor) { mySite = rSite; myRole = rRole; myNick = rNick || myLdap; myLdap = rLdap || myLdap; }
  }
  if (!mySite || !myRole) return null;

  var toPeers = [], ccOps = [];
  for (var j = 1; j < supData.length; j++) {
    var email = (supData[j][1] || "").toString().trim().toLowerCase();
    var site  = (supData[j][2] || "").toString().trim().toUpperCase();
    var role  = (supData[j][4] || "").toString().trim();
    if (!email) continue;
    if (/^operations\s*manager$/i.test(role)) { if (ccOps.indexOf(email) < 0) ccOps.push(email); }
    else if (site === mySite && email !== actor) { if (toPeers.indexOf(email) < 0) toPeers.push(email); }
  }

  var tag = kind === "cancel" ? "[BCP CANCELLED]" : (kind === "update" ? "[BCP UPDATE]" : "[BCP]");
  var link = ""; try { link = ScriptApp.getService().getUrl(); } catch (e) {}
  return {
    email: actor, ldap: myLdap, nick: myNick, site: mySite, role: myRole,
    subject: tag + " VL Coverage — " + myLdap + " (" + mySite + " · " + myRole + ") | " + _fmtMailRange(start, end),
    to: BCP_TEST_MODE ? [actor] : toPeers,
    cc: BCP_TEST_MODE ? [] : ccOps,
    cov: (coverage || []).map(function(l) {
      var c = l.toString().trim().toLowerCase();
      return nickByLdap[c] || c;
    }).join(", "),
    appUrl: link,
    testMode: BCP_TEST_MODE
  };
}

function getBcpPlan(data) {
  var actor = getSessionEmail().toLowerCase();
  if (!actor) return { success: false, message: "Could not identify your account." };
  if (!_checkRateLimit("bcpPlan_" + actor, 20, 60)) return { success: false, message: "Too many requests. Please wait a moment." };
  try {
    var start = (data.startDate || "").toString().trim();
    var end   = (data.endDate   || "").toString().trim() || start;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end)) return { success: false, message: "Invalid dates." };
    var kind = ["new", "update", "cancel"].indexOf(data.kind) >= 0 ? data.kind : "new";
    var cov = (Array.isArray(data.coverage) ? data.coverage : []).map(function(c) {
      return c.toString().trim().toLowerCase().replace(/[^a-z0-9._-]/g, "");
    }).filter(Boolean);
    var plan = _bcpPlan(actor, kind, start, end, cov);
    if (!plan) return { success: false, message: "Your email was not found in the Roles tab." };
    return { success: true, plan: plan };
  } catch (e) {
    _log("ERROR", "getBcpPlan", e.message);
    return { success: false, message: "Could not prepare the BCP email: " + e.message };
  }
}

function markBcpOpened(uuid, sentFor) {
  var actor = getSessionEmail().toLowerCase();
  uuid = (uuid || "").toString().trim();
  if (!actor || !uuid) return { success: false };
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    var sheet = getVLSheet();
    ensureColumn(sheet, "BCP Sent For");
    ensureColumn(sheet, "BCP Opened");
    var data = sheet.getDataRange().getValues();
    var norm = function(h) { return h.toString().replace(/[^a-zA-Z0-9]/g, "").toLowerCase(); };
    var H = data[0].map(norm);
    var cU = H.indexOf("rowuuid"), cE = H.indexOf("emailaddress"), cS = H.indexOf("bcpsentfor"), cO = H.indexOf("bcpopened");
    if (cU < 0 || cE < 0 || cS < 0 || cO < 0) return { success: false };
    for (var r = 1; r < data.length; r++) {
      if ((data[r][cU] || "").toString() === uuid && (data[r][cE] || "").toString().toLowerCase() === actor) {
        sheet.getRange(r + 1, cS + 1).setNumberFormat("@").setValue(_safeCell((sentFor || "").toString(), 40));
        sheet.getRange(r + 1, cO + 1).setValue(new Date());
        CacheService.getScriptCache().remove("vl_events_cache");
        return { success: true };
      }
    }
    return { success: false };
  } catch (e) {
    _log("WARN", "markBcpOpened", e.message);
    return { success: false };
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
