/**
 * ==================================================================
 * PROJECT EXODUS — Email Automation v2.0 (Corrected)
 * ==================================================================
 */

/**
 * ------------------------------------------------------------------
 * CONFIGURATION — Edit these values to customize behavior
 * ------------------------------------------------------------------
 */
const EMAIL_CONFIG = {
  SHEET_NAME:          "Support VL Requests",
  LOG_SHEET_NAME:      "Update Log 1.0",
  ADMIN_EMAIL:         "admin@google.com",           // ← Your email for summaries & error reports
  REPLY_TO_EMAIL:      "gup-play-ops@google.com",    // ← Reply-to address for agents
  SENDER_NAME:         "Project Pahinga System",
  DASHBOARD_LINK:      "https://docs.google.com/spreadsheets/d/1RTDHOg74c92d3jTivh5qA6XgpRPPsHPYsX6nTTl_aPs/edit?gid=727043204#gid=727043204",

  // Update Log: only watch these columns (1-indexed) on the Leave Results sheet
  LOG_WATCHED_SHEET:   "Support VL Requests",
  LOG_WATCHED_COLUMNS: [12, 13], // Status and Comments
};

/**
 * ------------------------------------------------------------------
 * COLUMN MAP — Update here if columns ever shift; never in the code
 * ------------------------------------------------------------------
 */
const EMAIL_COL = {
  TIMESTAMP:      0,   // Col A
  EMAIL:          1,   // Col B — agent email address
  LDAP:           2,   // Col C
  NICKNAME:       3,   // Col D
  ROLE:           4,   // Col E
  SITE:           5,   // Col F
  VL_DATE:        6,   // Col G
  END_DATE:       7,   // Col H
  VL_TYPE:        8,   // Col I
  REASON:         9,   // Col J
  SHIFT_START:    10,  // Col K
  STATUS:         11,  // Col L
  COMMENTS:       12,  // Col M
  CONFIRMATION:   13,  // Col N
  EMAIL_SENT:     14,  // Col O
  COVERAGE_SNAP:  15,  // Col P
  UUID:           16,  // Col Q
};

const EMAIL_TOTAL_COLS = 17;

/**
 * ------------------------------------------------------------------
 * MENU
 * ------------------------------------------------------------------
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Project Exodus')
    .addItem('📧 Send Status Emails',             'sendLeaveNotifications')
    .addItem('🔍 Dry Run (Preview Only)',          'dryRunNotifications')
    .addSeparator()
    .addItem('⏰ Enable Hourly Auto-Send',         'setupHourlyTrigger')
    .addItem('🛑 Disable Auto-Send',              'removeTriggers')
    .addSeparator()
    .addItem('🧪 Test Supervisor Summary Email',  'testSupervisorSummary')
    .addItem('🧪 Test Approval Reminder Email',   'testApprovalReminderEmail')
    .addToUi();
}

/**
 * ------------------------------------------------------------------
 * PUBLIC ENTRY POINTS
 * ------------------------------------------------------------------
 */
function sendLeaveNotifications() {
  _runNotifications(false);
}

function dryRunNotifications() {
  _runNotifications(true);
}

/**
 * ------------------------------------------------------------------
 * CORE ENGINE
 * ------------------------------------------------------------------
 */
function _runNotifications(isDryRun) {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const sheet       = ss.getSheetByName(EMAIL_CONFIG.SHEET_NAME);
  const ui          = SpreadsheetApp.getUi();
  const adminEmail  = Session.getActiveUser().getEmail() || EMAIL_CONFIG.ADMIN_EMAIL;

  if (!sheet) {
    ui.alert(`Error: Could not find the '${EMAIL_CONFIG.SHEET_NAME}' sheet.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("No data found to process.");
    return;
  }

  const data           = sheet.getRange(2, 1, lastRow - 1, EMAIL_TOTAL_COLS).getValues();
  let   emailsSent     = 0;
  let   emailsSkipped  = 0;
  const errors         = []; // { ldap, row, reason }
  const dryRunLog      = []; // { row, ldap, status, email }
  const supervisorMap  = {}; // { supervisorEmail: { smeEmail, entries: [...] } }

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    // --- DATA MAPPING via COL constants ---
    const timestamp    = row[EMAIL_COL.TIMESTAMP];
    const emailAddress = row[EMAIL_COL.EMAIL];
    const ldap         = row[EMAIL_COL.LDAP];
    const nickname     = row[EMAIL_COL.NICKNAME];
    const role         = row[EMAIL_COL.ROLE];
    const site         = row[EMAIL_COL.SITE];
    const vlDate       = row[EMAIL_COL.VL_DATE];
    const vlType       = row[EMAIL_COL.VL_TYPE];
    const reason       = row[EMAIL_COL.REASON];
    const status       = row[EMAIL_COL.STATUS];
    const comments     = row[EMAIL_COL.COMMENTS];
    const sentStatus   = row[EMAIL_COL.EMAIL_SENT];
    const confirmation = row[EMAIL_COL.CONFIRMATION];

    // --- SKIP CONDITIONS ---
    if (!status || status.toString().toLowerCase() === "pending") { emailsSkipped++; continue; }
    // Duplicate guard: checks string start so "Sent" and "Sent - 2026-..." both match
    if (sentStatus && sentStatus.toString().startsWith("Sent")) { emailsSkipped++; continue; }
    // Also skip cancelled
    if (status.toString().toLowerCase() === "cancelled") { emailsSkipped++; continue; }

    try {
      const formattedDate = Utilities.formatDate(new Date(vlDate), Session.getScriptTimeZone(), "MMMM dd, yyyy");
      const lowerStatus   = status.toString().toLowerCase();

      const subject   = `Pahinga Support VL Application - ${status}`;
      const emailBody = createEmailTemplate(
        nickname || ldap, formattedDate, status, vlType, site, comments, timestamp, confirmation
      );

      if (isDryRun) {
        dryRunLog.push({
          row:    i + 2,
          ldap:   ldap,
          status: status,
          email:  emailAddress,
        });
      } else {
        // --- SEND AGENT EMAIL ---
        MailApp.sendEmail({
          to:       emailAddress,
          subject:  subject,
          htmlBody: emailBody,
          name:     EMAIL_CONFIG.SENDER_NAME,
          replyTo:  EMAIL_CONFIG.REPLY_TO_EMAIL,
        });

        // Write "Sent"
        const sentAt      = new Date();
        const sentAtLabel = Utilities.formatDate(sentAt, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        sheet.getRange(i + 2, EMAIL_COL.EMAIL_SENT + 1).setValue(`Sent - ${sentAtLabel}`);

        emailsSent++;
      }

    } catch (e) {
      Logger.log(`Error processing ${ldap}: ${e.toString()}`);
      errors.push({ ldap: ldap, row: i + 2, reason: e.toString() });
      if (!isDryRun) {
        sheet.getRange(i + 2, COL.SENT_STATUS + 1).setValue("Error");
      }
    }
  }

  // --- POST-RUN ACTIONS ---
  if (isDryRun) {
    _showDryRunDialog(dryRunLog, errors);

  } else {
    // Send admin summary if anything happened
    if (emailsSent > 0 || errors.length > 0) {
      _sendAdminSummary(adminEmail, emailsSent, emailsSkipped, errors);
    }

    // Toast feedback
    if (emailsSent > 0) {
      const errNote = errors.length > 0 ? ` ⚠️ ${errors.length} error(s) — check your inbox.` : "";
      ss.toast(`✅ ${emailsSent} email(s) sent.${errNote}`, "Project Pahinga");
    } else if (errors.length > 0) {
      ss.toast(`⚠️ ${errors.length} error(s) occurred. Check your inbox.`, "Project Pahinga");
    } else {
      ss.toast("No new pending notifications found.", "Project Pahinga");
    }
  }
}

/**
 * ------------------------------------------------------------------
 * TEST — Sends a sample agent approval email to the active user's own email only.
 * ------------------------------------------------------------------
 */
function testApprovalReminderEmail() {
  const testEmail = Session.getActiveUser().getEmail();
  const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM dd, yyyy");

  const emailBody = createEmailTemplate(
    "jecylangela",         // nickname
    formattedDate,         // date
    "Approved",            // status — change to "Denied" to test other branches
    "Personal",            // vlType
    "CEB",                 // site
    "No additional comments", // comments
    new Date(),            // timestamp
    "TEST-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss") // confirmation
  );

  MailApp.sendEmail({
    to:       testEmail,
    subject:  "[TEST] Pahinga Support VL Application - Approved",
    htmlBody: emailBody,
    name:     EMAIL_CONFIG.SENDER_NAME,
    replyTo:  EMAIL_CONFIG.REPLY_TO_EMAIL,
  });

  SpreadsheetApp.getActiveSpreadsheet()
    .toast(`🧪 Test approval email sent to ${testEmail}`, "Project Pahinga");
}

/**
 * ------------------------------------------------------------------
 * DRY RUN DIALOG
 * ------------------------------------------------------------------
 */
function _showDryRunDialog(log, errors) {
  if (log.length === 0 && errors.length === 0) {
    SpreadsheetApp.getUi().alert("🔍 Dry Run Complete\n\nNo pending emails found to send.");
    return;
  }

  let msg = `🔍 DRY RUN — ${log.length} email(s) would be sent:\n\n`;
  log.forEach(entry => {
    msg += `Row ${entry.row}: ${entry.ldap} → ${entry.email}\n`;
    msg += `  Status : ${entry.status}\n\n`;
  });

  if (errors.length > 0) {
    msg += `\n⚠️ ${errors.length} row(s) would have errored:\n`;
    errors.forEach(err => {
      msg += `  Row ${err.row} (${err.ldap}): ${err.reason}\n`;
    });
  }

  msg += "\n✋ No emails were sent. Run 'Send Status Emails' to go live.";
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * ------------------------------------------------------------------
 * ADMIN SUMMARY EMAIL — sent to admin after each real run
 * ------------------------------------------------------------------
 */
function _sendAdminSummary(adminEmail, sent, skipped, errors) {
  const now = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), "MMMM dd, yyyy 'at' HH:mm:ss"
  );

  let errorSection = "";
  if (errors.length > 0) {
    const errorRows = errors.map(e => `
      <tr>
        <td style="padding:8px 12px;border-bottom:1px solid #f1f3f4;">${e.ldap}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #f1f3f4;">${e.row}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #f1f3f4;color:#d93025;font-size:11px;">${e.reason}</td>
      </tr>`).join("");

    errorSection = `
      <h3 style="color:#d93025;margin:25px 0 10px 0;">⚠️ Failed Rows (${errors.length})</h3>
      <table style="width:100%;border-collapse:collapse;font-size:13px;background:#fff8f8;border-radius:6px;overflow:hidden;border:1px solid #f5c6c6;">
        <tr style="background:#fce8e6;">
          <th style="padding:8px 12px;text-align:left;color:#c5221f;font-weight:600;">LDAP</th>
          <th style="padding:8px 12px;text-align:left;color:#c5221f;font-weight:600;">Row</th>
          <th style="padding:8px 12px;text-align:left;color:#c5221f;font-weight:600;">Error</th>
        </tr>
        ${errorRows}
      </table>`;
  }

  const summaryBody = `
    <div style="font-family:'Google Sans',Roboto,Helvetica,Arial,sans-serif;max-width:600px;border:1px solid #dadce0;border-radius:8px;overflow:hidden;margin:0 auto;background:#ffffff;">
      <table style="width:100%;border-collapse:collapse;height:6px;">
        <tr>
          <td style="background-color:#4285F4;width:25%;height:6px;"></td>
          <td style="background-color:#EA4335;width:25%;height:6px;"></td>
          <td style="background-color:#FBBC05;width:25%;height:6px;"></td>
          <td style="background-color:#34A853;width:25%;height:6px;"></td>
        </tr>
      </table>
      <div style="padding:30px;">
        <h2 style="color:#1a73e8;margin:0 0 4px 0;font-weight:400;">📊 Project Exodus — Run Summary</h2>
        <p style="color:#80868b;margin:0 0 25px 0;font-size:12px;">${now}</p>

        <table style="width:100%;border-collapse:collapse;font-size:14px;">
          <tr style="border-bottom:1px solid #f1f3f4;">
            <td style="padding:14px 0;color:#5f6368;">✅ Emails Sent</td>
            <td style="padding:14px 0;text-align:right;font-weight:bold;color:#188038;font-size:18px;">${sent}</td>
          </tr>
          <tr style="border-bottom:1px solid #f1f3f4;">
            <td style="padding:14px 0;color:#5f6368;">⏭️ Rows Skipped</td>
            <td style="padding:14px 0;text-align:right;color:#202124;font-size:18px;">${skipped}</td>
          </tr>
          <tr style="border-bottom:1px solid #f1f3f4;">
            <td style="padding:14px 0;color:#5f6368;">❌ Errors</td>
            <td style="padding:14px 0;text-align:right;font-weight:bold;color:${errors.length > 0 ? '#d93025' : '#202124'};font-size:18px;">${errors.length}</td>
          </tr>
        </table>

        ${errorSection}

        <div style="text-align:center;margin:30px 0;">
          <a href="${EMAIL_CONFIG.DASHBOARD_LINK}" style="background-color:#1a73e8;color:#ffffff;padding:12px 24px;text-decoration:none;border-radius:4px;font-size:14px;font-weight:500;display:inline-block;box-shadow:0 1px 2px rgba(60,64,67,0.3);">View Leave Dashboard</a>
        </div>

        <hr style="border:0;height:1px;background:#dadce0;margin:25px 0;">
        <p style="font-size:11px;color:#d93025;text-align:center;font-weight:500;">
          This is an automated summary from Project Pahinga (Support VL).
        </p>
      </div>
    </div>`;

  MailApp.sendEmail({
    to:       adminEmail,
    subject:  `📊 Project Pahinga Run Summary — ${sent} sent, ${errors.length} error(s)`,
    htmlBody: summaryBody,
    name:     "Project Pahinga Admin",
  });
}

/**
 * ------------------------------------------------------------------
 * TRIGGER MANAGEMENT — hourly auto-send
 * ------------------------------------------------------------------
 */
function setupHourlyTrigger() {
  removeTriggers(true);
  ScriptApp.newTrigger("sendLeaveNotifications")
    .timeBased()
    .everyHours(1)
    .create();
  SpreadsheetApp.getActiveSpreadsheet()
    .toast("⏰ Hourly auto-send enabled.", "Project Pahinga");
}

function removeTriggers(silent) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "sendLeaveNotifications")
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (!silent) {
    SpreadsheetApp.getActiveSpreadsheet()
      .toast("🛑 Auto-send disabled.", "Project Pahinga");
  }
}

/**
 * ------------------------------------------------------------------
 * HTML EMAIL TEMPLATE
 * ------------------------------------------------------------------
 */
function createEmailTemplate(nickname, date, status, vlType, site, comments, timestamp, confirmation) {
  var statusColor = "#5f6368";
  var statusBg = "#f1f3f4";
  var lowerStatus = status.toString().toLowerCase();

  if (lowerStatus.includes("approved")) { statusColor = "#137333"; statusBg = "#e6f4ea"; }
  else if (lowerStatus.includes("denied")) { statusColor = "#d93025"; statusBg = "#fce8e6"; }

  var dashboardLink = "";
  try {
    dashboardLink = ScriptApp.getService().getUrl();
  } catch (e) {
    // Fallback to CONFIG link if ScriptApp fails (e.g. not running as web app yet)
    dashboardLink = (typeof EMAIL_CONFIG !== 'undefined' && EMAIL_CONFIG.DASHBOARD_LINK) ? EMAIL_CONFIG.DASHBOARD_LINK : "#";
  }

  return `
    <div style="font-family: 'Google Sans', Roboto, Helvetica, Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #ffffff; border: 1px solid #e0e0e0; border-radius: 16px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.05);">
      <!-- Brand Header -->
      <div style="background-color: #f8f9fa; padding: 24px; border-bottom: 1px solid #f1f3f4; display: flex; align-items: center; gap: 12px;">
        <img src="https://upload.wikimedia.org/wikipedia/commons/a/a5/Google_Calendar_icon_%282020%29.svg" width="32" height="32" style="display:block">
        <span style="font-size: 20px; font-weight: 500; color: #3c4043; font-family: 'Google Sans Display', sans-serif;">Pahinga Support VL Calendar</span>
      </div>

      <!-- Status Banner -->
      <div style="padding: 40px 32px 32px; text-align: center;">
        <div style="display: inline-block; padding: 8px 16px; background-color: ${statusBg}; border-radius: 100px; margin-bottom: 16px;">
          <span style="font-size: 14px; font-weight: 700; color: ${statusColor}; text-transform: uppercase; letter-spacing: 0.8px;">${status}</span>
        </div>
        <h1 style="font-size: 36px; font-weight: 400; color: #202124; margin: 0; letter-spacing: -0.5px;">Update for ${nickname}</h1>
        <p style="font-size: 16px; color: #5f6368; margin-top: 12px;">Your leave request for <strong>${date}</strong> has been reviewed by the system.</p>
      </div>

      <!-- Details Card -->
      <div style="padding: 0 32px 32px;">
        <div style="background-color: #f8f9fa; border-radius: 12px; padding: 24px; border: 1px solid #f1f3f4;">
          <table style="width: 100%; border-collapse: collapse;">
            <tr>
              <td style="padding: 8px 0; font-size: 12px; font-weight: 700; color: #70757a; text-transform: uppercase; letter-spacing: 0.5px; width: 40%;">VL Date</td>
              <td style="padding: 8px 0; font-size: 15px; font-weight: 500; color: #202124; text-align: right;">${date}</td>
            </tr>
            <tr>
              <td style="padding: 8px 0; font-size: 12px; font-weight: 700; color: #70757a; text-transform: uppercase; letter-spacing: 0.5px;">VL Type</td>
              <td style="padding: 8px 0; font-size: 15px; font-weight: 500; color: #202124; text-align: right;">${vlType}</td>
            </tr>
            <tr>
              <td style="padding: 8px 0; font-size: 12px; font-weight: 700; color: #70757a; text-transform: uppercase; letter-spacing: 0.5px;">Site</td>
              <td style="padding: 8px 0; font-size: 15px; font-weight: 500; color: #202124; text-align: right;">${site || "N/A"}</td>
            </tr>
            <tr><td colspan="2" style="padding: 16px 0;"><div style="height: 1px; background-color: #dadce0;"></div></td></tr>
            <tr>
              <td style="padding: 8px 0; font-size: 12px; font-weight: 700; color: #70757a; text-transform: uppercase; letter-spacing: 0.5px;">Manager Note</td>
              <td style="padding: 8px 0; font-size: 14px; font-style: italic; color: #3c4043; text-align: right;">"${comments || "No additional comments"}"</td>
            </tr>
          </table>
        </div>

        <!-- Action Button -->
        <div style="text-align: center; margin-top: 32px;">
          <a href="${dashboardLink}" style="background-color: #1a73e8; color: #ffffff; padding: 14px 32px; text-decoration: none; border-radius: 24px; font-size: 14px; font-weight: 500; display: inline-block; box-shadow: 0 2px 4px rgba(26,115,232,0.25);">Open Calendar Dashboard</a>
        </div>
      </div>

      <!-- Footer Info -->
      <div style="padding: 24px 32px; background-color: #f8f9fa; border-top: 1px solid #f1f3f4; text-align: center;">
        <p style="font-size: 11px; color: #70757a; margin: 0 0 8px;">Submitted on ${timestamp}</p>
        <p style="font-size: 11px; color: #bdc1c6; margin: 0;">This is an automated notification from Project Pahinga. Ref: ${confirmation || "N/A"}</p>
      </div>
    </div>
  `;
}

/**
 * ------------------------------------------------------------------
 * PART 2: UPDATE LOGGER (Filtered — only Leave Results, key columns)
 * ------------------------------------------------------------------
 */
function onEdit(e) {
  if (!e) return;

  const ss      = e.source;
  const sheet   = ss.getActiveSheet();
  const range   = e.range;
  const tabName = sheet.getName();

  // Only log edits on the watched sheet AND watched columns
  if (tabName !== EMAIL_CONFIG.LOG_WATCHED_SHEET)                     return;
  if (!EMAIL_CONFIG.LOG_WATCHED_COLUMNS.includes(range.getColumn())) return;

  const editedValue = range.getValue();
  const timestamp   = new Date();

  let userEmail = "Anonymous Editor";
  if (e.user && e.user.getEmail) {
    userEmail = e.user.getEmail();
  } else if (e.user) {
    userEmail = e.user;
  }

  let logSheet = ss.getSheetByName(EMAIL_CONFIG.LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(EMAIL_CONFIG.LOG_SHEET_NAME);
    logSheet.getRange("B3:E3").setValues([["Email Address", "Name of Tab Edited", "Edit/Change Made", "Timestamp"]]);
    logSheet.setFrozenRows(3);

    // Protection is set ONCE here at creation time only.
    try {
      const protection = logSheet.protect().setDescription("Do not edit!");
      const me = Session.getEffectiveUser();
      protection.addEditor(me);
      protection.removeEditors(protection.getEditors());
      if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
      }
    } catch (error) {
      // Ignore permissions error
    }
  }

  const nextRow = logSheet.getLastRow() + 1;
  logSheet.getRange(nextRow, 2, 1, 4).setValues([
    [userEmail, tabName, `Cell ${range.getA1Notation()} changed to "${editedValue}"`, timestamp]
  ]);
}
