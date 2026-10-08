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
const CONFIG = {
  SHEET_NAME:          "Leave Results",
  LOG_SHEET_NAME:      "Update Log 1.0",
  ADMIN_EMAIL:         "admin@google.com",           // ← Your email for summaries & error reports
  REPLY_TO_EMAIL:      "gup-play-ops@google.com",    // ← Reply-to address for agents
  SENDER_NAME:         "gUp Play Support System",
  DASHBOARD_LINK:      "https://docs.google.com/spreadsheets/d/1RTDHOg74c92d3jTivh5qA6XgpRPPsHPYsX6nTTl_aPs/edit?gid=727043204#gid=727043204",
  WORKFORCECX_LINK:    "https://workforcecx.concentrix.com/login",

  // Update Log: only watch these columns (1-indexed) on the Leave Results sheet
  // N=Status(14), O=Queue(15), P=Comments(16), Q=Attendance(17)
  LOG_WATCHED_SHEET:   "Leave Results",
  LOG_WATCHED_COLUMNS: [14, 15, 16, 17],
};

/**
 * ------------------------------------------------------------------
 * COLUMN MAP — Update here if columns ever shift; never in the code
 * ------------------------------------------------------------------
 */
const COL = {
  TIMESTAMP:    0,   // Col A
  EMAIL:        1,   // Col B — agent email address
  LDAP:         2,   // Col C
  VL_DATE:      3,   // Col D
  TEAM:         4,   // Col E
  REASON:       5,   // Col F
  ACCRUALS:     8,   // Col I
  SITE:         12,  // Col M
  STATUS:       13,  // Col N
  QUEUE:        14,  // Col O
  COMMENTS:     15,  // Col P
  ATTENDANCE:   16,  // Col Q
  CHANNEL:      17,  // Col R
  POOL:         18,  // Col S
  SENT_STATUS:  21,  // Col V
  CONFIRMATION: 22,  // Col W
  SENT_AT:      23,  // Col X — send timestamp + admin
  SUPERVISOR:   26,  // Col AA — Team Captain (supervisor) email
  SME:          27,  // Col AB — SME email (CC)
};

const TOTAL_COLS = 28; // A through AB

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
  const sheet       = ss.getSheetByName(CONFIG.SHEET_NAME);
  const ui          = SpreadsheetApp.getUi();
  const adminEmail  = Session.getActiveUser().getEmail() || CONFIG.ADMIN_EMAIL;

  if (!sheet) {
    ui.alert(`Error: Could not find the '${CONFIG.SHEET_NAME}' sheet.`);
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert("No data found to process.");
    return;
  }

  const data           = sheet.getRange(2, 1, lastRow - 1, TOTAL_COLS).getValues();
  let   emailsSent     = 0;
  let   emailsSkipped  = 0;
  const errors         = []; // { ldap, row, reason }
  const dryRunLog      = []; // { row, ldap, status, email }
  const supervisorMap  = {}; // { supervisorEmail: { smeEmail, entries: [...] } }

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    // --- DATA MAPPING via COL constants ---
    const timestamp    = row[COL.TIMESTAMP];
    const ldap         = row[COL.LDAP];
    const vlDate       = row[COL.VL_DATE];
    const team         = row[COL.TEAM];
    const reason       = row[COL.REASON];
    const accruals     = row[COL.ACCRUALS];
    const site         = row[COL.SITE];
    const status       = row[COL.STATUS];
    const queue        = row[COL.QUEUE];
    const comments     = row[COL.COMMENTS];
    let   attendance   = row[COL.ATTENDANCE];
    const channel      = row[COL.CHANNEL];
    const pool         = row[COL.POOL];
    const sentStatus   = row[COL.SENT_STATUS];
    const confirmation = row[COL.CONFIRMATION];
    const supervisor   = row[COL.SUPERVISOR];
    const sme          = row[COL.SME];

    // --- WORKGROUP COMBINATION ---
    let workgroupCombined = "-";
    if (channel && pool)   workgroupCombined = `${channel} - ${pool}`;
    else                   workgroupCombined = channel || pool || "-";

    // --- FORCE PERCENTAGE FORMAT ---
    if (typeof attendance === 'number') {
      attendance = (attendance * 100).toFixed(2) + "%";
    }

    // --- SKIP CONDITIONS ---
    if (!status || !queue)                        { emailsSkipped++; continue; }
    // Duplicate guard: checks string start so "Sent" and "Sent - 2026-..." both match
    if (sentStatus.toString().startsWith("Sent")) { emailsSkipped++; continue; }

    try {
      const emailAddress  = row[COL.EMAIL];
      const formattedDate = Utilities.formatDate(new Date(vlDate), Session.getScriptTimeZone(), "MMMM dd, yyyy");
      const lowerStatus   = status.toString().toLowerCase();

      const subject   = `Play Leave Application - ${status}`;
      const emailBody = createEmailTemplate(
        ldap, formattedDate, status, queue, team, reason,
        workgroupCombined, comments, timestamp, "",
        accruals, attendance, site, confirmation
      );

      // --- FIX #5: COLLECT FOR SUPERVISOR SUMMARY BEFORE SENDING EMAIL ---
      // Moved here so the supervisor still gets their summary even if the
      // agent email fails. The !isDryRun guard prevents ghost entries in
      // dry run mode.
      if (supervisor && !isDryRun) {
        const supRaw = supervisor.toString().trim();
        const supKey = supRaw.includes("@") ? supRaw : supRaw + "@google.com";
        const smeRaw = sme ? sme.toString().trim() : "";
        const smeKey = smeRaw && !smeRaw.includes("@") ? smeRaw + "@google.com" : smeRaw;
        if (!supervisorMap[supKey]) {
          supervisorMap[supKey] = { smeEmail: smeKey, entries: [] };
        }
        supervisorMap[supKey].entries.push({
          ldap:         ldap,
          date:         formattedDate,
          status:       status,
          emoji:        "",
          queue:        queue,
          site:         site,
          workgroup:    workgroupCombined,
          attendance:   attendance,
          accruals:     accruals,
          comments:     comments,
          confirmation: confirmation,
          reason:       reason,
          team:         team,
          timestamp:    timestamp,
        });
      }

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
          name:     CONFIG.SENDER_NAME,
          replyTo:  CONFIG.REPLY_TO_EMAIL,
        });

        // Write "Sent" to Col V + timestamp and admin to Col X
        const sentAt      = new Date();
        const sentAtLabel = Utilities.formatDate(sentAt, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        sheet.getRange(i + 2, COL.SENT_STATUS + 1).setValue("Sent");
        sheet.getRange(i + 2, COL.SENT_AT + 1).setValue(`${sentAtLabel} by System-Automated`);

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
    // Send supervisor summaries
    if (Object.keys(supervisorMap).length > 0) {
      _sendSupervisorSummaries(supervisorMap);
    }

    // Send admin summary if anything happened
    if (emailsSent > 0 || errors.length > 0) {
      _sendAdminSummary(adminEmail, emailsSent, emailsSkipped, errors);
    }

    // Toast feedback
    if (emailsSent > 0) {
      const errNote = errors.length > 0 ? ` ⚠️ ${errors.length} error(s) — check your inbox.` : "";
      ss.toast(`✅ ${emailsSent} email(s) sent.${errNote}`, "Project Exodus");
    } else if (errors.length > 0) {
      ss.toast(`⚠️ ${errors.length} error(s) occurred. Check your inbox.`, "Project Exodus");
    } else {
      ss.toast("No new pending notifications found.", "Project Exodus");
    }
  }
}

/**
 * ------------------------------------------------------------------
 * SUPERVISOR SUMMARY EMAIL
 * Sent to Team Captain (Col AA), CC'd to SME (Col AB).
 * ------------------------------------------------------------------
 */
function _sendSupervisorSummaries(supervisorMap) {
  const now = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), "MMMM dd, yyyy 'at' HH:mm:ss"
  );

  for (const supervisorEmail in supervisorMap) {
    const { smeEmail, entries } = supervisorMap[supervisorEmail];

    const agentRows = entries.map(e => {
      let statusColor = "#5f6368";
      const ls = e.status.toString().toLowerCase();
      if      (ls.includes("birthday"))      statusColor = "#1a73e8";
      else if (ls.includes("approved"))       statusColor = "#188038";
      else if (ls.includes("denied"))         statusColor = "#d93025";
      else if (ls.includes("no allocation"))  statusColor = "#ea4335";
      else if (ls.includes("pending"))        statusColor = "#fbbc04";
      else if (ls.includes("no accruals"))    statusColor = "#e37400";
      else if (ls.includes("emergency"))      statusColor = "#a142f4";
      else if (ls.includes("duplicate"))      statusColor = "#e37400";
      else if (ls.includes("wrong date"))     statusColor = "#c5221f";
      else if (ls.includes("not in roster"))  statusColor = "#c5221f";
      return `
      <tr style="border-bottom: 1px solid #f1f3f4;">
        <td style="padding: 12px 10px; color: #202124; font-weight: 500;">${e.ldap}</td>
        <td style="padding: 12px 10px; color: #202124;">${e.date}</td>
        <td style="padding: 12px 10px; font-weight: bold; color: ${statusColor};">${e.status}</td>
        <td style="padding: 12px 10px; color: #202124;">${e.queue}</td>
        <td style="padding: 12px 10px; color: #202124;">${e.site || "-"}</td>
        <td style="padding: 12px 10px; color: #1a73e8; font-weight: 600;">${e.workgroup}</td>
        <td style="padding: 12px 10px; color: #188038; font-weight: 600;">${e.attendance || "-"}</td>
        <td style="padding: 12px 10px; color: #202124;">${e.accruals || "-"}</td>
        <td style="padding: 12px 10px; color: #202124;">${e.comments || "-"}</td>
        <td style="padding: 12px 10px; color: #202124;">${e.confirmation || "-"}</td>
      </tr>`;
    }).join("");

    const summaryBody = `
      <div style="font-family: 'Google Sans', Roboto, Helvetica, Arial, sans-serif; max-width: 900px; border: 1px solid #dadce0; border-radius: 8px; overflow: hidden; margin: 0 auto; background-color: #ffffff;">

        <table style="width: 100%; border-collapse: collapse; height: 6px;">
          <tr>
            <td style="background-color: #4285F4; width: 25%; height: 6px;"></td>
            <td style="background-color: #EA4335; width: 25%; height: 6px;"></td>
            <td style="background-color: #FBBC05; width: 25%; height: 6px;"></td>
            <td style="background-color: #34A853; width: 25%; height: 6px;"></td>
          </tr>
        </table>

        <div style="padding: 30px 30px 10px 30px;">
          <h2 style="color: #1a73e8; margin: 0 0 4px 0; font-weight: 400;">📋 Leave Application Summary — Your Team</h2>
          <p style="color: #80868b; margin: 0 0 20px 0; font-size: 12px;">${now}</p>
          <p style="color: #202124; font-size: 14px; margin-bottom: 6px;">Hi,</p>
          <p style="color: #5f6368; font-size: 14px; margin-bottom: 20px;">
            The following leave notifications were just sent to your agents.
            This is a summary for your records.
          </p>
        </div>

        <div style="padding: 0 30px 30px 30px; overflow-x: auto;">
          <table style="width: 100%; border-collapse: collapse; font-size: 13px; min-width: 750px;">
            <thead>
              <tr style="background-color: #f8f9fa; border-bottom: 2px solid #dadce0;">
                <th style="padding: 12px 10px; text-align: left; color: #5f6368; font-weight: 600;">Agent (LDAP)</th>
                <th style="padding: 12px 10px; text-align: left; color: #5f6368; font-weight: 600;">VL Date</th>
                <th style="padding: 12px 10px; text-align: left; color: #5f6368; font-weight: 600;">Status</th>
                <th style="padding: 12px 10px; text-align: left; color: #5f6368; font-weight: 600;">Queue / Slot</th>
                <th style="padding: 12px 10px; text-align: left; color: #5f6368; font-weight: 600;">Site</th>
                <th style="padding: 12px 10px; text-align: left; color: #5f6368; font-weight: 600;">Workgroup</th>
                <th style="padding: 12px 10px; text-align: left; color: #5f6368; font-weight: 600;">Attendance</th>
                <th style="padding: 12px 10px; text-align: left; color: #5f6368; font-weight: 600;">Accruals</th>
                <th style="padding: 12px 10px; text-align: left; color: #5f6368; font-weight: 600;">Comments</th>
                <th style="padding: 12px 10px; text-align: left; color: #5f6368; font-weight: 600;">Confirmation</th>
              </tr>
            </thead>
            <tbody>
              ${agentRows}
            </tbody>
          </table>
        </div>

        <div style="padding: 0 30px 30px 30px; text-align: center;">
          <a href="${CONFIG.DASHBOARD_LINK}" style="background-color: #1a73e8; color: #ffffff; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-size: 14px; font-weight: 500; display: inline-block; box-shadow: 0 1px 2px rgba(60,64,67,0.3);">View Leave Dashboard</a>
        </div>

        <div style="padding: 0 30px 30px 30px;">
          <hr style="border: 0; height: 1px; background: #dadce0; margin: 0 0 20px 0;">
          <p style="font-size: 11px; color: #d93025; text-align: center; font-weight: 500;">
            This is an automated message from Project Exodus (gUP Play Ops).
          </p>
        </div>

      </div>`;

    const mailOptions = {
      to:       supervisorEmail,
      subject:  `📋 Leave Summary — ${entries.length} agent notification(s) sent`,
      htmlBody: summaryBody,
      name:     CONFIG.SENDER_NAME,
      replyTo:  CONFIG.REPLY_TO_EMAIL,
    };

    // CC the SME (Col AB) if present
    if (smeEmail) {
      mailOptions.cc = smeEmail;
    }

    try {
      MailApp.sendEmail(mailOptions);
    } catch (e) {
      Logger.log(`Error sending supervisor summary to ${supervisorEmail}: ${e.toString()}`);
    }
  }
}

/**
 * ------------------------------------------------------------------
 * TEST — Sends a sample agent approval email (with WorkforceCX
 * reminder block) to the active user's own email only.
 * ------------------------------------------------------------------
 */
function testApprovalReminderEmail() {
  const testEmail = Session.getActiveUser().getEmail();
  const formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM dd, yyyy");

  const emailBody = createEmailTemplate(
    "jecylangela",        // ldap
    formattedDate,         // date
    "Approved",            // status — change to "Birthday Leave" or "Denied" to test other branches
    "Slot 1",              // queue
    "Team Steven",         // team
    "Personal",             // reason
    "Play NA",             // workgroup
    "No additional comments", // comments
    new Date(),             // timestamp
    "",                      // emoji (unused param, kept for signature compatibility)
    20,                       // accruals
    "100.00%",                // attendance
    "CEB",                     // site
    "TEST-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss") // confirmation
  );

  MailApp.sendEmail({
    to:       testEmail,
    subject:  "[TEST] Play Leave Application - Approved",
    htmlBody: emailBody,
    name:     CONFIG.SENDER_NAME,
    replyTo:  CONFIG.REPLY_TO_EMAIL,
  });

  SpreadsheetApp.getActiveSpreadsheet()
    .toast(`🧪 Test approval email sent to ${testEmail}`, "Project Exodus");
}

/**
 * ------------------------------------------------------------------
 * TEST — Sends a sample supervisor summary to the active user's email
 * ------------------------------------------------------------------
 */
function testSupervisorSummary() {
  const testEmail = Session.getActiveUser().getEmail();

  const mockSupervisorMap = {
    [testEmail]: {
      smeEmail: testEmail,
      entries: [
        {
          ldap:         "jdelacruz",
          date:         "April 05, 2026",
          status:       "Approved",
          emoji:        "✅",
          queue:        "Slot 3",
          site:         "BGC",
          workgroup:    "Play Support - Tier 1",
          attendance:   "97.50%",
          accruals:     "3.5",
          comments:     "Approved — within allocation.",
          confirmation: "CNF-20260401-001",
          reason:       "Personal",
          team:         "Alpha Team",
          timestamp:    new Date(),
        },
        {
          ldap:         "mreyes",
          date:         "April 06, 2026",
          status:       "Approved",
          emoji:        "✅",
          queue:        "Slot 1",
          site:         "Ortigas",
          workgroup:    "Play Support - Tier 2",
          attendance:   "93.00%",
          accruals:     "2.0",
          comments:     "Approved.",
          confirmation: "CNF-20260401-002",
          reason:       "Vacation",
          team:         "Beta Team",
          timestamp:    new Date(),
        },
      ],
    },
  };

  _sendSupervisorSummaries(mockSupervisorMap);

  SpreadsheetApp.getActiveSpreadsheet()
    .toast(`🧪 Test supervisor summary sent to ${testEmail}`, "Project Exodus");
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
          <a href="${CONFIG.DASHBOARD_LINK}" style="background-color:#1a73e8;color:#ffffff;padding:12px 24px;text-decoration:none;border-radius:4px;font-size:14px;font-weight:500;display:inline-block;box-shadow:0 1px 2px rgba(60,64,67,0.3);">View Leave Dashboard</a>
        </div>

        <hr style="border:0;height:1px;background:#dadce0;margin:25px 0;">
        <p style="font-size:11px;color:#d93025;text-align:center;font-weight:500;">
          This is an automated summary from Project Exodus (gUP Play Ops).
        </p>
      </div>
    </div>`;

  MailApp.sendEmail({
    to:       adminEmail,
    subject:  `📊 Project Exodus Run Summary — ${sent} sent, ${errors.length} error(s)`,
    htmlBody: summaryBody,
    name:     "Project Exodus Admin",
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
    .toast("⏰ Hourly auto-send enabled.", "Project Exodus");
}

function removeTriggers(silent) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "sendLeaveNotifications")
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (!silent) {
    SpreadsheetApp.getActiveSpreadsheet()
      .toast("🛑 Auto-send disabled.", "Project Exodus");
  }
}

/**
 * ------------------------------------------------------------------
 * HTML EMAIL TEMPLATE
 * ------------------------------------------------------------------
 */
function createEmailTemplate(ldap, date, status, queue, team, reason, workgroup, comments, timestamp, emoji, accruals, attendance, site, confirmation) {
  var statusColor = "#5f6368";
  var statusBg = "#f1f3f4";
  var lowerStatus = status.toString().toLowerCase();

  if (lowerStatus.includes("birthday")) { statusColor = "#673ab7"; statusBg = "#f3e5f5"; }
  else if (lowerStatus.includes("approved")) { statusColor = "#137333"; statusBg = "#e6f4ea"; }
  else if (lowerStatus.includes("denied")) { statusColor = "#d93025"; statusBg = "#fce8e6"; }
  else if (lowerStatus.includes("no alloc")) { statusColor = "#b06000"; statusBg = "#fef7e0"; }

  var dashboardLink = "";
  try {
    dashboardLink = ScriptApp.getService().getUrl();
  } catch (e) {
    // Fallback to CONFIG link if ScriptApp fails (e.g. not running as web app yet)
    dashboardLink = (typeof CONFIG !== 'undefined' && CONFIG.DASHBOARD_LINK) ? CONFIG.DASHBOARD_LINK : "#";
  }

  var workforceReminder = "";
  if (lowerStatus.includes("approved") || lowerStatus.includes("birthday")) {
    var wfLink = (typeof CONFIG !== 'undefined' && CONFIG.WORKFORCECX_LINK) ? CONFIG.WORKFORCECX_LINK : "#";
    workforceReminder = `
        <div style="margin-top: 24px; background-color: #fff8e1; border: 1px solid #fde293; border-left: 4px solid #f9ab00; border-radius: 12px; padding: 20px 24px;">
          <table style="width:100%;border-collapse:collapse;margin-bottom:14px;">
            <tr>
              <td style="vertical-align:middle;width:32px;"><span style="font-size:22px;">⏰</span></td>
              <td style="vertical-align:middle;"><span style="font-size:15px;font-weight:700;color:#7a4f00;">Action Required: File This in WorkforceCX</span></td>
            </tr>
          </table>
          <p style="font-size:13px;color:#7a4f00;margin:0 0 16px;line-height:1.5;">
            This approval is logged on the Play VL Calendar only. To make it official, you still need to submit a matching PTO request in <strong>WorkforceCX</strong>.
          </p>
          <p style="font-size:11px;font-weight:700;color:#7a4f00;text-transform:uppercase;letter-spacing:0.5px;margin:0 0 8px;">Step 1 — Log In</p>
          <table style="width:100%;border-collapse:collapse;margin-bottom:16px;">
            <tr><td style="padding:4px 0;font-size:12.5px;color:#3c2e00;"><span style="display:inline-block;width:20px;height:20px;line-height:20px;text-align:center;background-color:#1a73e8;color:#ffffff;border-radius:50%;font-size:11px;font-weight:700;margin-right:8px;">1</span>Access the <strong>WorkforceCX</strong> login portal</td></tr>
            <tr><td style="padding:4px 0;font-size:12.5px;color:#3c2e00;"><span style="display:inline-block;width:20px;height:20px;line-height:20px;text-align:center;background-color:#1a73e8;color:#ffffff;border-radius:50%;font-size:11px;font-weight:700;margin-right:8px;">2</span>Enter your <strong>SSO Credentials</strong></td></tr>
            <tr><td style="padding:4px 0;font-size:12.5px;color:#3c2e00;"><span style="display:inline-block;width:20px;height:20px;line-height:20px;text-align:center;background-color:#1a73e8;color:#ffffff;border-radius:50%;font-size:11px;font-weight:700;margin-right:8px;">3</span>Click <strong>SIGN IN</strong> to begin</td></tr>
          </table>
          <p style="font-size:11px;font-weight:700;color:#7a4f00;text-transform:uppercase;letter-spacing:0.5px;margin:0 0 8px;">Step 2 — Submit a PTO Request</p>
          <table style="width:100%;border-collapse:collapse;margin-bottom:18px;">
            <tr><td style="padding:4px 0;font-size:12.5px;color:#3c2e00;"><span style="display:inline-block;width:20px;height:20px;line-height:20px;text-align:center;background-color:#f9ab00;color:#ffffff;border-radius:50%;font-size:11px;font-weight:700;margin-right:8px;">1</span>Pick the highlighted day(s) on the calendar — <strong>${date}</strong></td></tr>
            <tr><td style="padding:4px 0;font-size:12.5px;color:#3c2e00;"><span style="display:inline-block;width:20px;height:20px;line-height:20px;text-align:center;background-color:#f9ab00;color:#ffffff;border-radius:50%;font-size:11px;font-weight:700;margin-right:8px;">2</span>Click <strong>Request</strong> to open the time-off pop-up</td></tr>
            <tr><td style="padding:4px 0;font-size:12.5px;color:#3c2e00;"><span style="display:inline-block;width:20px;height:20px;line-height:20px;text-align:center;background-color:#f9ab00;color:#ffffff;border-radius:50%;font-size:11px;font-weight:700;margin-right:8px;">3</span>Choose <strong>PTO</strong> from the Time Off Type list</td></tr>
            <tr><td style="padding:4px 0;font-size:12.5px;color:#3c2e00;"><span style="display:inline-block;width:20px;height:20px;line-height:20px;text-align:center;background-color:#f9ab00;color:#ffffff;border-radius:50%;font-size:11px;font-weight:700;margin-right:8px;">4</span><strong>Submit</strong> to finalize the request for approval</td></tr>
          </table>
          <div style="text-align:center;">
            <a href="${wfLink}" style="background-color:#f9ab00;color:#3c2e00;padding:12px 28px;text-decoration:none;border-radius:24px;font-size:13px;font-weight:700;display:inline-block;box-shadow:0 2px 4px rgba(249,171,0,0.3);">Open WorkforceCX Portal</a>
          </div>
        </div>`;
  }

  return `
    <div style="font-family: 'Google Sans', Roboto, Helvetica, Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #ffffff; border: 1px solid #e0e0e0; border-radius: 16px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.05);">
      <!-- Brand Header -->
      <div style="background-color: #f8f9fa; padding: 24px; border-bottom: 1px solid #f1f3f4; display: flex; align-items: center; gap: 12px;">
        <img src="https://upload.wikimedia.org/wikipedia/commons/a/a5/Google_Calendar_icon_%282020%29.svg" width="32" height="32" style="display:block">
        <span style="font-size: 20px; font-weight: 500; color: #3c4043; font-family: 'Google Sans Display', sans-serif;">Play VL Calendar</span>
      </div>

      <!-- Status Banner -->
      <div style="padding: 40px 32px 32px; text-align: center;">
        <div style="display: inline-block; padding: 8px 16px; background-color: ${statusBg}; border-radius: 100px; margin-bottom: 16px;">
          <span style="font-size: 14px; font-weight: 700; color: ${statusColor}; text-transform: uppercase; letter-spacing: 0.8px;">${status}</span>
        </div>
        <h1 style="font-size: 36px; font-weight: 400; color: #202124; margin: 0; letter-spacing: -0.5px;">Update for ${ldap}</h1>
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
              <td style="padding: 8px 0; font-size: 12px; font-weight: 700; color: #70757a; text-transform: uppercase; letter-spacing: 0.5px;">Site / Channel</td>
              <td style="padding: 8px 0; font-size: 15px; font-weight: 500; color: #202124; text-align: right;">${site || "N/A"} • ${workgroup}</td>
            </tr>
            <tr>
              <td style="padding: 8px 0; font-size: 12px; font-weight: 700; color: #70757a; text-transform: uppercase; letter-spacing: 0.5px;">Accruals</td>
              <td style="padding: 8px 0; font-size: 15px; font-weight: 500; color: #202124; text-align: right;">${accruals || "0"} credits</td>
            </tr>
            <tr>
              <td style="padding: 8px 0; font-size: 12px; font-weight: 700; color: #70757a; text-transform: uppercase; letter-spacing: 0.5px;">Attendance</td>
              <td style="padding: 8px 0; font-size: 15px; font-weight: 500; color: #1e8e3e; text-align: right;">${attendance || "N/A"}</td>
            </tr>
            <tr><td colspan="2" style="padding: 16px 0;"><div style="height: 1px; background-color: #dadce0;"></div></td></tr>
            <tr>
              <td style="padding: 8px 0; font-size: 12px; font-weight: 700; color: #70757a; text-transform: uppercase; letter-spacing: 0.5px;">Supervisor Note</td>
              <td style="padding: 8px 0; font-size: 14px; font-style: italic; color: #3c4043; text-align: right;">"${comments || "No additional comments"}"</td>
            </tr>
          </table>
        </div>

        <!-- Action Button -->
        <div style="text-align: center; margin-top: 32px;">
          <a href="${dashboardLink}" style="background-color: #1a73e8; color: #ffffff; padding: 14px 32px; text-decoration: none; border-radius: 24px; font-size: 14px; font-weight: 500; display: inline-block; box-shadow: 0 2px 4px rgba(26,115,232,0.25);">Open Calendar Dashboard</a>
        </div>
        ${workforceReminder}
      </div>

      <!-- Footer Info -->
      <div style="padding: 24px 32px; background-color: #f8f9fa; border-top: 1px solid #f1f3f4; text-align: center;">
        <p style="font-size: 11px; color: #70757a; margin: 0 0 8px;">Submitted on ${timestamp} • Reason: ${reason}</p>
        <p style="font-size: 11px; color: #bdc1c6; margin: 0;">This is an automated notification from Project Exodus (gUP Play Ops). Ref: ${confirmation || "N/A"}</p>
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
  if (tabName !== CONFIG.LOG_WATCHED_SHEET)                     return;
  if (!CONFIG.LOG_WATCHED_COLUMNS.includes(range.getColumn())) return;

  const editedValue = range.getValue();
  const timestamp   = new Date();

  let userEmail = "Anonymous Editor";
  if (e.user && e.user.getEmail) {
    userEmail = e.user.getEmail();
  } else if (e.user) {
    userEmail = e.user;
  }

  let logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
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
