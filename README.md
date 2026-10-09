# Project Pahinga - Support VL Platform

This repository contains the Google Apps Script project for the **Pahinga** Vacation Leave (VL) Calendar. This application replaces the legacy Play VL Calendar and has been repurposed for **Support staff (SMEs and TCs)**.

## Prerequisites & Data Model

This project is built to run on the V8 runtime and must be deployed as a web application. It relies on a bound Google Sheet with specific tabs for configuration and data storage, and one external read-only schedule file.

### External Schedule File (Read-Only)
- **ID:** `1J6X6Bxqh3vXFHP4iz5eqbrOpvYHHJC2w_4iBJI2haco`
- **Tab Name:** `NEW Schedule File`

### Bound Google Sheet Layout

Create the following tabs in the Google Sheet this script is bound to:

#### 1. Support VL Requests
Stores all VL applications.
**Columns:**
- Timestamp
- Email Address
- LDAP
- Nickname
- Role
- Site
- VL Date
- End Date
- VL Type
- Reason / Notes
- Shift Start
- Status
- Comments
- Confirmation on Status
- Email Sent
- Coverage Snapshot
- Row UUID

#### 2. Roles
Replaces the legacy Supervisors sheet.
**Columns:**
- Email
- Name
- Role (`SME` | `TC` | `Manager` | `Admin`)
- Site
- Active (TRUE/FALSE)

*Note:* Ensure to seed the Admin role (`stevenjosephc@google.com`) during initial setup.

#### 3. FormConfig
Configures application settings.
**Columns (must include):**
- IsOpen (TRUE/FALSE)
- ActiveMonths (comma-separated list, e.g., "May 2026")
- ClosedMsg (String)
- MinCoverage (Number, default `2`)
- ShiftHours (Number, default `8`)
- IntervalMinutes (Number, default `60`)
- TimeZone (String, default `Asia/Manila`)
- ScheduleSheetId (String, external file ID)
- ScheduleTabName (String, `NEW Schedule File`)

#### 4. Additional Sheets
The application utilizes the following sheets for legacy and system tracking:
- `System Log`
- `Feedback`
- `Deleted Requests`
- `Update Log 1.0`

## Development Notes
- `Coverage.gs` contains pure functions for schedule parsing and coverage logic. It can be unit-tested via Node.js (`npm test`).
- Uses `CacheService` to minimize reads to the external schedule spreadsheet.
- Implements lock patterns to avoid concurrency issues during status updates.

## Deployment
1. Open the Apps Script editor.
2. Deploy -> New Deployment.
3. Type: Web App.
4. Execute as: User accessing the web app.
5. Who has access: Anyone within your Google Workspace domain.
