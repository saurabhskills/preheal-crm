# Preheal CRM (Google Apps Script + Google Sheets)

Preheal CRM is a lightweight, role-based lead management system built as a **Google Apps Script Web App** with **Google Sheets** as the database. It supports user registration + admin approval, lead intake (single + bulk CSV), lead updates with full history, and dashboards for Admin and TeleCallers.

---

## Features

### Authentication & User Management
- Login / Register UI.
- User accounts stored in **Users** sheet.
- New registrations are marked **PENDING** and must be approved by Admin.
- Approved users can log in and are routed to role-specific pages.

### Roles (current behavior)
- **Admin**
  - Dashboard (overall counts + performance).
  - User approvals.
  - Create leads (intake) and assign telecallers during creation.
  - Update leads (all leads).
  - View all leads.
  - Export leads as CSV.
- **Manager** (depends on your current sidebar config)
  - Lead intake + bulk import.
  - View leads (server filtering applies).
- **Executive**
  - Create leads (intake) and assign telecallers during creation.
  - Update leads they created.
  - View their leads (server filtering applies).
  - Export leads as CSV.
- **TeleCaller**
  - Dashboard showing leads to contact today & upcoming.
  - Update their assigned leads (status, follow-up, call notes).
  - View their leads (server filtering applies).

### Lead Management
- **Create Lead**
  - From Intake Form.
  - Supports assigning a TeleCaller at creation time via dropdown.
- **Bulk Upload**
  - CSV upload (browser reads file, server parses, inserts into Leads sheet).
- **Update Lead**
  - Updates: LeadStatus, FollowUpStatus, FollowUpDate, Issue, CallNotes.
  - CallNotes are appended (not overwritten).
- **History Tracking**
  - Before any lead update, the system saves a snapshot of the old row into `LeadsHistory`.

### Dashboards
- **Admin Dashboard**
  - Total leads
  - Leads by status
  - Executive performance (CreatedBy)
  - TeleCaller performance (AssignedTo)
- **TeleCaller Dashboard**
  - Assigned count
  - Contact today (FollowUpDate == today)
  - Upcoming (FollowUpDate > today)

### CSV Export
- “Download CSV” button downloads leads visible to the logged-in role.
- Export is generated server-side and downloaded client-side.

---

## Tech Stack
- Google Apps Script (server)
- HTMLService (web app hosting)
- Google Sheets (database)
- TailwindCSS (UI styling via CDN)
- Font Awesome (icons)

---

## Project Structure

### `Code.gs` (Server-side)
Contains all backend logic:
- Auth: `loginUser()`, `registerUser()`
- Admin approval: `getPendingUsers()`, `approveUser()`
- Leads: `addLead()`, `getLeads()`, `updateLeadWithHistory()`, etc.
- Bulk import:
  - CSV import function (recommended)
- Dashboards:
  - `getAdminDashboard()`
  - `getTelecallerDashboard()`
- Telecaller list for assignment dropdown:
  - `getApprovedTeleCallers_()` (private helper)
  - `getApprovedTelecallers()` (public wrapper callable from UI)

### `Index.html` (Client-side UI)
Single-page interface containing:
- Login/Register screen
- Sidebar layout + main content area
- Views:
  - Admin Dashboard
  - TeleCaller Dashboard
  - Intake Form
  - Bulk Upload (CSV)
  - Leads table view
  - Update Leads view
  - Approval view
- Client-to-server calls using `google.script.run`

---

## Google Sheets Database

### 1) `Users` sheet
**Columns (recommended order):**
1. Email
2. Password
3. Role
4. Status
5. Name

**Status values:**
- `PENDING` (default on registration)
- `APPROVED` (admin approves)

### 2) `Leads` sheet
**Columns (A..M):<br>**
A. LeadID  
B. LeadName  
C. Phone  
D. LeadSource  
E. AssignedTo  
F. CallNote  
G. Issue  
H. FollowUpDate  
I. FollowUpStatus  
J. LeadStatus  
K. CreatedBy  
L. CreatedDate  
M. LastUpdatedDate  

### 3) `LeadsHistory` sheet
Stores a snapshot of the lead row BEFORE every update (audit trail), plus metadata like:
- LoggedAt
- UpdatedBy
- LeadID
- LeadRowNumber
- ChangeJSON
- Full row snapshot fields (mirroring the Leads row)

---

## LeadID Format

The project uses a date-based LeadID format:

`L0DDMMYY0XXX`

Where:
- `DD` = day (2 digits)
- `MM` = month (2 digits)
- `YY` = last two digits of year
- `XXX` = sequential number (3 digits), reset daily

Example:
- `L00801260 001` (spaces shown only for clarity)

This format helps generate daily/weekly/monthly/yearly reports by parsing the embedded date.

> Implementation uses a daily sequence counter stored in script properties and a document lock to prevent collisions.

---

## Date Format System

### Follow-up dates
- Stored as a string in `Leads!H` as: `DD/MM/YYYY`
- UI uses `<input type="date">` which provides `YYYY-MM-DD`, so the UI converts:
  - ISO → DMY when saving
  - DMY → ISO when showing in the date input

### CreatedDate / LastUpdatedDate
- Stored as actual Date objects in the sheet (Apps Script `new Date()`).

---

## Setup & Deployment

### 1) Create the Spreadsheet
Create a Google Sheet with tabs:
- `Users`
- `Leads`
- `LeadsHistory` (optional, can be auto-created if your code includes an “ensure” helper)

Add header rows matching the column definitions above.

### 2) Create Apps Script Project
- Open the Spreadsheet → Extensions → Apps Script
- Paste:
  - `Code.gs` content
  - `Index.html` content

### 3) Deploy as Web App
- Deploy → New deployment → Web app
- Execute as: **Me**
- Access: as required (usually “Anyone with link” for testing; tighten for production)

---

## CSV Import (Bulk Upload)

### Expected CSV headers
At minimum:
- `Name`

Supported:
- `Phone`
- `Source`
- `CallNotes`
- `FollowUpDate` (DD/MM/YYYY preferred)
- `AssignedTo` (telecaller email)

Example header row:
```csv
Name,Phone,Source,CallNotes,FollowUpDate,AssignedTo
