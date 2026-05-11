# Academic Management System (AMS) — Enderun Colleges

> Centralized institutional data and student records into a unified platform, creating a secure and scalable academic management ecosystem.

A web-based portal for academic departments to digitize, route, and audit faculty/student academic requests — from DTR discrepancies to make-up class scheduling — with role-based approvals, automated email notifications, and a complete audit trail. Built entirely on the Google Workspace stack (Apps Script + Sheets + Gmail + Drive) so it deploys without external servers.

---

## Table of Contents

- [Overview](#overview)
- [Key Features](#key-features)
- [Screenshots](#screenshots)
- [Tech Stack](#tech-stack)
- [System Architecture](#system-architecture)
- [Roles & Permissions](#roles--permissions)
- [Request Workflow](#request-workflow)
- [Setup & Deployment](#setup--deployment)
- [Project Structure](#project-structure)
- [Security Notes](#security-notes)

---

## Overview

The Academic Management System (AMS) replaces fragmented paper forms, email threads, and spreadsheets that academic departments traditionally use for routine internal requests. Faculty and staff submit applications through a single public-facing portal; department heads and deans review and approve through a private admin portal; and every action is logged for compliance.

The system supports five distinct request types:

| Form | Purpose |
|------|---------|
| **DTR Discrepancy** | Report attendance / time-record discrepancies |
| **Online Modality** | Request to move a class to an online platform |
| **Late Grades** | Submit late-grade entries with justification |
| **Evaluation** | Submit faculty academic performance evaluations |
| **Make-up Class** | Schedule a replacement for missed class sessions |

---

## Key Features

### For Requestors (public)
- Bento-style form gallery — pick the request type and submit in under a minute
- File attachments routed to a shared Drive folder (`AMS_Attachments`)
- Auto-generated **Request ID** returned on submit, also emailed for tracking
- Live status tracker — paste your Request ID to see real-time progress
- Email confirmation on every status change (submitted, endorsed, approved, rejected)

### For Department Heads / Deans (admin)
- **In-Progress queue** — only requests awaiting your review surface in your inbox
- **One-click email approval** — approve from the notification email without opening the portal (token-secured link)
- Row-level RBAC: department heads only see requests from their own department
- Bulk CSV export of any queue
- **Progressed view** — searchable archive of approved & rejected requests

### For Administrators
- **Dashboard** with submission trends, approvals-by-form-type charts, and recent activity feed
- **User Management** — add/edit/deactivate Department Heads & Deans, assign them to specific forms
- **Audit Logs** — every login, submission, approval, and admin action timestamped and searchable; exportable as CSV
- **Weekly digest emails** auto-sent to admin users summarizing the week's activity
- PDF generation of any approved request for record-keeping (stored in `AMS_PDFs` Drive folder)

### Cross-cutting
- SHA-256 password hashing
- First-login forced password change
- Forgot-password flow with temporary password via Gmail
- Account activation via email link
- Mobile-responsive UI

---

## Screenshots

### Public Request Portal
The landing page where faculty and staff pick a request type. The same view also exposes the live status tracker.

![Public request portal](AMS%20SCREENSHOTS/Screenshot%202026-05-09%20172246.png)

### Dynamic Request Form (DTR Discrepancy example)
Every form type renders into a consistent layout with the appropriate fields, validation, and optional file attachment.

![DTR Discrepancy form](AMS%20SCREENSHOTS/Screenshot%202026-05-09%20172339.png)

### Admin Dashboard
At-a-glance KPIs (total / in-progress / approved / rejected), a submission trend line, and an approvals-by-form-type breakdown.

![Admin dashboard](AMS%20SCREENSHOTS/Screenshot%202026-05-09%20172851.png)

### In-Progress Requests
The reviewer's working queue. Search, filter, and export. Empty state shown when the queue is clear.

![In-Progress requests](AMS%20SCREENSHOTS/Screenshot%202026-05-09%20172934.png)

### Progressed Requests
Searchable archive of approved and rejected requests with KPI totals at the top.

![Progressed requests](AMS%20SCREENSHOTS/Screenshot%202026-05-09%20173003.png)

### User Management
Admin-only — add department heads and deans, assign them to specific forms / departments, activate or deactivate accounts.

![User Management](AMS%20SCREENSHOTS/Screenshot%202026-05-09%20173029.png)

### Audit Logs
Immutable, timestamped record of every meaningful action: logins, submissions, approvals, user-management changes.

![Audit Logs](AMS%20SCREENSHOTS/Screenshot%202026-05-09%20173051.png)

### Profile Settings
Users update their display name, email, and password from a single panel.

![Profile Settings](AMS%20SCREENSHOTS/Screenshot%202026-05-09%20173114.png)

---

## Tech Stack

| Layer | Technology |
|-------|------------|
| **Frontend** | HTML, CSS, Vanilla JavaScript |
| **UI helpers** | Chart.js (dashboard charts), Font Awesome (icons), Google Fonts (Cormorant Garamond + DM Sans) |
| **Backend** | Google Apps Script (server-side JavaScript) |
| **Database** | Google Sheets (relational-style sheets per form + Users + System_Logs) |
| **Email** | Gmail (GmailApp service) |
| **File storage** | Google Drive (auto-created `AMS_Attachments` & `AMS_PDFs` folders) |
| **Auth** | Custom user table + SHA-256 hashed passwords |
| **Hosting** | Apps Script Web App deployment (no external server) |

---

## System Architecture

```
┌──────────────────────────────────────────────────────────────────┐
│                         Web Browser                              │
│   index.html  ⇄  google.script.run  ⇄  Apps Script (code.gs)     │
└──────────────────────────────────────────────────────────────────┘
                                │
                ┌───────────────┼───────────────┐
                ▼               ▼               ▼
        ┌──────────────┐ ┌────────────┐ ┌──────────────┐
        │ Google Sheet │ │  Gmail App │ │  Drive App   │
        │   (the DB)   │ │ (Emails)   │ │ (Attachments)│
        └──────────────┘ └────────────┘ └──────────────┘
```

**Sheets used (auto-created on first run via `_ensureSheet`):**

- `Users` — `User_ID, Name, Email, Role, Assigned_Form, Status, Date_Added, Password, IsFirstLogin`
- `Form1_Data` (DTR) — `Request_ID, Name, Department, Date_Of_Discrepancy, Time_In_Out, N_A, Reason, Status, Date_Submitted, Remarks, Attachment, Auth_Token`
- `Form2_Data` (Online Modality) — `Course_Details, Class_Sched, Proposed_Date, Workplan, …`
- `Form3_Data` (Late Grades) — `Semester, Term_Grade, …, Class_List, …`
- `Form4_Data` (Evaluation) — `Prof_Name, Subject, Rating, Feedback, …`
- `Form5_Data` (Make-up) — `Subject, Missed_Class, Makeup_Class, Reason, …`
- `System_Logs` — `Timestamp, User, Module, Description`

---

## Roles & Permissions

| Role | Sees | Can Approve |
|------|------|-------------|
| **Requestor** (public) | Their own request status via tracker | — |
| **Dept Head** | Only requests in their assigned form + their own department, in status `Pending Program Head` | Endorses to Dean |
| **Dean** | All requests in `Pending Dean` status | Final approve / reject |
| **Admin** | All requests, all users, audit logs, dashboards | Full administrative control |

Row-level RBAC is enforced server-side in [code.gs](code.gs) — clients cannot request data outside their scope.

---

## Request Workflow

```
  Requestor submits      →   Pending Program Head
        │                              │
        │  Email to Dept Head          │  Dept Head reviews
        │  (with Quick-Approve link)   │  (in portal or email)
        ▼                              ▼
  Email confirmation         Pending Dean
                                       │
                                       │  Dean reviews
                                       ▼
                              Approved / Rejected
                                       │
                                       │  Email to requestor +
                                       │  PDF archived to Drive
                                       ▼
                                Audit log entry
```

Every state transition writes to `System_Logs` and triggers an email to the relevant parties.

---

## Setup & Deployment

### 1. Create the host Google Sheet
1. Create a new Google Sheet — this will be your database
2. Open **Extensions → Apps Script**

### 2. Paste the code
1. Replace `Code.gs` with the contents of [code.gs](code.gs) from this repo
2. Add a new HTML file named `index` and paste the contents of [index.html](index.html)

### 3. Configure the super-admin
Open [code.gs](code.gs) and edit the two constants at the top of the authentication section:

```javascript
var ADMIN_EMAIL    = 'your-admin-email@example.com';
var ADMIN_PASSWORD = 'a-strong-password-you-choose';
```

These bootstrap the first admin who will then create other users from inside the portal.

### 4. Initialize the workbook (one-time)
From the Apps Script editor, run the `initializeAllSheets` function once to create all the sheets and headers automatically.

### 5. Deploy as a Web App
1. **Deploy → New deployment → Web app**
2. Execute as: **Me**
3. Who has access: **Anyone** (or your domain, per your policy)
4. Copy the generated URL — that's your portal

### 6. Grant Gmail / Drive scopes
On first deployment, Apps Script will prompt you to authorize Gmail and Drive scopes. These are required for outbound notifications and attachment storage.

---

## Project Structure

```
.
├── code.gs                  # Apps Script backend (routing, RBAC, email, PDF)
├── index.html               # Single-page frontend (login, forms, admin panels)
├── AMS SCREENSHOTS/         # UI screenshots referenced in this README
└── README.md
```

---

## Security Notes

- Passwords are hashed with **SHA-256** before storage; plaintext is never persisted
- **Quick-approve email links** are protected by per-request UUID tokens stored in the sheet — links cannot be guessed or reused after status changes
- All RBAC filtering happens **server-side** (`google.script.run` handlers in [code.gs](code.gs)) — the client never gets to see data outside its role
- The super-admin credentials live as constants in [code.gs](code.gs). **Do not commit real credentials**; the published version of this repo uses placeholders that must be replaced before deployment
- Apps Script runs under the deployer's Google account, so all Sheet/Drive/Gmail access is scoped to that identity

---

## License

This project is shared for educational and portfolio purposes. Please contact the author before reuse in production environments.
