# Cockrell Hill Tech Projects

Tech tools and applications for Chick-fil-A operations at Cockrell Hill.

**Live Site:** [brentb154.github.io/Brent-Chick-fil-A-Tech-Projects](https://brentb154.github.io/Brent-Chick-fil-A-Tech-Projects)

## Projects

### Payroll Management System (`/payroll-system`)
Complete payroll solution with PTO tracking, overtime management, uniform ordering, and year-end processing. Built with Google Apps Script and integrates with Google Sheets.

### Schedule Counter (`/schedule-counter-gas`)
Drop in your published schedule and it counts FOH/BOH by the hour, weights it against your sales curve, snapshots every week, and tracks productivity — all inside a Google Sheet, running itself on a weekly trigger. There's a no-setup local version you can try right in your browser first (linked from the landing page and kept in `/schedule-counter`); when you're ready for the automated version, follow `SETUP_GUIDE.md`.

### Inventory Analyzer (`/inventory-analyzer`)
Standalone HTML tool for tracking and analyzing inventory levels. Works entirely in your browser.

### Manager Accountability Hub (`/manager-hub`)
Employee management system for tracking performance and team accountability. Built with Google Apps Script.

### Catering Quote Generator (`/catering-quote-tool`)
Create, manage, email, and print professional catering quotes. Features a searchable menu picker, sequential quote IDs, PDF email integration, and quote history. Built with Google Apps Script and integrates with Google Sheets.

### Shared Table Email Summary (`/shared-table-email-summary`)
Automatically sends weekly waste summary emails from Google Form responses. Configurable send schedule, recipients, and categorized item reporting. Built with Google Apps Script.

### H.E.A.R.D. Log (`/guest-recovery-log`)
Guest-recovery complaint log. Log guest issues in under a minute, auto-flag repeat complainers via phone lookup, live dashboard of trends and resolutions, and configurable email alerts. React frontend with a Google Apps Script + Sheets backend and a one-click first-time setup function.

### Training Tracker (`/training-tracker`)
Turns a Google Form into a live training dashboard — daily training logs, certification tracking by position, a per-person timeline, built-in name deduplication, and automatic alerts when a trainee goes inactive. Built with Google Apps Script and Google Sheets. See `SETUP_GUIDE.md` in the folder.

## How to Use

**Standalone Tools (Inventory Analyzer, Schedule Counter local preview):**
1. Click the link on the website, or
2. Download the HTML file and open it in any browser

**Google Apps Script Projects (Payroll, Schedule Counter, Catering Quote, Training Tracker, H.E.A.R.D. Log, Manager Hub, Shared Table Email):**
1. See the setup guide in each project folder (`SETUP_GUIDE.md` or `README.md`)
2. Requires Google Sheets and Apps Script setup

---
Built with care for Cockrell Hill Chick-fil-A
