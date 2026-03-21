# Training Tracker — Technical Setup Guide

This guide walks you through the full installation of the Chick-fil-A Training Tracker system. It's written for whoever is actually setting it up — you don't need to be a developer, but you do need to follow the steps carefully.

**Time estimate:** 30-45 minutes for first-time setup.

---

## Table of Contents

1. [Prerequisites](#prerequisites)
2. [Create Your Google Sheet](#step-1-create-your-google-sheet)
3. [Open the Apps Script Editor](#step-2-open-the-apps-script-editor)
4. [Create the Script Files](#step-3-create-the-script-files)
5. [Create the HTML Files](#step-4-create-the-html-files)
6. [Save and Run Initial Setup](#step-5-save-and-run-initial-setup)
7. [Authorize the Script](#step-6-authorize-the-script)
8. [Create Your Google Form](#step-7-create-your-google-form)
9. [Link the Form to the Sheet](#step-8-link-the-form-to-the-sheet)
10. [Set Up Triggers](#step-9-set-up-triggers)
11. [Configure Email Alerts](#step-10-configure-email-alerts)
12. [Set Up Certification Forms (Optional)](#step-11-set-up-certification-forms-optional)
13. [Customize for Your Location](#step-12-customize-for-your-location)
14. [Set Up Monday Auto-Populate (Optional)](#step-13-set-up-monday-auto-populate-optional)
15. [Verify Everything Works](#step-14-verify-everything-works)
16. [Troubleshooting](#troubleshooting)
17. [File Reference](#file-reference)

---

## Prerequisites

You need:
- A Google account (the account that will "own" the system)
- Access to Google Sheets, Google Forms, and Google Apps Script (all free with any Google account)
- The script files from this repository (the `.gs` and `.html` files)

> **Tip:** Use a shared/store Google account rather than a personal one. This way the triggers and alerts keep running even if an individual manager leaves.

---

## Step 1: Create Your Google Sheet

1. Go to [sheets.google.com](https://sheets.google.com)
2. Create a new blank spreadsheet
3. Name it something like **"[Your Location] Training Tracker"**

This is the spreadsheet everything will run from. The script will create all the internal sheets (tabs) for you automatically.

> If your location already has a "Training Needs" sheet with a specific layout (days of the week stacked vertically, daypart sections with "FOH Training" headers), keep it — the system will detect and work with that layout automatically.

---

## Step 2: Open the Apps Script Editor

1. In your new spreadsheet, click **Extensions** in the menu bar
2. Click **Apps Script**
3. A new browser tab opens with the script editor
4. You'll see a default file called `Code.gs` — you can delete the contents of this file (or delete the file entirely)

---

## Step 3: Create the Script Files

For each script file below, do this:
1. In the Apps Script editor, click the **+** button next to **Files** in the left sidebar
2. Select **Script**
3. Name it exactly as shown (the `.gs` extension is added automatically — don't type it)
4. Delete any default content in the new file
5. Copy the full contents of the corresponding `.gs` file from this repository and paste it in

| Create this file in the editor | Paste contents from |
|-------------------------------|-------------------|
| `01_Menu_and_Setup` | `01_Menu_and_Setup.gs` |
| `02_Form_and_Dedup` | `02_Form_and_Dedup.gs` |
| `03_Dashboard` | `03_Dashboard.gs` |
| `04_Timeline` | `04_Timeline.gs` |
| `05_Certification` | `05_Certification.gs` |
| `06_Alerts` | `06_Alerts.gs` |
| `07_UI_Functions` | `07_UI_Functions.gs` |
| `08_DataSync` | `08_DataSync.gs` |

> **Optional:** You can also add `DIAGNOSTIC.gs` — it's a debug utility that helps troubleshoot the Training Needs layout detection. Useful during initial setup, can be removed once everything is working.

---

## Step 4: Create the HTML Files

These are the UI dialogs and sidebars. The process is similar but you select **HTML** instead of **Script**:

1. Click the **+** button next to **Files**
2. Select **HTML**
3. Name it exactly as shown (the `.html` extension is added automatically)
4. Delete any default content and paste in the file contents

| Create this file | Paste contents from |
|-----------------|-------------------|
| `TimelineDialog` | `TimelineDialog.html` |
| `DeduplicationSidebar` | `DeduplicationSidebar.html` |
| `AlertSettingsDialog` | `AlertSettingsDialog.html` |

When you're done, your Files panel should show 8 `.gs` files and 3 `.html` files (plus the original `Code.gs` if you didn't delete it — that's fine, just make sure it's empty).

---

## Step 5: Save and Run Initial Setup

1. Press **Ctrl+S** (or **Cmd+S** on Mac) to save all files
2. Go back to your spreadsheet tab in the browser
3. **Reload the page** (F5 or Cmd+R)
4. Wait 3-5 seconds — a new **Training Tools** menu should appear in the menu bar (next to Help)
5. Click **Training Tools > Initial Setup (Run Once)**

> If the Training Tools menu doesn't appear after reloading, wait a few more seconds and try reloading again. The `onOpen()` function needs a moment to register.

---

## Step 6: Authorize the Script

The first time you run any function, Google will ask you to authorize the script:

1. A dialog says "Authorization required" — click **Continue**
2. Select your Google account
3. You may see a warning that says **"Google hasn't verified this app"** — this is normal for custom scripts
   - Click **Advanced** (small text at the bottom left)
   - Click **"Go to [Your Project Name] (unsafe)"**
4. Review the permissions and click **Allow**

**What permissions does it need and why:**

| Permission | Why |
|-----------|-----|
| See, edit, create, and delete spreadsheets | Reads/writes training data, creates timeline sheets |
| Send email as you | Sends alert emails to managers |
| Display and run third-party web content | Shows the timeline dialog, dedup sidebar, and alert settings UI |

After authorizing, the Initial Setup runs and creates all the sheets. You should see a confirmation dialog listing what was created.

---

## Step 7: Create Your Google Form

If you don't already have a training log form, create one:

1. Go to [forms.google.com](https://forms.google.com)
2. Create a new blank form
3. Name it **"Daily Training Log"** (or similar)
4. Add these questions **in this exact order**:

| # | Question | Type | Required? |
|---|----------|------|-----------|
| 1 | Date | Date | Yes |
| 2 | Trainee Name | Short answer | Yes |
| 3 | Position Trained | Dropdown or Short answer | Yes |
| 4 | Hours | Short answer (number) | Yes |
| 5 | On Track? | Multiple choice (Yes / No) | Yes |
| 6 | Notes | Paragraph (long answer) | No |

> **For Position Trained:** A dropdown is recommended so names stay consistent. List all your positions exactly as they appear in the Position Requirements sheet (e.g., "iPOS", "Register/POS", "Breading", etc.). The system has alias handling, but consistent names are better.

> **Important:** The order matters. The script reads form responses by column position: A=Timestamp (auto), B=Date, C=Name, D=Position, E=Hours, F=On Track, G=Notes. If your form has a different question order, you'll need to adjust the column indices in `02_Form_and_Dedup.gs` and `08_DataSync.gs`.

---

## Step 8: Link the Form to the Sheet

This connects your form to the Training Spreadsheet so responses automatically land in the Daily Training Log.

1. Open your Google Form
2. Click the **Responses** tab at the top
3. Click the **Link to Sheets** icon (green Sheets icon)
4. Select **"Select existing spreadsheet"**
5. Find and select your Training Tracker spreadsheet
6. **Important:** When asked which sheet to link to, you have two options:
   - If Google offers to create a new sheet named "Form Responses 1" — that's fine. The system can read from it.
   - If you can select an existing sheet, choose **"Daily Training Log"**

> **If the form creates a "Form Responses 1" sheet:** That's OK. Use **Training Tools > Sync Form Data** to import responses into the Daily Training Log. The system checks for sheets named `Form_Responses`, `Form Responses 1`, `Form responses 1`, and `Form_Responses_1`.

**Verify the connection:**
1. Submit a test response through the form
2. Check your spreadsheet — the response should appear in either "Daily Training Log" or "Form Responses 1"
3. If it went to "Form Responses 1", click **Training Tools > Sync Form Data** to import it

---

## Step 9: Set Up Triggers

Triggers make the system run automatically. Go to the Apps Script editor:

1. Click the **clock icon** (Triggers) in the left sidebar
2. Click **+ Add Trigger** in the bottom-right corner
3. Create each of these triggers:

### Trigger 1: Form Submission Handler
| Setting | Value |
|---------|-------|
| Choose which function to run | `onFormSubmit` |
| Choose which deployment should run | Head |
| Select event source | From spreadsheet |
| Select event type | On form submit |

This is the most important trigger. It processes every form submission in real-time: resolves names, checks milestones, updates the dashboard, and sends alerts.

### Trigger 2: Daily Training Reminder
| Setting | Value |
|---------|-------|
| Choose which function to run | `dailyReminderCheck` |
| Choose which deployment should run | Head |
| Select event source | Time-driven |
| Select type of time based trigger | Day timer |
| Select time of day | 8pm to 9pm |

Sends an email if no training was logged today. Useful for catching days when forms weren't submitted.

### Trigger 3: Inactive Trainee Check
| Setting | Value |
|---------|-------|
| Choose which function to run | `checkInactiveTrainees` |
| Choose which deployment should run | Head |
| Select event source | Time-driven |
| Select type of time based trigger | Day timer |
| Select time of day | 6am to 7am |

Sends an email listing any trainees who haven't trained in 3+ days.

### Trigger 4: Duplicate Name Scan
| Setting | Value |
|---------|-------|
| Choose which function to run | `checkForDuplicates` |
| Choose which deployment should run | Head |
| Select event source | Time-driven |
| Select type of time based trigger | Day timer |
| Select time of day | 7am to 8am |

Scans all names in the training log and flags potential duplicates (typos, nicknames, etc.).

### Trigger 5 (Optional): Certification Form Response
Only set this up if your certification feedback forms feed responses back into this same spreadsheet.

| Setting | Value |
|---------|-------|
| Choose which function to run | `onCertificationFormSubmit` |
| Choose which deployment should run | Head |
| Select event source | From spreadsheet |
| Select event type | On form submit |

> **Note on trigger 5:** If you use standalone certification forms that are NOT linked to this spreadsheet, skip this trigger. Use **Training Tools > Manually Certify Trainee** instead.

---

## Step 10: Configure Email Alerts

1. Go to your spreadsheet
2. Click **Training Tools > Alert Settings**
3. A dialog opens with 5 alert types:

| Alert Type | What It Does | Recommended Recipients |
|-----------|-------------|----------------------|
| Daily Training Log Reminder | Fires if no training logged today | Training Director |
| Trainee Inactive 3+ Days | Lists trainees who haven't trained recently | Training Director, Ops Director |
| Position Completion Milestone | Fires when a trainee completes a position's minimum hours | Training Director |
| Trainee Ready for Certification | Fires when all position minimums are met | Training Director, Operator |
| Duplicate Name Detected | Fires when potential name duplicates are found | Whoever manages the spreadsheet |

4. For each alert, toggle it on/off and enter up to 3 email addresses
5. Click **Save Settings**

---

## Step 11: Set Up Certification Forms (Optional)

If you want to use Google Forms as the final sign-off step:

### Create the Forms
1. Create two Google Forms:
   - **"FOH Final Trainer Feedback"**
   - **"BOH Final Trainer Feedback"**
2. Include fields like: Trainee Name, Trainer Name, ratings/feedback per position, overall readiness, notes

### Connect Them to the Script
1. Open `05_Certification.gs` in the Apps Script editor
2. Find the `openCertificationForm` function (near the top)
3. Replace the `fohFormUrl` value with your FOH form's URL
4. Replace the `bohFormUrl` value with your BOH form's URL
5. Update the `entry.123456789` parameter with the actual entry ID for the name field:
   - Open your form
   - Click the **three-dot menu (...)** > **Get pre-filled link**
   - Fill in the trainee name field with any text
   - Click **Get link**
   - In the generated URL, find the part that says `entry.XXXXXXXXX` — copy that number
   - Replace `123456789` in the script with your number

> **If this feels too complicated:** Just skip it and use **Training Tools > Manually Certify Trainee** instead. It works the same way — adds the trainee to the Certification Log and removes them from the active dashboard.

---

## Step 12: Customize for Your Location

### Positions and Hours
1. Go to the **Position Requirements** sheet
2. Edit positions, add new ones, remove ones you don't use
3. Adjust the Minimum, Maximum, and Target hours
4. The system uses the **Target Hours** column when generating timelines and the **Minimum Hours** column when checking certification readiness

### Position Name Aliases
If your team uses shorthand names on the form that differ from Position Requirements, add aliases in `03_Dashboard.gs`:

Find the `normalizePositionName` function and add entries to the `aliases` object:
```javascript
var aliases = {
  'register': 'Register/POS',
  'pos':      'Register/POS',
  // Add your custom aliases here:
  'drive':    'DT Drinks',
  'front':    'FC Drinks',
};
```

### Daypart Time Windows
If your location's daypart boundaries differ from the defaults, edit `getDaypartsForShift()` in `04_Timeline.gs`:

```javascript
var DAYPARTS = [
  { name: 'Breakfast', start: 6,  end: 11 },  // 6 AM - 11 AM
  { name: 'Lunch',     start: 11, end: 15 },  // 11 AM - 3 PM
  { name: 'Afternoon', start: 15, end: 17 },  // 3 PM - 5 PM
  { name: 'Dinner',    start: 17, end: 22 }   // 5 PM - 10 PM
];
```

### Shift Options in the Timeline Dialog
To change the shift options available when generating a timeline, edit `TimelineDialog.html`:

Find the `<select id="shift">` element and edit the options:
```html
<select id="shift">
  <option value="Breakfast">Breakfast (6a-11a)</option>
  <option value="Lunch">Lunch (11a-3p)</option>
  <option value="Afternoon">Afternoon (3p-5p)</option>
  <option value="Dinner">Dinner (5p-10p)</option>
</select>
```

---

## Step 13: Set Up Monday Auto-Populate (Optional)

If you want the Training Needs sheet to be automatically filled in each Monday morning:

1. Click **Training Tools > Setup Monday Auto-Populate**
2. This creates a trigger that runs every Monday at ~5 AM
3. It reads the Training Schedule and populates the Training Needs sheet for the current week

> This only works if you've generated timelines for your trainees and have a Training Needs sheet with the expected layout (day names + "FOH Training" headers).

---

## Step 14: Verify Everything Works

Run through this checklist:

- [ ] **Training Tools menu appears** when you open/reload the spreadsheet
- [ ] **All sheets exist:** Daily Training Log, Position Requirements, Master Dashboard, Certification Log, Name Deduplication, Alert Settings, Training Schedule
- [ ] **Position Requirements** has your location's positions and hour targets
- [ ] **Submit a test form response** and verify it appears in the Daily Training Log
- [ ] **Master Dashboard updates** — click Training Tools > Refresh Dashboard
- [ ] **Generate a test timeline** — Training Tools > Generate Training Timeline
- [ ] **Alert Settings** has your email addresses entered
- [ ] **Triggers are set up** — check the clock icon in Apps Script editor

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| **Training Tools menu doesn't appear** | Reload the spreadsheet page. If still missing, open Apps Script and check for syntax errors (red underlines). |
| **"Authorization required" keeps popping up** | Click through the full authorization flow (Continue > Advanced > Go to app > Allow). You must use the same Google account that owns the spreadsheet. |
| **Form responses don't appear in the sheet** | Check that the form is linked to the correct spreadsheet. Go to the form's Responses tab and verify the Sheets link. |
| **Dashboard shows no data** | Click Training Tools > Sync Form Data first, then Training Tools > Refresh Dashboard. |
| **Alerts aren't sending** | Check Alert Settings — make sure the alert is enabled AND has a valid email address. Also check your spam folder. |
| **Timeline doesn't generate** | Make sure Position Requirements has data for the house you selected (FOH or BOH). |
| **Training Needs not populating** | Run the diagnostic: add `DIAGNOSTIC.gs` to your script files, then run `runTrainingDiagnostic()` from the Apps Script editor. It will show you exactly what the layout detection sees. |
| **"Script has exceeded the maximum execution time"** | This can happen with very large datasets. Try running Sync Form Data first to consolidate, then Refresh Dashboard. |
| **Duplicate detection finds too many false positives** | Review and click "Ignore" for non-duplicates. The system learns from your choices. |
| **Position hours seem wrong on dashboard** | Check if the position name on the form matches Position Requirements exactly. Run Training Tools > Sync Form Data to re-import with canonical names. |

### Checking Script Errors

1. Open **Extensions > Apps Script**
2. Click the **Executions** icon in the left sidebar (looks like a play button with a list)
3. This shows every script execution with status (Completed, Failed, Timed Out)
4. Click a failed execution to see the error message and stack trace

---

## File Reference

### Script Files

| File | Purpose | Key Functions |
|------|---------|--------------|
| `01_Menu_and_Setup.gs` | Menu + setup | `onOpen()`, `runInitialSetup()` |
| `02_Form_and_Dedup.gs` | Form handler + dedup | `onFormSubmit()`, `checkForDuplicates()`, `mergeDuplicateNames()` |
| `03_Dashboard.gs` | Dashboard engine | `updateDashboard()`, `normalizePositionName()`, `isCertificationReady()` |
| `04_Timeline.gs` | Timeline + scheduling | `createTimeline()`, `populateWeeklyTraining()`, `getDaypartsForShift()` |
| `05_Certification.gs` | Certification | `openCertificationForm()`, `manuallyCertifyTrainee()` |
| `06_Alerts.gs` | Email alerts | `sendAlert()`, `dailyReminderCheck()`, `checkInactiveTrainees()` |
| `07_UI_Functions.gs` | UI helpers | `showDeduplicationSidebar()`, `showAlertSettings()` |
| `08_DataSync.gs` | Data import | `syncFormData()`, `backfillCanonicalNames()` |

### HTML Files

| File | Purpose |
|------|---------|
| `TimelineDialog.html` | Timeline generation dialog with 3-week day picker |
| `DeduplicationSidebar.html` | Sidebar for reviewing/merging duplicate names |
| `AlertSettingsDialog.html` | Modal dialog for configuring email alerts |

### Sheets Created by Setup

| Sheet | Purpose | Populated By |
|-------|---------|-------------|
| Daily Training Log | Raw training data | Google Form + Sync |
| Position Requirements | Positions + hour targets | Initial Setup (then you edit) |
| Master Dashboard | Live summary view | `updateDashboard()` auto-refresh |
| Training Schedule | Planned weekly assignments | `createTimeline()` |
| Certification Log | Certified trainee archive | Certification form or manual |
| Name Deduplication | Duplicate name suggestions | `checkForDuplicates()` |
| Alert Settings | Email alert config | Initial Setup + Alert Settings dialog |

### Triggers Summary

| Function | Event | Purpose |
|----------|-------|---------|
| `onFormSubmit` | On form submit | Process each training log entry |
| `dailyReminderCheck` | Day timer, 8-9 PM | Remind if no training logged today |
| `checkInactiveTrainees` | Day timer, 6-7 AM | Flag trainees inactive 3+ days |
| `checkForDuplicates` | Day timer, 7-8 AM | Scan for duplicate names |
| `mondayAutoPopulate` | Weekly, Monday ~5 AM | Auto-fill Training Needs sheet |
| `onCertificationFormSubmit` | On form submit (optional) | Process certification form responses |
