# 🍗 Payroll System - Complete Setup Guide

**For Chick-fil-A Locations (or any restaurant!)**

This guide is written for someone with **zero technical experience**. Follow each step exactly and you'll have a working payroll system in about 30-45 minutes.

---

## 📋 Table of Contents

1. [Can Other Restaurants Use This?](#-can-other-restaurants-use-this)
2. [What You'll Need](#-what-youll-need)
3. [Part 1: Create the Spreadsheet](#part-1-create-the-spreadsheet-5-minutes)
4. [Part 2: Add the Code](#part-2-add-the-code-15-20-minutes)
5. [Part 3: Deploy the App](#part-3-deploy-the-app-5-minutes)
6. [Part 4: Configure Settings](#part-4-configure-settings-10-minutes)
7. [Part 5: Add Employees](#part-5-add-employees)
8. [Part 6: Set Up Employee Forms](#part-6-set-up-employee-forms)
9. [Troubleshooting](#-troubleshooting)
10. [Complete File List](#-complete-file-list)

---

## 🤔 Can Other Restaurants Use This?

**YES!** This system can be copied and used by any restaurant or business that needs to:
- Track overtime hours
- Manage uniform orders and payroll deductions
- Handle PTO requests
- Process bi-weekly payroll

### What You Need to Customize:

| Setting | Where to Change | Example |
|---------|-----------------|---------|
| Location Names | Settings page | "Downtown Store", "Mall Location" |
| Manager Passcode | PAYROLL_SETTINGS sheet | Any number you choose |
| Admin Email | Settings page | manager@email.com |
| Hourly Wage (for OT estimates) | Settings page | $15.00 |
| Logo/Branding | Code (optional) | Your restaurant name |

### What Works Out of the Box:
- ✅ Uniform ordering system
- ✅ Overtime tracking
- ✅ PTO requests (English/Spanish)
- ✅ Payroll deduction calculations
- ✅ Employee management
- ✅ All reports and dashboards

---

## 📦 What You'll Need

Before starting, make sure you have:

- [ ] A **Google account** (Gmail)
- [ ] Access to **Google Drive**
- [ ] The **OTTrackerPro folder** with all code files
- [ ] About **30-45 minutes** of uninterrupted time
- [ ] A computer (not a phone/tablet)

---

## Part 1: Create the Spreadsheet (5 minutes)

### Step 1.1: Open Google Drive

1. Open your web browser (Chrome works best)
2. Go to **drive.google.com**
3. Sign in with your Google account

### Step 1.2: Create a New Spreadsheet

1. Click the big **"+ New"** button (top left corner)
2. Click **"Google Sheets"**
3. A blank spreadsheet will open

### Step 1.3: Name Your Spreadsheet

1. Click on **"Untitled spreadsheet"** at the top left
2. Type a name like: `Payroll System 2025`
3. Press Enter

**💡 Tip:** This spreadsheet will store ALL your data - orders, employees, OT records, etc. Don't delete it!

---

## Part 2: Add the Code (15-20 minutes)

This is the longest part, but just follow each step carefully.

### Step 2.1: Open the Script Editor

1. In your spreadsheet, click **"Extensions"** in the menu bar
2. Click **"Apps Script"**
3. A new browser tab opens with a code editor
4. You'll see a file called `Code.gs` with some default code

### Step 2.2: Create All the Files

You need to create many files. Here's exactly how:

#### Creating Script Files (.gs files)

**To create a .gs file:**
1. Click the **"+"** button next to "Files" on the left
2. Select **"Script"**
3. Type the name (without .gs - it adds that automatically)

Create these script files:
- `Code` (already exists - just rename if needed)
- `DashboardModule`
- `PTOModule`
- `ReportsModule`

#### Creating HTML Files

**To create an HTML file:**
1. Click the **"+"** button next to "Files"
2. Select **"HTML"**
3. Type the name (without .html)

Create these HTML files:
```
Index
MainApp
MainAppJS
Styles
JavaScript
View_Dashboard
View_OTUpload
View_OTHistory
View_OTEmployee
View_OTTrends
View_UniformOrders
View_UniformDeductions
View_UniformCatalog
View_PTORecords
View_PayrollProcessing
View_SettingsPage
View_SystemHealth
View_Help
View_PayrollCalendar
View_YearEndWizard
EmployeeUniformRequest
EmployeePTORequest
```

### Step 2.3: Copy the Code Into Each File

**For EACH file:**

1. Open the matching file from the OTTrackerPro folder on your computer
2. Select all the text: **Ctrl+A** (Windows) or **Cmd+A** (Mac)
3. Copy: **Ctrl+C** (Windows) or **Cmd+C** (Mac)
4. Click on the matching file in the Apps Script editor
5. Delete any existing content
6. Paste: **Ctrl+V** (Windows) or **Cmd+V** (Mac)
7. **Save!** Press **Ctrl+S** (Windows) or **Cmd+S** (Mac)

**⚠️ Important:** Make sure you:
- Copy ALL the content (some files are very long)
- File names match EXACTLY
- Save after each file

### Step 2.4: Initialize the System

Before deploying, you need to set up the sheets:

1. In the code editor, make sure you're in **Code.gs**
2. Look at the top toolbar for a dropdown that says "Select function"
3. Click it and choose **`manualInit`**
4. Click the **"Run"** button (looks like ▶️)

**First time running? You'll see a permissions popup:**

1. Click **"Review permissions"**
2. Choose your Google account
3. You'll see "Google hasn't verified this app"
4. Click **"Advanced"** (small text at bottom left)
5. Click **"Go to Payroll System (unsafe)"**
6. Click **"Allow"**

**✅ Success looks like this in the log:**
```
Sheets initialized successfully!
OT_History sheet: EXISTS
Employees sheet: EXISTS
Settings sheet: EXISTS
...
```

---

## Part 3: Deploy the App (5 minutes)

This makes your app accessible via a URL.

### Step 3.1: Create Deployment

1. Click the blue **"Deploy"** button (top right)
2. Select **"New deployment"**

### Step 3.2: Configure Settings

1. Click the **gear icon ⚙️** next to "Select type"
2. Choose **"Web app"**
3. Fill in:
   - **Description:** `Version 1.0`
   - **Execute as:** `Me (your-email@gmail.com)`
   - **Who has access:** `Anyone`

### Step 3.3: Deploy

1. Click **"Deploy"**
2. If prompted for permissions again, go through the same process
3. **COPY THE URL** that appears - this is your app's address!

**🔖 Your URL looks like:**
```
https://script.google.com/macros/s/AKfycb...very-long-string.../exec
```

**Save this URL somewhere safe! Bookmark it!**

### Step 3.4: Test It

1. Open the URL in a new browser tab
2. Wait 10-20 seconds (first load is slow)
3. You should see the Dashboard!

---

## Part 4: Configure Settings (10 minutes)

### Step 4.1: Open Settings

1. In your app, click **"Settings"** in the left sidebar

### Step 4.2: General Settings

Configure these options:

| Setting | What to Enter | Example |
|---------|---------------|---------|
| Average Hourly Wage | Your average employee wage | $15.00 |
| OT Multiplier | Usually 1.5 | 1.5 |
| Location 1 | First store name | Cockrell Hill |
| Location 2 | Second store (or blank) | DBU OCV |

### Step 4.3: Email Settings

1. Toggle **"Enable Email Notifications"** ON
2. Enter admin email(s) separated by commas
3. Choose which notifications you want

### Step 4.4: Save

Click **"Save All Settings"** at the bottom!

### Step 4.5: Set Manager Passcode

To set the passcode for the manager receiving panel:

1. Go back to your **Google Spreadsheet** (not the app)
2. Find the sheet tab called **"PAYROLL_SETTINGS"** (bottom of screen)
3. Find the cell for "ManagerPasscode"
4. Enter your desired passcode (numbers only, like `1234`)

---

## Part 5: Add Employees

### Option A: Upload OT Data (Easiest)

When you upload overtime data, employees are automatically created!

1. Go to **Overtime → Upload Data**
2. Upload your CSV from CFA Home
3. Employees are auto-created

### Option B: Manual Entry

1. Go to your Google Spreadsheet
2. Click the **"Employees"** sheet tab
3. Add employees with columns:
   - Employee_ID (unique number)
   - Full_Name
   - Primary_Location
   - Status (Active)

---

## Part 6: Set Up Employee Forms

### Uniform Request Form

Employees can order uniforms via a QR code or link.

**Get your Uniform Form URL:**
```
YOUR-APP-URL?page=uniform-request
```

Example: `https://script.google.com/.../exec?page=uniform-request`

### PTO Request Form

**Get your PTO Form URL:**
```
YOUR-APP-URL?page=pto-request
```

### Create QR Codes

1. Go to [qr-code-generator.com](https://www.qr-code-generator.com/)
2. Paste your form URL
3. Download the QR code
4. Print and post in break room

---

## 🔧 Troubleshooting

### "Script function not found: doGet"
- **Cause:** Code.gs wasn't copied correctly
- **Fix:** Re-copy the entire Code.gs file, save, and redeploy

### App Shows White/Blank Screen
- **Cause:** Index.html is missing or incomplete
- **Fix:** Check that Index.html exists and has content

### "Authorization Required" Keeps Appearing
- **Cause:** Permissions weren't fully granted
- **Fix:** 
  1. Go to Apps Script
  2. Run any function manually
  3. Complete the authorization again

### Changes Not Showing Up
- **Cause:** Forgot to create a new version
- **Fix:**
  1. Click **Deploy → Manage deployments**
  2. Click pencil icon ✏️
  3. Under "Version", select **"New version"**
  4. Click **Deploy**

### Email Notifications Not Working
- Check email addresses are correct
- Check spam folder
- Google limits: ~100 emails/day

### ? Help Button Not Working
- Make sure Styles.html is complete (includes help modal CSS)
- Redeploy the app

---

## 📁 Complete File List

### Backend Files (.gs) - 4 files
| File | Purpose |
|------|---------|
| Code.gs | Main backend logic (10,000+ lines) |
| DashboardModule.gs | Dashboard calculations |
| PTOModule.gs | PTO management |
| ReportsModule.gs | Report generation |

### Core HTML Files - 6 files
| File | Purpose |
|------|---------|
| Index.html | Entry point |
| MainApp.html | App shell, sidebar, modals |
| MainAppJS.html | Navigation JavaScript |
| Styles.html | All CSS (4,000+ lines) |
| JavaScript.html | Main JavaScript (9,000+ lines) |

### View Files - 14 files
| File | Page |
|------|------|
| View_Dashboard.html | Home dashboard |
| View_OTUpload.html | Upload OT data |
| View_OTHistory.html | OT history table |
| View_OTEmployee.html | Individual employee OT |
| View_OTTrends.html | OT trends/graphs |
| View_UniformOrders.html | Manage uniform orders |
| View_UniformDeductions.html | Deduction schedules |
| View_UniformCatalog.html | Uniform item catalog |
| View_PTORecords.html | PTO requests |
| View_PayrollProcessing.html | Pre-payroll checklist |
| View_SettingsPage.html | App settings |
| View_SystemHealth.html | System diagnostics |
| View_Help.html | Help & documentation |
| View_PayrollCalendar.html | Payroll calendar |
| View_YearEndWizard.html | Year-end closing |

### Employee Forms - 2 files
| File | Purpose |
|------|---------|
| EmployeeUniformRequest.html | Uniform order form (bilingual) |
| EmployeePTORequest.html | PTO request form (bilingual) |

**Total: ~26 files, ~22,000+ lines of code**

---

## 🎉 You're Done!

Your payroll system is now ready to use. Here's what to do next:

1. **Upload your first OT data** - Tests the system and creates employees
2. **Create a test uniform order** - Make sure the flow works
3. **Print QR codes** - Post them for employees
4. **Train your team** - Show them the Help page

### Daily Workflow

**Monday:** Upload OT data from CFA Home
**As needed:** Process uniform orders, approve PTO
**Before payroll:** Run Pre-Payroll Validation, process deductions

---

## 🎮 Easter Eggs

There are 6 hidden features in the app. Can you find them all?

**Hint:** Check the footer for clues... 🥚

---

*Setup Guide v3.0 - December 2025*
*Built by Brent for Cockrell Hill DTO & Dallas Baptist University OCV*
*"Built to last"*


