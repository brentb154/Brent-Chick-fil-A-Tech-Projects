# Chick-fil-A Training Tracker

A complete training management system for Chick-fil-A operators built on Google Sheets, Google Forms, and Google Apps Script. It automates training timelines, tracks position-by-position progress, detects when team members are certification-ready, and sends email alerts — all from tools your team already uses.

---

## What Problem Does This Solve?

Training a new team member at CFA involves dozens of positions, minimum hour requirements for each, multiple trainers, and a timeline that shifts constantly. Most locations track this on paper, in scattered spreadsheets, or not at all. The result: team members fall through the cracks, managers don't know who's ready for what, and certification readiness is a guessing game.

This system replaces all of that with a single Google Sheet that:

- **Auto-generates a personalized training timeline** based on the new hire's schedule, start date, and position requirements
- **Tracks daily training progress** via a simple Google Form your trainers already know how to use
- **Shows a live dashboard** with every active trainee's hours, position progress, and certification status
- **Sends email alerts** when a trainee completes a position, goes inactive, or is ready for certification
- **Handles name deduplication** automatically (catches "Tim" vs "Timothy", typos, etc.)
- **Populates your weekly Training Needs sheet** so managers know exactly who is training where each day

---

## The Three Components

This system has three parts. They work together but are set up independently.

### Component 1: Training Timeline Spreadsheet (Google Sheets + Apps Script)

This is the core of the system — a Google Sheet with custom Apps Script code that powers everything.

**What it does:**
- Creates a personalized, week-by-week training timeline for each new hire
- Tracks all training hours logged via a connected Google Form
- Maintains a live Master Dashboard showing every trainee's progress
- Auto-populates a weekly Training Needs sheet (who's training where, which daypart)
- Manages certification workflow and archives certified team members
- Sends configurable email alerts to managers

**Sheets created automatically:**
| Sheet | Purpose |
|-------|---------|
| Daily Training Log | Where form responses land — the raw training data |
| Position Requirements | FOH and BOH positions with min/max/target hours (editable) |
| Master Dashboard | Live overview: active trainees, hours, progress, cert status |
| Training Schedule | Planned day-by-day assignments generated from timelines |
| Training Needs | Weekly view auto-populated for managers (who trains where today) |
| Certification Log | Archive of certified team members with total hours and duration |
| Name Deduplication | Catches and merges duplicate/variant name entries |
| Alert Settings | Configure which alerts fire and who receives them |

**Default position list (customize to your location):**

| FOH Positions | Target Hours | BOH Positions | Target Hours |
|---------------|-------------|---------------|-------------|
| iPOS | 14 | Breading | 14 |
| Register/POS | 10 | Raw (Filet) | 11 |
| Cash Cart | 4 | Fries | 9 |
| Server | 7.5 | Machines | 9 |
| FC Drinks | 7 | Dishes | 7 |
| Desserts | 5 | Prep | 9 |
| DT Drinks | 7 | Secondary | 9 |
| DT Stuffer | 9 | Primary (Buns) | 14 |
| FC Bagger | 5 | Truck | 5 |
| DT Bagger | 5 | | |
| Window | 5 | | |

> These are defaults based on Cockrell Hill's setup. **Edit the "Position Requirements" sheet** to match your location's positions and hour targets.

### Component 2: Position Progress Google Form (Daily Training Log)

A Google Form that trainers or managers fill out after each training session. This is how data gets into the system.

**What the form collects:**
| Field | Description |
|-------|-------------|
| Date | Date of training session |
| Trainee Name | Full name of the team member being trained |
| Position Trained | Which position they worked (dropdown recommended) |
| Hours | Number of hours trained |
| On Track? | Yes/No — is the trainee progressing as expected? |
| Notes | Optional freeform notes |

**How it connects:** The form is linked to the "Daily Training Log" sheet in your Training Spreadsheet. When someone submits the form, the response lands in that sheet, and the Apps Script automatically:
1. Resolves the trainee's canonical name (catches duplicates/variants)
2. Checks if they just completed a position milestone
3. Checks if they're now certification-ready
4. Refreshes the Master Dashboard
5. Sends any triggered email alerts

### Component 3: FOH and BOH Final Feedback Forms

Two standalone Google Forms — one for Front of House, one for Back of House — used as the **final sign-off** at the end of a trainee's training.

**Purpose:** These are filled out by a certified trainer (not the trainee) as the last checkpoint before a team member is considered fully trained. They serve as a formal record that a qualified trainer reviewed and approved the trainee's readiness.

**When they're used:** After the Master Dashboard shows a trainee as "READY" (all position minimums met), a certified trainer completes the appropriate feedback form to officially certify them.

**Important:** These forms are **independent** — they are not connected to the Training Spreadsheet. They exist as standalone records in Google Forms. However, the Apps Script can optionally open the correct form (FOH or BOH) with the trainee's name pre-filled when you certify someone from the spreadsheet.

> If you don't want to use separate certification forms, you can use the **"Manually Certify Trainee"** option in the Training Tools menu instead.

---

## How All Three Components Work Together

```
                    DAILY OPERATIONS
                    ================

  Trainer fills out              Training Spreadsheet
  Google Form after    ----->    automatically updates:
  each training                  - Daily Training Log
  session                        - Master Dashboard
                                 - Position progress
                                 - Email alerts

                    WEEKLY PLANNING
                    ===============

  Manager generates              System auto-populates
  training timeline    ----->    Training Needs sheet
  for new hire                   each Monday at 5 AM
                                 (or manually via menu)

                    CERTIFICATION
                    =============

  Dashboard shows               Certified trainer fills
  trainee is         ----->     out FOH or BOH Feedback
  "READY"                       Form as final sign-off
                                (or use manual certification)
```

---

## Setup Instructions (Quick Start)

> For the full technical walkthrough, see [SETUP_GUIDE.md](SETUP_GUIDE.md).

### Step 1: Make Your Copies

1. **Training Spreadsheet:** Make a copy of the master Training Spreadsheet (or create a new Google Sheet)
2. **Daily Training Log Form:** Make a copy of the training log Google Form (or create your own — see form fields above)
3. **Certification Forms (optional):** Make copies of the FOH and BOH feedback forms

### Step 2: Install the Apps Script Code

1. Open your Training Spreadsheet
2. Go to **Extensions > Apps Script**
3. Create the script files listed in [SETUP_GUIDE.md](SETUP_GUIDE.md) and paste in the code
4. Save all files

### Step 3: Run Initial Setup

1. Go back to your spreadsheet and reload the page
2. Click **Training Tools > Initial Setup (Run Once)**
3. Authorize the script when prompted by Google
4. All sheets are created and populated with defaults

### Step 4: Link Your Form

1. Open your training log Google Form
2. Go to **Responses > Link to Sheets**
3. Select your Training Spreadsheet
4. Choose the **Daily Training Log** sheet as the destination

### Step 5: Set Up Triggers

In the Apps Script editor, create these triggers (clock icon in the left sidebar):

| Trigger | Function | Event | Timing |
|---------|----------|-------|--------|
| Form submission | `onFormSubmit` | From spreadsheet > On form submit | — |
| Daily reminder | `dailyReminderCheck` | Time-driven > Day timer | 8-9 PM |
| Inactive check | `checkInactiveTrainees` | Time-driven > Day timer | 6-7 AM |
| Duplicate scan | `checkForDuplicates` | Time-driven > Day timer | 7-8 AM |

### Step 6: Configure Alerts

Click **Training Tools > Alert Settings** and enter email addresses for the managers who should receive alerts.

### Step 7: Customize Positions

Edit the **Position Requirements** sheet to match your location's positions and hour targets.

---

## Adapting This to Your Location

These are the things you'll want to change:

| What to Change | Where | Why |
|---------------|-------|-----|
| Position names and hours | "Position Requirements" sheet | Your location may have different positions or different hour targets |
| Alert recipients | Training Tools > Alert Settings | Enter your managers' email addresses |
| Daypart time windows | `04_Timeline.gs` — `getDaypartsForShift()` function | If your dayparts differ from Breakfast 6-11, Lunch 11-15, Afternoon 15-17, Dinner 17-22 |
| Certification form URLs | `05_Certification.gs` — `openCertificationForm()` function | Replace with your own form URLs and entry IDs |
| Position name aliases | `03_Dashboard.gs` — `normalizePositionName()` function | Add mappings if your team uses different shorthand for positions |
| Nickname mappings | `02_Form_and_Dedup.gs` — `areSimilar()` function | Add common nicknames specific to your team |
| Efficiency factor | `04_Timeline.gs` — line with `0.85` | The 85% factor accounts for breaks/setup; adjust if needed |

---

## Using the System Day-to-Day

### The Training Tools Menu

When you open the spreadsheet, a **Training Tools** menu appears with these options:

| Menu Item | What It Does |
|-----------|-------------|
| Generate Training Timeline | Opens a dialog to create a week-by-week plan for a new hire |
| Load This Week's Training | Populates the Training Needs sheet with this week's assignments |
| Load Next Week's Training | Same, but for next week |
| Review Duplicate Names | Opens a sidebar to merge or ignore duplicate name suggestions |
| Check for Duplicates Now | Runs the duplicate name scan immediately |
| Alert Settings | Opens a dialog to configure email alerts |
| Refresh Dashboard | Manually recalculates the Master Dashboard |
| Sync Form Data | Imports any unsynced form responses into the Daily Training Log |
| Manually Certify Trainee | Certifies a trainee without using the feedback form |
| Setup Monday Auto-Populate | Creates a trigger that auto-loads Training Needs every Monday at 5 AM |
| Initial Setup (Run Once) | Creates all required sheets (safe to re-run) |

### Generating a Timeline

1. Click **Training Tools > Generate Training Timeline**
2. Enter the trainee's name
3. Select their house (FOH or BOH)
4. Select their primary shift (Breakfast, Lunch, Afternoon, or Dinner)
5. Enter their hours per shift
6. Pick a start date (defaults to next Monday)
7. Click the days they'll be working across the 3-week calendar
8. Click **Generate Timeline**

The system creates a new "Timeline - [Name]" sheet with a day-by-day plan, writes the schedule to the Training Schedule sheet, and auto-populates Training Needs for the first week.

**Multi-daypart coverage:** If a trainee works an 8-hour shift starting at Breakfast (6 AM), they'll be scheduled through Lunch as well. The system automatically calculates which dayparts a shift covers based on the hours.

### Reading the Dashboard

The **Master Dashboard** shows:
- **Quick Stats:** Active trainees, weekly hours, cert-ready count
- **Active Trainees table:** Name, house, total hours, days since last training, last position, on-track %, certification status
- **Position Progress matrix:** Color-coded grid showing hours logged vs. required for every position
  - Green = completed (met minimum hours)
  - Yellow = in progress (some hours logged)
  - Gray = not started

### Certifying a Trainee

When the dashboard shows a trainee as **"READY"**:
1. Have a certified trainer complete the FOH or BOH Feedback Form, **or**
2. Use **Training Tools > Manually Certify Trainee**

Either way, the trainee moves to the Certification Log and is removed from the active dashboard.

---

## FAQ

**Q: Do I need to know how to code to use this?**
A: No. The initial setup requires copying and pasting code into the Apps Script editor, but after that everything runs through the spreadsheet menu and Google Forms. See the [SETUP_GUIDE.md](SETUP_GUIDE.md) for step-by-step instructions.

**Q: Can multiple managers use this at the same time?**
A: Yes. It's a Google Sheet — multiple people can have it open and the form can receive submissions simultaneously. The dashboard refreshes automatically after each form submission.

**Q: What if someone's name is entered differently on different days?**
A: The system has built-in duplicate detection. It runs daily and catches typos, nicknames (Tim/Timothy), and similar names using fuzzy matching. You review suggestions in the sidebar and click Merge or Ignore.

**Q: Can I add or remove positions?**
A: Yes. Edit the "Position Requirements" sheet directly. Add rows for new positions or delete rows for ones you don't use. The dashboard and timeline will pick up the changes automatically.

**Q: What if a trainee works multiple dayparts?**
A: The system handles this. If you set a trainee's shift to "Breakfast" and their hours per shift to 8, the system calculates that they'll be present from 6 AM to 2 PM and populates them in both the Breakfast and Lunch sections of Training Needs.

**Q: Can I use this for BOH too?**
A: Yes. The system supports both FOH and BOH. Timelines and dashboards work for both houses. Note: the weekly Training Needs auto-population currently writes FOH trainees only (BOH scheduling may be handled differently at your location). [TODO: confirm if BOH Training Needs population is needed for your setup]

**Q: What happens if I run Initial Setup more than once?**
A: It's safe to re-run. It skips any sheets that already exist and won't overwrite your data.

**Q: How do I stop getting email alerts?**
A: Go to **Training Tools > Alert Settings** and uncheck the alerts you don't want, or remove your email address from the recipient fields.

**Q: The Training Tools menu doesn't appear when I open the sheet.**
A: Reload the page. The menu is created by the `onOpen()` function which runs each time the spreadsheet loads. If it still doesn't appear, open **Extensions > Apps Script** and make sure all the script files are saved without errors.

---

## File Inventory

```
Scripts (8 files):
  01_Menu_and_Setup.gs      Menu creation + one-time sheet setup
  02_Form_and_Dedup.gs      Form submission handler + name deduplication
  03_Dashboard.gs           Dashboard calculations + position progress tracking
  04_Timeline.gs            Timeline generator + Training Needs auto-population
  05_Certification.gs       Certification workflow + archiving
  06_Alerts.gs              Email alerts + daily trigger functions
  07_UI_Functions.gs        Sidebar/dialog launchers + data getters
  08_DataSync.gs            Form response import + canonical name backfill

HTML (3 files):
  TimelineDialog.html       Timeline generation dialog UI
  DeduplicationSidebar.html Duplicate name review sidebar UI
  AlertSettingsDialog.html  Alert configuration dialog UI

Documentation:
  README.md                 This file
  SETUP_GUIDE.md            Detailed technical installation guide

Diagnostic (optional):
  DIAGNOSTIC.gs             Debug utility for troubleshooting layout detection
                            (can be removed after setup is confirmed working)
```

---

## License

This project was built for the Chick-fil-A operator community. Use it, adapt it, share it. If you make improvements, consider sharing them back so other operators can benefit.
