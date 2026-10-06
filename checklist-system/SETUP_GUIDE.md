# Checklist System – Setup Guide

Do these in order. Steps marked **(store account)** must be done while signed in as the store Gmail account. Triggers and forms belong to whoever creates them.

## 1. Transfer the master sheet (store account)
1. From Brent's account, open **Checklist Master (FOH + BOH)**, then Share, and make the store Gmail account the owner. Between personal Gmail accounts, the store account has to accept the transfer.
2. Keep Brent as an **Editor** so code can still be pushed with clasp.

## 2. Create the script (store account)
1. In the sheet: **Extensions > Apps Script**. Name the project `Checklist System`.
2. In the editor: **Project Settings (gear)**, then copy the **Script ID**.

## 3. Push the code (Brent's Mac, one-time clasp setup)
```bash
npm install -g @google/clasp
```
Turn on the Apps Script API at https://script.google.com/home/usersettings, then:
```bash
clasp login
```
In this `checklist-system` folder, create `.clasp.json` with the Script ID from step 2:
```json
{ "scriptId": "PASTE_SCRIPT_ID_HERE", "rootDir": "." }
```
Then push. This replaces the default `Code.gs` with these files:
```bash
clasp push --force
```

## 4. First run (store account)
1. Reload the sheet. A **Checklists** menu appears.
2. **Checklists > Validate sheet.** Fix anything listed as blocking. Blank emails are only warnings.
3. **Checklists > Rebuild forms now.** Approve the permissions prompt. This creates the three forms and fills in Form ID / Form link.
4. **Checklists > Install / repair triggers.** The dialog should say triggers will run as the store account.
5. **Checklists > Show form links.** Open each link in a private/incognito window. You should get the form without signing in. If it asks you to sign in or request access, open the form, click **Publish**, and set Responders to **Anyone with the link**.
6. Fill in **Settings**: FOH escalation email and Daily summary recipients.

## 5. Test
- Submit one response per position on Positional Closing, plus Manager Closing and the 2–3 PM checklist.
- Check **Submissions** (one row each), **Item Results** (one row per task), and **Daily Status** (rows flip to Complete).
- Confirm that submitting **Stocker** ends the form. Team Leader should never appear after Stocker.
- **Checklists > Send test alert to me** to see what a late email looks like.

## 6. Go live
1. Run one week with Alerts mode = `Digest only`. Only the 7 AM summary goes out. Tune Due by / Late after from what actually happens.
2. Set Alerts mode to `Live`. Heads-ups and late escalations start on the next 15-minute check.
3. Replace the old links and QR codes with the ones from *Show form links*. The links never change.
4. On each of the eight old forms, turn off **Accepting responses** and set the closed message to point to the new link. **Do not delete** the old forms or their response sheets.

## 7. Photo uploads (store account, one time)
The photo page lets team members upload pictures from their phone without signing in.

1. Signed in as the store account, open the sheet and go to **Extensions > Apps Script**.
2. Click **Deploy > New deployment**, then the gear icon, then **Web app**.
   - Description: `Photo upload`
   - Execute as: **Me** (the store account)
   - Who has access: **Anyone**
3. Click **Deploy** and approve the permissions. The script saves photos to the store account's Google Drive.
4. Copy the **Web app URL** (it ends in `/exec`) into **Settings > Photo upload page**.
5. Run **Checklists > Rebuild forms now**. Every form now shows the upload link at the top and again on the screen after submitting.
6. In **Checklists > Show form links**, copy the **Photo upload page** link and make a QR code from it. In Chrome: open the link, click Share, then Create QR code. Post it at each station.
7. Test on a phone that isn't signed in to Google. Photos start uploading as soon as they're picked (3 at a time, shrunk on the phone first), and people can keep adding more. The page only has to stay open until it shows ✓. Check that photos show up in Drive under **Checklist Photos** and on the **Photos** tab.
8. After a week on photos, clear **Settings > Photo email** so the forms stop mentioning email.
9. **Let managers view photos:** after the first upload, the store account opens Google Drive, right-clicks **Checklist Photos**, chooses **Share**, and adds each manager as a **Viewer**. New photos inherit the sharing. Managers then use **Checklists > View photos**. The first time, it asks them to approve the script's permissions.

**If the upload page's code ever changes:** as the store account, go to **Deploy > Manage deployments**, click the pencil, set Version to **New version**, and click **Deploy**. The link stays the same.

**To retire an old link or QR:** change **Settings > Photo upload key**, then make a new QR.

## If something goes wrong
- **Forms didn't rebuild:** the FOH escalation email gets a "NOT rebuilt" email listing what to fix. Fix the sheet; the next 15-minute check rebuilds on its own.
- **A script error:** one email per failing step per day goes to the FOH escalation email (or the store account if that's blank). Details are in the Alert Log.
- **"Authorization required" after a code update:** a new feature needs a new Google permission. As the store account, open the sheet and run **Checklists > Validate sheet** once, then approve.
- **Someone else installed triggers before:** they should run *Checklists > Remove my triggers* from their account. Until they do, their triggers do nothing.
