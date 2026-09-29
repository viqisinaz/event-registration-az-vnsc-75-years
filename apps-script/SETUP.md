# Registration backend setup

Registrations are saved to a **private Google Sheet** owned by
viqis.in.az@gmail.com. The sheet never lives in this repo. Only the web app
URL goes into `index.html`, and that URL can only add rows, not read them.

For every submission the backend:

- stores a row with the ID `vnsc_75_az_<email>`. If the same email registers again, that row is updated, not duplicated.
- recalculates the total from the category counts ($150 / $75 / $50).
- emails a confirmation to the registrant, with viqis.in.az@gmail.com on CC.

## One-time setup (about 10 minutes)

1. Sign in to Google as **viqis.in.az@gmail.com**.
2. Create a new Google Sheet, for example "VIQI 75 AZ Registrations".
3. In the sheet, open **Extensions → Apps Script**.
4. Delete the sample code, paste in all of `apps-script/Code.gs`, and click **Save**.
5. Choose `setup` in the function dropdown and click **Run**. Google will ask for
   permission to edit the sheet and send email as you. Click **Allow**. (If you see
   "Google hasn't verified this app", click **Advanced → Go to … (unsafe)**; it's
   your own script.) A **Registrations** tab with headers now appears in the sheet.
6. Click **Deploy → New deployment**. Pick the type **Web app** and set:
   - **Execute as:** Me (viqis.in.az@gmail.com)
   - **Who has access:** Anyone
7. Click **Deploy** and copy the **Web app URL** (it ends in `/exec`).
8. In `index.html`, replace `PASTE_APPS_SCRIPT_WEB_APP_URL_HERE` with that URL,
   then commit and push.

## Getting the spreadsheet as Excel

In the Google Sheet, use **File → Download → Microsoft Excel (.xlsx)**. Don't
commit the downloaded file. `.gitignore` already blocks `.xlsx` and `.csv`
files as a safeguard.

## Changing the script later

After editing `Code.gs` in the Apps Script editor, use **Deploy → Manage
deployments → Edit (pencil) → Version: New version → Deploy**. This keeps the
same URL, so `index.html` doesn't need to change.

If you change the prices, update both `index.html` and `CATEGORIES` in
`Code.gs`.

## Limits

A regular Gmail account can send email to about **100 recipients a day**.
Each registration uses 2 (the registrant plus the CC), so that's about 50
registrations a day. Registrations over the limit are still saved; only the
email is skipped for that day.
