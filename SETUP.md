# Reunion form setup (about 10 minutes)

## 1. Create the Google Sheet + backend
1. Go to sheets.google.com and create a blank sheet, e.g. "School Reunion Responses".
2. Menu **Extensions → Apps Script**. Delete the sample code, paste everything from `Code.gs`, click **Save**.
3. In the function dropdown pick `setup` and click **Run**. Approve the permission prompt (Advanced → Go to project → Allow). A "Responses" tab with headers appears in the sheet.
4. Click **Deploy → New deployment** → gear icon → **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy** and copy the **Web app URL** (ends with `/exec`).

## 2. Connect the page
Open `index.html` and, at the top of the `<script>`, paste the URL:

```js
SCRIPT_URL: "https://script.google.com/macros/s/XXXX/exec",
```

Edit the other `CONFIG` lines too: school name, batch, title, date/place options, budget ranges.

## 3. Put it online (free)
Pick one:
- **Netlify Drop**: open app.netlify.com/drop and drag the `reunion-form` folder in. You get a link instantly.
- **GitHub Pages**: new public repo → upload `index.html` → Settings → Pages → deploy from `main`.
- Or put it on any of your own servers / S3 bucket.

Share that link in the WhatsApp/Facebook group.

## Notes
- Answers land in the "Responses" tab; photos go into a Drive folder called "Reunion Photos" (shared view-only by link so they show on the page).
- Photos are shrunk to 600px in the browser before upload, so they stay small.
- Use a personal Gmail account for the script. Some work/school Google accounts block "anyone with the link" sharing, and then photos won't show.
- Submitting again from the same phone/browser updates that person's entry instead of adding a duplicate.
- If you change `Code.gs` later: **Deploy → Manage deployments → Edit → Version: New version → Deploy** (the URL stays the same).
- Use the sheet's **Data → Create a filter** or a pivot table to count people per city or per preferred date.
