# Hosting the Dashboard on GitHub Pages (Free)

## Why host?
- Google Sheets live fetch works (no CORS errors)
- Anyone with the URL can access the dashboard
- Free, secure (HTTPS), always online
- No server maintenance needed

## Steps (10 minutes, one-time setup)

### 1. Create GitHub account
- Go to https://github.com
- Sign up (free)

### 2. Create a repository
- Click "+" → "New repository"
- Name: `mis-dashboard`
- Set to **Public** (required for free GitHub Pages)
- Click "Create repository"

### 3. Upload the dashboard file
- In the repository, click "Add file" → "Upload files"
- Upload `index.html` from `C:\Users\MI\Downloads\MIS\`
- Click "Commit changes"

### 4. Enable GitHub Pages
- Go to repository Settings → Pages
- Source: "Deploy from a branch"
- Branch: `main` → folder: `/ (root)`
- Click Save
- Wait 1-2 minutes

### 5. Access your dashboard
- Your URL will be: `https://YOUR-USERNAME.github.io/mis-dashboard/`
- Share this URL with your team
- Everyone opens it → clicks "🔄 Load from Google Sheets (live)" → data loads!

## How it works after setup
1. Team edits data in Google Sheets (as usual)
2. Anyone opens the dashboard URL
3. Clicks "🔄 Load from Google Sheets (live)" → fetches latest data
4. Each person customizes their own filters/charts
5. No upload needed — always live from Google Sheets

## Updating the dashboard
If I make changes to the dashboard HTML:
1. Go to your GitHub repository
2. Click on `index.html`
3. Click the pencil icon (edit) or "Add file" → "Upload files"
4. Upload the new `index.html`
5. Changes go live in 1-2 minutes

## Security
- The dashboard HTML is public (anyone with URL can see it)
- But the DATA requires the Google Sheets to be published
- Google Sheets publish is read-only — no one can edit via the published link
- Editing still requires Google Sheets access (shared with your team only)

## Daily Backups (separate)
See `DailyBackup-GoogleSheets.js` — paste into each Google Sheet's Apps Script.
Creates daily Excel backups in your Drive's "MIS_Backups" folder.
