@echo off
git add -A
git -c user.email=agent@local -c user.name=Kiro commit -m "feat: SharePoint live (via Apps Script proxy) — same UX as Google Sheets endpoint"
git push
git tag -a v2026.05.17.14 -m "SharePoint live via Apps Script (CORS-free, searchable picker)"
git push origin v2026.05.17.14
