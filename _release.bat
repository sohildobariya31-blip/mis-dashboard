@echo off
git add -A
git -c user.email=agent@local -c user.name=Kiro commit -m "feat: 60s refresh cooldown for SharePoint Apps Script (with countdown)"
git push
git tag -a v2026.05.17.15 -m "SharePoint refresh cooldown"
git push origin v2026.05.17.15
git rm _release.bat
git -c user.email=agent@local -c user.name=Kiro commit -m "chore: remove temp script"
git push
