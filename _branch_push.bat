@echo off
git checkout -b apps-script-live
git add -A
git -c user.email=agent@local -c user.name=Kiro commit -m "feat(test): Apps Script live JSON endpoint for instant refresh"
git push -u origin apps-script-live
git checkout main
