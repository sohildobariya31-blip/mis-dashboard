@echo off
git add -A
git -c user.email=agent@local -c user.name=Kiro commit -m "feat: live edit reflection + dropdown filter overhaul + refresh cooldown"
git push
