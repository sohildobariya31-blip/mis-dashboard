@echo off
git add -A
git -c user.email=agent@local -c user.name=Kiro commit -m "feat: hardcoded Apps Script URLs, searchable picker, full Solar/Frame/Glass calc columns, info-chip header"
git push
git tag -a v2026.05.17.13 -m "Apps Script defaults + special-item calculations + chip-style header"
git push origin v2026.05.17.13
