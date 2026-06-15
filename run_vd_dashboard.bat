@echo off
REM Auto-refresh the VD Rebate dashboard for the current month, then deploy if changed.
cd /d "C:\Users\Administrator\olympic-paints-vd"
"C:\Python313\python.exe" build_vd_dashboard.py >> "C:\Users\Administrator\olympic-paints-vd\build.log" 2>&1
git add dashboard/index.html
git diff --cached --quiet
if errorlevel 1 (
  git commit -m "chore(vd): auto-refresh dashboard data for current month"
  git push
)
