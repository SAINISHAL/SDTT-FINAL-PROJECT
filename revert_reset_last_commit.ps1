# WARNING: This will remove the last commit (destructive). Only use if today's changes were already committed.
# Run from the repo root: PowerShell -ExecutionPolicy Bypass -File .\revert_reset_last_commit.ps1
git log --oneline -n 5
Write-Host "`nAbout to run: git reset --hard HEAD~1  (this will drop the latest commit). Press Enter to continue or Ctrl+C to abort."
Read-Host
git reflog -n 5
git reset --hard HEAD~1
git clean -fd
Write-Host "Done. Repository reset to previous commit. Use git reflog if you need to recover."
