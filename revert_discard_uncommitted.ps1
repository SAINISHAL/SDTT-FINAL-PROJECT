# Discard all uncommitted local changes in the repo (safe if you do NOT need today's edits).
# Run from the repo root: PowerShell -ExecutionPolicy Bypass -File .\revert_discard_uncommitted.ps1
git status --porcelain
Write-Host "`nAbout to discard ALL local uncommitted changes and untracked files. Press Enter to continue or Ctrl+C to abort."
Read-Host
git stash save "backup-before-discard-$(Get-Date -Format s)"
git checkout -- .
git clean -fd
Write-Host "Done. All uncommitted changes discarded (a stash backup was created)."
