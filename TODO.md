# GitHub Connected - Final Fix

## Push Error: Git not in PowerShell PATH (PS prompt)

**Solution: Switch Terminal**
1. Terminal panel (+ New Terminal)
2. Click dropdown → **Git Bash**
3. Prompt changes to `bgora@BGORA MINGW64 ~/Downloads/BarnData`

**Run:**
```
git status
git add .
git commit -m "Initial BarnData commit"
git push -u origin main
```

**Success:** 
- `git status` shows clean
- Repo https://github.com/Gora-Bhargavi/BarnData shows BarnData.sln etc.

**gh bonus:** `gh pr create` for future PRs.

Connected!





