@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

REM ===================================================================
REM  WorkBuddy Version Snapshot - one click capture
REM  * read-only collection (never modifies WorkBuddy)
REM  * auto commit into THIS folder's own git repo
REM  * does NOT touch the outer Codex repo
REM ===================================================================

cd /d "%~dp0.."
set REPO=%CD%

echo ==================================================
echo   WorkBuddy Version Snapshot
echo   Repo: %REPO%
echo ==================================================
echo.

REM ---- locate a usable python ----
set PY=
where python >nul 2>&1 && set PY=python
if not defined PY if exist "%USERPROFILE%\.workbuddy\binaries\python\versions\3.13.12\python.exe" set PY=%USERPROFILE%\.workbuddy\binaries\python\versions\3.13.12\python.exe
if not defined PY if exist "D:\04 Software\00 python\python.exe" set PY=D:\04 Software\00 python\python.exe

if not defined PY (
  echo [ERROR] Python not found.
  echo         Run manually:  python tools\snapshot.py
  echo.
  pause
  exit /b 1
)

echo [python] %PY%
echo.

REM ---- run snapshot ----
"%PY%" "tools\snapshot.py" %*
set RC=%ERRORLEVEL%
echo.

REM ---- commit only when something changed ----
git rev-parse --git-dir >nul 2>&1
if errorlevel 1 goto :nogit

set STATUSFILE=%TEMP%\wbv_git_status.txt
git status --porcelain > "%STATUSFILE%" 2>nul
for /f "usebackq delims=" %%L in ("%STATUSFILE%") do (
  set HASCHG=1
  goto :docommit
)
echo [git] No changes, nothing to commit.
goto :end

:docommit
echo [git] Changes detected, committing...
git add README.md TIMELINE.md records tools .gitignore >nul 2>&1
git commit -m "snapshot: WorkBuddy version record" >nul 2>&1
if errorlevel 1 (
  echo [git] Commit failed. Check git user.name / user.email.
) else (
  echo [git] Committed.
)
goto :end

:nogit
echo [git] Not a git repository, skip commit.

:end
if defined STATUSFILE del "%STATUSFILE%" >nul 2>&1
echo.
echo ==================================================
echo   Done. See TIMELINE.md
echo ==================================================
echo.
pause
exit /b %RC%
