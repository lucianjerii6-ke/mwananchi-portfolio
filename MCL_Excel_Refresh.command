#!/bin/bash
# Mwananchi Credit — Portfolio Excel Refresh
# Double-click this file (or right-click > Open on first run)

REPO_DIR="$HOME/mwananchi-portfolio"
VENV="$REPO_DIR/.venv"
SCRIPT="$REPO_DIR/mcl_refresh.py"
CREDS="$HOME/.mcl_creds"

step() { echo ""; echo "[ $1 ] $2"; }
ok()   { echo "       OK: $*"; }
fail() { echo ""; echo "  FAILED: $*"; echo ""; read -p "Press Enter to close..."; exit 1; }

clear
echo ""
echo "========================================================"
echo "  Mwananchi Credit Ltd"
echo "  Portfolio Excel Refresh"
echo "========================================================"

# ── STEP 1: Repo ─────────────────────────────────────────────
step 1 "Repository..."
if [ -d "$REPO_DIR/.git" ]; then
  git -C "$REPO_DIR" pull --quiet 2>/dev/null && ok "Up to date" || ok "Offline — using local copy"
else
  git clone https://github.com/lucianjerii6-ke/mwananchi-portfolio.git "$REPO_DIR" \
    || fail "git clone failed — check internet"
  ok "Cloned"
fi

# ── STEP 2: Python venv + deps ───────────────────────────────
step 2 "Python dependencies..."
cd "$REPO_DIR"

if [ -d "$VENV" ]; then
  "$VENV/bin/python3" -c "import pymssql, openpyxl" 2>/dev/null \
    || { echo "       Rebuilding venv..."; rm -rf "$VENV"; }
fi

if [ ! -d "$VENV" ]; then
  python3 -m venv "$VENV" || fail "python3 -m venv failed"
fi

source "$VENV/bin/activate"
pip install --upgrade pip setuptools wheel --quiet 2>/dev/null
pip install pymssql --prefer-binary --quiet 2>/dev/null \
  || fail "pymssql install failed"
pip install openpyxl --quiet 2>/dev/null \
  || fail "openpyxl install failed"
ok "All packages ready"

# ── STEP 3: SQL credentials ──────────────────────────────────
step 3 "SQL credentials..."
export SQL_SERVER="173.249.59.109"
export SQL_PORT="1433"
export SQL_DATABASE="mwananchidy365"
export SQL_USER="monday"

[ -f "$CREDS" ] && source "$CREDS" 2>/dev/null

if [ -z "$SQL_PASSWORD" ]; then
  read -s -p "       Enter SQL password (saved for future runs): " SQL_PASSWORD
  echo ""
  echo "SQL_PASSWORD='$SQL_PASSWORD'" > "$CREDS"
  chmod 600 "$CREDS"
  ok "Password saved to $CREDS"
else
  ok "Using saved password"
fi
export SQL_PASSWORD

# ── STEP 4: Date selection ───────────────────────────────────
step 4 "As at Date..."
DEFAULT_DATE=$(date +"%Y-%m-%d")
echo ""
echo "       Enter the date you want to report on."
echo "       Format: YYYY-MM-DD   (e.g. 2026-06-20)"
echo "       Press Enter to use today: $DEFAULT_DATE"
echo ""
read -p "       As at Date [${DEFAULT_DATE}]: " INPUT_DATE
echo ""

# Use default if nothing entered
AS_AT_DATE="${INPUT_DATE:-$DEFAULT_DATE}"

# Validate format
if ! [[ "$AS_AT_DATE" =~ ^[0-9]{4}-[0-9]{2}-[0-9]{2}$ ]]; then
  echo "       Invalid date format. Using today: $DEFAULT_DATE"
  AS_AT_DATE="$DEFAULT_DATE"
fi
ok "Reporting as at: $AS_AT_DATE"

# ── STEP 4b: Branch filter ────────────────────────────────────
echo ""
echo "       Enter a branch name to filter (or press Enter for ALL branches)."
echo "       Examples: KISUMU, ELDORET, MOMBASA, NAKURU, THIKA"
echo ""
read -p "       Branch [ALL]: " INPUT_BRANCH
echo ""

BRANCH="${INPUT_BRANCH:-ALL}"
BRANCH_UPPER=$(echo "$BRANCH" | tr '[:lower:]' '[:upper:]')
ok "Branch filter: $BRANCH_UPPER"

# ── STEP 5: Run refresh ───────────────────────────────────────
step 5 "Running SP queries and building Excel..."
echo ""
"$VENV/bin/python3" "$SCRIPT" "$AS_AT_DATE" "$BRANCH_UPPER"
EXIT_CODE=$?

if [ $EXIT_CODE -ne 0 ]; then
  echo ""
  echo "  The script encountered an error."
  echo "  Common causes:"
  echo "    - Wrong SQL password (delete ~/.mcl_creds and re-run)"
  echo "    - SQL Server unreachable (check VPN / network)"
  echo "    - Stored procedure name changed"
  read -p "  Press Enter to close..."
  exit 1
fi

# ── STEP 6: Open Excel ────────────────────────────────────────
if [ "$BRANCH_UPPER" = "ALL" ]; then
  EXCEL_FILE="$REPO_DIR/MCL_Portfolio_${AS_AT_DATE}.xlsx"
else
  EXCEL_FILE="$REPO_DIR/MCL_Portfolio_${AS_AT_DATE}_${BRANCH_UPPER}.xlsx"
fi
if [ -f "$EXCEL_FILE" ]; then
  echo "  Opening Excel..."
  open "$EXCEL_FILE"
fi

echo ""
echo "========================================================"
echo "  DONE — Portfolio Excel is ready"
echo "  Date:  $AS_AT_DATE"
echo "  File:  $EXCEL_FILE"
echo ""
echo "  To get Claude's analysis:"
echo "    1. Upload the Excel file to Claude in Cowork"
echo "    2. Ask: Analyse this portfolio and write a report"
echo "========================================================"
echo ""
echo "  Press Enter to close this window."
read
