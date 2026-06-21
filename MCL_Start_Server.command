#!/bin/bash
# Mwananchi Credit — Start Portfolio API Server
# Double-click to run. Right-click > Open on first launch.

REPO_DIR="$HOME/mwananchi-portfolio"
VENV="$REPO_DIR/.venv"
APP="$REPO_DIR/app.py"
CREDS="$HOME/.mcl_creds"
PORT=5050

step() { echo ""; echo "[ $1 ] $2"; }
ok()   { echo "       OK: $*"; }
fail() { echo ""; echo "  FAILED: $*"; echo ""; read -p "Press Enter to close..."; exit 1; }

clear
echo ""
echo "========================================================"
echo "  Mwananchi Credit Ltd"
echo "  Portfolio API Server"
echo "========================================================"

# ── STEP 1: Repo ─────────────────────────────────────────────
step 1 "Repository..."
[ -d "$REPO_DIR/.git" ] || fail "Repo not found. Run MCL_Excel_Refresh.command first."
git -C "$REPO_DIR" pull --quiet 2>/dev/null && ok "Up to date" || ok "Offline — using local copy"

# ── STEP 2: Python venv ──────────────────────────────────────
step 2 "Python environment..."
if [ -d "$VENV" ]; then
  "$VENV/bin/python3" -c "import pymssql, flask, flask_cors" 2>/dev/null \
    || { echo "       Rebuilding venv..."; rm -rf "$VENV"; }
fi
if [ ! -d "$VENV" ]; then
  python3 -m venv "$VENV" || fail "python3 -m venv failed"
  source "$VENV/bin/activate"
  pip install --upgrade pip setuptools wheel --quiet 2>/dev/null
  pip install pymssql --prefer-binary --quiet 2>/dev/null || fail "pymssql install failed"
  pip install flask flask-cors openpyxl --quiet 2>/dev/null  || fail "flask install failed"
fi
ok "Environment ready"

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
  ok "Password saved"
else
  ok "Using saved password"
fi
export SQL_PASSWORD

# ── STEP 4: Cloudflared tunnel (optional, for Monday.com) ────
step 4 "Tunnel for Monday.com embed..."
TUNNEL_URL=""
if command -v cloudflared &>/dev/null; then
  echo "       Starting cloudflared tunnel..."
  cloudflared tunnel --url "http://localhost:$PORT" --no-autoupdate > /tmp/mcl_tunnel.log 2>&1 &
  TUNNEL_PID=$!
  sleep 4
  TUNNEL_URL=$(grep -o 'https://[a-zA-Z0-9-]*\.trycloudflare\.com' /tmp/mcl_tunnel.log | head -1)
  if [ -n "$TUNNEL_URL" ]; then
    ok "Tunnel URL: $TUNNEL_URL"
  else
    echo "       Tunnel not available — server will run on localhost only"
  fi
else
  echo "       cloudflared not installed — running on localhost only"
  echo "       (Install cloudflared to get a public URL for Monday.com)"
fi

# ── STEP 5: Start Flask ──────────────────────────────────────
step 5 "Starting server on port $PORT..."
echo ""
echo "========================================================"
echo "  Server started"
echo ""
echo "  Local URL:    http://localhost:$PORT"
if [ -n "$TUNNEL_URL" ]; then
echo ""
echo "  Public URL:   $TUNNEL_URL"
echo ""
echo "  FOR MONDAY.COM EMBED:"
echo "    1. Open mcl_dashboard.html in a text editor"
echo "    2. Change this line near the top of the <script>:"
echo "         const API_BASE = 'http://localhost:$PORT';"
echo "       to:"
echo "         const API_BASE = '$TUNNEL_URL';"
echo "    3. Save and re-upload / refresh the embed"
fi
echo ""
echo "  Keep this window open while using the dashboard."
echo "  Press Ctrl+C to stop the server."
echo "========================================================"
echo ""

cd "$REPO_DIR"
"$VENV/bin/python3" "$APP"
