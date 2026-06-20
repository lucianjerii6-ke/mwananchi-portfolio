"""
Mwananchi Credit — Portfolio Dashboard
Monday.com App Backend + Static File Server

Deploy to Render.com:
  1. Push this folder to GitHub
  2. Connect repo on render.com → New Web Service
  3. Set environment variables (see render.yaml)
  4. Deploy — get your public URL (e.g. https://mwananchi-portfolio.onrender.com)
  5. Use that URL in the Monday.com app feature settings

Local development:
  python app.py   (runs on http://localhost:5050)
"""

import os
from flask import Flask, jsonify, request, render_template, send_from_directory
from flask_cors import CORS
import pymssql
from datetime import datetime

app = Flask(__name__, template_folder="templates", static_folder="static")
CORS(app)

# ── Credentials — read from environment variables (set in Render dashboard)
SQL_SERVER   = os.environ.get("SQL_SERVER",   "173.249.59.109")
SQL_PORT     = int(os.environ.get("SQL_PORT", "1433"))
SQL_DATABASE = os.environ.get("SQL_DATABASE", "mwananchidy365")
SQL_USER     = os.environ.get("SQL_USER",     "monday")
SQL_PASSWORD = os.environ.get("SQL_PASSWORD", "Monday@2026")

# ── DB helpers ────────────────────────────────────────────────────────────────

def get_conn():
    return pymssql.connect(
        server=SQL_SERVER, port=SQL_PORT,
        user=SQL_USER, password=SQL_PASSWORD,
        database=SQL_DATABASE, login_timeout=20, as_dict=True
    )

def run_sp(sp_name):
    conn = get_conn()
    cur  = conn.cursor()
    cur.execute(f"EXEC {sp_name}")
    rows = cur.fetchall()
    conn.close()
    clean = []
    for row in rows:
        r = {}
        for k, v in row.items():
            r[k] = v.isoformat() if isinstance(v, datetime) else v
        clean.append(r)
    return clean

def filter_branch(rows, key, branch):
    if not branch or branch.upper() == "ALL":
        return rows
    return [r for r in rows if str(r.get(key, "")).upper() == branch.upper()]

# ── Serve the widget HTML ─────────────────────────────────────────────────────

@app.route("/")
def index():
    """Serves the portfolio dashboard widget."""
    return render_template("widget.html")

@app.route("/health")
def health():
    return jsonify({"status": "ok", "time": datetime.now().isoformat()})

# ── API endpoints ─────────────────────────────────────────────────────────────

@app.route("/api/branches")
def api_branches():
    try:
        rows  = run_sp("sp_PortfolioInArrearsByBranch")
        names = sorted({
            r["BranchName"] for r in rows
            if r.get("BranchName") and str(r["BranchName"]).upper() != "TOTAL"
        })
        return jsonify({"status": "ok", "branches": names})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/summary")
def api_summary():
    try:
        conn = get_conn()
        cur  = conn.cursor()
        cur.execute("EXEC sp_ActiveLoanCount");   active    = cur.fetchone()["ActiveLoanCount"]
        cur.execute("EXEC sp_TotalLoanBalance");   total_bal = cur.fetchone()["PrincipalOutstandingBalance"]
        cur.execute("EXEC sp_LoanPrincipalBalance"); principal = cur.fetchone()["PrincipalOutstandingBalance"]
        cur.execute("EXEC sp_Portfoliainarrears"); par_row   = cur.fetchone()
        conn.close()
        return jsonify({
            "status": "ok",
            "active_loans": active,
            "total_balance": total_bal,
            "principal_outstanding": principal,
            "amount_in_arrears": par_row["TotalAmountInArrears"],
            "loan_count_in_arrears": par_row["LoanCount"],
            "par_percentage": par_row["PercentageInArrears"],
            "refreshed_at": datetime.now().isoformat()
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/branch-arrears")
def api_branch_arrears():
    branch = request.args.get("branch", "ALL")
    try:
        rows = run_sp("sp_PortfolioInArrearsByBranch")
        rows = [r for r in rows if str(r.get("BranchName", "")).upper() != "TOTAL"]
        rows = filter_branch(rows, "BranchName", branch)
        return jsonify({"status": "ok", "data": rows, "refreshed_at": datetime.now().isoformat()})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/loan-product")
def api_loan_product():
    try:
        rows = run_sp("sp_PortfolioByLoanProduct")
        rows = [r for r in rows if str(r.get("LoanProductType", "")).upper() != "TOTAL"]
        return jsonify({"status": "ok", "data": rows, "refreshed_at": datetime.now().isoformat()})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/loan-category")
def api_loan_category():
    try:
        rows = run_sp("sp_PortfolioByLoanCategory")
        rows = [r for r in rows if str(r.get("LoanCategory", "")).upper() != "TOTAL"]
        return jsonify({"status": "ok", "data": rows, "refreshed_at": datetime.now().isoformat()})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

# ── Run locally ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    print(f"\n  Mwananchi Credit Portfolio Server")
    print(f"  Running at http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=False)
