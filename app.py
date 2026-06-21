"""
Mwananchi Credit — Portfolio Dashboard
Monday.com App Backend + Static File Server
"""

import os
from flask import Flask, jsonify, request, render_template, send_from_directory
from flask_cors import CORS
import pymssql
from datetime import datetime, date

app = Flask(__name__, template_folder="templates", static_folder="static")
CORS(app)

SQL_SERVER   = os.environ.get("SQL_SERVER",   "173.249.59.109")
SQL_PORT     = int(os.environ.get("SQL_PORT", "1433"))
SQL_DATABASE = os.environ.get("SQL_DATABASE", "mwananchidy365")
SQL_USER     = os.environ.get("SQL_USER",     "monday")
SQL_PASSWORD = os.environ.get("SQL_PASSWORD", "")

def get_conn():
    return pymssql.connect(
        server=SQL_SERVER, port=SQL_PORT,
        user=SQL_USER, password=SQL_PASSWORD,
        database=SQL_DATABASE, login_timeout=20, as_dict=True
    )

def clean_rows(rows):
    out = []
    for row in rows:
        r = {}
        for k, v in row.items():
            if isinstance(v, (datetime, date)):
                r[k] = v.isoformat()
            elif v is None:
                r[k] = None
            else:
                r[k] = v
        out.append(r)
    return out

def run_sp(sp_name, as_at_date=None):
    """Call a stored procedure, optionally passing @AsAtDate."""
    conn = get_conn()
    cur  = conn.cursor()
    if as_at_date:
        try:
            cur.execute(f"EXEC {sp_name} @AsAtDate = %s", (as_at_date,))
            rows = cur.fetchall()
        except Exception:
            # SP may not accept @AsAtDate — retry without parameter
            conn.close()
            conn = get_conn()
            cur  = conn.cursor()
            cur.execute(f"EXEC {sp_name}")
            rows = cur.fetchall()
    else:
        cur.execute(f"EXEC {sp_name}")
        rows = cur.fetchall()
    conn.close()
    return clean_rows(rows)

def filter_branch(rows, key, branch):
    if not branch or branch.upper() == "ALL":
        return rows
    return [r for r in rows if str(r.get(key, "")).upper() == branch.upper()]

def today_str():
    return datetime.now().strftime("%Y-%m-%d")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

@app.route("/")
def index():
    # Serve the live dashboard as the default page
    return send_from_directory(BASE_DIR, "mcl_dashboard.html")

@app.route("/dashboard")
def dashboard():
    return send_from_directory(BASE_DIR, "mcl_dashboard.html")

@app.route("/report")
def report():
    return send_from_directory(BASE_DIR, "mcl_portfolio_report.html")

@app.route("/widget")
def widget():
    return render_template("widget.html")

@app.route("/health")
def health():
    return jsonify({"status": "ok", "time": datetime.now().isoformat()})

# ── Debug: see raw SP output ──────────────────────────────────
@app.route("/api/debug")
def api_debug():
    sp = request.args.get("sp", "sp_PortfolioInArrearsByBranch")
    dt = request.args.get("date", today_str())
    try:
        rows = run_sp(sp, as_at_date=dt)
        cols = list(rows[0].keys()) if rows else []
        return jsonify({
            "status": "ok",
            "sp": sp,
            "date_used": dt,
            "row_count": len(rows),
            "columns": cols,
            "sample": rows[:5]
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

# ── API endpoints ─────────────────────────────────────────────

@app.route("/api/branches")
def api_branches():
    dt = request.args.get("date", today_str())
    try:
        rows  = run_sp("sp_PortfolioInArrearsByBranch", as_at_date=dt)
        names = sorted({
            r["BranchName"] for r in rows
            if r.get("BranchName") and str(r["BranchName"]).upper() != "TOTAL"
        })
        return jsonify({"status": "ok", "branches": names})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/summary")
def api_summary():
    dt = request.args.get("date", today_str())
    try:
        rows_active    = run_sp("sp_ActiveLoanCount",      as_at_date=dt)
        rows_total     = run_sp("sp_TotalLoanBalance",     as_at_date=dt)
        rows_principal = run_sp("sp_LoanPrincipalBalance", as_at_date=dt)
        rows_par       = run_sp("sp_Portfoliainarrears",   as_at_date=dt)
        active    = rows_active[0].get("ActiveLoanCount", 0)            if rows_active    else 0
        total_bal = rows_total[0].get("PrincipalOutstandingBalance", 0) if rows_total     else 0
        principal = rows_principal[0].get("PrincipalOutstandingBalance", 0) if rows_principal else 0
        par_row   = rows_par[0] if rows_par else {}
        return jsonify({
            "status": "ok",
            "date": dt,
            "active_loans": active,
            "total_balance": total_bal,
            "principal_outstanding": principal,
            "amount_in_arrears":     par_row.get("TotalAmountInArrears", 0),
            "loan_count_in_arrears": par_row.get("LoanCount", 0),
            "par_percentage":        par_row.get("PercentageInArrears", 0),
            "refreshed_at": datetime.now().isoformat()
        })
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/branch-arrears")
def api_branch_arrears():
    branch = request.args.get("branch", "ALL")
    dt     = request.args.get("date", today_str())
    try:
        rows = run_sp("sp_PortfolioInArrearsByBranch", as_at_date=dt)
        rows = [r for r in rows if str(r.get("BranchName", "")).upper() != "TOTAL"]
        rows = filter_branch(rows, "BranchName", branch)
        return jsonify({"status": "ok", "data": rows, "refreshed_at": datetime.now().isoformat()})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/loan-product")
def api_loan_product():
    branch = request.args.get("branch", "ALL")
    dt     = request.args.get("date", today_str())
    try:
        rows = run_sp("sp_PortfolioByLoanProduct", as_at_date=dt)
        rows = [r for r in rows if str(r.get("LoanProductType", "")).upper() != "TOTAL"]
        return jsonify({"status": "ok", "data": rows, "refreshed_at": datetime.now().isoformat()})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

@app.route("/api/loan-category")
def api_loan_category():
    branch = request.args.get("branch", "ALL")
    dt     = request.args.get("date", today_str())
    try:
        rows = run_sp("sp_PortfolioByLoanCategory", as_at_date=dt)
        rows = [r for r in rows if str(r.get("LoanCategory", "")).upper() != "TOTAL"]
        return jsonify({"status": "ok", "data": rows, "refreshed_at": datetime.now().isoformat()})
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    print(f"\n  Mwananchi Credit Portfolio Server")
    print(f"  Running at http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=False)
