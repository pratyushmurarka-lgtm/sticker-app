import os
import re
import datetime
import io
import segno
from flask import Flask, request, jsonify, render_template, redirect, url_for, session, send_file
import database
import pdf_handler

app = Flask(__name__)
app.secret_key = "qr_code_dashboard_super_secret_key_12345"

# Folder to store customer PDFs
PDF_STORAGE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "customer_pdfs")
os.makedirs(PDF_STORAGE_DIR, exist_ok=True)

# ---------------------------------------------------------
# Authentication Middleware Helpers
# ---------------------------------------------------------
def is_logged_in():
    return "user_id" in session

def is_admin():
    return session.get("user_role") == "Admin"

def get_next_serial_number(item_type, item_code, item_brand):
    qr_month = datetime.date.today().strftime("%m%yyyy")
    
    # Special Brand 1612 / Type 81 shared sequence condition
    if item_brand == 1612 and item_type == 81:
        sql = """
            SELECT ISNULL(MAX(q.SerialNo), 0) + 1 AS NextSerial
            FROM In_QRmaster q
            INNER JOIN ItemMaster i ON q.ItemCode = i.CodeStr
            WHERE i.ItemBrand = 1612 AND i.ItemType = 81 AND q.QRMonth = ?
        """
        params = (qr_month,)
    else:
        # Standard serialization grouping rules
        if item_type in [97, 4797, 10991]:
            sql = "SELECT ISNULL(MAX(SerialNo), 0) + 1 AS NextSerial FROM In_QRmaster WHERE ItemCode=?"
            params = (item_code,)
        else:
            sql = "SELECT ISNULL(MAX(SerialNo), 0) + 1 AS NextSerial FROM In_QRmaster WHERE ItemCode=? AND QRMonth=?"
            params = (item_code, qr_month)
            
    res = database.query_row(sql, params)
    return res["NextSerial"] if res else 1

def is_valid_for_printing(qr_value):
    # 1. Check registry
    reg_check = database.query_row("SELECT QRValue FROM In_QRmaster WHERE QRValue = ?", (qr_value,))
    if reg_check:
        return False, "QR code already exists in QR registry (In_QRmaster)."
        
    # 2. Check packaging scan log
    scan_check = database.query_row("SELECT COUNT(1) AS ScanCount FROM esjobscandata WHERE ItemDetail = ?", (qr_value,))
    if scan_check and scan_check["ScanCount"] > 0:
        return False, "QR code has already been scanned on a product box (esjobscandata)."
        
    return True, ""

# ---------------------------------------------------------
# Page Routes
# ---------------------------------------------------------
@app.route("/")
def index():
    if is_logged_in():
        return redirect(url_for("dashboard"))
    return redirect(url_for("login"))

@app.route("/login", methods=["GET", "POST"])
def login():
    if is_logged_in():
        return redirect(url_for("dashboard"))
        
    if request.method == "POST":
        data = request.get_json() or request.form
        user_id = data.get("username", "").strip()
        password = data.get("password", "").strip()
        
        if not user_id or not password:
            return jsonify({"success": False, "message": "Please enter both Username and Password."}), 400
            
        sql = "SELECT UserID, UserRole FROM UserMaster WHERE UserID = ? AND PasswordHash = ?"
        user = database.query_row(sql, (user_id, password))
        
        if user:
            session["user_id"] = user["UserID"]
            session["user_role"] = user["UserRole"]
            return jsonify({"success": True, "role": user["UserRole"]})
        else:
            return jsonify({"success": False, "message": "Invalid Username or Password."}), 401
            
    return render_template("login.html")

@app.route("/dashboard")
def dashboard():
    if not is_logged_in():
        return redirect(url_for("login"))
        
    # Query loaded items count
    items_count = database.query_row("""
        SELECT COUNT(1) AS ItemCount 
        FROM ItemMaster 
        WHERE Blocked = 0 AND ItemType IN (81,13414,94,4796,4846,82,22)
    """)
    loaded_count = items_count["ItemCount"] if items_count else 0
    
    return render_template("dashboard.html", 
                           user_id=session["user_id"], 
                           user_role=session["user_role"], 
                           items_loaded=loaded_count)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

# ---------------------------------------------------------
# Core APIs
# ---------------------------------------------------------
@app.route("/api/items", methods=["GET"])
def get_items():
    if not is_logged_in():
        return jsonify({"message": "Unauthorized"}), 401
        
    sql = """
        SELECT CodeStr, Name 
        FROM ItemMaster 
        WHERE Blocked = 0 AND ItemType IN (81,13414,94,4796,4846,82,22) 
        ORDER BY CodeStr
    """
    items = database.query_all(sql)
    return jsonify(items)

@app.route("/api/last_serial", methods=["GET"])
def get_last_serial():
    if not is_logged_in():
        return jsonify({"message": "Unauthorized"}), 401
        
    item_code = request.args.get("item_code", "").strip()
    if not item_code:
        return jsonify({"last_serial": "N/A"})
        
    # Fetch details
    item_details = database.query_row("SELECT ItemType, ItemBrand FROM ItemMaster WHERE CodeStr=?", (item_code,))
    if not item_details:
        return jsonify({"last_serial": "N/A"})
        
    item_type = item_details["ItemType"]
    item_brand = item_details["ItemBrand"] or 0
    
    if item_brand == 1612 and item_type == 81:
        sql = """
            SELECT TOP 1 q.QRValue 
            FROM In_QRmaster q 
            INNER JOIN ItemMaster i ON q.ItemCode = i.CodeStr 
            WHERE i.ItemBrand = 1612 AND i.ItemType = 81 
            ORDER BY q.QRID DESC
        """
        res = database.query_row(sql)
    else:
        sql = "SELECT TOP 1 QRValue FROM In_QRmaster WHERE ItemCode = ? ORDER BY QRID DESC"
        res = database.query_row(sql, (item_code,))
        
    last_serial = res["QRValue"] if res else "None"
    return jsonify({"last_serial": last_serial})

@app.route("/api/generate_qrs", methods=["POST"])
def generate_qrs():
    if not is_logged_in():
        return jsonify({"message": "Unauthorized"}), 401
        
    data = request.get_json() or {}
    item_code = data.get("item_code", "").strip()
    qty = int(data.get("qty", 0))
    
    if not item_code or qty <= 0:
        return jsonify({"success": False, "message": "Please select a valid item and quantity."}), 400
        
    item_details = database.query_row("""
        SELECT Code, OtherDes, CodeStr, MRP, Name, ItemType, ItemBrand 
        FROM ItemMaster 
        WHERE CodeStr=?
    """, (item_code,))
    
    if not item_details:
        return jsonify({"success": False, "message": "Item details not found."}), 404
        
    item_numeric_code = item_details["Code"]
    item_other_des = item_details["OtherDes"] or ""
    item_mrp = item_details["MRP"] or 0.0
    item_name = item_details["Name"]
    item_type = item_details["ItemType"]
    item_brand = item_details["ItemBrand"] or 0
    
    qr_month = datetime.date.today().strftime("%m%yyyy")
    qr_list = []
    statements = []
    
    try:
        # Start a local offset for serials to avoid querying MAX inside the loop
        start_serial = get_next_serial_number(item_type, item_code, item_brand)
        
        for idx in range(qty):
            serial_no = start_serial + idx
            
            # Format: CODE|OtherDes|codestr|YYYYMM|MRP|NAME|SRNO
            qr_value = f"{item_numeric_code}|{item_other_des}|{item_code}|{datetime.datetime.now().strftime('%Y%m')}|{item_mrp}|{item_name}|{serial_no}"
            
            # Run validation
            valid, err = is_valid_for_printing(qr_value)
            if not valid:
                return jsonify({"success": False, "message": f"Collision bypass triggered: {err}"}), 409
                
            # Collect SQL insert statements to execute inside a single transaction
            sql = "INSERT INTO In_QRmaster (ItemCode, QRMonth, SerialNo, QRValue, GeneratedOn, Printed, IsCustomerQR) VALUES (?, ?, ?, ?, GETDATE(), 0, 0)"
            params = (item_code, qr_month, serial_no, qr_value)
            statements.append((sql, params))
            
            qr_list.append(qr_value)
            
        # Execute all inserts inside a single database transaction
        database.execute_transaction(statements)
        
        return jsonify({
            "success": True, 
            "count": len(qr_list), 
            "qrs": qr_list,
            "message": f"Generated {len(qr_list)} QR codes successfully."
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"Database error during generation: {str(e)}"}), 500

@app.route("/api/pdf_list", methods=["GET"])
def get_pdf_list():
    if not is_logged_in():
        return jsonify({"message": "Unauthorized"}), 401
        
    sql = "SELECT PDFID, ItemCode, FileName, TotalQty, Printed FROM In_CustomerPDF ORDER BY PDFID DESC"
    pdfs = database.query_all(sql)
    return jsonify(pdfs)

@app.route("/api/import_pdf", methods=["POST"])
def import_pdf():
    if not is_logged_in():
        return jsonify({"message": "Unauthorized"}), 401
        
    if "pdf_file" not in request.files:
        return jsonify({"success": False, "message": "No file uploaded."}), 400
        
    pdf_file = request.files["pdf_file"]
    item_code = request.form.get("item_code", "").strip()
    regex_pattern = request.form.get("regex_pattern", "").strip()
    
    if not item_code or not regex_pattern or pdf_file.filename == "":
        return jsonify({"success": False, "message": "Please select a model, regex pattern, and PDF file."}), 400
        
    # Save PDF temporarily to extract QR codes
    temp_pdf_path = os.path.join(PDF_STORAGE_DIR, f"temp_{datetime.datetime.now().timestamp()}.pdf")
    pdf_file.save(temp_pdf_path)
    
    try:
        extracted_qrs = pdf_handler.scan_pdf_for_qrs(temp_pdf_path, regex_pattern)
        
        if not extracted_qrs:
            if os.path.exists(temp_pdf_path):
                os.remove(temp_pdf_path)
            return jsonify({"success": False, "message": "No QR codes matching the regex format were found in this PDF."}), 422
            
        # Check duplicate filename in database
        filename = pdf_file.filename
        dup_check = database.query_row("SELECT PDFID FROM In_CustomerPDF WHERE FileName = ?", (filename,))
        if dup_check:
            if os.path.exists(temp_pdf_path):
                os.remove(temp_pdf_path)
            return jsonify({"success": False, "message": "A PDF with this filename has already been imported."}), 409
            
        # Rename file to a persistent location
        persistent_filename = f"{datetime.datetime.now().strftime('%y%m%d_%H%M%S')}_{filename}"
        persistent_path = os.path.join(PDF_STORAGE_DIR, persistent_filename)
        os.rename(temp_pdf_path, persistent_path)
        
        # Save to database in a single transaction
        statements = []
        
        # 1. Insert In_CustomerPDF
        sql_pdf = "INSERT INTO In_CustomerPDF (ItemCode, FileName, FilePath, TotalQty, Printed, ReprintCount) VALUES (?, ?, ?, ?, 0, 0)"
        params_pdf = (item_code, filename, persistent_path, len(extracted_qrs))
        statements.append((sql_pdf, params_pdf))
        
        # Get next PDFID via subquery instead of @@IDENTITY to bundle in one transaction block
        # We will insert PDF entry, then in Python we run it, get the ID, then insert QR records.
        # To keep database connection open in transaction, we run a custom block here:
        with database.get_connection() as conn:
            with conn.cursor() as cursor:
                conn.autocommit = False
                try:
                    cursor.execute(sql_pdf, params_pdf)
                    cursor.execute("SELECT @@IDENTITY AS NextID")
                    pdf_id = int(cursor.fetchone()[0])
                    
                    qr_month = datetime.date.today().strftime("%m%yyyy")
                    for qr in extracted_qrs:
                        sql_qr = "INSERT INTO In_QRmaster (ItemCode, QRMonth, SerialNo, QRValue, GeneratedOn, Printed, IsCustomerQR, CustomerPDFID) VALUES (?, ?, 0, ?, GETDATE(), 0, 1, ?)"
                        cursor.execute(sql_qr, (item_code, qr_month, qr, pdf_id))
                        
                    conn.commit()
                except Exception as e:
                    conn.rollback()
                    if os.path.exists(persistent_path):
                        os.remove(persistent_path)
                    raise e
                    
        return jsonify({
            "success": True, 
            "count": len(extracted_qrs),
            "message": f"Successfully imported PDF with {len(extracted_qrs)} unique QR codes."
        })
    except Exception as e:
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        return jsonify({"success": False, "message": f"Error during import: {str(e)}"}), 500

@app.route("/api/get_pdf_file/<int:pdf_id>", methods=["GET"])
def get_pdf_file(pdf_id):
    if not is_logged_in():
        return jsonify({"message": "Unauthorized"}), 401
        
    sql = "SELECT FilePath, FileName FROM In_CustomerPDF WHERE PDFID = ?"
    pdf = database.query_row(sql, (pdf_id,))
    if not pdf or not os.path.exists(pdf["FilePath"]):
        return jsonify({"message": "File not found"}), 404
        
    from_page = request.args.get("from_page")
    to_page = request.args.get("to_page")
    
    if from_page and to_page:
        try:
            from_p = int(from_page)
            to_p = int(to_page)
            temp_split_path = os.path.join(PDF_STORAGE_DIR, f"split_temp_{datetime.datetime.now().timestamp()}.pdf")
            pdf_handler.split_pdf_pages(pdf["FilePath"], temp_split_path, from_p, to_p)
            return send_file(temp_split_path, mimetype="application/pdf", as_attachment=True, download_name=f"split_{pdf['FileName']}")
        except Exception as e:
            return jsonify({"message": f"Failed to split PDF: {str(e)}"}), 500
            
    return send_file(pdf["FilePath"], mimetype="application/pdf", as_attachment=True, download_name=pdf["FileName"])

@app.route("/api/update_print_status", methods=["POST"])
def update_print_status():
    if not is_logged_in():
        return jsonify({"message": "Unauthorized"}), 401
        
    data = request.get_json() or {}
    pdf_id = data.get("pdf_id")
    is_reprint = data.get("is_reprint", False)
    supervisor_id = data.get("supervisor_id", "").strip()
    
    if not pdf_id:
        return jsonify({"success": False, "message": "Missing PDF ID"}), 400
        
    sql_check = "SELECT Printed, FileName FROM In_CustomerPDF WHERE PDFID = ?"
    pdf = database.query_row(sql_check, (pdf_id,))
    if not pdf:
        return jsonify({"success": False, "message": "PDF not found"}), 404
        
    try:
        statements = []
        if is_reprint or pdf["Printed"] == 1:
            # Audit log reprint action
            sql_log = "INSERT INTO QRReprintLog (QRValue, AuthorizedBy, ReprintedOn, Reason) VALUES (?, ?, GETDATE(), ?)"
            log_reason = f"Reprint of PDF: {pdf['FileName']}"
            auth_by = supervisor_id if supervisor_id else session["user_id"]
            statements.append((sql_log, (pdf["FileName"], auth_by, log_reason)))
            
            # Increment reprint count
            sql_pdf = "UPDATE In_CustomerPDF SET ReprintCount = ReprintCount + 1 WHERE PDFID = ?"
            statements.append((sql_pdf, (pdf_id,)))
        else:
            # Update printed status
            sql_pdf = "UPDATE In_CustomerPDF SET Printed = 1, PrintedOn = GETDATE() WHERE PDFID = ?"
            statements.append((sql_pdf, (pdf_id,)))
            
            sql_qrs = "UPDATE In_QRmaster SET Printed = 1 WHERE CustomerPDFID = ?"
            statements.append((sql_qrs, (pdf_id,)))
            
        database.execute_transaction(statements)
        return jsonify({"success": True, "message": "Database printed status updated."})
    except Exception as e:
        return jsonify({"success": False, "message": f"Database error: {str(e)}"}), 500

@app.route("/api/verify_supervisor", methods=["POST"])
def verify_supervisor():
    if not is_logged_in():
        return jsonify({"message": "Unauthorized"}), 401
        
    data = request.get_json() or {}
    username = data.get("username", "").strip()
    password = data.get("password", "").strip()
    
    if not username or not password:
        return jsonify({"success": False, "message": "Missing credentials"}), 400
        
    sql = "SELECT UserID, UserRole FROM UserMaster WHERE UserID = ? AND PasswordHash = ? AND UserRole IN ('Supervisor', 'Admin')"
    user = database.query_row(sql, (username, password))
    
    if user:
        return jsonify({"success": True, "supervisor_id": user["UserID"]})
    else:
        return jsonify({"success": False, "message": "Invalid Supervisor credentials."}), 401

@app.route("/api/reports", methods=["POST"])
def run_report():
    if not is_logged_in():
        return jsonify({"message": "Unauthorized"}), 401
        
    data = request.get_json() or {}
    report_type = int(data.get("report_type", 0))
    
    if report_type == 0:
        sql = "SELECT ReprintedOn, QRValue, AuthorizedBy, Reason FROM QRReprintLog ORDER BY LogID DESC"
    elif report_type == 1:
        sql = """
            SELECT q.ItemCode, q.QRValue, q.GeneratedOn, s.DATE AS ScanDate,
            CASE WHEN s.ItemDetail IS NOT NULL THEN 'Scanned (Completed)' ELSE 'Unscanned (WIP)' END AS Status
            FROM In_QRmaster q
            LEFT JOIN esjobscandata s ON q.QRValue = s.ItemDetail AND s.DATE >= DATEADD(day, -30, GETDATE())
            WHERE q.GeneratedOn >= DATEADD(day, -30, GETDATE())
            ORDER BY q.QRID DESC
        """
    elif report_type == 2:
        sql = "SELECT PDFID, ItemCode, FileName, TotalQty, Printed, PrintedOn, ReprintCount FROM In_CustomerPDF ORDER BY PDFID DESC"
    elif report_type == 3:
        sql = """
            SELECT CONVERT(date, GeneratedOn) AS ProdDate, ItemCode,
            COUNT(1) AS TotalGenerated,
            SUM(CASE WHEN Printed = 1 THEN 1 ELSE 0 END) AS TotalPrinted
            FROM In_QRmaster
            GROUP BY CONVERT(date, GeneratedOn), ItemCode
            ORDER BY ProdDate DESC, ItemCode
        """
    else:
        return jsonify({"success": False, "message": "Invalid report type."}), 400
        
    try:
        rows = database.query_all(sql)
        # Convert date objects to string for JSON serialization
        for r in rows:
            for k, v in r.items():
                if isinstance(v, (datetime.datetime, datetime.date)):
                    r[k] = v.strftime("%Y-%m-%d %H:%M:%S")
        return jsonify({"success": True, "data": rows})
    except Exception as e:
        return jsonify({"success": False, "message": f"Database error: {str(e)}"}), 500

# ---------------------------------------------------------
# User Management APIs (Admin Only)
# ---------------------------------------------------------
@app.route("/api/users", methods=["GET", "POST"])
def manage_users():
    if not is_logged_in() or not is_admin():
        return jsonify({"message": "Unauthorized"}), 401
        
    if request.method == "GET":
        users = database.query_all("SELECT UserID, PasswordHash, UserRole FROM UserMaster ORDER BY UserID")
        return jsonify(users)
        
    if request.method == "POST":
        data = request.get_json() or {}
        username = data.get("username", "").strip()
        password = data.get("password", "").strip()
        role = data.get("role", "").strip()
        
        if not username or not password or not role:
            return jsonify({"success": False, "message": "Please fill in all details."}), 400
            
        dup_check = database.query_row("SELECT UserID FROM UserMaster WHERE UserID = ?", (username,))
        if dup_check:
            return jsonify({"success": False, "message": "User ID already exists."}), 409
            
        try:
            database.execute_non_query("INSERT INTO UserMaster (UserID, PasswordHash, UserRole) VALUES (?, ?, ?)", 
                                       (username, password, role))
            return jsonify({"success": True, "message": f"User '{username}' created successfully."})
        except Exception as e:
            return jsonify({"success": False, "message": f"Database error: {str(e)}"}), 500

@app.route("/api/users/<username>", methods=["PUT", "DELETE"])
def update_delete_user(username):
    if not is_logged_in() or not is_admin():
        return jsonify({"message": "Unauthorized"}), 401
        
    username = username.strip()
    
    if request.method == "PUT":
        data = request.get_json() or {}
        password = data.get("password", "").strip()
        role = data.get("role", "").strip()
        
        if not password or not role:
            return jsonify({"success": False, "message": "Missing password or role."}), 400
            
        try:
            database.execute_non_query("UPDATE UserMaster SET PasswordHash = ?, UserRole = ? WHERE UserID = ?", 
                                       (password, role, username))
            return jsonify({"success": True, "message": f"User '{username}' updated successfully."})
        except Exception as e:
            return jsonify({"success": False, "message": f"Database error: {str(e)}"}), 500
            
    if request.method == "DELETE":
        if username.lower() == session["user_id"].lower():
            return jsonify({"success": False, "message": "You cannot delete your own logged-in account."}), 403
            
        try:
            database.execute_non_query("DELETE FROM UserMaster WHERE UserID = ?", (username,))
            return jsonify({"success": True, "message": f"User '{username}' deleted successfully."})
        except Exception as e:
            return jsonify({"success": False, "message": f"Database error: {str(e)}"}), 500

@app.route("/api/qr_image")
def get_qr_image():
    text = request.args.get("text", "")
    if not text:
        return "Missing text", 400
    try:
        qr = segno.make(text, error='M')
        buf = io.BytesIO()
        qr.save(buf, kind='png', scale=10, border=1)
        buf.seek(0)
        return send_file(buf, mimetype="image/png")
    except Exception as e:
        return str(e), 500

# ---------------------------------------------------------
# Start Server
# ---------------------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
