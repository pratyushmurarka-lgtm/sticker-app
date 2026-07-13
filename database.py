import os
import pyodbc

# Global defaults
DEFAULT_SERVER = r"ERPINTECH\SQLEXPRESS"
DEFAULT_DATABASE = "Es_Comp0001_2026"
DEFAULT_USER = "sa"
DEFAULT_PASSWORD = "12345678"

def get_db_credentials():
    config_path = r"C:\dbconfig.txt"
    db_mode = "TEST"
    
    config = {
        "LIVE_SERVER": "", "LIVE_DATABASE": "", "LIVE_USER": "", "LIVE_PASSWORD": "",
        "TEST_SERVER": "", "TEST_DATABASE": "", "TEST_USER": "", "TEST_PASSWORD": ""
    }
    
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if line and "=" in line:
                        parts = line.split("=", 1)
                        key = parts[0].strip().upper()
                        val = parts[1].strip()
                        if key == "MODE":
                            db_mode = val.upper()
                        elif key in config:
                            config[key] = val
        except Exception:
            pass  # Fall back to defaults on error
            
    if db_mode == "LIVE":
        server = config["LIVE_SERVER"] or DEFAULT_SERVER
        database = config["LIVE_DATABASE"] or DEFAULT_DATABASE
        user = config["LIVE_USER"] or DEFAULT_USER
        password = config["LIVE_PASSWORD"] or DEFAULT_PASSWORD
    else:
        server = config["TEST_SERVER"] or DEFAULT_SERVER
        database = config["TEST_DATABASE"] or DEFAULT_DATABASE
        user = config["TEST_USER"] or DEFAULT_USER
        password = config["TEST_PASSWORD"] or DEFAULT_PASSWORD
        
    return server, database, user, password

def get_connection():
    server, database, user, password = get_db_credentials()
    
    # Dynamically find the best ODBC driver on this Windows machine
    available_drivers = [d for d in pyodbc.drivers() if "sql server" in d.lower()]
    if not available_drivers:
        raise RuntimeError("No SQL Server ODBC drivers found on this system.")
        
    # Prefer newer drivers but fall back to the native "SQL Server" driver
    driver = "SQL Server"
    for d in ["ODBC Driver 17 for SQL Server", "ODBC Driver 13 for SQL Server", "SQL Server Native Client 11.0"]:
        if d in available_drivers:
            driver = d
            break
            
    conn_str = f"Driver={{{driver}}};Server={server};Database={database};UID={user};PWD={password};"
    return pyodbc.connect(conn_str)

def query_row(sql, params=()):
    with get_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(sql, params)
            row = cursor.fetchone()
            if row:
                # Convert pyodbc Row to dict using cursor description
                columns = [column[0] for column in cursor.description]
                return dict(zip(columns, row))
            return None

def query_all(sql, params=()):
    with get_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(sql, params)
            rows = cursor.fetchall()
            columns = [column[0] for column in cursor.description]
            return [dict(zip(columns, row)) for row in rows]

def execute_non_query(sql, params=()):
    with get_connection() as conn:
        with conn.cursor() as cursor:
            cursor.execute(sql, params)
            conn.commit()

def execute_transaction(statements_with_params):
    """
    Executes multiple SQL statements inside a single transaction.
    This is extremely fast and reduces deadlock issues by committing all updates at once.
    statements_with_params: List of tuples (sql_query, params_tuple)
    """
    with get_connection() as conn:
        with conn.cursor() as cursor:
            try:
                for sql, params in statements_with_params:
                    cursor.execute(sql, params)
                conn.commit()
            except Exception as e:
                conn.rollback()
                raise e
