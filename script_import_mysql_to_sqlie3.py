import pymysql
import sqlite3
from decimal import Decimal
from datetime import date, datetime

# --------------------------
# Conexiune MySQL
# --------------------------
mysql_conn = pymysql.connect(
    host="localhost",
    user="root",
    password="cipri",
    database="date_clienti",
    cursorclass=pymysql.cursors.DictCursor
)
mysql_cursor = mysql_conn.cursor()

# --------------------------
# Conexiune SQLite
# --------------------------
sqlite_conn = sqlite3.connect("baza_clienti.sqlite3")
sqlite_cursor = sqlite_conn.cursor()

# --------------------------
# Creare tabele SQLite daca nu exista
# --------------------------
sqlite_cursor.execute("""
CREATE TABLE IF NOT EXISTS tabela_date_clienti (
    Nr_Crt INTEGER PRIMARY KEY AUTOINCREMENT,
    Nume_Firma TEXT,
    Sediu_Social TEXT,
    Cui TEXT UNIQUE,
    Nr_Telefon TEXT,
    Mail TEXT,
    Reg_Comert TEXT,
    Tva TEXT,
    Administrator TEXT,
    Status_Firma TEXT
)
""")

sqlite_cursor.execute("""
CREATE TABLE IF NOT EXISTS tabela_sedii_secundare (
    Nr_Crt INTEGER PRIMARY KEY AUTOINCREMENT,
    Id_Client INTEGER,
    Punct_Lucru TEXT,
    Model_Amef TEXT,
    Serie_Amef TEXT,
    Nui TEXT,
    Tip_Abonament TEXT,
    Data_Conect_Anaf TEXT,
    Tehnician TEXT,
    Data_Exp_Abon TEXT,
    Val_Ctr REAL,
    Data_Exp_Gprs TEXT
)
""")

sqlite_cursor.execute("""
CREATE TABLE IF NOT EXISTS istoric_abonamente (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    id_client INTEGER,
    id_sediu INTEGER,
    serie_amef TEXT,
    tip_abonament TEXT,
    data_start TEXT,
    data_expirare TEXT,
    data_prelungire TEXT,
    observatii TEXT
)
""")

# --------------------------
# Functie de convertire valori
# --------------------------
def convert_value(v):
    if isinstance(v, Decimal):
        return float(v)
    elif isinstance(v, (date, datetime)):
        return v.isoformat()
    elif v is None:
        return None
    else:
        return v

# --------------------------
# Preluare si import clienti
# --------------------------
mysql_cursor.execute("SELECT * FROM tabela_date_clienti")
clienti = mysql_cursor.fetchall()
for c in clienti:
    c_safe = {k: convert_value(v) for k, v in c.items()}
    # verificam daca exista dupa CUI
    sqlite_cursor.execute("SELECT Nr_Crt FROM tabela_date_clienti WHERE Cui=?", (c_safe["Cui"],))
    existing = sqlite_cursor.fetchone()
    if existing:
        placeholders = ", ".join(f"{k}=?" for k in c_safe.keys() if k != "Nr_Crt")
        values = [c_safe[k] for k in c_safe.keys() if k != "Nr_Crt"]
        values.append(existing[0])
        sqlite_cursor.execute(f"UPDATE tabela_date_clienti SET {placeholders} WHERE Nr_Crt=?", values)
    else:
        columns = ", ".join(c_safe.keys())
        placeholders = ", ".join(["?"] * len(c_safe))
        values = [c_safe[k] for k in c_safe.keys()]
        sqlite_cursor.execute(f"INSERT INTO tabela_date_clienti ({columns}) VALUES ({placeholders})", values)

# --------------------------
# Preluare si import sedii secundare
# --------------------------
mysql_cursor.execute("SELECT * FROM tabela_sedii_secundare")
sedii = mysql_cursor.fetchall()
for s in sedii:
    s_safe = {k: convert_value(v) for k, v in s.items()}
    sqlite_cursor.execute("SELECT Nr_Crt FROM tabela_sedii_secundare WHERE Id_Client=? AND Serie_Amef=?",
                          (s_safe["Id_Client"], s_safe["Serie_Amef"]))
    existing = sqlite_cursor.fetchone()
    if existing:
        placeholders = ", ".join(f"{k}=?" for k in s_safe.keys() if k not in ["Nr_Crt", "Id_Client", "Serie_Amef"])
        values = [s_safe[k] for k in s_safe.keys() if k not in ["Nr_Crt", "Id_Client", "Serie_Amef"]]
        values.extend([s_safe["Id_Client"], s_safe["Serie_Amef"]])
        sqlite_cursor.execute(f"UPDATE tabela_sedii_secundare SET {placeholders} WHERE Id_Client=? AND Serie_Amef=?", values)
    else:
        columns = ", ".join(s_safe.keys())
        placeholders = ", ".join(["?"] * len(s_safe))
        values = [s_safe[k] for k in s_safe.keys()]
        sqlite_cursor.execute(f"INSERT INTO tabela_sedii_secundare ({columns}) VALUES ({placeholders})", values)

# --------------------------
# Preluare si import istoric abonamente
# --------------------------
mysql_cursor.execute("SELECT * FROM istoric_abonamente")
istoric = mysql_cursor.fetchall()
for i in istoric:
    i_safe = {k: convert_value(v) for k, v in i.items()}
    sqlite_cursor.execute("SELECT id FROM istoric_abonamente WHERE id_sediu=? AND tip_abonament=? AND data_start=?",
                          (i_safe["id_sediu"], i_safe["tip_abonament"], i_safe["data_start"]))
    existing = sqlite_cursor.fetchone()
    if existing:
        placeholders = ", ".join(f"{k}=?" for k in i_safe.keys() if k not in ["id", "id_sediu", "tip_abonament", "data_start"])
        values = [i_safe[k] for k in i_safe.keys() if k not in ["id", "id_sediu", "tip_abonament", "data_start"]]
        values.extend([i_safe["id_sediu"], i_safe["tip_abonament"], i_safe["data_start"]])
        sqlite_cursor.execute(f"UPDATE istoric_abonamente SET {placeholders} WHERE id_sediu=? AND tip_abonament=? AND data_start=?", values)
    else:
        columns = ", ".join(i_safe.keys())
        placeholders = ", ".join(["?"] * len(i_safe))
        values = [i_safe[k] for k in i_safe.keys()]
        sqlite_cursor.execute(f"INSERT INTO istoric_abonamente ({columns}) VALUES ({placeholders})", values)

# --------------------------
# Commit si inchidere conexiuni
# --------------------------
sqlite_conn.commit()
sqlite_conn.close()
mysql_conn.close()

print("Import finalizat cu succes!")