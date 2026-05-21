import os
import sqlite3

DB_PATH = os.path.join(os.path.dirname(__file__), 'healthcare.db')
conn = sqlite3.connect(DB_PATH)
cursor = conn.cursor()

cursor.executescript("""
-- Users (patients + doctors share this table, role column differentiates)
CREATE TABLE IF NOT EXISTS users(
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    name          TEXT    NOT NULL,
    email         TEXT    UNIQUE NOT NULL,
    password      TEXT    NOT NULL,
    age           INTEGER,
    gender        TEXT,
    role          TEXT    DEFAULT 'patient',
    phone         TEXT,
    created_at    TEXT    DEFAULT CURRENT_TIMESTAMP
);

-- Doctor professional details
CREATE TABLE IF NOT EXISTS doctors(
    id                INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id           INTEGER UNIQUE,
    specialization    TEXT,
    license_no        TEXT UNIQUE,
    hospital          TEXT,
    experience_years  INTEGER,
    available_days    TEXT,
    consultation_fee  REAL    DEFAULT 500,
    rating            REAL    DEFAULT 0.0,
    total_reviews     INTEGER DEFAULT 0,
    FOREIGN KEY(user_id) REFERENCES users(id)
);

-- Patient medical records
CREATE TABLE IF NOT EXISTS patients(
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id          INTEGER,
    blood_pressure   TEXT,
    blood_sugar      INTEGER,
    allergies        TEXT,
    past_diseases    TEXT,
    current_symptoms TEXT,
    previous_drugs   TEXT,
    diagnosed_disease TEXT,
    recommended_drug  TEXT,
    alternative_drug  TEXT,
    side_effects      TEXT,
    risk_level        TEXT,
    created_at        TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(user_id) REFERENCES users(id)
);

-- Appointments between patients and doctors
CREATE TABLE IF NOT EXISTS appointments(
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    patient_id       INTEGER,
    doctor_id        INTEGER,
    appointment_date TEXT,
    appointment_time TEXT,
    reason           TEXT,
    status           TEXT DEFAULT 'pending',
    notes            TEXT,
    created_at       TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(patient_id) REFERENCES users(id),
    FOREIGN KEY(doctor_id)  REFERENCES users(id)
);

-- Prescriptions (issued by AI or by doctor)
CREATE TABLE IF NOT EXISTS prescriptions(
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    patient_id      INTEGER,
    doctor_id       INTEGER,
    drug            TEXT,
    dosage          TEXT,
    frequency       TEXT,
    duration        TEXT,
    notes           TEXT,
    date_prescribed TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(patient_id) REFERENCES patients(id)
);

-- Doctor clinical notes per appointment
CREATE TABLE IF NOT EXISTS doctor_notes(
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    appointment_id  INTEGER,
    doctor_id       INTEGER,
    patient_user_id INTEGER,
    diagnosis       TEXT,
    treatment_plan  TEXT,
    follow_up_date  TEXT,
    created_at      TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(appointment_id) REFERENCES appointments(id)
);

-- Patient reviews of doctors
CREATE TABLE IF NOT EXISTS reviews(
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    doctor_id  INTEGER,
    patient_id INTEGER,
    rating     INTEGER CHECK(rating BETWEEN 1 AND 5),
    comment    TEXT,
    created_at TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(doctor_id)  REFERENCES doctors(id),
    FOREIGN KEY(patient_id) REFERENCES users(id)
);
""")

def ensure_column_exists(conn, table_name, column_def):
    cursor = conn.cursor()
    cursor.execute(f"PRAGMA table_info({table_name})")
    existing_cols = [row[1] for row in cursor.fetchall()]
    column_name = column_def.split()[0]
    if column_name not in existing_cols:
        cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_def}")
        conn.commit()

conn.commit()
ensure_column_exists(conn, "prescriptions", "doctor_id INTEGER")
conn.close()
print("✅ Database initialised with all tables.")

# ---- Safe migrations: add missing columns if DB was created earlier without them
def ensure_column(table, column_def):
    """Add a column to a table if it does not already exist. column_def should be like 'role TEXT DEFAULT "patient"'"""
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    try:
        cur.execute(f"PRAGMA table_info({table})")
        cols = [r[1] for r in cur.fetchall()]
        col_name = column_def.split()[0]
        if col_name not in cols:
            try:
                cur.execute(f"ALTER TABLE {table} ADD COLUMN {column_def}")
            except sqlite3.OperationalError as e:
                if "non-constant default" in str(e).lower():
                    # SQLite cannot add CURRENT_TIMESTAMP defaults to existing tables.
                    parts = column_def.split()
                    fallback_def = f"{parts[0]} {parts[1]}"
                    cur.execute(f"ALTER TABLE {table} ADD COLUMN {fallback_def}")
                else:
                    raise
            conn.commit()
    except Exception:
        pass
    finally:
        conn.close()

# Ensure commonly-missing columns are present
ensure_column('users', 'role TEXT DEFAULT "patient"')
ensure_column('users', 'phone TEXT')
ensure_column('users', 'created_at TEXT DEFAULT CURRENT_TIMESTAMP')
ensure_column('doctors', 'rating REAL DEFAULT 0.0')
ensure_column('doctors', 'total_reviews INTEGER DEFAULT 0')
ensure_column('patients', 'risk_level TEXT')
ensure_column('patients', 'recommended_drug TEXT')
ensure_column('patients', 'alternative_drug TEXT')
ensure_column('patients', 'side_effects TEXT')
ensure_column('patients', 'created_at TEXT DEFAULT CURRENT_TIMESTAMP')
ensure_column('appointments', 'created_at TEXT DEFAULT CURRENT_TIMESTAMP')
ensure_column('doctor_notes', 'created_at TEXT DEFAULT CURRENT_TIMESTAMP')
ensure_column('reviews', 'created_at TEXT DEFAULT CURRENT_TIMESTAMP')
ensure_column('prescriptions', 'doctor_id INTEGER')
ensure_column('prescriptions', 'frequency TEXT')
ensure_column('prescriptions', 'duration TEXT')
ensure_column('prescriptions', 'notes TEXT')
ensure_column('prescriptions', 'date_prescribed TEXT DEFAULT CURRENT_TIMESTAMP')

# Update missing timestamps so history and report views show row dates correctly
conn = sqlite3.connect(DB_PATH)
cur = conn.cursor()
cur.execute("UPDATE patients SET created_at = CURRENT_TIMESTAMP WHERE created_at IS NULL")
cur.execute("UPDATE prescriptions SET date_prescribed = CURRENT_TIMESTAMP WHERE date_prescribed IS NULL")
conn.commit()
conn.close()

print("✅ Database migrations completed (missing columns added where necessary).")