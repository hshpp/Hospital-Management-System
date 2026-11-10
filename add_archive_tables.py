import sqlite3

DB_NAME = 'hospital.db'

conn = sqlite3.connect(DB_NAME)
cursor = conn.cursor()

# ===== جدول archive_medical_records =====
cursor.execute('''
CREATE TABLE IF NOT EXISTS archive_medical_records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    original_id INTEGER NOT NULL,
    patient_id INTEGER,
    doctor_id INTEGER,
    record_date TEXT,
    chief_complaint TEXT,
    diagnosis TEXT,
    symptoms TEXT,
    blood_pressure TEXT,
    temperature REAL,
    weight REAL,
    height REAL,
    notes TEXT,
    patient_name TEXT,
    doctor_name TEXT,
    created_at TIMESTAMP,
    archived_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    archived_by TEXT
)
''')
print('✅ Table "archive_medical_records" created')

# ===== جدول archive_prescriptions =====
cursor.execute('''
CREATE TABLE IF NOT EXISTS archive_prescriptions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    original_id INTEGER NOT NULL,
    original_record_id INTEGER,
    medicine_id INTEGER,
    medicine_name TEXT,
    dosage TEXT,
    frequency TEXT,
    duration TEXT,
    quantity INTEGER,
    instructions TEXT,
    archived_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)
''')
print('✅ Table "archive_prescriptions" created')

# ===== Indexes =====
cursor.execute('CREATE INDEX IF NOT EXISTS idx_archive_mr_date ON archive_medical_records(record_date)')
cursor.execute('CREATE INDEX IF NOT EXISTS idx_archive_mr_patient ON archive_medical_records(patient_id)')
cursor.execute('CREATE INDEX IF NOT EXISTS idx_archive_px_record ON archive_prescriptions(original_record_id)')

print('✅ Indexes created')

conn.commit()
conn.close()
print('✅ Archive tables ready!')