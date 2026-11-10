import sqlite3

DB_NAME = 'hospital.db'

conn = sqlite3.connect(DB_NAME)
cursor = conn.cursor()

# ===== جدول medical_records =====
cursor.execute('''
CREATE TABLE IF NOT EXISTS medical_records (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    patient_id INTEGER NOT NULL,
    doctor_id INTEGER NOT NULL,
    record_date TEXT NOT NULL,
    chief_complaint TEXT,
    diagnosis TEXT,
    symptoms TEXT,
    blood_pressure TEXT,
    temperature REAL,
    weight REAL,
    height REAL,
    notes TEXT,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY(patient_id) REFERENCES patients(id),
    FOREIGN KEY(doctor_id) REFERENCES doctors(id)
)
''')
print('✅ Table "medical_records" created')

# ===== جدول prescriptions =====
cursor.execute('''
CREATE TABLE IF NOT EXISTS prescriptions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    record_id INTEGER NOT NULL,
    medicine_id INTEGER NOT NULL,
    dosage TEXT,
    frequency TEXT,
    duration TEXT,
    quantity INTEGER DEFAULT 1,
    instructions TEXT,
    FOREIGN KEY(record_id) REFERENCES medical_records(id) ON DELETE CASCADE,
    FOREIGN KEY(medicine_id) REFERENCES medicines(id)
)
''')
print('✅ Table "prescriptions" created')

conn.commit()
conn.close()
print('✅ Database setup complete!')