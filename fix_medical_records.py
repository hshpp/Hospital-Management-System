import sqlite3

DB_NAME = 'hospital.db'

conn = sqlite3.connect(DB_NAME)
cursor = conn.cursor()

# ===== ۱. اطلاعات فعلی =====
print("=" * 60)
print("BEFORE FIX:")
print("=" * 60)

cursor.execute('SELECT COUNT(*) FROM medical_records')
print(f"Total medical_records: {cursor.fetchone()[0]}")

cursor.execute('SELECT id, record_date, diagnosis FROM medical_records ORDER BY id')
print("\nExisting records:")
for row in cursor.fetchall():
    print(f"  ID: {row[0]} | Date: {row[1]} | Diagnosis: {(row[2] or '')[:30]}")

# ===== ۲. چک AUTOINCREMENT =====
cursor.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='medical_records'")
schema = cursor.fetchone()[0]
print(f"\nSchema:\n{schema}")

if 'AUTOINCREMENT' in schema.upper():
    print("\n✅ Already has AUTOINCREMENT - no fix needed!")
    conn.close()
    exit()

print("\n⚠️ AUTOINCREMENT missing! Rebuilding tables...")

# ===== ۳. بکاپ درون دیتابیس =====
cursor.execute('ALTER TABLE medical_records RENAME TO medical_records_backup')
cursor.execute('ALTER TABLE prescriptions RENAME TO prescriptions_backup')

# ===== ۴. ساخت جدول جدید با AUTOINCREMENT =====
cursor.execute('''
CREATE TABLE medical_records (
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
print('✅ New "medical_records" created with AUTOINCREMENT')

cursor.execute('''
CREATE TABLE prescriptions (
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
print('✅ New "prescriptions" created with AUTOINCREMENT')

# ===== ۵. کپی داده‌ها =====
cursor.execute('''
INSERT INTO medical_records
(id, patient_id, doctor_id, record_date, chief_complaint, diagnosis,
 symptoms, blood_pressure, temperature, weight, height, notes, created_at)
SELECT id, patient_id, doctor_id, record_date, chief_complaint, diagnosis,
       symptoms, blood_pressure, temperature, weight, height, notes, created_at
FROM medical_records_backup
ORDER BY id
''')
print(f'✅ Copied {cursor.rowcount} records')

cursor.execute('''
INSERT INTO prescriptions
(id, record_id, medicine_id, dosage, frequency, duration, quantity, instructions)
SELECT id, record_id, medicine_id, dosage, frequency, duration, quantity, instructions
FROM prescriptions_backup
ORDER BY id
''')
print(f'✅ Copied {cursor.rowcount} prescriptions')

# ===== ۶. حذف جدول‌های بکاپ =====
cursor.execute('DROP TABLE medical_records_backup')
cursor.execute('DROP TABLE prescriptions_backup')
print('✅ Backup tables dropped')

# ===== ۷. آپدیت sqlite_sequence =====
cursor.execute('''
UPDATE sqlite_sequence
SET seq = (SELECT COALESCE(MAX(id), 0) FROM medical_records)
WHERE name = 'medical_records'
''')
cursor.execute('''
UPDATE sqlite_sequence
SET seq = (SELECT COALESCE(MAX(id), 0) FROM prescriptions)
WHERE name = 'prescriptions'
''')
print('✅ AUTOINCREMENT sequence updated')

conn.commit()

# ===== ۸. نمایش نتیجه =====
print("\n" + "=" * 60)
print("AFTER FIX:")
print("=" * 60)

cursor.execute("SELECT sql FROM sqlite_master WHERE type='table' AND name='medical_records'")
print(f"New Schema:\n{cursor.fetchone()[0]}")

cursor.execute('SELECT COUNT(*) FROM medical_records')
print(f"\nTotal medical_records: {cursor.fetchone()[0]}")

cursor.execute('SELECT id, record_date, diagnosis FROM medical_records ORDER BY id')
print("\nRecords after fix:")
for row in cursor.fetchall():
    print(f"  ID: {row[0]} | Date: {row[1]} | Diagnosis: {(row[2] or '')[:30]}")

conn.close()
print("\n✅ FIX COMPLETE!")
print("🔄 Please RESTART the application.")