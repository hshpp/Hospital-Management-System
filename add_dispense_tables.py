import sqlite3

DB_NAME = 'hospital.db'

conn = sqlite3.connect(DB_NAME)
cursor = conn.cursor()

# ===== جدول dispensed_medicines (ثبت تحویل) =====
cursor.execute('''
CREATE TABLE IF NOT EXISTS dispensed_medicines (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    prescription_id INTEGER NOT NULL,
    patient_id INTEGER NOT NULL,
    medicine_id INTEGER NOT NULL,
    medicine_name TEXT NOT NULL,
    quantity INTEGER NOT NULL,
    unit_price REAL DEFAULT 0,
    total_price REAL DEFAULT 0,
    dispensed_at TIMESTAMP DEFAULT (datetime('now', 'localtime')),
    dispensed_by TEXT,
    notes TEXT,
    FOREIGN KEY(prescription_id) REFERENCES prescriptions(id),
    FOREIGN KEY(patient_id) REFERENCES patients(id),
    FOREIGN KEY(medicine_id) REFERENCES medicines(id)
)
''')
print('✅ Table "dispensed_medicines" created')

# ===== جدول pharmacy_payments (پرداخت داروخانه) =====
cursor.execute('''
CREATE TABLE IF NOT EXISTS pharmacy_payments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    receipt_number TEXT UNIQUE NOT NULL,
    patient_id INTEGER NOT NULL,
    dispense_id INTEGER,
    amount REAL NOT NULL,
    discount REAL DEFAULT 0,
    final_amount REAL NOT NULL,
    payment_method TEXT DEFAULT 'Cash',
    payment_status TEXT DEFAULT 'Paid',
    payment_date TIMESTAMP DEFAULT (datetime('now', 'localtime')),
    notes TEXT,
    created_by TEXT,
    FOREIGN KEY(patient_id) REFERENCES patients(id)
)
''')
print('✅ Table "pharmacy_payments" created')

# ===== Indexes =====
cursor.execute('CREATE INDEX IF NOT EXISTS idx_dispensed_patient ON dispensed_medicines(patient_id)')
cursor.execute('CREATE INDEX IF NOT EXISTS idx_dispensed_prescription ON dispensed_medicines(prescription_id)')
cursor.execute('CREATE INDEX IF NOT EXISTS idx_pharmacy_pay_patient ON pharmacy_payments(patient_id)')

print('✅ Indexes created')
print('\n✅ All dispense tables ready!')

conn.commit()
conn.close()