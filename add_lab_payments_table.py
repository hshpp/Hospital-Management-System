import sqlite3

DB_NAME = 'hospital.db'

conn = sqlite3.connect(DB_NAME)
cursor = conn.cursor()

# ===== ساخت جدول lab_payments =====
cursor.execute('''
CREATE TABLE IF NOT EXISTS lab_payments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    receipt_number TEXT UNIQUE NOT NULL,
    patient_id INTEGER NOT NULL,
    lab_request_id INTEGER,
    amount REAL NOT NULL,
    discount REAL DEFAULT 0,
    final_amount REAL NOT NULL,
    payment_method TEXT DEFAULT 'Cash',
    payment_status TEXT DEFAULT 'Paid',
    payment_date TIMESTAMP DEFAULT (datetime('now', 'localtime')),
    created_by TEXT,
    notes TEXT,
    FOREIGN KEY(patient_id) REFERENCES patients(id),
    FOREIGN KEY(lab_request_id) REFERENCES lab_requests(id)
)
''')
print("✅ Table 'lab_payments' created")

cursor.execute('CREATE INDEX IF NOT EXISTS idx_lab_pay_patient ON lab_payments(patient_id)')
cursor.execute('CREATE INDEX IF NOT EXISTS idx_lab_pay_request ON lab_payments(lab_request_id)')
print("✅ Indexes created")

conn.commit()
conn.close()
print("\n✅ Done!")