import sqlite3

DB_NAME = 'hospital.db'

conn = sqlite3.connect(DB_NAME)
cursor = conn.cursor()

cursor.execute('''
CREATE TABLE IF NOT EXISTS payment_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    payment_id INTEGER NOT NULL,
    service_id INTEGER,
    service_name TEXT NOT NULL,
    quantity INTEGER DEFAULT 1,
    unit_price REAL DEFAULT 0,
    subtotal REAL DEFAULT 0,
    FOREIGN KEY(payment_id) REFERENCES payments(id) ON DELETE CASCADE
)
''')
print('✅ Table "payment_items" created')

cursor.execute('CREATE INDEX IF NOT EXISTS idx_payment_items_payment ON payment_items(payment_id)')
print('✅ Index created')

conn.commit()
conn.close()
print('✅ Done!')