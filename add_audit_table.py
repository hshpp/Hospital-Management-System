import sqlite3

DB_NAME = 'hospital.db'

conn = sqlite3.connect(DB_NAME)
cursor = conn.cursor()

cursor.execute('''
CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id INTEGER,
    username TEXT,
    user_role TEXT,
    action TEXT NOT NULL,
    entity_type TEXT,
    entity_id INTEGER,
    details TEXT,
    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)
''')
print('✅ Table "audit_log" created')

# Index برای جستجوی سریع
cursor.execute('CREATE INDEX IF NOT EXISTS idx_audit_timestamp ON audit_log(timestamp)')
cursor.execute('CREATE INDEX IF NOT EXISTS idx_audit_user ON audit_log(user_id)')
cursor.execute('CREATE INDEX IF NOT EXISTS idx_audit_action ON audit_log(action)')
print('✅ Indexes created')

conn.commit()
conn.close()
print('✅ Audit log table ready!')