"""
Fix dates in database to use local time
"""
import sqlite3
from datetime import datetime, timedelta

DB_NAME = 'hospital.db'

conn = sqlite3.connect(DB_NAME)
cursor = conn.cursor()

print("=" * 70)
print("🔍 CURRENT DATES IN DATABASE")
print("=" * 70)

# ===== بررسی تعداد رکوردها =====
tables_to_check = [
    ('payments', 'payment_date'),
    ('medicine_sales', 'sale_date'),
    ('medical_records', 'created_at'),
    ('audit_log', 'timestamp'),
    ('users', 'last_login'),
]

for table, col in tables_to_check:
    try:
        cursor.execute(f'SELECT COUNT(*) FROM {table}')
        total = cursor.fetchone()[0]
        print(f'📋 {table}: {total} records')
    except Exception as e:
        print(f'⚠️ {table}: {e}')

# ===== بررسی تاریخ‌های اخیر =====
print("\n" + "=" * 70)
print("📅 LATEST RECORDS")
print("=" * 70)

# Payments
print("\n💰 Payments (last 5):")
try:
    cursor.execute('SELECT id, payment_date FROM payments ORDER BY id DESC LIMIT 5')
    for row in cursor.fetchall():
        print(f'  ID: {row[0]} | Date: {row[1]}')
except Exception as e:
    print(f'  ⚠️ {e}')

# Medicine sales
print("\n💊 Medicine Sales (last 5):")
try:
    cursor.execute('SELECT id, sale_date FROM medicine_sales ORDER BY id DESC LIMIT 5')
    for row in cursor.fetchall():
        print(f'  ID: {row[0]} | Date: {row[1]}')
except Exception as e:
    print(f'  ⚠️ {e}')

# Medical records
print("\n📋 Medical Records (last 5):")
try:
    cursor.execute('SELECT id, record_date, created_at FROM medical_records ORDER BY id DESC LIMIT 5')
    for row in cursor.fetchall():
        print(f'  ID: {row[0]} | Date: {row[1]} | Created: {row[2]}')
except Exception as e:
    print(f'  ⚠️ {e}')

# ===== آپدیت =====
print("\n" + "=" * 70)
print("🔧 FIXING DATES")
print("=" * 70)

confirm = input("\n⚠️ Update all dates to local time? (yes/no): ")

if confirm.lower() != 'yes':
    print("❌ Cancelled")
    conn.close()
    exit()

# ===== ۱. Payments =====
try:
    cursor.execute('''
        UPDATE payments
        SET payment_date = datetime(payment_date, 'localtime')
        WHERE payment_date IS NOT NULL
    ''')
    print(f'✅ Payments: {cursor.rowcount} rows updated')
except Exception as e:
    print(f'⚠️ Payments: {e}')

# ===== ۲. Medicine Sales =====
try:
    cursor.execute('''
        UPDATE medicine_sales
        SET sale_date = datetime(sale_date, 'localtime')
        WHERE sale_date IS NOT NULL
    ''')
    print(f'✅ Medicine Sales: {cursor.rowcount} rows updated')
except Exception as e:
    print(f'⚠️ Medicine Sales: {e}')

# ===== ۳. Medical Records =====
try:
    cursor.execute('''
        UPDATE medical_records
        SET created_at = datetime(created_at, 'localtime')
        WHERE created_at IS NOT NULL
    ''')
    print(f'✅ Medical Records: {cursor.rowcount} rows updated')
except Exception as e:
    print(f'⚠️ Medical Records: {e}')

# ===== ۴. Audit Log =====
try:
    cursor.execute('''
        UPDATE audit_log
        SET timestamp = datetime(timestamp, 'localtime')
        WHERE timestamp IS NOT NULL
    ''')
    print(f'✅ Audit Log: {cursor.rowcount} rows updated')
except Exception as e:
    print(f'⚠️ Audit Log: {e}')

# ===== ۵. Users =====
try:
    cursor.execute('''
        UPDATE users
        SET last_login = datetime(last_login, 'localtime')
        WHERE last_login IS NOT NULL
    ''')
    print(f'✅ Users: {cursor.rowcount} rows updated')
except Exception as e:
    print(f'⚠️ Users: {e}')

conn.commit()

# ===== نتیجه =====
print("\n" + "=" * 70)
print("✅ AFTER FIX")
print("=" * 70)

print("\n💰 Payments (last 5):")
cursor.execute('SELECT id, payment_date FROM payments ORDER BY id DESC LIMIT 5')
for row in cursor.fetchall():
    print(f'  ID: {row[0]} | Date: {row[1]}')

print("\n💊 Medicine Sales (last 5):")
cursor.execute('SELECT id, sale_date FROM medicine_sales ORDER BY id DESC LIMIT 5')
for row in cursor.fetchall():
    print(f'  ID: {row[0]} | Date: {row[1]}')

conn.close()
print("\n✅ Done! Please restart the application.")