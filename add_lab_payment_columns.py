import sqlite3

DB_NAME = 'hospital.db'

conn = sqlite3.connect(DB_NAME)
cursor = conn.cursor()

# ===== چک کردن ستون‌های موجود =====
cursor.execute("PRAGMA table_info(lab_requests)")
existing_columns = [col[1] for col in cursor.fetchall()]

print("=" * 60)
print("Current columns in lab_requests:")
for col in existing_columns:
    print(f"  - {col}")
print("=" * 60)

# ===== اضافه کردن ستون‌ها (اگه نبودن) =====

if 'payment_status' not in existing_columns:
    cursor.execute("ALTER TABLE lab_requests ADD COLUMN payment_status TEXT DEFAULT 'Pending'")
    print("✅ Added: payment_status")
else:
    print("ℹ️ Already exists: payment_status")

if 'final_amount' not in existing_columns:
    cursor.execute("ALTER TABLE lab_requests ADD COLUMN final_amount REAL DEFAULT 0")
    print("✅ Added: final_amount")
else:
    print("ℹ️ Already exists: final_amount")

if 'payment_method' not in existing_columns:
    cursor.execute("ALTER TABLE lab_requests ADD COLUMN payment_method TEXT")
    print("✅ Added: payment_method")
else:
    print("ℹ️ Already exists: payment_method")

if 'receipt_number' not in existing_columns:
    cursor.execute("ALTER TABLE lab_requests ADD COLUMN receipt_number TEXT")
    print("✅ Added: receipt_number")
else:
    print("ℹ️ Already exists: receipt_number")

conn.commit()

# ===== نمایش نتیجه =====
cursor.execute("PRAGMA table_info(lab_requests)")
new_columns = [col[1] for col in cursor.fetchall()]

print("\n" + "=" * 60)
print("New columns in lab_requests:")
for col in new_columns:
    print(f"  - {col}")
print("=" * 60)

conn.close()
print("\n✅ Done!")