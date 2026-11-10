import sqlite3

DB_NAME = 'hospital.db'

conn = sqlite3.connect(DB_NAME)
cursor = conn.cursor()

# چک کن ستون وجود داره یا نه
cursor.execute("PRAGMA table_info(users)")
columns = [col[1] for col in cursor.fetchall()]

if 'recovery_code' not in columns:
    cursor.execute("ALTER TABLE users ADD COLUMN recovery_code TEXT")
    print("✅ ستون recovery_code اضافه شد")
else:
    print("ℹ️ ستون recovery_code از قبل وجود دارد")

conn.commit()
conn.close()
print("✅ دیتابیس آپدیت شد!")