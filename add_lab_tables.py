import sqlite3

DB_NAME = 'hospital.db'

conn = sqlite3.connect(DB_NAME)
cursor = conn.cursor()

# ===== جدول lab_tests =====
cursor.execute('''
CREATE TABLE IF NOT EXISTS lab_tests (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE,
    category TEXT,
    price REAL DEFAULT 0,
    normal_range TEXT,
    unit TEXT,
    description TEXT
)
''')
print('✅ Table "lab_tests" created')

# ===== جدول lab_requests =====
cursor.execute('''
CREATE TABLE IF NOT EXISTS lab_requests (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    patient_id INTEGER NOT NULL,
    doctor_id INTEGER,
    test_id INTEGER NOT NULL,
    record_id INTEGER,
    request_date TEXT,
    status TEXT DEFAULT 'Pending',
    result TEXT,
    result_date TEXT,
    notes TEXT,
    created_at TIMESTAMP DEFAULT (datetime('now', 'localtime')),
    FOREIGN KEY(patient_id) REFERENCES patients(id),
    FOREIGN KEY(doctor_id) REFERENCES doctors(id),
    FOREIGN KEY(test_id) REFERENCES lab_tests(id),
    FOREIGN KEY(record_id) REFERENCES medical_records(id)
)
''')
print('✅ Table "lab_requests" created')

# ===== اضافه کردن آزمایش‌های پیش‌فرض =====
default_tests = [
    ('CBC (Complete Blood Count)', 'Hematology', 150000, '4.5-11.0', '10^3/µL', 'شمارش کامل خون'),
    ('FBS (Fasting Blood Sugar)', 'Biochemistry', 80000, '70-100', 'mg/dL', 'قند خون ناشتا'),
    ('HbA1c', 'Biochemistry', 200000, '4-5.6', '%', 'هموگلوبین گلیکوزیله'),
    ('TSH', 'Hormones', 200000, '0.4-4.0', 'mIU/L', 'هورمون محرک تیروئید'),
    ('T3', 'Hormones', 150000, '80-200', 'ng/dL', 'تری‌یدوتیرونین'),
    ('T4', 'Hormones', 150000, '5-12', 'µg/dL', 'تیروکسین'),
    ('Urine Analysis', 'Urinalysis', 100000, '—', '—', 'آنالیز ادرار'),
    ('Lipid Profile', 'Biochemistry', 250000, '—', '—', 'پروفایل چربی'),
    ('Liver Function Test (LFT)', 'Biochemistry', 300000, '—', '—', 'تست عملکرد کبد'),
    ('Kidney Function Test (KFT)', 'Biochemistry', 300000, '—', '—', 'تست عملکرد کلیه'),
    ('Vitamin D3', 'Vitamins', 250000, '30-100', 'ng/mL', 'ویتامین D'),
    ('Iron (Fe)', 'Minerals', 150000, '60-170', 'µg/dL', 'آهن'),
    ('Ferritin', 'Minerals', 200000, '20-250', 'ng/mL', 'فریتین'),
    ('CRP', 'Immunology', 150000, '< 10', 'mg/L', 'پروتئین واکنشی C'),
    ('ESR', 'Hematology', 100000, '0-20', 'mm/hr', 'سرعت رسوب گلبول'),
]

for test in default_tests:
    try:
        cursor.execute('''
            INSERT INTO lab_tests (name, category, price, normal_range, unit, description)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', test)
    except sqlite3.IntegrityError:
        pass  # اگه قبلاً بود، رد کن

cursor.execute('SELECT COUNT(*) FROM lab_tests')
count = cursor.fetchone()[0]

conn.commit()
conn.close()

print(f'✅ {count} lab tests in database')
print('✅ All lab tables ready!')