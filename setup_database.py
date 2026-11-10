import sqlite3

DB_NAME = 'hospital.db'


def create_all_tables():
    """Create all tables with the correct structure"""
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    # جدول کاربران
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            full_name TEXT NOT NULL,
            role TEXT NOT NULL,
            is_active INTEGER DEFAULT 1,
            last_login TIMESTAMP,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # جدول بیماران
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS patients (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            national_id TEXT UNIQUE NOT NULL,
            first_name TEXT NOT NULL,
            last_name TEXT NOT NULL,
            age INTEGER NOT NULL,
            gender TEXT NOT NULL,
            phone TEXT NOT NULL,
            email TEXT,
            address TEXT,
            blood_type TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # جدول پزشکان
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS doctors (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            national_id TEXT UNIQUE NOT NULL,
            first_name TEXT NOT NULL,
            last_name TEXT NOT NULL,
            specialty TEXT NOT NULL,
            phone TEXT NOT NULL,
            email TEXT,
            available_days TEXT,
            consultation_fee REAL
        )
    ''')

    # جدول نوبت‌ها
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS appointments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            patient_id INTEGER NOT NULL,
            doctor_id INTEGER NOT NULL,
            appointment_date TEXT NOT NULL,
            appointment_time TEXT NOT NULL,
            status TEXT DEFAULT 'Scheduled',
            notes TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (patient_id) REFERENCES patients(id),
            FOREIGN KEY (doctor_id) REFERENCES doctors(id)
        )
    ''')

    # جدول پرداخت‌ها
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            receipt_number TEXT UNIQUE NOT NULL,
            patient_id INTEGER NOT NULL,
            amount REAL NOT NULL,
            discount REAL DEFAULT 0,
            final_amount REAL NOT NULL,
            payment_method TEXT NOT NULL,
            payment_status TEXT DEFAULT 'Pending',
            payment_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            notes TEXT,
            FOREIGN KEY (patient_id) REFERENCES patients(id)
        )
    ''')

    # جدول خدمات
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS services (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            price REAL NOT NULL,
            description TEXT
        )
    ''')

    # جدول داروها
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS medicines (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            category TEXT,
            manufacturer TEXT,
            price REAL NOT NULL,
            stock_quantity INTEGER DEFAULT 0,
            min_stock INTEGER DEFAULT 10,
            expiry_date TEXT,
            description TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # ===== جدول فروش دارو با invoice_id =====
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS medicine_sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            medicine_id INTEGER NOT NULL,
            patient_id INTEGER NOT NULL,
            quantity INTEGER NOT NULL,
            unit_price REAL NOT NULL,
            total_price REAL NOT NULL,
            sale_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            invoice_id TEXT,
            notes TEXT,
            FOREIGN KEY (medicine_id) REFERENCES medicines(id),
            FOREIGN KEY (patient_id) REFERENCES patients(id)
        )
    ''')

    # اضافه کردن کاربر Admin پیش‌فرض
    import hashlib
    admin_password = hashlib.sha256('admin123'.encode()).hexdigest()
    try:
        cursor.execute('''
            INSERT INTO users (username, password_hash, full_name, role)
            VALUES (?, ?, ?, ?)
        ''', ('admin', admin_password, 'System Administrator', 'admin'))
    except sqlite3.IntegrityError:
        pass

    # اضافه کردن خدمات پیش‌فرض
    services = [
        ('Consultation', 500000, 'Doctor consultation fee'),
        ('Blood Test', 200000, 'Complete blood count'),
        ('X-Ray', 800000, 'Chest X-ray'),
        ('MRI', 3000000, 'Magnetic resonance imaging'),
        ('ECG', 600000, 'Electrocardiogram'),
        ('Ultrasound', 1200000, 'Ultrasound imaging'),
        ('Vaccination', 300000, 'Vaccine injection'),
        ('Dressing', 150000, 'Wound dressing'),
    ]
    for service in services:
        try:
            cursor.execute('INSERT INTO services (name, price, description) VALUES (?, ?, ?)', service)
        except sqlite3.IntegrityError:
            pass

    # اضافه کردن داروهای پیش‌فرض
    medicines = [
        ('Paracetamol 500mg', 'Painkiller', 'Tehran Pharma', 15000, 500, 50, '2027-06-30', 'For pain and fever'),
        ('Amoxicillin 500mg', 'Antibiotic', 'Pars Daru', 45000, 200, 30, '2027-03-15', 'Antibiotic'),
        ('Ibuprofen 400mg', 'Painkiller', 'Alborz Pharma', 25000, 300, 40, '2027-09-20', 'Anti-inflammatory'),
        ('Omeprazole 20mg', 'Antacid', 'Tehran Pharma', 35000, 150, 20, '2027-12-01', 'For stomach acid'),
        ('Metformin 500mg', 'Diabetes', 'Pars Daru', 30000, 400, 50, '2028-01-15', 'For diabetes'),
        ('Atorvastatin 20mg', 'Cholesterol', 'Alborz Pharma', 55000, 100, 25, '2027-08-10', 'For cholesterol'),
        ('Losartan 50mg', 'Blood Pressure', 'Tehran Pharma', 40000, 180, 30, '2027-11-20', 'For blood pressure'),
        ('Vitamin D3 1000IU', 'Supplement', 'Pars Daru', 20000, 600, 60, '2028-06-30', 'Vitamin supplement'),
        ('Iron Supplement', 'Supplement', 'Alborz Pharma', 18000, 350, 40, '2028-03-15', 'Iron supplement'),
        ('Cetirizine 10mg', 'Antihistamine', 'Tehran Pharma', 22000, 250, 30, '2027-10-05', 'For allergies'),
    ]
    for med in medicines:
        try:
            cursor.execute('''
                INSERT INTO medicines (name, category, manufacturer, price, stock_quantity, min_stock, expiry_date, description)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', med)
        except sqlite3.IntegrityError:
            pass

    conn.commit()
    conn.close()

    print('✅ Database setup completed!')
    print('📋 All tables created with correct structure')
    print('👤 Default user: admin / admin123')
    print('💊 Default medicines and services added')


if __name__ == '__main__':
    create_all_tables()
    print()
    print('🎉 Ready to use!')