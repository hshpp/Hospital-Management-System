<div align="center">

# 🏥 Hospital Management System (HMS)
### سیستم جامع مدیریت بیمارستان

<!-- Badges -->
![Python Version](https://badgen.net/badge/Python/3.10+/blue)
![Tkinter](https://badgen.net/badge/GUI/Tkinter/green)
![SQLite](https://badgen.net/badge/Database/SQLite/lightgrey)
![Status](https://badgen.net/badge/Status/Active/success)

</div>

---

## 🇬🇧 English Description

A comprehensive, fully-featured desktop application designed to manage all aspects of a clinic or hospital. Built with Python and Tkinter, it features a modern UI, complete patient lifecycle management, financial tracking, and inventory control.

### ✨ Key Features
*   **🔐 Role-Based Access Control:** Secure login with roles (Admin, Doctor, Receptionist, Pharmacist).
*   **📊 Dynamic Dashboard:** Real-time stats, 7-day revenue charts, Revenue Breakdown Pie Chart, Top Doctors/Medicines.
*   **👥 Patient & Doctor Management:** Full CRUD, patient timelines, and medical histories.
*   **📅 Appointment System:** Scheduling, reminders, and conflict prevention.
*   **📋 Medical Records:** EHR with vital signs, diagnoses, and smart prescription writer.
*   **🧪 Laboratory:** Direct & Prescription-based lab requests, result entry, linked to records.
*   **💊 Pharmacy:** Stock management, direct sales, dispensing, and low-stock alerts.
*   **💰 Billing & Reports:** Automated invoices. Detailed Daily/Weekly/Monthly/Yearly reports for Visits, Pharmacy, and Lab with **Export to Excel (with Charts)** and PDF Printing.
*   **💾 Backup & Security:** Automated backups, permanent archives, and system Audit Log.
*   **🎨 Modern UI:** Responsive design, rounded buttons, dark mode, and customized Tkinter interface.

### 🛠️ Tech Stack
*   **Language:** Python 3
*   **GUI Framework:** Tkinter
*   **Database:** SQLite3
*   **Data Visualization:** Matplotlib
*   **PDF Generation:** FPDF2, Arabic Reshaper, Python-Bidi
*   **Excel Export:** OpenPyXL

### 🚀 Installation & Usage

> ⚠️ **Note:** The final executable (`.exe`) is available. You can download it and run the app without installing Python.

#### Option 1: Run the Executable (Windows)
1. Download `HospitalMS.exe` from the Releases section.
2. Ensure `Vazirmatn-Regular.ttf` is in the same folder (if not bundled).
3. Double-click the `.exe` file to run.

#### Option 2: Run from Source (Developers)
1. **Clone the repository:**
   ```bash
   git clone https://github.com/hshpp/Hospital-Management-System.git
   cd Hospital-Management-System
   

2. Install dependencies:

pip install -r requirements.txt

3. Initialize the Database:

python setup_database.py

4. Run the application:

python hospital_gui_full.py

🔑 Default Credentials
Username: admin

Password: admin123
(⚠️ Please change this immediately after first login).

🇮🇷 توضیحات فارسی
یک نرم‌افزار دسکتاپ جامع برای مدیریت تمام بخش‌های یک کلینیک. این پروژه با پایتون و Tkinter ساخته شده و دارای رابط کاربری مدرن، مدیریت چرخه کامل بیمار، پیگیری مالی و کنترل موجودی است.

✨ ویژگی‌های کلیدی
🔐 کنترل دسترسی: سیستم ورود امن با نقش‌های مختلف.

📊 داشبورد پویا: آمار لحظه‌ای، نمودار درآمد ۷ روز اخیر، نمودار دایره‌ای سهم درآمدها، برترین پزشکان و داروها.

👥 مدیریت بیماران و پزشکان: عملیات کامل، خط زمانی بیمار.

📅 سیستم نوبت‌دهی: زمان‌بندی و یادآوری.

📋 پرونده‌های پزشکی: ثبت علائم حیاتی، تشخیص، نسخه‌نویسی هوشمند.

🧪 آزمایشگاه: تفکیک آزمایش آزاد و نسخه‌ای، ثبت نتیجه و اتصال به پرونده.

💊 داروخانه: مدیریت انبار، فروش آزاد، تحویل نسخه، هشدار موجودی.

💰 گزارش‌گیری مالی: صدور فاکتور. گزارش‌های تفصیلی روزانه/هفتگی/ماهانه/سالانه برای ویزیت، داروخانه و آزمایشگاه همراه با خروجی اکسل (دارای نمودار) و پرینت PDF.

💾 بکاپ و امنیت: بکاپ خودکار روزانه و سیستم ثبت لاگ (Audit Log).

🎨 رابط کاربری مدرن: واکنش‌گرا، دکمه‌های گرد، پشتیبانی از حالت تاریک.

🚀 نصب و راه‌اندازی
⚠️ نکته: نسخه نهایی و فایل نصبی (.exe) آماده شده است. نیازی به نصب پایتون برای اجرای آن نیست.

روش اول: اجرای فایل نصبی (ویندوز)
فایل HospitalMS.exe را دانلود کنید.

روی فایل دوبار کلیک کنید تا برنامه اجرا شود.

روش دوم: اجرای کدها (برای توسعه‌دهندگان)
۱. دریافت پروژه:

git clone https://github.com/hshpp/Hospital-Management-System.git

۲. نصب پیش‌نیازها:

pip install -r requirements.txt

۳. ساخت دیتابیس:

python setup_database.py

۴. اجرای برنامه:

python hospital_gui_full.py

🔑 اطلاعات ورود پیش‌فرض
نام کاربری: admin

رمز عبور: admin123

📸 Screenshots / تصاویر

<div align="center"> <img src="https://raw.githubusercontent.com/hshpp/Hospital-Management-System/main/login.png" width="48%" /> 
<img src="https://raw.githubusercontent.com/hshpp/Hospital-Management-System/main/dashboard.png" width="48%" /> <br><br> 
<img src="https://raw.githubusercontent.com/hshpp/Hospital-Management-System/main/medicalrecord.png" width="48%" /> 
<img src="https://raw.githubusercontent.com/hshpp/Hospital-Management-System/main/pharmacy.png" width="48%" /> </div> <br>

📄 License / مجوز
This project is licensed under the MIT License. / این پروژه تحت مجوز MIT منتشر شده است.

👨‍💻 Developer / توسعه‌دهنده
Hossein Shahabi / حسین شهابی

Email : spacex.nasa.starlink@gmail.com

GitHub: @hshpp

