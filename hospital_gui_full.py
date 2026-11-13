import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import sqlite3
import re
from fpdf.enums import XPos, YPos
# ==========================================
# Date Filter Helper
# ==========================================

from datetime import datetime, timedelta

import os
import shutil
from datetime import datetime

# ===== برای نمودارها =====
try:
    import matplotlib
    matplotlib.use('TkAgg')  # برای کار با Tkinter
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.figure import Figure
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False
    print('⚠️ matplotlib not installed. Charts will not work.')

DATE_FILTER_OPTIONS = [
    'All',
    'Today',
    'Yesterday',
    'Last 7 Days',
    'Last 30 Days',
    'This Month',
    'Last Month',
    'This Year',
    'Custom Range',
]


def get_date_range(filter_type, custom_from=None, custom_to=None):
    """
    برگرداندن بازه‌ی تاریخ بر اساس فیلتر

    Returns:
        (from_date, to_date) هر دو به صورت 'YYYY-MM-DD' یا (None, None)
    """
    today = datetime.now().date()

    if filter_type == 'All' or not filter_type:
        return None, None

    elif filter_type == 'Today':
        return today.strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d')

    elif filter_type == 'Yesterday':
        yesterday = today - timedelta(days=1)
        return yesterday.strftime('%Y-%m-%d'), yesterday.strftime('%Y-%m-%d')

    elif filter_type == 'Last 7 Days':
        from_date = today - timedelta(days=6)
        return from_date.strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d')

    elif filter_type == 'Last 30 Days':
        from_date = today - timedelta(days=29)
        return from_date.strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d')

    elif filter_type == 'This Month':
        from_date = today.replace(day=1)
        return from_date.strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d')

    elif filter_type == 'Last Month':
        first_of_this_month = today.replace(day=1)
        last_of_last_month = first_of_this_month - timedelta(days=1)
        first_of_last_month = last_of_last_month.replace(day=1)
        return (first_of_last_month.strftime('%Y-%m-%d'),
                last_of_last_month.strftime('%Y-%m-%d'))

    elif filter_type == 'This Year':
        from_date = today.replace(month=1, day=1)
        return from_date.strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d')

    elif filter_type == 'Custom Range':
        return custom_from, custom_to

    return None, None


def create_date_filter(parent, bg_color, on_change_callback):
    """
    ساخت یه Date Filter Widget

    Returns:
        (filter_var, from_var, to_var, widget_frame) - برای دسترسی از بیرون
    """
    container = tk.Frame(parent, bg=bg_color)

    tk.Label(container, text='📅 Date:',
             font=('Arial', 10, 'bold'),
             bg=bg_color, fg='#2c3e50').pack(side='left', padx=(0, 5))

    filter_var = tk.StringVar(value='All')
    combo = ttk.Combobox(container, textvariable=filter_var,
                         values=DATE_FILTER_OPTIONS,
                         state='readonly', width=14, font=('Arial', 10))
    combo.pack(side='left')

    # ===== Custom Range Frame (مخفی به صورت پیش‌فرض) =====
    custom_frame = tk.Frame(container, bg=bg_color)

    from_var = tk.StringVar()
    to_var = tk.StringVar()

    tk.Label(custom_frame, text='From:',
             font=('Arial', 9), bg=bg_color).pack(side='left', padx=(10, 3))

    from_entry = tk.Entry(custom_frame, textvariable=from_var,
                          font=('Arial', 9), width=11,
                          bg='white', relief='solid', bd=1)
    from_entry.pack(side='left', ipady=3)

    tk.Label(custom_frame, text='To:',
             font=('Arial', 9), bg=bg_color).pack(side='left', padx=(5, 3))

    to_entry = tk.Entry(custom_frame, textvariable=to_var,
                        font=('Arial', 9), width=11,
                        bg='white', relief='solid', bd=1)
    to_entry.pack(side='left', ipady=3)

    tk.Label(custom_frame, text='(YYYY-MM-DD)',
             font=('Arial', 8, 'italic'),
             bg=bg_color, fg='#95a5a6').pack(side='left', padx=5)

    # ===== رفتار Combo =====
    def on_combo_change(event=None):
        if filter_var.get() == 'Custom Range':
            custom_frame.pack(side='left', padx=(5, 0))
        else:
            custom_frame.pack_forget()
        on_change_callback()

    combo.bind('<<ComboboxSelected>>', on_combo_change)

    # ===== Auto-apply برای Custom =====
    def on_date_change(*args):
        if filter_var.get() == 'Custom Range':
            on_change_callback()

    from_var.trace_add('write', on_date_change)
    to_var.trace_add('write', on_date_change)

    return filter_var, from_var, to_var, container, custom_frame

# ===== برای پرینت =====
try:
    import win32print
    import win32api
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
    print('⚠️ pywin32 not installed. Print will fallback to PDF.')


DB_NAME = 'hospital.db'


# ==========================================
# Backup Functions
# ==========================================

import os
import shutil
from datetime import datetime

BACKUP_DIR = 'backups'
ARCHIVE_DIR = 'backups/archive'  # بکاپ‌های دائمی
MAX_BACKUPS = 100  # تعداد نسخه‌هایی که نگه داشته می‌شه (~3 ماه)
CONFIG_FILE = 'backup_config.json'  # فایل تنظیمات مسیر ثانویه
SETTINGS_FILE = 'settings.json'


def ensure_backup_dir():
    """پوشه‌ی backups رو بساز اگه وجود نداشت"""
    if not os.path.exists(BACKUP_DIR):
        os.makedirs(BACKUP_DIR)


def ensure_archive_dir():
    """پوشه‌ی archive رو بساز"""
    if not os.path.exists(ARCHIVE_DIR):
        os.makedirs(ARCHIVE_DIR)


def load_config():
    """تنظیمات رو از فایل بخون"""
    import json
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return {
        'secondary_path': '',  # مسیر ثانویه (هارد اکسترنال)
        'enabled': False
    }


def save_config(config):
    """تنظیمات رو ذخیره کن"""
    import json
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def create_backup_daily():
    """فقط اگه امروز بکاپ نگرفته، بگیر"""
    today = datetime.now().strftime('%Y%m%d')

    ensure_backup_dir()

    # چک کن امروز بکاپ گرفته یا نه
    existing = [
        f for f in os.listdir(BACKUP_DIR)
        if f.startswith(f'hospital_backup_{today}') and f.endswith('.db')
    ]

    if existing:
        return False, 'Already backed up today'

    return create_backup(auto=True)


def copy_to_secondary(backup_path):
    """کپی به مسیر ثانویه (هارد اکسترنال)"""
    try:
        config = load_config()

        if not config.get('enabled'):
            return False, 'Secondary backup disabled'

        secondary = config.get('secondary_path', '').strip()

        if not secondary:
            return False, 'No secondary path set'

        if not os.path.exists(secondary):
            return False, f'Secondary path not available: {secondary}'

        # کپی فایل
        filename = os.path.basename(backup_path)
        dest = os.path.join(secondary, filename)
        shutil.copy2(backup_path, dest)

        return True, dest

    except Exception as e:
        return False, str(e)


def create_backup(auto=False):
    """
    یه نسخه‌ی پشتیبان از دیتابیس بساز
    auto: اگه True باشه، پیام نمایش نمی‌ده
    """
    try:
        ensure_backup_dir()

        if not os.path.exists(DB_NAME):
            return False, '❌ Database file not found!'

        # اسم فایل با timestamp
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_name = f'hospital_backup_{timestamp}.db'
        backup_path = os.path.join(BACKUP_DIR, backup_name)

        # کپی فایل دیتابیس
        shutil.copy2(DB_NAME, backup_path)

        # پاک کردن نسخه‌های قدیمی (فقط از پوشه‌ی اصلی، نه archive)
        cleanup_old_backups()



        # کپی به مسیر ثانویه (اگه تنظیم شده)
        secondary_msg = ''
        sec_success, sec_result = copy_to_secondary(backup_path)
        if sec_success:
            secondary_msg = f'  |  📤 Copied to: {sec_result}'

        # اندازه فایل
        size_kb = os.path.getsize(backup_path) / 1024

        return True, f'{backup_path} ({size_kb:.1f} KB){secondary_msg}'

    except Exception as e:
        return False, str(e)


def cleanup_old_backups():
    """نگه‌داشتن فقط MAX_BACKUPS نسخه‌ی آخر (archive دست نخورده می‌مونه)"""
    try:
        if not os.path.exists(BACKUP_DIR):
            return

        # فقط فایل‌های hospital_backup_*.db (نه پوشه‌ی archive)
        backups = []
        for f in os.listdir(BACKUP_DIR):
            full_path = os.path.join(BACKUP_DIR, f)
            if (os.path.isfile(full_path) and
                    f.startswith('hospital_backup_') and
                    f.endswith('.db')):
                backups.append(f)

        # مرتب‌سازی نزولی (جدیدترین اول)
        backups.sort(reverse=True)

        # حذف قدیمی‌ها
        for old_file in backups[MAX_BACKUPS:]:
            try:
                os.remove(os.path.join(BACKUP_DIR, old_file))
                print(f'🗑️ Old backup deleted: {old_file}')
            except:
                pass

    except Exception as e:
        print(f'⚠️ Cleanup error: {e}')


def archive_backup(backup_path):
    """یه بکاپ رو به آرشیو دائمی منتقل کن"""
    try:
        ensure_archive_dir()

        filename = os.path.basename(backup_path)

        # جلوگیری از تکرار نام
        dest = os.path.join(ARCHIVE_DIR, filename)
        if os.path.exists(dest):
            timestamp = datetime.now().strftime('%H%M%S')
            name, ext = os.path.splitext(filename)
            dest = os.path.join(ARCHIVE_DIR, f'{name}_archived_{timestamp}{ext}')

        shutil.move(backup_path, dest)
        return True, dest

    except Exception as e:
        return False, str(e)


def get_archive_list():
    """لیست بکاپ‌های آرشیوی"""
    try:
        ensure_archive_dir()

        archives = [
            f for f in os.listdir(ARCHIVE_DIR)
            if f.endswith('.db')
        ]
        archives.sort(reverse=True)

        result = []
        for a in archives:
            path = os.path.join(ARCHIVE_DIR, a)
            size_kb = os.path.getsize(path) / 1024
            mtime = datetime.fromtimestamp(os.path.getmtime(path))
            result.append({
                'name': a,
                'path': path,
                'size_kb': size_kb,
                'time': mtime.strftime('%Y-%m-%d %H:%M:%S')
            })
        return result
    except Exception:
        return []


def get_backup_list():
    """لیست Backup های موجود رو برگردون"""
    try:
        ensure_backup_dir()

        backups = [
            f for f in os.listdir(BACKUP_DIR)
            if f.startswith('hospital_backup_') and f.endswith('.db')
        ]

        backups.sort(reverse=True)

        result = []
        for b in backups:
            path = os.path.join(BACKUP_DIR, b)
            size_kb = os.path.getsize(path) / 1024
            mtime = datetime.fromtimestamp(os.path.getmtime(path))
            result.append({
                'name': b,
                'path': path,
                'size_kb': size_kb,
                'time': mtime.strftime('%Y-%m-%d %H:%M:%S')
            })

        return result

    except Exception:
        return []


# ==========================================
# Theme Constants
# ==========================================

# ==========================================
# Theme
# ==========================================

try:
    from theme import get_theme, LIGHT_THEME, DARK_THEME
    THEME_AVAILABLE = True
except ImportError:
    THEME_AVAILABLE = False
    LIGHT_THEME = {'bg': '#f5f7fa', 'fg': '#2c3e50', 'bg_secondary': '#ffffff'}
    DARK_THEME = LIGHT_THEME
    def get_theme(mode='light'):
        return LIGHT_THEME

# ===== تم فعلی =====
CURRENT_THEME_MODE = 'light'  # این از settings.json خونده می‌شه
CURRENT_THEME = get_theme(CURRENT_THEME_MODE)

# ===== رنگ‌های اصلی (از تم) =====
WINDOW_BG = CURRENT_THEME['bg']
HEADER_BG = CURRENT_THEME['header_bg']
TEXT_FG = CURRENT_THEME['fg']
HEADER_BG = '#3b4a6b'

_BUTTON_COLORS = {
    'primary': '#3498db',
    'primary_dark': '#2980b9',
    'secondary': '#9b59b6',
    'secondary_dark': '#8e44ad',
    'success': '#27ae60',
    'success_dark': '#229954',
    'danger': '#e74c3c',
    'danger_dark': '#c0392b',
    'warning': '#f39c12',
    'warning_dark': '#d68910',
    'info': '#1abc9c',
    'info_dark': '#16a085',
}


# ==========================================
# Helper Functions
# ==========================================

def get_connection():
    conn = sqlite3.connect(DB_NAME, timeout=10)

    return conn


def log_action(action, entity_type=None, entity_id=None, details=None,
               user=None, silent=True):
    """
    ثبت یه عملیات توی audit_log

    action: 'CREATE', 'UPDATE', 'DELETE', 'LOGIN', 'LOGOUT', 'PRINT', 'BACKUP', 'RESTORE'
    entity_type: 'patient', 'doctor', 'appointment', 'payment', 'medicine', 'user', ...
    entity_id: ID رکورد
    details: توضیحات اضافی
    user: {'id': ..., 'username': ..., 'role': ...}
    silent: اگه True، خطا رو نمایش نده
    """
    try:
        conn = get_connection()
        cursor = conn.cursor()

        user_id = None
        username = None
        user_role = None

        if user:
            user_id = user.get('id')
            username = user.get('username')
            user_role = user.get('role')

        cursor.execute('''
            INSERT INTO audit_log
            (user_id, username, user_role, action, entity_type, entity_id, details)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (user_id, username, user_role, action, entity_type, entity_id, details))

        conn.commit()
        conn.close()
        return True
    except Exception as e:
        if not silent:
            print(f'⚠️ Audit log error: {e}')
        return False



def is_valid_name(name):
    pattern = r'^[A-Za-z\s\u0600-\u06FF]+$'
    return bool(re.match(pattern, name))


def is_valid_phone(phone):
    pattern = r'^09\d{9}$'
    return bool(re.match(pattern, phone))


def is_valid_national_id(nid):
    return nid.isdigit() and len(nid) == 10


def is_valid_email(email):
    if not email:
        return True
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email))


def get_stats():
    """گرفتن آمار کلی سیستم (شامل درآمد کل)"""
    try:
        conn = get_connection()
        cursor = conn.cursor()

        cursor.execute('SELECT COUNT(*) FROM patients')
        patients = cursor.fetchone()[0] or 0

        cursor.execute('SELECT COUNT(*) FROM doctors')
        doctors = cursor.fetchone()[0] or 0

        cursor.execute("SELECT COUNT(*) FROM appointments WHERE status = 'Scheduled'")
        scheduled = cursor.fetchone()[0] or 0

        # ✅ محاسبه درآمد کل (ویزیت + داروخانه + آزمایشگاه)
        cursor.execute('SELECT COALESCE(SUM(final_amount), 0) FROM payments WHERE payment_status = "Paid"')
        visit_rev = cursor.fetchone()[0]

        cursor.execute('SELECT COALESCE(SUM(final_amount), 0) FROM pharmacy_payments WHERE payment_status = "Paid"')
        pharmacy_rev = cursor.fetchone()[0]

        cursor.execute('SELECT COALESCE(SUM(final_amount), 0) FROM lab_payments WHERE payment_status = "Paid"')
        lab_rev = cursor.fetchone()[0]

        total_revenue = visit_rev + pharmacy_rev + lab_rev

        conn.close()

        return {
            'patients': patients,
            'doctors': doctors,
            'scheduled': scheduled,
            'total_revenue': total_revenue # ✅ فیلد جدید
        }

    except Exception as e:
        print(f'❌ get_stats error: {e}')
        return {'patients': 0, 'doctors': 0, 'scheduled': 0, 'total_revenue': 0}



def print_text_file(text_content, filename='temp_print.txt', open_pdf=True):
    """Create a PDF (with Persian support) and optionally open it"""
    import os
    import tempfile

    try:
        from fpdf import FPDF
        import arabic_reshaper
        from bidi.algorithm import get_display

        temp_dir = tempfile.gettempdir()
        pdf_filename = filename.replace('.txt', '.pdf')
        pdf_path = os.path.join(temp_dir, pdf_filename)

        pdf = FPDF()
        pdf.add_page()

        # ===== اضافه کردن فونت فارسی =====
        font_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                  'Vazirmatn-Regular.ttf')

        if not os.path.exists(font_path):
            return False, f'❌ Font not found: {font_path}'

        pdf.add_font('Vazir', '', font_path)
        pdf.set_font('Vazir', size=10)

        # ===== تابع کمکی برای اصلاح متن فارسی =====
        def fix_persian(text):
            try:
                reshaped = arabic_reshaper.reshape(text)
                return get_display(reshaped)
            except Exception:
                return text

        def has_persian(text):
            return any('\u0600' <= c <= '\u06FF' for c in text)

        # ===== حذف ایموجی‌ها =====
        import re as _re

        def remove_emoji(text):
            emoji_pattern = _re.compile(
                '['
                '\U0001F300-\U0001F9FF'
                '\U0001FA00-\U0001FAFF'
                '\U00002600-\U000027BF'
                '\U0001F000-\U0001F2FF'
                '\U0001F600-\U0001F64F'
                '\U0001F680-\U0001F6FF'
                '\U0001F1E0-\U0001F1FF'
                '\u2190-\u21FF'
                '\u2700-\u27BF'
                '\uFE00-\uFE0F'
                ']+',
                flags=_re.UNICODE
            )
            return emoji_pattern.sub('', text)

        # ===== پردازش خط به خط =====
        for line in text_content.split('\n'):
            clean_line = remove_emoji(line)

            if has_persian(clean_line):
                fixed_line = fix_persian(clean_line)
                pdf.cell(0, 5, fixed_line, new_x=XPos.LMARGIN, new_y=YPos.NEXT)
            else:
                pdf.cell(0, 5, clean_line, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

        # ===== ذخیره PDF =====
        pdf.output(pdf_path)

        # ===== باز کردن فایل (اختیاری) =====
        if open_pdf:
            if os.name == 'nt':
                try:
                    os.startfile(pdf_path)
                except OSError:
                    import subprocess
                    subprocess.run(['cmd', '/c', 'start', '', pdf_path], check=False)
            else:
                import subprocess
                subprocess.run(['xdg-open', pdf_path], check=False)

        return True, pdf_path

    except Exception as e:
        return False, str(e)



# ==========================================
# Real Printer Functions
# ==========================================

def get_available_printers():
    """لیست پرینترهای نصب‌شده روی سیستم"""
    if not WIN32_AVAILABLE:
        return []
    try:
        printers = []
        for printer in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS,
            None, 1
        ):
            # printer[2] = نام پرینتر
            printers.append(printer[2])
        return printers
    except Exception as e:
        print(f'⚠️ Error getting printers: {e}')
        return []


def get_default_printer():
    """پرینتر پیش‌فرض ویندوز"""
    if not WIN32_AVAILABLE:
        return None
    try:
        return win32print.GetDefaultPrinter()
    except Exception:
        return None


def print_pdf_to_printer(pdf_path, printer_name=None):
    """ارسال PDF مستقیم به پرینتر"""
    if not WIN32_AVAILABLE:
        return False, 'pywin32 not installed'

    if not os.path.exists(pdf_path):
        return False, f'PDF file not found: {pdf_path}'

    try:
        if not printer_name:
            printer_name = win32print.GetDefaultPrinter()

        if not printer_name:
            return False, 'No default printer found'

        # ===== تلاش برای ارسال =====
        win32api.ShellExecute(
            0,
            "printto",
            pdf_path,
            f'"{printer_name}"',
            ".",
            0
        )

        return True, printer_name

    except Exception as e:
        error_msg = str(e)

        # ===== اگه پرینتر مجازی بود =====
        if 'Microsoft Print to PDF' in (printer_name or '') or \
           'XPS' in (printer_name or '') or \
           'OneNote' in (printer_name or ''):
            return False, f'VIRTUAL_PRINTER: {printer_name}'

        return False, error_msg


def print_multiple_pdfs(pdf_paths, printer_name=None):
    """
    چاپ چند PDF به ترتیب

    Args:
        pdf_paths: لیست مسیرهای PDF
        printer_name: نام پرینتر (اگه None، پیش‌فرض)

    Returns:
        (success_count, error_count, errors)
    """
    if not WIN32_AVAILABLE:
        return 0, len(pdf_paths), ['pywin32 not installed']

    success_count = 0
    error_count = 0
    errors = []

    for i, pdf_path in enumerate(pdf_paths, 1):
        try:
            if not os.path.exists(pdf_path):
                error_count += 1
                errors.append(f'{i}. File not found: {pdf_path}')
                continue

            ok, msg = print_pdf_to_printer(pdf_path, printer_name)

            if ok:
                success_count += 1
                # کمی تاخیر بین چاپ‌ها (خیلی مهم برای پرینتر!)
                import time
                time.sleep(0.5)
            else:
                error_count += 1
                errors.append(f'{i}. {msg}')

        except Exception as e:
            error_count += 1
            errors.append(f'{i}. {str(e)}')

    return success_count, error_count, errors


def ask_printer(parent=None):
    """پنجره‌ی انتخاب پرینتر"""
    printers = get_available_printers()

    if not printers:
        messagebox.showerror(
            '❌ No Printers',
            'No printers found on this system!\n\n'
            'Please install a printer first or use "Save as PDF" option.'
        )
        return None

    # ===== پنجره =====
    win = tk.Toplevel(parent) if parent else tk.Tk()
    win.title('🖨️ Select Printer')
    win.geometry('450x400')
    win.configure(bg='#ecf0f1')
    win.resizable(False, False)

    win.update_idletasks()
    w, h = 450, 400
    x = (win.winfo_screenwidth() // 2) - (w // 2)
    y = (win.winfo_screenheight() // 2) - (h // 2)
    win.geometry(f'{w}x{h}+{x}+{y}')

    # Header
    tk.Label(win, text='🖨️', font=('Arial', 30),
             bg='#ecf0f1').pack(pady=(15, 5))
    tk.Label(win, text='Select Printer',
             font=('Arial', 14, 'bold'), bg='#ecf0f1',
             fg='#2c3e50').pack()
    tk.Label(win, text='Choose which printer to send the document to',
             font=('Arial', 9, 'italic'), bg='#ecf0f1',
             fg='#7f8c8d').pack(pady=(5, 15))

    default_printer = get_default_printer()

    # Radio buttons
    printer_var = tk.StringVar(value=default_printer or printers[0])

    list_frame = tk.Frame(win, bg='white', relief='solid', bd=1)
    list_frame.pack(fill='both', expand=True, padx=20, pady=10)

    for p in printers:
        icon = '⭐' if p == default_printer else '🖨️'
        label = f'{icon} {p}'
        if p == default_printer:
            label += '  (Default)'

        tk.Radiobutton(list_frame, text=label,
                       variable=printer_var, value=p,
                       bg='white', font=('Arial', 10),
                       anchor='w', padx=10, pady=5).pack(fill='x')

    result = {'printer': None}

    def do_print():
        result['printer'] = printer_var.get()
        win.destroy()

    def do_cancel():
        result['printer'] = None
        win.destroy()

    # Buttons
    btn_frame = tk.Frame(win, bg='#ecf0f1')
    btn_frame.pack(pady=15)

    tk.Button(btn_frame, text='🖨️ Print', command=do_print,
              bg='#27ae60', fg='white', font=('Arial', 11, 'bold'),
              relief='flat', cursor='hand2',
              padx=25, pady=8).pack(side='left', padx=5)

    tk.Button(btn_frame, text='❌ Cancel', command=do_cancel,
              bg='#e74c3c', fg='white', font=('Arial', 11, 'bold'),
              relief='flat', cursor='hand2',
              padx=25, pady=8).pack(side='left', padx=5)

    # Double-click to print
    def on_double_click(event):
        do_print()

    for child in list_frame.winfo_children():
        child.bind('<Double-Button-1>', on_double_click)

    win.grab_set()
    win.wait_window()

    return result['printer']


# ==========================================
# UI Helpers
# ==========================================

def apply_modern_notebook_style():
    """Apply modern style to ttk.Notebook"""
    style = ttk.Style()
    style.theme_use('clam')

    style.configure('Modern.TNotebook',
                     background='#eef1f7',
                     borderwidth=0,
                     tabmargins=[2, 5, 2, 0])

    style.configure('Modern.TNotebook.Tab',
                     background='#d5dae5',
                     foreground='#3b4a6b',
                     padding=[8, 6],           # ← کوچیک‌تر
                     font=('Arial', 9, 'bold'), # ← کوچیک‌تر
                     borderwidth=0)

    style.map('Modern.TNotebook.Tab',
              background=[('selected', '#667eea'),
                          ('active', '#8b9dc3')],
              foreground=[('selected', 'white'),
                          ('active', 'white')])

def apply_dark_style():
    """اعمال استایل تیره یا روشن برای ttk widget ها"""
    try:
        style = ttk.Style()
        style.theme_use('clam')

        if CURRENT_THEME_MODE == 'dark':
            # ===== تم تیره =====
            style.configure('Treeview',
                            background=CURRENT_THEME['tree_bg'],
                            foreground=CURRENT_THEME['tree_fg'],
                            fieldbackground=CURRENT_THEME['tree_bg'],
                            borderwidth=0)
            style.configure('Treeview.Heading',
                            background=CURRENT_THEME['header_bg'],
                            foreground='white',
                            font=('Arial', 10, 'bold'))
            style.map('Treeview',
                      background=[('selected', CURRENT_THEME['tree_selected'])])

            style.configure('Modern.TNotebook',
                            background=CURRENT_THEME['notebook_bg'])
            style.configure('Modern.TNotebook.Tab',
                            background=CURRENT_THEME['notebook_tab'],
                            foreground=CURRENT_THEME['notebook_tab_fg'])

            style.configure('TCombobox',
                            fieldbackground=CURRENT_THEME['entry_bg'],
                            foreground=CURRENT_THEME['fg'])

            style.configure('TEntry',
                            fieldbackground=CURRENT_THEME['entry_bg'],
                            foreground=CURRENT_THEME['fg'])

            style.configure('Vertical.TScrollbar',
                            background=CURRENT_THEME['bg_secondary'])
        else:
            # ===== تم روشن (پیش‌فرض) =====
            style.configure('Treeview',
                            background='white',
                            foreground='#2c3e50',
                            fieldbackground='white',
                            borderwidth=0)
            style.configure('Treeview.Heading',
                            background='#3b4a6b',
                            foreground='white',
                            font=('Arial', 10, 'bold'))
            style.map('Treeview',
                      background=[('selected', '#3498db')])

            style.configure('Modern.TNotebook',
                            background='#eef1f7')
            style.configure('Modern.TNotebook.Tab',
                            background='#d5dae5',
                            foreground='#3b4a6b')

            style.configure('TCombobox',
                            fieldbackground='white',
                            foreground='#2c3e50')

            style.configure('TEntry',
                            fieldbackground='white',
                            foreground='#2c3e50')

    except Exception as e:
        print(f'⚠️ Style error: {e}')


def create_standard_button(parent, text, command, color='primary',
                            width=140, height=42, font=('Arial', 10, 'bold')):
    """Create a standardized rounded button"""
    bg_color = _BUTTON_COLORS.get(color, _BUTTON_COLORS['primary'])
    hover_color = _BUTTON_COLORS.get(color + '_dark', bg_color)

    canvas = tk.Canvas(parent, width=width, height=height,
                        bg=parent['bg'], highlightthickness=0, bd=0)
    canvas.pack(side='left', padx=5, pady=2)

    def round_rect(x1, y1, x2, y2, r, **kwargs):
        points = [
            x1 + r, y1, x2 - r, y1, x2, y1,
            x2, y1 + r, x2, y2 - r, x2, y2,
            x2 - r, y2, x1 + r, y2, x1, y2,
            x1, y2 - r, x1, y1 + r, x1, y1,
        ]
        return canvas.create_polygon(points, **kwargs, smooth=True)

    rect = round_rect(2, 2, width - 2, height - 2, 15,
                       fill=bg_color, outline=bg_color, tags='btn')

    canvas.create_text(width / 2, height / 2, text=text,
                        fill='white', font=font, tags='btn')

    def on_enter(e):
        canvas.itemconfig(rect, fill=hover_color, outline=hover_color)

    def on_leave(e):
        canvas.itemconfig(rect, fill=bg_color, outline=bg_color)

    def on_click(e):
        command()

    canvas.tag_bind('btn', '<Enter>', on_enter)
    canvas.tag_bind('btn', '<Leave>', on_leave)
    canvas.tag_bind('btn', '<Button-1>', on_click)
    canvas.config(cursor='hand2')

    return canvas

def create_button_bar(parent, bg='#f5f7fa'):
    """Create a standardized button bar container (centered)"""
    # Frame بیرونی که تمام عرض رو پر می‌کنه
    outer = tk.Frame(parent, bg=bg)
    outer.pack(pady=15, padx=40, fill='x')

    # Frame داخلی که وسط‌چین می‌شه
    inner = tk.Frame(outer, bg=bg)
    inner.pack(anchor='center')

    return inner


def create_rounded_button(parent, text, command, bg_color, hover_color,
                            text_color='white', width=200, height=50,
                            corner_radius=20, font=('Arial', 12, 'bold')):
    """Create a rounded button using Canvas"""
    canvas = tk.Canvas(parent, width=width, height=height,
                        bg=parent['bg'], highlightthickness=0)
    canvas.pack(side='left', padx=8)

    def round_rect(x1, y1, x2, y2, r, **kwargs):
        points = [
            x1 + r, y1, x2 - r, y1, x2, y1,
            x2, y1 + r, x2, y2 - r, x2, y2,
            x2 - r, y2, x1 + r, y2, x1, y2,
            x1, y2 - r, x1, y1 + r, x1, y1,
        ]
        return canvas.create_polygon(points, **kwargs, smooth=True)

    rect = round_rect(2, 2, width - 2, height - 2, corner_radius,
                       fill=bg_color, outline=bg_color, tags='btn')

    canvas.create_text(width / 2, height / 2, text=text,
                        fill=text_color, font=font, tags='btn')

    def on_enter(e):
        canvas.itemconfig(rect, fill=hover_color, outline=hover_color)

    def on_leave(e):
        canvas.itemconfig(rect, fill=bg_color, outline=bg_color)

    def on_click(e):
        command()

    canvas.tag_bind('btn', '<Enter>', on_enter)
    canvas.tag_bind('btn', '<Leave>', on_leave)
    canvas.tag_bind('btn', '<Button-1>', on_click)
    canvas.config(cursor='hand2')

    return canvas


def create_rounded_stat_card(parent, title, value, bg_color='#3b4a6b', width=180, height=90):
    """ساخت کارت آماری با گوشه‌های گرد و افکت هاور"""
    canvas = tk.Canvas(parent, width=width, height=height, bg=parent['bg'], highlightthickness=0)

    def round_rect(x1, y1, x2, y2, r, **kwargs):
        points = [
            x1 + r, y1, x2 - r, y1, x2, y1,
            x2, y1 + r, x2, y2 - r, x2, y2,
            x2 - r, y2, x1 + r, y2, x1, y2,
            x1, y2 - r, x1, y1 + r, x1, y1,
        ]
        return canvas.create_polygon(points, **kwargs, smooth=True)

    round_rect(2, 2, width - 2, height - 2, 15, fill=bg_color, outline=bg_color, tags='card')

    canvas.create_text(width / 2, height / 2 - 12, text=title, font=('Arial', 9, 'bold'), fill='white', tags='card')
    val_id = canvas.create_text(width / 2, height / 2 + 15, text=str(value), font=('Arial', 14, 'bold'), fill='white',
                                tags='card')

    # تابعی برای آپدیت مقدار
    def update_value(new_val):
        canvas.itemconfig(val_id, text=new_val)

    return canvas, update_value


def create_popup_buttons(parent, buttons_config, bg='#ecf0f1'):
    """
    ساخت ردیف دکمه‌های گِرد برای پنجره‌های Popup

    buttons_config: لیست دیکشنری
    [
        {'text': '💾 Save', 'command': save_func, 'color': 'success'},
        {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger'},
    ]
    """
    frame = tk.Frame(parent, bg=bg)
    frame.pack(pady=15)

    for btn in buttons_config:
        color = btn.get('color', 'primary')
        bg_color = _BUTTON_COLORS.get(color, _BUTTON_COLORS['primary'])
        hover_color = _BUTTON_COLORS.get(color + '_dark', bg_color)

        create_rounded_button(
            frame,
            text=btn['text'],
            command=btn['command'],
            bg_color=bg_color,
            hover_color=hover_color,
            width=btn.get('width', 160),
            height=btn.get('height', 45),
            corner_radius=15,
            font=('Arial', 11, 'bold')
        )

    return frame


# ==========================================
# Excel Export Helper
# ==========================================

def export_to_excel(headers, rows, filename, sheet_name='Sheet1', title=None):
    """
    خروجی به فایل اکسل با فرمت زیبا

    Args:
        headers: لیست عنوان‌ها (مثلاً ['ID', 'Name', 'Age'])
        rows: لیست ردیف‌ها (مثلاً [(1, 'Ali', 30), (2, 'Sara', 25)])
        filename: مسیر فایل ذخیره
        sheet_name: نام شیت
        title: عنوان بالای جدول (اختیاری)
    """
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name

        # ===== رنگ‌ها =====
        header_color = '2C3E50'  # سرمه‌ای
        title_color = '3498DB'  # آبی
        alt_row_color = 'F8F9FA'  # طوسی روشن
        border_color = 'BDC3C7'

        # ===== تنظیمات Border =====
        thin_border = Border(
            left=Side(style='thin', color=border_color),
            right=Side(style='thin', color=border_color),
            top=Side(style='thin', color=border_color),
            bottom=Side(style='thin', color=border_color)
        )

        current_row = 1

        # ===== عنوان (اگه هست) =====
        if title:
            ws.merge_cells(start_row=1, start_column=1,
                           end_row=1, end_column=len(headers))
            title_cell = ws.cell(row=1, column=1, value=title)
            title_cell.font = Font(size=14, bold=True, color='FFFFFF')
            title_cell.fill = PatternFill(start_color=title_color,
                                          end_color=title_color,
                                          fill_type='solid')
            title_cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[1].height = 30

            # تاریخ تولید
            ws.merge_cells(start_row=2, start_column=1,
                           end_row=2, end_column=len(headers))
            date_cell = ws.cell(row=2, column=1,
                                value=f'Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
            date_cell.font = Font(size=10, italic=True, color='7F8C8D')
            date_cell.alignment = Alignment(horizontal='center')

            current_row = 4

        # ===== Header =====
        header_row_num = current_row
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=header_row_num, column=col_idx, value=header)
            cell.font = Font(size=11, bold=True, color='FFFFFF')
            cell.fill = PatternFill(start_color=header_color,
                                    end_color=header_color,
                                    fill_type='solid')
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
        ws.row_dimensions[header_row_num].height = 25

        # ===== داده‌ها =====
        for row_idx, row_data in enumerate(rows, 1):
            excel_row = header_row_num + row_idx
            row_color = alt_row_color if row_idx % 2 == 0 else 'FFFFFF'

            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=excel_row, column=col_idx, value=value)
                cell.font = Font(size=10)
                cell.fill = PatternFill(start_color=row_color,
                                        end_color=row_color,
                                        fill_type='solid')
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border

        # ===== تنظیم عرض ستون‌ها (خودکار) =====
        for col_idx in range(1, len(headers) + 1):
            max_length = len(str(headers[col_idx - 1]))
            for row in rows:
                if col_idx - 1 < len(row):
                    max_length = max(max_length, len(str(row[col_idx - 1])))
            column_letter = get_column_letter(col_idx)
            ws.column_dimensions[column_letter].width = max_length + 4

        # ===== Freeze Header =====
        ws.freeze_panes = f'A{header_row_num + 1}'

        # ===== ذخیره =====
        wb.save(filename)
        return True, filename

    except ImportError:
        return False, 'openpyxl not installed. Run: pip install openpyxl'
    except Exception as e:
        return False, str(e)


def enable_multi_click_selection(treeview, on_selection_change=None):
    """
    فعال‌سازی انتخاب چندتایی با کلیک معمولی

    کاربر می‌تونه روی هر ردیف کلیک کنه تا به انتخاب اضافه/کم بشه
    """

    def on_click(event):
        # پیدا کردن ردیفی که کلیک شده
        item = treeview.identify_row(event.y)
        if not item:
            return

        # اگه قبلاً انتخاب شده بود، حذفش کن
        if item in treeview.selection():
            treeview.selection_remove(item)
        else:
            # اضافه کن به انتخاب
            treeview.selection_add(item)

        # جلوگیری از رفتار پیش‌فرض
        if on_selection_change:
            on_selection_change()

        return 'break'

    treeview.bind('<Button-1>', on_click)


# ==========================================
# Settings Functions
# ==========================================

def load_settings():
    """بارگذاری تنظیمات از فایل"""
    import json

    default_settings = {
        'printers': {
            'default': '',
            'prescription': '',
            'invoice': '',
            'report': '',
            'receipt': '',
        },
        'dark_mode': False,  # ← این خط رو اضافه کنید
    }
    # ... بقیه کد



    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # merge با default (برای مقادیر جدید)
                for key in default_settings:
                    if key not in data:
                        data[key] = default_settings[key]
                    elif isinstance(default_settings[key], dict):
                        for subkey in default_settings[key]:
                            if subkey not in data[key]:
                                data[key][subkey] = default_settings[key][subkey]
                return data
        except Exception as e:
            print(f'⚠️ Settings load error: {e}')

    return default_settings


def save_settings(settings):
    """ذخیره تنظیمات"""
    import json
    with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


# ==========================================
# Archive Functions
# ==========================================

# def get_archive_stats():
#     """گرفتن آمار Archive"""
#     conn = get_connection()
#     cursor = conn.cursor()
#
#     # Active records
#     cursor.execute('SELECT COUNT(*) FROM medical_records')
#     active_records = cursor.fetchone()[0]
#
#     cursor.execute('SELECT COUNT(*) FROM prescriptions')
#     active_prescriptions = cursor.fetchone()[0]
#
#     # Archived records
#     cursor.execute('SELECT COUNT(*) FROM archive_medical_records')
#     archived_records = cursor.fetchone()[0]
#
#     cursor.execute('SELECT COUNT(*) FROM archive_prescriptions')
#     archived_prescriptions = cursor.fetchone()[0]
#
#     # Oldest active date
#     cursor.execute('SELECT MIN(record_date) FROM medical_records')
#     oldest_date = cursor.fetchone()[0]
#
#     # Date range of archive
#     cursor.execute('SELECT MIN(record_date), MAX(record_date) FROM archive_medical_records')
#     archive_range = cursor.fetchone()
#
#     # DB file size
#     import os
#     db_size_mb = 0
#     if os.path.exists(DB_NAME):
#         db_size_mb = os.path.getsize(DB_NAME) / (1024 * 1024)
#
#     conn.close()
#
#     return {
#         'active_records': active_records,
#         'active_prescriptions': active_prescriptions,
#         'archived_records': archived_records,
#         'archived_prescriptions': archived_prescriptions,
#         'oldest_active': oldest_date,
#         'archive_from': archive_range[0] if archive_range else None,
#         'archive_to': archive_range[1] if archive_range else None,
#         'db_size_mb': db_size_mb,
#     }


# def count_records_to_archive(months):
#     """شمارش رکوردهایی که باید آرشیو بشن"""
#     from datetime import datetime, timedelta
#
#     cutoff = (datetime.now() - timedelta(days=months * 30)).strftime('%Y-%m-%d')
#
#     conn = get_connection()
#     cursor = conn.cursor()
#     cursor.execute('SELECT COUNT(*) FROM medical_records WHERE record_date < ?', (cutoff,))
#     count = cursor.fetchone()[0]
#     conn.close()
#
#     return count, cutoff


# def archive_records(months, archived_by='system', progress_callback=None):
#     """
#     آرشیو کردن رکوردهای قدیمی
#
#     Args:
#         months: تعداد ماه (رکوردهای قدیمی‌تر از این ماه آرشیو می‌شن)
#         archived_by: نام کاربر
#         progress_callback: تابع برای آپدیت progress
#
#     Returns:
#         (success, message, records_count)
#     """
#     from datetime import datetime, timedelta
#
#     try:
#         cutoff = (datetime.now() - timedelta(days=months * 30)).strftime('%Y-%m-%d')
#
#         conn = get_connection()
#         cursor = conn.cursor()
#
#         # ===== ۱. پیدا کردن رکوردهای قدیمی =====
#         cursor.execute('''
#             SELECT
#                 mr.id, mr.patient_id, mr.doctor_id, mr.record_date,
#                 mr.chief_complaint, mr.diagnosis, mr.symptoms,
#                 mr.blood_pressure, mr.temperature, mr.weight, mr.height,
#                 mr.notes, mr.created_at,
#                 p.first_name || ' ' || p.last_name AS patient_name,
#                 d.first_name || ' ' || d.last_name AS doctor_name
#             FROM medical_records mr
#             LEFT JOIN patients p ON mr.patient_id = p.id
#             LEFT JOIN doctors d ON mr.doctor_id = d.id
#             WHERE mr.record_date < ?
#             ORDER BY mr.record_date
#         ''', (cutoff,))
#         old_records = cursor.fetchall()
#
#         if not old_records:
#             conn.close()
#             return False, f'No records older than {months} months', 0
#
#         record_ids = [r[0] for r in old_records]
#         total = len(record_ids)
#
#         if progress_callback:
#             progress_callback(0, total, 'Copying records...')
#
#         # ===== ۲. کپی به Archive =====
#         for idx, record in enumerate(old_records, 1):
#             cursor.execute('''
#                 INSERT INTO archive_medical_records
#                 (original_id, patient_id, doctor_id, record_date,
#                  chief_complaint, diagnosis, symptoms, blood_pressure,
#                  temperature, weight, height, notes, patient_name,
#                  doctor_name, created_at, archived_by)
#                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
#             ''', (record[0], record[1], record[2], record[3],
#                   record[4], record[5], record[6], record[7],
#                   record[8], record[9], record[10], record[11],
#                   record[13], record[14], record[12], archived_by))
#
#             if progress_callback and idx % 10 == 0:
#                 progress_callback(idx, total, f'Archiving record {idx}/{total}...')
#
#         # ===== ۳. کپی prescriptions =====
#         if progress_callback:
#             progress_callback(total, total, 'Archiving prescriptions...')
#
#         placeholders = ','.join('?' * len(record_ids))
#         cursor.execute(f'''
#             SELECT
#                 pr.id, pr.record_id, pr.medicine_id,
#                 m.name, pr.dosage, pr.frequency, pr.duration,
#                 pr.quantity, pr.instructions
#             FROM prescriptions pr
#             LEFT JOIN medicines m ON pr.medicine_id = m.id
#             WHERE pr.record_id IN ({placeholders})
#         ''', record_ids)
#         old_prescriptions = cursor.fetchall()
#
#         for rx in old_prescriptions:
#             cursor.execute('''
#                 INSERT INTO archive_prescriptions
#                 (original_id, original_record_id, medicine_id, medicine_name,
#                  dosage, frequency, duration, quantity, instructions)
#                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
#             ''', (rx[0], rx[1], rx[2], rx[3], rx[4], rx[5], rx[6], rx[7], rx[8]))
#
#         # ===== ۴. حذف از فعال =====
#         if progress_callback:
#             progress_callback(total, total, 'Removing from active...')
#
#         cursor.execute(f'DELETE FROM prescriptions WHERE record_id IN ({placeholders})', record_ids)
#         cursor.execute(f'DELETE FROM medical_records WHERE id IN ({placeholders})', record_ids)
#
#         conn.commit()
#         conn.close()
#
#         return True, f'Archived {total} records + {len(old_prescriptions)} prescriptions', total
#
#     except Exception as e:
#         return False, str(e), 0


# def restore_from_archive(date_from, date_to):
#     """بازگردانی از Archive"""
#     try:
#         conn = get_connection()
#         cursor = conn.cursor()
#
#         # ===== ۱. پیدا کردن رکوردهای Archive =====
#         cursor.execute('''
#             SELECT * FROM archive_medical_records
#             WHERE record_date BETWEEN ? AND ?
#             ORDER BY record_date
#         ''', (date_from, date_to))
#         archived = cursor.fetchall()
#
#         if not archived:
#             conn.close()
#             return False, f'No archived records between {date_from} and {date_to}', 0
#
#         # ===== ۲. برگرداندن به active =====
#         restored_count = 0
#         for row in archived:
#             try:
#                 cursor.execute('''
#                     INSERT INTO medical_records
#                     (patient_id, doctor_id, record_date, chief_complaint,
#                      diagnosis, symptoms, blood_pressure, temperature,
#                      weight, height, notes, created_at)
#                     VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
#                 ''', (row[2], row[3], row[4], row[5], row[6], row[7],
#                       row[8], row[9], row[10], row[11], row[12], row[16]))
#
#                 new_record_id = cursor.lastrowid
#
#                 # برگرداندن prescriptions
#                 cursor.execute('''
#                     SELECT * FROM archive_prescriptions
#                     WHERE original_record_id = ?
#                 ''', (row[1],))
#                 archived_rx = cursor.fetchall()
#
#                 for rx in archived_rx:
#                     cursor.execute('''
#                         INSERT INTO prescriptions
#                         (record_id, medicine_id, dosage, frequency,
#                          duration, quantity, instructions)
#                         VALUES (?, ?, ?, ?, ?, ?, ?)
#                     ''', (new_record_id, rx[3], rx[5], rx[6], rx[7], rx[8], rx[9]))
#
#                 # حذف از archive
#                 cursor.execute('DELETE FROM archive_prescriptions WHERE original_record_id = ?', (row[1],))
#                 cursor.execute('DELETE FROM archive_medical_records WHERE id = ?', (row[0],))
#
#                 restored_count += 1
#
#             except Exception:
#                 continue
#
#         conn.commit()
#         conn.close()
#
#         return True, f'Restored {restored_count} records', restored_count
#
#     except Exception as e:
#         return False, str(e), 0


# def delete_archive_permanent(date_from=None, date_to=None):
#     """حذف دائمی از Archive"""
#     try:
#         conn = get_connection()
#         cursor = conn.cursor()
#
#         if date_from and date_to:
#             cursor.execute('SELECT COUNT(*) FROM archive_medical_records WHERE record_date BETWEEN ? AND ?',
#                            (date_from, date_to))
#             count = cursor.fetchone()[0]
#
#             if count == 0:
#                 conn.close()
#                 return False, 'No records to delete', 0
#
#             # حذف prescriptions
#             cursor.execute('''
#                 DELETE FROM archive_prescriptions
#                 WHERE original_record_id IN (
#                     SELECT original_id FROM archive_medical_records
#                     WHERE record_date BETWEEN ? AND ?
#                 )
#             ''', (date_from, date_to))
#
#             cursor.execute('DELETE FROM archive_medical_records WHERE record_date BETWEEN ? AND ?',
#                            (date_from, date_to))
#         else:
#             # حذف همه
#             cursor.execute('SELECT COUNT(*) FROM archive_medical_records')
#             count = cursor.fetchone()[0]
#             cursor.execute('DELETE FROM archive_prescriptions')
#             cursor.execute('DELETE FROM archive_medical_records')
#
#         conn.commit()
#         conn.close()
#
#         return True, f'Permanently deleted {count} archived records', count
#
#     except Exception as e:
#         return False, str(e), 0



def get_printer_for(doc_type):
    """
    گرفتن پرینتر برای یک نوع سند خاص

    Args:
        doc_type: 'default', 'prescription', 'invoice', 'report', 'receipt'

    Returns:
        نام پرینتر یا None
    """
    settings = load_settings()
    printers = settings.get('printers', {})

    # 1. چک پرینتر خاص (مثلاً invoice)
    specific = printers.get(doc_type, '').strip()
    if specific and WIN32_AVAILABLE:
        available = get_available_printers()
        if specific in available:
            return specific

    # 2. چک پرینتر پیش‌فرض برنامه
    default_app = printers.get('default', '').strip()
    if default_app and WIN32_AVAILABLE:
        available = get_available_printers()
        if default_app in available:
            return default_app

    # 3. Fallback به پیش‌فرض ویندوز
    return get_default_printer()


def is_virtual_printer(printer_name):
    """تشخیص اینکه پرینتر مجازیه یا فیزیکی"""
    if not printer_name:
        return False

    virtual_keywords = [
        'Print to PDF',
        'PDF',
        'XPS',
        'OneNote',
        'Fax',
        'Virtual',
        'Writer',
    ]

    printer_lower = printer_name.lower()
    return any(kw.lower() in printer_lower for kw in virtual_keywords)


def center_window(win, width, height):
    """وسط‌چین کردن پنجره با اندازه‌ی مشخص"""
    win.update_idletasks()
    screen_w = win.winfo_screenwidth()
    screen_h = win.winfo_screenheight()

    # اگه ارتفاع بیشتر از صفحه بود، کمش کن
    if height > screen_h - 80:
        height = screen_h - 80
    if width > screen_w - 80:
        width = screen_w - 80

    x = (screen_w // 2) - (width // 2)
    y = (screen_h // 2) - (height // 2)
    win.geometry(f'{width}x{height}+{x}+{y}')


def auto_scrollbar(canvas, scrollbar, scrollable):
    """
    اسکرول‌بار خودکار:
    - اگه محتوا بزرگ‌تر از canvas هست → نمایش
    - اگه نه → مخفی
    """

    def check_scrollbar(*args):
        try:
            # به‌روزرسانی scrollregion
            canvas.configure(scrollregion=canvas.bbox('all'))

            # اندازه‌ی محتوا
            bbox = canvas.bbox('all')
            if not bbox:
                return

            content_height = bbox[3] - bbox[1]
            canvas_height = canvas.winfo_height()

            # اگه محتوا بزرگ‌تر از canvas هست → نمایش
            if content_height > canvas_height + 5:
                scrollbar.pack(side='right', fill='y')
            else:
                scrollbar.pack_forget()
        except Exception as e:
            pass

    # اتصال به تغییرات
    canvas.bind('<Configure>', check_scrollbar, add='+')
    scrollable.bind('<Configure>', check_scrollbar, add='+')

    # اجرای اولیه
    canvas.after(100, check_scrollbar)


def auto_scrollbar_horizontal(canvas, x_scrollbar, scrollable):
    """
    اسکرول افقی خودکار:
    - اگه محتوا عریض‌تر از canvas هست → نمایش
    - اگه نه → مخفی
    """

    def check_scrollbar(*args):
        try:
            canvas.configure(scrollregion=canvas.bbox('all'))

            bbox = canvas.bbox('all')
            if not bbox:
                return

            content_width = bbox[2] - bbox[0]
            canvas_width = canvas.winfo_width()

            if content_width > canvas_width + 5:
                x_scrollbar.pack(side='bottom', fill='x')
            else:
                x_scrollbar.pack_forget()
        except Exception as e:
            pass

    canvas.bind('<Configure>', check_scrollbar, add='+')
    scrollable.bind('<Configure>', check_scrollbar, add='+')
    canvas.after(100, check_scrollbar)


# ==========================================
# Main Application Class
# ==========================================

class HospitalApp:
    def __init__(self, root):
        self.root = root
        self.root.title('🏥 Hospital Management System')
        self.root.geometry('1100x750')

        # ===== اعمال تم =====
        try:
            settings = load_settings()
            dark = settings.get('dark_mode', False)

            global CURRENT_THEME_MODE, CURRENT_THEME, WINDOW_BG, TEXT_FG
            CURRENT_THEME_MODE = 'dark' if dark else 'light'
            CURRENT_THEME = get_theme(CURRENT_THEME_MODE)
            WINDOW_BG = CURRENT_THEME['bg']
            TEXT_FG = CURRENT_THEME['fg']
        except Exception as e:
            print(f'⚠️ Theme error: {e}')

        self.root.configure(bg=WINDOW_BG)
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.create_widgets()

        # ===== اعمال استایل =====
        try:
            apply_dark_style()
        except Exception as e:
            print(f'⚠️ Dark style error: {e}')

        self.start_daily_backup()

    def show_audit_log(self):
        """نمایش لاگ سیستم (واکنش‌گرا)"""
        win = tk.Toplevel(self.root)
        win.title('🔐 Audit Log - System Activity')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        # ===== ۱. هدر =====
        header = tk.Frame(win, bg='#34495e', height=55)
        header.pack(fill='x', side='top')
        header.pack_propagate(False)
        tk.Label(header, text='🔐 Audit Log - System Activity',
                 font=('Arial', 15, 'bold'), bg='#34495e',
                 fg='white').pack(pady=12)

        # ===== ۲. دکمه‌های پایین (ثابت) =====
        btn_frame = tk.Frame(win, bg='#ecf0f1')
        btn_frame.pack(side='bottom', fill='x', pady=12)

        # ===== ۳. بدنه اصلی با اسکرول =====
        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=5, pady=5)

        canvas = tk.Canvas(body_container, bg='#ecf0f1', highlightthickness=0)
        v_scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        h_scrollbar = ttk.Scrollbar(body_container, orient='horizontal', command=canvas.xview)

        scrollable = tk.Frame(canvas, bg='#ecf0f1')
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')

        canvas.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')

        body_container.grid_rowconfigure(0, weight=1)
        body_container.grid_columnconfigure(0, weight=1)

        # ===== تابع مدیریت هوشمند اسکرول‌بارها =====
        def check_scrollbars(event=None):
            canvas.configure(scrollregion=canvas.bbox('all'))
            bbox = canvas.bbox('all')
            if bbox:
                content_height = bbox[3] - bbox[1]
                canvas_height = canvas.winfo_height()
                if content_height > canvas_height + 5:
                    v_scrollbar.grid(row=0, column=1, sticky='ns')
                else:
                    v_scrollbar.grid_remove()

                content_width = bbox[2] - bbox[0]
                canvas_width = canvas.winfo_width()
                if content_width > canvas_width + 5:
                    h_scrollbar.grid(row=1, column=0, sticky='ew')
                else:
                    h_scrollbar.grid_remove()

        canvas.bind('<Configure>', check_scrollbars)
        scrollable.bind('<Configure>', check_scrollbars)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۴. محتوای اصلی =====
        content_frame = tk.Frame(scrollable, bg='#ecf0f1')
        content_frame.pack(fill='both', expand=True, padx=20, pady=10)

        # ===== فیلترها =====
        filter_frame = tk.Frame(content_frame, bg='#ecf0f1')
        filter_frame.pack(fill='x', pady=10)

        # ردیف ۱: Search + Action + Entity
        row1 = tk.Frame(filter_frame, bg='#ecf0f1')
        row1.pack(fill='x', pady=5)

        tk.Label(row1, text='🔍 Search:', font=('Arial', 10, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=(0, 5))

        search_var = tk.StringVar()
        search_entry = tk.Entry(row1, textvariable=search_var,
                                font=('Arial', 10), width=25,
                                bg='white', relief='solid', bd=1)
        search_entry.pack(side='left', ipady=4, padx=(0, 15))

        # Action filter
        tk.Label(row1, text='Action:', font=('Arial', 10, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=(5, 5))

        action_var = tk.StringVar(value='All')
        action_combo = ttk.Combobox(row1, textvariable=action_var,
                                    values=['All', 'CREATE', 'UPDATE', 'DELETE',
                                            'LOGIN', 'LOGOUT', 'PRINT', 'BACKUP', 'RESTORE'],
                                    state='readonly', width=12, font=('Arial', 10))
        action_combo.pack(side='left', padx=(0, 15))

        # Entity filter
        tk.Label(row1, text='Entity:', font=('Arial', 10, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=(5, 5))

        entity_var = tk.StringVar(value='All')
        entity_combo = ttk.Combobox(row1, textvariable=entity_var,
                                    values=['All', 'patient', 'doctor', 'appointment',
                                            'payment', 'medicine', 'medical_record',
                                            'prescription', 'user', 'database', 'sale',
                                            'lab_request', 'lab_test'],
                                    state='readonly', width=14, font=('Arial', 10))
        entity_combo.pack(side='left', padx=(0, 10))

        # ردیف ۲: Date filter + Count
        row2 = tk.Frame(filter_frame, bg='#ecf0f1')
        row2.pack(fill='x', pady=5)

        tk.Label(row2, text='📅 Date:', font=('Arial', 10, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=(0, 5))

        date_filter_var, date_from_var, date_to_var, df_container, _ = create_date_filter(
            row2, '#ecf0f1', lambda: load_logs()
        )
        df_container.pack(side='left')

        count_var = tk.StringVar(value='')
        tk.Label(row2, textvariable=count_var,
                 font=('Arial', 10, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(side='right')

        # ===== جدول =====
        tree_frame = tk.Frame(content_frame, bg='#ecf0f1')
        tree_frame.pack(fill='both', expand=True, pady=(0, 10))

        cols = ('ID', 'Timestamp', 'User', 'Role', 'Action', 'Entity', 'ID', 'Details')

        # ارتفاع رو کم میکنیم تا با فیلترها و دکمه‌ها در فول اسکرین جا بشه
        tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=15)

        widths = {'ID': 50, 'Timestamp': 140, 'User': 100, 'Role': 80,
                  'Action': 100, 'Entity': 120, 'ID': 60, 'Details': 400}

        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=widths[c], anchor='w' if c in ('Details', 'User', 'Entity') else 'center')

        scroll = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)

        tree.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        # رنگ‌ها
        tree.tag_configure('CREATE', background='#d5f4e6')
        tree.tag_configure('UPDATE', background='#fff3cd')
        tree.tag_configure('DELETE', background='#f8d7da')
        tree.tag_configure('LOGIN', background='#d1ecf1')
        tree.tag_configure('LOGOUT', background='#d1ecf1')
        tree.tag_configure('PRINT', background='#e2e3e5')
        tree.tag_configure('BACKUP', background='#cce5ff')
        tree.tag_configure('RESTORE', background='#ffe5cc')

        def load_logs():
            from_date, to_date = get_date_range(
                date_filter_var.get(),
                date_from_var.get().strip(),
                date_to_var.get().strip()
            )

            search = search_var.get().strip()
            action = action_var.get()
            entity = entity_var.get()

            conn = get_connection()
            cursor = conn.cursor()

            query = 'SELECT id, timestamp, username, user_role, action, entity_type, entity_id, details FROM audit_log WHERE 1=1'
            params = []

            if search:
                query += ' AND (username LIKE ? OR details LIKE ? OR action LIKE ?)'
                pattern = f'%{search}%'
                params.extend([pattern, pattern, pattern])

            if action != 'All':
                query += ' AND action = ?'
                params.append(action)

            if entity != 'All':
                query += ' AND entity_type = ?'
                params.append(entity)

            if from_date:
                query += ' AND DATE(timestamp) >= ?'
                params.append(from_date)

            if to_date:
                query += ' AND DATE(timestamp) <= ?'
                params.append(to_date)

            query += ' ORDER BY timestamp DESC LIMIT 1000'

            cursor.execute(query, params)
            logs = cursor.fetchall()
            conn.close()

            for item in tree.get_children():
                tree.delete(item)

            for log in logs:
                log_id, timestamp, username, role, action_val, entity_type, entity_id, details = log

                icon = {
                    'CREATE': '➕',
                    'UPDATE': '✏️',
                    'DELETE': '🗑️',
                    'LOGIN': '🔓',
                    'LOGOUT': '🔒',
                    'PRINT': '🖨️',
                    'BACKUP': '💾',
                    'RESTORE': '♻️',
                }.get(action_val, '📌')

                tree.insert('', 'end', values=(
                    log_id,
                    timestamp or '—',
                    username or '—',
                    role or '—',
                    f'{icon} {action_val}',
                    entity_type or '—',
                    entity_id or '—',
                    details or '—'
                ), tags=(action_val,))

            count_var.set(f'📋 {len(logs)} log entries')
            check_scrollbars()  # بعد از لود داده‌ها اسکرول‌بار چک شود

        # Bind فیلترها
        search_var.trace_add('write', lambda *args: load_logs())
        action_combo.bind('<<ComboboxSelected>>', lambda e: load_logs())
        entity_combo.bind('<<ComboboxSelected>>', lambda e: load_logs())

        # ===== Export =====
        def export_log():
            from tkinter import filedialog
            file_path = filedialog.asksaveasfilename(
                defaultextension='.csv',
                filetypes=[('CSV files', '*.csv')],
                initialfile=f'audit_log_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv'
            )
            if not file_path:
                return
            try:
                with open(file_path, 'w', encoding='utf-8-sig') as f:
                    f.write('ID,Timestamp,User,Role,Action,Entity,Entity ID,Details\n')
                    for item in tree.get_children():
                        vals = tree.item(item)['values']
                        f.write(','.join(f'"{v}"' for v in vals) + '\n')
                messagebox.showinfo('✅ Export Successful',
                                    f'Exported {len(tree.get_children())} entries to:\n{file_path}')
            except Exception as e:
                messagebox.showerror('Export Error', f'❌ {e}')

        def clear_old_logs():
            if not messagebox.askyesno('Confirm',
                                       'Delete audit logs older than 6 months?\n\nThis action cannot be undone!'):
                return
            try:
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("DELETE FROM audit_log WHERE timestamp < DATE('now', '-6 months')")
                deleted = cursor.rowcount
                conn.commit()
                conn.close()

                log_action('DELETE', entity_type='audit_log',
                           details=f'Cleaned {deleted} old log entries',
                           user=getattr(self.root, 'current_user', None))

                messagebox.showinfo('Success', f'✅ Deleted {deleted} old entries')
                load_logs()
            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        # ===== دکمه‌ها =====
        create_popup_buttons(btn_frame, [
            {'text': '🔄 Refresh', 'command': load_logs, 'color': 'primary', 'width': 140},
            {'text': '📊 Export CSV', 'command': export_log, 'color': 'info', 'width': 160},
            {'text': '🗑️ Clean Old', 'command': clear_old_logs, 'color': 'warning', 'width': 160},
            {'text': '❌ Close', 'command': win.destroy, 'color': 'danger', 'width': 120},
        ], bg='#ecf0f1')

        # ===== بارگذاری =====
        load_logs()

        # ===== Maximize =====
        win.update_idletasks()
        win.state('zoomed')


    # ==========================================
    # Main Layout
    # ==========================================

    def create_widgets(self):
        # ===== پس‌زمینه‌ی اصلی =====
        self.root.configure(bg=WINDOW_BG)

        # ===== هدر با گرادیانت =====
        header_frame = tk.Frame(self.root, height=80)
        header_frame.pack(fill='x', side='top')
        header_frame.pack_propagate(False)

        header_canvas = tk.Canvas(header_frame, highlightthickness=0, bd=0)
        header_canvas.pack(fill='both', expand=True)

        def on_header_resize(event=None):
            w = header_canvas.winfo_width()
            h = header_canvas.winfo_height()
            if w <= 1 or h <= 1:
                return

            # گرادیانت
            header_canvas.delete('gradient')
            top_color = '#3b4a6b'
            bottom_color = '#667eea'
            r1, g1, b1 = int(top_color[1:3], 16), int(top_color[3:5], 16), int(top_color[5:7], 16)
            r2, g2, b2 = int(bottom_color[1:3], 16), int(bottom_color[3:5], 16), int(bottom_color[5:7], 16)

            for i in range(h):
                ratio = i / h
                r = int(r1 + (r2 - r1) * ratio)
                g = int(g1 + (g2 - g1) * ratio)
                b = int(b1 + (b2 - b1) * ratio)
                header_canvas.create_line(0, i, w, i, fill=f'#{r:02x}{g:02x}{b:02x}', tags='gradient')

            # متن و آیکون
            header_canvas.delete('header_text')

            header_canvas.create_text(
                w / 2 - 170, h / 2,
                text='🏥',
                font=('Arial', 32),
                fill='white',
                anchor='center',
                tags='header_text'
            )

            header_canvas.create_text(
                w / 2 + 70, h / 2,
                text='Hospital Management System',
                font=('Arial', 22, 'bold'),
                fill='white',
                anchor='center',
                tags='header_text'
            )

            # ===== تاریخ، ساعت و نسخه (سمت چپ) =====
            from datetime import datetime
            now = datetime.now()

            # ===== تاریخ، ساعت و نسخه (سمت چپ) =====
            from datetime import datetime
            now = datetime.now()

            # ---------- خط ۱: تاریخ ----------
            # ایموجی تاریخ
            header_canvas.create_text(
                30, h / 2 - 22,
                text='📅',
                font=('Arial', 12),
                anchor='w',
                tags='header_text'
            )
            # متن تاریخ
            header_canvas.create_text(
                55, h / 2 - 22,
                text=now.strftime('%Y-%m-%d'),
                font=('Arial', 11, 'bold'),
                fill='#a8d8ff',
                anchor='w',
                tags='header_text'
            )

            # ---------- خط ۲: ساعت ----------
            # ایموجی ساعت
            header_canvas.create_text(
                30, h / 2,
                text='⏰',
                font=('Arial', 12),
                anchor='w',
                tags='header_text'
            )
            # متن ساعت
            header_canvas.create_text(
                55, h / 2,
                text=now.strftime('%H:%M:%S'),
                font=('Arial', 11),
                fill='#c5d9f1',
                anchor='w',
                tags='header_text'
            )

            # ---------- خط ۳: نسخه ----------
            # ایموجی نسخه
            header_canvas.create_text(
                30, h / 2 + 22,
                text='🏥',
                font=('Arial', 12),
                anchor='w',
                tags='header_text'
            )
            # متن نسخه
            header_canvas.create_text(
                55, h / 2 + 22,
                text='Hospital MS v1.0',
                font=('Arial', 10, 'bold'),
                fill='#c5d9f1',
                anchor='w',
                tags='header_text'
            )

            # اطلاعات کاربر
            if hasattr(self.root, 'current_user') and self.root.current_user:
                user = self.root.current_user
                user_text = f"👤 {user['full_name']}  |  {user['role'].upper()}"
                header_canvas.create_text(
                    w - 30, h / 2,
                    text=user_text,
                    font=('Arial', 11, 'bold'),
                    fill='#a8d8ff',
                    anchor='e',
                    tags='header_text'
                )

        header_canvas.bind('<Configure>', on_header_resize)

        # ✅ تابع آپدیت زمان
        def update_clock():
            try:
                on_header_resize()
                self.root.after(1000, update_clock)
            except:
                pass

        # شروع آپدیت
        self.root.after(1000, update_clock)

        # ===== نوار تب‌ها =====
        apply_modern_notebook_style()

        try:
            apply_dark_style()
        except Exception as e:
            print(f'⚠️ Dark style error: {e}')
        apply_dark_style()

        self.notebook = ttk.Notebook(self.root, style='Modern.TNotebook')
        self.notebook.pack(fill='both', expand=True, padx=20, pady=(15, 20))

        # نقش کاربر
        role = 'admin'
        if hasattr(self.root, 'current_user') and self.root.current_user:
            role = self.root.current_user.get('role', 'admin')

        tab_bg = WINDOW_BG

        # ===== تب‌ها با try/except (پایدار) =====

        # Dashboard
        if role in ['admin', 'doctor', 'receptionist']:
            try:
                self.dashboard_tab = tk.Frame(self.notebook, bg=tab_bg)
                self.notebook.add(self.dashboard_tab, text='🏠 Dashboard')
                self.create_dashboard()
            except Exception as e:
                print(f'❌ Dashboard error: {e}')
                import traceback
                traceback.print_exc()

        # Patients
        if role in ['admin', 'doctor', 'receptionist']:
            try:
                self.patients_tab = tk.Frame(self.notebook, bg=tab_bg)
                self.notebook.add(self.patients_tab, text='👥 Patients')
                self.create_patients_tab()
            except Exception as e:
                print(f'❌ Patients error: {e}')
                import traceback
                traceback.print_exc()

        # Doctors
        if role in ['admin', 'doctor', 'receptionist']:
            try:
                self.doctors_tab = tk.Frame(self.notebook, bg=tab_bg)
                self.notebook.add(self.doctors_tab, text='👨‍⚕️ Doctors')
                self.create_doctors_tab()
            except Exception as e:
                print(f'❌ Doctors error: {e}')
                import traceback
                traceback.print_exc()

        # Medical Records
        if role in ['admin', 'doctor', 'receptionist']:
            try:
                self.medical_records_tab = tk.Frame(self.notebook, bg=tab_bg)
                self.notebook.add(self.medical_records_tab, text='📋 Medical Records')
                self.create_medical_records_tab()
            except Exception as e:
                print(f'❌ Medical Records error: {e}')
                import traceback
                traceback.print_exc()

        # Appointments
        if role in ['admin', 'doctor', 'receptionist']:
            try:
                self.appointments_tab = tk.Frame(self.notebook, bg=tab_bg)
                self.notebook.add(self.appointments_tab, text='📅 Appointments')
                self.create_appointments_tab()
                # ===== Laboratory Tab =====
                if role in ['admin', 'doctor', 'receptionist']:
                    try:
                        self.lab_tab = tk.Frame(self.notebook, bg=tab_bg)
                        self.notebook.add(self.lab_tab, text='🧪 Laboratory')
                        self.create_lab_tab()
                    except Exception as e:
                        print(f'❌ Laboratory error: {e}')
                        import traceback
                        traceback.print_exc()
            except Exception as e:
                print(f'❌ Appointments error: {e}')
                import traceback
                traceback.print_exc()

        # Payments
        if role in ['admin', 'receptionist']:
            try:
                self.payments_tab = tk.Frame(self.notebook, bg=tab_bg)
                self.notebook.add(self.payments_tab, text='💰 Payments')
                self.create_payments_tab()
            except Exception as e:
                print(f'❌ Payments error: {e}')
                import traceback
                traceback.print_exc()

        # Pharmacy
        if role in ['admin', 'pharmacist']:
            try:
                self.pharmacy_tab = tk.Frame(self.notebook, bg=tab_bg)
                self.notebook.add(self.pharmacy_tab, text='💊 Pharmacy')
                self.create_pharmacy_tab()
            except Exception as e:
                print(f'❌ Pharmacy error: {e}')
                import traceback
                traceback.print_exc()

        # Users
        if role == 'admin':
            try:
                self.users_tab = tk.Frame(self.notebook, bg=tab_bg)
                self.notebook.add(self.users_tab, text='👤 Users')
                self.create_users_tab()
            except Exception as e:
                print(f'❌ Users error: {e}')
                import traceback
                traceback.print_exc()

        # Settings
        if role == 'admin':
            try:
                self.settings_tab = tk.Frame(self.notebook, bg=tab_bg)
                self.notebook.add(self.settings_tab, text='⚙️ Settings')
                self.create_settings_tab()
            except Exception as e:
                print(f'❌ Settings error: {e}')
                import traceback
                traceback.print_exc()

        # ===== Reports Tab =====
        if role in ['admin', 'doctor', 'receptionist']:
            try:
                self.reports_tab = tk.Frame(self.notebook, bg=tab_bg)
                self.notebook.add(self.reports_tab, text='📊 Reports')
                self.create_reports_tab()
            except Exception as e:
                print(f'❌ Reports error: {e}')
                import traceback
                traceback.print_exc()

        # ===== یادآوری هنگام لاگین =====
        try:
            self.root.after(2000, self.show_login_reminder)
        except Exception as e:
            print(f'⚠️ Login reminder setup error: {e}')

    def show_login_reminder(self):
        """نمایش پیام یادآوری هنگام لاگین"""
        from datetime import datetime, timedelta

        today = datetime.now().strftime('%Y-%m-%d')
        tomorrow = (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d')

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            'SELECT COUNT(*) FROM appointments WHERE appointment_date = ? AND status = "Scheduled"',
            (today,)
        )
        today_count = cursor.fetchone()[0]
        cursor.execute(
            'SELECT COUNT(*) FROM appointments WHERE appointment_date = ? AND status = "Scheduled"',
            (tomorrow,)
        )
        tomorrow_count = cursor.fetchone()[0]
        conn.close()

        if today_count > 0 or tomorrow_count > 0:
            parts = []
            if today_count > 0:
                parts.append(f'📅 {today_count} appointment(s) TODAY')
            if tomorrow_count > 0:
                parts.append(f'📅 {tomorrow_count} appointment(s) TOMORROW')

            result = messagebox.askyesno(
                '🔔 Appointment Reminders',
                f'You have:\n\n' + '\n'.join(parts) + '\n\n'
                                                      f'View the list now?'
            )

            if result:
                self.show_appointment_reminders()

    def create_dashboard(self):
        """Dashboard (نسخه‌ی درست و ساده)"""
        # ===== پاک کردن محتوای قبلی =====
        for widget in self.dashboard_tab.winfo_children():
            widget.destroy()

        # ===== دکمه‌ها (پایین ثابت) =====
        btn_container = tk.Frame(self.dashboard_tab, bg=WINDOW_BG)
        btn_container.pack(side='bottom', fill='x', pady=15)

        btn_bar = tk.Frame(btn_container, bg=WINDOW_BG)
        btn_bar.pack()

        create_standard_button(btn_bar, '🔄 Refresh', self.refresh_dashboard, color='success')
        create_standard_button(btn_bar, '💾 Backup Now', self.do_backup, color='primary')
        create_standard_button(btn_bar, '📂 Backups', self.show_backup_list, color='info')
        create_standard_button(btn_bar, '📈 Statistics', self.show_clinic_statistics, color='warning')
        create_standard_button(btn_bar, '🔐 Audit Log', self.show_audit_log, color='secondary')
        create_standard_button(btn_bar, '🔔 Reminders', self.show_appointment_reminders, color='warning')

        # ===== محتوا (Scrollable) =====
        content_frame = tk.Frame(self.dashboard_tab, bg=WINDOW_BG)
        content_frame.pack(fill='both', expand=True)

        canvas = tk.Canvas(content_frame, bg=WINDOW_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(content_frame, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg=WINDOW_BG)

        scrollable.bind('<Configure>',
                        lambda e: canvas.configure(scrollregion=canvas.bbox('all')))

        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)

        # ✅ تنظیم عرض کامل
        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)

        canvas.bind('<Configure>', _on_canvas_configure)

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== هشدار نوبت‌ها =====
        # ===== هشدار نوبت‌ها =====
        from datetime import datetime, timedelta
        today = datetime.now().strftime('%Y-%m-%d')
        tomorrow = (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d')

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(
            'SELECT COUNT(*) FROM appointments WHERE appointment_date = ? AND status = "Scheduled"',
            (today,))
        today_count = cursor.fetchone()[0]
        cursor.execute(
            'SELECT COUNT(*) FROM appointments WHERE appointment_date = ? AND status = "Scheduled"',
            (tomorrow,))
        tomorrow_count = cursor.fetchone()[0]
        conn.close()

        if today_count > 0 or tomorrow_count > 0:
            # ✅ تغییر رنگ پس‌زمینه به #2c3e50 (هم‌رنگ جدول Revenue)
            alert_frame = tk.Frame(scrollable, bg='#2c3e50')
            alert_frame.pack(fill='x', padx=40, pady=(15, 5))

            parts = []
            if today_count > 0:
                parts.append(f'📅 {today_count} appointment(s) TODAY')
            if tomorrow_count > 0:
                parts.append(f'📅 {tomorrow_count} appointment(s) TOMORROW')

            lbl = tk.Label(alert_frame, text='🔔 ' + '  |  '.join(parts),
                           font=('Arial', 11, 'bold'), bg='#2c3e50', # ✅ تغییر رنگ
                           fg='white', pady=8, cursor='hand2')
            lbl.pack(fill='x')
            lbl.bind('<Button-1>', lambda e: self.show_appointment_reminders())

        # ===== کارت‌های آماری =====

        # ===== کارت‌های آماری (گِرد و کشیده) =====
        stats = get_stats()

        stats_frame = tk.Frame(scrollable, bg=WINDOW_BG)
        stats_frame.pack(pady=(15, 10))

        # ✅ تغییر کارت چهارم از "Completed" به "Total Revenue"
        cards = [
            ('👥', 'Patients', stats['patients'], '#3498db'),
            ('👨‍⚕️', 'Doctors', stats['doctors'], '#9b59b6'),
            ('📅', 'Scheduled', stats['scheduled'], '#f39c12'),
            ('💰', 'Total Revenue', f"{stats['total_revenue']:,.0f}", '#2ecc71'),
        ]

        cards_inner = tk.Frame(stats_frame, bg=WINDOW_BG)
        cards_inner.pack()

        def make_rounded_card(parent, icon, label, value, color, width=220, height=110):
            """ساخت کارت گِرد و کشیده"""
            # Canvas برای گوشه‌های گِرد
            canvas = tk.Canvas(parent, width=width, height=height,
                               bg=WINDOW_BG, highlightthickness=0, bd=0)
            canvas.pack(side='left', padx=8, pady=5)

            # رسم مستطیل گِرد
            def round_rect(x1, y1, x2, y2, r, **kwargs):
                points = [
                    x1 + r, y1, x2 - r, y1, x2, y1,
                    x2, y1 + r, x2, y2 - r, x2, y2,
                    x2 - r, y2, x1 + r, y2, x1, y2,
                    x1, y2 - r, x1, y1 + r, x1, y1,
                ]
                return canvas.create_polygon(points, **kwargs, smooth=True)

            # سایه (اختیاری)
            round_rect(4, 4, width - 2, height - 2, 20,
                       fill='#d0d0d0', outline='#d0d0d0')

            # مستطیل اصلی
            round_rect(2, 2, width - 4, height - 4, 20,
                       fill=color, outline=color)

            # آیکون
            canvas.create_text(width / 2, 30, text=icon,
                               font=('Arial', 20), fill='white')

            # مقدار
            canvas.create_text(width / 2, 60, text=str(value),
                               font=('Arial', 22, 'bold'), fill='white')

            # برچسب
            canvas.create_text(width / 2, 88, text=label,
                               font=('Arial', 10, 'bold'), fill='white')

            return canvas

        for icon, label, value, color in cards:
            make_rounded_card(cards_inner, icon, label, value, color,
                              width=220, height=110)

        # ===== نمودار درآمد =====
        if MATPLOTLIB_AVAILABLE:
            try:
                chart_frame = tk.Frame(scrollable, bg='white',
                                       relief='solid', bd=1)
                chart_frame.pack(fill='x', padx=40, pady=10)

                header = tk.Frame(chart_frame, bg='#2c3e50')
                header.pack(fill='x')

                # محاسبه‌ی Total
                today_dt = datetime.now()
                total_rev = 0
                conn = get_connection()
                cursor = conn.cursor()
                for i in range(7):
                    date = today_dt - timedelta(days=i)
                    cursor.execute('''
                        SELECT COALESCE(SUM(final_amount), 0)
                        FROM payments
                        WHERE DATE(payment_date) = ? AND payment_status = 'Paid'
                    ''', (date.strftime('%Y-%m-%d'),))
                    total_rev += cursor.fetchone()[0]
                conn.close()

                tk.Label(header, text='📊 Revenue - Last 7 Days',
                         font=('Arial', 12, 'bold'), bg='#2c3e50',
                         fg='white', pady=6).pack(side='left', padx=10)

                tk.Label(header, text=f'💰 Total: {total_rev:,.0f} IRR',
                         font=('Arial', 12, 'bold'), bg='#2c3e50',
                         fg='#f1c40f', pady=6).pack(side='right', padx=10)

                # نمودار
                self._draw_revenue_chart(chart_frame)

            except Exception as e:
                print(f'⚠️ Chart error: {e}')

        # ===== Top Lists =====
        self._draw_top_lists(scrollable)

        # ✅ اضافه کردن نمودار دایره‌ای (Pie Chart)
        pie_frame = tk.Frame(scrollable, bg='white', relief='solid', bd=1)
        pie_frame.pack(fill='x', padx=40, pady=10)

        pie_header = tk.Frame(pie_frame, bg='#2c3e50')
        pie_header.pack(fill='x')
        tk.Label(pie_header, text='🥧 Revenue Breakdown', font=('Arial', 12, 'bold'),
                 bg='#2c3e50', fg='white', pady=6).pack()

        self._draw_revenue_pie_chart(pie_frame)

        # فضای خالی آخر
        tk.Frame(scrollable, bg=WINDOW_BG, height=20).pack()

    def _draw_revenue_chart(self, parent):
        """نمودار درآمد ۷ روز اخیر (ترکیب ویزیت، داروخانه و آزمایشگاه)"""
        from datetime import datetime, timedelta

        today = datetime.now()
        dates = []
        visit_revs = []
        pharm_revs = []
        lab_revs = []
        total_revs = []

        conn = get_connection()
        cursor = conn.cursor()

        for i in range(6, -1, -1):
            date = today - timedelta(days=i)
            date_str = date.strftime('%Y-%m-%d')
            dates.append(date.strftime('%b %d'))

            # درآمد ویزیت
            cursor.execute(
                'SELECT COALESCE(SUM(final_amount), 0) FROM payments WHERE DATE(payment_date) = ? AND payment_status = "Paid"',
                (date_str,))
            v_rev = cursor.fetchone()[0]

            # درآمد داروخانه
            cursor.execute(
                'SELECT COALESCE(SUM(final_amount), 0) FROM pharmacy_payments WHERE DATE(payment_date) = ? AND payment_status = "Paid"',
                (date_str,))
            p_rev = cursor.fetchone()[0]

            # درآمد آزمایشگاه
            cursor.execute(
                'SELECT COALESCE(SUM(final_amount), 0) FROM lab_payments WHERE DATE(payment_date) = ? AND payment_status = "Paid"',
                (date_str,))
            l_rev = cursor.fetchone()[0]

            visit_revs.append(v_rev)
            pharm_revs.append(p_rev)
            lab_revs.append(l_rev)
            total_revs.append(v_rev + p_rev + l_rev)

        conn.close()

        try:
            fig = Figure(figsize=(10, 2.8), dpi=80)
            ax = fig.add_subplot(111)

            # ✅ رسم نمودار ستونی روی هم (Stacked Bar)
            ax.bar(dates, visit_revs, label='Visits', color='#3498db', edgecolor='white')
            ax.bar(dates, pharm_revs, bottom=visit_revs, label='Pharmacy', color='#27ae60', edgecolor='white')

            # محاسبه موقعیت پایه برای آزمایشگاه (ویزیت + داروخانه)
            bottoms_lab = [v + p for v, p in zip(visit_revs, pharm_revs)]
            ax.bar(dates, lab_revs, bottom=bottoms_lab, label='Laboratory', color='#e67e22', edgecolor='white')

            # نمایش مجموع روی هر ستون
            def fmt(v):
                return f'{v / 1000000:.1f}M' if v >= 1000000 else f'{v / 1000:.0f}K' if v >= 1000 else f'{v:.0f}'

            for i, total in enumerate(total_revs):
                if total > 0:
                    ax.annotate(fmt(total), xy=(i, total), xytext=(0, 3), textcoords="offset points",
                                ha='center', va='bottom', fontsize=8, fontweight='bold')

            ax.set_title('Total Revenue - Last 7 Days (Visits + Pharmacy + Lab)', fontsize=11, fontweight='bold',
                         color='#2c3e50')
            ax.set_ylabel('Revenue (IRR)', fontsize=9)
            ax.legend(loc='upper left', fontsize=8)
            ax.grid(axis='y', alpha=0.3, linestyle='--')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)

            from matplotlib.ticker import FuncFormatter
            ax.yaxis.set_major_formatter(FuncFormatter(lambda v, t: fmt(v)))

            fig.tight_layout()

            canvas = FigureCanvasTkAgg(fig, master=parent)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='x', padx=10, pady=5)

        except Exception as e:
            tk.Label(parent, text=f'⚠️ Chart error: {e}',
                     font=('Arial', 9, 'italic'),
                     bg='white', fg='#e74c3c',
                     pady=20).pack()


    def _draw_top_lists(self, parent):
        """Top Medicines و Top Doctors برای Dashboard"""
        # ===== خواندن داده‌ها =====
        conn = get_connection()
        cursor = conn.cursor()

        cursor.execute('''
            SELECT m.name, COUNT(pr.id) as cnt
            FROM prescriptions pr
            JOIN medicines m ON pr.medicine_id = m.id
            GROUP BY m.id
            ORDER BY cnt DESC
            LIMIT 5
        ''')
        top_meds = cursor.fetchall()

        cursor.execute('''
            SELECT d.first_name || ' ' || d.last_name, COUNT(mr.id) as cnt
            FROM medical_records mr
            JOIN doctors d ON mr.doctor_id = d.id
            GROUP BY d.id
            ORDER BY cnt DESC
            LIMIT 5
        ''')
        top_docs = cursor.fetchall()
        conn.close()

        # ===== Frame اصلی =====
        top_frame = tk.Frame(parent, bg=WINDOW_BG)
        top_frame.pack(fill='x', padx=40, pady=10)

        # ===== Top Medicines =====
        med_frame = tk.Frame(top_frame, bg='white', relief='solid', bd=1)
        med_frame.pack(side='left', fill='both', expand=True, padx=(0, 5))

        # ✅ تغییر رنگ هدر به #2c3e50
        med_header = tk.Frame(med_frame, bg='#2c3e50')
        med_header.pack(fill='x')
        med_title = f'🏆 Top {len(top_meds)} Medicines' if top_meds else '🏆 Top Medicines'
        tk.Label(med_header, text=med_title,
                 font=('Arial', 11, 'bold'), bg='#2c3e50', # ✅ تغییر رنگ
                 fg='white', pady=6).pack()

        med_content = tk.Frame(med_frame, bg='white')
        med_content.pack(fill='both', expand=True, padx=10, pady=8)

        if top_meds:
            for i, (name, cnt) in enumerate(top_meds, 1):
                row = tk.Frame(med_content, bg='white')
                row.pack(fill='x', pady=2)

                rank_colors = ['#f1c40f', '#95a5a6', '#e67e22', '#3498db', '#1abc9c']
                rank_color = rank_colors[i - 1] if i <= 5 else '#95a5a6'

                tk.Label(row, text=f' {i} ', font=('Arial', 11, 'bold'),
                         bg=rank_color, fg='white', width=3).pack(side='left', padx=(0, 8))

                tk.Label(row, text=name[:25], font=('Arial', 10),
                         bg='white', anchor='w').pack(side='left', fill='x', expand=True)
                # ✅ تغییر رنگ عدد به #2c3e50
                tk.Label(row, text=f'💊 {cnt}', font=('Arial', 10, 'bold'),
                         bg='white', fg='#2c3e50', width=8, anchor='e').pack(side='right')
        else:
            tk.Label(med_content, text='📭 No data yet',
                     font=('Arial', 10, 'italic'),
                     bg='white', fg='#95a5a6').pack(pady=20)

        # ===== Top Doctors =====
        doc_frame = tk.Frame(top_frame, bg='white', relief='solid', bd=1)
        doc_frame.pack(side='left', fill='both', expand=True, padx=(5, 0))

        # ✅ تغییر رنگ هدر به #2c3e50
        doc_header = tk.Frame(doc_frame, bg='#2c3e50')
        doc_header.pack(fill='x')
        doc_title = f'🏆 Top {len(top_docs)} Doctors' if top_docs else '🏆 Top Doctors'
        tk.Label(doc_header, text=doc_title,
                 font=('Arial', 11, 'bold'), bg='#2c3e50', # ✅ تغییر رنگ
                 fg='white', pady=6).pack()

        doc_content = tk.Frame(doc_frame, bg='white')
        doc_content.pack(fill='both', expand=True, padx=10, pady=8)

        if top_docs:
            for i, (name, cnt) in enumerate(top_docs, 1):
                row = tk.Frame(doc_content, bg='white')
                row.pack(fill='x', pady=2)

                rank_colors = ['#f1c40f', '#95a5a6', '#e67e22', '#3498db', '#1abc9c']
                rank_color = rank_colors[i - 1] if i <= 5 else '#95a5a6'

                tk.Label(row, text=f' {i} ', font=('Arial', 11, 'bold'),
                         bg=rank_color, fg='white', width=3).pack(side='left', padx=(0, 8))

                doctor_clean = name.replace('Dr.', '').replace('dr.', '').strip()
                tk.Label(row, text=f'Dr. {doctor_clean}'[:25], font=('Arial', 10),
                         bg='white', anchor='w').pack(side='left', fill='x', expand=True)
                # ✅ تغییر رنگ عدد به #2c3e50
                tk.Label(row, text=f'👁️ {cnt}', font=('Arial', 10, 'bold'),
                         bg='white', fg='#2c3e50', width=8, anchor='e').pack(side='right')
        else:
            tk.Label(doc_content, text='📭 No data yet',
                     font=('Arial', 10, 'italic'),
                     bg='white', fg='#95a5a6').pack(pady=20)




    def _draw_revenue_pie_chart(self, parent):
        """نمودار دایره‌ای سهم درآمد بخش‌های مختلف"""
        try:
            conn = get_connection()
            cursor = conn.cursor()

            # محاسبه درآمد هر بخش (کل زمان‌ها)
            cursor.execute('SELECT COALESCE(SUM(final_amount), 0) FROM payments WHERE payment_status = "Paid"')
            visit_rev = cursor.fetchone()[0]

            cursor.execute('SELECT COALESCE(SUM(final_amount), 0) FROM pharmacy_payments WHERE payment_status = "Paid"')
            pharm_rev = cursor.fetchone()[0]

            cursor.execute('SELECT COALESCE(SUM(final_amount), 0) FROM lab_payments WHERE payment_status = "Paid"')
            lab_rev = cursor.fetchone()[0]

            conn.close()

            total = visit_rev + pharm_rev + lab_rev
            if total == 0:
                tk.Label(parent, text='📭 No revenue data to display',
                         font=('Arial', 10, 'italic'), bg='white', fg='#95a5a6').pack(pady=20)
                return

            labels = ['Visits', 'Pharmacy', 'Laboratory']
            sizes = [visit_rev, pharm_rev, lab_rev]
            colors = ['#3498db', '#27ae60', '#e67e22']

            fig = Figure(figsize=(10, 3), dpi=80)
            ax = fig.add_subplot(111)

            # رسم نمودار دایره‌ای
            wedges, texts, autotexts = ax.pie(
                sizes, labels=labels, colors=colors,
                autopct='%1.1f%%', startangle=90,
                wedgeprops=dict(width=0.4, edgecolor='w')
            )

            # تنظیم فونت
            for text in texts:
                text.set_fontsize(10)
                text.set_fontweight('bold')
            for autotext in autotexts:
                autotext.set_fontsize(9)
                autotext.set_fontweight('bold')

            ax.axis('equal')  # برای دایره‌ای ماندن
            ax.set_title('Total Revenue Breakdown (All Time)', fontsize=12, fontweight='bold', color='#2c3e50')

            fig.tight_layout()

            canvas = FigureCanvasTkAgg(fig, master=parent)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='x', padx=10, pady=10)

        except Exception as e:
            tk.Label(parent, text=f'⚠️ Chart error: {e}',
                     font=('Arial', 9, 'italic'), bg='white', fg='#e74c3c').pack(pady=20)



    def create_revenue_chart(self):
        """نمودار درآمد ۷ روز اخیر"""
        from datetime import datetime, timedelta

        # ===== خواندن داده‌ها =====
        today = datetime.now()
        dates = []
        revenues = []

        conn = get_connection()
        cursor = conn.cursor()

        for i in range(6, -1, -1):  # ۶ روز قبل تا امروز
            date = today - timedelta(days=i)
            date_str = date.strftime('%Y-%m-%d')
            dates.append(date.strftime('%b %d'))  # Oct 22

            cursor.execute('''
                SELECT COALESCE(SUM(final_amount), 0)
                FROM payments
                WHERE DATE(payment_date) = ?
                  AND payment_status = 'Paid'
            ''', (date_str,))
            revenues.append(cursor.fetchone()[0])

        conn.close()

        total = sum(revenues)

        # ===== Frame اصلی =====
        chart_frame = tk.Frame(self.dashboard_tab, bg='white',
                               relief='solid', bd=1)
        chart_frame.pack(fill='x', padx=40, pady=10)

        # هدر
        header = tk.Frame(chart_frame, bg='#2c3e50')
        header.pack(fill='x')

        tk.Label(header, text='📊 Revenue - Last 7 Days',
                 font=('Arial', 12, 'bold'), bg='#2c3e50',
                 fg='white', pady=6).pack(side='left', padx=10)

        tk.Label(header, text=f'💰 Total: {total:,.0f} IRR',
                 font=('Arial', 12, 'bold'), bg='#2c3e50',
                 fg='#f1c40f', pady=6).pack(side='right', padx=10)

        # ===== ساخت نمودار با matplotlib =====
        try:
            # ساخت Figure
            fig = Figure(figsize=(10, 2.5), dpi=80)
            ax = fig.add_subplot(111)

            # رنگ ستون‌ها (بر اساس بیشترین درآمد)
            # رنگ ستون‌ها (بر اساس بیشترین درآمد)
            max_rev = max(revenues) if revenues else 0
            colors = []
            for r in revenues:
                if max_rev == 0:
                    colors.append('#bdc3c7')
                else:
                    ratio = r / max_rev
                    if ratio >= 0.95:
                        colors.append('#27ae60')  # سبز تیره (بهترین)
                    elif ratio >= 0.75:
                        colors.append('#2ecc71')  # سبز روشن
                    elif ratio >= 0.5:
                        colors.append('#3498db')  # آبی
                    elif ratio >= 0.25:
                        colors.append('#f39c12')  # نارنجی
                    elif ratio > 0:
                        colors.append('#e67e22')  # پرتقالی
                    else:
                        colors.append('#e74c3c')  # قرمز (صفر)

            bars = ax.bar(dates, revenues, color=colors,
                          edgecolor='#2c3e50', linewidth=1)

            # متن بالای ستون
            for bar, rev in zip(bars, revenues):
                height = bar.get_height()
                ax.text(bar.get_x() + bar.get_width() / 2., height,
                        f'{rev / 1000:.0f}K' if rev > 0 else '',
                        ha='center', va='bottom',
                        fontsize=8, fontweight='bold')

            # تنظیمات
            ax.set_ylabel('Revenue (IRR)', fontsize=9)
            ax.set_ylim(0, max(max_rev * 1.15, 1000))
            ax.grid(axis='y', alpha=0.3, linestyle='--')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.tick_params(axis='x', labelsize=9)
            ax.tick_params(axis='y', labelsize=8)

            # فرمت y-axis
            def format_func(value, tick_number):
                if value >= 1_000_000:
                    return f'{value / 1_000_000:.1f}M'
                elif value >= 1000:
                    return f'{value / 1000:.0f}K'
                return f'{value:.0f}'

            from matplotlib.ticker import FuncFormatter
            ax.yaxis.set_major_formatter(FuncFormatter(format_func))

            fig.tight_layout()

            # Embed در Tkinter
            canvas = FigureCanvasTkAgg(fig, master=chart_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill='x', padx=10, pady=5)

        except Exception as e:
            tk.Label(chart_frame, text=f'⚠️ Chart error: {e}',
                     font=('Arial', 9, 'italic'),
                     bg='white', fg='#e74c3c',
                     pady=20).pack()





    def create_top_lists_section(self):
        """Top 5 Medicines و Top 5 Doctors"""
        # ===== Frame اصلی =====
        top_frame = tk.Frame(self.dashboard_tab, bg=WINDOW_BG)
        top_frame.pack(fill='x', padx=40, pady=10)

        # ===== Top Medicines =====
        med_frame = tk.Frame(top_frame, bg='white', relief='solid', bd=1)
        med_frame.pack(side='left', fill='both', expand=True, padx=(0, 5))

        # هدر
        med_header = tk.Frame(med_frame, bg='#c0392b')
        med_header.pack(fill='x')
        tk.Label(med_header, text='🏆 Top 5 Medicines',
                 font=('Arial', 11, 'bold'), bg='#c0392b',
                 fg='white', pady=6).pack()

        # محتوا
        med_content = tk.Frame(med_frame, bg='white')
        med_content.pack(fill='both', expand=True, padx=10, pady=8)

        try:
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                SELECT m.name, COUNT(pr.id) as cnt
                FROM prescriptions pr
                JOIN medicines m ON pr.medicine_id = m.id
                GROUP BY m.id
                ORDER BY cnt DESC
                LIMIT 5
            ''')
            top_meds = cursor.fetchall()
            conn.close()

            if top_meds:
                for i, (name, cnt) in enumerate(top_meds, 1):
                    row = tk.Frame(med_content, bg='white')
                    row.pack(fill='x', pady=2)

                    icon = '🥇' if i == 1 else '🥈' if i == 2 else '🥉' if i == 3 else f'{i}.'

                    tk.Label(row, text=icon, font=('Arial', 11),
                             bg='white', width=3, anchor='w').pack(side='left')
                    tk.Label(row, text=name[:25], font=('Arial', 10),
                             bg='white', anchor='w').pack(side='left', fill='x', expand=True)
                    tk.Label(row, text=f'💊 {cnt}', font=('Arial', 10, 'bold'),
                             bg='white', fg='#c0392b', width=8, anchor='e').pack(side='right')
            else:
                tk.Label(med_content, text='📭 No data yet',
                         font=('Arial', 10, 'italic'),
                         bg='white', fg='#95a5a6').pack(pady=20)

        except Exception as e:
            tk.Label(med_content, text=f'⚠️ {e}',
                     font=('Arial', 9), bg='white',
                     fg='#e74c3c').pack(pady=20)

        # ===== Top Doctors =====
        doc_frame = tk.Frame(top_frame, bg='white', relief='solid', bd=1)
        doc_frame.pack(side='left', fill='both', expand=True, padx=(5, 0))

        # هدر
        doc_header = tk.Frame(doc_frame, bg='#8e44ad')
        doc_header.pack(fill='x')
        tk.Label(doc_header, text='🏆 Top 5 Doctors',
                 font=('Arial', 11, 'bold'), bg='#8e44ad',
                 fg='white', pady=6).pack()

        # محتوا
        doc_content = tk.Frame(doc_frame, bg='white')
        doc_content.pack(fill='both', expand=True, padx=10, pady=8)

        try:
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                SELECT d.first_name || ' ' || d.last_name, COUNT(mr.id) as cnt
                FROM medical_records mr
                JOIN doctors d ON mr.doctor_id = d.id
                GROUP BY d.id
                ORDER BY cnt DESC
                LIMIT 5
            ''')
            top_docs = cursor.fetchall()
            conn.close()

            if top_docs:
                for i, (name, cnt) in enumerate(top_docs, 1):
                    row = tk.Frame(doc_content, bg='white')
                    row.pack(fill='x', pady=2)

                    icon = '🥇' if i == 1 else '🥈' if i == 2 else '🥉' if i == 3 else f'{i}.'
                    doctor_clean = name.replace('Dr.', '').replace('dr.', '').strip()

                    tk.Label(row, text=icon, font=('Arial', 11),
                             bg='white', width=3, anchor='w').pack(side='left')
                    tk.Label(row, text=f'Dr. {doctor_clean}'[:25], font=('Arial', 10),
                             bg='white', anchor='w').pack(side='left', fill='x', expand=True)
                    tk.Label(row, text=f'👁️ {cnt}', font=('Arial', 10, 'bold'),
                             bg='white', fg='#8e44ad', width=8, anchor='e').pack(side='right')
            else:
                tk.Label(doc_content, text='📭 No data yet',
                         font=('Arial', 10, 'italic'),
                         bg='white', fg='#95a5a6').pack(pady=20)

        except Exception as e:
            tk.Label(doc_content, text=f'⚠️ {e}',
                     font=('Arial', 9), bg='white',
                     fg='#e74c3c').pack(pady=20)



    def refresh_dashboard(self):
        for widget in self.dashboard_tab.winfo_children():
            widget.destroy()
        self.dashboard_tab.configure(bg=WINDOW_BG)
        self.create_dashboard()






    def on_closing(self):
        try:
            success, msg = create_backup_daily()
            if success:
                print(f'✅ Auto-backup: {msg}')
                log_action(
                    'BACKUP',
                    entity_type='database',
                    details=f'Auto backup on close: {msg}',
                    user=getattr(self.root, 'current_user', None)
                )
            else:
                print(f'ℹ️ {msg}')
        except Exception as e:
            print(f'⚠️ {e}')

        self.root.destroy()

    def start_daily_backup(self):
        """شروع تایمر Backup روزانه"""

        def backup_loop():
            try:
                success, msg = create_backup_daily()
                if success:
                    print(f'✅ Daily backup at {datetime.now().strftime("%H:%M:%S")}')
                    log_action(
                        'BACKUP',
                        entity_type='database',
                        details=f'Daily auto-backup: {msg}',
                        user=getattr(self.root, 'current_user', None)
                    )
                else:
                    print(f'ℹ️ Skipped: {msg}')
            except Exception as e:
                print(f'⚠️ Daily backup failed: {e}')

            self.root.after(24 * 60 * 60 * 1000, backup_loop)

        self.root.after(60 * 1000, backup_loop)

    def do_backup(self):
        """دستی Backup بگیر"""
        success, result = create_backup(auto=False)

        if success:
            # لاگ اینجا ثبت بشه (که self داریم)
            log_action(
                'BACKUP',
                entity_type='database',
                details=f'Manual backup created: {result}',
                user=getattr(self.root, 'current_user', None)
            )

            db_size = os.path.getsize(DB_NAME) / 1024
            messagebox.showinfo(
                '✅ Backup Successful',
                f'Database backed up successfully!\n\n'
                f'📁 Path: {result}\n\n'
                f'📊 Total backups: {len(get_backup_list())}'
            )
        else:
            messagebox.showerror(
                '❌ Backup Failed',
                f'Could not create backup:\n\n{result}'
            )

    def show_backup_list(self):
        """فرم مدیریت بکاپ‌ها"""
        win = tk.Toplevel(self.root)
        win.title('📂 Backup Management')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#ecf0f1')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه =====
        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        # ===== ۳. عنوان =====
        tk.Label(body_container, text='📂 Backup Management',
                 font=('Arial', 16, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(pady=10)

        # ===== ۴. Notebook =====
        style = ttk.Style()
        style.theme_use('clam')
        notebook = ttk.Notebook(body_container)
        notebook.pack(fill='both', expand=True, padx=20, pady=10)

        # ===== تب ۱: Daily Backups =====
        tab1 = tk.Frame(notebook, bg='#ecf0f1')
        notebook.add(tab1, text='📅 Daily Backups')

        backups = get_backup_list()

        info_frame = tk.Frame(tab1, bg='#ecf0f1')
        info_frame.pack(fill='x', padx=20, pady=10)

        count_label = tk.Label(info_frame,
                               text=f'💾 {len(backups)} backup(s)  |  Max kept: {MAX_BACKUPS}',
                               font=('Arial', 10, 'italic'), bg='#ecf0f1', fg='#7f8c8d')
        count_label.pack(side='left')

        config = load_config()
        if config.get('enabled') and config.get('secondary_path'):
            sec_path = config['secondary_path']
            status = '✅' if os.path.exists(sec_path) else '❌'
            sec_label = tk.Label(info_frame,
                                 text=f'{status} Secondary: {sec_path}',
                                 font=('Arial', 10, 'italic'),
                                 bg='#ecf0f1', fg='#27ae60' if os.path.exists(sec_path) else '#e74c3c')
            sec_label.pack(side='right')
        else:
            tk.Label(info_frame, text='⚠️ No secondary path set',
                     font=('Arial', 10, 'italic'), bg='#ecf0f1', fg='#95a5a6').pack(side='right')

        columns = ('#', 'Date', 'Time', 'Size (KB)', 'File Name')
        tree = ttk.Treeview(tab1, columns=columns, show='headings', height=13)

        tree.column('#', width=50, anchor='center')
        tree.column('Date', width=120, anchor='center')
        tree.column('Time', width=100, anchor='center')
        tree.column('Size (KB)', width=100, anchor='center')
        tree.column('File Name', width=450, anchor='w')

        tree.heading('#', text='#')
        tree.heading('Date', text='📅 Date')
        tree.heading('Time', text='⏰ Time')
        tree.heading('Size (KB)', text='📊 Size')
        tree.heading('File Name', text='📁 File Name')

        for idx, b in enumerate(backups, 1):
            date_part, time_part = b['time'].split(' ')
            tree.insert('', 'end', values=(
                idx, date_part, time_part,
                f'{b["size_kb"]:.1f}',
                b['name']
            ), tags=(b['path'],))

        tree.pack(pady=10, padx=20, fill='both', expand=True)

        # ===== تب ۲: Archived =====
        tab2 = tk.Frame(notebook, bg='#ecf0f1')
        notebook.add(tab2, text='📌 Archived (Permanent)')

        archives = get_archive_list()

        arch_info = tk.Label(tab2,
                             text=f'📌 {len(archives)} archived backup(s)  (never auto-deleted)',
                             font=('Arial', 10, 'italic'), bg='#ecf0f1', fg='#7f8c8d')
        arch_info.pack(fill='x', padx=20, pady=10)

        arch_tree = ttk.Treeview(tab2, columns=columns, show='headings', height=13)

        arch_tree.column('#', width=50, anchor='center')
        arch_tree.column('Date', width=120, anchor='center')
        arch_tree.column('Time', width=100, anchor='center')
        arch_tree.column('Size (KB)', width=100, anchor='center')
        arch_tree.column('File Name', width=450, anchor='w')

        arch_tree.heading('#', text='#')
        arch_tree.heading('Date', text='📅 Date')
        arch_tree.heading('Time', text='⏰ Time')
        arch_tree.heading('Size (KB)', text='📊 Size')
        arch_tree.heading('File Name', text='📁 File Name')

        for idx, b in enumerate(archives, 1):
            date_part, time_part = b['time'].split(' ')
            arch_tree.insert('', 'end', values=(
                idx, date_part, time_part,
                f'{b["size_kb"]:.1f}',
                b['name']
            ), tags=(b['path'],))

        arch_tree.pack(pady=10, padx=20, fill='both', expand=True)

        # ===== Status =====
        status_var = tk.StringVar(value='')
        status_label = tk.Label(body_container, textvariable=status_var,
                                font=('Arial', 10, 'italic'),
                                bg='#ecf0f1', fg='#27ae60')
        status_label.pack(pady=5)

        def show_status(msg, color='#27ae60'):
            status_var.set(msg)
            status_label.config(fg=color)
            win.after(3500, lambda: status_var.set(''))

        def refresh_lists():
            for item in tree.get_children():
                tree.delete(item)
            for idx, b in enumerate(get_backup_list(), 1):
                date_part, time_part = b['time'].split(' ')
                tree.insert('', 'end', values=(idx, date_part, time_part,
                                               f'{b["size_kb"]:.1f}', b['name']),
                            tags=(b['path'],))
            count_label.config(text=f'💾 {len(get_backup_list())} backup(s)  |  Max kept: {MAX_BACKUPS}')

            for item in arch_tree.get_children():
                arch_tree.delete(item)
            for idx, b in enumerate(get_archive_list(), 1):
                date_part, time_part = b['time'].split(' ')
                arch_tree.insert('', 'end', values=(idx, date_part, time_part,
                                                    f'{b["size_kb"]:.1f}', b['name']),
                                 tags=(b['path'],))
            arch_info.config(text=f'📌 {len(get_archive_list())} archived backup(s)  (never auto-deleted)')

        def do_manual_backup():
            success, result = create_backup(auto=False)
            if success:
                show_status(f'✅ New backup: {result}')
                refresh_lists()
            else:
                show_status(f'❌ {result}', '#e74c3c')

        def do_restore():
            current_tab = notebook.index(notebook.select())
            active_tree = tree if current_tab == 0 else arch_tree
            selected = active_tree.selection()
            if not selected:
                show_status('⚠️ Select a backup first!', '#e67e22')
                return

            backup_path = active_tree.item(selected[0])['tags'][0]
            backup_name = active_tree.item(selected[0])['values'][4]

            if not messagebox.askyesno('⚠️ Restore Confirmation',
                                       f'Restore from:\n\n📁 {backup_name}\n\n'
                                       f'⚠️ WARNING:\n'
                                       f'• Current database will be OVERWRITTEN\n'
                                       f'• All changes after this backup will be LOST\n'
                                       f'• A safety backup of current DB will be created first\n\n'
                                       f'Continue?'):
                return

            try:
                create_backup(auto=True)
                shutil.copy2(backup_path, DB_NAME)
                messagebox.showinfo('✅ Restore Successful',
                                    f'Database restored from:\n{backup_name}\n\n'
                                    f'⚠️ RESTART the application!')
                show_status(f'✅ Restored from {backup_name}')
            except Exception as e:
                messagebox.showerror('❌ Restore Failed', f'Error:\n\n{e}')

        def do_archive():
            selected = tree.selection()
            if not selected:
                show_status('⚠️ Select a backup from "Daily Backups"!', '#e67e22')
                return

            backup_path = tree.item(selected[0])['tags'][0]
            backup_name = tree.item(selected[0])['values'][4]

            if not messagebox.askyesno('📌 Archive Backup',
                                       f'Archive this backup permanently?\n\n'
                                       f'📁 {backup_name}\n\n'
                                       f'✅ It will NEVER be auto-deleted.'):
                return

            success, result = archive_backup(backup_path)
            if success:
                show_status(f'✅ Archived: {backup_name}')
                refresh_lists()
            else:
                show_status(f'❌ {result}', '#e74c3c')

        def do_delete():
            current_tab = notebook.index(notebook.select())
            active_tree = tree if current_tab == 0 else arch_tree
            selected = active_tree.selection()
            if not selected:
                show_status('⚠️ Select a backup first!', '#e67e22')
                return

            backup_path = active_tree.item(selected[0])['tags'][0]
            backup_name = active_tree.item(selected[0])['values'][4]

            if not messagebox.askyesno('Confirm Delete',
                                       f'Delete this backup?\n\n📁 {backup_name}\n\n'
                                       f'⚠️ Cannot be undone!'):
                return

            try:
                os.remove(backup_path)
                show_status(f'✅ Deleted: {backup_name}')
                refresh_lists()
            except Exception as e:
                show_status(f'❌ Error: {e}', '#e74c3c')

        def do_settings():
            settings_win = tk.Toplevel(win)
            settings_win.title('⚙️ Backup Settings')
            settings_win.configure(bg='#ecf0f1')
            settings_win.resizable(True, True)

            # دکمه‌ها پایین
            set_btn_container = tk.Frame(settings_win, bg='#ecf0f1')
            set_btn_container.pack(side='bottom', fill='x', pady=15)

            # بدنه
            set_body = tk.Frame(settings_win, bg='#ecf0f1')
            set_body.pack(fill='both', expand=True, padx=10, pady=5)

            tk.Label(set_body, text='⚙️ Secondary Backup Settings',
                     font=('Arial', 14, 'bold'), bg='#ecf0f1',
                     fg='#2c3e50').pack(pady=15)

            tk.Label(set_body,
                     text='Automatically copy each backup to another location\n'
                          '(e.g., External HDD, USB, Network Drive)',
                     font=('Arial', 9, 'italic'), bg='#ecf0f1',
                     fg='#7f8c8d').pack(pady=(0, 15), padx=20)

            cfg = load_config()

            enabled_var = tk.BooleanVar(value=cfg.get('enabled', False))
            tk.Checkbutton(set_body, text='Enable secondary backup',
                           variable=enabled_var, bg='#ecf0f1',
                           font=('Arial', 11, 'bold')).pack(pady=5)

            tk.Label(set_body, text='Secondary Path:',
                     font=('Arial', 10, 'bold'), bg='#ecf0f1',
                     anchor='w').pack(anchor='w', padx=40, pady=(15, 5))

            path_frame = tk.Frame(set_body, bg='#ecf0f1')
            path_frame.pack(fill='x', padx=40)

            path_var = tk.StringVar(value=cfg.get('secondary_path', ''))
            path_entry = tk.Entry(path_frame, textvariable=path_var,
                                  font=('Arial', 10), bg='white')
            path_entry.pack(side='left', fill='x', expand=True, ipady=6)

            def browse():
                from tkinter import filedialog
                folder = filedialog.askdirectory(title='Select backup folder')
                if folder:
                    path_var.set(folder)

            tk.Button(path_frame, text='📁 Browse', command=browse,
                      bg='#3498db', fg='white', font=('Arial', 9, 'bold'),
                      relief='flat', cursor='hand2',
                      padx=10, pady=6).pack(side='left', padx=(5, 0))

            def save_settings():
                new_cfg = {
                    'secondary_path': path_var.get().strip(),
                    'enabled': enabled_var.get()
                }
                save_config(new_cfg)
                messagebox.showinfo('✅ Saved', 'Settings saved!')
                settings_win.destroy()
                refresh_lists()

            create_popup_buttons(set_btn_container, [
                {'text': '💾 Save', 'command': save_settings, 'color': 'success', 'width': 140},
                {'text': '❌ Cancel', 'command': settings_win.destroy, 'color': 'danger', 'width': 120},
            ], bg='#ecf0f1')

            settings_win.update_idletasks()
            settings_win.state('zoomed')

        # ===== ۵. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '💾 Backup Now', 'command': do_manual_backup, 'color': 'primary', 'width': 160},
            {'text': '📌 Archive', 'command': do_archive, 'color': 'secondary', 'width': 140},
            {'text': '♻️ Restore', 'command': do_restore, 'color': 'warning', 'width': 140},
            {'text': '🗑️ Delete', 'command': do_delete, 'color': 'danger', 'width': 140},
            {'text': '⚙️ Settings', 'command': do_settings, 'color': 'info', 'width': 140},
            {'text': '❌ Close', 'command': win.destroy, 'color': 'danger', 'width': 120},
        ], bg='#ecf0f1')

        # ===== ۶. Maximize =====
        win.update_idletasks()
        win.state('zoomed')


    # ==========================================
    # Patients Tab
    # ==========================================

    def create_patients_tab(self):
        tk.Label(self.patients_tab, text='👥 Patient Management',
                 font=('Arial', 14, 'bold'), bg=WINDOW_BG, fg='#2c3e50').pack(pady=15)

        # ===== دکمه‌ها =====
        btn_bar = create_button_bar(self.patients_tab, bg=WINDOW_BG)

        create_standard_button(btn_bar, '➕ Add', self.add_patient_window, color='success')
        create_standard_button(btn_bar, '✏️ Edit', self.edit_patient_window, color='warning')
        create_standard_button(btn_bar, '👁️ Refresh', self.load_patients, color='primary')
        create_standard_button(btn_bar, '📅 Timeline', self.patient_timeline_window, color='secondary')
        create_standard_button(btn_bar, '🗑️ Delete', self.delete_patient, color='danger')
        create_standard_button(btn_bar, '📊 Export Excel', self.export_patients_excel, color='info')

        # ===== ردیف ۱: Search =====
        row1 = tk.Frame(self.patients_tab, bg=WINDOW_BG)
        row1.pack(pady=(5, 3), padx=40, fill='x')

        tk.Label(row1, text='🔍 Search:', font=('Arial', 11, 'bold'),
                 bg=WINDOW_BG, fg='#2c3e50').pack(side='left', padx=(0, 8))

        self.patient_search_var = tk.StringVar()
        search_entry = tk.Entry(row1, textvariable=self.patient_search_var,
                                font=('Arial', 11), bg='white',
                                relief='solid', bd=1, width=40)
        search_entry.pack(side='left', ipady=6, padx=(0, 8))

        self.patient_search_var.trace_add('write',
                                          lambda *args: self.load_patients())

        def clear_search():
            self.patient_search_var.set('')

        tk.Button(row1, text='✖ Clear', command=clear_search,
                  bg='#95a5a6', fg='white', font=('Arial', 9, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=10, pady=4).pack(side='left')

        # ===== ردیف ۲: Count =====
        row2 = tk.Frame(self.patients_tab, bg=WINDOW_BG)
        row2.pack(pady=(0, 10), padx=40, fill='x')

        self.patient_count_var = tk.StringVar(value='')
        tk.Label(row2, textvariable=self.patient_count_var,
                 font=('Arial', 10, 'italic'), bg=WINDOW_BG,
                 fg='#7f8c8d').pack(side='right')

        # ===== Treeview =====
        columns = ('ID', 'Name', 'Age', 'Gender', 'Phone', 'Blood')
        tree_frame = tk.Frame(self.patients_tab, bg=WINDOW_BG)
        tree_frame.pack(pady=(0, 10), padx=40, fill='both', expand=True)

        self.patients_tree = ttk.Treeview(tree_frame, columns=columns,
                                          show='headings', height=15,
                                          selectmode='extended')

        for col in columns:
            self.patients_tree.heading(col, text=col)
            self.patients_tree.column(col, width=140, anchor='center')

        # اسکرول
        scroll = ttk.Scrollbar(tree_frame, orient='vertical',
                               command=self.patients_tree.yview)
        self.patients_tree.configure(yscrollcommand=scroll.set)

        self.patients_tree.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        # فعال‌سازی انتخاب چندتایی
        enable_multi_click_selection(self.patients_tree)

        self.load_patients()

    def export_patients_excel(self):
        """Export Patients به Excel"""
        from tkinter import filedialog

        search = ''
        if hasattr(self, 'patient_search_var'):
            search = self.patient_search_var.get().strip()

        conn = get_connection()
        cursor = conn.cursor()

        if search:
            pattern = f'%{search}%'
            cursor.execute('''
                SELECT national_id, first_name, last_name, age, gender,
                       phone, email, blood_type
                FROM patients
                WHERE first_name LIKE ? OR last_name LIKE ?
                    OR national_id LIKE ? OR phone LIKE ?
                ORDER BY last_name
            ''', (pattern, pattern, pattern, pattern))
        else:
            cursor.execute('''
                SELECT national_id, first_name, last_name, age, gender,
                       phone, email, blood_type
                FROM patients ORDER BY last_name
            ''')

        patients = cursor.fetchall()
        conn.close()

        if not patients:
            messagebox.showwarning('Warning', '📭 No data to export!')
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[('Excel files', '*.xlsx')],
            initialfile=f'patients_{datetime.now().strftime("%Y%m%d")}.xlsx'
        )

        if not file_path:
            return

        headers = ['National ID', 'First Name', 'Last Name', 'Age', 'Gender',
                   'Phone', 'Email', 'Blood Type']
        rows = [tuple(p) for p in patients]

        success, msg = export_to_excel(
            headers, rows, file_path,
            sheet_name='Patients',
            title=f'Patients List ({len(patients)} records)'
        )

        if success:
            messagebox.showinfo('✅ Export Successful', f'Excel saved to:\n{msg}')
            try:
                os.startfile(msg)
            except:
                pass
        else:
            messagebox.showerror('Export Error', f'❌ {msg}')


    def load_patients(self):
        # ===== خواندن مقدار جستجو =====
        search_text = ''
        if hasattr(self, 'patient_search_var'):
            search_text = self.patient_search_var.get().strip()

        conn = get_connection()
        cursor = conn.cursor()

        if search_text:
            # جستجو در نام، نام خانوادگی، کد ملی، موبایل
            search_pattern = f'%{search_text}%'
            cursor.execute('''
                SELECT id, first_name, last_name, age, gender, phone, blood_type
                FROM patients
                WHERE
                    first_name LIKE ?
                    OR last_name LIKE ?
                    OR (first_name || ' ' || last_name) LIKE ?
                    OR national_id LIKE ?
                    OR phone LIKE ?
                ORDER BY id
            ''', (search_pattern, search_pattern, search_pattern,
                  search_pattern, search_pattern))
        else:
            cursor.execute('''
                SELECT id, first_name, last_name, age, gender, phone, blood_type
                FROM patients ORDER BY id
            ''')

        patients = cursor.fetchall()
        conn.close()

        # پاک کردن جدول
        for item in self.patients_tree.get_children():
            self.patients_tree.delete(item)

        # اضافه کردن ردیف‌ها
        for index, p in enumerate(patients, 1):
            real_id = p[0]
            self.patients_tree.insert(
                '', 'end',
                values=(index, f'{p[1]} {p[2]}', p[3], p[4], p[5], p[6]),
                tags=(str(real_id),)
            )

        # آپدیت تعداد
        if hasattr(self, 'patient_count_var'):
            if search_text:
                self.patient_count_var.set(f'🔍 {len(patients)} result(s)')
            else:
                self.patient_count_var.set(f'📋 {len(patients)} patient(s)')

    def add_patient_window(self):
        """فرم اضافه کردن بیمار جدید"""
        win = tk.Toplevel(self.root)
        win.title('➕ Add New Patient')
        win.configure(bg='#f0f0f0')
        win.resizable(True, True)


        # ===== ۱. دکمه‌ها (پایین ثابت) =====
        btn_container = tk.Frame(win, bg='#f0f0f0')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#f0f0f0')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f0f0f0')

        scrollable.bind('<Configure>',
                        lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>',
                    lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text='➕ Add New Patient',
                 font=('Arial', 16, 'bold'), bg='#f0f0f0',
                 fg='#2c3e50').pack(pady=15)

        # ===== ۴. فرم =====
        form_frame = tk.Frame(scrollable, bg='#f0f0f0')
        form_frame.pack(pady=10, padx=40, fill='x')

        fields = {}
        labels = [
            ('National ID *', 'national_id', '10 digits'),
            ('First Name *', 'first_name', 'Letters only'),
            ('Last Name *', 'last_name', 'Letters only'),
            ('Age *', 'age', 'Number (1-119)'),
            ('Gender *', 'gender', 'Male or Female'),
            ('Phone *', 'phone', '09xxxxxxxxx'),
            ('Email (Optional)', 'email', 'user@domain.com'),
            ('Address (Optional)', 'address', 'Full address'),
            ('Blood Type *', 'blood_type', 'A+, A-, B+, B-, AB+, AB-, O+, O-'),
        ]

        for i, (label, key, hint) in enumerate(labels):
            tk.Label(form_frame, text=label, bg='#f0f0f0',
                     font=('Arial', 10, 'bold'), anchor='w').grid(
                row=i * 2, column=0, sticky='w', pady=(5, 0), padx=5)
            entry = tk.Entry(form_frame, width=40, font=('Arial', 10))
            entry.grid(row=i * 2 + 1, column=0, pady=(0, 5), padx=5, sticky='ew')
            fields[key] = entry
            tk.Label(form_frame, text=f'💡 {hint}', bg='#f0f0f0',
                     font=('Arial', 8), fg='#7f8c8d', anchor='w').grid(
                row=i * 2 + 1, column=1, sticky='w', padx=5)

        form_frame.grid_columnconfigure(0, weight=1)

        # ===== ۵. تابع ذخیره =====
        def save():
            try:
                nid = fields['national_id'].get().strip()
                if not is_valid_national_id(nid):
                    messagebox.showerror('Error', '❌ National ID must be exactly 10 digits!')
                    return

                first = fields['first_name'].get().strip()
                if not is_valid_name(first):
                    messagebox.showerror('Error', '❌ First Name must contain only letters!')
                    return

                last = fields['last_name'].get().strip()
                if not is_valid_name(last):
                    messagebox.showerror('Error', '❌ Last Name must contain only letters!')
                    return

                age_str = fields['age'].get().strip()
                try:
                    age = int(age_str)
                except ValueError:
                    messagebox.showerror('Error', '❌ Age must be a number!')
                    return
                if not (0 < age < 120):
                    messagebox.showerror('Error', '❌ Age must be between 1-119!')
                    return

                gender = fields['gender'].get().strip().capitalize()
                if gender not in ['Male', 'Female']:
                    messagebox.showerror('Error', '❌ Gender must be "Male" or "Female"!')
                    return

                phone = fields['phone'].get().strip()
                if not is_valid_phone(phone):
                    messagebox.showerror('Error', '❌ Phone must start with 09 (11 digits)!')
                    return

                email = fields['email'].get().strip()
                if email and not is_valid_email(email):
                    messagebox.showerror('Error', '❌ Invalid email format!')
                    return

                address = fields['address'].get().strip()

                blood = fields['blood_type'].get().strip().upper()
                if blood not in ['A+', 'A-', 'B+', 'B-', 'AB+', 'AB-', 'O+', 'O-']:
                    messagebox.showerror('Error', '❌ Invalid blood type!')
                    return

                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO patients (national_id, first_name, last_name, age, gender, phone, email, address, blood_type)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (nid, first, last, age, gender, phone, email, address, blood))
                conn.commit()
                conn.close()

                messagebox.showinfo('Success', f'✅ Patient "{first} {last}" added!')
                try:
                    log_action('CREATE', entity_type='patient',
                               details=f'Added patient: {first} {last} (NID: {nid})',
                               user=getattr(self.root, 'current_user', None))
                except:
                    pass
                win.destroy()
                self.load_patients()
                self.refresh_dashboard()

            except sqlite3.IntegrityError:
                messagebox.showerror('Error', '❌ National ID already exists!')
            except Exception as e:
                messagebox.showerror('Error', f'❌ Error: {e}')

        # ===== ۶. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '💾 Save', 'command': save, 'color': 'success', 'width': 160},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#f0f0f0')

        center_window(win)  # ← Maximize

        fields['national_id'].focus()

    def delete_patient(self):
        selected = self.patients_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', 'Please select at least one patient first!')
            return

        count = len(selected)

        if count == 1:
            values = self.patients_tree.item(selected[0])['values']
            confirm_msg = f'Delete "{values[1]}"?'
        else:
            confirm_msg = f'Delete {count} patient(s)?\n\nThis action cannot be undone!'

        if not messagebox.askyesno('Confirm', confirm_msg):
            return

        try:
            conn = get_connection()
            cursor = conn.cursor()

            for item in selected:
                real_id = self.patients_tree.item(item)['tags'][0]
                cursor.execute('DELETE FROM patients WHERE id = ?', (real_id,))

            conn.commit()
            conn.close()

            messagebox.showinfo('Success', f'✅ {count} patient(s) deleted!')
            self.load_patients()
            self.refresh_dashboard()

        except Exception as e:
            messagebox.showerror('Error', f'❌ Error: {e}')

    def patient_timeline_window(self):
        """خط زمانی کامل بیمار"""
        selected = self.patients_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select a patient first!')
            return

        patient_id = self.patients_tree.item(selected[0])['tags'][0]

        # ===== خواندن اطلاعات بیمار =====
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT id, first_name, last_name, national_id, phone,
                   age, gender, blood_type, email, address
            FROM patients WHERE id = ?
        ''', (patient_id,))
        patient = cursor.fetchone()

        if not patient:
            conn.close()
            messagebox.showerror('Error', '❌ Patient not found!')
            return

        # ===== خواندن همه‌ی ویزیت‌ها =====
        cursor.execute('''
            SELECT
                mr.id, mr.record_date, mr.chief_complaint, mr.diagnosis,
                mr.symptoms, mr.blood_pressure, mr.temperature,
                mr.weight, mr.height, mr.notes,
                d.first_name || ' ' || d.last_name AS doctor_name,
                d.specialty,
                (SELECT COUNT(*) FROM prescriptions WHERE record_id = mr.id) AS rx_count
            FROM medical_records mr
            JOIN doctors d ON mr.doctor_id = d.id
            WHERE mr.patient_id = ?
            ORDER BY mr.record_date DESC, mr.id DESC
        ''', (patient_id,))
        visits = cursor.fetchall()
        conn.close()

        # ===== پنجره =====
        win = tk.Toplevel(self.root)
        win.title(f'📅 Patient Timeline: {patient[1]} {patient[2]}')
        win.geometry('1000x850')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)
        win.minsize(800, 600)

        win.update_idletasks()
        w, h = 1000, 850
        x = (win.winfo_screenwidth() // 2) - (w // 2)
        y = (win.winfo_screenheight() // 2) - (h // 2)
        win.geometry(f'{w}x{h}+{x}+{y}')

        # ===== هدر =====
        header = tk.Frame(win, bg='#9b59b6', height=55)
        header.pack(fill='x', side='top')
        header.pack_propagate(False)
        tk.Label(header, text=f'📅 Patient Timeline: {patient[1]} {patient[2]}',
                 font=('Arial', 15, 'bold'), bg='#9b59b6',
                 fg='white').pack(pady=12)

        # ===== دکمه‌های پایین =====
        btn_frame = tk.Frame(win, bg='#ecf0f1')
        btn_frame.pack(side='bottom', fill='x', pady=12)

        # ===== بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#ecf0f1', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#ecf0f1')
        scrollable.bind('<Configure>',
                        lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>',
                    lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== بخش اطلاعات بیمار =====
        info_frame = tk.LabelFrame(scrollable, text=' 👤 Patient Information ',
                                   font=('Arial', 12, 'bold'),
                                   bg='#ecf0f1', fg='#9b59b6',
                                   padx=15, pady=12)
        info_frame.pack(fill='x', pady=(10, 15), padx=15)

        info_inner = tk.Frame(info_frame, bg='#ecf0f1')
        info_inner.pack(fill='x')

        # ردیف ۱
        r1 = tk.Frame(info_inner, bg='#ecf0f1')
        r1.pack(fill='x', pady=2)

        for label, value in [
            ('🆔 National ID:', patient[3] or '—'),
            ('📞 Phone:', patient[4] or '—'),
            ('🎂 Age:', f'{patient[5]} years' if patient[5] else '—'),
        ]:
            tk.Label(r1, text=label, font=('Arial', 10, 'bold'),
                     bg='#ecf0f1', width=14, anchor='w').pack(side='left')
            tk.Label(r1, text=value, font=('Arial', 10),
                     bg='#ecf0f1', width=18, anchor='w').pack(side='left')

        # ردیف ۲
        r2 = tk.Frame(info_inner, bg='#ecf0f1')
        r2.pack(fill='x', pady=2)

        for label, value in [
            ('👤 Gender:', patient[6] or '—'),
            ('🩸 Blood Type:', patient[7] or '—'),
            ('📧 Email:', patient[8] or '—'),
        ]:
            tk.Label(r2, text=label, font=('Arial', 10, 'bold'),
                     bg='#ecf0f1', width=14, anchor='w').pack(side='left')
            tk.Label(r2, text=value, font=('Arial', 10),
                     bg='#ecf0f1', width=18, anchor='w').pack(side='left')

        # ردیف ۳: آدرس
        if patient[9]:
            r3 = tk.Frame(info_inner, bg='#ecf0f1')
            r3.pack(fill='x', pady=2)
            tk.Label(r3, text='📍 Address:', font=('Arial', 10, 'bold'),
                     bg='#ecf0f1', width=14, anchor='w').pack(side='left')
            tk.Label(r3, text=patient[9], font=('Arial', 10),
                     bg='#ecf0f1', anchor='w', wraplength=700,
                     justify='left').pack(side='left', fill='x', expand=True)

        # ===== آمار کلی =====
        stats_frame = tk.Frame(scrollable, bg='#ecf0f1')
        stats_frame.pack(fill='x', pady=(0, 15), padx=15)

        total_visits = len(visits)
        total_rx = sum(v[12] for v in visits)

        def mini_stat(icon, label, value, color):
            card = tk.Frame(stats_frame, bg=color, height=70)
            card.pack(side='left', padx=5, fill='x', expand=True)
            card.pack_propagate(False)

            inner = tk.Frame(card, bg=color)
            inner.pack(pady=10)

            tk.Label(inner, text=f'{icon} {value}',
                     font=('Arial', 16, 'bold'),
                     bg=color, fg='white').pack()
            tk.Label(inner, text=label,
                     font=('Arial', 9),
                     bg=color, fg='white').pack()

        mini_stat('🏥', 'Total Visits', total_visits, '#3498db')
        mini_stat('💊', 'Prescriptions', total_rx, '#e74c3c')
        mini_stat('👨‍⚕️', 'Unique Doctors',
                  len(set(v[10] for v in visits)), '#27ae60')
        mini_stat('📅', 'Last Visit',
                  visits[0][1] if visits else '—', '#f39c12')

        # ===== خط زمانی =====
        tk.Label(scrollable, text='📅 Visit History (Timeline):',
                 font=('Arial', 12, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(anchor='w', padx=15, pady=(10, 10))

        if not visits:
            empty_frame = tk.Frame(scrollable, bg='#ecf0f1')
            empty_frame.pack(fill='x', pady=40, padx=15)

            tk.Label(empty_frame, text='📭',
                     font=('Arial', 40), bg='#ecf0f1',
                     fg='#95a5a6').pack()
            tk.Label(empty_frame, text='No visits recorded yet',
                     font=('Arial', 12, 'italic'), bg='#ecf0f1',
                     fg='#95a5a6').pack(pady=5)
            tk.Label(empty_frame, text='Click "New Record" in Medical Records tab to add one.',
                     font=('Arial', 10), bg='#ecf0f1',
                     fg='#95a5a6').pack()
        else:
            for i, v in enumerate(visits):
                record_id = v[0]
                visit_date = v[1]
                chief = v[2]
                diagnosis = v[3]
                doctor = v[10].replace('Dr.', '').replace('dr.', '').strip()
                specialty = v[11]
                rx_count = v[12]

                # ===== کارت ویزیت =====
                card = tk.Frame(scrollable, bg='white',
                                relief='solid', bd=1)
                card.pack(fill='x', pady=6, padx=15)

                # هدر کارت
                card_header = tk.Frame(card, bg='#9b59b6', height=38)
                card_header.pack(fill='x')
                card_header.pack_propagate(False)

                tk.Label(card_header, text=f'  📅 {visit_date}',
                         font=('Arial', 11, 'bold'),
                         bg='#9b59b6', fg='white',
                         anchor='w').pack(side='left', padx=10)

                tk.Label(card_header, text=f'👨‍⚕️ Dr. {doctor} | {specialty}  ',
                         font=('Arial', 10),
                         bg='#9b59b6', fg='white',
                         anchor='e').pack(side='right', padx=10)

                # بدنه
                body = tk.Frame(card, bg='white')
                body.pack(fill='x', padx=15, pady=10)

                # Diagnosis
                if diagnosis:
                    d_row = tk.Frame(body, bg='white')
                    d_row.pack(fill='x', pady=2)
                    tk.Label(d_row, text='🩺 Diagnosis:',
                             font=('Arial', 10, 'bold'),
                             bg='white', fg='#27ae60',
                             width=15, anchor='w').pack(side='left')
                    tk.Label(d_row, text=diagnosis,
                             font=('Arial', 10), bg='white',
                             anchor='w', justify='left',
                             wraplength=750).pack(side='left', fill='x', expand=True)

                # Chief Complaint
                if chief:
                    c_row = tk.Frame(body, bg='white')
                    c_row.pack(fill='x', pady=2)
                    tk.Label(c_row, text='💬 Complaint:',
                             font=('Arial', 10, 'bold'),
                             bg='white', fg='#3498db',
                             width=15, anchor='w').pack(side='left')
                    tk.Label(c_row, text=chief,
                             font=('Arial', 10), bg='white',
                             anchor='w', justify='left',
                             wraplength=750).pack(side='left', fill='x', expand=True)

                # Vital Signs (خلاصه)
                vital_parts = []
                if v[5]:
                    vital_parts.append(f'BP: {v[5]}')
                if v[6]:
                    vital_parts.append(f'Temp: {v[6]}°C')
                if v[7]:
                    vital_parts.append(f'Weight: {v[7]}kg')
                if v[8]:
                    vital_parts.append(f'Height: {v[8]}cm')

                if vital_parts:
                    v_row = tk.Frame(body, bg='white')
                    v_row.pack(fill='x', pady=2)
                    tk.Label(v_row, text='🩺 Vital:',
                             font=('Arial', 10, 'bold'),
                             bg='white', fg='#e67e22',
                             width=15, anchor='w').pack(side='left')
                    tk.Label(v_row, text='  |  '.join(vital_parts),
                             font=('Arial', 10), bg='white',
                             anchor='w').pack(side='left', fill='x', expand=True)

                # Prescriptions
                if rx_count > 0:
                    p_row = tk.Frame(body, bg='white')
                    p_row.pack(fill='x', pady=2)
                    tk.Label(p_row, text='💊 Prescriptions:',
                             font=('Arial', 10, 'bold'),
                             bg='white', fg='#e74c3c',
                             width=15, anchor='w').pack(side='left')
                    tk.Label(p_row, text=f'{rx_count} medicine(s) prescribed',
                             font=('Arial', 10, 'bold'), bg='white',
                             fg='#e74c3c', anchor='w').pack(side='left', fill='x', expand=True)
                else:
                    p_row = tk.Frame(body, bg='white')
                    p_row.pack(fill='x', pady=2)
                    tk.Label(p_row, text='💊 Prescriptions:',
                             font=('Arial', 10, 'bold'),
                             bg='white', fg='#7f8c8d',
                             width=15, anchor='w').pack(side='left')
                    tk.Label(p_row, text='No prescription',
                             font=('Arial', 10, 'italic'), bg='white',
                             fg='#95a5a6', anchor='w').pack(side='left', fill='x', expand=True)

                # Notes
                if v[9]:
                    n_row = tk.Frame(body, bg='white')
                    n_row.pack(fill='x', pady=2)
                    tk.Label(n_row, text='📝 Notes:',
                             font=('Arial', 10, 'bold'),
                             bg='white', fg='#7f8c8d',
                             width=15, anchor='w').pack(side='left')
                    tk.Label(n_row, text=v[9],
                             font=('Arial', 9, 'italic'), bg='white',
                             anchor='w', justify='left',
                             wraplength=750).pack(side='left', fill='x', expand=True)

                # ===== دکمه‌های اکشن =====
                action_row = tk.Frame(card, bg='#f8f9fa')
                action_row.pack(fill='x', padx=10, pady=(5, 10))

                tk.Button(action_row, text='👁️ View Full',
                          command=lambda rid=record_id: self.view_medical_record_by_id(rid),
                          bg='#3498db', fg='white', font=('Arial', 9, 'bold'),
                          relief='flat', cursor='hand2',
                          padx=12, pady=4).pack(side='left', padx=3)

                tk.Button(action_row, text='🖨️ Print',
                          command=lambda rid=record_id: self.print_medical_record_by_id(rid),
                          bg='#9b59b6', fg='white', font=('Arial', 9, 'bold'),
                          relief='flat', cursor='hand2',
                          padx=12, pady=4).pack(side='left', padx=3)

                # شماره‌ی ویزیت
                tk.Label(action_row, text=f'Record #{record_id}',
                         font=('Arial', 8, 'italic'),
                         bg='#f8f9fa', fg='#95a5a6').pack(side='right', padx=10)

        # ===== تابع Export =====
        def export_timeline():
            from tkinter import filedialog

            file_path = filedialog.asksaveasfilename(
                defaultextension='.txt',
                filetypes=[('Text files', '*.txt')],
                initialfile=f'timeline_{patient[1]}_{patient[2]}_{datetime.now().strftime("%Y%m%d")}.txt'
            )

            if not file_path:
                return

            try:
                lines = []
                lines.append('=' * 70)
                lines.append(' ' * 15 + 'PATIENT TIMELINE')
                lines.append('=' * 70)
                lines.append('')
                lines.append(f'  Patient: {patient[1]} {patient[2]}')
                lines.append(f'  National ID: {patient[3]}')
                lines.append(f'  Phone: {patient[4]}')
                lines.append(f'  Age: {patient[5]} years')
                lines.append(f'  Gender: {patient[6]}')
                lines.append(f'  Blood Type: {patient[7]}')
                if patient[9]:
                    lines.append(f'  Address: {patient[9]}')
                lines.append('')
                lines.append(f'  Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
                lines.append('')
                lines.append(f'  Total Visits: {total_visits}')
                lines.append(f'  Total Prescriptions: {total_rx}')
                lines.append('')
                lines.append('=' * 70)
                lines.append('')

                for v in visits:
                    lines.append('-' * 70)
                    lines.append(f'  📅 {v[1]}')
                    lines.append(f'  👨‍⚕️ Dr. {v[10]} | {v[11]}')
                    lines.append(f'  Record #{v[0]}')
                    lines.append('-' * 70)

                    if v[3]:
                        lines.append(f'  🩺 Diagnosis: {v[3]}')
                    if v[2]:
                        lines.append(f'  💬 Complaint: {v[2]}')

                    vital = []
                    if v[5]:
                        vital.append(f'BP: {v[5]}')
                    if v[6]:
                        vital.append(f'Temp: {v[6]}C')
                    if v[7]:
                        vital.append(f'Weight: {v[7]}kg')
                    if v[8]:
                        vital.append(f'Height: {v[8]}cm')
                    if vital:
                        lines.append(f'  🩺 Vital: {" | ".join(vital)}')

                    if v[12] > 0:
                        lines.append(f'  💊 Prescriptions: {v[12]} medicine(s)')

                    if v[9]:
                        lines.append(f'  📝 Notes: {v[9]}')

                    lines.append('')

                lines.append('=' * 70)

                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(lines))

                messagebox.showinfo('✅ Export Successful',
                                    f'Timeline saved to:\n{file_path}')

            except Exception as e:
                messagebox.showerror('Export Error', f'❌ {e}')

        # ===== دکمه‌های پایین =====
        create_popup_buttons(btn_frame, [
            {'text': '📊 Export Timeline', 'command': export_timeline,
             'color': 'info', 'width': 200},
            {'text': '❌ Close', 'command': win.destroy,
             'color': 'danger', 'width': 140},
        ], bg='#ecf0f1')

    def edit_patient_window(self):
        """فرم ویرایش بیمار"""
        selected = self.patients_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select a patient first!')
            return

        real_id = self.patients_tree.item(selected[0])['tags'][0]

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT national_id, first_name, last_name, age, gender,
                   phone, email, address, blood_type
            FROM patients WHERE id = ?
        ''', (real_id,))
        p = cursor.fetchone()
        conn.close()

        if not p:
            messagebox.showerror('Error', '❌ Patient not found!')
            return

        win = tk.Toplevel(self.root)
        win.title('✏️ Edit Patient')
        win.configure(bg='#f0f0f0')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#f0f0f0')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#f0f0f0')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f0f0f0')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text='✏️ Edit Patient',
                 font=('Arial', 16, 'bold'), bg='#f0f0f0',
                 fg='#2c3e50').pack(pady=15)

        # ===== ۴. فرم =====
        form_frame = tk.Frame(scrollable, bg='#f0f0f0')
        form_frame.pack(pady=10, padx=40, fill='x')

        fields = {}
        labels = [
            ('National ID *', 'national_id', '10 digits'),
            ('First Name *', 'first_name', 'Letters only'),
            ('Last Name *', 'last_name', 'Letters only'),
            ('Age *', 'age', 'Number (1-119)'),
            ('Gender *', 'gender', 'Male or Female'),
            ('Phone *', 'phone', '09xxxxxxxxx'),
            ('Email (Optional)', 'email', 'user@domain.com'),
            ('Address (Optional)', 'address', 'Full address'),
            ('Blood Type *', 'blood_type', 'A+, A-, B+, B-, AB+, AB-, O+, O-'),
        ]

        initial_values = {
            'national_id': p[0],
            'first_name': p[1],
            'last_name': p[2],
            'age': str(p[3]),
            'gender': p[4],
            'phone': p[5],
            'email': p[6] or '',
            'address': p[7] or '',
            'blood_type': p[8],
        }

        for i, (label, key, hint) in enumerate(labels):
            tk.Label(form_frame, text=label, bg='#f0f0f0',
                     font=('Arial', 10, 'bold'), anchor='w').grid(
                row=i * 2, column=0, sticky='w', pady=(5, 0), padx=5)
            entry = tk.Entry(form_frame, width=40, font=('Arial', 10))
            entry.grid(row=i * 2 + 1, column=0, pady=(0, 5), padx=5, sticky='ew')
            entry.insert(0, initial_values[key])
            fields[key] = entry
            tk.Label(form_frame, text=f'💡 {hint}', bg='#f0f0f0',
                     font=('Arial', 8), fg='#7f8c8d', anchor='w').grid(
                row=i * 2 + 1, column=1, sticky='w', padx=5)

        form_frame.grid_columnconfigure(0, weight=1)

        # ===== ۵. تابع ذخیره =====
        def save():
            try:
                nid = fields['national_id'].get().strip()
                if not is_valid_national_id(nid):
                    messagebox.showerror('Error', '❌ National ID must be exactly 10 digits!')
                    return
                first = fields['first_name'].get().strip()
                if not is_valid_name(first):
                    messagebox.showerror('Error', '❌ First Name must contain only letters!')
                    return
                last = fields['last_name'].get().strip()
                if not is_valid_name(last):
                    messagebox.showerror('Error', '❌ Last Name must contain only letters!')
                    return
                age_str = fields['age'].get().strip()
                try:
                    age = int(age_str)
                except ValueError:
                    messagebox.showerror('Error', '❌ Age must be a number!')
                    return
                if not (0 < age < 120):
                    messagebox.showerror('Error', '❌ Age must be between 1-119!')
                    return
                gender = fields['gender'].get().strip().capitalize()
                if gender not in ['Male', 'Female']:
                    messagebox.showerror('Error', '❌ Gender must be "Male" or "Female"!')
                    return
                phone = fields['phone'].get().strip()
                if not is_valid_phone(phone):
                    messagebox.showerror('Error', '❌ Phone must start with 09 (11 digits)!')
                    return
                email = fields['email'].get().strip()
                if email and not is_valid_email(email):
                    messagebox.showerror('Error', '❌ Invalid email format!')
                    return
                address = fields['address'].get().strip()
                blood = fields['blood_type'].get().strip().upper()
                if blood not in ['A+', 'A-', 'B+', 'B-', 'AB+', 'AB-', 'O+', 'O-']:
                    messagebox.showerror('Error', '❌ Invalid blood type!')
                    return

                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute('''
                    UPDATE patients SET
                        national_id = ?, first_name = ?, last_name = ?,
                        age = ?, gender = ?, phone = ?, email = ?,
                        address = ?, blood_type = ?
                    WHERE id = ?
                ''', (nid, first, last, age, gender, phone, email, address, blood, real_id))
                conn.commit()
                conn.close()

                messagebox.showinfo('Success', f'✅ Patient "{first} {last}" updated!')
                try:
                    log_action('UPDATE', entity_type='patient',
                               details=f'Updated patient: {first} {last}',
                               user=getattr(self.root, 'current_user', None))
                except:
                    pass
                win.destroy()
                self.load_patients()
                self.refresh_dashboard()
            except sqlite3.IntegrityError:
                messagebox.showerror('Error', '❌ National ID already exists!')
            except Exception as e:
                messagebox.showerror('Error', f'❌ Error: {e}')

        # ===== ۶. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '💾 Save', 'command': save, 'color': 'success', 'width': 160},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#f0f0f0')

        # ===== ۷. Maximize =====
        win.update_idletasks()
        win.state('zoomed')

        fields['national_id'].focus()



    # ==========================================
    # Doctors Tab
    # ==========================================

    def create_doctors_tab(self):
        tk.Label(self.doctors_tab, text='👨‍⚕️ Doctor Management',
                 font=('Arial', 14, 'bold'), bg=WINDOW_BG, fg='#2c3e50').pack(pady=15)

        # ===== دکمه‌ها =====
        btn_bar = create_button_bar(self.doctors_tab, bg=WINDOW_BG)

        create_standard_button(btn_bar, '➕ Add', self.add_doctor_window, color='success')
        create_standard_button(btn_bar, '✏️ Edit', self.edit_doctor_window, color='warning')
        create_standard_button(btn_bar, '👁️ Refresh', self.load_doctors, color='primary')
        create_standard_button(btn_bar, '🗑️ Delete', self.delete_doctor, color='danger')

        # ===== ردیف ۱: Search =====
        row1 = tk.Frame(self.doctors_tab, bg=WINDOW_BG)
        row1.pack(pady=(5, 3), padx=40, fill='x')

        tk.Label(row1, text='🔍 Search:', font=('Arial', 11, 'bold'),
                 bg=WINDOW_BG, fg='#2c3e50').pack(side='left', padx=(0, 8))

        self.doctor_search_var = tk.StringVar()
        search_entry = tk.Entry(row1, textvariable=self.doctor_search_var,
                                font=('Arial', 11), bg='white',
                                relief='solid', bd=1, width=40)
        search_entry.pack(side='left', ipady=6, padx=(0, 8))

        self.doctor_search_var.trace_add('write',
                                         lambda *args: self.load_doctors())

        def clear_search():
            self.doctor_search_var.set('')

        tk.Button(row1, text='✖ Clear', command=clear_search,
                  bg='#95a5a6', fg='white', font=('Arial', 9, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=10, pady=4).pack(side='left')

        # ===== ردیف ۲: Count =====
        row2 = tk.Frame(self.doctors_tab, bg=WINDOW_BG)
        row2.pack(pady=(0, 10), padx=40, fill='x')

        self.doctor_count_var = tk.StringVar(value='')
        tk.Label(row2, textvariable=self.doctor_count_var,
                 font=('Arial', 10, 'italic'), bg=WINDOW_BG,
                 fg='#7f8c8d').pack(side='right')

        # ===== Treeview =====
        columns = ('ID', 'Name', 'Specialty', 'Phone', 'Fee')
        tree_frame = tk.Frame(self.doctors_tab, bg=WINDOW_BG)
        tree_frame.pack(pady=(0, 10), padx=40, fill='both', expand=True)

        self.doctors_tree = ttk.Treeview(tree_frame, columns=columns,
                                         show='headings', height=15,
                                         selectmode='extended')

        for col in columns:
            self.doctors_tree.heading(col, text=col)
            self.doctors_tree.column(col, width=170, anchor='center')

        scroll = ttk.Scrollbar(tree_frame, orient='vertical',
                               command=self.doctors_tree.yview)
        self.doctors_tree.configure(yscrollcommand=scroll.set)

        self.doctors_tree.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        enable_multi_click_selection(self.doctors_tree)

        self.load_doctors()

    def load_doctors(self):
        # ===== خواندن مقدار جستجو =====
        search_text = ''
        if hasattr(self, 'doctor_search_var'):
            search_text = self.doctor_search_var.get().strip()

        conn = get_connection()
        cursor = conn.cursor()

        if search_text:
            search_pattern = f'%{search_text}%'
            cursor.execute('''
                SELECT id, first_name, last_name, specialty, phone, consultation_fee
                FROM doctors
                WHERE
                    first_name LIKE ?
                    OR last_name LIKE ?
                    OR (first_name || ' ' || last_name) LIKE ?
                    OR specialty LIKE ?
                    OR phone LIKE ?
                    OR national_id LIKE ?
                ORDER BY id
            ''', (search_pattern, search_pattern, search_pattern,
                  search_pattern, search_pattern, search_pattern))
        else:
            cursor.execute('''
                SELECT id, first_name, last_name, specialty, phone, consultation_fee
                FROM doctors ORDER BY id
            ''')

        doctors = cursor.fetchall()
        conn.close()

        # پاک کردن
        for item in self.doctors_tree.get_children():
            self.doctors_tree.delete(item)

        # اضافه کردن
        for index, d in enumerate(doctors, 1):
            real_id = d[0]
            first_name = d[1].replace('Dr.', '').replace('dr.', '').strip()
            full_name = f'Dr. {first_name} {d[2]}'

            self.doctors_tree.insert(
                '', 'end',
                values=(index, full_name, d[3], d[4], f'{d[5]:,.0f}'),
                tags=(str(real_id),)
            )

        # آپدیت تعداد
        if hasattr(self, 'doctor_count_var'):
            if search_text:
                self.doctor_count_var.set(f'🔍 {len(doctors)} result(s)')
            else:
                self.doctor_count_var.set(f'📋 {len(doctors)} doctor(s)')

    def add_doctor_window(self):
        """فرم اضافه کردن دکتر جدید"""
        win = tk.Toplevel(self.root)
        win.title('➕ Add New Doctor')
        win.configure(bg='#f0f0f0')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین ثابت) =====
        btn_container = tk.Frame(win, bg='#f0f0f0')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#f0f0f0')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f0f0f0')

        # ===== ۳. اتصال اسکرول =====
        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۴. عنوان =====
        tk.Label(scrollable, text='➕ Add New Doctor',
                 font=('Arial', 16, 'bold'), bg='#f0f0f0',
                 fg='#2c3e50').pack(pady=15)

        # ===== ۵. فرم =====
        form_frame = tk.Frame(scrollable, bg='#f0f0f0')
        form_frame.pack(pady=10, padx=40, fill='x')

        fields = {}
        labels = [
            ('National ID *', 'national_id', '10 digits'),
            ('First Name *', 'first_name', 'Without Dr.'),
            ('Last Name *', 'last_name', 'Letters only'),
            ('Phone *', 'phone', '09xxxxxxxxx'),
            ('Email (Optional)', 'email', 'user@domain.com'),
            ('Available Days *', 'available_days', 'e.g., Mon,Wed,Fri'),
            ('Consultation Fee *', 'fee', 'e.g., 500000'),
        ]

        for i, (label, key, hint) in enumerate(labels):
            tk.Label(form_frame, text=label, bg='#f0f0f0',
                     font=('Arial', 10, 'bold'), anchor='w').grid(
                row=i * 2, column=0, sticky='w', pady=(5, 0), padx=5)
            entry = tk.Entry(form_frame, width=40, font=('Arial', 10))
            entry.grid(row=i * 2 + 1, column=0, pady=(0, 5), padx=5, sticky='ew')
            fields[key] = entry
            tk.Label(form_frame, text=f'💡 {hint}', bg='#f0f0f0',
                     font=('Arial', 8), fg='#7f8c8d', anchor='w').grid(
                row=i * 2 + 1, column=1, sticky='w', padx=5)

        # ===== تخصص =====
        tk.Label(form_frame, text='Specialty *', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=len(labels) * 2, column=0, sticky='w', pady=(5, 0), padx=5)

        specialties = ['Cardiology', 'Pediatrics', 'Orthopedics', 'Neurology',
                       'Dermatology', 'Internal Medicine', 'Surgery', 'Radiology']
        specialty_var = tk.StringVar(value=specialties[0])
        ttk.Combobox(form_frame, textvariable=specialty_var,
                     values=specialties, width=38, state='readonly').grid(
            row=len(labels) * 2 + 1, column=0, pady=(0, 5), padx=5, sticky='ew')

        form_frame.grid_columnconfigure(0, weight=1)

        # ===== ۶. تابع ذخیره =====
        def save():
            try:
                nid = fields['national_id'].get().strip()
                if not is_valid_national_id(nid):
                    messagebox.showerror('Error', '❌ National ID must be 10 digits!')
                    return
                first = fields['first_name'].get().strip()
                last = fields['last_name'].get().strip()
                if not is_valid_name(first) or not is_valid_name(last):
                    messagebox.showerror('Error', '❌ Name must contain only letters!')
                    return
                phone = fields['phone'].get().strip()
                if not is_valid_phone(phone):
                    messagebox.showerror('Error', '❌ Phone must start with 09!')
                    return
                email = fields['email'].get().strip()
                if email and not is_valid_email(email):
                    messagebox.showerror('Error', '❌ Invalid email format!')
                    return
                days = fields['available_days'].get().strip()
                if not days:
                    messagebox.showerror('Error', '❌ Available days is empty!')
                    return
                try:
                    fee = float(fields['fee'].get().strip())
                    if fee < 0: raise ValueError
                except ValueError:
                    messagebox.showerror('Error', '❌ Fee must be a positive number!')
                    return

                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO doctors (national_id, first_name, last_name, specialty, phone, email, available_days, consultation_fee)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (nid, first, last, specialty_var.get(), phone, email, days, fee))
                conn.commit()
                conn.close()

                messagebox.showinfo('Success', f'✅ Doctor "Dr. {first} {last}" added!')
                try:
                    log_action('CREATE', entity_type='doctor', details=f'Added doctor: Dr. {first} {last}',
                               user=getattr(self.root, 'current_user', None))
                except:
                    pass
                win.destroy()
                self.load_doctors()
                self.refresh_dashboard()
            except sqlite3.IntegrityError:
                messagebox.showerror('Error', '❌ National ID already exists!')
            except Exception as e:
                messagebox.showerror('Error', f'❌ Error: {e}')

        # ===== ۷. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '💾 Save', 'command': save, 'color': 'success', 'width': 160},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#f0f0f0')

        # ===== ۸. Maximize (آخرین مرحله) =====
        win.update_idletasks()
        win.state('zoomed')

        fields['national_id'].focus()


    def delete_doctor(self):
        selected = self.doctors_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', 'Please select at least one doctor first!')
            return

        count = len(selected)

        if count == 1:
            values = self.doctors_tree.item(selected[0])['values']
            confirm_msg = f'Delete "{values[1]}"?'
        else:
            confirm_msg = f'Delete {count} doctor(s)?\n\nThis action cannot be undone!'

        if not messagebox.askyesno('Confirm', confirm_msg):
            return

        try:
            conn = get_connection()
            cursor = conn.cursor()

            for item in selected:
                real_id = self.doctors_tree.item(item)['tags'][0]
                cursor.execute('DELETE FROM doctors WHERE id = ?', (real_id,))

            conn.commit()
            conn.close()

            messagebox.showinfo('Success', f'✅ {count} doctor(s) deleted!')
            self.load_doctors()
            self.refresh_dashboard()

        except Exception as e:
            messagebox.showerror('Error', f'❌ Error: {e}')

    def edit_doctor_window(self):
        """فرم ویرایش دکتر"""
        selected = self.doctors_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select a doctor first!')
            return

        real_id = self.doctors_tree.item(selected[0])['tags'][0]

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT national_id, first_name, last_name, specialty,
                   phone, email, available_days, consultation_fee
            FROM doctors WHERE id = ?
        ''', (real_id,))
        d = cursor.fetchone()
        conn.close()

        if not d:
            messagebox.showerror('Error', '❌ Doctor not found!')
            return

        win = tk.Toplevel(self.root)
        win.title('✏️ Edit Doctor')
        win.configure(bg='#f0f0f0')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#f0f0f0')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#f0f0f0')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f0f0f0')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text='✏️ Edit Doctor',
                 font=('Arial', 16, 'bold'), bg='#f0f0f0',
                 fg='#2c3e50').pack(pady=15)

        # ===== ۴. فرم =====
        form_frame = tk.Frame(scrollable, bg='#f0f0f0')
        form_frame.pack(pady=10, padx=40, fill='x')

        fields = {}
        labels = [
            ('National ID *', 'national_id', '10 digits'),
            ('First Name *', 'first_name', 'Without Dr.'),
            ('Last Name *', 'last_name', 'Letters only'),
            ('Phone *', 'phone', '09xxxxxxxxx'),
            ('Email (Optional)', 'email', 'user@domain.com'),
            ('Available Days *', 'available_days', 'e.g., Mon,Wed,Fri'),
            ('Consultation Fee *', 'fee', 'e.g., 500000'),
        ]

        initial_values = {
            'national_id': d[0],
            'first_name': d[1].replace('Dr.', '').replace('dr.', '').strip(),
            'last_name': d[2],
            'phone': d[4],
            'email': d[5] or '',
            'available_days': d[6] or '',
            'fee': str(int(d[7])),
        }

        for i, (label, key, hint) in enumerate(labels):
            tk.Label(form_frame, text=label, bg='#f0f0f0',
                     font=('Arial', 10, 'bold'), anchor='w').grid(
                row=i * 2, column=0, sticky='w', pady=(5, 0), padx=5)
            entry = tk.Entry(form_frame, width=40, font=('Arial', 10))
            entry.grid(row=i * 2 + 1, column=0, pady=(0, 5), padx=5, sticky='ew')
            entry.insert(0, initial_values[key])
            fields[key] = entry
            tk.Label(form_frame, text=f'💡 {hint}', bg='#f0f0f0',
                     font=('Arial', 8), fg='#7f8c8d', anchor='w').grid(
                row=i * 2 + 1, column=1, sticky='w', padx=5)

        # ===== تخصص =====
        tk.Label(form_frame, text='Specialty *', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=len(labels) * 2, column=0, sticky='w', pady=(5, 0), padx=5)

        specialties = ['Cardiology', 'Pediatrics', 'Orthopedics', 'Neurology',
                       'Dermatology', 'Internal Medicine', 'Surgery', 'Radiology']
        specialty_var = tk.StringVar(value=d[3])
        ttk.Combobox(form_frame, textvariable=specialty_var,
                     values=specialties, width=38, state='readonly').grid(
            row=len(labels) * 2 + 1, column=0, pady=(0, 5), padx=5, sticky='ew')

        form_frame.grid_columnconfigure(0, weight=1)

        # ===== ۵. تابع ذخیره =====
        def save():
            try:
                nid = fields['national_id'].get().strip()
                if not is_valid_national_id(nid):
                    messagebox.showerror('Error', '❌ National ID must be 10 digits!')
                    return
                first = fields['first_name'].get().strip()
                last = fields['last_name'].get().strip()
                if not is_valid_name(first) or not is_valid_name(last):
                    messagebox.showerror('Error', '❌ Name must contain only letters!')
                    return
                phone = fields['phone'].get().strip()
                if not is_valid_phone(phone):
                    messagebox.showerror('Error', '❌ Phone must start with 09!')
                    return
                email = fields['email'].get().strip()
                if email and not is_valid_email(email):
                    messagebox.showerror('Error', '❌ Invalid email format!')
                    return
                days = fields['available_days'].get().strip()
                if not days:
                    messagebox.showerror('Error', '❌ Available days is empty!')
                    return
                try:
                    fee = float(fields['fee'].get().strip())
                    if fee < 0: raise ValueError
                except ValueError:
                    messagebox.showerror('Error', '❌ Fee must be a positive number!')
                    return

                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute('''
                    UPDATE doctors SET
                        national_id = ?, first_name = ?, last_name = ?,
                        specialty = ?, phone = ?, email = ?,
                        available_days = ?, consultation_fee = ?
                    WHERE id = ?
                ''', (nid, first, last, specialty_var.get(), phone, email, days, fee, real_id))
                conn.commit()
                conn.close()

                messagebox.showinfo('Success', f'✅ Doctor "Dr. {first} {last}" updated!')
                try:
                    log_action('UPDATE', entity_type='doctor',
                               details=f'Updated doctor: Dr. {first} {last}',
                               user=getattr(self.root, 'current_user', None))
                except:
                    pass
                win.destroy()
                self.load_doctors()
                self.refresh_dashboard()
            except sqlite3.IntegrityError:
                messagebox.showerror('Error', '❌ National ID already exists!')
            except Exception as e:
                messagebox.showerror('Error', f'❌ Error: {e}')

        # ===== ۶. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '💾 Save', 'command': save, 'color': 'success', 'width': 160},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#f0f0f0')

        # ===== ۷. Maximize =====
        win.update_idletasks()
        win.state('zoomed')

        fields['national_id'].focus()



    def create_medical_records_tab(self):
        tk.Label(self.medical_records_tab, text='📋 Medical Records',
                 font=('Arial', 14, 'bold'), bg=WINDOW_BG, fg='#2c3e50').pack(pady=15)

        # ===== دکمه‌ها =====
        btn_bar1 = create_button_bar(self.medical_records_tab, bg=WINDOW_BG)

        create_standard_button(btn_bar1, '➕ New', self.add_medical_record_window, color='success')
        create_standard_button(btn_bar1, '✏️ Edit', self.edit_medical_record_window, color='warning')
        create_standard_button(btn_bar1, '👁️ View', self.view_medical_record_window, color='primary')
        create_standard_button(btn_bar1, '🖨️ Print', self.print_medical_record, color='secondary')
        create_standard_button(btn_bar1, '🖨️ Batch Print', self.batch_print_prescriptions, color='info')
        create_standard_button(btn_bar1, '🗑️ Delete', self.delete_medical_record, color='danger')

        # ===== ردیف ۱: Search =====
        row1 = tk.Frame(self.medical_records_tab, bg=WINDOW_BG)
        row1.pack(pady=(5, 3), padx=40, fill='x')

        tk.Label(row1, text='🔍 Search:', font=('Arial', 11, 'bold'),
                 bg=WINDOW_BG, fg='#2c3e50').pack(side='left', padx=(0, 8))

        self.medical_search_var = tk.StringVar()
        search_entry = tk.Entry(row1, textvariable=self.medical_search_var,
                                font=('Arial', 11), bg='white',
                                relief='solid', bd=1, width=30)
        search_entry.pack(side='left', ipady=6, padx=(0, 8))

        self.medical_search_var.trace_add('write',
                                          lambda *args: self.load_medical_records())

        def clear_search():
            self.medical_search_var.set('')

        tk.Button(row1, text='✖ Clear', command=clear_search,
                  bg='#95a5a6', fg='white', font=('Arial', 9, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=10, pady=4).pack(side='left')

        # ===== ردیف ۲: Date Filter + Count =====
        row2 = tk.Frame(self.medical_records_tab, bg=WINDOW_BG)
        row2.pack(pady=(0, 10), padx=40, fill='x')

        filter_var, from_var, to_var, df_container, custom_frame = create_date_filter(
            row2, WINDOW_BG, lambda: self.load_medical_records()
        )
        self.medical_date_filter = filter_var
        self.medical_date_from = from_var
        self.medical_date_to = to_var
        df_container.pack(side='left')

        self.medical_count_var = tk.StringVar(value='')
        tk.Label(row2, textvariable=self.medical_count_var,
                 font=('Arial', 10, 'italic'), bg=WINDOW_BG,
                 fg='#7f8c8d').pack(side='right')

        # ===== Treeview =====
        columns = ('ID', 'Date', 'Patient', 'Doctor', 'Diagnosis', 'Has Rx')

        tree_frame = tk.Frame(self.medical_records_tab, bg=WINDOW_BG)
        tree_frame.pack(fill='both', expand=True, padx=40, pady=(5, 10))

        self.medical_tree = ttk.Treeview(tree_frame, columns=columns,
                                         show='headings', height=14,
                                         selectmode='extended')

        col_widths = {
            'ID': 60, 'Date': 120, 'Patient': 200,
            'Doctor': 200, 'Diagnosis': 300, 'Has Rx': 100,
        }

        for col in columns:
            self.medical_tree.heading(col, text=col)
            self.medical_tree.column(col, width=col_widths[col], anchor='center')

        scroll = ttk.Scrollbar(tree_frame, orient='vertical',
                               command=self.medical_tree.yview)
        self.medical_tree.configure(yscrollcommand=scroll.set)

        self.medical_tree.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        self.medical_tree.bind('<Double-Button-1>',
                               lambda e: self.view_medical_record_window())

        enable_multi_click_selection(self.medical_tree)

        self.load_medical_records()


    def load_medical_records(self):
        """بارگذاری پرونده‌های پزشکی"""
        search_text = ''
        if hasattr(self, 'medical_search_var'):
            search_text = self.medical_search_var.get().strip()

        from_date, to_date = None, None
        if hasattr(self, 'medical_date_filter'):
            from_date, to_date = get_date_range(
                self.medical_date_filter.get(),
                self.medical_date_from.get().strip() if hasattr(self, 'medical_date_from') else '',
                self.medical_date_to.get().strip() if hasattr(self, 'medical_date_to') else '',
            )

        conn = get_connection()
        cursor = conn.cursor()

        query = '''
            SELECT
                mr.id,
                mr.record_date,
                p.first_name || ' ' || p.last_name AS patient_name,
                d.first_name || ' ' || d.last_name AS doctor_name,
                mr.diagnosis,
                (SELECT COUNT(*) FROM prescriptions WHERE record_id = mr.id) AS rx_count
            FROM medical_records mr
            JOIN patients p ON mr.patient_id = p.id
            JOIN doctors d ON mr.doctor_id = d.id
            WHERE 1=1
        '''
        params = []

        if search_text:
            pattern = f'%{search_text}%'
            query += '''
                AND (
                    p.first_name LIKE ?
                    OR p.last_name LIKE ?
                    OR (p.first_name || ' ' || p.last_name) LIKE ?
                    OR d.first_name LIKE ?
                    OR d.last_name LIKE ?
                    OR mr.diagnosis LIKE ?
                    OR mr.chief_complaint LIKE ?
                )
            '''
            params.extend([pattern] * 7)

        if from_date:
            query += ' AND mr.record_date >= ?'
            params.append(from_date)
        if to_date:
            query += ' AND mr.record_date <= ?'
            params.append(to_date)

        query += ' ORDER BY mr.record_date DESC, mr.id DESC'

        cursor.execute(query, params)
        records = cursor.fetchall()
        conn.close()

        # پاک کردن
        for item in self.medical_tree.get_children():
            self.medical_tree.delete(item)

        for r in records:
            record_id, record_date, patient, doctor, diagnosis, rx_count = r

            doctor_clean = doctor.replace('Dr.', '').replace('dr.', '').strip()
            rx_display = f'💊 {rx_count}' if rx_count > 0 else '—'

            self.medical_tree.insert('', 'end', values=(
                record_id, record_date, patient, f'Dr. {doctor_clean}',
                diagnosis or '—', rx_display
            ), tags=(str(record_id),))

        if hasattr(self, 'medical_count_var'):
            active = search_text or (hasattr(self, 'medical_date_filter')
                                     and self.medical_date_filter.get() != 'All')
            prefix = '🔍' if active else '📋'
            self.medical_count_var.set(f'{prefix} {len(records)} record(s)')

    def add_medical_record_window(self):
        """فرم ثبت پرونده‌ی پزشکی جدید"""
        win = tk.Toplevel(self.root)
        win.title('➕ New Medical Record')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین ثابت) =====
        btn_container = tk.Frame(win, bg='#ecf0f1')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#ecf0f1', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#ecf0f1')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text='➕ New Medical Record',
                 font=('Arial', 16, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(pady=15)

        # ===== ۴. خواندن بیماران و دکترها =====
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT id, first_name, last_name, national_id FROM patients ORDER BY last_name')
        patients = cursor.fetchall()
        cursor.execute('SELECT id, first_name, last_name, specialty FROM doctors ORDER BY last_name')
        doctors = cursor.fetchall()
        conn.close()

        if not patients or not doctors:
            messagebox.showerror('Error', '❌ Need at least 1 patient and 1 doctor!')
            win.destroy()
            return

        # ===== ۵. Patient =====
        row1 = tk.Frame(scrollable, bg='#ecf0f1')
        row1.pack(fill='x', pady=5, padx=20)

        tk.Label(row1, text='👤 Patient *:', font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', width=15, anchor='w').pack(side='left')

        patient_options = [f'{p[0]} - {p[1]} {p[2]} ({p[3]})' for p in patients]
        patient_var = tk.StringVar(value=patient_options[0])
        ttk.Combobox(row1, textvariable=patient_var, values=patient_options,
                     state='readonly', width=55, font=('Arial', 10)).pack(side='left', padx=5)

        # ===== ۶. Doctor =====
        row2 = tk.Frame(scrollable, bg='#ecf0f1')
        row2.pack(fill='x', pady=5, padx=20)

        tk.Label(row2, text='👨‍⚕️ Doctor *:', font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', width=15, anchor='w').pack(side='left')

        doctor_options = [f'{d[0]} - Dr. {d[1]} {d[2]} ({d[3]})' for d in doctors]
        doctor_var = tk.StringVar(value=doctor_options[0])
        ttk.Combobox(row2, textvariable=doctor_var, values=doctor_options,
                     state='readonly', width=55, font=('Arial', 10)).pack(side='left', padx=5)

        # ===== ۷. Date =====
        row3 = tk.Frame(scrollable, bg='#ecf0f1')
        row3.pack(fill='x', pady=5, padx=20)

        tk.Label(row3, text='📅 Date *:', font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', width=15, anchor='w').pack(side='left')

        from datetime import datetime
        date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        tk.Entry(row3, textvariable=date_var, font=('Arial', 11),
                 width=20, bg='white').pack(side='left', padx=5)
        tk.Label(row3, text='(YYYY-MM-DD)', font=('Arial', 9, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(side='left')

        # ===== ۸. Vital Signs =====
        vital_outer = tk.Frame(scrollable, bg='#ecf0f1')
        vital_outer.pack(fill='x', pady=10, padx=20)

        vital_frame = tk.LabelFrame(vital_outer, text=' 🩺 Vital Signs ',
                                    font=('Arial', 11, 'bold'),
                                    bg='#ecf0f1', fg='#2c3e50',
                                    padx=15, pady=10)
        vital_frame.pack(fill='x')

        # ردیف ۱
        tk.Label(vital_frame, text='Blood Pressure:', bg='#ecf0f1',
                 font=('Arial', 10), anchor='w', width=15).grid(
            row=0, column=0, sticky='w', pady=3)

        bp_var = tk.StringVar()
        tk.Entry(vital_frame, textvariable=bp_var, font=('Arial', 10),
                 width=15, bg='white').grid(row=0, column=1, padx=5, pady=3, sticky='w')

        tk.Label(vital_frame, text='(e.g., 120/80)', font=('Arial', 8, 'italic'),
                 bg='#ecf0f1', fg='#95a5a6').grid(row=0, column=2, sticky='w', padx=(0, 40))

        tk.Label(vital_frame, text='🌡️ Temp (°C):', bg='#ecf0f1',
                 font=('Arial', 10), anchor='w', width=13).grid(
            row=0, column=3, sticky='w', padx=(0, 5), pady=3)

        temp_var = tk.StringVar()
        tk.Entry(vital_frame, textvariable=temp_var, font=('Arial', 10),
                 width=10, bg='white').grid(row=0, column=4, padx=5, pady=3, sticky='w')

        # ردیف ۲
        tk.Label(vital_frame, text='⚖️ Weight (kg):', bg='#ecf0f1',
                 font=('Arial', 10), anchor='w', width=15).grid(
            row=1, column=0, sticky='w', pady=3)

        weight_var = tk.StringVar()
        tk.Entry(vital_frame, textvariable=weight_var, font=('Arial', 10),
                 width=15, bg='white').grid(row=1, column=1, padx=5, pady=3, sticky='w')

        tk.Label(vital_frame, text='📏 Height (cm):', bg='#ecf0f1',
                 font=('Arial', 10), anchor='w', width=13).grid(
            row=1, column=3, sticky='w', padx=(0, 5), pady=3)

        height_var = tk.StringVar()
        tk.Entry(vital_frame, textvariable=height_var, font=('Arial', 10),
                 width=10, bg='white').grid(row=1, column=4, padx=5, pady=3, sticky='w')

        # ===== ۹. Chief Complaint =====
        tk.Label(scrollable, text='💬 Chief Complaint:',
                 font=('Arial', 11, 'bold'), bg='#ecf0f1',
                 anchor='w').pack(anchor='w', pady=(10, 3), padx=20)
        complaint_text = tk.Text(scrollable, height=2, font=('Arial', 10),
                                 bg='white', relief='solid', bd=1)
        complaint_text.pack(fill='x', pady=3, padx=20)

        # ===== ۱۰. Symptoms =====
        tk.Label(scrollable, text='🔍 Symptoms:',
                 font=('Arial', 11, 'bold'), bg='#ecf0f1',
                 anchor='w').pack(anchor='w', pady=(10, 3), padx=20)
        symptoms_text = tk.Text(scrollable, height=2, font=('Arial', 10),
                                bg='white', relief='solid', bd=1)
        symptoms_text.pack(fill='x', pady=3, padx=20)

        # ===== ۱۱. Diagnosis =====
        tk.Label(scrollable, text='🩺 Diagnosis:',
                 font=('Arial', 11, 'bold'), bg='#ecf0f1',
                 anchor='w').pack(anchor='w', pady=(10, 3), padx=20)
        diagnosis_text = tk.Text(scrollable, height=2, font=('Arial', 10),
                                 bg='white', relief='solid', bd=1)
        diagnosis_text.pack(fill='x', pady=3, padx=20)

        # ===== ۱۲. Notes =====
        tk.Label(scrollable, text='📝 Notes:',
                 font=('Arial', 11, 'bold'), bg='#ecf0f1',
                 anchor='w').pack(anchor='w', pady=(10, 3), padx=20)
        notes_text = tk.Text(scrollable, height=2, font=('Arial', 10),
                             bg='white', relief='solid', bd=1)
        notes_text.pack(fill='x', pady=(3, 15), padx=20)

        # ===== ۱۳. تابع ذخیره =====
        def save():
            try:
                patient_id = int(patient_var.get().split(' - ')[0])
                doctor_id = int(doctor_var.get().split(' - ')[0])
                record_date = date_var.get().strip()

                if not re.match(r'^\d{4}-\d{2}-\d{2}$', record_date):
                    messagebox.showerror('Error', '❌ Invalid date! Use YYYY-MM-DD')
                    return

                try:
                    temp = float(temp_var.get()) if temp_var.get().strip() else None
                    weight = float(weight_var.get()) if weight_var.get().strip() else None
                    height = float(height_var.get()) if height_var.get().strip() else None
                except ValueError:
                    messagebox.showerror('Error', '❌ Temperature/Weight/Height must be numbers!')
                    return

                bp = bp_var.get().strip()
                complaint = complaint_text.get('1.0', 'end-1c').strip()
                symptoms = symptoms_text.get('1.0', 'end-1c').strip()
                diagnosis = diagnosis_text.get('1.0', 'end-1c').strip()
                notes = notes_text.get('1.0', 'end-1c').strip()

                if not diagnosis and not complaint:
                    messagebox.showerror('Error', '❌ Please enter at least Diagnosis or Chief Complaint!')
                    return

                from datetime import datetime
                now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO medical_records
                    (patient_id, doctor_id, record_date, chief_complaint,
                     diagnosis, symptoms, blood_pressure, temperature,
                     weight, height, notes, created_at)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (patient_id, doctor_id, record_date, complaint,
                      diagnosis, symptoms, bp, temp, weight, height, notes, now))

                record_id = cursor.lastrowid
                conn.commit()
                conn.close()

                messagebox.showinfo('✅ Success',
                                    f'Medical record #{record_id} saved!')

                try:
                    log_action('CREATE', entity_type='medical_record',
                               entity_id=record_id,
                               details=f'New medical record for patient #{patient_id}',
                               user=getattr(self.root, 'current_user', None))
                except:
                    pass

                win.destroy()
                self.load_medical_records()

            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        # ===== ۱۴. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '💾 Save', 'command': save, 'color': 'success', 'width': 180},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#ecf0f1')

        # ===== ۱۵. Maximize =====
        win.update_idletasks()
        win.state('zoomed')

    def edit_medical_record_window(self):
        """فرم ویرایش پرونده‌ی پزشکی"""
        selected = self.medical_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select a record first!')
            return

        record_id = self.medical_tree.item(selected[0])['tags'][0]

        # ===== خواندن اطلاعات فعلی =====
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT
                mr.patient_id, mr.doctor_id, mr.record_date,
                mr.chief_complaint, mr.diagnosis, mr.symptoms,
                mr.blood_pressure, mr.temperature, mr.weight,
                mr.height, mr.notes
            FROM medical_records mr
            WHERE mr.id = ?
        ''', (record_id,))
        record = cursor.fetchone()

        if not record:
            conn.close()
            messagebox.showerror('Error', '❌ Record not found!')
            return

        # ===== بیماران و دکترها =====
        cursor.execute('SELECT id, first_name, last_name, national_id FROM patients ORDER BY last_name')
        patients = cursor.fetchall()

        cursor.execute('SELECT id, first_name, last_name, specialty FROM doctors ORDER BY last_name')
        doctors = cursor.fetchall()

        conn.close()

        win = tk.Toplevel(self.root)
        win.title(f'✏️ Edit Medical Record #{record_id}')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#ecf0f1')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#ecf0f1', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#ecf0f1')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text=f'✏️ Edit Medical Record #{record_id}',
                 font=('Arial', 16, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(pady=15)

        # ===== ۴. Patient =====
        row1 = tk.Frame(scrollable, bg='#ecf0f1')
        row1.pack(fill='x', pady=5, padx=20)

        tk.Label(row1, text='👤 Patient *:', font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', width=15, anchor='w').pack(side='left')

        patient_options = [f'{p[0]} - {p[1]} {p[2]} ({p[3]})' for p in patients]
        current_patient = next((opt for opt in patient_options
                                if opt.startswith(f'{record[0]} - ')),
                               patient_options[0])
        patient_var = tk.StringVar(value=current_patient)
        ttk.Combobox(row1, textvariable=patient_var, values=patient_options,
                     state='readonly', width=55, font=('Arial', 10)).pack(side='left', padx=5)

        # ===== ۵. Doctor =====
        row2 = tk.Frame(scrollable, bg='#ecf0f1')
        row2.pack(fill='x', pady=5, padx=20)

        tk.Label(row2, text='👨‍⚕️ Doctor *:', font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', width=15, anchor='w').pack(side='left')

        doctor_options = [f'{d[0]} - Dr. {d[1]} {d[2]} ({d[3]})' for d in doctors]
        current_doctor = next((opt for opt in doctor_options
                               if opt.startswith(f'{record[1]} - ')),
                              doctor_options[0])
        doctor_var = tk.StringVar(value=current_doctor)
        ttk.Combobox(row2, textvariable=doctor_var, values=doctor_options,
                     state='readonly', width=55, font=('Arial', 10)).pack(side='left', padx=5)

        # ===== ۶. Date =====
        row3 = tk.Frame(scrollable, bg='#ecf0f1')
        row3.pack(fill='x', pady=5, padx=20)

        tk.Label(row3, text='📅 Date *:', font=('Arial', 11, 'bold'),
                 bg='#ecf0f1', width=15, anchor='w').pack(side='left')

        date_var = tk.StringVar(value=record[2])
        tk.Entry(row3, textvariable=date_var, font=('Arial', 11),
                 width=20, bg='white').pack(side='left', padx=5)
        tk.Label(row3, text='(YYYY-MM-DD)', font=('Arial', 9, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(side='left')

        # ===== ۷. Vital Signs =====
        vital_outer = tk.Frame(scrollable, bg='#ecf0f1')
        vital_outer.pack(fill='x', pady=10, padx=20)

        vital_frame = tk.LabelFrame(vital_outer, text=' 🩺 Vital Signs ',
                                    font=('Arial', 11, 'bold'),
                                    bg='#ecf0f1', fg='#2c3e50',
                                    padx=15, pady=10)
        vital_frame.pack(fill='x')

        # ردیف ۱
        tk.Label(vital_frame, text='Blood Pressure:', bg='#ecf0f1',
                 font=('Arial', 10), anchor='w', width=15).grid(
            row=0, column=0, sticky='w', pady=3)

        bp_var = tk.StringVar(value=record[6] or '')
        tk.Entry(vital_frame, textvariable=bp_var, font=('Arial', 10),
                 width=15, bg='white').grid(row=0, column=1, padx=5, pady=3, sticky='w')

        tk.Label(vital_frame, text='(e.g., 120/80)', font=('Arial', 8, 'italic'),
                 bg='#ecf0f1', fg='#95a5a6').grid(row=0, column=2, sticky='w', padx=(0, 40))

        tk.Label(vital_frame, text='🌡️ Temp (°C):', bg='#ecf0f1',
                 font=('Arial', 10), anchor='w', width=13).grid(
            row=0, column=3, sticky='w', padx=(0, 5), pady=3)

        temp_var = tk.StringVar(value=str(record[7]) if record[7] else '')
        tk.Entry(vital_frame, textvariable=temp_var, font=('Arial', 10),
                 width=10, bg='white').grid(row=0, column=4, padx=5, pady=3, sticky='w')

        # ردیف ۲
        tk.Label(vital_frame, text='⚖️ Weight (kg):', bg='#ecf0f1',
                 font=('Arial', 10), anchor='w', width=15).grid(
            row=1, column=0, sticky='w', pady=3)

        weight_var = tk.StringVar(value=str(record[8]) if record[8] else '')
        tk.Entry(vital_frame, textvariable=weight_var, font=('Arial', 10),
                 width=15, bg='white').grid(row=1, column=1, padx=5, pady=3, sticky='w')

        tk.Label(vital_frame, text='📏 Height (cm):', bg='#ecf0f1',
                 font=('Arial', 10), anchor='w', width=13).grid(
            row=1, column=3, sticky='w', padx=(0, 5), pady=3)

        height_var = tk.StringVar(value=str(record[9]) if record[9] else '')
        tk.Entry(vital_frame, textvariable=height_var, font=('Arial', 10),
                 width=10, bg='white').grid(row=1, column=4, padx=5, pady=3, sticky='w')

        # ===== ۸. Chief Complaint =====
        tk.Label(scrollable, text='💬 Chief Complaint:',
                 font=('Arial', 11, 'bold'), bg='#ecf0f1',
                 anchor='w').pack(anchor='w', pady=(10, 3), padx=20)
        complaint_text = tk.Text(scrollable, height=2, font=('Arial', 10),
                                 bg='white', relief='solid', bd=1)
        complaint_text.pack(fill='x', pady=3, padx=20)
        if record[3]:
            complaint_text.insert('1.0', record[3])

        # ===== ۹. Symptoms =====
        tk.Label(scrollable, text='🔍 Symptoms:',
                 font=('Arial', 11, 'bold'), bg='#ecf0f1',
                 anchor='w').pack(anchor='w', pady=(10, 3), padx=20)
        symptoms_text = tk.Text(scrollable, height=2, font=('Arial', 10),
                                bg='white', relief='solid', bd=1)
        symptoms_text.pack(fill='x', pady=3, padx=20)
        if record[5]:
            symptoms_text.insert('1.0', record[5])

        # ===== ۱۰. Diagnosis =====
        tk.Label(scrollable, text='🩺 Diagnosis:',
                 font=('Arial', 11, 'bold'), bg='#ecf0f1',
                 anchor='w').pack(anchor='w', pady=(10, 3), padx=20)
        diagnosis_text = tk.Text(scrollable, height=2, font=('Arial', 10),
                                 bg='white', relief='solid', bd=1)
        diagnosis_text.pack(fill='x', pady=3, padx=20)
        if record[4]:
            diagnosis_text.insert('1.0', record[4])

        # ===== ۱۱. Notes =====
        tk.Label(scrollable, text='📝 Notes:',
                 font=('Arial', 11, 'bold'), bg='#ecf0f1',
                 anchor='w').pack(anchor='w', pady=(10, 3), padx=20)
        notes_text = tk.Text(scrollable, height=2, font=('Arial', 10),
                             bg='white', relief='solid', bd=1)
        notes_text.pack(fill='x', pady=(3, 15), padx=20)
        if record[10]:
            notes_text.insert('1.0', record[10])

        # ===== ۱۲. تابع ذخیره =====
        def save():
            try:
                patient_id = int(patient_var.get().split(' - ')[0])
                doctor_id = int(doctor_var.get().split(' - ')[0])
                record_date = date_var.get().strip()

                if not re.match(r'^\d{4}-\d{2}-\d{2}$', record_date):
                    messagebox.showerror('Error', '❌ Invalid date! Use YYYY-MM-DD')
                    return

                try:
                    temp = float(temp_var.get()) if temp_var.get().strip() else None
                    weight = float(weight_var.get()) if weight_var.get().strip() else None
                    height = float(height_var.get()) if height_var.get().strip() else None
                except ValueError:
                    messagebox.showerror('Error', '❌ Temperature/Weight/Height must be numbers!')
                    return

                bp = bp_var.get().strip()
                complaint = complaint_text.get('1.0', 'end-1c').strip()
                symptoms = symptoms_text.get('1.0', 'end-1c').strip()
                diagnosis = diagnosis_text.get('1.0', 'end-1c').strip()
                notes = notes_text.get('1.0', 'end-1c').strip()

                if not diagnosis and not complaint:
                    messagebox.showerror('Error', '❌ Please enter at least Diagnosis or Chief Complaint!')
                    return

                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute('''
                    UPDATE medical_records
                    SET patient_id = ?, doctor_id = ?, record_date = ?,
                        chief_complaint = ?, diagnosis = ?, symptoms = ?,
                        blood_pressure = ?, temperature = ?, weight = ?,
                        height = ?, notes = ?
                    WHERE id = ?
                ''', (patient_id, doctor_id, record_date, complaint,
                      diagnosis, symptoms, bp, temp, weight, height, notes,
                      record_id))
                conn.commit()
                conn.close()

                messagebox.showinfo('✅ Success',
                                    f'Medical record #{record_id} updated successfully!')

                try:
                    log_action('UPDATE', entity_type='medical_record',
                               entity_id=record_id,
                               details=f'Updated medical record #{record_id}',
                               user=getattr(self.root, 'current_user', None))
                except:
                    pass

                win.destroy()
                self.load_medical_records()

            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        # ===== ۱۳. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '💾 Save Changes', 'command': save, 'color': 'success', 'width': 200},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#ecf0f1')

        # ===== ۱۴. Maximize =====
        win.update_idletasks()
        win.state('zoomed')

    def view_medical_record_window(self):
        """نمایش کامل پرونده‌ی پزشکی همراه با آزمایشات و پرداخت ویزیت"""
        selected = self.medical_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select a record first!')
            return

        record_id = self.medical_tree.item(selected[0])['tags'][0]

        # ===== خواندن اطلاعات =====
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT
                mr.id, mr.record_date, mr.chief_complaint, mr.diagnosis,
                mr.symptoms, mr.blood_pressure, mr.temperature,
                mr.weight, mr.height, mr.notes, mr.created_at,
                p.id AS patient_id, p.first_name || ' ' || p.last_name AS patient_name,
                p.national_id, p.phone, p.age, p.gender, p.blood_type,
                d.id AS doctor_id, d.first_name || ' ' || d.last_name AS doctor_name,
                d.specialty, d.consultation_fee
            FROM medical_records mr
            JOIN patients p ON mr.patient_id = p.id
            JOIN doctors d ON mr.doctor_id = d.id
            WHERE mr.id = ?
        ''', (record_id,))
        record = cursor.fetchone()

        cursor.execute('''
            SELECT
                pr.id, m.name, m.category, pr.dosage, pr.frequency,
                pr.duration, pr.quantity, pr.instructions
            FROM prescriptions pr
            JOIN medicines m ON pr.medicine_id = m.id
            WHERE pr.record_id = ?
            ORDER BY pr.id
        ''', (record_id,))
        prescriptions = cursor.fetchall()

        # خواندن آزمایشات مرتبط با این پرونده
        cursor.execute('''
            SELECT lr.id, lt.name, lr.status, lr.result, lr.request_date
            FROM lab_requests lr
            JOIN lab_tests lt ON lr.test_id = lt.id
            WHERE lr.record_id = ?
            ORDER BY lr.id
        ''', (record_id,))
        lab_requests = cursor.fetchall()

        conn.close()

        if not record:
            messagebox.showerror('Error', '❌ Record not found!')
            return

        win = tk.Toplevel(self.root)
        win.title(f'📋 Medical Record #{record_id}')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        btn_container = tk.Frame(win, bg='#ecf0f1')
        btn_container.pack(side='bottom', fill='x', pady=15)

        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#ecf0f1', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#ecf0f1')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        tk.Label(scrollable, text=f'📋 Medical Record #{record_id}',
                 font=('Arial', 16, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(pady=15)

        def section(parent, title, color='#3498db'):
            sec = tk.LabelFrame(parent, text=f' {title} ',
                                font=('Arial', 11, 'bold'),
                                bg='#ecf0f1', fg=color,
                                padx=15, pady=10)
            sec.pack(fill='x', pady=8, padx=20)
            return sec

        def info_row(parent, label, value):
            row = tk.Frame(parent, bg='#ecf0f1')
            row.pack(fill='x', pady=2)
            tk.Label(row, text=label, font=('Arial', 10, 'bold'),
                     bg='#ecf0f1', width=18, anchor='w').pack(side='left')
            tk.Label(row, text=value or '—', font=('Arial', 10),
                     bg='#ecf0f1', anchor='w', justify='left',
                     wraplength=600).pack(side='left', fill='x', expand=True)

        # Patient Info
        sec_patient = section(scrollable, '👤 Patient Information', '#27ae60')
        info_row(sec_patient, 'Name:', record[12])
        info_row(sec_patient, 'National ID:', record[13])
        info_row(sec_patient, 'Phone:', record[14])
        info_row(sec_patient, 'Age:', f'{record[15]} years')
        info_row(sec_patient, 'Gender:', record[16])
        info_row(sec_patient, 'Blood Type:', record[17])

        # Doctor Info
        sec_doctor = section(scrollable, '👨‍⚕️ Doctor Information', '#9b59b6')
        doctor_clean = record[19].replace('Dr.', '').replace('dr.', '').strip()
        info_row(sec_doctor, 'Name:', f'Dr. {doctor_clean}')
        info_row(sec_doctor, 'Specialty:', record[20])
        info_row(sec_doctor, 'Visit Date:', record[1])
        info_row(sec_doctor, 'Consultation Fee:', f'{record[21]:,.0f} IRR')

        # Vital Signs
        sec_vital = section(scrollable, '🩺 Vital Signs', '#e67e22')
        vital_data = [
            ('Blood Pressure:', record[5] or '—'),
            ('Temperature:', f'{record[6]} °C' if record[6] else '—'),
            ('Weight:', f'{record[7]} kg' if record[7] else '—'),
            ('Height:', f'{record[8]} cm' if record[8] else '—'),
        ]
        vital_grid = tk.Frame(sec_vital, bg='#ecf0f1')
        vital_grid.pack(fill='x')
        for i, (lbl, val) in enumerate(vital_data):
            r = i // 2
            c = (i % 2) * 2
            tk.Label(vital_grid, text=lbl, font=('Arial', 10, 'bold'),
                     bg='#ecf0f1', width=15, anchor='w').grid(
                row=r, column=c, sticky='w', padx=(0, 5), pady=3)
            tk.Label(vital_grid, text=val, font=('Arial', 10),
                     bg='#ecf0f1', anchor='w').grid(
                row=r, column=c + 1, sticky='w', padx=(0, 40), pady=3)

        # Chief Complaint
        if record[2]:
            sec_cc = section(scrollable, '💬 Chief Complaint', '#3498db')
            tk.Label(sec_cc, text=record[2], font=('Arial', 10),
                     bg='#ecf0f1', justify='left', anchor='w',
                     wraplength=800).pack(fill='x', anchor='w')

        # Symptoms
        if record[4]:
            sec_sym = section(scrollable, '🔍 Symptoms', '#e74c3c')
            tk.Label(sec_sym, text=record[4], font=('Arial', 10),
                     bg='#ecf0f1', justify='left', anchor='w',
                     wraplength=800).pack(fill='x', anchor='w')

        # Diagnosis
        if record[3]:
            sec_diag = section(scrollable, '🩺 Diagnosis', '#27ae60')
            tk.Label(sec_diag, text=record[3], font=('Arial', 10, 'bold'),
                     bg='#ecf0f1', justify='left', anchor='w',
                     fg='#2c3e50', wraplength=800).pack(fill='x', anchor='w')

        # Prescriptions
        sec_rx = section(scrollable, f'💊 Prescriptions ({len(prescriptions)})', '#c0392b')
        if prescriptions:
            hdr = tk.Frame(sec_rx, bg='#c0392b')
            hdr.pack(fill='x', pady=(0, 3))
            for text, w in [('Medicine', 30), ('Dosage', 10), ('Freq', 8), ('Duration', 10), ('Qty', 6)]:
                tk.Label(hdr, text=text, font=('Arial', 9, 'bold'),
                         bg='#c0392b', fg='white', width=w, anchor='w').pack(side='left', padx=2)

            for i, pr in enumerate(prescriptions):
                bg = 'white' if i % 2 == 0 else '#f8f9fa'
                row = tk.Frame(sec_rx, bg=bg)
                row.pack(fill='x', pady=1)
                tk.Label(row, text=pr[1], font=('Arial', 9), bg=bg, width=30, anchor='w').pack(side='left', padx=2)
                tk.Label(row, text=pr[3] or '—', font=('Arial', 9), bg=bg, width=10, anchor='w').pack(side='left',
                                                                                                      padx=2)
                tk.Label(row, text=pr[4] or '—', font=('Arial', 9), bg=bg, width=8, anchor='w').pack(side='left',
                                                                                                     padx=2)
                tk.Label(row, text=pr[5] or '—', font=('Arial', 9), bg=bg, width=10, anchor='w').pack(side='left',
                                                                                                      padx=2)
                tk.Label(row, text=str(pr[6]), font=('Arial', 9, 'bold'), bg=bg, fg='#27ae60', width=6,
                         anchor='w').pack(side='left', padx=2)
        else:
            tk.Label(sec_rx, text='📭 No prescriptions yet', font=('Arial', 10, 'italic'),
                     bg='#ecf0f1', fg='#95a5a6').pack(pady=10)

        # Lab Tests Section
        sec_lab = section(scrollable, f'🧪 Lab Tests ({len(lab_requests)})', '#e67e22')
        if lab_requests:
            for lr in lab_requests:
                lr_id, test_name, status, result, req_date = lr
                l_row = tk.Frame(sec_lab, bg='white', relief='solid', bd=1)
                l_row.pack(fill='x', pady=2)

                status_icon = '✅' if status == 'Done' else '⏳' if status == 'Pending' else '🔄'

                tk.Label(l_row, text=f'📅 {req_date} | 🧪 {test_name}', font=('Arial', 10, 'bold'), bg='white',
                         anchor='w').pack(side='left', padx=10, pady=5)
                tk.Label(l_row, text=f'{status_icon} {status}', font=('Arial', 10), bg='white', fg='#e67e22').pack(
                    side='right', padx=10)

                if result:
                    res_row = tk.Frame(sec_lab, bg='#f8f9fa')
                    res_row.pack(fill='x', padx=10, pady=(0, 5))
                    tk.Label(res_row, text=f'↳ Result: {result}', font=('Arial', 9, 'italic'), bg='#f8f9fa',
                             fg='#27ae60', anchor='w', justify='left', wraplength=700).pack(fill='x', padx=10,
                                                                                            pady=3)
        else:
            tk.Label(sec_lab, text='📭 No lab tests requested', font=('Arial', 10, 'italic'),
                     bg='#ecf0f1', fg='#95a5a6').pack(pady=10)

        # Notes
        if record[9]:
            sec_notes = section(scrollable, '📝 Notes', '#7f8c8d')
            tk.Label(sec_notes, text=record[9], font=('Arial', 10),
                     bg='#ecf0f1', justify='left', anchor='w',
                     wraplength=800).pack(fill='x', anchor='w')

        tk.Label(scrollable, text=f'🕐 Created: {record[10]}',
                 font=('Arial', 8, 'italic'),
                 bg='#ecf0f1', fg='#95a5a6').pack(pady=10)

        # ===== تابع باز کردن پنجره پرداخت ویزیت =====
        def open_visit_payment_window():
            pay_win = tk.Toplevel(win)
            pay_win.title('💳 Visit Payment')
            pay_win.geometry('500x600')
            pay_win.configure(bg='#ecf0f1')
            pay_win.resizable(False, False)

            # وسط چین کردن
            pay_win.update_idletasks()
            w, h = 500, 600
            x = (pay_win.winfo_screenwidth() // 2) - (w // 2)
            y = (pay_win.winfo_screenheight() // 2) - (h // 2)
            pay_win.geometry(f'{w}x{h}+{x}+{y}')

            # هدر
            header = tk.Frame(pay_win, bg='#27ae60', height=55)
            header.pack(fill='x')
            header.pack_propagate(False)
            tk.Label(header, text='💳 Visit Payment', font=('Arial', 15, 'bold'),
                     bg='#27ae60', fg='white').pack(pady=12)

            # اطلاعات
            info_frame = tk.Frame(pay_win, bg='#f8f9fa', relief='solid', bd=1)
            info_frame.pack(fill='x', padx=20, pady=15)

            tk.Label(info_frame, text=f'👤 Patient: {record[12]}', font=('Arial', 11, 'bold'),
                     bg='#f8f9fa', anchor='w').pack(anchor='w', padx=15, pady=(10, 3))
            tk.Label(info_frame, text=f'👨‍⚕️ Doctor: Dr. {doctor_clean}', font=('Arial', 11),
                     bg='#f8f9fa', anchor='w').pack(anchor='w', padx=15, pady=(0, 10))

            # فرم
            form = tk.Frame(pay_win, bg='#ecf0f1')
            form.pack(fill='both', expand=True, padx=20, pady=5)

            tk.Label(form, text='Base Amount (IRR):', font=('Arial', 11, 'bold'),
                     bg='#ecf0f1', anchor='w').pack(anchor='w', pady=(10, 5))

            # مبلغ پیش فرض از دیتابیس خوانده شده
            amount_var = tk.StringVar(value=str(int(record[21])))
            amount_entry = tk.Entry(form, textvariable=amount_var, font=('Arial', 12), bg='white')
            amount_entry.pack(fill='x', ipady=6, pady=(0, 15))

            tk.Label(form, text='Discount (IRR):', font=('Arial', 11, 'bold'),
                     bg='#ecf0f1', anchor='w').pack(anchor='w', pady=(0, 5))
            discount_var = tk.StringVar(value='0')
            discount_entry = tk.Entry(form, textvariable=discount_var, font=('Arial', 12), bg='white')
            discount_entry.pack(fill='x', ipady=6, pady=(0, 15))

            tk.Label(form, text='Payment Method:', font=('Arial', 11, 'bold'),
                     bg='#ecf0f1', anchor='w').pack(anchor='w', pady=(0, 5))
            method_var = tk.StringVar(value='Cash')
            method_frame = tk.Frame(form, bg='#ecf0f1')
            method_frame.pack(fill='x', pady=(0, 15))
            for m in ['Cash', 'Card', 'Insurance']:
                tk.Radiobutton(method_frame, text=m, variable=method_var, value=m,
                               bg='#ecf0f1', font=('Arial', 10)).pack(side='left', padx=10)

            # جمع کل
            total_frame = tk.Frame(form, bg='white', relief='solid', bd=2)
            total_frame.pack(fill='x', pady=10)

            total_var = tk.StringVar()

            def update_total(*args):
                try:
                    amt = float(amount_var.get() or 0)
                    disc = float(discount_var.get() or 0)
                    total = max(amt - disc, 0)
                    total_var.set(f'{total:,.0f} IRR')
                except ValueError:
                    total_var.set('Invalid Input')

            amount_var.trace_add('write', update_total)
            discount_var.trace_add('write', update_total)
            update_total()

            tk.Label(total_frame, text='TOTAL:', font=('Arial', 14, 'bold'), bg='white').pack(side='left', padx=15,
                                                                                              pady=10)
            tk.Label(total_frame, textvariable=total_var, font=('Arial', 14, 'bold'), fg='#27ae60',
                     bg='white').pack(side='left')

            # دکمه ها
            def confirm_payment():
                try:
                    amt = float(amount_var.get())
                    disc = float(discount_var.get() or 0)
                    final = amt - disc

                    if final < 0:
                        messagebox.showerror('Error', '❌ Discount cannot exceed amount!')
                        return

                    import random
                    from datetime import datetime
                    receipt = f'VIS-{datetime.now().strftime("%Y%m%d%H%M%S")}-{random.randint(100, 999)}'

                    # ✅ استفاده از زمان محلی (هماهنگ با بقیه پرداخت‌ها)
                    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                    conn = get_connection()
                    cursor = conn.cursor()
                    cursor.execute('''
                        INSERT INTO payments
                        (receipt_number, patient_id, amount, discount, final_amount,
                         payment_method, payment_status, notes, payment_date)
                        VALUES (?, ?, ?, ?, ?, ?, 'Paid', ?, ?)
                    ''', (receipt, record[11], amt, disc, final, method_var.get(),
                          f'Visit for Record #{record_id}', now))
                    conn.commit()
                    conn.close()

                    messagebox.showinfo('✅ Success',
                                        f'Visit payment registered!\n\nReceipt: {receipt}\nTotal: {final:,.0f} IRR')

                    try:
                        log_action('PAYMENT', entity_type='visit_payment',
                                   details=f'Visit payment for record #{record_id}: {final:,.0f} IRR',
                                   user=getattr(self.root, 'current_user', None))
                    except:
                        pass

                    pay_win.destroy()

                    # ✅ رفرش خودکار تب Payments
                    try:
                        self.load_payments()
                    except Exception as ex:
                        print(f"Could not auto-refresh payments: {ex}")

                    # رفرش گزارش‌های داروخانه و آزمایشگاه
                    try:
                        self.load_pharmacy_stats()
                        self.load_lab_stats()
                    except:
                        pass

                    # چاپ رسید (اختیاری)
                    self._print_visit_receipt(receipt, record[12], final, doctor_clean, record[1])

                except ValueError:
                    messagebox.showerror('Error', '❌ Please enter valid numbers!')
                except Exception as e:
                    messagebox.showerror('Error', f'❌ {e}')

        # ✅ دکمه‌ها
        create_popup_buttons(btn_container, [
            {'text': '💊 Add Prescription',
             'command': lambda: [win.destroy(), self.add_prescription_window(record_id)],
             'color': 'warning', 'width': 180},
            {'text': '🧪 Request Lab Test',
             'command': lambda: [win.destroy(), self.request_lab_from_record(record_id)],
             'color': 'info', 'width': 180},
            {'text': '💳 Pay Visit',
             'command': open_visit_payment_window,
             'color': 'success', 'width': 150},
            {'text': '🖨️ Print',
             'command': lambda: self.print_medical_record_by_id(record_id),
             'color': 'secondary', 'width': 120},
            {'text': '❌ Close',
             'command': win.destroy,
             'color': 'danger', 'width': 120},
        ], bg='#ecf0f1')

        win.update_idletasks()
        win.state('zoomed')

    def _print_visit_receipt(self, receipt, patient_name, amount, doctor_name, date):
        """تابع ساخت، چاپ یا نمایش رسید پرداخت ویزیت"""
        from datetime import datetime
        import os

        # ساخت متن رسید
        content = (
            "========================================\n"
            "       HOSPITAL MANAGEMENT SYSTEM\n"
            "          VISIT PAYMENT RECEIPT\n"
            "========================================\n\n"
            f"Receipt #: {receipt}\n"
            f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"Patient: {patient_name}\n"
            f"Doctor: Dr. {doctor_name}\n"
            f"Visit Date: {date}\n\n"
            "----------------------------------------\n"
            f"AMOUNT PAID: {amount:,.0f} IRR\n"
            "----------------------------------------\n\n"
            "Thank you!\n"
            "========================================\n"
        )

        # ساخت فایل PDF
        success, pdf_path = print_text_file(content, f'visit_receipt_{receipt}.txt', open_pdf=False)

        if not success:
            messagebox.showerror('❌ PDF Error', f'Could not create PDF:\n{pdf_path}')
            return

        # بررسی وضعیت پرینتر و چاپ/نمایش فایل
        if WIN32_AVAILABLE:
            printer = get_printer_for('receipt')

            # چک کردن اینکه پرینتر مجازی است یا نه
            virtual_keywords = ['Print to PDF', 'XPS', 'OneNote', 'Fax', 'Virtual', 'Writer']
            is_virtual = any(kw.lower() in (printer or '').lower() for kw in virtual_keywords)

            if is_virtual or not printer:
                # اگر پرینتر مجازی بود یا نبود، فایل رو باز کن تا کاربر ببیند
                try:
                    os.startfile(pdf_path)
                except:
                    pass
                messagebox.showinfo('📄 PDF Ready', f'Receipt PDF created and opened.\n\nPath: {pdf_path}')
            else:
                # اگر پرینتر فیزیکی بود، چاپ کن
                ok, msg = print_pdf_to_printer(pdf_path, printer)
                if ok:
                    messagebox.showinfo('🖨️ Sent to Printer', f'Receipt sent to:\n🖨️ {printer}')
                else:
                    try:
                        os.startfile(pdf_path)
                    except:
                        pass
                    messagebox.showwarning('⚠️ Could Not Print', f'Error: {msg}\n\nPDF opened for manual printing.')
        else:
            # اگر ویندوز نبود، فقط باز کن
            try:
                os.startfile(pdf_path)
            except:
                pass
            messagebox.showinfo('📄 PDF Opened', 'PDF opened for manual printing.')


    def add_prescription_window(self, record_id):
        """فرم نسخه‌نویسی (با کادرهای Spinbox برای انتخاب آسان‌تر)"""
        win = tk.Toplevel(self.root)
        win.title(f'💊 Prescription for Record #{record_id}')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین ثابت) =====
        btn_container = tk.Frame(win, bg='#ecf0f1')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#ecf0f1', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#ecf0f1')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text=f'💊 Prescription for Record #{record_id}',
                 font=('Arial', 16, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(pady=15)

        # ===== ۴. خواندن اطلاعات =====
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT p.first_name || ' ' || p.last_name, mr.record_date, mr.diagnosis
            FROM medical_records mr
            JOIN patients p ON mr.patient_id = p.id
            WHERE mr.id = ?
        ''', (record_id,))
        info = cursor.fetchone()

        cursor.execute('SELECT id, name, price, stock_quantity FROM medicines WHERE stock_quantity > 0 ORDER BY name')
        medicines = cursor.fetchall()

        cursor.execute('''
            SELECT pr.id, m.name, pr.dosage, pr.frequency,
                   pr.duration, pr.quantity, pr.instructions,
                   pr.medicine_id
            FROM prescriptions pr
            JOIN medicines m ON pr.medicine_id = m.id
            WHERE pr.record_id = ?
            ORDER BY pr.id
        ''', (record_id,))
        existing_rx = cursor.fetchall()
        conn.close()

        if not medicines:
            messagebox.showerror('Error', '❌ No medicines in stock!')
            win.destroy()
            return

        # ===== ۵. Info bar =====
        info_bar = tk.Frame(scrollable, bg='#f8f9fa', relief='solid', bd=1)
        info_bar.pack(fill='x', pady=(0, 10), padx=20)

        if info:
            tk.Label(info_bar, text=f'👤 {info[0]}  |  📅 {info[1]}  |  🩺 {info[2] or "—"}',
                     font=('Arial', 10, 'bold'), bg='#f8f9fa',
                     fg='#2c3e50').pack(pady=8, padx=15, anchor='w')

        # ===== ۶. بخش اضافه کردن دارو =====
        add_frame = tk.LabelFrame(scrollable, text=' ➕ Add Medicine ',
                                  font=('Arial', 11, 'bold'),
                                  bg='#ecf0f1', fg='#27ae60',
                                  padx=15, pady=10)
        add_frame.pack(fill='x', pady=(0, 10), padx=20)

        # ردیف ۱: دارو
        r1 = tk.Frame(add_frame, bg='#ecf0f1')
        r1.pack(fill='x', pady=3)

        tk.Label(r1, text='💊 Medicine *:', font=('Arial', 10, 'bold'),
                 bg='#ecf0f1', width=14, anchor='w').pack(side='left')

        med_options = [f'{m[0]} - {m[1]} - Price: {m[2]:,.0f} | Stock: {m[3]}' for m in medicines]
        med_var = tk.StringVar(value=med_options[0])
        med_combo = ttk.Combobox(r1, textvariable=med_var, values=med_options,
                                 state='readonly', width=55, font=('Arial', 10))
        med_combo.pack(side='left', padx=5)

        # ===== پیش‌فرض داروها =====
        MEDICINE_DEFAULTS = {
            'Amoxicillin': {'dosage': '500mg', 'freq': '3x/day', 'duration': '7 days'},
            'Azithromycin': {'dosage': '250mg', 'freq': '1x/day', 'duration': '5 days'},
            'Ceftriaxone': {'dosage': '1g', 'freq': '2x/day', 'duration': '5 days'},
            'Ciprofloxacin': {'dosage': '500mg', 'freq': '2x/day', 'duration': '7 days'},
            'Metronidazole': {'dosage': '250mg', 'freq': '3x/day', 'duration': '7 days'},
            'Penicillin': {'dosage': '500mg', 'freq': '4x/day', 'duration': '10 days'},
            'Paracetamol': {'dosage': '500mg', 'freq': '3x/day', 'duration': '3 days'},
            'Ibuprofen': {'dosage': '400mg', 'freq': '3x/day', 'duration': '5 days'},
            'Naproxen': {'dosage': '250mg', 'freq': '2x/day', 'duration': '7 days'},
            'Aspirin': {'dosage': '100mg', 'freq': '1x/day', 'duration': '30 days'},
            'Diclofenac': {'dosage': '50mg', 'freq': '2x/day', 'duration': '5 days'},
            'Atorvastatin': {'dosage': '20mg', 'freq': '1x/day', 'duration': '30 days'},
            'Losartan': {'dosage': '50mg', 'freq': '1x/day', 'duration': '30 days'},
            'Metoprolol': {'dosage': '50mg', 'freq': '2x/day', 'duration': '30 days'},
            'Amlodipine': {'dosage': '5mg', 'freq': '1x/day', 'duration': '30 days'},
            'Metformin': {'dosage': '500mg', 'freq': '2x/day', 'duration': '30 days'},
            'Glibenclamide': {'dosage': '5mg', 'freq': '1x/day', 'duration': '30 days'},
            'Omeprazole': {'dosage': '20mg', 'freq': '1x/day', 'duration': '14 days'},
            'Pantoprazole': {'dosage': '40mg', 'freq': '1x/day', 'duration': '14 days'},
            'Ranitidine': {'dosage': '150mg', 'freq': '2x/day', 'duration': '14 days'},
            'Vitamin D3': {'dosage': '1000IU', 'freq': '1x/day', 'duration': '30 days'},
            'Vitamin C': {'dosage': '500mg', 'freq': '1x/day', 'duration': '30 days'},
            'Iron Supplement': {'dosage': '65mg', 'freq': '1x/day', 'duration': '30 days'},
            'Calcium': {'dosage': '500mg', 'freq': '2x/day', 'duration': '30 days'},
            'Folic Acid': {'dosage': '1mg', 'freq': '1x/day', 'duration': '30 days'},
            'Zinc': {'dosage': '50mg', 'freq': '1x/day', 'duration': '14 days'},
            'Cetirizine': {'dosage': '10mg', 'freq': '1x/day', 'duration': '7 days'},
            'Loratadine': {'dosage': '10mg', 'freq': '1x/day', 'duration': '7 days'},
            'Diphenhydramine': {'dosage': '25mg', 'freq': '3x/day', 'duration': '5 days'},
            'Salbutamol': {'dosage': '2 puff', 'freq': '4x/day', 'duration': '30 days'},
            'Montelukast': {'dosage': '10mg', 'freq': '1x/day', 'duration': '30 days'},
            'Gabapentin': {'dosage': '300mg', 'freq': '3x/day', 'duration': '14 days'},
            'Sertraline': {'dosage': '50mg', 'freq': '1x/day', 'duration': '30 days'},
            'Prednisolone': {'dosage': '5mg', 'freq': '1x/day', 'duration': '7 days'},
            'Dexamethasone': {'dosage': '0.5mg', 'freq': '2x/day', 'duration': '5 days'},
        }

        # ===== تعریف لیست‌های پیش‌فرض برای Spinbox ها =====
        dosage_options = ('1mg', '2.5mg', '5mg', '10mg', '20mg', '25mg', '50mg', '100mg', '250mg', '500mg', '1000mg', '1 tab', '2 tabs')
        freq_options = ('1x/day', '2x/day', '3x/day', '4x/day', '5x/day', 'every 6 hours', 'every 8 hours', 'every 12 hours', 'as needed')
        duration_options = ('1 day', '2 days', '3 days', '5 days', '7 days', '10 days', '14 days', '21 days', '30 days', '60 days', '90 days')

        def on_medicine_change(event=None):
            selected = med_var.get()
            try:
                name_part = selected.split(' - ', 1)[1]
                med_name = name_part.split(' - ', 1)[0].strip()
            except Exception:
                return

            matched = None
            for key, defaults in MEDICINE_DEFAULTS.items():
                if key.lower() in med_name.lower() or med_name.lower() in key.lower():
                    matched = defaults
                    break

            if matched:
                dosage_var.set(matched['dosage'])
                freq_var.set(matched['freq'])
                duration_var.set(matched['duration'])
                qty_var.set('1') # ✅ مقدار پیش‌فرض تعداد بسته همیشه 1 است
            else:
                dosage_var.set('')
                freq_var.set('')
                duration_var.set('')
                qty_var.set('1')

        med_combo.bind('<<ComboboxSelected>>', on_medicine_change)

        # ردیف ۲: دوز + تعداد در روز (استفاده از Spinbox)
        r2 = tk.Frame(add_frame, bg='#ecf0f1')
        r2.pack(fill='x', pady=3)

        tk.Label(r2, text='Dosage:', font=('Arial', 10, 'bold'),
                 bg='#ecf0f1', width=14, anchor='w').pack(side='left')
        dosage_var = tk.StringVar(value='500mg')
        tk.Spinbox(r2, textvariable=dosage_var, font=('Arial', 10),
                 width=15, bg='white', values=dosage_options).pack(side='left', padx=5)

        tk.Label(r2, text='Frequency:', font=('Arial', 10, 'bold'),
                 bg='#ecf0f1', width=11, anchor='w').pack(side='left', padx=(15, 0))
        freq_var = tk.StringVar(value='3x/day')
        tk.Spinbox(r2, textvariable=freq_var, font=('Arial', 10),
                 width=15, bg='white', values=freq_options).pack(side='left', padx=5)

        # ردیف ۳: مدت + تعداد کل (استفاده از Spinbox)
        r3 = tk.Frame(add_frame, bg='#ecf0f1')
        r3.pack(fill='x', pady=3)

        tk.Label(r3, text='Duration:', font=('Arial', 10, 'bold'),
                 bg='#ecf0f1', width=14, anchor='w').pack(side='left')
        duration_var = tk.StringVar(value='7 days')
        tk.Spinbox(r3, textvariable=duration_var, font=('Arial', 10),
                 width=15, bg='white', values=duration_options).pack(side='left', padx=5)

        tk.Label(r3, text='Quantity:', font=('Arial', 10, 'bold'),
                 bg='#ecf0f1', width=11, anchor='w').pack(side='left', padx=(15, 0))
        qty_var = tk.StringVar(value='1') # ✅ مقدار پیش‌فرض 1 بسته
        tk.Spinbox(r3, from_=1, to=1000, textvariable=qty_var, font=('Arial', 10),
                 width=12, bg='white').pack(side='left', padx=5)

        # ردیف ۴: توضیحات
        r4 = tk.Frame(add_frame, bg='#ecf0f1')
        r4.pack(fill='x', pady=3)

        tk.Label(r4, text='Instructions:', font=('Arial', 10, 'bold'),
                 bg='#ecf0f1', width=14, anchor='w').pack(side='left')
        instr_var = tk.StringVar()
        tk.Entry(r4, textvariable=instr_var, font=('Arial', 10),
                 width=55, bg='white').pack(side='left', padx=5)
        tk.Label(r4, text='(e.g., after meals)',
                 font=('Arial', 8, 'italic'),
                 bg='#ecf0f1', fg='#95a5a6').pack(side='left')

        # ===== ۷. بخش سبد نسخه =====
        basket_frame = tk.LabelFrame(scrollable, text=' 💊 Pending Prescription ',
                                     font=('Arial', 11, 'bold'),
                                     bg='#ecf0f1', fg='#c0392b',
                                     padx=10, pady=10)
        basket_frame.pack(fill='both', expand=True, padx=20, pady=10)

        basket_cols = ('#', 'Medicine', 'Dosage', 'Freq', 'Duration', 'Qty')

        tree_scroll_frame = tk.Frame(basket_frame, bg='#ecf0f1')
        tree_scroll_frame.pack(fill='both', expand=True)

        basket_tree = ttk.Treeview(tree_scroll_frame, columns=basket_cols,
                                   show='headings', height=8)

        basket_widths = {'#': 50, 'Medicine': 250, 'Dosage': 100,
                         'Freq': 100, 'Duration': 100, 'Qty': 80}
        for col in basket_cols:
            basket_tree.heading(col, text=col)
            basket_tree.column(col, width=basket_widths[col], anchor='center', minwidth=60)

        y_scroll = ttk.Scrollbar(tree_scroll_frame, orient='vertical', command=basket_tree.yview)
        basket_tree.configure(yscrollcommand=y_scroll.set)

        x_scroll = ttk.Scrollbar(tree_scroll_frame, orient='horizontal', command=basket_tree.xview)
        basket_tree.configure(xscrollcommand=x_scroll.set)

        basket_tree.grid(row=0, column=0, sticky='nsew')
        y_scroll.grid(row=0, column=1, sticky='ns')
        x_scroll.grid(row=1, column=0, sticky='ew')

        tree_scroll_frame.grid_rowconfigure(0, weight=1)
        tree_scroll_frame.grid_columnconfigure(0, weight=1)

        pending_rx = []

        def refresh_pending_display():
            for item in basket_tree.get_children():
                basket_tree.delete(item)

            for i, rx in enumerate(pending_rx, 1):
                basket_tree.insert('', 'end', values=(
                    i, rx['medicine_name'], rx['dosage'],
                    rx['frequency'], rx['duration'], rx['quantity']
                ))

        def add_to_pending():
            try:
                med_id = int(med_var.get().split(' - ')[0])
                med_name = med_var.get().split(' - ')[1].split(' - ')[0].strip()
                qty = int(qty_var.get())

                if qty <= 0:
                    messagebox.showerror('Error', '❌ Quantity must be positive!')
                    return

                pending_rx.append({
                    'medicine_id': med_id,
                    'medicine_name': med_name,
                    'dosage': dosage_var.get().strip(),
                    'frequency': freq_var.get().strip(),
                    'duration': duration_var.get().strip(),
                    'quantity': qty,
                    'instructions': instr_var.get().strip(),
                })

                refresh_pending_display()
                dosage_var.set('500mg')
                freq_var.set('3x/day')
                duration_var.set('7 days')
                qty_var.set('1') # ✅ بعد از اضافه شدن، مجددا 1 می‌شود
                instr_var.set('')

            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        def remove_selected():
            selected = basket_tree.selection()
            if not selected:
                messagebox.showwarning('Warning', '⚠️ Select a row first!')
                return
            idx = int(basket_tree.item(selected[0])['values'][0]) - 1
            if 0 <= idx < len(pending_rx):
                pending_rx.pop(idx)
                refresh_pending_display()

        # نمایش اولیه‌ی نسخه‌های موجود
        for rx in existing_rx:
            pending_rx.append({
                'medicine_id': rx[7],
                'medicine_name': rx[1],
                'dosage': rx[2] or '',
                'frequency': rx[3] or '',
                'duration': rx[4] or '',
                'quantity': rx[5],
                'instructions': rx[6] or '',
                'existing_id': rx[0],
            })
        refresh_pending_display()

        # ===== ۸. تابع ذخیره =====
        def save_all():
            if not pending_rx:
                messagebox.showwarning('Warning', '⚠️ No prescriptions to save!')
                return

            try:
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute('DELETE FROM prescriptions WHERE record_id = ?', (record_id,))

                for rx in pending_rx:
                    if not rx.get('medicine_id'):
                        continue
                    cursor.execute('''
                        INSERT INTO prescriptions
                        (record_id, medicine_id, dosage, frequency,
                         duration, quantity, instructions)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (record_id, rx['medicine_id'], rx['dosage'],
                          rx['frequency'], rx['duration'], rx['quantity'],
                          rx['instructions']))

                conn.commit()
                conn.close()

                messagebox.showinfo('✅ Success',
                                    f'Prescription saved!\n\n'
                                    f'💊 {len(pending_rx)} medicine(s) prescribed.')

                try:
                    log_action('UPDATE', entity_type='prescription',
                               entity_id=record_id,
                               details=f'Prescription for record #{record_id}: {len(pending_rx)} medicine(s)',
                               user=getattr(self.root, 'current_user', None))
                except:
                    pass

                win.destroy()
                self.load_medical_records()

            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        # ===== ۹. دکمه‌ها =====
        create_popup_buttons(btn_container, [
            {'text': '➕ Add to Prescription', 'command': add_to_pending,
             'color': 'success', 'width': 200},
            {'text': '🗑️ Remove Selected', 'command': remove_selected,
             'color': 'danger', 'width': 180},
            {'text': '💾 Save Prescription', 'command': save_all,
             'color': 'primary', 'width': 200},
            {'text': '❌ Cancel', 'command': win.destroy,
             'color': 'warning', 'width': 120},
        ], bg='#ecf0f1')

        win.update_idletasks()
        win.state('zoomed')

        med_combo.focus()



    def print_medical_record_by_id(self, record_id):
        """پرینت پرونده‌ی پزشکی"""
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT
                mr.id, mr.record_date, mr.chief_complaint, mr.diagnosis,
                mr.symptoms, mr.blood_pressure, mr.temperature,
                mr.weight, mr.height, mr.notes,
                p.first_name || ' ' || p.last_name,
                p.national_id, p.phone, p.age, p.gender, p.blood_type,
                d.first_name || ' ' || d.last_name, d.specialty
            FROM medical_records mr
            JOIN patients p ON mr.patient_id = p.id
            JOIN doctors d ON mr.doctor_id = d.id
            WHERE mr.id = ?
        ''', (record_id,))
        record = cursor.fetchone()

        cursor.execute('''
            SELECT m.name, pr.dosage, pr.frequency, pr.duration,
                   pr.quantity, pr.instructions
            FROM prescriptions pr
            JOIN medicines m ON pr.medicine_id = m.id
            WHERE pr.record_id = ?
            ORDER BY pr.id
        ''', (record_id,))
        prescriptions = cursor.fetchall()
        conn.close()

        if not record:
            messagebox.showerror('Error', '❌ Record not found!')
            return

        # ===== ساخت متن نسخه =====
        lines = []
        lines.append('=' * 70)
        lines.append(' ' * 15 + 'HOSPITAL MANAGEMENT SYSTEM')
        lines.append(' ' * 22 + 'MEDICAL PRESCRIPTION')
        lines.append('=' * 70)
        lines.append('')
        lines.append(f'  Record #: {record[0]}')
        lines.append(f'  Date: {record[1]}')
        lines.append('')
        lines.append('-' * 70)
        lines.append('  PATIENT INFORMATION')
        lines.append('-' * 70)
        lines.append(f'  Name: {record[10]}')
        lines.append(f'  National ID: {record[11]}')
        lines.append(f'  Age: {record[13]} years    Gender: {record[14]}')
        lines.append(f'  Phone: {record[12]}')
        lines.append(f'  Blood Type: {record[15]}')
        lines.append('')
        lines.append('-' * 70)
        lines.append('  DOCTOR INFORMATION')
        lines.append('-' * 70)
        doctor_clean = record[16].replace('Dr.', '').replace('dr.', '').strip()
        lines.append(f'  Doctor: Dr. {doctor_clean}')
        lines.append(f'  Specialty: {record[17]}')
        lines.append('')

        # Vital Signs
        if any([record[5], record[6], record[7], record[8]]):
            lines.append('-' * 70)
            lines.append('  VITAL SIGNS')
            lines.append('-' * 70)
            if record[5]:
                lines.append(f'  Blood Pressure: {record[5]}')
            if record[6]:
                lines.append(f'  Temperature: {record[6]} C')
            if record[7]:
                lines.append(f'  Weight: {record[7]} kg')
            if record[8]:
                lines.append(f'  Height: {record[8]} cm')
            lines.append('')

        # Chief Complaint
        if record[2]:
            lines.append('-' * 70)
            lines.append('  CHIEF COMPLAINT')
            lines.append('-' * 70)
            lines.append(f'  {record[2]}')
            lines.append('')

        # Symptoms
        if record[4]:
            lines.append('-' * 70)
            lines.append('  SYMPTOMS')
            lines.append('-' * 70)
            lines.append(f'  {record[4]}')
            lines.append('')

        # Diagnosis
        if record[3]:
            lines.append('-' * 70)
            lines.append('  DIAGNOSIS')
            lines.append('-' * 70)
            lines.append(f'  {record[3]}')
            lines.append('')

        # Prescriptions
        lines.append('-' * 70)
        lines.append('  PRESCRIPTION (Rx)')
        lines.append('-' * 70)
        lines.append('')

        if prescriptions:
            for i, rx in enumerate(prescriptions, 1):
                lines.append(f'  {i}. {rx[0]}')
                lines.append(f'     Dosage: {rx[1] or "—"}    Frequency: {rx[2] or "—"}')
                lines.append(f'     Duration: {rx[3] or "—"}    Packages: {rx[4]}')
                if rx[5]:
                    lines.append(f'     Instructions: {rx[5]}')
                lines.append('')
        else:
            lines.append('  (No medicines prescribed)')
            lines.append('')

        # Notes
        if record[9]:
            lines.append('-' * 70)
            lines.append('  NOTES')
            lines.append('-' * 70)
            lines.append(f'  {record[9]}')
            lines.append('')

        lines.append('=' * 70)
        lines.append(f'  Doctor Signature: _______________________')
        lines.append('')
        lines.append(f'  Printed: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        lines.append('=' * 70)

        content = '\n'.join(lines)

        # ===== ساخت PDF =====
        success, result = print_text_file(content, f'prescription_{record_id}.txt',
                                          open_pdf=False)

        if not success:
            messagebox.showerror('❌ PDF Error', f'Could not create PDF:\n{result}')
            return

        pdf_path = result

        # ===== پرینت به پیش‌فرض =====
        # ===== پرینت به پرینتر نسخه‌ها =====
        if WIN32_AVAILABLE:
            printer = get_printer_for('prescription')  # ← تغییر
            if printer:
                ok, msg = print_pdf_to_printer(pdf_path, printer)
                if ok:
                    messagebox.showinfo(
                        '🖨️ Sent to Printer',
                        f'Prescription sent to:\n\n🖨️ {printer}'
                    )
                else:
                    try:
                        os.startfile(pdf_path)
                    except:
                        pass
                    messagebox.showwarning(
                        '⚠️ Could Not Print',
                        f'PDF opened for manual printing.'
                    )
            else:
                try:
                    os.startfile(pdf_path)
                except:
                    pass
                messagebox.showwarning('⚠️ No Printer',
                                       'No default printer. PDF opened.')
        else:
            try:
                os.startfile(pdf_path)
            except:
                pass

    def batch_print_prescriptions(self):
        """چاپ گروهی نسخه‌ها"""
        selected = self.medical_tree.selection()

        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select at least one record!\n\n'
                                              '💡 Tip: Use Ctrl+Click or Shift+Click.')
            return

        if not messagebox.askyesno(
                '🖨️ Batch Print',
                f'Print {len(selected)} prescription(s)?'
        ):
            return

        printer = get_printer_for('prescription')
        if not printer and WIN32_AVAILABLE:
            printer = get_default_printer()

        # ===== Progress =====
        progress_win = tk.Toplevel(self.root)
        progress_win.title('🖨️ Printing...')
        progress_win.geometry('400x150')
        progress_win.configure(bg='white')
        progress_win.resizable(False, False)

        progress_win.update_idletasks()
        w, h = 400, 150
        x = (progress_win.winfo_screenwidth() // 2) - (w // 2)
        y = (progress_win.winfo_screenheight() // 2) - (h // 2)
        progress_win.geometry(f'{w}x{h}+{x}+{y}')

        tk.Label(progress_win, text='🖨️ Preparing prescriptions...',
                 font=('Arial', 13, 'bold'), bg='white',
                 fg='#2c3e50').pack(pady=15)

        progress_var = tk.StringVar(value=f'0 / {len(selected)}')
        tk.Label(progress_win, textvariable=progress_var,
                 font=('Arial', 11), bg='white').pack()

        progress_bar = ttk.Progressbar(progress_win, length=350,
                                       mode='determinate',
                                       maximum=len(selected))
        progress_bar.pack(pady=10, padx=25)
        progress_win.update()

        # ===== پردازش =====
        pdf_paths = []
        errors = []

        conn = get_connection()
        cursor = conn.cursor()

        for idx, item in enumerate(selected, 1):
            try:
                tags = self.medical_tree.item(item)['tags']
                if not tags:
                    continue

                record_id = tags[0]

                # خواندن اطلاعات
                cursor.execute('''
                    SELECT
                        mr.id, mr.record_date, mr.chief_complaint, mr.diagnosis,
                        mr.symptoms, mr.blood_pressure, mr.temperature,
                        mr.weight, mr.height, mr.notes,
                        p.first_name || ' ' || p.last_name,
                        p.national_id, p.phone, p.age, p.gender, p.blood_type,
                        d.first_name || ' ' || d.last_name, d.specialty
                    FROM medical_records mr
                    JOIN patients p ON mr.patient_id = p.id
                    JOIN doctors d ON mr.doctor_id = d.id
                    WHERE mr.id = ?
                ''', (record_id,))
                record = cursor.fetchone()

                if not record:
                    continue

                cursor.execute('''
                    SELECT m.name, pr.dosage, pr.frequency, pr.duration,
                           pr.quantity, pr.instructions
                    FROM prescriptions pr
                    JOIN medicines m ON pr.medicine_id = m.id
                    WHERE pr.record_id = ?
                ''', (record_id,))
                prescriptions = cursor.fetchall()

                content = self._build_prescription_text(record, prescriptions)

                success, pdf_path = print_text_file(
                    content,
                    f'batch_rx_{record_id}.txt',
                    open_pdf=False
                )

                if success:
                    pdf_paths.append(pdf_path)

                progress_var.set(f'{idx} / {len(selected)}')
                progress_bar['value'] = idx
                progress_win.update()

            except Exception as e:
                errors.append(str(e))

        conn.close()
        progress_win.destroy()

        # ===== چاپ =====
        if not pdf_paths:
            messagebox.showerror('❌ Error', 'Could not prepare any PDFs!')
            return

        if printer:
            # ===== چک کن آیا مجازیه یا فیزیکی =====
            if is_virtual_printer(printer):
                # 🎯 پرینتر مجازیه → PDF ها رو ذخیره کن
                save_dir = os.path.join(os.path.expanduser('~'), 'Desktop', 'Hospital PDFs')
                os.makedirs(save_dir, exist_ok=True)

                saved_files = []
                for pdf in pdf_paths:
                    try:
                        dest = os.path.join(save_dir, os.path.basename(pdf))
                        shutil.copy2(pdf, dest)
                        saved_files.append(dest)
                    except:
                        pass

                # باز کردن پوشه
                try:
                    os.startfile(save_dir)
                except:
                    pass

                messagebox.showinfo(
                    '📄 PDFs Saved',
                    f'✅ {len(saved_files)} PDF(s) saved!\n\n'
                    f'📁 Location:\n{save_dir}\n\n'
                    f'💡 Folder has been opened.\n'
                    f'You can now view, print, or share the PDFs.'
                )
            else:
                # 🎯 پرینتر فیزیکی → چاپ مستقیم
                success_count, error_count, print_errors = print_multiple_pdfs(
                    pdf_paths, printer
                )

                if error_count == 0:
                    messagebox.showinfo('✅ Success',
                                        f'🖨️ Batch Print Complete!\n\n'
                                        f'✅ {success_count} printed successfully!')
                else:
                    msg = f'⚠️ Partial Success\n\n'
                    msg += f'✅ Success: {success_count}\n'
                    msg += f'❌ Failed: {error_count}\n\n'
                    msg += 'Errors:\n' + '\n'.join(print_errors[:3])
                    messagebox.showwarning('⚠️ Partial Success', msg)
        else:
            # پرینتر نیست → PDF ذخیره کن
            save_dir = os.path.join(os.path.expanduser('~'), 'Desktop', 'Hospital PDFs')
            os.makedirs(save_dir, exist_ok=True)

            for pdf in pdf_paths:
                try:
                    dest = os.path.join(save_dir, os.path.basename(pdf))
                    shutil.copy2(pdf, dest)
                except:
                    pass

            try:
                os.startfile(save_dir)
            except:
                pass

            messagebox.showinfo(
                '📄 PDFs Saved',
                f'{len(pdf_paths)} PDF(s) saved to:\n\n{save_dir}'
            )

    def _build_prescription_text(self, record, prescriptions):
        """ساخت متن نسخه (helper)"""
        (r_id, r_date, cc, diag, sym, bp, temp, w, h, notes,
         patient_name, nid, phone, age, gender, blood,
         doctor_name, specialty) = record

        lines = []
        lines.append('=' * 70)
        lines.append('         🏥 HOSPITAL MANAGEMENT SYSTEM')
        lines.append('              MEDICAL PRESCRIPTION')
        lines.append('=' * 70)
        lines.append('')
        lines.append(f'Record #: {r_id}    Date: {r_date}')
        lines.append('')
        lines.append('-' * 70)
        lines.append('  PATIENT INFORMATION')
        lines.append('-' * 70)
        lines.append(f'  Name: {patient_name}')
        lines.append(f'  National ID: {nid}')
        lines.append(f'  Age: {age} years    Gender: {gender}')
        lines.append(f'  Phone: {phone}    Blood: {blood}')
        lines.append('')
        lines.append('-' * 70)
        lines.append('  DOCTOR')
        lines.append('-' * 70)
        doctor_clean = doctor_name.replace('Dr.', '').replace('dr.', '').strip()
        lines.append(f'  Dr. {doctor_clean} ({specialty})')
        lines.append('')

        # Vital Signs
        if any([bp, temp, w, h]):
            lines.append('-' * 70)
            lines.append('  VITAL SIGNS')
            lines.append('-' * 70)
            if bp: lines.append(f'  BP: {bp}')
            if temp: lines.append(f'  Temp: {temp} C')
            if w: lines.append(f'  Weight: {w} kg')
            if h: lines.append(f'  Height: {h} cm')
            lines.append('')

        if cc:
            lines.append(f'CHIEF COMPLAINT: {cc}')
            lines.append('')

        if diag:
            lines.append(f'DIAGNOSIS: {diag}')
            lines.append('')

        if sym:
            lines.append(f'SYMPTOMS: {sym}')
            lines.append('')

        # Prescriptions
        lines.append('-' * 70)
        lines.append('  PRESCRIPTION (Rx)')
        lines.append('-' * 70)
        lines.append('')

        if prescriptions:
            for i, rx in enumerate(prescriptions, 1):
                lines.append(f'  {i}. {rx[0]}')
                lines.append(f'     Dosage: {rx[1]} | Freq: {rx[2]}')
                lines.append(f'     Duration: {rx[3]} | Packages: {rx[4]}')
                if rx[5]:
                    lines.append(f'     Instructions: {rx[5]}')
                lines.append('')
        else:
            lines.append('  (No medicines prescribed)')
            lines.append('')

        if notes:
            lines.append(f'NOTES: {notes}')
            lines.append('')

        lines.append('=' * 70)
        lines.append('  Doctor Signature: _______________________')
        lines.append('=' * 70)

        return '\n'.join(lines)


    def print_medical_record(self):
        """پرینت از توی تب Medical Records"""
        selected = self.medical_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select a record first!')
            return

        record_id = self.medical_tree.item(selected[0])['tags'][0]
        self.print_medical_record_by_id(record_id)

    def delete_medical_record(self):
        selected = self.medical_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select at least one record!')
            return

        count = len(selected)

        if count == 1:
            confirm_msg = 'Delete this medical record?'
        else:
            confirm_msg = f'Delete {count} medical record(s)?\n\nThis action cannot be undone!'

        if not messagebox.askyesno('Confirm', confirm_msg):
            return

        try:
            conn = get_connection()
            cursor = conn.cursor()

            for item in selected:
                real_id = self.medical_tree.item(item)['tags'][0]
                cursor.execute('DELETE FROM prescriptions WHERE record_id = ?', (real_id,))
                cursor.execute('DELETE FROM medical_records WHERE id = ?', (real_id,))

            conn.commit()
            conn.close()

            messagebox.showinfo('Success', f'✅ {count} record(s) deleted!')
            self.load_medical_records()

        except Exception as e:
            messagebox.showerror('Error', f'❌ Error: {e}')


    # ==========================================
    # Appointments Tab
    # ==========================================

    def create_appointments_tab(self):
        tk.Label(self.appointments_tab, text='📅 Appointment Management',
                 font=('Arial', 14, 'bold'), bg=WINDOW_BG, fg='#2c3e50').pack(pady=15)

        # ===== دکمه‌ها =====
        btn_bar = create_button_bar(self.appointments_tab, bg=WINDOW_BG)

        create_standard_button(btn_bar, '➕ New', self.add_appointment_window, color='success')
        create_standard_button(btn_bar, '✏️ Update', self.update_appointment_status, color='warning')
        create_standard_button(btn_bar, '👁️ Refresh', self.load_appointments, color='primary')
        create_standard_button(btn_bar, '🔔 Reminders', self.show_appointment_reminders, color='info')
        create_standard_button(btn_bar, '🗑️ Delete', self.delete_appointment, color='danger')
        create_standard_button(btn_bar, '🧹 Clean Orphans', self.clean_orphan_appointments, color='secondary')

        # ===== ردیف ۱: Search =====
        row1 = tk.Frame(self.appointments_tab, bg=WINDOW_BG)
        row1.pack(pady=(5, 3), padx=40, fill='x')

        tk.Label(row1, text='🔍 Search:', font=('Arial', 11, 'bold'),
                 bg=WINDOW_BG, fg='#2c3e50').pack(side='left', padx=(0, 8))

        self.appointment_search_var = tk.StringVar()
        search_entry = tk.Entry(row1, textvariable=self.appointment_search_var,
                                font=('Arial', 11), bg='white',
                                relief='solid', bd=1, width=30)
        search_entry.pack(side='left', ipady=6, padx=(0, 8))

        self.appointment_search_var.trace_add('write',
                                              lambda *args: self.load_appointments())

        def clear_search():
            self.appointment_search_var.set('')

        tk.Button(row1, text='✖ Clear', command=clear_search,
                  bg='#95a5a6', fg='white', font=('Arial', 9, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=10, pady=4).pack(side='left')

        # ===== ردیف ۲: Filters =====
        row2 = tk.Frame(self.appointments_tab, bg=WINDOW_BG)
        row2.pack(pady=(0, 10), padx=40, fill='x')

        filter_var, from_var, to_var, df_container, custom_frame = create_date_filter(
            row2, WINDOW_BG, lambda: self.load_appointments()
        )
        self.appointment_date_filter = filter_var
        self.appointment_date_from = from_var
        self.appointment_date_to = to_var
        df_container.pack(side='left')

        # Status
        tk.Label(row2, text='  Status:', font=('Arial', 10, 'bold'),
                 bg=WINDOW_BG, fg='#2c3e50').pack(side='left', padx=(15, 5))

        self.appointment_status_filter = tk.StringVar(value='All')
        status_combo = ttk.Combobox(row2, textvariable=self.appointment_status_filter,
                                    values=['All', 'Scheduled', 'Completed', 'Cancelled'],
                                    state='readonly', width=12, font=('Arial', 10))
        status_combo.pack(side='left')
        status_combo.bind('<<ComboboxSelected>>', lambda e: self.load_appointments())

        self.appointment_count_var = tk.StringVar(value='')
        tk.Label(row2, textvariable=self.appointment_count_var,
                 font=('Arial', 10, 'italic'), bg=WINDOW_BG,
                 fg='#7f8c8d').pack(side='right')

        # ===== Treeview =====
        columns = ('ID', 'Date', 'Time', 'Patient', 'Doctor', 'Status')

        tree_frame = tk.Frame(self.appointments_tab, bg=WINDOW_BG)
        tree_frame.pack(fill='both', expand=True, padx=40, pady=(5, 10))

        self.appointments_tree = ttk.Treeview(tree_frame, columns=columns,
                                              show='headings', height=14,
                                              selectmode='extended')

        for col in columns:
            self.appointments_tree.heading(col, text=col)
            self.appointments_tree.column(col, width=150, anchor='center')

        scroll = ttk.Scrollbar(tree_frame, orient='vertical',
                               command=self.appointments_tree.yview)
        self.appointments_tree.configure(yscrollcommand=scroll.set)

        self.appointments_tree.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        enable_multi_click_selection(self.appointments_tree)

        self.load_appointments()

    def create_lab_tab(self):
        """تب مدیریت آزمایشگاه (تفکیک آزمایش آزاد و نسخه‌ای)"""
        tk.Label(self.lab_tab, text='🧪 Laboratory Management',
                 font=('Arial', 14, 'bold'), bg=WINDOW_BG, fg='#2c3e50').pack(pady=10)

        # ===== بخش آمار و فیلتر =====
        stats_outer = tk.Frame(self.lab_tab, bg=WINDOW_BG)
        stats_outer.pack(fill='x', padx=40, pady=(0, 10))

        filter_row = tk.Frame(stats_outer, bg=WINDOW_BG)
        filter_row.pack(fill='x', pady=(0, 5))

        tk.Label(filter_row, text='📅 Stats Date:', font=('Arial', 10, 'bold'),
                 bg=WINDOW_BG, fg='#2c3e50').pack(side='left', padx=(0, 5))

        self.lab_stat_filter_var, self.lab_stat_from_var, self.lab_stat_to_var, df_container, _ = create_date_filter(
            filter_row, WINDOW_BG, lambda: self.load_lab_stats()
        )
        df_container.pack(side='left')

        create_standard_button(
            filter_row, text='📊 Export Excel', command=self.export_lab_stats_excel,
            color='info', width=150, height=32
        )

        # ردیف کارت‌های آمار
        cards_frame = tk.Frame(stats_outer, bg=WINDOW_BG)
        cards_frame.pack(pady=5)

        cards_inner = tk.Frame(cards_frame, bg=WINDOW_BG)
        cards_inner.pack()

        self.lab_stat_cards = {}
        CARD_COLOR = '#3b4a6b'

        card_configs = [
            ('total_revenue', '💰 Total Revenue'),
            ('direct_revenue', '🛒 Direct / آزاد'),
            ('rx_revenue', '📋 Doctor Rx / نسخه'),
            ('pending_tests', '⏳ Pending Pay')
        ]

        for key, title in card_configs:
            canvas, updater = create_rounded_stat_card(cards_inner, title, '0', bg_color=CARD_COLOR, width=180,
                                                       height=85)
            canvas.pack(side='left', padx=5)
            self.lab_stat_cards[key] = updater

        # ===== دکمه‌های اصلی =====
        btn_bar1 = create_button_bar(self.lab_tab, bg=WINDOW_BG)
        create_standard_button(btn_bar1, '➕ Direct Request', self.add_lab_request_window, color='success')
        create_standard_button(btn_bar1, '✏️ Update Result', self.update_lab_result_window, color='warning')
        create_standard_button(btn_bar1, '👁️ View', self.view_lab_request_window, color='primary')
        create_standard_button(btn_bar1, '💳 Pay Lab', self.pay_lab_window, color='success')
        create_standard_button(btn_bar1, '📊 Full Report', self.lab_full_report_window, color='info')

        btn_bar2 = create_button_bar(self.lab_tab, bg=WINDOW_BG)
        create_standard_button(btn_bar2, '🔄 Refresh', self.load_lab_requests, color='info')
        create_standard_button(btn_bar2, '⚙️ Test Types', self.manage_lab_tests_window, color='secondary')
        create_standard_button(btn_bar2, '🖨️ Print Selected', self.batch_print_lab_requests, color='info')
        create_standard_button(btn_bar2, '🗑️ Delete', self.delete_lab_request, color='danger')

        # ===== ردیف ۱: Search =====
        row1 = tk.Frame(self.lab_tab, bg=WINDOW_BG)
        row1.pack(pady=(5, 3), padx=40, fill='x')

        tk.Label(row1, text='🔍 Search:', font=('Arial', 11, 'bold'),
                 bg=WINDOW_BG, fg='#2c3e50').pack(side='left', padx=(0, 8))

        self.lab_search_var = tk.StringVar()
        search_entry = tk.Entry(row1, textvariable=self.lab_search_var,
                                font=('Arial', 11), bg='white',
                                relief='solid', bd=1, width=30)
        search_entry.pack(side='left', ipady=6, padx=(0, 8))
        self.lab_search_var.trace_add('write', lambda *args: self.load_lab_requests())

        def clear_search():
            self.lab_search_var.set('')

        tk.Button(row1, text='✖ Clear', command=clear_search,
                  bg='#95a5a6', fg='white', font=('Arial', 9, 'bold'),
                  relief='flat', cursor='hand2', padx=10, pady=4).pack(side='left')

        # ===== ردیف ۲: Filters =====
        row2 = tk.Frame(self.lab_tab, bg=WINDOW_BG)
        row2.pack(pady=(0, 10), padx=40, fill='x')

        filter_var, from_var, to_var, df_container2, _ = create_date_filter(
            row2, WINDOW_BG, lambda: self.load_lab_requests()
        )
        self.lab_date_filter = filter_var
        self.lab_date_from = from_var
        self.lab_date_to = to_var
        df_container2.pack(side='left')

        tk.Label(row2, text='  Status:', font=('Arial', 10, 'bold'),
                 bg=WINDOW_BG, fg='#2c3e50').pack(side='left', padx=(15, 5))

        self.lab_status_filter = tk.StringVar(value='All')
        status_combo = ttk.Combobox(row2, textvariable=self.lab_status_filter,
                                    values=['All', 'Pending', 'In Progress', 'Done', 'Cancelled'],
                                    state='readonly', width=14, font=('Arial', 10))
        status_combo.pack(side='left')
        status_combo.bind('<<ComboboxSelected>>', lambda e: self.load_lab_requests())

        self.lab_count_var = tk.StringVar(value='')
        tk.Label(row2, textvariable=self.lab_count_var,
                 font=('Arial', 10, 'italic'), bg=WINDOW_BG,
                 fg='#7f8c8d').pack(side='right')

        # ===== Treeview (اضافه شدن ستون Type) =====
        columns = ('ID', 'Type', 'Date', 'Patient', 'Test', 'Doctor', 'Status', 'Payment', 'Result')
        tree_frame = tk.Frame(self.lab_tab, bg=WINDOW_BG)
        tree_frame.pack(fill='both', expand=True, padx=40, pady=(5, 10))

        self.lab_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=12, selectmode='extended')

        col_widths = {'ID': 50, 'Type': 90, 'Date': 90, 'Patient': 150, 'Test': 180, 'Doctor': 150, 'Status': 100,
                      'Payment': 90, 'Result': 80}
        for col in columns:
            self.lab_tree.heading(col, text=col)
            self.lab_tree.column(col, width=col_widths[col], anchor='center')

        scroll = ttk.Scrollbar(tree_frame, orient='vertical', command=self.lab_tree.yview)
        self.lab_tree.configure(yscrollcommand=scroll.set)
        self.lab_tree.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        self.lab_tree.bind('<Double-Button-1>', lambda e: self.view_lab_request_window())
        enable_multi_click_selection(self.lab_tree)

        self.load_lab_requests()
        self.load_lab_stats()

    def add_lab_request_window(self):
        """فرم ثبت درخواست آزمایش جدید (آزاد - بدون پزشک)"""
        win = tk.Toplevel(self.root)
        win.title('➕ New Direct Lab Request')
        win.configure(bg='#f0f0f0')
        win.resizable(True, True)

        btn_container = tk.Frame(win, bg='#f0f0f0')
        btn_container.pack(side='bottom', fill='x', pady=15)

        body_container = tk.Frame(win, bg='#f0f0f0')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f0f0f0')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        tk.Label(scrollable, text='➕ New Direct Lab Request',
                 font=('Arial', 16, 'bold'), bg='#f0f0f0',
                 fg='#2c3e50').pack(pady=15)

        form_frame = tk.Frame(scrollable, bg='#f0f0f0')
        form_frame.pack(pady=10, padx=40, fill='x')

        # Patient
        tk.Label(form_frame, text='👤 Patient *', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=0, column=0, sticky='w', pady=(5, 0), padx=5)

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT id, first_name, last_name, national_id FROM patients ORDER BY last_name')
        patients = cursor.fetchall()
        cursor.execute('SELECT id, name, price FROM lab_tests ORDER BY category, name')
        lab_tests = cursor.fetchall()
        conn.close()

        if not patients:
            messagebox.showerror('Error', '❌ No patients in database!')
            win.destroy()
            return

        if not lab_tests:
            messagebox.showerror('Error', '❌ No lab tests defined!')
            win.destroy()
            return

        patient_options = [f'{p[0]} - {p[1]} {p[2]} ({p[3]})' for p in patients]
        patient_var = tk.StringVar(value=patient_options[0])
        ttk.Combobox(form_frame, textvariable=patient_var, values=patient_options,
                     width=38, state='readonly').grid(
            row=1, column=0, pady=(0, 10), padx=5, sticky='ew')

        # Lab Test
        tk.Label(form_frame, text='🧪 Lab Test *', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=2, column=0, sticky='w', pady=(5, 0), padx=5)

        test_options = [f'{t[0]} - {t[1]} ({t[2]:,.0f} IRR)' for t in lab_tests]
        test_var = tk.StringVar(value=test_options[0])
        ttk.Combobox(form_frame, textvariable=test_var, values=test_options,
                     width=38, state='readonly').grid(
            row=3, column=0, pady=(0, 10), padx=5, sticky='ew')

        # Request Date
        tk.Label(form_frame, text='📅 Request Date *', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=4, column=0, sticky='w', pady=(5, 0), padx=5)

        from datetime import datetime
        date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        tk.Entry(form_frame, textvariable=date_var, font=('Arial', 10),
                 bg='white', relief='solid', bd=1).grid(
            row=5, column=0, pady=(0, 5), padx=5, sticky='ew')

        tk.Label(form_frame, text='💡 Format: YYYY-MM-DD', bg='#f0f0f0',
                 font=('Arial', 8), fg='#7f8c8d', anchor='w').grid(
            row=5, column=1, sticky='w', padx=5)

        # Notes
        tk.Label(form_frame, text='📝 Notes (Optional)', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=6, column=0, sticky='w', pady=(5, 0), padx=5)

        notes_text = tk.Text(form_frame, height=3, font=('Arial', 10),
                             bg='white', relief='solid', bd=1)
        notes_text.grid(row=7, column=0, pady=(0, 10), padx=5, sticky='ew')

        form_frame.grid_columnconfigure(0, weight=1)

        def save():
            try:
                patient_id = int(patient_var.get().split(' - ')[0])
                test_id = int(test_var.get().split(' - ')[0])

                # پزشک برای آزمایش آزاد خالی است
                doctor_id = None

                request_date = date_var.get().strip()
                if not re.match(r'^\d{4}-\d{2}-\d{2}$', request_date):
                    messagebox.showerror('Error', '❌ Invalid date! Use YYYY-MM-DD')
                    return

                notes = notes_text.get('1.0', 'end-1c').strip()

                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO lab_requests
                    (patient_id, doctor_id, test_id, request_date, status, notes)
                    VALUES (?, ?, ?, ?, 'Pending', ?)
                ''', (patient_id, doctor_id, test_id, request_date, notes))
                conn.commit()
                request_id = cursor.lastrowid
                conn.close()

                messagebox.showinfo('✅ Success', f'Lab request #{request_id} created!\n\nStatus: Pending')
                self.load_lab_requests()
                self.load_lab_stats()
                win.destroy()

            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        create_popup_buttons(btn_container, [
            {'text': '💾 Save Request', 'command': save, 'color': 'success', 'width': 180},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#f0f0f0')

        win.update_idletasks()
        win.state('zoomed')

    def view_lab_request_window(self):
        """نمایش کامل درخواست آزمایش"""
        selected = self.lab_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select a lab request first!')
            return

        request_id = self.lab_tree.item(selected[0])['tags'][0]

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT
                lr.id, lr.request_date, lr.status, lr.result, lr.result_date,
                lr.notes, lr.created_at,
                p.first_name || ' ' || p.last_name AS patient_name,
                p.national_id, p.phone, p.age, p.gender,
                lt.name AS test_name, lt.category, lt.price,
                lt.normal_range, lt.unit, lt.description,
                d.first_name || ' ' || d.last_name AS doctor_name
            FROM lab_requests lr
            JOIN patients p ON lr.patient_id = p.id
            JOIN lab_tests lt ON lr.test_id = lt.id
            LEFT JOIN doctors d ON lr.doctor_id = d.id
            WHERE lr.id = ?
        ''', (request_id,))
        record = cursor.fetchone()
        conn.close()

        if not record:
            messagebox.showerror('Error', '❌ Lab request not found!')
            return

        win = tk.Toplevel(self.root)
        win.title(f'🧪 Lab Request #{request_id}')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#ecf0f1')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#ecf0f1', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#ecf0f1')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        status_colors = {
            'Pending': '#f39c12',
            'In Progress': '#3498db',
            'Done': '#27ae60',
            'Cancelled': '#e74c3c',
        }
        header_color = status_colors.get(record[2], '#2c3e50')

        tk.Label(scrollable, text=f'🧪 Lab Request #{request_id}  —  {record[2].upper()}',
                 font=('Arial', 16, 'bold'), bg='#ecf0f1',
                 fg=header_color).pack(pady=15)

        def section(title, color='#3498db'):
            sec = tk.LabelFrame(scrollable, text=f' {title} ',
                                font=('Arial', 11, 'bold'),
                                bg='#ecf0f1', fg=color,
                                padx=15, pady=10)
            sec.pack(fill='x', pady=8, padx=20)
            return sec

        def info_row(parent, label, value):
            row = tk.Frame(parent, bg='#ecf0f1')
            row.pack(fill='x', pady=2)
            tk.Label(row, text=label, font=('Arial', 10, 'bold'),
                     bg='#ecf0f1', width=18, anchor='w').pack(side='left')
            tk.Label(row, text=value or '—', font=('Arial', 10),
                     bg='#ecf0f1', anchor='w', justify='left',
                     wraplength=600).pack(side='left', fill='x', expand=True)

        # ===== Patient Info =====
        sec_p = section('👤 Patient Information', '#27ae60')
        info_row(sec_p, 'Name:', record[7])
        info_row(sec_p, 'National ID:', record[8])
        info_row(sec_p, 'Phone:', record[9])
        info_row(sec_p, 'Age:', f'{record[10]} years')
        info_row(sec_p, 'Gender:', record[11])

        if record[18]:
            sec_d = section('👨‍⚕️ Requesting Doctor', '#9b59b6')
            doctor_clean = record[18].replace('Dr.', '').replace('dr.', '').strip()
            info_row(sec_d, 'Name:', f'Dr. {doctor_clean}')

        sec_t = section('🧪 Test Information', '#e67e22')
        info_row(sec_t, 'Test Name:', record[12])
        info_row(sec_t, 'Category:', record[13])
        info_row(sec_t, 'Price:', f'{record[14]:,.0f} IRR')
        if record[15]:
            info_row(sec_t, 'Normal Range:', record[15])
        if record[16]:
            info_row(sec_t, 'Unit:', record[16])
        if record[17]:
            info_row(sec_t, 'Description:', record[17])

        sec_r = section('📅 Request Info', '#3498db')
        info_row(sec_r, 'Request Date:', record[1])
        info_row(sec_r, 'Status:', record[2])
        info_row(sec_r, 'Created:', record[6])

        sec_res = section('📊 Result', '#c0392b')
        if record[3]:
            result_text = tk.Text(sec_res, height=6, font=('Courier New', 10),
                                  bg='white', relief='solid', bd=1, wrap='word')
            result_text.pack(fill='x', pady=5)
            result_text.insert('1.0', record[3])
            result_text.config(state='disabled')

            if record[4]:
                info_row(sec_res, 'Result Date:', record[4])
        else:
            tk.Label(sec_res, text='⏳ No result yet',
                     font=('Arial', 10, 'italic'),
                     bg='#ecf0f1', fg='#95a5a6').pack(pady=10)

        if record[5]:
            sec_n = section('📝 Notes', '#7f8c8d')
            tk.Label(sec_n, text=record[5], font=('Arial', 10),
                     bg='#ecf0f1', justify='left', anchor='w',
                     wraplength=800).pack(fill='x', anchor='w')

        def open_update():
            win.destroy()
            for item in self.lab_tree.get_children():
                if str(self.lab_tree.item(item)['tags'][0]) == str(request_id):
                    self.lab_tree.selection_set(item)
                    break
            self.update_lab_result_window()

        # ===== ۴. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '✏️ Update Result', 'command': open_update,
             'color': 'warning', 'width': 180},
            {'text': '❌ Close', 'command': win.destroy,
             'color': 'danger', 'width': 120},
        ], bg='#ecf0f1')

        # ===== ۵. Maximize =====
        win.update_idletasks()
        win.state('zoomed')

    def update_lab_result_window(self):
        """فرم بروزرسانی نتیجه‌ی آزمایش (استاندارد و واکنش‌گرا)"""
        selected = self.lab_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select a lab request first!')
            return

        request_id = self.lab_tree.item(selected[0])['tags'][0]

        # ===== خواندن اطلاعات فعلی =====
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT
                lr.id, lr.request_date, lr.status, lr.result, lr.result_date,
                lr.notes,
                p.first_name || ' ' || p.last_name AS patient_name,
                lt.name AS test_name, lt.normal_range, lt.unit
            FROM lab_requests lr
            JOIN patients p ON lr.patient_id = p.id
            JOIN lab_tests lt ON lr.test_id = lt.id
            WHERE lr.id = ?
        ''', (request_id,))
        record = cursor.fetchone()
        conn.close()

        if not record:
            messagebox.showerror('Error', '❌ Lab request not found!')
            return

        # ===== پنجره =====
        win = tk.Toplevel(self.root)
        win.title(f'✏️ Update Result - Lab Request #{request_id}')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        # ===== ۱. هدر =====
        header = tk.Frame(win, bg='#f39c12', height=55)
        header.pack(fill='x', side='top')
        header.pack_propagate(False)
        tk.Label(header, text=f'✏️ Update Lab Result #{request_id}',
                 font=('Arial', 15, 'bold'), bg='#f39c12',
                 fg='white').pack(pady=12)

        # ===== ۲. دکمه‌ها (پایین ثابت) =====
        btn_frame = tk.Frame(win, bg='#ecf0f1')
        btn_frame.pack(side='bottom', fill='x', pady=15)

        # ===== ۳. بدنه اصلی با اسکرول =====
        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=5, pady=5)

        canvas = tk.Canvas(body_container, bg='#ecf0f1', highlightthickness=0)
        v_scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        h_scrollbar = ttk.Scrollbar(body_container, orient='horizontal', command=canvas.xview)

        scrollable = tk.Frame(canvas, bg='#ecf0f1')
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')

        # چیدمان گرید
        canvas.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')

        body_container.grid_rowconfigure(0, weight=1)
        body_container.grid_columnconfigure(0, weight=1)

        # ===== تابع مدیریت هوشمند اسکرول‌بارها =====
        def check_scrollbars(event=None):
            canvas.configure(scrollregion=canvas.bbox('all'))
            bbox = canvas.bbox('all')
            if bbox:
                content_height = bbox[3] - bbox[1]
                canvas_height = canvas.winfo_height()
                if content_height > canvas_height + 5:
                    v_scrollbar.grid(row=0, column=1, sticky='ns')
                else:
                    v_scrollbar.grid_remove()

                content_width = bbox[2] - bbox[0]
                canvas_width = canvas.winfo_width()
                if content_width > canvas_width + 5:
                    h_scrollbar.grid(row=1, column=0, sticky='ew')
                else:
                    h_scrollbar.grid_remove()

        canvas.bind('<Configure>', check_scrollbars)
        scrollable.bind('<Configure>', check_scrollbars)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۴. محتوای اصلی =====
        content_frame = tk.Frame(scrollable, bg='#ecf0f1')
        content_frame.pack(fill='both', expand=True, padx=20, pady=10)

        # ===== اطلاعات بیمار و آزمایش =====
        info_frame = tk.Frame(content_frame, bg='#f8f9fa', relief='solid', bd=1)
        info_frame.pack(fill='x', pady=(0, 15))

        tk.Label(info_frame, text=f'👤 Patient: {record[6]}',
                 font=('Arial', 11, 'bold'), bg='#f8f9fa',
                 fg='#2c3e50', anchor='w').pack(anchor='w', padx=15, pady=(10, 3))

        tk.Label(info_frame, text=f'🧪 Test: {record[7]}',
                 font=('Arial', 11, 'bold'), bg='#f8f9fa',
                 fg='#2c3e50', anchor='w').pack(anchor='w', padx=15, pady=(0, 3))

        if record[8]:
            tk.Label(info_frame, text=f'📊 Normal Range: {record[8]}  ({record[9] or ""})',
                     font=('Arial', 10), bg='#f8f9fa',
                     fg='#7f8c8d', anchor='w').pack(anchor='w', padx=15, pady=(0, 10))
        else:
            tk.Label(info_frame, text='', bg='#f8f9fa').pack(pady=5)

        # ===== فرم اصلی =====
        form = tk.Frame(content_frame, bg='#ecf0f1')
        form.pack(fill='both', expand=True)

        # ===== Status =====
        tk.Label(form, text='📊 Status *:',
                 font=('Arial', 11, 'bold'), bg='#ecf0f1',
                 anchor='w').pack(anchor='w', pady=(0, 5))

        status_var = tk.StringVar(value=record[2])
        statuses = ['Pending', 'In Progress', 'Done', 'Cancelled']

        status_frame = tk.Frame(form, bg='#ecf0f1')
        status_frame.pack(fill='x', pady=(0, 15))

        for s in statuses:
            colors = {
                'Pending': '#f39c12',
                'In Progress': '#3498db',
                'Done': '#27ae60',
                'Cancelled': '#e74c3c',
            }
            tk.Radiobutton(status_frame, text=s, variable=status_var,
                           value=s, bg='#ecf0f1', fg=colors[s],
                           font=('Arial', 10, 'bold'),
                           activebackground='#ecf0f1').pack(side='left', padx=10)

        # ===== Result =====
        tk.Label(form, text='📋 Test Result:',
                 font=('Arial', 11, 'bold'), bg='#ecf0f1',
                 anchor='w').pack(anchor='w', pady=(0, 5))

        # در حالت فول اسکرین، این کادر فضای باقی‌مانده را پر می‌کند
        result_text_frame = tk.Frame(form, bg='#ecf0f1')
        result_text_frame.pack(fill='both', expand=True, pady=(0, 15))

        result_text = tk.Text(result_text_frame, height=15, font=('Courier New', 11),
                              bg='white', relief='solid', bd=1, wrap='word')
        result_text.pack(fill='both', expand=True)

        # پر کردن با نتیجه قبلی
        if record[3]:
            result_text.insert('1.0', record[3])

        # ===== Notes =====
        tk.Label(form, text='📝 Notes (Optional):',
                 font=('Arial', 11, 'bold'), bg='#ecf0f1',
                 anchor='w').pack(anchor='w', pady=(0, 5))

        notes_entry = tk.Entry(form, font=('Arial', 10), bg='white',
                               relief='solid', bd=1)
        notes_entry.pack(fill='x', pady=(0, 15), ipady=4)

        if record[5]:
            notes_entry.insert(0, record[5])

        # ===== Result Date =====
        tk.Label(form, text='📅 Result Date (auto for Done):',
                 font=('Arial', 10), bg='#ecf0f1',
                 fg='#7f8c8d', anchor='w').pack(anchor='w', pady=(0, 5))

        from datetime import datetime
        result_date_var = tk.StringVar(
            value=record[4] if record[4] else datetime.now().strftime('%Y-%m-%d')
        )
        tk.Entry(form, textvariable=result_date_var, font=('Arial', 10),
                 bg='white', relief='solid', bd=1).pack(fill='x', ipady=4)

        # ===== ۵. تابع ذخیره =====
        def save():
            try:
                new_status = status_var.get()
                new_result = result_text.get('1.0', 'end-1c').strip()
                new_notes = notes_entry.get().strip()
                new_result_date = result_date_var.get().strip()

                # اگه status = Done، حتماً result باید پر باشه
                if new_status == 'Done' and not new_result:
                    messagebox.showerror('Error',
                                         '❌ For "Done" status, Result is required!')
                    return

                # اعتبارسنجی تاریخ
                if new_result_date:
                    if not re.match(r'^\d{4}-\d{2}-\d{2}$', new_result_date):
                        messagebox.showerror('Error',
                                             '❌ Invalid date! Use YYYY-MM-DD')
                        return

                conn = get_connection()
                cursor = conn.cursor()

                # اگه status به Done تغییر کرد و تاریخ نتیجه خالیه، الان پر کن
                if new_status == 'Done' and not new_result_date:
                    new_result_date = datetime.now().strftime('%Y-%m-%d')

                cursor.execute('''
                    UPDATE lab_requests
                    SET status = ?, result = ?, result_date = ?, notes = ?
                    WHERE id = ?
                ''', (new_status, new_result, new_result_date or None,
                      new_notes, request_id))
                conn.commit()
                conn.close()

                messagebox.showinfo('✅ Success',
                                    f'Lab request #{request_id} updated!\n\n'
                                    f'Status: {new_status}')

                try:
                    log_action(
                        'UPDATE',
                        entity_type='lab_request',
                        entity_id=request_id,
                        details=f'Updated lab request #{request_id}: status={new_status}',
                        user=getattr(self.root, 'current_user', None)
                    )
                except:
                    pass

                win.destroy()
                self.load_lab_requests()

            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        # ===== ۶. دکمه‌ها =====
        create_popup_buttons(btn_frame, [
            {'text': '💾 Save Result', 'command': save,
             'color': 'success', 'width': 200},
            {'text': '❌ Cancel', 'command': win.destroy,
             'color': 'danger', 'width': 140},
        ], bg='#ecf0f1')

        # ===== ۷. Maximize =====
        win.update_idletasks()
        win.state('zoomed')



    def print_lab_request(self):
        """چاپ برگه‌ی درخواست آزمایش"""
        selected = self.lab_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select a lab request first!')
            return

        request_id = self.lab_tree.item(selected[0])['tags'][0]

        # ===== خواندن اطلاعات =====
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT
                lr.id, lr.request_date, lr.status, lr.result, lr.result_date,
                lr.notes,
                p.first_name || ' ' || p.last_name AS patient_name,
                p.national_id, p.phone, p.age, p.gender, p.blood_type,
                lt.name AS test_name, lt.category, lt.price,
                lt.normal_range, lt.unit,
                d.first_name || ' ' || d.last_name AS doctor_name
            FROM lab_requests lr
            JOIN patients p ON lr.patient_id = p.id
            JOIN lab_tests lt ON lr.test_id = lt.id
            LEFT JOIN doctors d ON lr.doctor_id = d.id
            WHERE lr.id = ?
        ''', (request_id,))
        record = cursor.fetchone()
        conn.close()

        if not record:
            messagebox.showerror('Error', '❌ Lab request not found!')
            return

        # ===== ساخت متن برگه =====
        content = self._build_lab_request_text(record)



        # ===== ساخت PDF =====
        success, pdf_path = print_text_file(
            content,
            f'lab_request_{request_id}.txt',
            open_pdf=False
        )

        if not success:
            messagebox.showerror('❌ PDF Error', f'Could not create PDF:\n{pdf_path}')
            return

        # ===== پرینت =====
        if WIN32_AVAILABLE:
            printer = get_printer_for('report')

            # ===== چک کن پرینتر مجازیه یا فیزیکی =====
            virtual_keywords = ['Print to PDF', 'XPS', 'OneNote', 'Fax', 'Virtual', 'Writer']
            is_virtual = any(kw.lower() in (printer or '').lower() for kw in virtual_keywords)

            if is_virtual or not printer:
                # ✅ پرینتر مجازی → PDF باز کن
                try:
                    os.startfile(pdf_path)
                    messagebox.showinfo('📄 PDF Ready',
                                        f'PDF created for printing.\n\n'
                                        f'📁 File: {pdf_path}\n\n'
                                        f'💡 The PDF has been opened.\n'
                                        f'Use Ctrl+P in the PDF viewer to save or print.')
                except Exception as e:
                    messagebox.showwarning('⚠️ PDF Error',
                                           f'PDF created at:\n{pdf_path}\n\nError: {e}')
            else:
                # ✅ پرینتر فیزیکی → چاپ کن
                ok, msg = print_pdf_to_printer(pdf_path, printer)
                if ok:
                    messagebox.showinfo('🖨️ Sent to Printer',
                                        f'Lab request sent to:\n\n🖨️ {printer}')
                else:
                    try:
                        os.startfile(pdf_path)
                    except:
                        pass
                    messagebox.showwarning('⚠️ Could Not Print',
                                           f'Error: {msg}\n\nPDF opened for manual printing.')
        else:
            try:
                os.startfile(pdf_path)
            except:
                pass
            messagebox.showinfo('📄 PDF Opened',
                                'PDF opened for manual printing.')

    def pay_lab_window(self):
        """فرم پرداخت آزمایش با کد ملی"""
        win = tk.Toplevel(self.root)
        win.title('💳 Pay Lab Test')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#ecf0f1')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#ecf0f1', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#ecf0f1')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text='💳 Pay Lab Test',
                 font=('Arial', 16, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(pady=15)

        # ===== ۴. جستجو با کد ملی =====
        search_frame = tk.Frame(scrollable, bg='#ecf0f1')
        search_frame.pack(fill='x', padx=40, pady=10)

        tk.Label(search_frame, text='🔍 National ID:',
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(side='left', padx=(0, 10))

        nid_var = tk.StringVar()
        nid_entry = tk.Entry(search_frame, textvariable=nid_var,
                             font=('Arial', 12), width=20,
                             bg='white', relief='solid', bd=1)
        nid_entry.pack(side='left', ipady=4, padx=(0, 10))

        # ===== ۵. نتایج =====
        result_frame = tk.Frame(scrollable, bg='#ecf0f1')
        result_frame.pack(fill='both', expand=True, padx=40, pady=10)

        cols = ('Select', 'Request ID', 'Date', 'Test', 'Price', 'Status')
        tree = ttk.Treeview(result_frame, columns=cols, show='headings', height=10)

        col_widths = {'Select': 60, 'Request ID': 80, 'Date': 110,
                      'Test': 220, 'Price': 120, 'Status': 100}
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=col_widths[col], anchor='center', minwidth=60)

        scroll = ttk.Scrollbar(result_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        # ===== ۶. جمع کل =====
        total_frame = tk.Frame(scrollable, bg='#f8f9fa', relief='solid', bd=2)
        total_frame.pack(fill='x', padx=40, pady=10)

        total_var = tk.StringVar(value='0 IRR')
        tk.Label(total_frame, text='💰 TOTAL:',
                 font=('Arial', 14, 'bold'), bg='#f8f9fa').pack(side='left', padx=15, pady=8)
        tk.Label(total_frame, textvariable=total_var,
                 font=('Arial', 14, 'bold'), bg='#f8f9fa', fg='#27ae60').pack(side='left')

        # ===== ۷. ذخیره‌ی داده‌ها =====
        pay_data = {'patient_id': None, 'patient_name': '', 'items': []}

        def search_lab_requests(event=None):
            nid = nid_var.get().strip()

            if not nid:
                messagebox.showwarning('Warning', '⚠️ Please enter National ID!')
                return

            if not nid.isdigit() or len(nid) != 10:
                messagebox.showerror('Error', '❌ National ID must be 10 digits!')
                return

            conn = get_connection()
            cursor = conn.cursor()

            cursor.execute('''
                SELECT id, first_name || ' ' || last_name
                FROM patients WHERE national_id = ?
            ''', (nid,))
            patient = cursor.fetchone()

            if not patient:
                conn.close()
                messagebox.showerror('Error', f'❌ No patient found with National ID: {nid}')
                return

            patient_id = patient[0]
            patient_name = patient[1]
            pay_data['patient_id'] = patient_id
            pay_data['patient_name'] = patient_name

            cursor.execute('''
                SELECT lr.id, lr.request_date, lt.name, lt.price, lr.status
                FROM lab_requests lr
                JOIN lab_tests lt ON lr.test_id = lt.id
                WHERE lr.patient_id = ?
                  AND (lr.payment_status IS NULL OR lr.payment_status = 'Pending')
                  AND lr.status != 'Cancelled'
                ORDER BY lr.request_date DESC
            ''', (patient_id,))

            requests = cursor.fetchall()
            conn.close()

            for item in tree.get_children():
                tree.delete(item)

            pay_data['items'] = []

            if not requests:
                messagebox.showinfo('Info',
                                    f'✅ No pending lab requests for {patient_name}')
                return

            total_amount = 0
            for r in requests:
                req_id, date, test_name, price, status = r
                pay_data['items'].append({
                    'request_id': req_id,
                    'test_name': test_name,
                    'price': price,
                    'total': price,
                })
                tree.insert('', 'end', values=(
                    '✅', req_id, date, test_name, f'{price:,.0f}', status
                ))
                total_amount += price

            total_var.set(f'{total_amount:,.0f} IRR')

        nid_entry.bind('<Return>', search_lab_requests)

        # ===== ۸. تابع پرداخت =====
        def do_payment():
            if not pay_data['items']:
                messagebox.showwarning('Warning', '⚠️ No lab requests to pay!')
                return
            payment_window(win)

        def payment_window(parent_win):
            pay_win = tk.Toplevel(parent_win)
            pay_win.title('💳 Lab Payment')
            pay_win.configure(bg='#ecf0f1')
            pay_win.resizable(True, True)

            pay_btn_container = tk.Frame(pay_win, bg='#ecf0f1')
            pay_btn_container.pack(side='bottom', fill='x', pady=15)

            pay_body = tk.Frame(pay_win, bg='#ecf0f1')
            pay_body.pack(fill='both', expand=True, padx=10, pady=5)

            pay_canvas = tk.Canvas(pay_body, bg='#ecf0f1', highlightthickness=0)
            pay_scrollbar = ttk.Scrollbar(pay_body, orient='vertical', command=pay_canvas.yview)
            pay_scrollable = tk.Frame(pay_canvas, bg='#ecf0f1')

            pay_scrollable.bind('<Configure>', lambda e: pay_canvas.configure(scrollregion=pay_canvas.bbox('all')))
            pay_canvas_window = pay_canvas.create_window((0, 0), window=pay_scrollable, anchor='nw')
            pay_canvas.configure(yscrollcommand=pay_scrollbar.set)
            pay_canvas.bind('<Configure>', lambda e: pay_canvas.itemconfig(pay_canvas_window, width=e.width))

            pay_canvas.pack(side='left', fill='both', expand=True)
            auto_scrollbar(pay_canvas, pay_scrollbar, pay_scrollable)

            tk.Label(pay_scrollable, text='💳 Lab Payment',
                     font=('Arial', 16, 'bold'), bg='#ecf0f1',
                     fg='#2c3e50').pack(pady=15)

            info = tk.Frame(pay_scrollable, bg='#f8f9fa', relief='solid', bd=1)
            info.pack(fill='x', padx=40, pady=10)

            tk.Label(info, text=f'👤 Patient: {pay_data["patient_name"]}',
                     font=('Arial', 11, 'bold'), bg='#f8f9fa',
                     anchor='w').pack(anchor='w', padx=15, pady=(10, 3))

            subtotal = sum(item['total'] for item in pay_data['items'])
            tk.Label(info, text=f'💰 Total: {subtotal:,.0f} IRR',
                     font=('Arial', 12, 'bold'), bg='#f8f9fa',
                     fg='#27ae60', anchor='w').pack(anchor='w', padx=15, pady=(3, 10))

            form = tk.Frame(pay_scrollable, bg='#ecf0f1')
            form.pack(fill='x', padx=40, pady=10)

            tk.Label(form, text='Discount (IRR):',
                     font=('Arial', 11, 'bold'), bg='#ecf0f1',
                     anchor='w').pack(anchor='w', pady=(0, 5))

            discount_var = tk.StringVar(value='0')
            tk.Entry(form, textvariable=discount_var,
                     font=('Arial', 11), bg='white').pack(fill='x', ipady=4, pady=(0, 15))

            tk.Label(form, text='Payment Method:',
                     font=('Arial', 11, 'bold'), bg='#ecf0f1',
                     anchor='w').pack(anchor='w', pady=(0, 5))

            method_var = tk.StringVar(value='Cash')
            for method in ['Cash', 'Card', 'Insurance']:
                tk.Radiobutton(form, text=method, variable=method_var,
                               value=method, bg='#ecf0f1',
                               font=('Arial', 10)).pack(anchor='w')

            final_frame = tk.Frame(pay_scrollable, bg='white', relief='solid', bd=2)
            final_frame.pack(fill='x', padx=40, pady=15)

            final_var = tk.StringVar()

            def update_final(*args):
                try:
                    discount = float(discount_var.get() or 0)
                except:
                    discount = 0
                final = max(subtotal - discount, 0)
                final_var.set(f'{final:,.0f}')

            discount_var.trace_add('write', update_final)
            update_final()

            tk.Label(final_frame, text='FINAL AMOUNT:',
                     font=('Arial', 13, 'bold'), bg='white').pack(side='left', padx=15, pady=10)
            tk.Label(final_frame, textvariable=final_var,
                     font=('Arial', 14, 'bold'), bg='white', fg='#27ae60').pack(side='left', padx=5)
            tk.Label(final_frame, text='IRR',
                     font=('Arial', 12), bg='white').pack(side='left')

            def confirm_payment():
                try:
                    discount = float(discount_var.get() or 0)
                except:
                    discount = 0
                final = max(subtotal - discount, 0)

                if final < 0:
                    messagebox.showerror('Error', '❌ Discount cannot exceed total!')
                    return

                if not messagebox.askyesno('Confirm Payment',
                                           f'Confirm payment?\n\n'
                                           f'Patient: {pay_data["patient_name"]}\n'
                                           f'Subtotal: {subtotal:,.0f} IRR\n'
                                           f'Discount: {discount:,.0f} IRR\n'
                                           f'Final: {final:,.0f} IRR\n'
                                           f'Method: {method_var.get()}'):
                    return

                try:
                    import random
                    from datetime import datetime
                    receipt = f'LAB-{datetime.now().strftime("%Y%m%d%H%M%S")}-{random.randint(100, 999)}'

                    conn = get_connection()
                    cursor = conn.cursor()

                    cursor.execute('''
                        INSERT INTO lab_payments
                        (receipt_number, patient_id, amount, discount, final_amount,
                         payment_method, payment_status, created_by)
                        VALUES (?, ?, ?, ?, ?, ?, 'Paid', ?)
                    ''', (receipt, pay_data['patient_id'], subtotal, discount,
                          final, method_var.get(),
                          getattr(self.root, 'current_user', {}).get('username', 'admin')))

                    for item in pay_data['items']:
                        cursor.execute('''
                            UPDATE lab_requests
                            SET payment_status = 'Paid',
                                final_amount = ?,
                                payment_method = ?,
                                receipt_number = ?
                            WHERE id = ?
                        ''', (item['price'], method_var.get(), receipt, item['request_id']))

                    conn.commit()
                    conn.close()

                    self.print_lab_receipt(
                        receipt=receipt,
                        patient_name=pay_data['patient_name'],
                        items=pay_data['items'],
                        subtotal=subtotal,
                        discount=discount,
                        final=final,
                        method=method_var.get()
                    )

                    messagebox.showinfo('✅ Success',
                                        f'✅ Lab payment completed!\n\n'
                                        f'Receipt: {receipt}\n'
                                        f'Total: {final:,.0f} IRR\n'
                                        f'Method: {method_var.get()}')

                    try:
                        log_action('PAYMENT', entity_type='lab_payment',
                                   details=f'Lab payment: {receipt}, Total: {final:,.0f} IRR',
                                   user=getattr(self.root, 'current_user', None))
                    except:
                        pass

                    self.load_lab_requests()
                    self.load_lab_stats()

                    pay_win.destroy()
                    win.destroy()


                except Exception as e:
                    messagebox.showerror('Error', f'❌ {e}')

            create_popup_buttons(pay_btn_container, [
                {'text': '💾 Confirm Payment', 'command': confirm_payment,
                 'color': 'success', 'width': 220},
                {'text': '❌ Cancel', 'command': pay_win.destroy,
                 'color': 'danger', 'width': 140},
            ], bg='#ecf0f1')

            pay_win.update_idletasks()
            pay_win.state('zoomed')

        # ===== ۹. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '🔍 Search', 'command': search_lab_requests,
             'color': 'primary', 'width': 140},
            {'text': '💳 Pay Now', 'command': do_payment,
             'color': 'success', 'width': 160},
            {'text': '❌ Close', 'command': win.destroy,
             'color': 'danger', 'width': 120},
        ], bg='#ecf0f1')

        # ===== ۱۰. Maximize =====
        win.update_idletasks()
        win.state('zoomed')

        nid_entry.focus()


    def print_lab_receipt(self, receipt, patient_name, items, subtotal, discount, final, method):
        """چاپ رسید پرداخت آزمایش"""
        from datetime import datetime

        lines = []
        lines.append('=' * 60)
        lines.append(' ' * 15 + 'HOSPITAL MANAGEMENT SYSTEM')
        lines.append(' ' * 18 + 'LAB PAYMENT RECEIPT')
        lines.append('=' * 60)
        lines.append('')
        lines.append(f'  Receipt #: {receipt}')
        lines.append(f'  Date: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        lines.append(f'  Patient: {patient_name}')
        lines.append('')
        lines.append('-' * 60)
        lines.append('  LAB TESTS')
        lines.append('-' * 60)

        for i, item in enumerate(items, 1):
            lines.append(f'  {i}. {item["test_name"]}')
            lines.append(f'     Price: {item["price"]:,.0f} IRR')

        lines.append('')
        lines.append('-' * 60)
        lines.append(f'  Subtotal:    {subtotal:>20,.0f} IRR')
        lines.append(f'  Discount:    {discount:>20,.0f} IRR')
        lines.append('=' * 60)
        lines.append(f'  TOTAL:       {final:>20,.0f} IRR')
        lines.append('=' * 60)
        lines.append('')
        lines.append(f'  Payment Method: {method}')
        lines.append('')
        lines.append('=' * 60)
        lines.append('  Thank you!')
        lines.append('=' * 60)

        content = '\n'.join(lines)

        success, pdf_path = print_text_file(content, f'lab_receipt_{receipt}.txt', open_pdf=False)

        if success:
            if WIN32_AVAILABLE:
                printer = get_printer_for('receipt')
                virtual_keywords = ['Print to PDF', 'XPS', 'OneNote', 'Fax', 'Virtual', 'Writer']
                is_virtual = any(kw.lower() in (printer or '').lower() for kw in virtual_keywords)

                if is_virtual or not printer:
                    try:
                        os.startfile(pdf_path)
                    except:
                        pass
                else:
                    print_pdf_to_printer(pdf_path, printer)



    def batch_print_lab_requests(self):
        """چاپ گروهی درخواست‌های آزمایش"""
        selected = self.lab_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select at least one lab request!')
            return

        count = len(selected)

        if not messagebox.askyesno(
                '🖨️ Batch Print',
                f'Print {count} lab request(s)?'
        ):
            return

        # ===== گرفتن پرینتر =====
        printer = get_printer_for('report')
        if not printer and WIN32_AVAILABLE:
            printer = get_default_printer()

        # ===== ساخت PDFها =====
        pdf_paths = []
        errors = []

        # ===== Progress Window =====
        progress_win = tk.Toplevel(self.root)
        progress_win.title('🖨️ Printing...')
        progress_win.geometry('400x150')
        progress_win.configure(bg='white')
        progress_win.resizable(False, False)

        progress_win.update_idletasks()
        w, h = 400, 150
        x = (progress_win.winfo_screenwidth() // 2) - (w // 2)
        y = (progress_win.winfo_screenheight() // 2) - (h // 2)
        progress_win.geometry(f'{w}x{h}+{x}+{y}')

        tk.Label(progress_win, text='🖨️ Preparing lab requests...',
                 font=('Arial', 13, 'bold'), bg='white',
                 fg='#2c3e50').pack(pady=15)

        progress_var = tk.StringVar(value=f'0 / {count}')
        tk.Label(progress_win, textvariable=progress_var,
                 font=('Arial', 11), bg='white').pack()

        progress_bar = ttk.Progressbar(progress_win, length=350,
                                       mode='determinate', maximum=count)
        progress_bar.pack(pady=10, padx=25)
        progress_win.update()

        # ===== پردازش =====
        conn = get_connection()
        cursor = conn.cursor()

        for idx, item in enumerate(selected, 1):
            try:
                tags = self.lab_tree.item(item)['tags']
                if not tags:
                    continue

                request_id = tags[0]

                cursor.execute('''
                    SELECT
                        lr.id, lr.request_date, lr.status, lr.result, lr.result_date,
                        lr.notes,
                        p.first_name || ' ' || p.last_name,
                        p.national_id, p.phone, p.age, p.gender, p.blood_type,
                        lt.name, lt.category, lt.price,
                        lt.normal_range, lt.unit,
                        d.first_name || ' ' || d.last_name
                    FROM lab_requests lr
                    JOIN patients p ON lr.patient_id = p.id
                    JOIN lab_tests lt ON lr.test_id = lt.id
                    LEFT JOIN doctors d ON lr.doctor_id = d.id
                    WHERE lr.id = ?
                ''', (request_id,))
                record = cursor.fetchone()

                if not record:
                    errors.append(f'Request #{request_id} not found')
                    continue

                # ساخت متن
                content = self._build_lab_request_text(record)

                # ساخت PDF
                success, pdf_path = print_text_file(
                    content,
                    f'lab_request_{request_id}.txt',
                    open_pdf=False
                )

                if success:
                    pdf_paths.append(pdf_path)

                progress_var.set(f'{idx} / {count}')
                progress_bar['value'] = idx
                progress_win.update()

            except Exception as e:
                errors.append(f'Error: {e}')

        conn.close()
        progress_win.destroy()

        if not pdf_paths:
            messagebox.showerror('❌ Error', 'Could not prepare any PDFs!')
            return

        # ===== چاپ یا ذخیره =====
        virtual_keywords = ['Print to PDF', 'XPS', 'OneNote', 'Fax', 'Virtual', 'Writer']
        is_virtual = any(kw.lower() in (printer or '').lower() for kw in virtual_keywords)

        if is_virtual or not printer:
            # ✅ پرینتر مجازی → ذخیره در Desktop
            import shutil
            save_dir = os.path.join(os.path.expanduser('~'), 'Desktop', 'Lab PDFs')
            os.makedirs(save_dir, exist_ok=True)

            saved = 0
            for pdf in pdf_paths:
                try:
                    dest = os.path.join(save_dir, os.path.basename(pdf))
                    shutil.copy2(pdf, dest)
                    saved += 1
                except:
                    pass

            # باز کردن پوشه
            try:
                os.startfile(save_dir)
            except:
                pass

            messagebox.showinfo('📄 PDFs Saved',
                                f'✅ {saved} PDF(s) saved!\n\n'
                                f'📁 Location:\n{save_dir}\n\n'
                                f'💡 Folder has been opened.')
        else:
            # ✅ پرینتر فیزیکی → چاپ مستقیم
            success_count = 0
            for pdf in pdf_paths:
                ok, msg = print_pdf_to_printer(pdf, printer)
                if ok:
                    success_count += 1
                import time
                time.sleep(0.5)

            if success_count == len(pdf_paths):
                messagebox.showinfo('✅ Success',
                                    f'🖨️ {success_count} lab request(s) printed!')
            else:
                messagebox.showwarning('⚠️ Partial',
                                       f'Printed: {success_count} / {len(pdf_paths)}')

        # ===== لاگ =====
        try:
            log_action(
                'PRINT',
                entity_type='lab_requests',
                details=f'Batch printed {len(pdf_paths)} lab request(s)',
                user=getattr(self.root, 'current_user', None)
            )
        except:
            pass

    def _build_lab_request_text(self, record):
        """ساخت متن برگه‌ی آزمایش (Helper)"""
        lines = []
        lines.append('=' * 70)
        lines.append(' ' * 15 + 'HOSPITAL MANAGEMENT SYSTEM')
        lines.append(' ' * 20 + 'LABORATORY REQUEST FORM')
        lines.append('=' * 70)
        lines.append('')
        lines.append(f'  Request #: {record[0]}')
        lines.append(f'  Date: {record[1]}')
        lines.append(f'  Status: {record[2].upper()}')
        lines.append('')
        lines.append('-' * 70)
        lines.append('  PATIENT INFORMATION')
        lines.append('-' * 70)
        lines.append(f'  Name: {record[6]}')
        lines.append(f'  National ID: {record[7]}')
        lines.append(f'  Age: {record[9]} years    Gender: {record[10]}')
        lines.append(f'  Phone: {record[8]}')
        lines.append(f'  Blood Type: {record[11]}')
        lines.append('')

        if record[17]:
            doctor_clean = record[17].replace('Dr.', '').replace('dr.', '').strip()
            lines.append('-' * 70)
            lines.append('  REQUESTING DOCTOR')
            lines.append('-' * 70)
            lines.append(f'  Dr. {doctor_clean}')
            lines.append('')

        lines.append('-' * 70)
        lines.append('  TEST INFORMATION')
        lines.append('-' * 70)
        lines.append(f'  Test: {record[12]}')
        lines.append(f'  Category: {record[13] or "—"}')
        lines.append(f'  Price: {record[14]:,.0f} IRR')
        if record[15]:
            lines.append(f'  Normal Range: {record[15]}')
        if record[16]:
            lines.append(f'  Unit: {record[16]}')
        lines.append('')

        if record[2] == 'Done' and record[3]:
            lines.append('-' * 70)
            lines.append('  TEST RESULT')
            lines.append('-' * 70)
            lines.append(f'  {record[3]}')
            if record[4]:
                lines.append(f'  Result Date: {record[4]}')
            lines.append('')

        if record[5]:
            lines.append('-' * 70)
            lines.append('  NOTES')
            lines.append('-' * 70)
            lines.append(f'  {record[5]}')
            lines.append('')

        lines.append('=' * 70)
        lines.append('  Technician Signature: _______________________')
        lines.append('')
        lines.append(f'  Printed: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        lines.append('=' * 70)

        return '\n'.join(lines)

    def manage_lab_tests_window(self):
        """مدیریت انواع آزمایش‌ها (CRUD)"""
        win = tk.Toplevel(self.root)
        win.title('⚙️ Manage Lab Tests')
        win.geometry('1000x700')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        win.update_idletasks()
        w, h = 1000, 700
        x = (win.winfo_screenwidth() // 2) - (w // 2)
        y = (win.winfo_screenheight() // 2) - (h // 2)
        win.geometry(f'{w}x{h}+{x}+{y}')

        # ===== هدر =====
        header = tk.Frame(win, bg='#2c3e50', height=55)
        header.pack(fill='x')
        header.pack_propagate(False)
        tk.Label(header, text='⚙️ Manage Lab Tests',
                 font=('Arial', 15, 'bold'), bg='#2c3e50',
                 fg='white').pack(pady=12)

        # ===== Treeview (اول تعریف کن) =====
        columns = ('ID', 'Name', 'Category', 'Price', 'Normal Range', 'Unit')
        tree_frame = tk.Frame(win, bg='#ecf0f1')
        tree_frame.pack(fill='both', expand=True, padx=20, pady=(10, 10))

        tree = ttk.Treeview(tree_frame, columns=columns,
                            show='headings', height=15,
                            selectmode='extended')

        col_widths = {
            'ID': 50, 'Name': 280, 'Category': 150,
            'Price': 130, 'Normal Range': 150, 'Unit': 100,
        }

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=col_widths[col], anchor='center')

        scroll = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        # ===== شمارنده =====
        count_var = tk.StringVar(value='')
        tk.Label(win, textvariable=count_var,
                 font=('Arial', 10, 'italic'), bg='#ecf0f1',
                 fg='#7f8c8d').pack(anchor='e', padx=30)

        # ===== تابع بارگذاری (قبل از دکمه‌ها!) =====
        def refresh_callback():
            # پاک کردن
            for item in tree.get_children():
                tree.delete(item)

            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, name, category, price, normal_range, unit
                FROM lab_tests ORDER BY category, name
            ''')
            tests = cursor.fetchall()
            conn.close()

            for t in tests:
                tree.insert('', 'end', values=(
                    t[0], t[1], t[2] or '—',
                    f'{t[3]:,.0f}',
                    t[4] or '—',
                    t[5] or '—'
                ), tags=(str(t[0]),))

            count_var.set(f'⚙️ {len(tests)} test(s) defined')

        # ===== تابع ادیت =====
        def edit_selected_test():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning('Warning', '⚠️ Please select a test first!')
                return

            test_id = tree.item(selected[0])['tags'][0]
            self.add_edit_lab_test_window(win, test_id, refresh_callback)

        # ===== تابع حذف =====
        def delete_selected_test():
            selected = tree.selection()
            if not selected:
                messagebox.showwarning('Warning', '⚠️ Please select a test first!')
                return

            if not messagebox.askyesno('Confirm',
                                       f'Delete {len(selected)} test(s)?\n\n'
                                       f'⚠️ This will also delete all requests using these tests!'):
                return

            try:
                conn = get_connection()
                cursor = conn.cursor()

                deleted = 0
                skipped = 0

                for item in selected:
                    test_id = tree.item(item)['tags'][0]

                    cursor.execute('SELECT COUNT(*) FROM lab_requests WHERE test_id = ?', (test_id,))
                    usage = cursor.fetchone()[0]

                    if usage > 0:
                        skipped += 1
                        continue

                    cursor.execute('DELETE FROM lab_tests WHERE id = ?', (test_id,))
                    deleted += 1

                conn.commit()
                conn.close()

                msg = f'✅ {deleted} test(s) deleted!'
                if skipped > 0:
                    msg += f'\n\n⚠️ {skipped} test(s) skipped (used in requests)'

                messagebox.showinfo('Result', msg)
                refresh_callback()

            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        # ===== تابع افزودن (جدید) =====
        def add_new_test():
            self.add_edit_lab_test_window(win, None, refresh_callback)

        # ===== حالا دکمه‌ها (با توابع تعریف‌شده) =====
        btn_bar = tk.Frame(win, bg='#ecf0f1')
        btn_bar.pack(fill='x', padx=20, pady=10)

        create_standard_button(btn_bar, '➕ Add Test', add_new_test, color='success')
        create_standard_button(btn_bar, '✏️ Edit', edit_selected_test, color='warning')
        create_standard_button(btn_bar, '🗑️ Delete', delete_selected_test, color='danger')
        create_standard_button(btn_bar, '🔄 Refresh', refresh_callback, color='primary')

        # ===== Double click =====
        tree.bind('<Double-Button-1>', lambda e: edit_selected_test())

        # ===== بارگذاری اولیه =====
        refresh_callback()

        # ===== دکمه‌ی بستن =====
        btn_close = tk.Frame(win, bg='#ecf0f1')
        btn_close.pack(fill='x', pady=10)

        create_popup_buttons(btn_close, [
            {'text': '❌ Close', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#ecf0f1')

    def add_edit_lab_test_window(self, parent, test_id, refresh_callback):
        """فرم افزودن/ویرایش آزمایش"""
        is_edit = test_id is not None

        # ===== خواندن اطلاعات فعلی (اگه ویرایش) =====
        current = None
        if is_edit:
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                SELECT name, category, price, normal_range, unit, description
                FROM lab_tests WHERE id = ?
            ''', (test_id,))
            current = cursor.fetchone()
            conn.close()

            if not current:
                messagebox.showerror('Error', '❌ Test not found!')
                return

        win = tk.Toplevel(parent)
        win.title(f'{"✏️ Edit" if is_edit else "➕ Add"} Lab Test')
        win.configure(bg='#f0f0f0')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#f0f0f0')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#f0f0f0')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f0f0f0')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text=f'{"✏️ Edit" if is_edit else "➕ Add"} Lab Test',
                 font=('Arial', 16, 'bold'), bg='#f0f0f0',
                 fg='#2c3e50').pack(pady=15)

        # ===== ۴. فرم =====
        form_frame = tk.Frame(scrollable, bg='#f0f0f0')
        form_frame.pack(pady=10, padx=40, fill='x')

        fields = {}
        labels = [
            ('Test Name *', 'name', 'e.g., CBC'),
            ('Category', 'category', 'e.g., Hematology'),
            ('Price (IRR) *', 'price', 'e.g., 150000'),
            ('Normal Range', 'normal_range', 'e.g., 4.5-11.0'),
            ('Unit', 'unit', 'e.g., 10^3/µL'),
            ('Description', 'description', 'Optional'),
        ]

        for i, (label, key, hint) in enumerate(labels):
            tk.Label(form_frame, text=label, font=('Arial', 10, 'bold'),
                     bg='#f0f0f0', anchor='w').grid(
                row=i * 2, column=0, sticky='w', pady=(8, 2), padx=5)

            entry = tk.Entry(form_frame, width=40, font=('Arial', 10))
            entry.grid(row=i * 2 + 1, column=0, pady=(0, 5), padx=5, sticky='ew')

            if current:
                if key == 'name':
                    entry.insert(0, current[0])
                elif key == 'category':
                    entry.insert(0, current[1] or '')
                elif key == 'price':
                    entry.insert(0, str(int(current[2])))
                elif key == 'normal_range':
                    entry.insert(0, current[3] or '')
                elif key == 'unit':
                    entry.insert(0, current[4] or '')
                elif key == 'description':
                    entry.insert(0, current[5] or '')

            fields[key] = entry

            tk.Label(form_frame, text=f'💡 {hint}', font=('Arial', 8, 'italic'),
                     bg='#f0f0f0', fg='#7f8c8d', anchor='w').grid(
                row=i * 2 + 1, column=1, sticky='w', padx=5)

        form_frame.grid_columnconfigure(0, weight=1)

        # ===== ۵. تابع ذخیره =====
        def save():
            try:
                name = fields['name'].get().strip()
                if not name:
                    messagebox.showerror('Error', '❌ Test name is required!')
                    return

                category = fields['category'].get().strip()
                normal_range = fields['normal_range'].get().strip()
                unit = fields['unit'].get().strip()
                description = fields['description'].get().strip()

                try:
                    price = float(fields['price'].get().strip() or 0)
                    if price < 0: raise ValueError
                except ValueError:
                    messagebox.showerror('Error', '❌ Price must be a positive number!')
                    return

                conn = get_connection()
                cursor = conn.cursor()

                if is_edit:
                    cursor.execute('''
                        UPDATE lab_tests
                        SET name = ?, category = ?, price = ?,
                            normal_range = ?, unit = ?, description = ?
                        WHERE id = ?
                    ''', (name, category, price, normal_range, unit, description, test_id))
                else:
                    cursor.execute('''
                        INSERT INTO lab_tests
                        (name, category, price, normal_range, unit, description)
                        VALUES (?, ?, ?, ?, ?, ?)
                    ''', (name, category, price, normal_range, unit, description))

                conn.commit()
                conn.close()

                messagebox.showinfo('✅ Success',
                                    f'Lab test "{name}" {"updated" if is_edit else "added"}!')

                try:
                    log_action('UPDATE' if is_edit else 'CREATE', entity_type='lab_test',
                               entity_id=test_id if is_edit else None,
                               details=f'{"Updated" if is_edit else "Added"} lab test: {name}',
                               user=getattr(self.root, 'current_user', None))
                except:
                    pass

                win.destroy()
                if refresh_callback:
                    refresh_callback()

            except sqlite3.IntegrityError:
                messagebox.showerror('Error', f'❌ Test name "{name}" already exists!')
            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        # ===== ۶. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '💾 Save', 'command': save, 'color': 'success', 'width': 160},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#f0f0f0')

        # ===== ۷. Maximize =====
        win.update_idletasks()
        win.state('zoomed')


    def delete_lab_request(self):
        """حذف درخواست(های) آزمایش"""
        selected = self.lab_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select at least one lab request!')
            return

        count = len(selected)

        if count == 1:
            values = self.lab_tree.item(selected[0])['values']
            confirm_msg = f'Delete lab request #{values[0]}?\n\n' \
                          f'Patient: {values[2]}\n' \
                          f'Test: {values[3]}\n\n' \
                          f'⚠️ This action cannot be undone!'
        else:
            confirm_msg = f'Delete {count} lab request(s)?\n\n' \
                          f'⚠️ This action cannot be undone!'

        if not messagebox.askyesno('Confirm Delete', confirm_msg):
            return

        try:
            conn = get_connection()
            cursor = conn.cursor()

            deleted = 0
            for item in selected:
                try:
                    request_id = self.lab_tree.item(item)['tags'][0]
                    cursor.execute('DELETE FROM lab_requests WHERE id = ?', (request_id,))
                    deleted += 1
                except Exception as e:
                    print(f'⚠️ Error deleting {item}: {e}')

            conn.commit()
            conn.close()

            messagebox.showinfo('✅ Success', f'{deleted} lab request(s) deleted!')

            # لاگ
            try:
                log_action(
                    'DELETE',
                    entity_type='lab_request',
                    details=f'Deleted {deleted} lab request(s)',
                    user=getattr(self.root, 'current_user', None)
                )
            except:
                pass

            self.load_lab_requests()

        except Exception as e:
            messagebox.showerror('Error', f'❌ {e}')

    def request_lab_from_record(self, record_id):
        """درخواست آزمایش از توی پرونده (با اتصال به record_id)"""
        # ===== خواندن اطلاعات پرونده =====
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT
                mr.id, mr.patient_id, mr.doctor_id,
                p.first_name || ' ' || p.last_name AS patient_name,
                d.first_name || ' ' || d.last_name AS doctor_name
            FROM medical_records mr
            JOIN patients p ON mr.patient_id = p.id
            LEFT JOIN doctors d ON mr.doctor_id = d.id
            WHERE mr.id = ?
        ''', (record_id,))
        record = cursor.fetchone()

        if not record:
            conn.close()
            messagebox.showerror('Error', '❌ Medical record not found!')
            return

        cursor.execute('SELECT id, name, price FROM lab_tests ORDER BY category, name')
        lab_tests = cursor.fetchall()
        conn.close()

        if not lab_tests:
            messagebox.showerror('Error', '❌ No lab tests defined!')
            return

        # ===== پنجره =====
        win = tk.Toplevel(self.root)
        win.title(f'🧪 Request Lab Test - Record #{record_id}')
        win.geometry('600x650')
        win.configure(bg='#ecf0f1')
        win.resizable(False, False)

        win.update_idletasks()
        w, h = 600, 650
        x = (win.winfo_screenwidth() // 2) - (w // 2)
        y = (win.winfo_screenheight() // 2) - (h // 2)
        win.geometry(f'{w}x{h}+{x}+{y}')

        # ===== هدر =====
        header = tk.Frame(win, bg='#e67e22', height=55)
        header.pack(fill='x')
        header.pack_propagate(False)
        tk.Label(header, text=f'🧪 Request Lab Test for Record #{record_id}',
                 font=('Arial', 14, 'bold'), bg='#e67e22',
                 fg='white').pack(pady=12)

        # ===== Info bar =====
        info_frame = tk.Frame(win, bg='#f8f9fa', relief='solid', bd=1)
        info_frame.pack(fill='x', padx=20, pady=10)

        tk.Label(info_frame, text=f'👤 Patient: {record[3]}',
                 font=('Arial', 11, 'bold'), bg='#f8f9fa',
                 fg='#2c3e50', anchor='w').pack(anchor='w', padx=15, pady=(10, 3))

        if record[4]:
            doctor_clean = record[4].replace('Dr.', '').replace('dr.', '').strip()
            tk.Label(info_frame, text=f'👨‍⚕️ Doctor: Dr. {doctor_clean}',
                     font=('Arial', 10), bg='#f8f9fa',
                     fg='#7f8c8d', anchor='w').pack(anchor='w', padx=15, pady=(0, 10))
        else:
            tk.Label(info_frame, text='', bg='#f8f9fa').pack(pady=5)

        # ===== فرم =====
        form = tk.Frame(win, bg='#ecf0f1')
        form.pack(fill='both', expand=True, padx=20, pady=10)

        tk.Label(form, text='🧪 Select Lab Test *:',
                 font=('Arial', 11, 'bold'), bg='#ecf0f1',
                 anchor='w').pack(anchor='w', pady=(0, 5))

        test_options = [f'{t[0]} - {t[1]} ({t[2]:,.0f} IRR)' for t in lab_tests]
        test_var = tk.StringVar(value=test_options[0])
        test_combo = ttk.Combobox(form, textvariable=test_var, values=test_options,
                                  state='readonly', width=60, font=('Arial', 10))
        test_combo.pack(fill='x', pady=(0, 15))

        # ===== تاریخ =====
        tk.Label(form, text='📅 Request Date *:',
                 font=('Arial', 11, 'bold'), bg='#ecf0f1',
                 anchor='w').pack(anchor='w', pady=(0, 5))

        date_var = tk.StringVar(value=datetime.now().strftime('%Y-%m-%d'))
        tk.Entry(form, textvariable=date_var, font=('Arial', 10),
                 bg='white', relief='solid', bd=1).pack(fill='x', ipady=3, pady=(0, 15))

        # ===== Notes =====
        tk.Label(form, text='📝 Notes (Optional):',
                 font=('Arial', 11, 'bold'), bg='#ecf0f1',
                 anchor='w').pack(anchor='w', pady=(0, 5))

        notes_text = tk.Text(form, height=3, font=('Arial', 10),
                             bg='white', relief='solid', bd=1)
        notes_text.pack(fill='x', pady=(0, 15))

        # ===== دکمه‌ها =====
        btn_frame = tk.Frame(win, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        def save():
            try:
                test_id = int(test_var.get().split(' - ')[0])
                request_date = date_var.get().strip()

                if not re.match(r'^\d{4}-\d{2}-\d{2}$', request_date):
                    messagebox.showerror('Error', '❌ Invalid date! Use YYYY-MM-DD')
                    return

                notes = notes_text.get('1.0', 'end-1c').strip()

                from datetime import datetime
                now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO lab_requests
                    (patient_id, doctor_id, test_id, record_id, request_date, status, notes, created_at)
                    VALUES (?, ?, ?, ?, ?, 'Pending', ?, ?)
                ''', (record[1], record[2], test_id, record_id, request_date, notes, now))
                conn.commit()
                request_id = cursor.lastrowid
                conn.close()

                messagebox.showinfo('✅ Success',
                                    f'Lab request #{request_id} created!\n\n'
                                    f'✅ Linked to medical record #{record_id}')

                try:
                    log_action(
                        'CREATE',
                        entity_type='lab_request',
                        entity_id=request_id,
                        details=f'Lab request from medical record #{record_id}',
                        user=getattr(self.root, 'current_user', None)
                    )
                except:
                    pass

                win.destroy()

                # اگه توی Medical Records بودیم، رفرش کنه
                try:
                    if hasattr(self, 'medical_tree'):
                        # ذخیره‌ی انتخاب فعلی
                        current_selection = None
                        for item in self.medical_tree.get_children():
                            if str(self.medical_tree.item(item)['tags'][0]) == str(record_id):
                                current_selection = item
                                break

                        if current_selection:
                            self.medical_tree.selection_set(current_selection)
                            self.view_medical_record_window()
                except:
                    pass

                # رفرش Laboratory (اگه بازه)
                try:
                    if hasattr(self, 'lab_tree'):
                        self.load_lab_requests()
                except:
                    pass

            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        create_popup_buttons(btn_frame, [
            {'text': '💾 Save Request', 'command': save, 'color': 'success', 'width': 200},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#ecf0f1')

    def load_lab_requests(self):
        search_text = self.lab_search_var.get().strip() if hasattr(self, 'lab_search_var') else ''
        status_filter = self.lab_status_filter.get() if hasattr(self, 'lab_status_filter') else 'All'

        from_date, to_date = None, None
        if hasattr(self, 'lab_date_filter'):
            from_date, to_date = get_date_range(
                self.lab_date_filter.get(),
                self.lab_date_from.get().strip() if hasattr(self, 'lab_date_from') else '',
                self.lab_date_to.get().strip() if hasattr(self, 'lab_date_to') else '',
            )

        query = """
            SELECT
                lr.id, lr.request_date, lr.record_id,
                p.first_name || ' ' || p.last_name AS patient_name,
                lt.name AS test_name,
                d.first_name || ' ' || d.last_name AS doctor_name,
                lr.status, lr.result,
                COALESCE(lr.payment_status, 'Pending') as payment_status,
                COALESCE(lr.final_amount, 0) as final_amount
            FROM lab_requests lr
            JOIN patients p ON lr.patient_id = p.id
            JOIN lab_tests lt ON lr.test_id = lt.id
            LEFT JOIN doctors d ON lr.doctor_id = d.id
            WHERE 1=1
        """
        params = []

        if search_text:
            pattern = f'%{search_text}%'
            query += " AND (p.first_name LIKE ? OR p.last_name LIKE ? OR lt.name LIKE ?)"
            params.extend([pattern, pattern, pattern])

        if status_filter and status_filter != 'All':
            query += " AND lr.status = ?"
            params.append(status_filter)

        if from_date:
            query += " AND lr.request_date >= ?"
            params.append(from_date)
        if to_date:
            query += " AND lr.request_date <= ?"
            params.append(to_date)

        query += " ORDER BY lr.request_date DESC, lr.id DESC"

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute(query, params)
        requests = cursor.fetchall()
        conn.close()

        for item in self.lab_tree.get_children():
            self.lab_tree.delete(item)

        status_icons = {'Pending': '⏳', 'In Progress': '🔄', 'Done': '✅', 'Cancelled': '❌'}

        for r in requests:
            req_id, req_date, record_id, patient, test, doctor, status, result, pay_status, final_amt = r

            # ✅ تشخیص نوع آزمایش
            req_type = "📋 Doctor Rx" if record_id else "🛒 Direct"

            icon = status_icons.get(status, '📌')
            pay_icon = '✅ Paid' if pay_status == 'Paid' else '⏳ Pending'
            doctor_clean = (doctor or '').replace('Dr.', '').strip()

            self.lab_tree.insert('', 'end', values=(
                req_id, req_type, req_date, patient, test,
                f'Dr. {doctor_clean}' if doctor_clean else '—',
                f'{icon} {status}', pay_icon, '✅' if result else '—'
            ), tags=(str(req_id),))

        if hasattr(self, 'lab_count_var'):
            self.lab_count_var.set(f'🧪 {len(requests)} request(s)')



    def load_appointments(self):
        # ===== خواندن فیلترها =====
        search_text = ''
        if hasattr(self, 'appointment_search_var'):
            search_text = self.appointment_search_var.get().strip()

        status_filter = 'All'
        if hasattr(self, 'appointment_status_filter'):
            status_filter = self.appointment_status_filter.get()

        # ===== Date Filter =====
        from_date, to_date = None, None
        if hasattr(self, 'appointment_date_filter'):
            from_date, to_date = get_date_range(
                self.appointment_date_filter.get(),
                self.appointment_date_from.get().strip() if hasattr(self, 'appointment_date_from') else '',
                self.appointment_date_to.get().strip() if hasattr(self, 'appointment_date_to') else '',
            )

        conn = get_connection()
        cursor = conn.cursor()

        # ===== ساخت کوئری داینامیک =====

        query = '''
            SELECT
                a.id, a.appointment_date, a.appointment_time, a.status,
                COALESCE(p.first_name || ' ' || p.last_name, '— (deleted)') AS patient_name,
                COALESCE(d.first_name || ' ' || d.last_name, '— (deleted)') AS doctor_name
            FROM appointments a
            LEFT JOIN patients p ON a.patient_id = p.id
            LEFT JOIN doctors d ON a.doctor_id = d.id
            WHERE 1=1
        '''

        params = []

        # جستجو
        if search_text:
            search_pattern = f'%{search_text}%'
            query += '''
                AND (
                    p.first_name LIKE ?
                    OR p.last_name LIKE ?
                    OR (p.first_name || ' ' || p.last_name) LIKE ?
                    OR d.first_name LIKE ?
                    OR d.last_name LIKE ?
                    OR (d.first_name || ' ' || d.last_name) LIKE ?
                    OR a.appointment_date LIKE ?
                    OR a.status LIKE ?
                )
            '''
            params.extend([search_pattern] * 8)

        # فیلتر وضعیت
        if status_filter and status_filter != 'All':
            query += ' AND a.status = ?'
            params.append(status_filter)

        # ===== Date Filter =====
        if from_date:
            query += " AND a.appointment_date >= ?"
            params.append(from_date)
        if to_date:
            query += " AND a.appointment_date <= ?"
            params.append(to_date)

        query += ' ORDER BY a.appointment_date DESC, a.appointment_time DESC'

        print(f'📊 Query: {query}')
        print(f'📊 Params: {params}')

        cursor.execute(query, params)
        appointments = cursor.fetchall()

        print(f'📊 Appointments loaded: {len(appointments)}')

        cursor.execute(query, params)
        appointments = cursor.fetchall()
        conn.close()

        # پاک کردن
        for item in self.appointments_tree.get_children():
            self.appointments_tree.delete(item)

        # اضافه کردن
        for index, a in enumerate(appointments, 1):
            real_id = a[0]
            doctor_name = a[5].replace('Dr.', '').replace('dr.', '').strip()
            status_icon = '✅' if a[3] == 'Scheduled' else '✔️' if a[3] == 'Completed' else '❌'

            self.appointments_tree.insert(
                '', 'end',
                values=(index, a[1], a[2], a[4], f'Dr. {doctor_name}', f'{status_icon} {a[3]}'),
                tags=(str(real_id),)
            )

        # ===== آپدیت تعداد =====
        if hasattr(self, 'appointment_count_var'):
            active_filters = (
                    search_text or status_filter != 'All'
                    or (hasattr(self, 'appointment_date_filter') and self.appointment_date_filter.get() != 'All')
            )

            filter_info = ''
            if hasattr(self, 'appointment_date_filter') and self.appointment_date_filter.get() != 'All':
                filter_info = f'  |  📅 {self.appointment_date_filter.get()}'

            if active_filters:
                self.appointment_count_var.set(f'🔍 {len(appointments)} result(s){filter_info}')
            else:
                self.appointment_count_var.set(f'📋 {len(appointments)} appointment(s)')

    def add_appointment_window(self):
        """فرم ثبت نوبت جدید"""
        from datetime import datetime

        win = tk.Toplevel(self.root)
        win.title('➕ New Appointment')
        win.configure(bg='#f0f0f0')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#f0f0f0')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#f0f0f0')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f0f0f0')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text='➕ New Appointment',
                 font=('Arial', 16, 'bold'), bg='#f0f0f0',
                 fg='#2c3e50').pack(pady=15)

        # ===== ۴. فرم =====
        form_frame = tk.Frame(scrollable, bg='#f0f0f0')
        form_frame.pack(pady=10, padx=40, fill='x')

        # ===== Patient =====
        tk.Label(form_frame, text='Patient *', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=0, column=0, sticky='w', pady=(5, 0), padx=5)

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT id, first_name, last_name FROM patients ORDER BY last_name')
        patients = cursor.fetchall()
        conn.close()

        if not patients:
            messagebox.showerror('Error', '❌ No patients in database!')
            win.destroy()
            return

        patient_options = [f'{p[0]} - {p[1]} {p[2]}' for p in patients]
        patient_var = tk.StringVar(value=patient_options[0])
        ttk.Combobox(form_frame, textvariable=patient_var, values=patient_options,
                     width=38, state='readonly').grid(
            row=1, column=0, pady=(0, 10), padx=5, sticky='ew')

        # ===== Doctor =====
        tk.Label(form_frame, text='Doctor *', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=2, column=0, sticky='w', pady=(5, 0), padx=5)

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT id, first_name, last_name, specialty FROM doctors ORDER BY last_name')
        doctors = cursor.fetchall()
        conn.close()

        if not doctors:
            messagebox.showerror('Error', '❌ No doctors in database!')
            win.destroy()
            return

        doctor_options = [f'{d[0]} - Dr. {d[1]} {d[2]} ({d[3]})' for d in doctors]
        doctor_var = tk.StringVar(value=doctor_options[0])
        ttk.Combobox(form_frame, textvariable=doctor_var, values=doctor_options,
                     width=38, state='readonly').grid(
            row=3, column=0, pady=(0, 10), padx=5, sticky='ew')

        # ===== Date (✅ تغییر یافته: تاریخ امروز به صورت پیش‌فرض) =====
        tk.Label(form_frame, text='Date * (YYYY-MM-DD)', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=4, column=0, sticky='w', pady=(5, 0), padx=5)

        date_entry = tk.Entry(form_frame, width=40, font=('Arial', 10))
        date_entry.grid(row=5, column=0, pady=(0, 5), padx=5, sticky='ew')

        today_str = datetime.now().strftime('%Y-%m-%d')
        date_entry.insert(0, today_str)  # ✅ تاریخ سیستم جایگزین شد

        tk.Label(form_frame, text=f'💡 Today: {today_str}', bg='#f0f0f0',
                 font=('Arial', 8), fg='#7f8c8d', anchor='w').grid(
            row=5, column=1, sticky='w', padx=5)

        # ===== Time =====
        tk.Label(form_frame, text='Time * (HH:MM)', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=6, column=0, sticky='w', pady=(5, 0), padx=5)
        time_entry = tk.Entry(form_frame, width=40, font=('Arial', 10))
        time_entry.grid(row=7, column=0, pady=(0, 5), padx=5, sticky='ew')
        time_entry.insert(0, '10:30')
        tk.Label(form_frame, text='💡 Example: 10:30', bg='#f0f0f0',
                 font=('Arial', 8), fg='#7f8c8d', anchor='w').grid(
            row=7, column=1, sticky='w', padx=5)

        # ===== Notes =====
        tk.Label(form_frame, text='Notes (Optional)', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=8, column=0, sticky='w', pady=(5, 0), padx=5)
        notes_entry = tk.Entry(form_frame, width=40, font=('Arial', 10))
        notes_entry.grid(row=9, column=0, pady=(0, 5), padx=5, sticky='ew')

        form_frame.grid_columnconfigure(0, weight=1)

        # ===== ۵. تابع ذخیره =====
        def save():
            try:
                patient_id = int(patient_var.get().split(' - ')[0])
                doctor_id = int(doctor_var.get().split(' - ')[0])

                date_str = date_entry.get().strip()
                if not re.match(r'^\d{4}-\d{2}-\d{2}$', date_str):
                    messagebox.showerror('Error', '❌ Invalid date format! Use YYYY-MM-DD')
                    return

                time_str = time_entry.get().strip()
                if not re.match(r'^\d{2}:\d{2}$', time_str):
                    messagebox.showerror('Error', '❌ Invalid time format! Use HH:MM')
                    return

                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT COUNT(*) FROM appointments
                    WHERE doctor_id = ? AND appointment_date = ? AND appointment_time = ?
                    AND status = 'Scheduled'
                ''', (doctor_id, date_str, time_str))
                if cursor.fetchone()[0] > 0:
                    messagebox.showerror('Error', '❌ This time slot is already taken!')
                    conn.close()
                    return

                notes = notes_entry.get().strip()

                cursor.execute('''
                    INSERT INTO appointments (patient_id, doctor_id, appointment_date, appointment_time, status, notes)
                    VALUES (?, ?, ?, ?, 'Scheduled', ?)
                ''', (patient_id, doctor_id, date_str, time_str, notes))
                conn.commit()
                conn.close()

                messagebox.showinfo('Success', f'✅ Appointment created for {date_str} at {time_str}!')
                try:
                    log_action('CREATE', entity_type='appointment',
                               details=f'New appointment: Patient#{patient_id} → Dr#{doctor_id} on {date_str} {time_str}',
                               user=getattr(self.root, 'current_user', None))
                except:
                    pass
                win.destroy()
                self.load_appointments()
                self.refresh_dashboard()

            except Exception as e:
                messagebox.showerror('Error', f'❌ Error: {e}')

        # ===== ۶. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '💾 Save', 'command': save, 'color': 'success', 'width': 160},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#f0f0f0')

        # ===== ۷. Maximize =====
        win.update_idletasks()
        win.state('zoomed')


    def update_appointment_status(self):
        """فرم تغییر وضعیت نوبت"""
        selected = self.appointments_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select an appointment first!')
            return

        real_id = self.appointments_tree.item(selected[0])['tags'][0]
        values = self.appointments_tree.item(selected[0])['values']

        win = tk.Toplevel(self.root)
        win.title('✏️ Update Appointment Status')
        win.configure(bg='#f0f0f0')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#f0f0f0')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#f0f0f0')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f0f0f0')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text='✏️ Update Appointment Status',
                 font=('Arial', 16, 'bold'), bg='#f0f0f0',
                 fg='#2c3e50').pack(pady=15)

        # ===== ۴. اطلاعات نوبت =====
        info_frame = tk.Frame(scrollable, bg='#f8f9fa', relief='solid', bd=1)
        info_frame.pack(fill='x', padx=40, pady=10)

        tk.Label(info_frame, text=f'📅 Appointment #{values[0]}',
                 font=('Arial', 12, 'bold'), bg='#f8f9fa',
                 fg='#2c3e50', anchor='w').pack(anchor='w', padx=15, pady=(10, 5))

        tk.Label(info_frame, text=f'👤 Patient: {values[3]}',
                 font=('Arial', 11), bg='#f8f9fa',
                 anchor='w').pack(anchor='w', padx=15, pady=3)

        tk.Label(info_frame, text=f'👨‍⚕️ Doctor: {values[4]}',
                 font=('Arial', 11), bg='#f8f9fa',
                 anchor='w').pack(anchor='w', padx=15, pady=3)

        tk.Label(info_frame, text=f'📅 Date: {values[1]}',
                 font=('Arial', 11), bg='#f8f9fa',
                 anchor='w').pack(anchor='w', padx=15, pady=3)

        tk.Label(info_frame, text=f'⏰ Time: {values[2]}',
                 font=('Arial', 11), bg='#f8f9fa',
                 anchor='w').pack(anchor='w', padx=15, pady=3)

        tk.Label(info_frame, text=f'📊 Current Status: {values[5]}',
                 font=('Arial', 11, 'bold'), bg='#f8f9fa',
                 fg='#e67e22', anchor='w').pack(anchor='w', padx=15, pady=(3, 10))

        # ===== ۵. انتخاب وضعیت جدید =====
        tk.Label(scrollable, text='📊 New Status:',
                 font=('Arial', 12, 'bold'), bg='#f0f0f0',
                 fg='#2c3e50').pack(anchor='w', padx=40, pady=(15, 5))

        status_var = tk.StringVar(value='Scheduled')

        status_frame = tk.Frame(scrollable, bg='#f0f0f0')
        status_frame.pack(fill='x', padx=40, pady=5)

        statuses = [
            ('Scheduled', '✅ Scheduled (نوبت فعال)'),
            ('Completed', '✔️ Completed (انجام شده)'),
            ('Cancelled', '❌ Cancelled (لغو شده)'),
        ]

        for status, label in statuses:
            tk.Radiobutton(status_frame, text=label,
                           variable=status_var, value=status,
                           bg='#f0f0f0', font=('Arial', 11),
                           activebackground='#f0f0f0',
                           padx=10, pady=5).pack(anchor='w', pady=2)

        # ===== ۶. تابع ذخیره =====
        def apply():
            try:
                new_status = status_var.get()

                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute('UPDATE appointments SET status = ? WHERE id = ?',
                               (new_status, real_id))
                conn.commit()
                conn.close()

                messagebox.showinfo('✅ Success',
                                    f'✅ Status updated to "{new_status}"')

                try:
                    log_action('UPDATE', entity_type='appointment',
                               entity_id=real_id,
                               details=f'Appointment #{real_id} status → {new_status}',
                               user=getattr(self.root, 'current_user', None))
                except:
                    pass

                win.destroy()
                self.load_appointments()
                self.refresh_dashboard()

            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        # ===== ۷. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '💾 Apply', 'command': apply, 'color': 'success', 'width': 160},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#f0f0f0')

        # ===== ۸. Maximize =====
        win.update_idletasks()
        win.state('zoomed')


    def delete_appointment(self):
        selected = self.appointments_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', 'Please select at least one appointment first!')
            return

        count = len(selected)

        if count == 1:
            values = self.appointments_tree.item(selected[0])['values']
            confirm_msg = f'Delete appointment #{values[0]}?\nPatient: {values[3]}'
        else:
            confirm_msg = f'Delete {count} appointment(s)?\n\nThis action cannot be undone!'

        if not messagebox.askyesno('Confirm', confirm_msg):
            return

        try:
            conn = get_connection()
            cursor = conn.cursor()

            for item in selected:
                real_id = self.appointments_tree.item(item)['tags'][0]
                cursor.execute('DELETE FROM appointments WHERE id = ?', (real_id,))

            conn.commit()
            conn.close()

            messagebox.showinfo('Success', f'✅ {count} appointment(s) deleted!')
            self.load_appointments()
            self.refresh_dashboard()

        except Exception as e:
            messagebox.showerror('Error', f'❌ Error: {e}')

    def clean_orphan_appointments(self):
        """پاک کردن نوبت‌های بیمار/دکتر حذف‌شده"""
        # ===== شمارش =====
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT COUNT(*) FROM appointments 
            WHERE patient_id NOT IN (SELECT id FROM patients)
               OR doctor_id NOT IN (SELECT id FROM doctors)
        ''')
        orphan_count = cursor.fetchone()[0]
        conn.close()

        if orphan_count == 0:
            messagebox.showinfo('✅ All Good', 'No orphan appointments found!')
            return

        if not messagebox.askyesno('Confirm',
                                   f'Found {orphan_count} appointment(s) with deleted patient/doctor.\n\n'
                                   f'Delete them?'):
            return

        try:
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                DELETE FROM appointments 
                WHERE patient_id NOT IN (SELECT id FROM patients)
                   OR doctor_id NOT IN (SELECT id FROM doctors)
            ''')
            conn.commit()
            conn.close()

            messagebox.showinfo('✅ Success', f'{orphan_count} orphan appointment(s) deleted!')
            self.load_appointments()

        except Exception as e:
            messagebox.showerror('Error', f'❌ {e}')

    def show_appointment_reminders(self):
        """نمایش یادآوری نوبت‌های امروز و فردا (با اسکرول هوشمند)"""
        from datetime import datetime, timedelta

        today = datetime.now().strftime('%Y-%m-%d')
        tomorrow = (datetime.now() + timedelta(days=1)).strftime('%Y-%m-%d')

        conn = get_connection()
        cursor = conn.cursor()

        # ===== امروز =====
        cursor.execute('''
            SELECT
                a.id, a.appointment_time, a.status,
                COALESCE(p.first_name || ' ' || p.last_name, '— (deleted)') AS patient_name,
                COALESCE(p.phone, '—') AS patient_phone,
                COALESCE(d.first_name || ' ' || d.last_name, '— (deleted)') AS doctor_name,
                COALESCE(d.specialty, '—') AS doctor_specialty
            FROM appointments a
            LEFT JOIN patients p ON a.patient_id = p.id
            LEFT JOIN doctors d ON a.doctor_id = d.id
            WHERE a.appointment_date = ? AND a.status = 'Scheduled'
            ORDER BY a.appointment_time
        ''', (today,))
        today_appointments = cursor.fetchall()

        # ===== فردا =====
        cursor.execute('''
            SELECT
                a.id, a.appointment_time, a.status,
                COALESCE(p.first_name || ' ' || p.last_name, '— (deleted)') AS patient_name,
                COALESCE(p.phone, '—') AS patient_phone,
                COALESCE(d.first_name || ' ' || d.last_name, '— (deleted)') AS doctor_name,
                COALESCE(d.specialty, '—') AS doctor_specialty
            FROM appointments a
            LEFT JOIN patients p ON a.patient_id = p.id
            LEFT JOIN doctors d ON a.doctor_id = d.id
            WHERE a.appointment_date = ? AND a.status = 'Scheduled'
            ORDER BY a.appointment_time
        ''', (tomorrow,))
        tomorrow_appointments = cursor.fetchall()

        conn.close()

        # ===== پنجره =====
        win = tk.Toplevel(self.root)
        win.title('🔔 Appointment Reminders')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        # ===== ۱. هدر (بالا ثابت) =====
        header = tk.Frame(win, bg='#e67e22', height=55)
        header.pack(fill='x', side='top')
        header.pack_propagate(False)

        total = len(today_appointments) + len(tomorrow_appointments)
        tk.Label(header, text=f'🔔 Appointment Reminders ({total} total)',
                 font=('Arial', 15, 'bold'), bg='#e67e22',
                 fg='white').pack(pady=12)

        # ===== ۲. دکمه‌های پایین (ثابت) =====
        btn_frame = tk.Frame(win, bg='#ecf0f1')
        btn_frame.pack(side='bottom', fill='x', pady=12)

        # ===== ۳. بدنه اصلی =====
        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=5, pady=5)

        canvas = tk.Canvas(body_container, bg='#ecf0f1', highlightthickness=0)
        v_scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        h_scrollbar = ttk.Scrollbar(body_container, orient='horizontal', command=canvas.xview)

        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        # چیدمان اولیه (اسکرول‌ها در ابتدا پک می‌شوند تا بعدا مخفی شوند)
        canvas.pack(side='left', fill='both', expand=True)
        v_scrollbar.pack(side='right', fill='y')
        h_scrollbar.pack(side='bottom', fill='x')

        # --- فریم داخلی ---
        scrollable = tk.Frame(canvas, bg='#ecf0f1')
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')

        # ===== ترفند حداقل عرض 850 پیکسل =====
        def _on_canvas_configure(event):
            width = max(event.width, 850)
            canvas.itemconfig(canvas_window, width=width)

        canvas.bind('<Configure>', _on_canvas_configure)

        # ===== 🌟 تابع جادویی برای مخفی/نمایان کردن خودکار اسکرول‌ها =====
        def check_scrollbars(event=None):
            canvas.configure(scrollregion=canvas.bbox('all'))
            bbox = canvas.bbox('all')
            if not bbox:
                return

            content_width = bbox[2] - bbox[0]
            content_height = bbox[3] - bbox[1]

            canvas_width = canvas.winfo_width()
            canvas_height = canvas.winfo_height()

            # مدیریت اسکرول عمودی
            if content_height > canvas_height + 2:
                if not v_scrollbar.winfo_ismapped():
                    v_scrollbar.pack(side='right', fill='y')
            else:
                if v_scrollbar.winfo_ismapped():
                    v_scrollbar.pack_forget()

            # مدیریت اسکرول افقی
            if content_width > canvas_width + 2:
                if not h_scrollbar.winfo_ismapped():
                    h_scrollbar.pack(side='bottom', fill='x')
            else:
                if h_scrollbar.winfo_ismapped():
                    h_scrollbar.pack_forget()

        scrollable.bind('<Configure>', check_scrollbars)
        # استفاده از add='+' برای اینکه رویداد _on_canvas_configure پاک نشود
        canvas.bind('<Configure>', check_scrollbars, add='+')

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== محتوای اصلی =====
        content_frame = tk.Frame(scrollable, bg='#ecf0f1')
        content_frame.pack(fill='x', padx=20, pady=10)

        def make_appointment_card(parent, appt):
            appt_id, appt_time, status, patient_name, phone, doctor_name, specialty = appt

            # بدون pack_propagate(False) تا ارتفاع محتوا آزاد باشد
            card = tk.Frame(parent, bg='white', relief='solid', bd=1)
            card.pack(fill='x', pady=4)

            card_inner = tk.Frame(card, bg='white')
            card_inner.pack(fill='both', expand=True, padx=15, pady=10)

            r1 = tk.Frame(card_inner, bg='white')
            r1.pack(fill='x')

            tk.Label(r1, text=f'⏰ {appt_time}',
                     font=('Arial', 13, 'bold'),
                     bg='white', fg='#e67e22',
                     width=10, anchor='w').pack(side='left')

            tk.Label(r1, text=f'👤 {patient_name}',
                     font=('Arial', 11, 'bold'),
                     bg='white', fg='#2c3e50',
                     anchor='w').pack(side='left', padx=10)

            tk.Label(r1, text=f'📞 {phone}',
                     font=('Arial', 9),
                     bg='white', fg='#7f8c8d',
                     anchor='w').pack(side='left', padx=10)

            r2 = tk.Frame(card_inner, bg='white')
            r2.pack(fill='x', pady=(5, 0))

            doctor_clean = doctor_name.replace('Dr.', '').replace('dr.', '').strip()

            tk.Label(r2, text=f'   👨‍⚕️ Dr. {doctor_clean}',
                     font=('Arial', 10),
                     bg='white', fg='#3498db',
                     anchor='w').pack(side='left')

            tk.Label(r2, text=f'({specialty})',
                     font=('Arial', 9, 'italic'),
                     bg='white', fg='#95a5a6',
                     anchor='w').pack(side='left', padx=5)

        # ===== TODAY =====
        today_frame = tk.LabelFrame(content_frame,
                                    text=f'  📅 TODAY - {today}  ({len(today_appointments)} appointment{"s" if len(today_appointments) != 1 else ""})  ',
                                    font=('Arial', 12, 'bold'),
                                    bg='#ecf0f1', fg='#27ae60',
                                    padx=10, pady=10)
        today_frame.pack(fill='x', pady=(5, 15))

        if today_appointments:
            for appt in today_appointments:
                make_appointment_card(today_frame, appt)
        else:
            tk.Label(today_frame, text='📭 No appointments today',
                     font=('Arial', 11, 'italic'),
                     bg='#ecf0f1', fg='#95a5a6').pack(pady=20)

        # ===== TOMORROW =====
        tomorrow_frame = tk.LabelFrame(content_frame,
                                       text=f'  📅 TOMORROW - {tomorrow}  ({len(tomorrow_appointments)} appointment{"s" if len(tomorrow_appointments) != 1 else ""})  ',
                                       font=('Arial', 12, 'bold'),
                                       bg='#ecf0f1', fg='#3498db',
                                       padx=10, pady=10)
        tomorrow_frame.pack(fill='x', pady=(5, 15))

        if tomorrow_appointments:
            for appt in tomorrow_appointments:
                make_appointment_card(tomorrow_frame, appt)
        else:
            tk.Label(tomorrow_frame, text='📭 No appointments tomorrow',
                     font=('Arial', 11, 'italic'),
                     bg='#ecf0f1', fg='#95a5a6').pack(pady=20)

        # ===== ۵. دکمه‌ها =====
        create_popup_buttons(btn_frame, [
            {'text': '🔄 Refresh',
             'command': lambda: [win.destroy(), self.show_appointment_reminders()],
             'color': 'primary', 'width': 140},
            {'text': '❌ Close', 'command': win.destroy,
             'color': 'danger', 'width': 120},
        ], bg='#ecf0f1')

        # ===== ۶. Maximize =====
        win.update_idletasks()
        win.state('zoomed')

        # یک بار اجرای تابع بعد از باز شدن پنجره برای تنظیم اولیه
        win.after(100, check_scrollbars)


    # ==========================================
    # Payments Tab
    # ==========================================

    def create_payments_tab(self):
        tk.Label(self.payments_tab, text='💰 Payment Management',
                 font=('Arial', 14, 'bold'), bg=WINDOW_BG, fg='#2c3e50').pack(pady=15)

        btn_bar = create_button_bar(self.payments_tab, bg=WINDOW_BG)

        create_standard_button(btn_bar, '➕ New', self.add_payment_window, color='success')
        create_standard_button(btn_bar, '📊 Visit Report', self.visit_full_report_window, color='secondary')
        create_standard_button(btn_bar, '📊 Report', self.show_daily_report, color='warning')
        create_standard_button(btn_bar, '👁️ Refresh', self.load_payments, color='primary')
        create_standard_button(btn_bar, '🖨️ Print Selected', self.batch_print_payments, color='secondary')  # ← جدید
        create_standard_button(btn_bar, '🗑️ Delete', self.delete_payment, color='danger')

        # ===== Search Bar =====
        # ===== Search Bar (2 rows) =====
        # ردیف ۱: Search + Clear + Count
        row1 = tk.Frame(self.payments_tab, bg=WINDOW_BG)
        row1.pack(pady=(5, 3), padx=40, fill='x')

        tk.Label(row1, text='🔍 Search:', font=('Arial', 11, 'bold'),
                 bg=WINDOW_BG, fg='#2c3e50').pack(side='left', padx=(0, 8))

        self.payment_search_var = tk.StringVar()
        search_entry = tk.Entry(row1, textvariable=self.payment_search_var,
                                font=('Arial', 11), bg='white',
                                relief='solid', bd=1, width=35)
        search_entry.pack(side='left', ipady=6, padx=(0, 8))

        self.payment_search_var.trace_add('write',
                                          lambda *args: self.load_payments())

        def clear_search():
            self.payment_search_var.set('')

        tk.Button(row1, text='✖ Clear', command=clear_search,
                  bg='#95a5a6', fg='white', font=('Arial', 9, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=10, pady=4).pack(side='left', padx=(0, 15))

        # ===== Date Filter =====
        filter_var, from_var, to_var, df_container, custom_frame = create_date_filter(
            row1, WINDOW_BG, lambda: self.load_payments()
        )
        self.payment_date_filter = filter_var
        self.payment_date_from = from_var
        self.payment_date_to = to_var
        df_container.pack(side='left')

        # ===== Method Filter =====
        tk.Label(row1, text='  Method:', font=('Arial', 10, 'bold'),
                 bg=WINDOW_BG, fg='#2c3e50').pack(side='left', padx=(15, 5))

        self.payment_method_filter = tk.StringVar(value='All')
        method_combo = ttk.Combobox(row1, textvariable=self.payment_method_filter,
                                    values=['All', 'Cash', 'Card', 'Insurance', 'Online'],
                                    state='readonly', width=10, font=('Arial', 10))
        method_combo.pack(side='left')
        method_combo.bind('<<ComboboxSelected>>', lambda e: self.load_payments())

        # ===== Status Filter =====
        tk.Label(row1, text='  Status:', font=('Arial', 10, 'bold'),
                 bg=WINDOW_BG, fg='#2c3e50').pack(side='left', padx=(15, 5))

        self.payment_status_filter = tk.StringVar(value='All')
        status_combo = ttk.Combobox(row1, textvariable=self.payment_status_filter,
                                    values=['All', 'Paid', 'Pending'],
                                    state='readonly', width=10, font=('Arial', 10))
        status_combo.pack(side='left')
        status_combo.bind('<<ComboboxSelected>>', lambda e: self.load_payments())

        # ===== Total & Count =====
        info_frame = tk.Frame(self.payments_tab, bg=WINDOW_BG)
        info_frame.pack(pady=(0, 5), padx=40, fill='x')

        self.payment_count_var = tk.StringVar(value='')
        tk.Label(info_frame, textvariable=self.payment_count_var,
                 font=('Arial', 10, 'italic'), bg=WINDOW_BG,
                 fg='#7f8c8d').pack(side='left')

        self.payment_total_var = tk.StringVar(value='')
        tk.Label(info_frame, textvariable=self.payment_total_var,
                 font=('Arial', 11, 'bold'), bg=WINDOW_BG,
                 fg='#27ae60').pack(side='right')

        # ===== Total & Count =====
        info_frame = tk.Frame(self.payments_tab, bg=WINDOW_BG)
        info_frame.pack(pady=(0, 5), padx=40, fill='x')

        self.payment_count_var = tk.StringVar(value='')
        tk.Label(info_frame, textvariable=self.payment_count_var,
                 font=('Arial', 10, 'italic'), bg=WINDOW_BG,
                 fg='#7f8c8d').pack(side='left')

        self.payment_total_var = tk.StringVar(value='')
        tk.Label(info_frame, textvariable=self.payment_total_var,
                 font=('Arial', 11, 'bold'), bg=WINDOW_BG,
                 fg='#27ae60').pack(side='right')

        # ===== Treeview =====
        columns = ('ID', 'Receipt', 'Patient', 'Amount', 'Method', 'Status', 'Date')
        self.payments_tree = ttk.Treeview(self.payments_tab, columns=columns,
                                          show='headings', height=15,
                                          selectmode='extended')  # ← اضافه کنید
        for col in columns:
            self.payments_tree.heading(col, text=col)
            self.payments_tree.column(col, width=130, anchor='center')
        self.payments_tree.pack(pady=10, padx=40, fill='both', expand=True)
        self.load_payments()

        enable_multi_click_selection(self.payments_tree)

    def load_payments(self):
        # ===== خواندن فیلترها =====
        search_text = ''
        if hasattr(self, 'payment_search_var'):
            search_text = self.payment_search_var.get().strip()

        method_filter = 'All'
        if hasattr(self, 'payment_method_filter'):
            method_filter = self.payment_method_filter.get()

        status_filter = 'All'
        if hasattr(self, 'payment_status_filter'):
            status_filter = self.payment_status_filter.get()

        # ===== Date Filter =====
        from_date, to_date = None, None
        if hasattr(self, 'payment_date_filter'):
            from_date, to_date = get_date_range(
                self.payment_date_filter.get(),
                self.payment_date_from.get().strip() if hasattr(self, 'payment_date_from') else '',
                self.payment_date_to.get().strip() if hasattr(self, 'payment_date_to') else '',
            )

        conn = get_connection()
        cursor = conn.cursor()

        query = '''
            SELECT
                p.id, p.receipt_number, p.final_amount, p.payment_method,
                p.payment_status, p.payment_date,
                COALESCE(pt.first_name || ' ' || pt.last_name, '— (deleted)') AS patient_name
            FROM payments p
            LEFT JOIN patients pt ON p.patient_id = pt.id
            WHERE 1=1
        '''

        params = []

        # جستجو
        if search_text:
            search_pattern = f'%{search_text}%'
            query += '''
                AND (
                    p.receipt_number LIKE ?
                    OR pt.first_name LIKE ?
                    OR pt.last_name LIKE ?
                    OR (pt.first_name || ' ' || pt.last_name) LIKE ?
                    OR p.payment_method LIKE ?
                    OR p.payment_status LIKE ?
                    OR CAST(p.final_amount AS TEXT) LIKE ?
                )
            '''
            params.extend([search_pattern] * 7)

        # فیلتر روش
        if method_filter and method_filter != 'All':
            query += ' AND p.payment_method = ?'
            params.append(method_filter)

        # فیلتر وضعیت
        if status_filter and status_filter != 'All':
            query += ' AND p.payment_status = ?'
            params.append(status_filter)

        # ===== Date Filter =====
        if from_date:
            query += " AND DATE(p.payment_date) >= ?"
            params.append(from_date)
        if to_date:
            query += " AND DATE(p.payment_date) <= ?"
            params.append(to_date)

        query += ' ORDER BY p.payment_date DESC'

        cursor.execute(query, params)
        payments = cursor.fetchall()
        conn.close()

        # پاک کردن
        for item in self.payments_tree.get_children():
            self.payments_tree.delete(item)

        # ===== محاسبه مجموع =====
        total_paid = 0
        total_pending = 0

        for index, p in enumerate(payments, 1):
            real_id = p[0]
            status_icon = '✅' if p[4] == 'Paid' else '⏳' if p[4] == 'Pending' else '↩️'

            if p[4] == 'Paid':
                total_paid += p[2]
            elif p[4] == 'Pending':
                total_pending += p[2]

            self.payments_tree.insert(
                '', 'end',
                values=(index, p[1], p[6], f'{p[2]:,.0f}', p[3],
                        f'{status_icon} {p[4]}', p[5][:10]),
                tags=(str(real_id),)
            )

        # ===== آپدیت اطلاعات =====
        if hasattr(self, 'payment_count_var'):
            active_filters = (
                    search_text or method_filter != 'All'
                    or status_filter != 'All'
                    or (hasattr(self, 'payment_date_filter') and self.payment_date_filter.get() != 'All')
            )
            if active_filters:
                self.payment_count_var.set(f'🔍 {len(payments)} result(s)')
            else:
                self.payment_count_var.set(f'📋 {len(payments)} payment(s)')

        if hasattr(self, 'payment_total_var'):
            # ✅ محاسبه جمع کل
            grand_total = total_paid + total_pending

            parts = []
            if grand_total > 0:
                parts.append(f'💰 Total: {grand_total:,.0f}')
            if total_paid > 0:
                parts.append(f'✅ Paid: {total_paid:,.0f}')
            if total_pending > 0:
                parts.append(f'⏳ Pending: {total_pending:,.0f}')

            # اضافه کردن فیلتر به خلاصه
            filter_info = ''
            if hasattr(self, 'payment_date_filter') and self.payment_date_filter.get() != 'All':
                filter_info = f'  |  📅 {self.payment_date_filter.get()}'

            text = '  |  '.join(parts) if parts else ''
            if text:
                text += ' IRR' + filter_info

            self.payment_total_var.set(text)

    def batch_print_payments(self):
        """چاپ گروهی پرداخت‌ها"""
        selected = self.payments_tree.selection()

        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select at least one payment!\n\n'
                                              '💡 Tip: Use Ctrl+Click or Shift+Click to select multiple.')
            return

        # ===== تایید =====
        if not messagebox.askyesno(
                '🖨️ Batch Print',
                f'Print {len(selected)} payment receipt(s)?'
        ):
            return

        # ===== گرفتن پرینتر =====
        printer = get_printer_for('receipt')

        if not printer and WIN32_AVAILABLE:
            # اگه پرینتر نیست، از پیش‌فرض استفاده کن
            printer = get_default_printer()

        # ===== ساخت PDF ها =====
        pdf_paths = []
        errors = []

        conn = get_connection()
        cursor = conn.cursor()

        # ===== Progress window =====
        progress_win = tk.Toplevel(self.root)
        progress_win.title('🖨️ Printing...')
        progress_win.geometry('400x150')
        progress_win.configure(bg='white')
        progress_win.resizable(False, False)

        progress_win.update_idletasks()
        w, h = 400, 150
        x = (progress_win.winfo_screenwidth() // 2) - (w // 2)
        y = (progress_win.winfo_screenheight() // 2) - (h // 2)
        progress_win.geometry(f'{w}x{h}+{x}+{y}')

        tk.Label(progress_win, text='🖨️ Preparing PDFs...',
                 font=('Arial', 13, 'bold'), bg='white',
                 fg='#2c3e50').pack(pady=15)

        progress_var = tk.StringVar(value=f'0 / {len(selected)}')
        tk.Label(progress_win, textvariable=progress_var,
                 font=('Arial', 11), bg='white').pack()

        progress_bar = ttk.Progressbar(progress_win, length=350,
                                       mode='determinate',
                                       maximum=len(selected))
        progress_bar.pack(pady=10, padx=25)

        progress_win.update()

        # ===== پردازش هر پرداخت =====
        for idx, item in enumerate(selected, 1):
            try:
                tags = self.payments_tree.item(item)['tags']
                if not tags:
                    errors.append(f'Row {idx}: No payment ID')
                    continue

                payment_id = tags[0]

                # خواندن اطلاعات پرداخت
                cursor.execute('''
                    SELECT
                        p.id, p.receipt_number, p.final_amount,
                        p.payment_method, p.payment_status,
                        p.payment_date, p.discount, p.amount, p.notes,
                        pt.first_name || ' ' || pt.last_name,
                        pt.national_id, pt.phone
                    FROM payments p
                    JOIN patients pt ON p.patient_id = pt.id
                    WHERE p.id = ?
                ''', (payment_id,))
                payment = cursor.fetchone()

                if not payment:
                    errors.append(f'Payment #{payment_id} not found')
                    continue

                # ساخت متن فاکتور
                content = self._build_payment_receipt_text(payment)

                success, pdf_path = print_text_file(
                    content,
                    f'Receipt_{payment[1]}.txt',  # ← شماره رسید
                    open_pdf=False
                )

                if success:
                    pdf_paths.append(pdf_path)
                else:
                    errors.append(f'PDF error for #{payment_id}')

                # آپدیت progress
                progress_var.set(f'{idx} / {len(selected)}')
                progress_bar['value'] = idx
                progress_win.update()

            except Exception as e:
                errors.append(f'Error: {str(e)}')

        conn.close()

        progress_win.destroy()

        # ===== چاپ =====
        if not pdf_paths:
            messagebox.showerror('❌ Error',
                                 'Could not prepare any PDFs!\n\n' +
                                 '\n'.join(errors[:5]))
            return

        if printer:
            # ===== چک کن آیا مجازیه یا فیزیکی =====
            if is_virtual_printer(printer):
                # 🎯 پرینتر مجازیه → PDF ها رو ذخیره کن
                save_dir = os.path.join(os.path.expanduser('~'), 'Desktop', 'Hospital PDFs')
                os.makedirs(save_dir, exist_ok=True)

                saved_files = []
                for pdf in pdf_paths:
                    try:
                        dest = os.path.join(save_dir, os.path.basename(pdf))
                        shutil.copy2(pdf, dest)
                        saved_files.append(dest)
                    except:
                        pass

                # باز کردن پوشه
                try:
                    os.startfile(save_dir)
                except:
                    pass

                messagebox.showinfo(
                    '📄 PDFs Saved',
                    f'✅ {len(saved_files)} PDF(s) saved!\n\n'
                    f'📁 Location:\n{save_dir}\n\n'
                    f'💡 Folder has been opened.\n'
                    f'You can now view, print, or share the PDFs.'
                )
            else:
                # 🎯 پرینتر فیزیکی → چاپ مستقیم
                success_count, error_count, print_errors = print_multiple_pdfs(
                    pdf_paths, printer
                )

                if error_count == 0:
                    messagebox.showinfo('✅ Success',
                                        f'🖨️ Batch Print Complete!\n\n'
                                        f'✅ {success_count} printed successfully!')
                else:
                    msg = f'⚠️ Partial Success\n\n'
                    msg += f'✅ Success: {success_count}\n'
                    msg += f'❌ Failed: {error_count}\n\n'
                    msg += 'Errors:\n' + '\n'.join(print_errors[:3])
                    messagebox.showwarning('⚠️ Partial Success', msg)
        else:
            # پرینتر نیست → PDF ذخیره کن
            save_dir = os.path.join(os.path.expanduser('~'), 'Desktop', 'Hospital PDFs')
            os.makedirs(save_dir, exist_ok=True)

            for pdf in pdf_paths:
                try:
                    dest = os.path.join(save_dir, os.path.basename(pdf))
                    shutil.copy2(pdf, dest)
                except:
                    pass

            try:
                os.startfile(save_dir)
            except:
                pass

            messagebox.showinfo(
                '📄 PDFs Saved',
                f'{len(pdf_paths)} PDF(s) saved to:\n\n{save_dir}'
            )

    def _build_payment_receipt_text(self, payment):
        """ساخت متن فاکتور پرداخت با جزئیات سرویس‌ها"""
        (p_id, receipt, final_amount, method, status,
         p_date, discount, amount, notes, patient_name, nid, phone) = payment

        # ===== خواندن سرویس‌های این پرداخت =====
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT service_name, quantity, unit_price, subtotal
            FROM payment_items
            WHERE payment_id = ?
            ORDER BY id
        ''', (p_id,))
        items = cursor.fetchall()
        conn.close()

        lines = []
        lines.append('=' * 55)
        lines.append('         🏥 HOSPITAL MANAGEMENT SYSTEM')
        lines.append('                PAYMENT RECEIPT')
        lines.append('=' * 55)
        lines.append('')
        lines.append(f'📅 Date: {p_date}')
        lines.append(f'🧾 Receipt: {receipt}')
        lines.append('')
        lines.append('-' * 55)
        lines.append('👤 PATIENT:')
        lines.append('-' * 55)
        lines.append(f'   Name: {patient_name}')
        lines.append(f'   National ID: {nid}')
        lines.append(f'   Phone: {phone}')
        lines.append('')
        lines.append('-' * 55)
        lines.append('📋 SERVICES:')
        lines.append('-' * 55)

        # ===== نمایش سرویس‌ها =====
        if items:
            for item_name, qty, unit_price, subtotal in items:
                lines.append(f'   {item_name}')
                lines.append(f'      {qty} x {unit_price:,.0f} = {subtotal:,.0f} IRR')
        else:
            lines.append('   (No services details available)')
            lines.append(f'   Total Services: {amount:,.0f} IRR')

        lines.append('')
        lines.append('-' * 55)
        lines.append(f'  Subtotal:    {amount:>20,.0f} IRR')
        lines.append(f'  Discount:    {discount:>20,.0f} IRR')
        lines.append('-' * 55)
        lines.append(f'  TOTAL:       {final_amount:>20,.0f} IRR')
        lines.append('-' * 55)
        lines.append('')
        lines.append(f'💳 Payment Method: {method}')
        lines.append(f'📊 Status: {status}')
        if notes:
            lines.append(f'📝 Notes: {notes}')
        lines.append('')
        lines.append('=' * 55)
        lines.append('   Thank you for choosing our hospital!')
        lines.append('=' * 55)

        return '\n'.join(lines)



    def add_payment_window(self):
        """Open a full-screen window to add a new payment"""
        win = tk.Toplevel(self.root)
        win.title('💰 New Payment')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)
        win.bind('<Escape>', lambda e: win.destroy())
        win.attributes('-fullscreen', True)

        # ===== دکمه‌ها و Tip (پایین ثابت) =====
        bottom = tk.Frame(win, bg='#ecf0f1')
        bottom.pack(fill='x', side='bottom')

        tk.Label(bottom, text='💡 Fill the form, then click "SAVE PAYMENT"',
                 bg='#f39c12', fg='white', font=('Arial', 10, 'bold'),
                 pady=8).pack(fill='x')

        btn_container = tk.Frame(bottom, bg='#ecf0f1')
        btn_container.pack(fill='x', pady=10)

        # ===== متغیرهای سراسری برای دسترسی از همه‌جا =====
        service_vars = {}
        qty_vars = {}
        subtotal_labels = {}

        # ===== هدر (بالا) =====
        header = tk.Frame(win, bg='#2c3e50', height=60)
        header.pack(fill='x', side='top')
        header.pack_propagate(False)
        tk.Label(header, text='💰 New Payment',
                 font=('Arial', 16, 'bold'), bg='#2c3e50', fg='white').pack(pady=15)

        # ===== محتوا (وسط) =====
        main_frame = tk.Frame(win, bg='#ecf0f1')
        main_frame.pack(fill='both', expand=True, padx=15, pady=10)
        main_frame.grid_columnconfigure(0, weight=3)
        main_frame.grid_columnconfigure(1, weight=2)
        main_frame.grid_rowconfigure(0, weight=1)

        # ===== LEFT: Patient & Services =====
        left_frame = tk.LabelFrame(main_frame, text=' 📋 Patient & Services ',
                                   bg='#ecf0f1', font=('Arial', 11, 'bold'),
                                   fg='#2c3e50', padx=10, pady=10)
        left_frame.grid(row=0, column=0, sticky='nsew', padx=5)

        tk.Label(left_frame, text='Patient:', bg='#ecf0f1',
                 font=('Arial', 11, 'bold'), anchor='w').pack(anchor='w', pady=(0, 5))

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT id, first_name, last_name FROM patients ORDER BY last_name')
        patients = cursor.fetchall()
        cursor.execute('SELECT id, name, price FROM services ORDER BY name')
        services = cursor.fetchall()
        conn.close()

        if not patients or not services:
            messagebox.showerror('Error', '❌ No patients or services found!')
            win.destroy()
            return

        patient_options = [f'{p[0]} - {p[1]} {p[2]}' for p in patients]
        patient_var = tk.StringVar(value=patient_options[0])
        patient_combo = ttk.Combobox(left_frame, textvariable=patient_var,
                                     values=patient_options, width=50,
                                     state='readonly', font=('Arial', 10))
        patient_combo.pack(pady=(0, 10), fill='x')

        tk.Label(left_frame, text='Services:', bg='#ecf0f1',
                 font=('Arial', 11, 'bold'), anchor='w').pack(anchor='w', pady=(5, 5))

        canvas = tk.Canvas(left_frame, bg='white', highlightthickness=1,
                           highlightbackground='#bdc3c7')
        scrollbar = tk.Scrollbar(left_frame, orient='vertical', command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg='white')
        scrollable_frame.bind("<Configure>",
                              lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        canvas.bind('<Configure>', lambda e: scrollable_frame.config(width=e.width))

        header_row = tk.Frame(scrollable_frame, bg='#3498db')
        header_row.pack(fill='x', padx=2, pady=2)
        tk.Label(header_row, text='Service', bg='#3498db', fg='white',
                 font=('Arial', 10, 'bold'), width=25, anchor='w').pack(side='left', padx=5)
        tk.Label(header_row, text='Price', bg='#3498db', fg='white',
                 font=('Arial', 10, 'bold'), width=12, anchor='e').pack(side='left', padx=5)
        tk.Label(header_row, text='Qty', bg='#3498db', fg='white',
                 font=('Arial', 10, 'bold'), width=5).pack(side='left', padx=5)
        tk.Label(header_row, text='Subtotal', bg='#3498db', fg='white',
                 font=('Arial', 10, 'bold'), width=15, anchor='e').pack(side='left', padx=5)

        for s in services:
            row = tk.Frame(scrollable_frame, bg='white', relief='ridge', bd=1)
            row.pack(fill='x', padx=2, pady=1)
            var = tk.BooleanVar(value=False)
            service_vars[s[0]] = var
            tk.Checkbutton(row, text=s[1], variable=var, bg='white',
                           font=('Arial', 10), width=25, anchor='w',
                           command=lambda: update_totals()).pack(side='left', padx=5)
            tk.Label(row, text=f'{s[2]:,.0f}', bg='white',
                     font=('Arial', 10), width=12, anchor='e').pack(side='left', padx=5)
            qty_var = tk.StringVar(value='1')
            qty_vars[s[0]] = qty_var
            tk.Spinbox(row, from_=1, to=100, textvariable=qty_var,
                       width=5, font=('Arial', 10),
                       command=lambda: update_totals()).pack(side='left', padx=5)
            subtotal_label = tk.Label(row, text='0', bg='white',
                                      font=('Arial', 10, 'bold'), width=15,
                                      anchor='e', fg='#27ae60')
            subtotal_labels[s[0]] = subtotal_label
            subtotal_label.pack(side='left', padx=5)

        # ===== RIGHT: Payment Details =====
        right_frame = tk.LabelFrame(main_frame, text=' 💳 Payment Details ',
                                    bg='#ecf0f1', font=('Arial', 11, 'bold'),
                                    fg='#2c3e50', padx=10, pady=10)
        right_frame.grid(row=0, column=1, sticky='nsew', padx=5)

        tk.Label(right_frame, text='Discount (IRR):', bg='#ecf0f1',
                 font=('Arial', 10, 'bold'), anchor='w').pack(anchor='w', pady=(0, 5))
        discount_var = tk.StringVar(value='0')
        discount_entry = tk.Entry(right_frame, textvariable=discount_var,
                                   font=('Arial', 11), width=25)
        discount_entry.pack(pady=(0, 10), fill='x')
        # ✅ Bind برای آپدیت خودکار تخفیف
        discount_entry.bind('<KeyRelease>', lambda e: update_totals())

        tk.Label(right_frame, text='Payment Method:', bg='#ecf0f1',
                 font=('Arial', 10, 'bold'), anchor='w').pack(anchor='w', pady=(0, 5))
        method_var = tk.StringVar(value='Cash')
        for method in ['Cash', 'Card', 'Insurance', 'Online']:
            tk.Radiobutton(right_frame, text=method, variable=method_var,
                           value=method, bg='#ecf0f1',
                           font=('Arial', 10)).pack(anchor='w')

        tk.Label(right_frame, text='Status:', bg='#ecf0f1',
                 font=('Arial', 10, 'bold'), anchor='w').pack(anchor='w', pady=(10, 5))
        status_var = tk.StringVar(value='Paid')
        for status in ['Paid', 'Pending']:
            tk.Radiobutton(right_frame, text=status, variable=status_var,
                           value=status, bg='#ecf0f1',
                           font=('Arial', 10)).pack(anchor='w')

        tk.Label(right_frame, text='Notes:', bg='#ecf0f1',
                 font=('Arial', 10, 'bold'), anchor='w').pack(anchor='w', pady=(10, 5))
        notes_entry = tk.Entry(right_frame, font=('Arial', 10), width=25)
        notes_entry.pack(pady=(0, 10), fill='x')

        summary_frame = tk.Frame(right_frame, bg='white', relief='solid', bd=2)
        summary_frame.pack(fill='x', pady=(10, 5))
        summary_frame.grid_columnconfigure(0, weight=1)
        summary_frame.grid_columnconfigure(1, weight=1)

        tk.Label(summary_frame, text='Subtotal:', bg='white',
                 font=('Arial', 11), anchor='w').grid(row=0, column=0, sticky='w', padx=10, pady=5)
        subtotal_label = tk.Label(summary_frame, text='0 IRR', bg='white',
                                  font=('Arial', 11, 'bold'), anchor='e', fg='#2c3e50')
        subtotal_label.grid(row=0, column=1, sticky='e', padx=10, pady=5)

        tk.Label(summary_frame, text='Discount:', bg='white',
                 font=('Arial', 11), anchor='w').grid(row=1, column=0, sticky='w', padx=10, pady=5)
        discount_display = tk.Label(summary_frame, text='0 IRR', bg='white',
                                    font=('Arial', 11, 'bold'), anchor='e', fg='#e74c3c')
        discount_display.grid(row=1, column=1, sticky='e', padx=10, pady=5)

        tk.Label(summary_frame, text='TOTAL:', bg='white',
                 font=('Arial', 13, 'bold'), anchor='w', fg='#27ae60').grid(row=2, column=0, sticky='w', padx=10, pady=8)
        total_label = tk.Label(summary_frame, text='0 IRR', bg='white',
                               font=('Arial', 13, 'bold'), anchor='e', fg='#27ae60')
        total_label.grid(row=2, column=1, sticky='e', padx=10, pady=8)

        # ===== تابع آپدیت جمع‌ها =====
        def update_totals(*args):
            subtotal = 0
            for s in services:
                if service_vars[s[0]].get():
                    try:
                        qty = int(qty_vars[s[0]].get() or 1)
                    except ValueError:
                        qty = 1
                    line_total = s[2] * qty
                    subtotal += line_total
                    subtotal_labels[s[0]].config(text=f'{line_total:,.0f}')
                else:
                    subtotal_labels[s[0]].config(text='0')

            try:
                discount = float(discount_var.get() or 0)
            except ValueError:
                discount = 0

            total = max(subtotal - discount, 0)
            subtotal_label.config(text=f'{subtotal:,.0f} IRR')
            discount_display.config(text=f'-{discount:,.0f} IRR')
            total_label.config(text=f'{total:,.0f} IRR')

        # ===== تابع ذخیره =====
        def save(event=None):
            try:
                update_totals()
                selected_services = [s for s in services if service_vars[s[0]].get()]
                if not selected_services:
                    messagebox.showerror('Error', '❌ Please select at least one service!')
                    return
                patient_id = int(patient_var.get().split(' - ')[0])
                subtotal = 0
                for s in selected_services:
                    try:
                        qty = int(qty_vars[s[0]].get() or 1)
                    except ValueError:
                        qty = 1
                    subtotal += s[2] * qty
                try:
                    discount = float(discount_var.get() or 0)
                except ValueError:
                    discount = 0
                final_amount = subtotal - discount
                if final_amount < 0:
                    messagebox.showerror('Error', '❌ Discount cannot exceed total!')
                    return

                import random
                from datetime import datetime
                receipt = f'RCP-{datetime.now().strftime("%Y%m%d%H%M%S")}-{random.randint(100, 999)}'

                conn = get_connection()
                cursor = conn.cursor()
                from datetime import datetime
                now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

                cursor.execute('''
                    INSERT INTO payments
                    (receipt_number, patient_id, amount, discount, final_amount, 
                     payment_method, payment_status, notes, payment_date)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (receipt, patient_id, subtotal, discount, final_amount,
                      method_var.get(), status_var.get(), notes_entry.get().strip(), now))

                payment_id = cursor.lastrowid

                # ===== ذخیره سرویس‌های انتخاب‌شده =====
                for s in selected_services:
                    try:
                        qty = int(qty_vars[s[0]].get() or 1)
                    except ValueError:
                        qty = 1
                    line_total = s[2] * qty

                    cursor.execute('''
                        INSERT INTO payment_items
                        (payment_id, service_id, service_name, quantity, unit_price, subtotal)
                        VALUES (?, ?, ?, ?, ?, ?)
                    ''', (payment_id, s[0], s[1], qty, s[2], line_total))
                conn.commit()
                conn.close()
                self.load_payments()
                self.refresh_dashboard()
                messagebox.showinfo('Success',
                                    f'✅ Payment recorded!\n\n'
                                    f'🧾 Receipt: {receipt}\n'
                                    f'💵 TOTAL: {final_amount:,.0f} IRR')
                log_action(
                    'CREATE',
                    entity_type='payment',
                    entity_id=None,
                    details=f'Payment: {receipt} for {final_amount:,.0f} IRR ({method_var.get()})',
                    user=getattr(self.root, 'current_user', None)
                )
                win.destroy()
            except Exception as e:
                messagebox.showerror('Error', f'❌ Error: {e}')

        # ===== تابع نمایش فاکتور پیش‌نمایش =====
        def show_receipt():
            selected_services = [s for s in services if service_vars[s[0]].get()]
            if not selected_services:
                messagebox.showwarning('Warning', '⚠️ Please select at least one service first!')
                return

            patient_id = int(patient_var.get().split(' - ')[0])
            patient_name = patient_var.get().split(' - ')[1]

            subtotal = 0
            receipt_lines = []
            for s in selected_services:
                try:
                    qty = int(qty_vars[s[0]].get() or 1)
                except ValueError:
                    qty = 1
                line_total = s[2] * qty
                subtotal += line_total
                receipt_lines.append(f'  {s[1]:<25} x{qty:<3} = {line_total:>15,.0f}')

            try:
                discount = float(discount_var.get() or 0)
            except ValueError:
                discount = 0

            final_amount = max(subtotal - discount, 0)

            receipt_win = tk.Toplevel(win)
            receipt_win.title('🧾 Receipt Preview')
            receipt_win.geometry('500x600')
            receipt_win.configure(bg='white')
            receipt_win.resizable(False, False)

            receipt_win.update_idletasks()
            w, h = 500, 600
            x = (receipt_win.winfo_screenwidth() // 2) - (w // 2)
            y = (receipt_win.winfo_screenheight() // 2) - (h // 2)
            receipt_win.geometry(f'{w}x{h}+{x}+{y}')

            receipt_text = tk.Text(receipt_win, font=('Courier New', 10),
                                   bg='white', fg='#2c3e50', padx=20, pady=20,
                                   relief='flat', wrap='word')
            receipt_text.pack(fill='both', expand=True, padx=10, pady=10)

            from datetime import datetime
            receipt_text.insert('end', '=' * 55 + '\n')
            receipt_text.insert('end', '         🏥 HOSPITAL MANAGEMENT SYSTEM\n')
            receipt_text.insert('end', '                PAYMENT RECEIPT\n')
            receipt_text.insert('end', '=' * 55 + '\n\n')
            receipt_text.insert('end', f'📅 Date: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n')
            receipt_text.insert('end', f'👤 Patient: {patient_name}\n')
            receipt_text.insert('end', f'🆔 Patient ID: {patient_id}\n')
            receipt_text.insert('end', '\n' + '-' * 55 + '\n')
            receipt_text.insert('end', '📋 SERVICES:\n')
            receipt_text.insert('end', '-' * 55 + '\n')
            for line in receipt_lines:
                receipt_text.insert('end', line + '\n')
            receipt_text.insert('end', '\n' + '-' * 55 + '\n')
            receipt_text.insert('end', f'  {"Subtotal:":<25} {subtotal:>20,.0f} IRR\n')
            receipt_text.insert('end', f'  {"Discount:":<25} {"-" + f"{discount:,.0f}":>20} IRR\n')
            receipt_text.insert('end', '=' * 55 + '\n')
            receipt_text.insert('end', f'  {"TOTAL:":<25} {final_amount:>20,.0f} IRR\n')
            receipt_text.insert('end', '=' * 55 + '\n\n')
            receipt_text.insert('end', f'💳 Payment Method: {method_var.get()}\n')
            receipt_text.insert('end', f'📊 Status: {status_var.get()}\n')
            if notes_entry.get().strip():
                receipt_text.insert('end', f'\n📝 Notes: {notes_entry.get().strip()}\n')
            receipt_text.insert('end', '\n' + '=' * 55 + '\n')
            receipt_text.insert('end', '   Thank you for choosing our hospital!\n')
            receipt_text.insert('end', '=' * 55 + '\n')

            receipt_text.config(state='disabled')

            receipt_btn_frame = tk.Frame(receipt_win, bg='white')
            receipt_btn_frame.pack(fill='x', pady=10, padx=10)

            def save_as_txt():
                from tkinter import filedialog
                file_path = filedialog.asksaveasfilename(
                    defaultextension='.txt',
                    filetypes=[('Text files', '*.txt')],
                    initialfile=f'receipt_{datetime.now().strftime("%Y%m%d_%H%M%S")}.txt'
                )
                if file_path:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(receipt_text.get('1.0', 'end-1c'))
                    messagebox.showinfo('Success', f'✅ Receipt saved to:\n{file_path}')

            def print_receipt():
                # ===== ۱. ساخت PDF =====
                content = receipt_text.get('1.0', 'end-1c')
                success, result = print_text_file(content, 'receipt_print.txt', open_pdf=False)

                if not success:
                    messagebox.showerror('❌ PDF Error', f'Could not create PDF:\n{result}')
                    return

                pdf_path = result  # ← مسیر PDF

                # ===== ۲. اگه پرینتر موجود بود =====
                if WIN32_AVAILABLE:
                    printer = get_printer_for('receipt')  # ← تغییر
                    if printer:
                        ok, msg = print_pdf_to_printer(pdf_path, printer)
                        if ok:
                            messagebox.showinfo(
                                '🖨️ Sent to Printer',
                                f'Document sent to:\n\n🖨️ {printer}\n\n'
                                f'Check the printer for output.'
                            )
                        else:
                            # ===== اگه پرینت نشد، PDF رو باز کن =====
                            try:
                                os.startfile(pdf_path)
                            except:
                                pass

                            messagebox.showwarning(
                                '⚠️ Could Not Print',
                                f'Printer: {printer}\n\n'
                                f'The document could not be sent to the printer.\n'
                                f'This usually happens when:\n'
                                f'• The default printer is not connected\n'
                                f'• The printer is a virtual PDF printer\n'
                                f'• The printer is offline\n\n'
                                f'📄 PDF has been opened so you can print manually.'
                            )
                    else:
                        try:
                            os.startfile(pdf_path)
                        except:
                            pass
                        messagebox.showwarning(
                            '⚠️ No Default Printer',
                            'No default printer set in Windows.\n\n'
                            'PDF opened for manual printing.'
                        )
                else:
                    try:
                        os.startfile(pdf_path)
                    except:
                        pass
                    messagebox.showinfo(
                        '📄 PDF Opened',
                        'pywin32 not installed.\n\nPDF opened for manual printing.'
                    )

            tk.Button(receipt_btn_frame, text='💾 Save as TXT', command=save_as_txt,
                      bg='#3498db', fg='white', font=('Arial', 11, 'bold'),
                      padx=15, pady=8, cursor='hand2').pack(side='left', padx=3)

            tk.Button(receipt_btn_frame, text='🖨️ Print', command=print_receipt,
                      bg='#9b59b6', fg='white', font=('Arial', 11, 'bold'),
                      padx=15, pady=8, cursor='hand2').pack(side='left', padx=3)

            tk.Button(receipt_btn_frame, text='❌ Close', command=receipt_win.destroy,
                      bg='#e74c3c', fg='white', font=('Arial', 11, 'bold'),
                      padx=15, pady=8, cursor='hand2').pack(side='right', padx=3)

        # ===== دکمه‌های پایین =====
        btn_row_container = tk.Frame(btn_container, bg='#ecf0f1')
        btn_row_container.pack(side='right')

        create_rounded_button(
            btn_row_container,
            text='🧾 Show Receipt',
            command=lambda: show_receipt(),
            bg_color='#3498db',
            hover_color='#2980b9',
            width=180, height=50,
            corner_radius=20,
            font=('Arial', 12, 'bold')
        )

        create_rounded_button(
            btn_row_container,
            text='❌ Cancel',
            command=win.destroy,
            bg_color='#e74c3c',
            hover_color='#c0392b',
            width=140, height=50,
            corner_radius=20,
            font=('Arial', 12, 'bold')
        )

        create_rounded_button(
            btn_row_container,
            text='💾 SAVE PAYMENT',
            command=lambda: save(),
            bg_color='#27ae60',
            hover_color='#229954',
            width=200, height=50,
            corner_radius=20,
            font=('Arial', 12, 'bold')
        )

        win.bind('<Control-s>', save)

        # ===== محاسبه‌ی اولیه =====
        update_totals()

    def delete_payment(self):
        selected = self.payments_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', 'Please select at least one payment first!')
            return

        count = len(selected)

        if count == 1:
            confirm_msg = 'Delete this payment?'
        else:
            confirm_msg = f'Delete {count} payment(s)?\n\nThis action cannot be undone!'

        if not messagebox.askyesno('Confirm', confirm_msg):
            return

        try:
            conn = get_connection()
            cursor = conn.cursor()

            for item in selected:
                real_id = self.payments_tree.item(item)['tags'][0]
                # حذف سرویس‌های مرتبط
                cursor.execute('DELETE FROM payment_items WHERE payment_id = ?', (real_id,))
                # حذف پرداخت
                cursor.execute('DELETE FROM payments WHERE id = ?', (real_id,))

            conn.commit()
            conn.close()

            messagebox.showinfo('Success', f'✅ {count} payment(s) deleted!')
            self.load_payments()
            self.refresh_dashboard()

        except Exception as e:
            messagebox.showerror('Error', f'❌ Error: {e}')

    def show_daily_report(self):
        """گزارش روزانه پرداخت‌ها"""
        from datetime import datetime

        date_str = simpledialog.askstring(
            'Daily Report',
            'Enter date (YYYY-MM-DD) or leave empty for today:'
        )

        if date_str is None:
            return
        if not date_str.strip():
            date_str = datetime.now().strftime('%Y-%m-%d')

        try:
            datetime.strptime(date_str, '%Y-%m-%d')
        except ValueError:
            messagebox.showerror('Error', '❌ Invalid date! Use YYYY-MM-DD')
            return

        conn = get_connection()
        cursor = conn.cursor()

        # آمار
        cursor.execute('''
            SELECT
                COUNT(*) as total_count,
                SUM(CASE WHEN payment_status = 'Paid' THEN final_amount ELSE 0 END) as total_paid,
                SUM(CASE WHEN payment_status = 'Pending' THEN final_amount ELSE 0 END) as total_pending
            FROM payments
            WHERE DATE(payment_date) = ?
        ''', (date_str,))
        result = cursor.fetchone()

        # لیست پرداخت‌ها
        cursor.execute('''
            SELECT
                p.receipt_number,
                pt.first_name || ' ' || pt.last_name,
                p.final_amount,
                p.payment_method,
                p.payment_status
            FROM payments p
            JOIN patients pt ON p.patient_id = pt.id
            WHERE DATE(p.payment_date) = ?
            ORDER BY p.id
        ''', (date_str,))
        payments_list = cursor.fetchall()
        conn.close()

        # ===== پنجره‌ی گزارش =====
        win = tk.Toplevel(self.root)
        win.title(f'📊 Daily Report - {date_str}')
        win.geometry('800x600')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        win.update_idletasks()
        w, h = 800, 600
        x = (win.winfo_screenwidth() // 2) - (w // 2)
        y = (win.winfo_screenheight() // 2) - (h // 2)
        win.geometry(f'{w}x{h}+{x}+{y}')

        # هدر
        header = tk.Frame(win, bg='#2c3e50', height=55)
        header.pack(fill='x')
        header.pack_propagate(False)
        tk.Label(header, text=f'📊 Daily Report - {date_str}',
                 font=('Arial', 15, 'bold'), bg='#2c3e50',
                 fg='white').pack(pady=12)

        # آمار
        stats_frame = tk.Frame(win, bg='#ecf0f1')
        stats_frame.pack(fill='x', padx=20, pady=15)

        cards = tk.Frame(stats_frame, bg='#ecf0f1')
        cards.pack()

        def make_card(parent, icon, label, value, color):
            card = tk.Frame(parent, bg=color, width=200, height=90)
            card.pack(side='left', padx=8)
            card.pack_propagate(False)
            tk.Label(card, text=f'{icon} {label}', font=('Arial', 10),
                     bg=color, fg='white').pack(pady=(10, 2))
            tk.Label(card, text=str(value), font=('Arial', 16, 'bold'),
                     bg=color, fg='white').pack()

        make_card(cards, '📋', 'Total Payments', result[0], '#3498db')
        make_card(cards, '✅', 'Paid', f'{result[1] or 0:,.0f} IRR', '#27ae60')
        make_card(cards, '⏳', 'Pending', f'{result[2] or 0:,.0f} IRR', '#f39c12')

        # جدول پرداخت‌ها
        tk.Label(win, text='Payment Details:',
                 font=('Arial', 12, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(anchor='w', padx=20, pady=(15, 5))

        tree_frame = tk.Frame(win, bg='#ecf0f1')
        tree_frame.pack(fill='both', expand=True, padx=20, pady=(0, 10))

        columns = ('Receipt', 'Patient', 'Amount', 'Method', 'Status')
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=12)

        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=140, anchor='center')

        for p in payments_list:
            status_icon = '✅' if p[4] == 'Paid' else '⏳'
            tree.insert('', 'end', values=(
                p[0], p[1], f'{p[2]:,.0f}', p[3], f'{status_icon} {p[4]}'
            ))

        scroll = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        # ===== دکمه‌ها =====
        btn_frame = tk.Frame(win, bg='#ecf0f1')
        btn_frame.pack(pady=15)

        def export_daily_excel():
            if not payments_list:
                messagebox.showwarning('Warning', '📭 No data to export!')
                return

            from tkinter import filedialog
            file_path = filedialog.asksaveasfilename(
                defaultextension='.xlsx',
                filetypes=[('Excel files', '*.xlsx')],
                initialfile=f'daily_report_{date_str}.xlsx'
            )

            if not file_path:
                return

            headers = ['Receipt', 'Patient', 'Amount (IRR)', 'Method', 'Status']
            rows = []
            for p in payments_list:
                rows.append((p[0], p[1], p[2], p[3], p[4]))

            title = f'Daily Report - {date_str} | Total: {result[0]} payments | Paid: {result[1] or 0:,.0f} IRR'

            success, msg = export_to_excel(
                headers, rows, file_path,
                sheet_name='Daily Report',
                title=title
            )

            if success:
                messagebox.showinfo('✅ Export Successful',
                                    f'Excel saved to:\n{msg}')
                try:
                    os.startfile(msg)
                except:
                    pass
            else:
                messagebox.showerror('Export Error', f'❌ {msg}')

        create_popup_buttons(btn_frame, [
            {'text': '📊 Export Excel', 'command': export_daily_excel,
             'color': 'info', 'width': 180},
            {'text': '❌ Close', 'command': win.destroy,
             'color': 'danger', 'width': 120},
        ], bg='#ecf0f1')

    # ==========================================
    # Pharmacy Tab
    # ==========================================

    def create_pharmacy_tab(self):
        """تب مدیریت داروخانه (با کادر آمار و فیلتر تاریخ)"""
        tk.Label(self.pharmacy_tab, text='💊 Pharmacy Management',
                 font=('Arial', 14, 'bold'), bg=WINDOW_BG, fg='#2c3e50').pack(pady=10)

        # ===== بخش آمار و فیلتر =====
        stats_outer = tk.Frame(self.pharmacy_tab, bg=WINDOW_BG)
        stats_outer.pack(fill='x', padx=40, pady=(0, 10))

        # ردیف فیلتر تاریخ و دکمه اکسل
        filter_row = tk.Frame(stats_outer, bg=WINDOW_BG)
        filter_row.pack(fill='x', pady=(0, 5))

        tk.Label(filter_row, text='📅 Stats Date:', font=('Arial', 10, 'bold'),
                 bg=WINDOW_BG, fg='#2c3e50').pack(side='left', padx=(0, 5))

        # فیلتر تاریخ اختصاصی برای آمار داروخانه
        self.pharma_stat_filter_var, self.pharma_stat_from_var, self.pharma_stat_to_var, df_container, _ = create_date_filter(
            filter_row, WINDOW_BG, lambda: self.load_pharmacy_stats()
        )
        df_container.pack(side='left')

        # دکمه Export Excel
        create_standard_button(
            filter_row,
            text='📊 Export Excel',
            command=self.export_pharmacy_stats_excel,
            color='info',
            width=150,
            height=32
        )

        # ردیف کارت‌های آمار
        cards_frame = tk.Frame(stats_outer, bg=WINDOW_BG)
        cards_frame.pack(pady=5)  # ← fill='x' حذف شد

        self.pharma_stat_cards = {}
        CARD_COLOR = '#3b4a6b'

        card_configs = [
            ('total_revenue', '💰 Total Revenue'),
            ('direct_sales', '🛒 Direct Sales'),
            ('dispense_revenue', '📋 Dispensed Rx'),
            ('items_sold', '📦 Items Sold')
        ]

        for key, title in card_configs:
            canvas, updater = create_rounded_stat_card(cards_frame, title, '0', bg_color=CARD_COLOR, width=200,
                                                       height=85)
            canvas.pack(side='left', padx=5)  # ← fill='x', expand=True حذف شد
            self.pharma_stat_cards[key] = updater

        # ===== دکمه‌های اصلی داروخانه =====
        btn_bar1 = create_button_bar(self.pharmacy_tab, bg=WINDOW_BG)
        create_standard_button(btn_bar1, '➕ Add', self.add_medicine_window, color='success')
        create_standard_button(btn_bar1, '💰 Sell', self.sell_medicine_window, color='warning')
        create_standard_button(btn_bar1, '📦 Add Stock', self.add_stock_window, color='secondary')
        create_standard_button(btn_bar1, '👁️ Refresh', self.load_medicines, color='primary')

        btn_bar2 = create_button_bar(self.pharmacy_tab, bg=WINDOW_BG)
        create_standard_button(btn_bar2, '📊 Full Report', self.pharmacy_full_report_window, color='success') # ← دکمه جدید
        create_standard_button(btn_bar2, '📋 Dispense', self.dispense_medicines_window, color='success')
        create_standard_button(btn_bar2, '📋 History', self.medicine_history_window, color='secondary')
        create_standard_button(btn_bar2, '⚠️ Low Stock', self.low_stock_window, color='warning')
        create_standard_button(btn_bar2, '📊 Report', self.pharmacy_sales_report, color='info')
        create_standard_button(btn_bar2, '📊 Rx Report', self.prescription_report_window, color='primary')
        create_standard_button(btn_bar2, '🗑️ Delete', self.delete_medicine, color='danger')

        # ===== ردیف جستجو و فیلتر موجودی =====
        row1 = tk.Frame(self.pharmacy_tab, bg=WINDOW_BG)
        row1.pack(pady=(5, 3), padx=40, fill='x')

        tk.Label(row1, text='🔍 Search:', font=('Arial', 11, 'bold'),
                 bg=WINDOW_BG, fg='#2c3e50').pack(side='left', padx=(0, 8))

        self.medicine_search_var = tk.StringVar()
        search_entry = tk.Entry(row1, textvariable=self.medicine_search_var,
                                font=('Arial', 11), bg='white',
                                relief='solid', bd=1, width=35)
        search_entry.pack(side='left', ipady=6, padx=(0, 8))

        self.medicine_search_var.trace_add('write', lambda *args: self.load_medicines())

        def clear_search():
            self.medicine_search_var.set('')

        tk.Button(row1, text='✖ Clear', command=clear_search,
                  bg='#95a5a6', fg='white', font=('Arial', 9, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=10, pady=4).pack(side='left')

        row2 = tk.Frame(self.pharmacy_tab, bg=WINDOW_BG)
        row2.pack(pady=(0, 10), padx=40, fill='x')

        tk.Label(row2, text='📦 Stock:', font=('Arial', 10, 'bold'),
                 bg=WINDOW_BG, fg='#2c3e50').pack(side='left', padx=(0, 5))

        self.medicine_stock_filter = tk.StringVar(value='All')
        stock_combo = ttk.Combobox(row2, textvariable=self.medicine_stock_filter,
                                   values=['All', 'In Stock', 'Low Stock', 'Out of Stock'],
                                   state='readonly', width=13, font=('Arial', 10))
        stock_combo.pack(side='left')
        stock_combo.bind('<<ComboboxSelected>>', lambda e: self.load_medicines())

        self.medicine_count_var = tk.StringVar(value='')
        tk.Label(row2, textvariable=self.medicine_count_var,
                 font=('Arial', 10, 'italic'), bg=WINDOW_BG,
                 fg='#7f8c8d').pack(side='left', padx=(20, 0))

        self.medicine_alert_var = tk.StringVar(value='')
        tk.Label(row2, textvariable=self.medicine_alert_var,
                 font=('Arial', 10, 'bold'), bg=WINDOW_BG,
                 fg='#e67e22').pack(side='right')

        # ===== Treeview =====
        columns = ('ID', 'Name', 'Category', 'Price', 'Stock', 'Min', 'Expiry')
        tree_frame = tk.Frame(self.pharmacy_tab, bg=WINDOW_BG)
        tree_frame.pack(fill='both', expand=True, padx=40, pady=(5, 10))

        self.medicines_tree = ttk.Treeview(tree_frame, columns=columns,
                                           show='headings', height=12,
                                           selectmode='extended')

        for col in columns:
            self.medicines_tree.heading(col, text=col)
            self.medicines_tree.column(col, width=130, anchor='center')

        scroll = ttk.Scrollbar(tree_frame, orient='vertical', command=self.medicines_tree.yview)
        self.medicines_tree.configure(yscrollcommand=scroll.set)

        self.medicines_tree.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        enable_multi_click_selection(self.medicines_tree)

        # ===== بارگذاری اولیه =====
        self.load_medicines()
        self.load_pharmacy_stats()

    def load_pharmacy_stats(self):
        if not hasattr(self, 'pharma_stat_filter_var'): return

        from_date, to_date = get_date_range(
            self.pharma_stat_filter_var.get(),
            self.pharma_stat_from_var.get().strip(),
            self.pharma_stat_to_var.get().strip()
        )

        conn = get_connection()
        cursor = conn.cursor()

        date_cond = ""
        params = []
        if from_date:
            date_cond += " AND DATE(payment_date, 'localtime') >= ?"
            params.append(from_date)
        if to_date:
            date_cond += " AND DATE(payment_date, 'localtime') <= ?"
            params.append(to_date)

        # ✅ چاپ اطلاعات دیتابیس برای بررسی
        cursor.execute(f"SELECT receipt_number, final_amount FROM pharmacy_payments WHERE 1=1 {date_cond}", params)
        all_payments = cursor.fetchall()
        print("=========================================")
        print("DEBUG: All Pharmacy Payments:", all_payments)

        # 1. درآمد فروش آزاد (رسیدهای INV-)
        q_direct = f"SELECT COALESCE(SUM(final_amount), 0) FROM pharmacy_payments WHERE receipt_number LIKE 'INV-%' {date_cond}"
        cursor.execute(q_direct, params)
        direct_rev = cursor.fetchone()[0]

        # 2. درآمد نسخه‌ها (رسیدهای PH-)
        q_disp = f"SELECT COALESCE(SUM(final_amount), 0) FROM pharmacy_payments WHERE receipt_number LIKE 'PH-%' {date_cond}"
        cursor.execute(q_disp, params)
        dispense_rev = cursor.fetchone()[0]

        print("DEBUG: Direct Revenue:", direct_rev)
        print("DEBUG: Dispense Revenue:", dispense_rev)
        print("=========================================")

        # 3. تعداد اقلام
        date_cond_sales = ""
        params_sales = []
        if from_date:
            date_cond_sales += " AND DATE(sale_date, 'localtime') >= ?"
            params_sales.append(from_date)
        if to_date:
            date_cond_sales += " AND DATE(sale_date, 'localtime') <= ?"
            params_sales.append(to_date)

        cursor.execute(f"SELECT COALESCE(SUM(quantity), 0) FROM medicine_sales WHERE 1=1 {date_cond_sales}",
                       params_sales)
        items_sold = cursor.fetchone()[0]

        conn.close()

        total_rev = direct_rev + dispense_rev

        self.pharma_stat_cards['total_revenue'](f'{total_rev:,.0f}')
        self.pharma_stat_cards['direct_sales'](f'{direct_rev:,.0f}')
        self.pharma_stat_cards['dispense_revenue'](f'{dispense_rev:,.0f}')
        self.pharma_stat_cards['items_sold'](f'{items_sold:,.0f}')


    def export_pharmacy_stats_excel(self):
        """خروجی اکسل از آمار داروخانه"""
        from tkinter import filedialog
        from datetime import datetime

        from_date, to_date = get_date_range(
            self.pharma_stat_filter_var.get(),
            self.pharma_stat_from_var.get().strip(),
            self.pharma_stat_to_var.get().strip()
        )

        file_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[('Excel files', '*.xlsx')],
            initialfile=f'pharmacy_stats_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )

        if not file_path:
            return

        try:
            conn = get_connection()
            cursor = conn.cursor()

            date_cond_sales = ""
            params_sales = []
            if from_date:
                date_cond_sales += " AND DATE(sale_date) >= ?"
                params_sales.append(from_date)
            if to_date:
                date_cond_sales += " AND DATE(sale_date) <= ?"
                params_sales.append(to_date)

            # خواندن داده‌ها برای اکسل
            cursor.execute(f'''
                SELECT 
                    m.name, 
                    SUM(s.quantity) as qty, 
                    SUM(s.total_price) as total
                FROM medicine_sales s
                JOIN medicines m ON s.medicine_id = m.id
                WHERE 1=1 {date_cond_sales}
                GROUP BY m.id
                ORDER BY total DESC
            ''', params_sales)
            sales_data = cursor.fetchall()

            date_cond_payments = ""
            params_payments = []
            if from_date:
                date_cond_payments += " AND DATE(payment_date) >= ?"
                params_payments.append(from_date)
            if to_date:
                date_cond_payments += " AND DATE(payment_date) <= ?"
                params_payments.append(to_date)

            cursor.execute(f'''
                SELECT 
                    COUNT(id) as rx_count, 
                    SUM(final_amount) as total
                FROM pharmacy_payments
                WHERE 1=1 {date_cond_payments}
            ''', params_payments)
            dispense_summary = cursor.fetchone()

            conn.close()

            # ساخت فایل اکسل
            headers = ['Medicine / Item', 'Quantity Sold', 'Total Revenue (IRR)']
            rows = [(row[0], row[1], row[2]) for row in sales_data]

            # اضافه کردن خلاصه نسخه‌ها به انتهای جدول
            if dispense_summary[0] > 0:
                rows.append(('--- DISPENSED PRESCRIPTIONS ---', dispense_summary[0], dispense_summary[1] or 0))

            title = f"Pharmacy Sales Report ({from_date or 'All'} to {to_date or 'Now'})"

            success, msg = export_to_excel(
                headers, rows, file_path,
                sheet_name='Pharmacy Stats',
                title=title
            )

            if success:
                messagebox.showinfo('✅ Export Successful', f'Excel saved to:\n{msg}')
                try:
                    os.startfile(msg)
                except:
                    pass
            else:
                messagebox.showerror('Export Error', f'❌ {msg}')

        except Exception as e:
            messagebox.showerror('Export Error', f'❌ {e}')





    def export_pharmacy_stats_excel(self):
        """خروجی اکسل از آمار داروخانه"""
        from tkinter import filedialog
        from datetime import datetime

        from_date, to_date = get_date_range(
            self.pharma_stat_filter_var.get(),
            self.pharma_stat_from_var.get().strip(),
            self.pharma_stat_to_var.get().strip()
        )

        file_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[('Excel files', '*.xlsx')],
            initialfile=f'pharmacy_stats_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )

        if not file_path:
            return

        try:
            conn = get_connection()
            cursor = conn.cursor()

            date_cond_sales = ""
            params_sales = []
            if from_date:
                date_cond_sales += " AND DATE(sale_date) >= ?"
                params_sales.append(from_date)
            if to_date:
                date_cond_sales += " AND DATE(sale_date) <= ?"
                params_sales.append(to_date)

            # خواندن داده‌ها برای اکسل
            cursor.execute(f'''
                SELECT 
                    m.name, 
                    SUM(s.quantity) as qty, 
                    SUM(s.total_price) as total
                FROM medicine_sales s
                JOIN medicines m ON s.medicine_id = m.id
                WHERE 1=1 {date_cond_sales}
                GROUP BY m.id
                ORDER BY total DESC
            ''', params_sales)
            sales_data = cursor.fetchall()

            date_cond_payments = ""
            params_payments = []
            if from_date:
                date_cond_payments += " AND DATE(payment_date) >= ?"
                params_payments.append(from_date)
            if to_date:
                date_cond_payments += " AND DATE(payment_date) <= ?"
                params_payments.append(to_date)

            cursor.execute(f'''
                SELECT 
                    COUNT(id) as rx_count, 
                    SUM(final_amount) as total
                FROM pharmacy_payments
                WHERE 1=1 {date_cond_payments}
            ''', params_payments)
            dispense_summary = cursor.fetchone()

            conn.close()

            # ساخت فایل اکسل
            headers = ['Medicine / Item', 'Quantity Sold', 'Total Revenue (IRR)']
            rows = [(row[0], row[1], row[2]) for row in sales_data]

            # اضافه کردن خلاصه نسخه‌ها به انتهای جدول
            if dispense_summary[0] > 0:
                rows.append(('--- DISPENSED PRESCRIPTIONS ---', dispense_summary[0], dispense_summary[1] or 0))

            title = f"Pharmacy Sales Report ({from_date or 'All'} to {to_date or 'Now'})"

            success, msg = export_to_excel(
                headers, rows, file_path,
                sheet_name='Pharmacy Stats',
                title=title
            )

            if success:
                messagebox.showinfo('✅ Export Successful', f'Excel saved to:\n{msg}')
                try:
                    os.startfile(msg)
                except:
                    pass
            else:
                messagebox.showerror('Export Error', f'❌ {msg}')

        except Exception as e:
            messagebox.showerror('Export Error', f'❌ {e}')


    def add_medicine_window(self):
        """فرم افزودن داروی جدید"""
        win = tk.Toplevel(self.root)
        win.title('💊 Add Medicine')
        win.configure(bg='#f0f0f0')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#f0f0f0')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#f0f0f0')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f0f0f0')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text='💊 Add New Medicine',
                 font=('Arial', 16, 'bold'), bg='#f0f0f0',
                 fg='#2c3e50').pack(pady=15)

        # ===== ۴. فرم =====
        form_frame = tk.Frame(scrollable, bg='#f0f0f0')
        form_frame.pack(pady=10, padx=40, fill='x')

        fields = {}
        labels = [
            ('Name *', 'name', 'e.g., Aspirin 100mg'),
            ('Category *', 'category', 'e.g., Painkiller'),
            ('Manufacturer', 'manufacturer', 'e.g., Tehran Pharma'),
            ('Price (IRR) *', 'price', 'e.g., 25000'),
            ('Initial Stock *', 'stock', 'e.g., 100'),
            ('Min Stock *', 'min_stock', 'e.g., 20'),
            ('Expiry Date', 'expiry', 'YYYY-MM-DD'),
            ('Description', 'desc', 'Optional'),
        ]

        for i, (label, key, hint) in enumerate(labels):
            tk.Label(form_frame, text=label, bg='#f0f0f0',
                     font=('Arial', 10, 'bold'), anchor='w').grid(
                row=i * 2, column=0, sticky='w', pady=(5, 0), padx=5)
            entry = tk.Entry(form_frame, width=40, font=('Arial', 10))
            entry.grid(row=i * 2 + 1, column=0, pady=(0, 5), padx=5, sticky='ew')
            fields[key] = entry
            tk.Label(form_frame, text=f'💡 {hint}', bg='#f0f0f0',
                     font=('Arial', 8), fg='#7f8c8d', anchor='w').grid(
                row=i * 2 + 1, column=1, sticky='w', padx=5)

        form_frame.grid_columnconfigure(0, weight=1)

        # ===== ۵. تابع ذخیره =====
        def save():
            try:
                name = fields['name'].get().strip()
                if not name:
                    messagebox.showerror('Error', '❌ Name is required!')
                    return

                category = fields['category'].get().strip()
                if not category:
                    messagebox.showerror('Error', '❌ Category is required!')
                    return

                manufacturer = fields['manufacturer'].get().strip()

                try:
                    price = float(fields['price'].get().strip())
                    if price <= 0: raise ValueError
                except ValueError:
                    messagebox.showerror('Error', '❌ Price must be positive!')
                    return

                try:
                    stock = int(fields['stock'].get().strip())
                    if stock < 0: raise ValueError
                except ValueError:
                    messagebox.showerror('Error', '❌ Stock must be >= 0!')
                    return

                try:
                    min_stock = int(fields['min_stock'].get().strip())
                    if min_stock < 0: raise ValueError
                except ValueError:
                    messagebox.showerror('Error', '❌ Min stock must be >= 0!')
                    return

                expiry = fields['expiry'].get().strip()
                if expiry:
                    try:
                        from datetime import datetime
                        datetime.strptime(expiry, '%Y-%m-%d')
                    except ValueError:
                        messagebox.showerror('Error', '❌ Expiry date must be YYYY-MM-DD!')
                        return

                desc = fields['desc'].get().strip()

                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO medicines (name, category, manufacturer, price, stock_quantity, min_stock, expiry_date, description)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (name, category, manufacturer, price, stock, min_stock, expiry, desc))
                conn.commit()
                conn.close()

                self.load_pharmacy_stats()

                messagebox.showinfo('Success', f'✅ Medicine "{name}" added!')
                try:
                    log_action('CREATE', entity_type='medicine',
                               details=f'Added medicine: {name} (Stock: {stock})',
                               user=getattr(self.root, 'current_user', None))
                except:
                    pass
                win.destroy()
                self.load_medicines()

            except sqlite3.IntegrityError:
                messagebox.showerror('Error', '❌ Medicine name already exists!')
            except Exception as e:
                messagebox.showerror('Error', f'❌ Error: {e}')

        # ===== ۶. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '💾 Save', 'command': save, 'color': 'success', 'width': 160},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#f0f0f0')

        # ===== ۷. Maximize =====
        win.update_idletasks()
        win.state('zoomed')

        fields['name'].focus()

    def sell_medicine_window(self):
        """فرم فروش آزاد دارو"""
        win = tk.Toplevel(self.root)
        win.title('💰 Sell Medicine (Direct)')
        win.configure(bg='#f0f0f0')
        win.resizable(True, True)

        btn_container = tk.Frame(win, bg='#f0f0f0')
        btn_container.pack(side='bottom', fill='x', pady=15)

        body_container = tk.Frame(win, bg='#f0f0f0')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f0f0f0')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        tk.Label(scrollable, text='💰 Sell Medicine (Direct)',
                 font=('Arial', 16, 'bold'), bg='#f0f0f0', fg='#2c3e50').pack(pady=15)

        med_frame = tk.Frame(scrollable, bg='#f0f0f0')
        med_frame.pack(fill='x', pady=5, padx=40)

        tk.Label(med_frame, text='💊 Select Medicine:', font=('Arial', 11, 'bold'),
                 bg='#f0f0f0', width=15, anchor='w').pack(side='left')

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('SELECT id, name, price, stock_quantity FROM medicines WHERE stock_quantity > 0 ORDER BY name')
        meds = cursor.fetchall()
        conn.close()

        if not meds:
            messagebox.showerror('Error', '❌ No medicines in stock!')
            win.destroy()
            return

        med_options = [f'{m[0]} - {m[1]} ({m[2]:,.0f}) - Stock: {m[3]}' for m in meds]
        med_var = tk.StringVar(value=med_options[0])
        ttk.Combobox(med_frame, textvariable=med_var, values=med_options,
                     width=45, state='readonly').pack(side='left', padx=5)

        qty_var = tk.StringVar(value='1')
        tk.Spinbox(med_frame, from_=1, to=1000, textvariable=qty_var,
                   width=8, font=('Arial', 11)).pack(side='left', padx=5)

        tk.Label(scrollable, text='🛒 Cart:', font=('Arial', 12, 'bold'),
                 bg='#f0f0f0', fg='#2c3e50', anchor='w').pack(anchor='w', padx=40, pady=(15, 5))

        cart_container = tk.Frame(scrollable, bg='white', relief='solid', bd=1)
        cart_container.pack(fill='x', padx=40, pady=5)

        cart_cols = ('#', 'Medicine', 'Qty', 'Unit Price', 'Subtotal')
        cart_tree = ttk.Treeview(cart_container, columns=cart_cols, show='headings', height=6)
        for col in cart_cols:
            cart_tree.heading(col, text=col)
        cart_tree.column('#', width=40, anchor='center')
        cart_tree.column('Medicine', width=250, anchor='w')
        cart_tree.column('Qty', width=70, anchor='center')
        cart_tree.column('Unit Price', width=120, anchor='e')
        cart_tree.column('Subtotal', width=130, anchor='e')
        cart_tree.pack(fill='x', padx=5, pady=5)

        total_frame = tk.Frame(scrollable, bg='#f8f9fa', relief='solid', bd=2)
        total_frame.pack(fill='x', padx=40, pady=10)

        total_var = tk.StringVar(value='0 IRR')
        tk.Label(total_frame, text='💰 TOTAL:', font=('Arial', 14, 'bold'),
                 bg='#f8f9fa').pack(side='left', padx=15, pady=10)
        tk.Label(total_frame, textvariable=total_var, font=('Arial', 14, 'bold'),
                 bg='#f8f9fa', fg='#27ae60').pack(side='left')

        cart = []

        def update_cart_display():
            for item in cart_tree.get_children():
                cart_tree.delete(item)

            for idx, item in enumerate(cart, 1):
                cart_tree.insert('', 'end', values=(
                    idx, item['name'], item['qty'],
                    f'{item["price"]:,.0f}', f'{item["total"]:,.0f}'
                ))

            total = sum(item['total'] for item in cart)
            total_var.set(f'{total:,.0f} IRR')

        def add_to_cart():
            try:
                med_id = int(med_var.get().split(' - ')[0])
                qty = int(qty_var.get())
                if qty <= 0:
                    messagebox.showerror('Error', '❌ Quantity must be positive!')
                    return

                med = next((m for m in meds if m[0] == med_id), None)
                if not med:
                    messagebox.showerror('Error', '❌ Medicine not found!')
                    return

                already_in_cart = sum(item['qty'] for item in cart if item['id'] == med_id)
                if qty + already_in_cart > med[3]:
                    messagebox.showerror('Error', f'❌ Only {med[3] - already_in_cart} available!')
                    return

                cart.append({
                    'id': med_id, 'name': med[1], 'qty': qty,
                    'price': med[2], 'total': med[2] * qty
                })
                update_cart_display()
                qty_var.set('1')
            except ValueError:
                messagebox.showerror('Error', '❌ Invalid input!')

        def remove_selected():
            selected = cart_tree.selection()
            if not selected:
                messagebox.showwarning('Warning', '⚠️ Select an item first!')
                return
            idx = int(cart_tree.item(selected[0])['values'][0]) - 1
            if 0 <= idx < len(cart):
                cart.pop(idx)
                update_cart_display()

        def confirm_sale():
            if not cart:
                messagebox.showerror('Error', '❌ Cart is empty!')
                return

            total = sum(item['total'] for item in cart)
            if not messagebox.askyesno('Confirm', f'Sell {len(cart)} item(s)?\nTotal: {total:,.0f} IRR'):
                return

            try:
                import random
                from datetime import datetime

                # ✅ تضمین استفاده از پیشوند INV-
                receipt = f'INV-{datetime.now().strftime("%Y%m%d%H%M%S")}-{random.randint(100, 999)}'

                conn = get_connection()
                cursor = conn.cursor()

                cursor.execute("SELECT id FROM patients LIMIT 1")
                row = cursor.fetchone()
                default_patient_id = row[0] if row else None

                cursor.execute('''
                    INSERT INTO pharmacy_payments
                    (receipt_number, patient_id, amount, discount, final_amount,
                     payment_method, payment_status, created_by)
                    VALUES (?, ?, ?, ?, ?, ?, 'Paid', ?)
                ''', (receipt, default_patient_id, total, 0, total, 'Cash',
                      getattr(self.root, 'current_user', {}).get('username', 'admin')))

                for item in cart:
                    cursor.execute('''
                        INSERT INTO medicine_sales
                        (medicine_id, patient_id, quantity, unit_price, total_price, invoice_id)
                        VALUES (?, ?, ?, ?, ?, ?)
                    ''', (item['id'], default_patient_id, item['qty'],
                          item['price'], item['total'], receipt))

                    cursor.execute('UPDATE medicines SET stock_quantity = stock_quantity - ? WHERE id = ?',
                                   (item['qty'], item['id']))

                conn.commit()
                conn.close()

                self.load_medicines()
                self.load_pharmacy_stats()
                win.destroy()

                messagebox.showinfo('✅ Success', f'✅ Sale complete!\n\nReceipt: {receipt}\nTotal: {total:,.0f} IRR')

            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        create_popup_buttons(btn_container, [
            {'text': '➕ Add to Cart', 'command': add_to_cart, 'color': 'primary', 'width': 160},
            {'text': '🗑️ Remove Selected', 'command': remove_selected, 'color': 'warning', 'width': 180},
            {'text': '✅ Confirm Sale', 'command': confirm_sale, 'color': 'success', 'width': 180},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 120},
        ], bg='#f0f0f0')

        win.update_idletasks()
        win.state('zoomed')
        update_cart_display()


    def add_stock_window(self):
        """فرم افزودن موجودی دارو"""
        selected = self.medicines_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select a medicine first!')
            return

        real_id = self.medicines_tree.item(selected[0])['tags'][0]
        values = self.medicines_tree.item(selected[0])['values']
        med_name = values[1]

        win = tk.Toplevel(self.root)
        win.title('📦 Add Stock')
        win.configure(bg='#f0f0f0')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#f0f0f0')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#f0f0f0')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f0f0f0')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text='📦 Add Stock',
                 font=('Arial', 16, 'bold'), bg='#f0f0f0',
                 fg='#2c3e50').pack(pady=15)

        # ===== ۴. اطلاعات دارو =====
        info_frame = tk.Frame(scrollable, bg='#f8f9fa', relief='solid', bd=1)
        info_frame.pack(fill='x', padx=40, pady=10)

        tk.Label(info_frame, text=f'💊 Medicine: {med_name}',
                 font=('Arial', 12, 'bold'), bg='#f8f9fa',
                 fg='#2c3e50', anchor='w').pack(anchor='w', padx=15, pady=(10, 5))

        tk.Label(info_frame, text=f'📊 Current Stock: {values[4]}',
                 font=('Arial', 11), bg='#f8f9fa',
                 fg='#7f8c8d', anchor='w').pack(anchor='w', padx=15, pady=(3, 10))

        # ===== ۵. فرم =====
        form_frame = tk.Frame(scrollable, bg='#f0f0f0')
        form_frame.pack(pady=10, padx=40, fill='x')

        tk.Label(form_frame, text='📦 Quantity to add:',
                 font=('Arial', 11, 'bold'), bg='#f0f0f0',
                 anchor='w').pack(anchor='w', pady=(0, 5))

        qty_var = tk.StringVar(value='50')
        qty_entry = tk.Entry(form_frame, textvariable=qty_var,
                             font=('Arial', 12), width=20,
                             justify='center')
        qty_entry.pack(fill='x', ipady=6, pady=(0, 10))

        tk.Label(form_frame, text='💡 Enter a positive number',
                 font=('Arial', 8, 'italic'), bg='#f0f0f0',
                 fg='#7f8c8d', anchor='w').pack(anchor='w')

        # ===== ۶. تابع ذخیره =====
        def do_add():
            try:
                qty = int(qty_var.get())
                if qty <= 0:
                    messagebox.showerror('Error', '❌ Quantity must be positive!')
                    return

                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute('''
                    UPDATE medicines SET stock_quantity = stock_quantity + ?
                    WHERE id = ?
                ''', (qty, real_id))
                conn.commit()
                conn.close()

                messagebox.showinfo('Success', f'✅ Added {qty} units!')

                try:
                    log_action('UPDATE', entity_type='medicine',
                               entity_id=real_id,
                               details=f'Added {qty} units to "{med_name}" stock',
                               user=getattr(self.root, 'current_user', None))
                except:
                    pass

                win.destroy()
                self.load_medicines()

            except ValueError:
                messagebox.showerror('Error', '❌ Please enter a number!')
            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        # ===== ۷. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '📦 Add Stock', 'command': do_add, 'color': 'success', 'width': 160},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#f0f0f0')

        # ===== ۸. Maximize =====
        win.update_idletasks()
        win.state('zoomed')

        qty_entry.focus()


    def low_stock_window(self):
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT name, stock_quantity, min_stock
            FROM medicines
            WHERE stock_quantity < min_stock
            ORDER BY stock_quantity
        ''')
        low = cursor.fetchall()
        conn.close()

        if not low:
            messagebox.showinfo('✅ All Good', 'All medicines have sufficient stock!')
            return

        win = tk.Toplevel(self.root)
        win.title('⚠️ Low Stock Alert')
        win.geometry('500x400')
        win.configure(bg='#ecf0f1')

        win.update_idletasks()
        w, h = 500, 400
        x = (win.winfo_screenwidth() // 2) - (w // 2)
        y = (win.winfo_screenheight() // 2) - (h // 2)
        win.geometry(f'{w}x{h}+{x}+{y}')

        tk.Label(win, text='⚠️ Low Stock Alert', font=('Arial', 14, 'bold'),
                 bg='#ecf0f1', fg='#e67e22').pack(pady=10)

        text = tk.Text(win, font=('Courier New', 11), bg='white',
                       padx=15, pady=15, wrap='word')
        text.pack(fill='both', expand=True, padx=15, pady=10)

        for name, stock, min_stock in low:
            if stock < 0:
                text.insert('end', f'❌ {name}\n')
                text.insert('end', f'   Stock: {stock} (منفی!)\n')
            elif stock == 0:
                text.insert('end', f'🔴 {name}\n')
                text.insert('end', f'   Stock: 0 | Min: {min_stock}\n')
            else:
                text.insert('end', f'⚠️  {name}\n')
                text.insert('end', f'   Stock: {stock} | Min: {min_stock}\n')
            text.insert('end', '-' * 40 + '\n')

        text.config(state='disabled')

        tk.Button(win, text='❌ Close', command=win.destroy,
                  bg='#e74c3c', fg='white', font=('Arial', 11, 'bold'),
                  padx=25, pady=8, cursor='hand2').pack(pady=10)

    def pharmacy_sales_report(self):
        from datetime import datetime
        date_str = simpledialog.askstring('Sales Report',
            'Date (YYYY-MM-DD) or leave empty for today:')

        if date_str is None:
            return
        if not date_str.strip():
            date_str = datetime.now().strftime('%Y-%m-%d')

        try:
            datetime.strptime(date_str, '%Y-%m-%d')
        except ValueError:
            messagebox.showerror('Error', '❌ Invalid date! Use YYYY-MM-DD')
            return

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT
                COUNT(*) as total_sales,
                SUM(quantity) as total_items,
                SUM(total_price) as total_revenue
            FROM medicine_sales
            WHERE DATE(sale_date) = ?
        ''', (date_str,))
        result = cursor.fetchone()

        cursor.execute('''
            SELECT m.name, SUM(s.quantity) as qty, SUM(s.total_price) as total
            FROM medicine_sales s
            JOIN medicines m ON s.medicine_id = m.id
            WHERE DATE(s.sale_date) = ?
            GROUP BY m.id, m.name
            ORDER BY total DESC
            LIMIT 10
        ''', (date_str,))
        top_sales = cursor.fetchall()
        conn.close()

        msg = f'📊 Sales Report - {date_str}\n\n'
        msg += f'Total Sales: {result[0]}\n'
        msg += f'Total Items: {result[1] or 0}\n'
        msg += f'Total Revenue: {result[2] or 0:,.0f} IRR\n'

        if top_sales:
            msg += '\n🔥 Top Sales:\n'
            for name, qty, total in top_sales:
                msg += f'  {name}: {qty} units ({total:,.0f} IRR)\n'

        messagebox.showinfo('📊 Sales Report', msg)

    def prescription_report_window(self):
        """گزارش نسخه‌ها (Top Medicines + Summary + Daily)"""
        win = tk.Toplevel(self.root)
        win.title('📊 Prescription Report')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#ecf0f1')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه =====
        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        # ===== ۳. عنوان =====
        tk.Label(body_container, text='📊 Prescription Report',
                 font=('Arial', 16, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(pady=10)

        # ===== ۴. ردیف فیلتر =====
        filter_row = tk.Frame(body_container, bg='#ecf0f1')
        filter_row.pack(fill='x', padx=20, pady=5)

        tk.Label(filter_row, text='📅 Date Range:',
                 font=('Arial', 11, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(side='left', padx=(0, 8))

        filter_var, from_var, to_var, df_container, _ = create_date_filter(
            filter_row, '#ecf0f1', lambda: load_report()
        )
        df_container.pack(side='left')

        # ===== ۵. Notebook =====
        style = ttk.Style()
        style.theme_use('clam')
        notebook = ttk.Notebook(body_container)
        notebook.pack(fill='both', expand=True, pady=10, padx=20)

        # ===== تب ۱: Top Medicines =====
        tab_top = tk.Frame(notebook, bg='#ecf0f1')
        notebook.add(tab_top, text='🏆 Top Medicines')

        top_info = tk.Label(tab_top, text='',
                            font=('Arial', 11, 'bold'),
                            bg='#ecf0f1', fg='#27ae60')
        top_info.pack(pady=5)

        cols = ('Rank', 'Medicine', 'Category', 'Times Prescribed', 'Total Quantity')
        tree_frame = tk.Frame(tab_top, bg='#ecf0f1')
        tree_frame.pack(fill='both', expand=True, padx=10, pady=(0, 10))

        top_tree = ttk.Treeview(tree_frame, columns=cols, show='headings', height=15)

        widths = {'Rank': 70, 'Medicine': 280, 'Category': 180,
                  'Times Prescribed': 150, 'Total Quantity': 150}
        for c in cols:
            top_tree.heading(c, text=c)
            top_tree.column(c, width=widths[c], anchor='center', minwidth=60)

        top_scroll = ttk.Scrollbar(tree_frame, orient='vertical', command=top_tree.yview)
        top_tree.configure(yscrollcommand=top_scroll.set)

        top_tree.pack(side='left', fill='both', expand=True)
        top_scroll.pack(side='right', fill='y')

        # ===== تب ۲: Summary =====
        tab_summary = tk.Frame(notebook, bg='#ecf0f1')
        notebook.add(tab_summary, text='📊 Summary')

        # Scrollable برای Summary
        sum_canvas = tk.Canvas(tab_summary, bg='#ecf0f1', highlightthickness=0)
        sum_scrollbar = ttk.Scrollbar(tab_summary, orient='vertical', command=sum_canvas.yview)
        sum_scrollable = tk.Frame(sum_canvas, bg='#ecf0f1')

        sum_scrollable.bind('<Configure>', lambda e: sum_canvas.configure(scrollregion=sum_canvas.bbox('all')))
        sum_canvas_window = sum_canvas.create_window((0, 0), window=sum_scrollable, anchor='nw')
        sum_canvas.configure(yscrollcommand=sum_scrollbar.set)
        sum_canvas.bind('<Configure>', lambda e: sum_canvas.itemconfig(sum_canvas_window, width=e.width))

        sum_canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(sum_canvas, sum_scrollbar, sum_scrollable)

        def on_mousewheel(event):
            sum_canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        sum_canvas.bind_all('<MouseWheel>', on_mousewheel)

        cards_frame = tk.Frame(sum_scrollable, bg='#ecf0f1')
        cards_frame.pack(pady=30)

        def create_stat_card(parent, icon, title, value, color):
            card = tk.Frame(parent, bg=color, width=200, height=130)
            card.pack(side='left', padx=8)
            card.pack_propagate(False)

            tk.Label(card, text=icon, font=('Arial', 28),
                     bg=color, fg='white').pack(pady=(10, 0))
            tk.Label(card, text=title, font=('Arial', 10, 'bold'),
                     bg=color, fg='white').pack()
            value_label = tk.Label(card, text=value, font=('Arial', 20, 'bold'),
                                   bg=color, fg='white')
            value_label.pack(pady=(5, 10))

            return value_label

        total_rx_label = create_stat_card(cards_frame, '📋', 'Total Prescriptions', '0', '#3498db')
        unique_meds_label = create_stat_card(cards_frame, '💊', 'Unique Medicines', '0', '#9b59b6')
        unique_patients_label = create_stat_card(cards_frame, '👥', 'Unique Patients', '0', '#e67e22')
        unique_doctors_label = create_stat_card(cards_frame, '👨‍⚕️', 'Active Doctors', '0', '#27ae60')

        tk.Label(sum_scrollable, text='📅 Daily Prescription Count:',
                 font=('Arial', 12, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(pady=(20, 10))

        daily_cols = ('Date', 'Prescriptions', 'Unique Patients', 'Medicines')
        daily_frame = tk.Frame(sum_scrollable, bg='#ecf0f1')
        daily_frame.pack(fill='both', expand=True, padx=30, pady=(0, 20))

        daily_tree = ttk.Treeview(daily_frame, columns=daily_cols, show='headings', height=8)

        daily_widths = {'Date': 150, 'Prescriptions': 150,
                        'Unique Patients': 150, 'Medicines': 150}
        for c in daily_cols:
            daily_tree.heading(c, text=c)
            daily_tree.column(c, width=daily_widths[c], anchor='center', minwidth=60)

        daily_scroll = ttk.Scrollbar(daily_frame, orient='vertical', command=daily_tree.yview)
        daily_tree.configure(yscrollcommand=daily_scroll.set)

        daily_tree.pack(side='left', fill='both', expand=True)
        daily_scroll.pack(side='right', fill='y')

        # ===== ۶. گزارش =====
        report_data = {
            'top_medicines': [],
            'summary': {},
            'daily': [],
            'date_range': ('', ''),
        }

        def load_report():
            from_date, to_date = get_date_range(
                filter_var.get(),
                from_var.get().strip(),
                to_var.get().strip()
            )

            report_data['date_range'] = (from_date or 'All', to_date or 'Now')

            conn = get_connection()
            cursor = conn.cursor()

            date_where = ''
            date_params = []

            if from_date and to_date:
                date_where = ' WHERE DATE(mr.record_date) BETWEEN ? AND ?'
                date_params = [from_date, to_date]
            elif from_date:
                date_where = ' WHERE DATE(mr.record_date) >= ?'
                date_params = [from_date]
            elif to_date:
                date_where = ' WHERE DATE(mr.record_date) <= ?'
                date_params = [to_date]

            cursor.execute(f'''
                SELECT m.name, m.category,
                       COUNT(pr.id) AS times_prescribed,
                       SUM(pr.quantity) AS total_qty
                FROM prescriptions pr
                JOIN medicines m ON pr.medicine_id = m.id
                JOIN medical_records mr ON pr.record_id = mr.id
                {date_where}
                GROUP BY m.id, m.name, m.category
                ORDER BY times_prescribed DESC
                LIMIT 20
            ''', date_params)
            top_medicines = cursor.fetchall()

            cursor.execute(f'''
                SELECT
                    COUNT(DISTINCT mr.id) AS total_records,
                    COUNT(DISTINCT mr.patient_id) AS unique_patients,
                    COUNT(DISTINCT mr.doctor_id) AS unique_doctors,
                    COUNT(pr.id) AS total_prescriptions,
                    COUNT(DISTINCT pr.medicine_id) AS unique_medicines
                FROM medical_records mr
                LEFT JOIN prescriptions pr ON mr.id = pr.record_id
                {date_where}
            ''', date_params)
            summary = cursor.fetchone()

            cursor.execute(f'''
                SELECT mr.record_date,
                       COUNT(DISTINCT mr.id) AS record_count,
                       COUNT(DISTINCT mr.patient_id) AS patient_count,
                       COUNT(pr.id) AS rx_count
                FROM medical_records mr
                LEFT JOIN prescriptions pr ON mr.id = pr.record_id
                {date_where}
                GROUP BY mr.record_date
                ORDER BY mr.record_date DESC
                LIMIT 30
            ''', date_params)
            daily = cursor.fetchall()

            conn.close()

            report_data['top_medicines'] = top_medicines
            report_data['summary'] = {
                'total_records': summary[0] or 0,
                'unique_patients': summary[1] or 0,
                'unique_doctors': summary[2] or 0,
                'total_prescriptions': summary[3] or 0,
                'unique_medicines': summary[4] or 0,
            }
            report_data['daily'] = daily

            for item in top_tree.get_children():
                top_tree.delete(item)

            for i, (name, category, times, qty) in enumerate(top_medicines, 1):
                icon = '🥇' if i == 1 else '🥈' if i == 2 else '🥉' if i == 3 else f'{i}.'
                top_tree.insert('', 'end', values=(
                    icon, name, category or '—',
                    f'💊 {times}',
                    f'📦 {qty or 0}'
                ))

            total_rx_label.config(text=str(report_data['summary']['total_prescriptions']))
            unique_meds_label.config(text=str(report_data['summary']['unique_medicines']))
            unique_patients_label.config(text=str(report_data['summary']['unique_patients']))
            unique_doctors_label.config(text=str(report_data['summary']['unique_doctors']))

            for item in daily_tree.get_children():
                daily_tree.delete(item)

            for d in daily:
                daily_tree.insert('', 'end', values=(
                    d[0], f'{d[1]}', f'{d[2]}', f'{d[3]}'
                ))

            range_text = ''
            if from_date or to_date:
                range_text = f'  ({from_date or "..."} → {to_date or "now"})'
            top_info.config(
                text=f'🏆 Top {len(top_medicines)} medicines  |  '
                     f'📋 {report_data["summary"]["total_prescriptions"]} total prescriptions{range_text}'
            )

        # ===== ۷. Export Excel =====
        def export_excel():
            from tkinter import filedialog

            if not report_data['top_medicines']:
                messagebox.showwarning('Warning', '📭 No data to export!')
                return

            file_path = filedialog.asksaveasfilename(
                defaultextension='.xlsx',
                filetypes=[('Excel files', '*.xlsx')],
                initialfile=f'prescription_report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
            )

            if not file_path:
                return

            try:
                headers = ['Rank', 'Medicine', 'Category', 'Times Prescribed', 'Total Quantity']
                rows = []
                for i, (name, cat, times, qty) in enumerate(report_data['top_medicines'], 1):
                    rows.append((i, name, cat or '', times, qty or 0))

                title = f'Prescription Report ({report_data["date_range"][0]} → {report_data["date_range"][1]})'

                success, result = export_to_excel(
                    headers, rows, file_path,
                    sheet_name='Top Medicines',
                    title=title
                )

                if success:
                    messagebox.showinfo('✅ Export Successful',
                                        f'Excel file saved to:\n{result}')
                    try:
                        os.startfile(result)
                    except:
                        pass
                else:
                    messagebox.showerror('Export Error', f'❌ {result}')

            except Exception as e:
                messagebox.showerror('Export Error', f'❌ {e}')

        # ===== ۸. چاپ =====
        def print_report():
            lines = []
            lines.append('=' * 70)
            lines.append(' ' * 15 + 'HOSPITAL MANAGEMENT SYSTEM')
            lines.append(' ' * 20 + 'PRESCRIPTION REPORT')
            lines.append('=' * 70)
            lines.append('')
            lines.append(f'  Date Range: {report_data["date_range"][0]} → {report_data["date_range"][1]}')
            lines.append(f'  Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
            lines.append('')

            lines.append('-' * 70)
            lines.append('  SUMMARY')
            lines.append('-' * 70)
            s = report_data['summary']
            lines.append(f'  Total Prescriptions: {s["total_prescriptions"]}')
            lines.append(f'  Unique Medicines: {s["unique_medicines"]}')
            lines.append(f'  Unique Patients: {s["unique_patients"]}')
            lines.append(f'  Active Doctors: {s["unique_doctors"]}')
            lines.append('')

            lines.append('-' * 70)
            lines.append('  TOP PRESCRIBED MEDICINES')
            lines.append('-' * 70)
            lines.append(f'  {"#":<5} {"Medicine":<30} {"Times":<10} {"Qty":<10}')
            lines.append('-' * 70)
            for i, (name, cat, times, qty) in enumerate(report_data['top_medicines'], 1):
                lines.append(f'  {i:<5} {name[:28]:<30} {times:<10} {qty or 0:<10}')
            lines.append('')

            if report_data['daily']:
                lines.append('-' * 70)
                lines.append('  DAILY STATS')
                lines.append('-' * 70)
                for d in report_data['daily'][:15]:
                    lines.append(f'  {d[0]}  →  {d[1]} records, {d[2]} patients, {d[3]} prescriptions')
                lines.append('')

            lines.append('=' * 70)
            content = '\n'.join(lines)

            success, pdf_path = print_text_file(content, 'prescription_report.txt', open_pdf=False)
            if not success:
                messagebox.showerror('❌ PDF Error', f'Could not create PDF:\n{pdf_path}')
                return

            if WIN32_AVAILABLE:
                printer = get_printer_for('report')
                virtual_keywords = ['Print to PDF', 'XPS', 'OneNote', 'Fax', 'Virtual', 'Writer']
                is_virtual = any(kw.lower() in (printer or '').lower() for kw in virtual_keywords)

                if is_virtual or not printer:
                    try:
                        os.startfile(pdf_path)
                    except:
                        pass
                    messagebox.showinfo('📄 PDF Ready',
                                        f'PDF created.\n\nFile: {pdf_path}')
                else:
                    ok, msg = print_pdf_to_printer(pdf_path, printer)
                    if ok:
                        messagebox.showinfo('🖨️ Sent to Printer',
                                            f'Report sent to:\n\n🖨️ {printer}')
                    else:
                        try:
                            os.startfile(pdf_path)
                        except:
                            pass

        # ===== ۹. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '🔄 Refresh', 'command': load_report, 'color': 'primary', 'width': 140},
            {'text': '📊 Export Excel', 'command': export_excel, 'color': 'info', 'width': 160},
            {'text': '🖨️ Print', 'command': print_report, 'color': 'secondary', 'width': 140},
            {'text': '❌ Close', 'command': win.destroy, 'color': 'danger', 'width': 120},
        ], bg='#ecf0f1')

        # ===== ۱۰. Maximize =====
        win.update_idletasks()
        win.state('zoomed')

        load_report()


    def medicine_history_window(self):
        """تاریخچه خرید دارو با فیلتر تاریخ"""
        win = tk.Toplevel(self.root)
        win.title('📋 Medicine Purchase History')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#ecf0f1')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه =====
        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        # ===== ۳. عنوان =====
        tk.Label(body_container, text='📋 Medicine Purchase History',
                 font=('Arial', 16, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(pady=10)

        # ===== ۴. جستجو =====
        search_frame = tk.Frame(body_container, bg='#ecf0f1')
        search_frame.pack(fill='x', padx=20, pady=5)

        tk.Label(search_frame, text='🔍 National ID:',
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(side='left', padx=(0, 5))

        nid_var = tk.StringVar()
        nid_entry = tk.Entry(search_frame, textvariable=nid_var,
                             font=('Arial', 12), width=15,
                             bg='white', relief='solid', bd=1)
        nid_entry.pack(side='left', padx=(0, 15), ipady=4)

        # Date Filter
        filter_var, from_var, to_var, df_container, _ = create_date_filter(
            search_frame, '#ecf0f1', lambda: apply_filters()
        )
        df_container.pack(side='left')

        # ===== ۵. Notebook (3 تب) =====
        style = ttk.Style()
        style.theme_use('clam')
        notebook = ttk.Notebook(body_container)
        notebook.pack(fill='both', expand=True, padx=20, pady=10)

        # ===== تب ۱: Invoices =====
        tab_inv = tk.Frame(notebook, bg='#ecf0f1')
        notebook.add(tab_inv, text='📄 Invoices (Grouped)')

        inv_cols = ('Invoice #', 'Date', 'Time', 'Items', 'Total', 'Patient')
        inv_tree_frame = tk.Frame(tab_inv, bg='#ecf0f1')
        inv_tree_frame.pack(fill='both', expand=True, padx=10, pady=10)

        inv_tree = ttk.Treeview(inv_tree_frame, columns=inv_cols,
                                show='headings', height=13)

        inv_widths = {'Invoice #': 240, 'Date': 110, 'Time': 90,
                      'Items': 80, 'Total': 130, 'Patient': 220}
        for col in inv_cols:
            inv_tree.heading(col, text=col)
            inv_tree.column(col, width=inv_widths[col], anchor='center', minwidth=60)

        inv_scroll = ttk.Scrollbar(inv_tree_frame, orient='vertical', command=inv_tree.yview)
        inv_tree.configure(yscrollcommand=inv_scroll.set)
        inv_tree.pack(side='left', fill='both', expand=True)
        inv_scroll.pack(side='right', fill='y')

        # ===== تب ۲: All Items =====
        tab_items = tk.Frame(notebook, bg='#ecf0f1')
        notebook.add(tab_items, text='💊 All Items')

        item_cols = ('Invoice #', 'Date', 'Time', 'Medicine', 'Qty',
                     'Unit Price', 'Total', 'Patient')
        item_tree_frame = tk.Frame(tab_items, bg='#ecf0f1')
        item_tree_frame.pack(fill='both', expand=True, padx=10, pady=10)

        item_tree = ttk.Treeview(item_tree_frame, columns=item_cols,
                                 show='headings', height=13)

        item_widths = {'Invoice #': 210, 'Date': 100, 'Time': 80,
                       'Medicine': 200, 'Qty': 60, 'Unit Price': 100,
                       'Total': 110, 'Patient': 180}
        for col in item_cols:
            item_tree.heading(col, text=col)
            item_tree.column(col, width=item_widths[col], anchor='center', minwidth=60)

        item_scroll = ttk.Scrollbar(item_tree_frame, orient='vertical', command=item_tree.yview)
        item_tree.configure(yscrollcommand=item_scroll.set)
        item_tree.pack(side='left', fill='both', expand=True)
        item_scroll.pack(side='right', fill='y')

        # ===== تب ۳: Payments =====
        tab_payments = tk.Frame(notebook, bg='#ecf0f1')
        notebook.add(tab_payments, text='💳 Payments')

        pay_cols = ('ID', 'Receipt', 'Patient', 'Amount', 'Discount', 'Final', 'Method', 'Date')
        pay_tree_frame = tk.Frame(tab_payments, bg='#ecf0f1')
        pay_tree_frame.pack(fill='both', expand=True, padx=10, pady=10)

        pay_tree = ttk.Treeview(pay_tree_frame, columns=pay_cols,
                                show='headings', height=13)

        pay_widths = {'ID': 50, 'Receipt': 180, 'Patient': 180,
                      'Amount': 100, 'Discount': 80, 'Final': 100,
                      'Method': 80, 'Date': 150}
        for col in pay_cols:
            pay_tree.heading(col, text=col)
            pay_tree.column(col, width=pay_widths[col], anchor='center', minwidth=60)

        pay_scroll = ttk.Scrollbar(pay_tree_frame, orient='vertical', command=pay_tree.yview)
        pay_tree.configure(yscrollcommand=pay_scroll.set)
        pay_tree.pack(side='left', fill='both', expand=True)
        pay_scroll.pack(side='right', fill='y')

        # ===== ۶. جمع کل =====
        total_label = tk.Label(body_container, text='',
                               font=('Arial', 12, 'bold'),
                               bg='#ecf0f1', fg='#27ae60')
        total_label.pack(pady=5)

        status_var = tk.StringVar(value='')
        status_label = tk.Label(body_container, textvariable=status_var,
                                font=('Arial', 10, 'italic'),
                                bg='#ecf0f1', fg='#27ae60')
        status_label.pack(pady=2)

        # ===== ۷. State =====
        state = {
            'patient_id': None,
            'patient_name': '',
            'all_sales': [],
            'invoices': {},
        }

        def show_status(msg, color='#27ae60', duration=3000):
            status_var.set(msg)
            status_label.config(fg=color)
            win.after(duration, lambda: status_var.set(''))

        def search_history():
            nid = nid_var.get().strip()

            if not nid:
                show_status('⚠️ Please enter a National ID!', '#e67e22')
                return

            if not nid.isdigit() or len(nid) != 10:
                show_status('⚠️ National ID must be 10 digits!', '#e67e22')
                return

            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT id, first_name, last_name FROM patients WHERE national_id = ?', (nid,))
            patient = cursor.fetchone()

            if not patient:
                conn.close()
                show_status(f'❌ No patient found: {nid}', '#e74c3c', 4000)
                for item in inv_tree.get_children():
                    inv_tree.delete(item)
                for item in item_tree.get_children():
                    item_tree.delete(item)
                total_label.config(text='')
                state['patient_id'] = None
                state['all_sales'] = []
                state['invoices'] = {}
                return

            state['patient_id'] = patient[0]
            state['patient_name'] = f'{patient[1]} {patient[2]}'

            cursor.execute('''
                SELECT s.id, s.sale_date, m.name, s.quantity,
                       s.unit_price, s.total_price, s.invoice_id
                FROM medicine_sales s
                JOIN medicines m ON s.medicine_id = m.id
                WHERE s.patient_id = ?
                ORDER BY s.sale_date DESC
            ''', (patient[0],))
            state['all_sales'] = cursor.fetchall()
            conn.close()

            state['invoices'] = {}
            for sale in state['all_sales']:
                s_id, s_date, med_name, qty, price, total, inv_id = sale
                group_key = inv_id if inv_id else f'#{s_id}'

                if group_key not in state['invoices']:
                    state['invoices'][group_key] = {
                        'invoice_id': group_key,
                        'date': s_date.split(' ')[0] if s_date else '',
                        'time': s_date.split(' ')[1] if s_date and ' ' in s_date else '—',
                        'full_datetime': s_date,
                        'items': [],
                        'total': 0,
                        'items_count': 0,
                    }

                state['invoices'][group_key]['items'].append(sale)
                state['invoices'][group_key]['total'] += total
                state['invoices'][group_key]['items_count'] += 1

            apply_filters()
            show_status(f'✅ Found {len(state["invoices"])} invoice(s) for "{state["patient_name"]}"')

        def apply_filters():
            if state['patient_id'] is None:
                return

            from_date, to_date = get_date_range(
                filter_var.get(),
                from_var.get().strip(),
                to_var.get().strip()
            )

            filtered_invoices = {}
            for inv_key, inv_data in state['invoices'].items():
                inv_date = inv_data['date']
                if from_date and inv_date < from_date:
                    continue
                if to_date and inv_date > to_date:
                    continue
                filtered_invoices[inv_key] = inv_data

            for item in inv_tree.get_children():
                inv_tree.delete(item)

            grand_total = 0
            for inv_key, inv_data in sorted(filtered_invoices.items(),
                                            key=lambda x: x[1]['full_datetime'],
                                            reverse=True):
                inv_tree.insert('', 'end', values=(
                    inv_data['invoice_id'], inv_data['date'], inv_data['time'],
                    f"{inv_data['items_count']} items",
                    f"{inv_data['total']:,.0f}",
                    state['patient_name'],
                ), tags=(inv_data['invoice_id'],))
                grand_total += inv_data['total']

            for item in item_tree.get_children():
                item_tree.delete(item)

            for inv_key, inv_data in sorted(filtered_invoices.items(),
                                            key=lambda x: x[1]['full_datetime'],
                                            reverse=True):
                for sale in inv_data['items']:
                    s_id, s_date, med_name, qty, price, total, inv_id = sale
                    date_part = s_date.split(' ')[0] if s_date else ''
                    time_part = s_date.split(' ')[1] if s_date and ' ' in s_date else '—'
                    display_id = inv_id if inv_id else f'#{s_id}'

                    item_tree.insert('', 'end', values=(
                        display_id, date_part, time_part,
                        med_name, qty, f'{price:,.0f}',
                        f'{total:,.0f}', state['patient_name']
                    ))

            filter_info = ''
            if filter_var.get() != 'All':
                filter_info = f'  |  📅 {filter_var.get()}'

            if not filtered_invoices:
                total_label.config(text='')
                show_status('📭 No invoices in the selected range', '#e67e22')
            else:
                total_label.config(
                    text=f'💰 Invoices: {len(filtered_invoices)}  |  '
                         f'Total: {grand_total:,.0f} IRR{filter_info}'
                )

        def load_pharmacy_payments():
            for item in pay_tree.get_children():
                pay_tree.delete(item)

            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                SELECT
                    pp.id, pp.receipt_number,
                    p.first_name || ' ' || p.last_name,
                    pp.amount, pp.discount, pp.final_amount,
                    pp.payment_method, pp.payment_date
                FROM pharmacy_payments pp
                JOIN patients p ON pp.patient_id = p.id
                ORDER BY pp.payment_date DESC
            ''')
            payments = cursor.fetchall()
            conn.close()

            for p in payments:
                pay_tree.insert('', 'end', values=(
                    p[0], p[1], p[2], f'{p[3]:,.0f}', f'{p[4]:,.0f}',
                    f'{p[5]:,.0f}', p[6], p[7][:16] if p[7] else '—'
                ))

        load_pharmacy_payments()

        def view_selected_invoice():
            current_tab = notebook.index(notebook.select())
            if current_tab == 0:
                selected = inv_tree.selection()
                if not selected:
                    show_status('⚠️ Select an invoice first!', '#e67e22')
                    return
                inv_id = inv_tree.item(selected[0])['tags'][0]
            else:
                selected = item_tree.selection()
                if not selected:
                    show_status('⚠️ Select a row first!', '#e67e22')
                    return
                inv_id = str(item_tree.item(selected[0])['values'][0]).strip()

            if inv_id.startswith('INV-'):
                self.show_invoice_by_invoice_id(inv_id)
            else:
                try:
                    sale_id = int(inv_id.replace('#', '').strip())
                    self.show_invoice_by_sale_id(sale_id)
                except ValueError:
                    messagebox.showerror('Error', f'❌ Invalid ID: {inv_id}')

        # ===== ۸. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '🔍 Search', 'command': search_history,
             'color': 'primary', 'width': 140},
            {'text': '🧾 View Invoice', 'command': view_selected_invoice,
             'color': 'warning', 'width': 180},
            {'text': '❌ Close', 'command': win.destroy,
             'color': 'danger', 'width': 120},
        ], bg='#ecf0f1')

        # ===== ۹. Maximize =====
        win.update_idletasks()
        win.state('zoomed')

        nid_entry.focus()


    def dispense_medicines_window(self):
        """فرم تحویل دارو با کد ملی + پرداخت"""
        win = tk.Toplevel(self.root)
        win.title('📋 Dispense Medicines')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#ecf0f1')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#ecf0f1', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#ecf0f1')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text='📋 Dispense Medicines (with Payment)',
                 font=('Arial', 16, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(pady=15)

        # ===== ۴. جستجو با کد ملی =====
        search_frame = tk.Frame(scrollable, bg='#ecf0f1')
        search_frame.pack(fill='x', padx=40, pady=10)

        tk.Label(search_frame, text='🔍 National ID:',
                 font=('Arial', 12, 'bold'), bg='#ecf0f1').pack(side='left', padx=(0, 10))

        nid_var = tk.StringVar()
        nid_entry = tk.Entry(search_frame, textvariable=nid_var,
                             font=('Arial', 12), width=20,
                             bg='white', relief='solid', bd=1)
        nid_entry.pack(side='left', ipady=4, padx=(0, 10))

        # ===== ۵. نتایج =====
        result_frame = tk.Frame(scrollable, bg='#ecf0f1')
        result_frame.pack(fill='both', expand=True, padx=40, pady=10)

        cols = ('Select', 'Prescription ID', 'Date', 'Medicine', 'Qty', 'Price', 'Total')
        tree = ttk.Treeview(result_frame, columns=cols, show='headings', height=12)

        col_widths = {'Select': 60, 'Prescription ID': 100, 'Date': 110,
                      'Medicine': 200, 'Qty': 60, 'Price': 100, 'Total': 120}
        for col in cols:
            tree.heading(col, text=col)
            tree.column(col, width=col_widths[col], anchor='center', minwidth=60)

        scroll = ttk.Scrollbar(result_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scroll.set)
        tree.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        # ===== ۶. جمع کل =====
        total_frame = tk.Frame(scrollable, bg='#f8f9fa', relief='solid', bd=2)
        total_frame.pack(fill='x', padx=40, pady=10)

        total_var = tk.StringVar(value='0 IRR')
        tk.Label(total_frame, text='💰 TOTAL:',
                 font=('Arial', 14, 'bold'), bg='#f8f9fa').pack(side='left', padx=15, pady=8)
        tk.Label(total_frame, textvariable=total_var,
                 font=('Arial', 14, 'bold'), bg='#f8f9fa', fg='#27ae60').pack(side='left')

        # ===== ۷. ذخیره‌ی نسخه‌ها =====
        dispense_data = {'patient_id': None, 'patient_name': '', 'items': []}

        def search_prescriptions(event=None):
            nid = nid_var.get().strip()

            if not nid:
                messagebox.showwarning('Warning', '⚠️ Please enter National ID!')
                return

            if not nid.isdigit() or len(nid) != 10:
                messagebox.showerror('Error', '❌ National ID must be 10 digits!')
                return

            conn = get_connection()
            cursor = conn.cursor()

            cursor.execute('''
                SELECT id, first_name || ' ' || last_name
                FROM patients WHERE national_id = ?
            ''', (nid,))
            patient = cursor.fetchone()

            if not patient:
                conn.close()
                messagebox.showerror('Error', f'❌ No patient found with National ID: {nid}')
                return

            patient_id = patient[0]
            patient_name = patient[1]
            dispense_data['patient_id'] = patient_id
            dispense_data['patient_name'] = patient_name

            from datetime import datetime
            today = datetime.now().strftime('%Y-%m-%d')

            cursor.execute('''
                SELECT
                    pr.id, mr.record_date, m.id, m.name, pr.quantity,
                    m.price, (pr.quantity * m.price) as total
                FROM prescriptions pr
                JOIN medical_records mr ON pr.record_id = mr.id
                JOIN medicines m ON pr.medicine_id = m.id
                WHERE mr.patient_id = ?
                  AND DATE(mr.record_date) = ?
                  AND pr.id NOT IN (
                      SELECT prescription_id FROM dispensed_medicines
                  )
                ORDER BY pr.id
            ''', (patient_id, today))

            prescriptions = cursor.fetchall()
            conn.close()

            for item in tree.get_children():
                tree.delete(item)

            dispense_data['items'] = []

            if not prescriptions:
                messagebox.showinfo('Info',
                                    f'✅ No pending prescriptions for {patient_name}\n\n'
                                    f'All prescriptions have been dispensed.')
                return

            total_amount = 0
            for pr in prescriptions:
                pr_id, date, med_id, med_name, qty, price, total = pr
                dispense_data['items'].append({
                    'prescription_id': pr_id,
                    'medicine_id': med_id,
                    'medicine_name': med_name,
                    'quantity': qty,
                    'price': price,
                    'total': total,
                    'selected': True,
                })
                tree.insert('', 'end', values=(
                    '✅', pr_id, date, med_name, qty,
                    f'{price:,.0f}', f'{total:,.0f}'
                ))
                total_amount += total

            total_var.set(f'{total_amount:,.0f} IRR')

        nid_entry.bind('<Return>', search_prescriptions)

        # ===== ۸. تابع پرداخت =====
        def do_dispense():
            if not dispense_data['items']:
                messagebox.showwarning('Warning', '⚠️ No prescriptions to dispense!')
                return

            payment_window(win)

        def payment_window(parent_win):
            """پنجره‌ی پرداخت"""
            pay_win = tk.Toplevel(parent_win)
            pay_win.title('💳 Payment')
            pay_win.configure(bg='#ecf0f1')
            pay_win.resizable(True, True)

            # دکمه‌ها پایین
            pay_btn_container = tk.Frame(pay_win, bg='#ecf0f1')
            pay_btn_container.pack(side='bottom', fill='x', pady=15)

            # بدنه Scrollable
            pay_body = tk.Frame(pay_win, bg='#ecf0f1')
            pay_body.pack(fill='both', expand=True, padx=10, pady=5)

            pay_canvas = tk.Canvas(pay_body, bg='#ecf0f1', highlightthickness=0)
            pay_scrollbar = ttk.Scrollbar(pay_body, orient='vertical', command=pay_canvas.yview)
            pay_scrollable = tk.Frame(pay_canvas, bg='#ecf0f1')

            pay_scrollable.bind('<Configure>', lambda e: pay_canvas.configure(scrollregion=pay_canvas.bbox('all')))
            pay_canvas_window = pay_canvas.create_window((0, 0), window=pay_scrollable, anchor='nw')
            pay_canvas.configure(yscrollcommand=pay_scrollbar.set)
            pay_canvas.bind('<Configure>', lambda e: pay_canvas.itemconfig(pay_canvas_window, width=e.width))

            pay_canvas.pack(side='left', fill='both', expand=True)
            auto_scrollbar(pay_canvas, pay_scrollbar, pay_scrollable)

            # عنوان
            tk.Label(pay_scrollable, text='💳 Payment',
                     font=('Arial', 16, 'bold'), bg='#ecf0f1',
                     fg='#2c3e50').pack(pady=15)

            # اطلاعات بیمار
            info = tk.Frame(pay_scrollable, bg='#f8f9fa', relief='solid', bd=1)
            info.pack(fill='x', padx=40, pady=10)

            tk.Label(info, text=f'👤 Patient: {dispense_data["patient_name"]}',
                     font=('Arial', 11, 'bold'), bg='#f8f9fa',
                     anchor='w').pack(anchor='w', padx=15, pady=(10, 3))

            subtotal = sum(item['total'] for item in dispense_data['items'])
            tk.Label(info, text=f'💰 Total: {subtotal:,.0f} IRR',
                     font=('Arial', 12, 'bold'), bg='#f8f9fa',
                     fg='#27ae60', anchor='w').pack(anchor='w', padx=15, pady=(3, 10))

            # فرم
            form = tk.Frame(pay_scrollable, bg='#ecf0f1')
            form.pack(fill='x', padx=40, pady=10)

            tk.Label(form, text='Discount (IRR):',
                     font=('Arial', 11, 'bold'), bg='#ecf0f1',
                     anchor='w').pack(anchor='w', pady=(0, 5))

            discount_var = tk.StringVar(value='0')
            tk.Entry(form, textvariable=discount_var,
                     font=('Arial', 11), bg='white').pack(fill='x', ipady=4, pady=(0, 15))

            tk.Label(form, text='Payment Method:',
                     font=('Arial', 11, 'bold'), bg='#ecf0f1',
                     anchor='w').pack(anchor='w', pady=(0, 5))

            method_var = tk.StringVar(value='Cash')
            for method in ['Cash', 'Card', 'Insurance']:
                tk.Radiobutton(form, text=method, variable=method_var,
                               value=method, bg='#ecf0f1',
                               font=('Arial', 10)).pack(anchor='w')

            # جمع نهایی
            final_frame = tk.Frame(pay_scrollable, bg='white', relief='solid', bd=2)
            final_frame.pack(fill='x', padx=40, pady=15)

            final_var = tk.StringVar()

            def update_final(*args):
                try:
                    discount = float(discount_var.get() or 0)
                except:
                    discount = 0
                final = max(subtotal - discount, 0)
                final_var.set(f'{final:,.0f}')

            discount_var.trace_add('write', update_final)
            update_final()

            tk.Label(final_frame, text='FINAL AMOUNT:',
                     font=('Arial', 13, 'bold'), bg='white').pack(side='left', padx=15, pady=10)
            tk.Label(final_frame, textvariable=final_var,
                     font=('Arial', 14, 'bold'), bg='white', fg='#27ae60').pack(side='left', padx=5)
            tk.Label(final_frame, text='IRR',
                     font=('Arial', 12), bg='white').pack(side='left')

            def confirm_payment():
                try:
                    discount = float(discount_var.get() or 0)
                except:
                    discount = 0

                final = max(subtotal - discount, 0)

                if final < 0:
                    messagebox.showerror('Error', '❌ Discount cannot exceed total!')
                    return

                if not messagebox.askyesno('Confirm Payment',
                                           f'Confirm payment?\n\n'
                                           f'Patient: {dispense_data["patient_name"]}\n'
                                           f'Subtotal: {subtotal:,.0f} IRR\n'
                                           f'Discount: {discount:,.0f} IRR\n'
                                           f'Final: {final:,.0f} IRR\n'
                                           f'Method: {method_var.get()}'):
                    return

                try:
                    import random
                    from datetime import datetime
                    receipt = f'PH-{datetime.now().strftime("%Y%m%d%H%M%S")}-{random.randint(100, 999)}'

                    conn = get_connection()
                    cursor = conn.cursor()

                    cursor.execute('''
                        INSERT INTO pharmacy_payments
                        (receipt_number, patient_id, amount, discount, final_amount,
                         payment_method, payment_status, created_by)
                        VALUES (?, ?, ?, ?, ?, ?, 'Paid', ?)
                    ''', (receipt, dispense_data['patient_id'], subtotal, discount,
                          final, method_var.get(),
                          getattr(self.root, 'current_user', {}).get('username', 'admin')))

                    payment_id = cursor.lastrowid

                    for item in dispense_data['items']:
                        cursor.execute('''
                            INSERT INTO dispensed_medicines
                            (prescription_id, patient_id, medicine_id,
                             medicine_name, quantity, unit_price, total_price, dispensed_by)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (item['prescription_id'], dispense_data['patient_id'],
                              item['medicine_id'], item['medicine_name'], item['quantity'],
                              item['price'], item['total'],
                              getattr(self.root, 'current_user', {}).get('username', 'admin')))

                        cursor.execute('''
                            UPDATE medicines
                            SET stock_quantity = stock_quantity - ?
                            WHERE id = ?
                        ''', (item['quantity'], item['medicine_id']))

                    conn.commit()
                    conn.close()

                    self.load_pharmacy_stats()

                    self.print_dispense_receipt(
                        receipt=receipt,
                        patient_name=dispense_data['patient_name'],
                        items=dispense_data['items'],
                        subtotal=subtotal,
                        discount=discount,
                        final=final,
                        method=method_var.get()
                    )

                    messagebox.showinfo('✅ Success',
                                        f'✅ Medicines dispensed!\n\n'
                                        f'Receipt: {receipt}\n'
                                        f'Total: {final:,.0f} IRR\n'
                                        f'Method: {method_var.get()}')

                    try:
                        log_action('DISPENSE', entity_type='pharmacy_payment',
                                   entity_id=payment_id,
                                   details=f'Dispensed {len(dispense_data["items"])} medicines, Receipt: {receipt}',
                                   user=getattr(self.root, 'current_user', None))
                    except:
                        pass

                    self.load_medicines()
                    self.load_pharmacy_stats()

                    pay_win.destroy()
                    win.destroy()


                except Exception as e:
                    messagebox.showerror('Error', f'❌ {e}')

            create_popup_buttons(pay_btn_container, [
                {'text': '💾 Confirm & Dispense', 'command': confirm_payment,
                 'color': 'success', 'width': 220},
                {'text': '❌ Cancel', 'command': pay_win.destroy,
                 'color': 'danger', 'width': 140},
            ], bg='#ecf0f1')

            pay_win.update_idletasks()
            pay_win.state('zoomed')

        # ===== ۹. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '🔍 Search', 'command': search_prescriptions,
             'color': 'primary', 'width': 140},
            {'text': '💳 Dispense & Pay', 'command': do_dispense,
             'color': 'success', 'width': 200},
            {'text': '❌ Close', 'command': win.destroy,
             'color': 'danger', 'width': 120},
        ], bg='#ecf0f1')

        # ===== ۱۰. Maximize =====
        win.update_idletasks()
        win.state('zoomed')

        nid_entry.focus()



    def print_dispense_receipt(self, receipt, patient_name, items, subtotal, discount, final, method):
        """چاپ رسید تحویل دارو"""
        from datetime import datetime

        lines = []
        lines.append('=' * 60)
        lines.append(' ' * 15 + 'HOSPITAL MANAGEMENT SYSTEM')
        lines.append(' ' * 18 + 'PHARMACY RECEIPT')
        lines.append('=' * 60)
        lines.append('')
        lines.append(f'  Receipt #: {receipt}')
        lines.append(f'  Date: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        lines.append(f'  Patient: {patient_name}')
        lines.append('')
        lines.append('-' * 60)
        lines.append('  MEDICINES')
        lines.append('-' * 60)

        for i, item in enumerate(items, 1):
            lines.append(f'  {i}. {item["medicine_name"]}')
            lines.append(f'     {item["quantity"]} package(s) x {item["price"]:,.0f} = {item["total"]:,.0f} IRR')

        lines.append('')
        lines.append('-' * 60)
        lines.append(f'  Subtotal:    {subtotal:>20,.0f} IRR')
        lines.append(f'  Discount:    {discount:>20,.0f} IRR')
        lines.append('=' * 60)
        lines.append(f'  TOTAL:       {final:>20,.0f} IRR')
        lines.append('=' * 60)
        lines.append('')
        lines.append(f'  Payment Method: {method}')
        lines.append('')
        lines.append('=' * 60)
        lines.append('  Thank you!')
        lines.append('=' * 60)

        content = '\n'.join(lines)

        success, pdf_path = print_text_file(content, f'pharmacy_receipt_{receipt}.txt', open_pdf=False)

        if success:
            # پرینت یا باز کردن
            if WIN32_AVAILABLE:
                printer = get_printer_for('receipt')
                virtual_keywords = ['Print to PDF', 'XPS', 'OneNote', 'Fax', 'Virtual', 'Writer']
                is_virtual = any(kw.lower() in (printer or '').lower() for kw in virtual_keywords)

                if is_virtual or not printer:
                    try:
                        os.startfile(pdf_path)
                    except:
                        pass
                else:
                    print_pdf_to_printer(pdf_path, printer)



    def show_invoice_by_sale_id(self, sale_id):
        """Show detailed invoice for a specific sale"""
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT
                s.id, s.sale_date, s.quantity, s.unit_price, s.total_price,
                m.name, m.category, m.manufacturer, m.expiry_date,
                p.national_id, p.first_name, p.last_name, p.phone
            FROM medicine_sales s
            JOIN medicines m ON s.medicine_id = m.id
            JOIN patients p ON s.patient_id = p.id
            WHERE s.id = ?
        ''', (sale_id,))
        sale = cursor.fetchone()
        conn.close()

        if not sale:
            messagebox.showerror('Error', f'❌ Sale #{sale_id} not found!')
            return

        (s_id, sale_date, qty, unit_price, total_price,
         med_name, med_category, med_manufacturer, med_expiry,
         nid, first_name, last_name, phone) = sale

        # پنجره‌ی فاکتور
        inv_win = tk.Toplevel(self.root)
        inv_win.title(f'🧾 Invoice #{s_id}')
        inv_win.geometry('550x650')
        inv_win.configure(bg='white')
        inv_win.resizable(False, False)

        inv_win.update_idletasks()
        w, h = 550, 650
        x = (inv_win.winfo_screenwidth() // 2) - (w // 2)
        y = (inv_win.winfo_screenheight() // 2) - (h // 2)
        inv_win.geometry(f'{w}x{h}+{x}+{y}')

        # متن فاکتور
        inv_text = tk.Text(inv_win, font=('Courier New', 10),
                            bg='white', fg='#2c3e50',
                            padx=20, pady=20, relief='flat', wrap='word')
        inv_text.pack(fill='both', expand=True, padx=10, pady=10)

        inv_text.insert('end', '=' * 55 + '\n')
        inv_text.insert('end', '         🏥 HOSPITAL MANAGEMENT SYSTEM\n')
        inv_text.insert('end', '           PHARMACY INVOICE\n')
        inv_text.insert('end', '=' * 55 + '\n\n')

        inv_text.insert('end', f'🧾 Invoice #: {s_id}\n')
        inv_text.insert('end', f'📅 Date: {sale_date}\n')
        inv_text.insert('end', '\n' + '-' * 55 + '\n')
        inv_text.insert('end', '👤 PATIENT INFORMATION:\n')
        inv_text.insert('end', '-' * 55 + '\n')
        inv_text.insert('end', f'   Name: {first_name} {last_name}\n')
        inv_text.insert('end', f'   National ID: {nid}\n')
        inv_text.insert('end', f'   Phone: {phone}\n')
        inv_text.insert('end', '\n' + '-' * 55 + '\n')
        inv_text.insert('end', '💊 MEDICINE DETAILS:\n')
        inv_text.insert('end', '-' * 55 + '\n')
        inv_text.insert('end', f'   Name: {med_name}\n')
        inv_text.insert('end', f'   Category: {med_category}\n')
        inv_text.insert('end', f'   Manufacturer: {med_manufacturer or "N/A"}\n')
        inv_text.insert('end', f'   Expiry: {med_expiry or "N/A"}\n')
        inv_text.insert('end', '\n' + '-' * 55 + '\n')
        inv_text.insert('end', f'   Quantity: {qty}\n')
        inv_text.insert('end', f'   Unit Price: {unit_price:,.0f} IRR\n')
        inv_text.insert('end', '=' * 55 + '\n')
        inv_text.insert('end', f'   TOTAL: {total_price:,.0f} IRR\n')
        inv_text.insert('end', '=' * 55 + '\n\n')
        inv_text.insert('end', '   Thank you for choosing our pharmacy!\n')
        inv_text.insert('end', '=' * 55 + '\n')

        inv_text.config(state='disabled')

        # دکمه‌ها
        btn_frame = tk.Frame(inv_win, bg='white')  # یا receipt_win
        btn_frame.pack(fill='x', pady=10, padx=10)

        def save_as_txt():
            from tkinter import filedialog
            file_path = filedialog.asksaveasfilename(
                defaultextension='.txt',
                filetypes=[('Text files', '*.txt')],
                initialfile=f'invoice_{datetime.now().strftime("%Y%m%d_%H%M%S")}.txt'
            )
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(inv_text.get('1.0', 'end-1c'))
                messagebox.showinfo('Success', f'✅ Saved to:\n{file_path}')

        def print_invoice():
            # ===== ۱. ساخت PDF =====
            content = inv_text.get('1.0', 'end-1c')
            success, result = print_text_file(content, 'invoice_print.txt', open_pdf=False)

            if not success:
                messagebox.showerror('❌ PDF Error', f'Could not create PDF:\n{result}')
                return

            pdf_path = result  # ← مسیر PDF

            # ===== ۲. اگه پرینتر موجود بود =====
            if WIN32_AVAILABLE:
                printer = get_printer_for('invoice')  # ← تغییر
                if printer:
                    ok, msg = print_pdf_to_printer(pdf_path, printer)
                    if ok:
                        messagebox.showinfo(
                            '🖨️ Sent to Printer',
                            f'Document sent to:\n\n🖨️ {printer}\n\n'
                            f'Check the printer for output.'
                        )
                    else:
                        # ===== اگه پرینت نشد، PDF رو باز کن =====
                        try:
                            os.startfile(pdf_path)
                        except:
                            pass

                        messagebox.showwarning(
                            '⚠️ Could Not Print',
                            f'Printer: {printer}\n\n'
                            f'The document could not be sent to the printer.\n'
                            f'This usually happens when:\n'
                            f'• The default printer is not connected\n'
                            f'• The printer is a virtual PDF printer\n'
                            f'• The printer is offline\n\n'
                            f'📄 PDF has been opened so you can print manually.'
                        )
                else:
                    try:
                        os.startfile(pdf_path)
                    except:
                        pass
                    messagebox.showwarning(
                        '⚠️ No Default Printer',
                        'No default printer set in Windows.\n\n'
                        'PDF opened for manual printing.'
                    )
            else:
                try:
                    os.startfile(pdf_path)
                except:
                    pass
                messagebox.showinfo(
                    '📄 PDF Opened',
                    'pywin32 not installed.\n\nPDF opened for manual printing.'
                )

        tk.Button(btn_frame, text='💾 Save as TXT', command=save_as_txt,
                  bg='#3498db', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8, cursor='hand2').pack(side='left', padx=3)

        tk.Button(btn_frame, text='🖨️ Print', command=print_invoice,
                  bg='#9b59b6', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8, cursor='hand2').pack(side='left', padx=3)

        tk.Button(btn_frame, text='❌ Close', command=inv_win.destroy,
                  bg='#e74c3c', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8, cursor='hand2').pack(side='right', padx=3)


    def show_invoice_by_invoice_id(self, invoice_id):
        """Show full invoice for a group of sales with same invoice_id"""
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT
                s.id, s.sale_date, s.quantity, s.unit_price, s.total_price,
                m.name, m.category, m.manufacturer, m.expiry_date,
                p.national_id, p.first_name, p.last_name, p.phone
            FROM medicine_sales s
            JOIN medicines m ON s.medicine_id = m.id
            JOIN patients p ON s.patient_id = p.id
            WHERE s.invoice_id = ?
            ORDER BY s.id
        ''', (invoice_id,))
        sales = cursor.fetchall()
        conn.close()

        if not sales:
            messagebox.showerror('Error', f'❌ Invoice "{invoice_id}" not found!')
            return

        # اطلاعات بیمار از اولین ردیف
        first_sale = sales[0]
        nid = first_sale[9]
        first_name = first_sale[10]
        last_name = first_sale[11]
        phone = first_sale[12]
        sale_date = first_sale[1]

        # پنجره‌ی فاکتور
        inv_win = tk.Toplevel(self.root)
        inv_win.title(f'🧾 Invoice {invoice_id}')
        inv_win.geometry('600x700')
        inv_win.configure(bg='white')
        inv_win.resizable(False, False)

        inv_win.update_idletasks()
        w, h = 600, 700
        x = (inv_win.winfo_screenwidth() // 2) - (w // 2)
        y = (inv_win.winfo_screenheight() // 2) - (h // 2)
        inv_win.geometry(f'{w}x{h}+{x}+{y}')

        # متن فاکتور
        inv_text = tk.Text(inv_win, font=('Courier New', 10),
                            bg='white', fg='#2c3e50',
                            padx=20, pady=20, relief='flat', wrap='word')
        inv_text.pack(fill='both', expand=True, padx=10, pady=10)

        inv_text.insert('end', '=' * 55 + '\n')
        inv_text.insert('end', '         🏥 HOSPITAL MANAGEMENT SYSTEM\n')
        inv_text.insert('end', '           PHARMACY INVOICE\n')
        inv_text.insert('end', '=' * 55 + '\n\n')

        inv_text.insert('end', f'🧾 Invoice #: {invoice_id}\n')
        inv_text.insert('end', f'📅 Date: {sale_date}\n')
        inv_text.insert('end', '\n' + '-' * 55 + '\n')
        inv_text.insert('end', '👤 PATIENT INFORMATION:\n')
        inv_text.insert('end', '-' * 55 + '\n')
        inv_text.insert('end', f'   Name: {first_name} {last_name}\n')
        inv_text.insert('end', f'   National ID: {nid}\n')
        inv_text.insert('end', f'   Phone: {phone}\n')
        inv_text.insert('end', '\n' + '-' * 55 + '\n')
        inv_text.insert('end', '💊 MEDICINES:\n')
        inv_text.insert('end', '-' * 55 + '\n')

        grand_total = 0
        for sale in sales:
            (s_id, s_date, qty, unit_price, total_price,
             med_name, med_category, med_manufacturer, med_expiry,
             _, _, _, _) = sale

            inv_text.insert('end',
                f'  {med_name:<25} x{qty:<3} = {total_price:>15,.0f}\n')
            grand_total += total_price

        inv_text.insert('end', '\n' + '-' * 55 + '\n')
        inv_text.insert('end', '=' * 55 + '\n')
        inv_text.insert('end', f'  {"TOTAL:":<40} {grand_total:>10,.0f} IRR\n')
        inv_text.insert('end', '=' * 55 + '\n\n')
        inv_text.insert('end', '   Thank you for choosing our pharmacy!\n')
        inv_text.insert('end', '=' * 55 + '\n')

        inv_text.config(state='disabled')

        # دکمه‌ها
        btn_frame = tk.Frame(inv_win, bg='white')  # یا receipt_win
        btn_frame.pack(fill='x', pady=10, padx=10)

        def save_as_txt():
            from tkinter import filedialog
            file_path = filedialog.asksaveasfilename(
                defaultextension='.txt',
                filetypes=[('Text files', '*.txt')],
                initialfile=f'invoice_{datetime.now().strftime("%Y%m%d_%H%M%S")}.txt'
            )
            if file_path:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(inv_text.get('1.0', 'end-1c'))
                messagebox.showinfo('Success', f'✅ Saved to:\n{file_path}')

        def print_invoice():
            # ===== ۱. ساخت PDF =====
            content = inv_text.get('1.0', 'end-1c')
            success, result = print_text_file(content, 'invoice_print.txt', open_pdf=False)

            if not success:
                messagebox.showerror('❌ PDF Error', f'Could not create PDF:\n{result}')
                return

            pdf_path = result  # ← مسیر PDF

            # ===== ۲. اگه پرینتر موجود بود =====
            if WIN32_AVAILABLE:
                printer = get_printer_for('invoice')  # ← تغییر
                if printer:
                    ok, msg = print_pdf_to_printer(pdf_path, printer)
                    if ok:
                        messagebox.showinfo(
                            '🖨️ Sent to Printer',
                            f'Document sent to:\n\n🖨️ {printer}\n\n'
                            f'Check the printer for output.'
                        )
                    else:
                        # ===== اگه پرینت نشد، PDF رو باز کن =====
                        try:
                            os.startfile(pdf_path)
                        except:
                            pass

                        messagebox.showwarning(
                            '⚠️ Could Not Print',
                            f'Printer: {printer}\n\n'
                            f'The document could not be sent to the printer.\n'
                            f'This usually happens when:\n'
                            f'• The default printer is not connected\n'
                            f'• The printer is a virtual PDF printer\n'
                            f'• The printer is offline\n\n'
                            f'📄 PDF has been opened so you can print manually.'
                        )
                else:
                    try:
                        os.startfile(pdf_path)
                    except:
                        pass
                    messagebox.showwarning(
                        '⚠️ No Default Printer',
                        'No default printer set in Windows.\n\n'
                        'PDF opened for manual printing.'
                    )
            else:
                try:
                    os.startfile(pdf_path)
                except:
                    pass
                messagebox.showinfo(
                    '📄 PDF Opened',
                    'pywin32 not installed.\n\nPDF opened for manual printing.'
                )

        tk.Button(btn_frame, text='💾 Save as TXT', command=save_as_txt,
                  bg='#3498db', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8, cursor='hand2').pack(side='left', padx=3)

        tk.Button(btn_frame, text='🖨️ Print', command=print_invoice,
                  bg='#9b59b6', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8, cursor='hand2').pack(side='left', padx=3)

        tk.Button(btn_frame, text='❌ Close', command=inv_win.destroy,
                  bg='#e74c3c', fg='white', font=('Arial', 11, 'bold'),
                  padx=15, pady=8, cursor='hand2').pack(side='right', padx=3)

    def delete_medicine(self):
        selected = self.medicines_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select at least one medicine!')
            return

        count = len(selected)

        # چک کردن وجود فروش
        conn = get_connection()
        cursor = conn.cursor()
        has_sales = False
        for item in selected:
            real_id = self.medicines_tree.item(item)['tags'][0]
            cursor.execute('SELECT COUNT(*) FROM medicine_sales WHERE medicine_id = ?', (real_id,))
            if cursor.fetchone()[0] > 0:
                has_sales = True
                break
        conn.close()

        if has_sales:
            if not messagebox.askyesno('Warning',
                                       f'⚠️ Some of the selected medicines have sales!\n'
                                       'Deleting will affect sales history.\n\nContinue?'):
                return

        if count == 1:
            confirm_msg = f'Delete this medicine?'
        else:
            confirm_msg = f'Delete {count} medicine(s)?\n\nThis action cannot be undone!'

        if not messagebox.askyesno('Confirm', confirm_msg):
            return

        try:
            conn = get_connection()
            cursor = conn.cursor()

            deleted = 0
            for item in selected:
                real_id = self.medicines_tree.item(item)['tags'][0]
                try:
                    cursor.execute('DELETE FROM medicines WHERE id = ?', (real_id,))
                    deleted += 1
                except sqlite3.IntegrityError:
                    pass  # رد کردن داروهایی که فروش دارن

            conn.commit()
            conn.close()

            messagebox.showinfo('Success', f'✅ {deleted} medicine(s) deleted!')
            self.load_medicines()

        except Exception as e:
            messagebox.showerror('Error', f'❌ Error: {e}')

    # ==========================================
    # Users Tab
    # ==========================================

    def create_users_tab(self):
        tk.Label(self.users_tab, text='👤 User Management',
                 font=('Arial', 14, 'bold'), bg=WINDOW_BG, fg='#2c3e50').pack(pady=15)

        # ===== دکمه‌ها ردیف ۱ =====
        btn_bar1 = create_button_bar(self.users_tab, bg=WINDOW_BG)

        create_standard_button(btn_bar1, '➕ Add User', self.add_user_window, color='success')
        create_standard_button(btn_bar1, '✏️ Edit', self.edit_user_window, color='warning')
        create_standard_button(btn_bar1, '🔑 Change Password', self.change_password_window, color='info')

        # ===== دکمه‌ها ردیف ۲ =====
        btn_bar2 = create_button_bar(self.users_tab, bg=WINDOW_BG)

        create_standard_button(btn_bar2, '🔄 Refresh', self.load_users, color='primary')
        create_standard_button(btn_bar2, '✅ Toggle Active', self.toggle_user_active, color='secondary')
        create_standard_button(btn_bar2, '🗑️ Delete', self.delete_user, color='danger')

        # ===== ردیف ۱: Search =====
        row1 = tk.Frame(self.users_tab, bg=WINDOW_BG)
        row1.pack(pady=(5, 3), padx=40, fill='x')

        tk.Label(row1, text='🔍 Search:', font=('Arial', 11, 'bold'),
                 bg=WINDOW_BG, fg='#2c3e50').pack(side='left', padx=(0, 8))

        self.user_search_var = tk.StringVar()
        search_entry = tk.Entry(row1, textvariable=self.user_search_var,
                                font=('Arial', 11), bg='white',
                                relief='solid', bd=1, width=40)
        search_entry.pack(side='left', ipady=6, padx=(0, 8))

        self.user_search_var.trace_add('write',
                                       lambda *args: self.load_users())

        def clear_search():
            self.user_search_var.set('')

        tk.Button(row1, text='✖ Clear', command=clear_search,
                  bg='#95a5a6', fg='white', font=('Arial', 9, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=10, pady=4).pack(side='left')

        # ===== ردیف ۲: Count =====
        row2 = tk.Frame(self.users_tab, bg=WINDOW_BG)
        row2.pack(pady=(0, 10), padx=40, fill='x')

        self.user_count_var = tk.StringVar(value='')
        tk.Label(row2, textvariable=self.user_count_var,
                 font=('Arial', 10, 'italic'), bg=WINDOW_BG,
                 fg='#7f8c8d').pack(side='right')

        # ===== Treeview =====
        columns = ('ID', 'Username', 'Full Name', 'Role', 'Active', 'Last Login')

        tree_frame = tk.Frame(self.users_tab, bg=WINDOW_BG)
        tree_frame.pack(fill='both', expand=True, padx=40, pady=(5, 10))

        self.users_tree = ttk.Treeview(tree_frame, columns=columns,
                                       show='headings', height=14,
                                       selectmode='extended')

        for col in columns:
            self.users_tree.heading(col, text=col)
            self.users_tree.column(col, width=150, anchor='center')

        scroll = ttk.Scrollbar(tree_frame, orient='vertical',
                               command=self.users_tree.yview)
        self.users_tree.configure(yscrollcommand=scroll.set)

        self.users_tree.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        enable_multi_click_selection(self.users_tree)

        self.load_users()



    def create_settings_tab(self):
        """تنظیمات برنامه"""
        tk.Label(self.settings_tab, text='⚙️ Program Settings',
                 font=('Arial', 14, 'bold'), bg=WINDOW_BG,
                 fg='#2c3e50').pack(pady=15)

        style = ttk.Style()
        style.theme_use('clam')
        settings_notebook = ttk.Notebook(self.settings_tab)
        settings_notebook.pack(fill='both', expand=True, padx=20, pady=10)

        # ===== Tab 1: Printer Settings =====
        printer_tab = tk.Frame(settings_notebook, bg='#f8f9fa')
        # ===== Tab 2: Appearance =====
        appearance_tab = tk.Frame(settings_notebook, bg='#f8f9fa')
        settings_notebook.add(appearance_tab, text='🎨 Appearance')
        self.create_appearance_tab(appearance_tab)
        settings_notebook.add(printer_tab, text='🖨️ Printer Settings')
        self.create_printer_settings_tab(printer_tab)

        # ===== Tab 2: Archive Management (ساده) =====
        archive_tab = tk.Frame(settings_notebook, bg='#f8f9fa')
        settings_notebook.add(archive_tab, text='🗂️ Archive Management')
        self.create_archive_tab(archive_tab)

        # ===== Tab 3: About =====
        about_tab = tk.Frame(settings_notebook, bg='#f8f9fa')
        settings_notebook.add(about_tab, text='ℹ️ About')

        tk.Label(about_tab, text='🏥 Hospital Management System',
                 font=('Arial', 18, 'bold'), bg='#f8f9fa',
                 fg='#2c3e50').pack(pady=20)

        info_frame = tk.Frame(about_tab, bg='white', relief='solid', bd=1)
        info_frame.pack(pady=20, padx=40, fill='x')

        info_data = [
            ('📦 Version:', 'v1.0'),
            ('👨‍💻 Developer:', 'Your Name'),
            ('📅 Build Date:', '2026-10-28'),
            ('🔧 Python:', '3.15.0'),
            ('🎨 UI:', 'Tkinter'),
            ('💾 Database:', 'SQLite'),
        ]

        for label, value in info_data:
            row = tk.Frame(info_frame, bg='white')
            row.pack(fill='x', padx=20, pady=5)
            tk.Label(row, text=label, font=('Arial', 10, 'bold'),
                     bg='white', width=18, anchor='w').pack(side='left')
            tk.Label(row, text=value, font=('Arial', 10),
                     bg='white', anchor='w').pack(side='left')

        tk.Label(about_tab, text='© 2026 - All Rights Reserved',
                 font=('Arial', 9, 'italic'), bg='#f8f9fa',
                 fg='#95a5a6').pack(pady=20)

    def create_reports_tab(self):
        """تب گزارش‌های ماهانه"""
        # ===== Scrollable =====
        canvas = tk.Canvas(self.reports_tab, bg=WINDOW_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.reports_tab, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg=WINDOW_BG)

        scrollable.bind('<Configure>',
                        lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>',
                    lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== هدر =====
        tk.Label(scrollable, text='📊 Monthly Reports',
                 font=('Arial', 16, 'bold'), bg=WINDOW_BG,
                 fg='#2c3e50').pack(pady=15)

        # ===== انتخاب ماه =====
        filter_frame = tk.Frame(scrollable, bg=WINDOW_BG)
        filter_frame.pack(pady=10)

        tk.Label(filter_frame, text='📅 Select Month:',
                 font=('Arial', 11, 'bold'), bg=WINDOW_BG,
                 fg='#2c3e50').pack(side='left', padx=(0, 10))

        # ساخت لیست ماه‌ها (از امروز تا ۲۴ ماه قبل)
        from datetime import datetime, timedelta
        months = []
        today = datetime.now()
        for i in range(24):
            month_date = today - timedelta(days=i * 30)
            month_str = month_date.strftime('%B %Y')
            if month_str not in months:
                months.append(month_str)

        self.report_month_var = tk.StringVar(value=months[0])
        month_combo = ttk.Combobox(filter_frame, textvariable=self.report_month_var,
                                   values=months, state='readonly',
                                   width=20, font=('Arial', 11))
        month_combo.pack(side='left', padx=5)

        tk.Button(filter_frame, text='🔄 Generate Report',
                  command=self.load_monthly_report,
                  bg='#3498db', fg='white', font=('Arial', 10, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=15, pady=6).pack(side='left', padx=15)

        # ===== Content Container =====
        self.report_content_frame = tk.Frame(scrollable, bg=WINDOW_BG)
        self.report_content_frame.pack(fill='both', expand=True, padx=20, pady=10)

        # ===== دکمه‌های پایین =====
        btn_frame = tk.Frame(scrollable, bg=WINDOW_BG)
        btn_frame.pack(pady=20)

        tk.Button(btn_frame, text='📊 Export Excel',
                  command=self.export_monthly_report_excel,
                  bg='#27ae60', fg='white', font=('Arial', 11, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=20, pady=10).pack(side='left', padx=5)

        tk.Button(btn_frame, text='🖨️ Print Report',
                  command=self.print_monthly_report,
                  bg='#9b59b6', fg='white', font=('Arial', 11, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=20, pady=10).pack(side='left', padx=5)

        # ===== لود اولیه =====
        self.load_monthly_report()

    def load_monthly_report(self):
        """لود گزارش ماهانه"""
        # ===== پاک کردن قبلی =====
        for widget in self.report_content_frame.winfo_children():
            widget.destroy()

        # ===== محاسبه بازه =====
        from datetime import datetime, timedelta

        month_str = self.report_month_var.get()
        try:
            month_start = datetime.strptime(month_str, '%B %Y')
        except:
            month_start = datetime.now().replace(day=1)

        if month_start.month == 12:
            next_month = month_start.replace(year=month_start.year + 1, month=1)
        else:
            next_month = month_start.replace(month=month_start.month + 1)

        from_date = month_start.strftime('%Y-%m-%d')
        to_date = (next_month - timedelta(days=1)).strftime('%Y-%m-%d')

        conn = get_connection()
        cursor = conn.cursor()

        # ===== ۱. آمار مالی =====
        cursor.execute('''
            SELECT
                COALESCE(SUM(final_amount), 0),
                COUNT(*),
                SUM(CASE WHEN payment_status = 'Paid' THEN final_amount ELSE 0 END),
                SUM(CASE WHEN payment_status = 'Pending' THEN final_amount ELSE 0 END)
            FROM payments
            WHERE DATE(payment_date) BETWEEN ? AND ?
        ''', (from_date, to_date))
        financial = cursor.fetchone()

        total_revenue = financial[0] or 0
        payment_count = financial[1] or 0
        total_paid = financial[2] or 0
        total_pending = financial[3] or 0

        # ===== ۲. آمار بیماران =====
        cursor.execute('''
            SELECT COUNT(*) FROM medical_records
            WHERE DATE(record_date) BETWEEN ? AND ?
        ''', (from_date, to_date))
        total_visits = cursor.fetchone()[0]

        cursor.execute('''
            SELECT COUNT(DISTINCT patient_id) FROM medical_records
            WHERE DATE(record_date) BETWEEN ? AND ?
        ''', (from_date, to_date))
        unique_patients = cursor.fetchone()[0]

        # ===== ۳. آمار نسخه‌ها =====
        cursor.execute('''
            SELECT COUNT(*) FROM prescriptions pr
            JOIN medical_records mr ON pr.record_id = mr.id
            WHERE DATE(mr.record_date) BETWEEN ? AND ?
        ''', (from_date, to_date))
        total_prescriptions = cursor.fetchone()[0]

        # ===== ۴. آمار آزمایش‌ها =====
        cursor.execute('''
            SELECT COUNT(*) FROM lab_requests
            WHERE DATE(request_date) BETWEEN ? AND ?
        ''', (from_date, to_date))
        total_lab = cursor.fetchone()[0]

        # ===== ۵. آمار نوبت‌ها =====
        cursor.execute('''
            SELECT
                COUNT(*),
                SUM(CASE WHEN status = 'Completed' THEN 1 ELSE 0 END),
                SUM(CASE WHEN status = 'Cancelled' THEN 1 ELSE 0 END)
            FROM appointments
            WHERE DATE(appointment_date) BETWEEN ? AND ?
        ''', (from_date, to_date))
        appt = cursor.fetchone()
        total_appts = appt[0] or 0
        completed_appts = appt[1] or 0
        cancelled_appts = appt[2] or 0

        # ===== ۶. Top Doctors =====
        cursor.execute('''
            SELECT
                d.first_name || ' ' || d.last_name,
                COUNT(mr.id) as cnt
            FROM medical_records mr
            JOIN doctors d ON mr.doctor_id = d.id
            WHERE DATE(mr.record_date) BETWEEN ? AND ?
            GROUP BY d.id
            ORDER BY cnt DESC
            LIMIT 5
        ''', (from_date, to_date))
        top_doctors = cursor.fetchall()

        # ===== ۷. Top Medicines =====
        cursor.execute('''
            SELECT
                m.name,
                COUNT(pr.id) as cnt
            FROM prescriptions pr
            JOIN medicines m ON pr.medicine_id = m.id
            JOIN medical_records mr ON pr.record_id = mr.id
            WHERE DATE(mr.record_date) BETWEEN ? AND ?
            GROUP BY m.id
            ORDER BY cnt DESC
            LIMIT 5
        ''', (from_date, to_date))
        top_medicines = cursor.fetchall()

        conn.close()

        # ===== ۸. نمایش =====
        avg_per_day = total_revenue / 30 if total_revenue > 0 else 0

        # ===== آمار مالی =====
        fin_frame = tk.LabelFrame(self.report_content_frame,
                                  text=f' 💰 Financial Summary ({month_str}) ',
                                  font=('Arial', 12, 'bold'),
                                  bg='#f8f9fa', fg='#27ae60',
                                  padx=15, pady=10)
        fin_frame.pack(fill='x', pady=5)

        fin_inner = tk.Frame(fin_frame, bg='#f8f9fa')
        fin_inner.pack(fill='x')

        def make_stat(parent, icon, label, value, color):
            card = tk.Frame(parent, bg='white', relief='solid', bd=1)
            card.pack(side='left', padx=5, fill='x', expand=True)

            tk.Label(card, text=icon, font=('Arial', 20),
                     bg='white', fg=color).pack(pady=(8, 0))
            tk.Label(card, text=value, font=('Arial', 14, 'bold'),
                     bg='white', fg=color).pack()
            tk.Label(card, text=label, font=('Arial', 9),
                     bg='white', fg='#7f8c8d').pack(pady=(0, 8))

        make_stat(fin_inner, '💵', 'Total Revenue',
                  f'{total_revenue:,.0f} IRR', '#27ae60')
        make_stat(fin_inner, '✅', 'Paid',
                  f'{total_paid:,.0f} IRR', '#3498db')
        make_stat(fin_inner, '⏳', 'Pending',
                  f'{total_pending:,.0f} IRR', '#f39c12')
        make_stat(fin_inner, '📊', 'Avg/Day',
                  f'{avg_per_day:,.0f} IRR', '#9b59b6')
        make_stat(fin_inner, '📋', 'Payments',
                  str(payment_count), '#e67e22')

        # ===== آمار کلی =====
        stats_frame = tk.LabelFrame(self.report_content_frame,
                                    text=' 📊 Activity Summary ',
                                    font=('Arial', 12, 'bold'),
                                    bg='#f8f9fa', fg='#3498db',
                                    padx=15, pady=10)
        stats_frame.pack(fill='x', pady=5)

        stats_inner = tk.Frame(stats_frame, bg='#f8f9fa')
        stats_inner.pack(fill='x')

        make_stat(stats_inner, '👥', 'Patients',
                  str(unique_patients), '#3498db')
        make_stat(stats_inner, '📋', 'Visits',
                  str(total_visits), '#27ae60')
        make_stat(stats_inner, '💊', 'Prescriptions',
                  str(total_prescriptions), '#e74c3c')
        make_stat(stats_inner, '🧪', 'Lab Tests',
                  str(total_lab), '#e67e22')
        make_stat(stats_inner, '📅', 'Appointments',
                  str(total_appts), '#9b59b6')

        # ===== Top Doctors و Top Medicines =====
        top_frame = tk.Frame(self.report_content_frame, bg=WINDOW_BG)
        top_frame.pack(fill='x', pady=10)

        # ===== Top Doctors =====
        doc_frame = tk.Frame(top_frame, bg='white', relief='solid', bd=1)
        doc_frame.pack(side='left', fill='both', expand=True, padx=(0, 5))

        doc_header = tk.Frame(doc_frame, bg='#8e44ad')
        doc_header.pack(fill='x')
        doc_title = f'🏆 Top {len(top_doctors)} Doctors' if top_doctors else '🏆 Top Doctors'
        tk.Label(doc_header, text=doc_title,
                 font=('Arial', 11, 'bold'), bg='#8e44ad',
                 fg='white', pady=6).pack()

        doc_content = tk.Frame(doc_frame, bg='white')
        doc_content.pack(fill='both', expand=True, padx=10, pady=8)

        if top_doctors:
            for i, (name, cnt) in enumerate(top_doctors, 1):
                row = tk.Frame(doc_content, bg='white')
                row.pack(fill='x', pady=2)

                colors = ['#f1c40f', '#95a5a6', '#e67e22', '#3498db', '#1abc9c']
                color = colors[i - 1] if i <= 5 else '#95a5a6'

                tk.Label(row, text=f' {i} ', font=('Arial', 10, 'bold'),
                         bg=color, fg='white', width=2).pack(side='left', padx=(0, 8))

                doctor_clean = name.replace('Dr.', '').replace('dr.', '').strip()
                tk.Label(row, text=f'Dr. {doctor_clean}'[:25],
                         font=('Arial', 10), bg='white',
                         anchor='w').pack(side='left', fill='x', expand=True)
                tk.Label(row, text=f'👁️ {cnt}', font=('Arial', 10, 'bold'),
                         bg='white', fg='#8e44ad', width=8,
                         anchor='e').pack(side='right')
        else:
            tk.Label(doc_content, text='📭 No data',
                     font=('Arial', 10, 'italic'),
                     bg='white', fg='#95a5a6').pack(pady=20)

        # ===== Top Medicines =====
        med_frame = tk.Frame(top_frame, bg='white', relief='solid', bd=1)
        med_frame.pack(side='left', fill='both', expand=True, padx=(5, 0))

        med_header = tk.Frame(med_frame, bg='#c0392b')
        med_header.pack(fill='x')
        med_title = f'🏆 Top {len(top_medicines)} Medicines' if top_medicines else '🏆 Top Medicines'
        tk.Label(med_header, text=med_title,
                 font=('Arial', 11, 'bold'), bg='#c0392b',
                 fg='white', pady=6).pack()

        med_content = tk.Frame(med_frame, bg='white')
        med_content.pack(fill='both', expand=True, padx=10, pady=8)

        if top_medicines:
            for i, (name, cnt) in enumerate(top_medicines, 1):
                row = tk.Frame(med_content, bg='white')
                row.pack(fill='x', pady=2)

                colors = ['#f1c40f', '#95a5a6', '#e67e22', '#3498db', '#1abc9c']
                color = colors[i - 1] if i <= 5 else '#95a5a6'

                tk.Label(row, text=f' {i} ', font=('Arial', 10, 'bold'),
                         bg=color, fg='white', width=2).pack(side='left', padx=(0, 8))

                tk.Label(row, text=name[:25], font=('Arial', 10),
                         bg='white', anchor='w').pack(side='left', fill='x', expand=True)
                tk.Label(row, text=f'💊 {cnt}', font=('Arial', 10, 'bold'),
                         bg='white', fg='#c0392b', width=8,
                         anchor='e').pack(side='right')
        else:
            tk.Label(med_content, text='📭 No data',
                     font=('Arial', 10, 'italic'),
                     bg='white', fg='#95a5a6').pack(pady=20)

        # ===== ذخیره برای Export/Print =====
        self.current_report_data = {
            'month': month_str,
            'from_date': from_date,
            'to_date': to_date,
            'total_revenue': total_revenue,
            'total_paid': total_paid,
            'total_pending': total_pending,
            'payment_count': payment_count,
            'avg_per_day': avg_per_day,
            'unique_patients': unique_patients,
            'total_visits': total_visits,
            'total_prescriptions': total_prescriptions,
            'total_lab': total_lab,
            'total_appts': total_appts,
            'completed_appts': completed_appts,
            'cancelled_appts': cancelled_appts,
            'top_doctors': top_doctors,
            'top_medicines': top_medicines,
        }


    def export_monthly_report_excel(self):
        """Export گزارش ماهانه به Excel"""
        if not hasattr(self, 'current_report_data'):
            messagebox.showwarning('Warning', '⚠️ Generate a report first!')
            return

        from tkinter import filedialog
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        data = self.current_report_data

        file_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[('Excel files', '*.xlsx')],
            initialfile=f'monthly_report_{data["from_date"]}.xlsx'
        )

        if not file_path:
            return

        try:
            wb = Workbook()

            # ===== Sheet 1: Summary =====
            ws = wb.active
            ws.title = 'Summary'

            # عنوان
            ws.merge_cells('A1:C1')
            title = ws.cell(row=1, column=1, value=f'Monthly Report - {data["month"]}')
            title.font = Font(size=16, bold=True, color='FFFFFF')
            title.fill = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')
            title.alignment = Alignment(horizontal='center', vertical='center')
            ws.row_dimensions[1].height = 30

            # بازه
            ws.merge_cells('A2:C2')
            ws.cell(row=2, column=1, value=f'Period: {data["from_date"]} to {data["to_date"]}')
            ws.cell(row=2, column=1).alignment = Alignment(horizontal='center')

            # ===== بخش مالی =====
            current_row = 4
            ws.cell(row=current_row, column=1, value='💰 FINANCIAL SUMMARY')
            ws.cell(row=current_row, column=1).font = Font(size=12, bold=True, color='FFFFFF')
            ws.cell(row=current_row, column=1).fill = PatternFill(start_color='27AE60', end_color='27AE60',
                                                                  fill_type='solid')
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)

            current_row += 1
            financial_data = [
                ('Total Revenue', f'{data["total_revenue"]:,.0f} IRR'),
                ('Paid', f'{data["total_paid"]:,.0f} IRR'),
                ('Pending', f'{data["total_pending"]:,.0f} IRR'),
                ('Average per Day', f'{data["avg_per_day"]:,.0f} IRR'),
                ('Payment Count', data['payment_count']),
            ]

            for label, value in financial_data:
                ws.cell(row=current_row, column=1, value=label).font = Font(bold=True)
                ws.cell(row=current_row, column=2, value=value)
                current_row += 1

            # ===== بخش فعالیت =====
            current_row += 1
            ws.cell(row=current_row, column=1, value='📊 ACTIVITY SUMMARY')
            ws.cell(row=current_row, column=1).font = Font(size=12, bold=True, color='FFFFFF')
            ws.cell(row=current_row, column=1).fill = PatternFill(start_color='3498DB', end_color='3498DB',
                                                                  fill_type='solid')
            ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)

            current_row += 1
            activity_data = [
                ('Unique Patients', data['unique_patients']),
                ('Total Visits', data['total_visits']),
                ('Prescriptions', data['total_prescriptions']),
                ('Lab Tests', data['total_lab']),
                ('Appointments', data['total_appts']),
                ('Completed Appointments', data['completed_appts']),
                ('Cancelled Appointments', data['cancelled_appts']),
            ]

            for label, value in activity_data:
                ws.cell(row=current_row, column=1, value=label).font = Font(bold=True)
                ws.cell(row=current_row, column=2, value=value)
                current_row += 1

            # عرض ستون‌ها
            ws.column_dimensions['A'].width = 25
            ws.column_dimensions['B'].width = 30

            # ===== Sheet 2: Top Doctors =====
            ws2 = wb.create_sheet('Top Doctors')
            ws2.cell(row=1, column=1, value='Rank').font = Font(bold=True)
            ws2.cell(row=1, column=2, value='Doctor').font = Font(bold=True)
            ws2.cell(row=1, column=3, value='Visits').font = Font(bold=True)

            for i, (name, cnt) in enumerate(data['top_doctors'], 1):
                ws2.cell(row=i + 1, column=1, value=i)
                ws2.cell(row=i + 1, column=2, value=name)
                ws2.cell(row=i + 1, column=3, value=cnt)

            ws2.column_dimensions['A'].width = 10
            ws2.column_dimensions['B'].width = 30
            ws2.column_dimensions['C'].width = 15

            # ===== Sheet 3: Top Medicines =====
            ws3 = wb.create_sheet('Top Medicines')
            ws3.cell(row=1, column=1, value='Rank').font = Font(bold=True)
            ws3.cell(row=1, column=2, value='Medicine').font = Font(bold=True)
            ws3.cell(row=1, column=3, value='Prescribed').font = Font(bold=True)

            for i, (name, cnt) in enumerate(data['top_medicines'], 1):
                ws3.cell(row=i + 1, column=1, value=i)
                ws3.cell(row=i + 1, column=2, value=name)
                ws3.cell(row=i + 1, column=3, value=cnt)

            ws3.column_dimensions['A'].width = 10
            ws3.column_dimensions['B'].width = 30
            ws3.column_dimensions['C'].width = 15

            wb.save(file_path)

            messagebox.showinfo('✅ Export Successful',
                                f'Excel file saved to:\n{file_path}')

            try:
                os.startfile(file_path)
            except:
                pass

        except Exception as e:
            messagebox.showerror('Export Error', f'❌ {e}')

    def print_monthly_report(self):
        """پرینت گزارش ماهانه"""
        if not hasattr(self, 'current_report_data'):
            messagebox.showwarning('Warning', '⚠️ Generate a report first!')
            return

        data = self.current_report_data

        lines = []
        lines.append('=' * 70)
        lines.append(' ' * 20 + 'HOSPITAL MANAGEMENT SYSTEM')
        lines.append(' ' * 22 + 'MONTHLY REPORT')
        lines.append('=' * 70)
        lines.append('')
        lines.append(f'  Month: {data["month"]}')
        lines.append(f'  Period: {data["from_date"]} to {data["to_date"]}')
        lines.append(f'  Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        lines.append('')

        # Financial
        lines.append('-' * 70)
        lines.append('  💰 FINANCIAL SUMMARY')
        lines.append('-' * 70)
        lines.append(f'  Total Revenue:    {data["total_revenue"]:>20,.0f} IRR')
        lines.append(f'  Paid:             {data["total_paid"]:>20,.0f} IRR')
        lines.append(f'  Pending:          {data["total_pending"]:>20,.0f} IRR')
        lines.append(f'  Average per Day:  {data["avg_per_day"]:>20,.0f} IRR')
        lines.append(f'  Payment Count:    {data["payment_count"]:>20}')
        lines.append('')

        # Activity
        lines.append('-' * 70)
        lines.append('  📊 ACTIVITY SUMMARY')
        lines.append('-' * 70)
        lines.append(f'  Unique Patients:       {data["unique_patients"]:>10}')
        lines.append(f'  Total Visits:          {data["total_visits"]:>10}')
        lines.append(f'  Prescriptions:         {data["total_prescriptions"]:>10}')
        lines.append(f'  Lab Tests:             {data["total_lab"]:>10}')
        lines.append(f'  Total Appointments:    {data["total_appts"]:>10}')
        lines.append(f'  Completed:             {data["completed_appts"]:>10}')
        lines.append(f'  Cancelled:             {data["cancelled_appts"]:>10}')
        lines.append('')

        # Top Doctors
        if data['top_doctors']:
            lines.append('-' * 70)
            lines.append('  🏆 TOP 5 DOCTORS')
            lines.append('-' * 70)
            for i, (name, cnt) in enumerate(data['top_doctors'], 1):
                doctor_clean = name.replace('Dr.', '').replace('dr.', '').strip()
                lines.append(f'  {i}. Dr. {doctor_clean:<30} {cnt:>5} visits')
            lines.append('')

        # Top Medicines
        if data['top_medicines']:
            lines.append('-' * 70)
            lines.append('  🏆 TOP 5 MEDICINES')
            lines.append('-' * 70)
            for i, (name, cnt) in enumerate(data['top_medicines'], 1):
                lines.append(f'  {i}. {name:<35} {cnt:>5} times')
            lines.append('')

        lines.append('=' * 70)
        lines.append('  END OF REPORT')
        lines.append('=' * 70)

        content = '\n'.join(lines)

        success, pdf_path = print_text_file(
            content,
            f'monthly_report_{data["from_date"]}.txt',
            open_pdf=False
        )

        if not success:
            messagebox.showerror('❌ PDF Error', f'Could not create PDF:\n{pdf_path}')
            return

        if WIN32_AVAILABLE:
            printer = get_printer_for('report')
            virtual_keywords = ['Print to PDF', 'XPS', 'OneNote', 'Fax', 'Virtual', 'Writer']
            is_virtual = any(kw.lower() in (printer or '').lower() for kw in virtual_keywords)

            if is_virtual or not printer:
                try:
                    os.startfile(pdf_path)
                except:
                    pass
                messagebox.showinfo('📄 PDF Ready',
                                    f'PDF created for printing.\n\n'
                                    f'File: {pdf_path}')
            else:
                ok, msg = print_pdf_to_printer(pdf_path, printer)
                if ok:
                    messagebox.showinfo('🖨️ Sent to Printer',
                                        f'Report sent to:\n\n🖨️ {printer}')
                else:
                    try:
                        os.startfile(pdf_path)
                    except:
                        pass
                    messagebox.showwarning('⚠️ Could Not Print',
                                           f'Error: {msg}\n\nPDF opened.')
        else:
            try:
                os.startfile(pdf_path)
            except:
                pass



    def create_printer_settings_tab(self, parent):
        """تب تنظیمات پرینتر"""
        # ===== هدر =====
        header = tk.Frame(parent, bg='#ecf0f1', height=50)
        header.pack(fill='x')
        header.pack_propagate(False)
        tk.Label(header, text='🖨️ Printer Configuration',
                 font=('Arial', 13, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(pady=12)

        # ===== توضیحات =====
        info = tk.Label(parent,
                        text='💡 Select which printer to use for each type of document.\n'
                             'Leave empty to use Windows default printer.',
                        font=('Arial', 9, 'italic'), bg='#f8f9fa',
                        fg='#7f8c8d', justify='left')
        info.pack(pady=(0, 15), padx=20, anchor='w')

        # ===== لیست پرینترها =====
        available_printers = get_available_printers()

        if not available_printers:
            tk.Label(parent,
                     text='⚠️ No printers found on this system.\n\n'
                          'Please install a printer first via:\n'
                          'Windows → Settings → Devices → Printers & Scanners',
                     font=('Arial', 11), bg='#f8f9fa',
                     fg='#e74c3c', justify='center').pack(pady=40)
            return

        # ===== Load settings فعلی =====
        settings = load_settings()
        current_printers = settings.get('printers', {})

        # ===== فرم =====
        form_frame = tk.Frame(parent, bg='white',
                              relief='solid', bd=1)
        form_frame.pack(fill='x', padx=20, pady=10)

        # هدر فرم
        hdr = tk.Frame(form_frame, bg='#3498db')
        hdr.pack(fill='x')

        tk.Label(hdr, text='  📄 Document Type',
                 font=('Arial', 10, 'bold'), bg='#3498db',
                 fg='white', width=30, anchor='w').pack(side='left', padx=5, pady=8)
        tk.Label(hdr, text='🖨️ Printer',
                 font=('Arial', 10, 'bold'), bg='#3498db',
                 fg='white', anchor='w').pack(side='left', padx=5, pady=8)

        # ===== ردیف‌ها =====
        printer_vars = {}

        printer_configs = [
            ('default', '⭐ Default (پیش‌فرض)',
             'Used when no specific printer is set'),
            ('prescription', '💊 Prescription (نسخه‌ها)',
             'For medical prescriptions'),
            ('invoice', '💰 Invoice (فاکتورها)',
             'For pharmacy invoices and receipts'),
            ('report', '📊 Report (گزارش‌ها)',
             'For statistical reports'),
            ('receipt', '🧾 Receipt (رسیدها)',
             'For payment receipts'),
        ]

        # گزینه‌ی "None" به ابتدای لیست
        printer_options = ['(use Windows default)'] + available_printers

        for key, label, hint in printer_configs:
            row = tk.Frame(form_frame, bg='white')
            row.pack(fill='x', pady=1)

            # سمت چپ: label
            left = tk.Frame(row, bg='white', width=280)
            left.pack(side='left', padx=5, pady=8, fill='x')
            left.pack_propagate(False)

            tk.Label(left, text=label, font=('Arial', 10, 'bold'),
                     bg='white', anchor='w').pack(anchor='w')
            tk.Label(left, text=hint, font=('Arial', 8, 'italic'),
                     bg='white', fg='#95a5a6',
                     anchor='w').pack(anchor='w')

            # سمت راست: combo
            var = tk.StringVar()
            current_value = current_printers.get(key, '')

            if current_value and current_value in available_printers:
                var.set(current_value)
            else:
                var.set('(use Windows default)')

            printer_vars[key] = var

            combo = ttk.Combobox(row, textvariable=var,
                                 values=printer_options,
                                 state='readonly', width=40, font=('Arial', 10))
            combo.pack(side='left', padx=10, pady=8)

            # خط جداکننده
            tk.Frame(form_frame, bg='#ecf0f1', height=1).pack(fill='x')

        # ===== وضعیت =====
        status_var = tk.StringVar(value='')
        status_label = tk.Label(parent, textvariable=status_var,
                                font=('Arial', 10, 'italic'),
                                bg='#f8f9fa', fg='#27ae60')
        status_label.pack(pady=10)

        def show_status(msg, color='#27ae60', duration=3000):
            status_var.set(msg)
            status_label.config(fg=color)
            parent.after(duration, lambda: status_var.set(''))

        # ===== ذخیره =====
        def save_printer_settings():
            try:
                new_printers = {}
                for key, var in printer_vars.items():
                    value = var.get()
                    if value == '(use Windows default)':
                        new_printers[key] = ''
                    else:
                        new_printers[key] = value

                settings = load_settings()
                settings['printers'] = new_printers
                save_settings(settings)

                show_status('✅ Printer settings saved successfully!', '#27ae60')
                print(f'✅ Settings saved: {new_printers}')

                # لاگ
                log_action(
                    'UPDATE',
                    entity_type='settings',
                    details=f'Printer settings updated',
                    user=getattr(self.root, 'current_user', None)
                )

            except Exception as e:
                show_status(f'❌ Save failed: {e}', '#e74c3c')

        # ===== تست پرینتر =====
        def test_printer():
            try:
                # گرفتن پرینتر default
                test_printer_name = printer_vars['default'].get()
                if test_printer_name == '(use Windows default)':
                    test_printer_name = get_default_printer()

                if not test_printer_name:
                    show_status('❌ No printer available for test', '#e74c3c')
                    return

                # ساخت یه PDF ساده برای تست
                test_content = (
                        '=' * 50 + '\n'
                                   '     PRINTER TEST PAGE\n'
                                   '=' * 50 + '\n\n'
                                              'This is a test page to verify\n'
                                              'printer configuration.\n\n'
                                              f'Printer: {test_printer_name}\n'
                                              f'Time: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}\n\n'
                                              '=' * 50 + '\n'
                                                         '   If you see this, printer works!\n'
                                                         '=' * 50 + '\n'
                )

                success, result = print_text_file(
                    test_content, 'printer_test.txt', open_pdf=False
                )

                if not success:
                    show_status(f'❌ PDF error: {result}', '#e74c3c')
                    return

                # پرینت
                ok, msg = print_pdf_to_printer(result, test_printer_name)

                if ok:
                    show_status(f'✅ Test sent to: {test_printer_name}', '#27ae60')
                else:
                    show_status(f'❌ Print failed: {msg}', '#e74c3c')

            except Exception as e:
                show_status(f'❌ Test error: {e}', '#e74c3c')

        # ===== دکمه‌ها =====
        btn_frame = tk.Frame(parent, bg='#f8f9fa')
        btn_frame.pack(pady=15)

        create_popup_buttons(btn_frame, [
            {'text': '💾 Save Settings', 'command': save_printer_settings,
             'color': 'success', 'width': 200},
            {'text': '🖨️ Test Printer', 'command': test_printer,
             'color': 'primary', 'width': 180},
            {'text': '🔄 Reload',
             'command': lambda: [w.destroy() for w in parent.winfo_children()] or self.create_printer_settings_tab(
                 parent),
             'color': 'warning', 'width': 140},
        ], bg='#f8f9fa')

        # ===== راهنما =====
        help_frame = tk.Frame(parent, bg='#fff3cd',
                              relief='solid', bd=1)
        help_frame.pack(fill='x', padx=20, pady=(10, 20))

        tk.Label(help_frame, text='💡 How it works:',
                 font=('Arial', 10, 'bold'), bg='#fff3cd',
                 fg='#856404', anchor='w').pack(anchor='w', padx=15, pady=(8, 3))

        help_text = (
            '• If you set a printer for a specific document type, that document will always go to that printer\n'
            '• If you leave it empty, the Default printer will be used\n'
            '• If Default is also empty, Windows default printer is used\n'
            '• Click "Test Printer" to verify your configuration'
        )

        tk.Label(help_frame, text=help_text,
                 font=('Arial', 9), bg='#fff3cd',
                 fg='#856404', justify='left', anchor='w').pack(
            anchor='w', padx=15, pady=(0, 8))

    def create_appearance_tab(self, parent):
        """تب ظاهر (Dark Mode)"""
        # ===== هدر =====
        header = tk.Frame(parent, bg='#ecf0f1', height=50)
        header.pack(fill='x')
        header.pack_propagate(False)
        tk.Label(header, text='🎨 Appearance Settings',
                 font=('Arial', 13, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(pady=12)

        # ===== توضیحات =====
        tk.Label(parent,
                 text='💡 Choose your preferred theme.\n'
                      'Dark mode is easier on the eyes at night.',
                 font=('Arial', 9, 'italic'), bg='#f8f9fa',
                 fg='#7f8c8d', justify='left').pack(pady=(0, 20), padx=20, anchor='w')

        # ===== کارت‌های تم =====
        cards_frame = tk.Frame(parent, bg='#f8f9fa')
        cards_frame.pack(pady=20)

        settings = load_settings()
        current_mode = 'dark' if settings.get('dark_mode', False) else 'light'

        def make_theme_card(parent, icon, title, description, mode, bg_color, fg_color):
            card = tk.Frame(parent, bg='white', relief='solid', bd=2,
                            highlightbackground='#3498db' if current_mode == mode else '#bdc3c7',
                            highlightthickness=3 if current_mode == mode else 1)
            card.pack(side='left', padx=10)

            # پیش‌نمایش
            preview = tk.Frame(card, bg=bg_color, width=180, height=100)
            preview.pack(pady=(10, 5))
            preview.pack_propagate(False)

            tk.Label(preview, text=icon, font=('Arial', 30),
                     bg=bg_color, fg=fg_color).pack(pady=(15, 0))
            tk.Label(preview, text=title, font=('Arial', 10, 'bold'),
                     bg=bg_color, fg=fg_color).pack()

            # عنوان و توضیح
            tk.Label(card, text=title, font=('Arial', 11, 'bold'),
                     bg='white', fg='#2c3e50').pack(pady=(10, 2))
            tk.Label(card, text=description, font=('Arial', 8, 'italic'),
                     bg='white', fg='#7f8c8d').pack()

            # دکمه‌ی انتخاب
            btn_text = '✅ Active' if current_mode == mode else 'Apply'
            btn_color = '#27ae60' if current_mode == mode else '#3498db'

            def apply_theme():
                apply_and_restart(mode)

            tk.Button(card, text=btn_text, command=apply_theme,
                      bg=btn_color, fg='white', font=('Arial', 10, 'bold'),
                      relief='flat', cursor='hand2',
                      padx=20, pady=6).pack(pady=10)

        make_theme_card(cards_frame, '☀️', 'Light Mode',
                        'Classic bright appearance',
                        'light', '#f5f7fa', '#2c3e50')

        make_theme_card(cards_frame, '🌙', 'Dark Mode',
                        'Easier on the eyes',
                        'dark', '#1e1e2e', '#e0e0e0')

        # ===== توضیح =====
        info_frame = tk.Frame(parent, bg='#fff3cd', relief='solid', bd=1)
        info_frame.pack(fill='x', padx=20, pady=20)

        tk.Label(info_frame,
                 text='⚠️ After changing theme, the application will restart.\n'
                      'Make sure to save your work before applying.',
                 font=('Arial', 9), bg='#fff3cd',
                 fg='#856404', justify='left').pack(pady=10, padx=15, anchor='w')

        def apply_and_restart(mode):
            try:
                settings = load_settings()
                settings['dark_mode'] = (mode == 'dark')
                save_settings(settings)

                # ===== اعمال تم بدون بستن برنامه =====
                self.apply_theme_live(mode)

            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

    def apply_theme_live(self, mode):
        """تغییر تم بدون بستن برنامه"""
        try:
            # ===== ۱. به‌روزرسانی متغیرهای سراسری =====
            global CURRENT_THEME_MODE, CURRENT_THEME, WINDOW_BG, TEXT_FG, HEADER_BG
            CURRENT_THEME_MODE = mode
            CURRENT_THEME = get_theme(mode)
            WINDOW_BG = CURRENT_THEME['bg']
            TEXT_FG = CURRENT_THEME['fg']
            HEADER_BG = CURRENT_THEME['header_bg']

            # ===== ۲. پاک کردن همه‌ی widget ها =====
            for widget in self.root.winfo_children():
                widget.destroy()

            # ===== ۳. پس‌زمینه‌ی اصلی =====
            self.root.configure(bg=WINDOW_BG)

            # ===== ۴. بازسازی همه‌ی widget ها =====
            self.create_widgets()

            # ===== ۵. اعمال استایل تیره به ttk =====
            try:
                apply_dark_style()
            except Exception as e:
                print(f'⚠️ Dark style error: {e}')

            # ===== ۶. پیام موفقیت =====
            messagebox.showinfo('🎨 Theme Changed',
                                f'Theme changed to {mode.upper()} mode successfully!')

        except Exception as e:
            messagebox.showerror('Error', f'❌ Theme change failed:\n{e}')
            print(f'❌ Theme error: {e}')
            import traceback
            traceback.print_exc()



    def create_archive_tab(self, parent):
        """تب ساده‌ی Archive"""
        # ===== Scrollable Canvas =====
        canvas = tk.Canvas(parent, bg='#f8f9fa', highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f8f9fa')

        scrollable.bind('<Configure>',
                        lambda e: canvas.configure(scrollregion=canvas.bbox('all')))

        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>',
                    lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== هدر =====
        header = tk.Frame(scrollable, bg='#ecf0f1', height=50)
        header.pack(fill='x')
        header.pack_propagate(False)
        tk.Label(header, text='🗂️ Archive Management',
                 font=('Arial', 13, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(pady=12)

        # ===== آمار ساده =====
        stats_frame = tk.Frame(scrollable, bg='#f8f9fa')
        stats_frame.pack(fill='x', padx=20, pady=20)

        try:
            conn = get_connection()
            cursor = conn.cursor()

            cursor.execute('SELECT COUNT(*) FROM medical_records')
            active_records = cursor.fetchone()[0]

            cursor.execute('SELECT COUNT(*) FROM prescriptions')
            active_prescriptions = cursor.fetchone()[0]

            import os
            db_size_mb = 0
            if os.path.exists(DB_NAME):
                db_size_mb = os.path.getsize(DB_NAME) / (1024 * 1024)

            conn.close()

            # کارت‌های آمار
            cards_row = tk.Frame(stats_frame, bg='#f8f9fa')
            cards_row.pack(fill='x')

            def make_card(parent, icon, label, value, color, subtitle=''):
                card = tk.Frame(parent, bg=color, height=140)  # ← ارتفاع بیشتر
                card.pack(side='left', padx=8, fill='x', expand=True)
                card.pack_propagate(False)

                tk.Label(card, text=icon, font=('Arial', 24),
                         bg=color, fg='white').pack(pady=(15, 5))
                tk.Label(card, text=str(value), font=('Arial', 22, 'bold'),
                         bg=color, fg='white').pack()
                tk.Label(card, text=label, font=('Arial', 10),
                         bg=color, fg='white').pack(pady=(5, 0))
                if subtitle:
                    tk.Label(card, text=subtitle, font=('Arial', 8, 'italic'),
                             bg=color, fg='#f0f0f0').pack(pady=(2, 10))

            make_card(cards_row, '📋', 'Medical Records',
                      active_records, '#3498db',
                      f'{active_prescriptions} prescriptions')
            make_card(cards_row, '💾', 'Database Size',
                      f'{db_size_mb:.1f}', '#27ae60', 'MB')
            make_card(cards_row, '📅', 'Today',
                      datetime.now().strftime('%Y-%m-%d'), '#e67e22')

        except Exception as e:
            tk.Label(stats_frame, text=f'⚠️ Error: {e}',
                     font=('Arial', 10), bg='#f8f9fa',
                     fg='#e74c3c').pack(pady=20)

        # ===== راهنما =====
        guide_frame = tk.LabelFrame(scrollable, text=' 💡 How to Archive ',
                                    font=('Arial', 12, 'bold'),
                                    bg='#f8f9fa', fg='#2c3e50',
                                    padx=20, pady=15)
        guide_frame.pack(fill='x', padx=20, pady=20)

        guide_text = (
            'Archive is handled through Backup Management.\n\n'
            '📌 Steps to archive your data:\n\n'
            '1. Go to Dashboard\n'
            '2. Click "📂 Backups"\n'
            '3. Select a backup from "Daily Backups"\n'
            '4. Click "📌 Archive"\n'
            '5. Find it in "Archived (Permanent)" tab\n\n'
            '✅ Archived backups are NEVER auto-deleted\n'
            '✅ You can restore anytime\n'
            '✅ Safe and simple'
        )

        tk.Label(guide_frame, text=guide_text,
                 font=('Arial', 10),
                 bg='#f8f9fa', fg='#2c3e50',
                 justify='left', anchor='w').pack(anchor='w', fill='x')

        # ===== دکمه‌ها =====
        btn_frame = tk.Frame(scrollable, bg='#f8f9fa')
        btn_frame.pack(pady=20)

        def open_backup_management():
            self.show_backup_list()

        tk.Button(btn_frame, text='📂 Open Backup Management',
                  command=open_backup_management,
                  bg='#3498db', fg='white',
                  font=('Arial', 11, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=20, pady=10).pack(side='left', padx=5)

        tk.Button(btn_frame, text='💾 Create Backup Now',
                  command=self.do_backup,
                  bg='#27ae60', fg='white',
                  font=('Arial', 11, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=20, pady=10).pack(side='left', padx=5)

    def load_users(self):
        """بارگذاری کاربران با فیلتر"""
        search_text = ''
        if hasattr(self, 'user_search_var'):
            search_text = self.user_search_var.get().strip()

        conn = get_connection()
        cursor = conn.cursor()

        if search_text:
            pattern = f'%{search_text}%'
            cursor.execute('''
                SELECT id, username, full_name, role, is_active, last_login
                FROM users
                WHERE username LIKE ?
                    OR full_name LIKE ?
                    OR role LIKE ?
                ORDER BY id
            ''', (pattern, pattern, pattern))
        else:
            cursor.execute('''
                SELECT id, username, full_name, role, is_active, last_login
                FROM users ORDER BY id
            ''')

        users = cursor.fetchall()
        conn.close()

        for item in self.users_tree.get_children():
            self.users_tree.delete(item)

        for u in users:
            real_id = u[0]
            active_icon = '✅ Active' if u[4] == 1 else '❌ Disabled'
            last_login = u[5][:16] if u[5] else 'Never'

            self.users_tree.insert(
                '', 'end',
                values=(u[0], u[1], u[2], u[3].upper(), active_icon, last_login),
                tags=(str(real_id),)
            )

        # آپدیت شمارنده
        if hasattr(self, 'user_count_var'):
            if search_text:
                self.user_count_var.set(f'🔍 {len(users)} result(s)')
            else:
                self.user_count_var.set(f'👤 {len(users)} user(s)')

    def add_user_window(self):
        """فرم افزودن کاربر جدید"""
        win = tk.Toplevel(self.root)
        win.title('➕ Add New User')
        win.configure(bg='#f0f0f0')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#f0f0f0')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#f0f0f0')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f0f0f0')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text='➕ Add New User',
                 font=('Arial', 16, 'bold'), bg='#f0f0f0',
                 fg='#2c3e50').pack(pady=15)

        # ===== ۴. فرم =====
        form_frame = tk.Frame(scrollable, bg='#f0f0f0')
        form_frame.pack(pady=10, padx=40, fill='x')

        fields = {}

        # Username
        tk.Label(form_frame, text='Username *', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=0, column=0, sticky='w', pady=(10, 5), padx=5)
        fields['username'] = tk.Entry(form_frame, width=40, font=('Arial', 10))
        fields['username'].grid(row=1, column=0, pady=(0, 10), padx=5, sticky='ew')
        tk.Label(form_frame, text='💡 English letters/numbers only',
                 bg='#f0f0f0', font=('Arial', 8), fg='#7f8c8d').grid(
            row=1, column=1, sticky='w', padx=5)

        # Full Name
        tk.Label(form_frame, text='Full Name *', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=2, column=0, sticky='w', pady=(10, 5), padx=5)
        fields['full_name'] = tk.Entry(form_frame, width=40, font=('Arial', 10))
        fields['full_name'].grid(row=3, column=0, pady=(0, 10), padx=5, sticky='ew')
        tk.Label(form_frame, text='💡 e.g., Dr. Ahmad Tehrani',
                 bg='#f0f0f0', font=('Arial', 8), fg='#7f8c8d').grid(
            row=3, column=1, sticky='w', padx=5)

        # Password
        tk.Label(form_frame, text='Password *', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=4, column=0, sticky='w', pady=(10, 5), padx=5)
        fields['password'] = tk.Entry(form_frame, width=40, font=('Arial', 10), show='●')
        fields['password'].grid(row=5, column=0, pady=(0, 10), padx=5, sticky='ew')
        tk.Label(form_frame, text='💡 Minimum 6 characters',
                 bg='#f0f0f0', font=('Arial', 8), fg='#7f8c8d').grid(
            row=5, column=1, sticky='w', padx=5)

        # Confirm Password
        tk.Label(form_frame, text='Confirm Password *', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=6, column=0, sticky='w', pady=(10, 5), padx=5)
        fields['confirm'] = tk.Entry(form_frame, width=40, font=('Arial', 10), show='●')
        fields['confirm'].grid(row=7, column=0, pady=(0, 10), padx=5, sticky='ew')

        # Role
        tk.Label(form_frame, text='Role *', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=8, column=0, sticky='w', pady=(10, 5), padx=5)

        roles = ['admin', 'doctor', 'receptionist', 'pharmacist']
        role_var = tk.StringVar(value='doctor')
        ttk.Combobox(form_frame, textvariable=role_var, values=roles,
                     width=38, state='readonly').grid(
            row=9, column=0, pady=(0, 10), padx=5, sticky='ew')

        form_frame.grid_columnconfigure(0, weight=1)

        # ===== ۵. تابع ذخیره =====
        def save():
            try:
                username = fields['username'].get().strip()
                full_name = fields['full_name'].get().strip()
                password = fields['password'].get()
                confirm = fields['confirm'].get()

                if not username:
                    messagebox.showerror('Error', '❌ Username is required!')
                    return
                if len(username) < 3:
                    messagebox.showerror('Error', '❌ Username must be at least 3 characters!')
                    return
                if not username.replace('_', '').replace('.', '').isalnum():
                    messagebox.showerror('Error', '❌ Username can only contain letters, numbers, _ and .')
                    return
                if not full_name:
                    messagebox.showerror('Error', '❌ Full Name is required!')
                    return
                if len(password) < 6:
                    messagebox.showerror('Error', '❌ Password must be at least 6 characters!')
                    return
                if password != confirm:
                    messagebox.showerror('Error', '❌ Passwords do not match!')
                    return

                role = role_var.get()

                import hashlib
                pwd_hash = hashlib.sha256(password.encode()).hexdigest()

                import secrets
                recovery_code = secrets.token_hex(4).upper()

                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute('''
                    INSERT INTO users (username, password_hash, full_name, role, recovery_code)
                    VALUES (?, ?, ?, ?, ?)
                ''', (username, pwd_hash, full_name, role, recovery_code))
                conn.commit()
                conn.close()

                messagebox.showinfo('✅ User Created',
                                    f'✅ User "{username}" created!\n\n'
                                    f'👤 Role: {role}\n'
                                    f'👤 Full Name: {full_name}\n'
                                    f'🔑 Password: {password}\n'
                                    f'🆘 Recovery Code: {recovery_code}\n\n'
                                    f'⚠️ Save these credentials!')

                try:
                    log_action('CREATE', entity_type='user',
                               details=f'Created user "{username}" ({role})',
                               user=getattr(self.root, 'current_user', None))
                except:
                    pass

                win.destroy()
                self.load_users()

            except sqlite3.IntegrityError:
                messagebox.showerror('Error', f'❌ Username "{username}" already exists!')
            except Exception as e:
                messagebox.showerror('Error', f'❌ Error: {e}')

        # ===== ۶. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '💾 Create User', 'command': save, 'color': 'success', 'width': 180},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#f0f0f0')

        # ===== ۷. Maximize =====
        win.update_idletasks()
        win.state('zoomed')

        fields['username'].focus()


    def show_credentials_window(self, username, password, full_name, role, recovery_code):
        """پنجره‌ی نمایش Credentials با قابلیت کپی"""
        win = tk.Toplevel(self.root)
        win.title('🔐 User Credentials')
        win.geometry('600x650')
        win.configure(bg='white')
        win.resizable(False, False)

        win.update_idletasks()
        w, h = 600, 650
        x = (win.winfo_screenwidth() // 2) - (w // 2)
        y = (win.winfo_screenheight() // 2) - (h // 2)
        win.geometry(f'{w}x{h}+{x}+{y}')

        # ===== هدر =====
        header = tk.Frame(win, bg='#27ae60', height=60)
        header.pack(fill='x')
        header.pack_propagate(False)
        tk.Label(header, text='✅ User Created Successfully!',
                 font=('Arial', 14, 'bold'), bg='#27ae60',
                 fg='white').pack(pady=15)

        # ===== هشدار =====
        warning_frame = tk.Frame(win, bg='#fff3cd', relief='solid', bd=1)
        warning_frame.pack(fill='x', padx=15, pady=10)
        tk.Label(warning_frame, text='⚠️ IMPORTANT: Copy these credentials now!',
                 font=('Arial', 11, 'bold'), bg='#fff3cd',
                 fg='#856404').pack(pady=8)

        # ===== متن Credentials =====
        text_frame = tk.Frame(win, bg='#f8f9fa', relief='solid', bd=2)
        text_frame.pack(fill='both', expand=True, padx=15, pady=10)

        credentials_text = f"""👤 Full Name: {full_name}
    🎭 Role: {role.upper()}

    ━━━━━━━━━━━━━━━━━━━━━━━━━━

    🔑 Username: {username}
    🔒 Password: {password}

    🆘 Recovery Code: {recovery_code}

    ━━━━━━━━━━━━━━━━━━━━━━━━━━

    ⚠️ Recovery Code is CRITICAL!
    Keep it safe. You'll need it to reset
    the password if forgotten.
    """

        text_widget = tk.Text(text_frame, font=('Courier New', 11),
                              bg='#f8f9fa', fg='#2c3e50',
                              padx=20, pady=20, relief='flat',
                              wrap='word', height=14)
        text_widget.pack(fill='both', expand=True, padx=5, pady=5)
        text_widget.insert('1.0', credentials_text)
        text_widget.config(state='disabled')

        # ===== Status =====
        status_var = tk.StringVar(value='')
        status_label = tk.Label(win, textvariable=status_var,
                                font=('Arial', 10, 'italic'),
                                bg='white', fg='#27ae60')
        status_label.pack(pady=5)

        def show_status(msg, color='#27ae60', duration=2500):
            status_var.set(msg)
            status_label.config(fg=color)
            win.after(duration, lambda: status_var.set(''))

        # ===== کپی کامل =====
        def copy_all():
            full_text = (
                f'🔐 Account Credentials\n\n'
                f'Full Name: {full_name}\n'
                f'Role: {role.upper()}\n\n'
                f'Username: {username}\n'
                f'Password: {password}\n'
                f'Recovery Code: {recovery_code}\n\n'
                f'⚠️ Keep Recovery Code safe!'
            )
            try:
                win.clipboard_clear()
                win.clipboard_append(full_text)
                show_status('✅ All credentials copied to clipboard!')
            except Exception as e:
                show_status(f'❌ Copy failed: {e}', '#e74c3c')

        # ===== کپی فقط رمز =====
        def copy_password():
            try:
                win.clipboard_clear()
                win.clipboard_append(password)
                show_status('✅ Password copied!')
            except Exception as e:
                show_status(f'❌ Copy failed: {e}', '#e74c3c')

        # ===== کپی فقط Recovery Code =====
        def copy_recovery():
            try:
                win.clipboard_clear()
                win.clipboard_append(recovery_code)
                show_status('✅ Recovery code copied!')
            except Exception as e:
                show_status(f'❌ Copy failed: {e}', '#e74c3c')

        # ===== دکمه‌ها =====
        btn_frame = tk.Frame(win, bg='white')
        btn_frame.pack(fill='x', padx=15, pady=15)

        # ردیف اول: کپی
        row1 = tk.Frame(btn_frame, bg='white')
        row1.pack(fill='x', pady=5)

        tk.Button(row1, text='📋 Copy All',
                  command=copy_all,
                  bg='#3498db', fg='white',
                  font=('Arial', 11, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=20, pady=10).pack(side='left', padx=3, fill='x', expand=True)

        tk.Button(row1, text='🔒 Copy Password',
                  command=copy_password,
                  bg='#9b59b6', fg='white',
                  font=('Arial', 10, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=15, pady=10).pack(side='left', padx=3, fill='x', expand=True)

        tk.Button(row1, text='🆘 Copy Code',
                  command=copy_recovery,
                  bg='#e67e22', fg='white',
                  font=('Arial', 10, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=15, pady=10).pack(side='left', padx=3, fill='x', expand=True)

        # ردیف دوم: بستن
        row2 = tk.Frame(btn_frame, bg='white')
        row2.pack(fill='x', pady=(10, 0))

        tk.Button(row2, text='✅ Done',
                  command=win.destroy,
                  bg='#27ae60', fg='white',
                  font=('Arial', 12, 'bold'),
                  relief='flat', cursor='hand2',
                  padx=30, pady=10).pack(fill='x')

        # ===== کپی خودکار به کلیپ‌بورد (اولین بار) =====
        copy_all()

    def edit_user_window(self):
        """فرم ویرایش کاربر"""
        selected = self.users_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select a user first!')
            return

        real_id = self.users_tree.item(selected[0])['tags'][0]
        values = self.users_tree.item(selected[0])['values']
        current_role = values[3].lower()

        # چک کردن اینکه کاربر خودش رو ویرایش نکنه (اگه admin هست)
        if hasattr(self.root, 'current_user') and self.root.current_user:
            if str(self.root.current_user['id']) == str(real_id):
                if current_role == 'admin':
                    messagebox.showwarning('Warning', '⚠️ You cannot change your own role!')
                    return

        win = tk.Toplevel(self.root)
        win.title('✏️ Edit User')
        win.configure(bg='#f0f0f0')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#f0f0f0')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#f0f0f0')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f0f0f0')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text='✏️ Edit User',
                 font=('Arial', 16, 'bold'), bg='#f0f0f0',
                 fg='#2c3e50').pack(pady=15)

        tk.Label(scrollable, text=f'Username: {values[1]}',
                 font=('Arial', 11, 'bold'), bg='#f0f0f0', fg='#7f8c8d').pack(pady=5)

        # ===== ۴. فرم =====
        form_frame = tk.Frame(scrollable, bg='#f0f0f0')
        form_frame.pack(pady=10, padx=40, fill='x')

        # Full Name
        tk.Label(form_frame, text='Full Name *', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=0, column=0, sticky='w', pady=(10, 5), padx=5)
        full_name_entry = tk.Entry(form_frame, width=40, font=('Arial', 10))
        full_name_entry.grid(row=1, column=0, pady=(0, 15), padx=5, sticky='ew')
        full_name_entry.insert(0, values[2])

        # Role
        tk.Label(form_frame, text='Role *', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=2, column=0, sticky='w', pady=(10, 5), padx=5)

        roles = ['admin', 'doctor', 'receptionist', 'pharmacist']
        role_var = tk.StringVar(value=current_role)
        ttk.Combobox(form_frame, textvariable=role_var, values=roles,
                     width=38, state='readonly').grid(
            row=3, column=0, pady=(0, 15), padx=5, sticky='ew')

        form_frame.grid_columnconfigure(0, weight=1)

        # ===== ۵. تابع ذخیره =====
        def save():
            try:
                new_name = full_name_entry.get().strip()
                if not new_name:
                    messagebox.showerror('Error', '❌ Full Name is required!')
                    return

                new_role = role_var.get()

                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute('''
                    UPDATE users SET full_name = ?, role = ?
                    WHERE id = ?
                ''', (new_name, new_role, real_id))
                conn.commit()
                conn.close()

                messagebox.showinfo('Success', '✅ User updated!')
                try:
                    log_action('UPDATE', entity_type='user',
                               details=f'Updated user "{values[1]}": name={new_name}, role={new_role}',
                               user=getattr(self.root, 'current_user', None))
                except:
                    pass
                win.destroy()
                self.load_users()

            except Exception as e:
                messagebox.showerror('Error', f'❌ Error: {e}')

        # ===== ۶. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '💾 Save', 'command': save, 'color': 'success', 'width': 160},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#f0f0f0')

        # ===== ۷. Maximize =====
        win.update_idletasks()
        win.state('zoomed')

        full_name_entry.focus()

    def change_password_window(self):
        """فرم تغییر رمز عبور کاربر"""
        selected = self.users_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select a user first!')
            return

        real_id = self.users_tree.item(selected[0])['tags'][0]
        values = self.users_tree.item(selected[0])['values']
        username = values[1]

        win = tk.Toplevel(self.root)
        win.title('🔑 Change Password')
        win.configure(bg='#f0f0f0')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#f0f0f0')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#f0f0f0')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f0f0f0')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text='🔑 Change Password',
                 font=('Arial', 16, 'bold'), bg='#f0f0f0',
                 fg='#2c3e50').pack(pady=15)

        tk.Label(scrollable, text=f'User: {username}',
                 font=('Arial', 11, 'bold'), bg='#f0f0f0', fg='#7f8c8d').pack(pady=5)

        # ===== ۴. فرم =====
        form_frame = tk.Frame(scrollable, bg='#f0f0f0')
        form_frame.pack(pady=10, padx=40, fill='x')

        # New Password
        tk.Label(form_frame, text='New Password *', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=0, column=0, sticky='w', pady=(10, 5), padx=5)
        pwd_entry = tk.Entry(form_frame, width=40, font=('Arial', 10), show='●')
        pwd_entry.grid(row=1, column=0, pady=(0, 10), padx=5, sticky='ew')
        tk.Label(form_frame, text='💡 Minimum 6 characters',
                 bg='#f0f0f0', font=('Arial', 8), fg='#7f8c8d').grid(
            row=1, column=1, sticky='w', padx=5)

        # Confirm Password
        tk.Label(form_frame, text='Confirm Password *', bg='#f0f0f0',
                 font=('Arial', 10, 'bold'), anchor='w').grid(
            row=2, column=0, sticky='w', pady=(10, 5), padx=5)
        confirm_entry = tk.Entry(form_frame, width=40, font=('Arial', 10), show='●')
        confirm_entry.grid(row=3, column=0, pady=(0, 10), padx=5, sticky='ew')

        form_frame.grid_columnconfigure(0, weight=1)

        # ===== ۵. تابع ذخیره =====
        def save():
            try:
                pwd = pwd_entry.get()
                confirm = confirm_entry.get()

                if len(pwd) < 6:
                    messagebox.showerror('Error', '❌ Password must be at least 6 characters!')
                    return
                if pwd != confirm:
                    messagebox.showerror('Error', '❌ Passwords do not match!')
                    return

                import hashlib
                pwd_hash = hashlib.sha256(pwd.encode()).hexdigest()

                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute('UPDATE users SET password_hash = ? WHERE id = ?',
                               (pwd_hash, real_id))
                conn.commit()
                conn.close()

                messagebox.showinfo('Success',
                                    f'✅ Password changed!\n\n'
                                    f'User: {username}\n'
                                    f'New Password: {pwd}')

                try:
                    log_action('UPDATE', entity_type='user',
                               details=f'Password changed for user "{username}"',
                               user=getattr(self.root, 'current_user', None))
                except:
                    pass
                win.destroy()

            except Exception as e:
                messagebox.showerror('Error', f'❌ Error: {e}')

        # ===== ۶. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '💾 Change', 'command': save, 'color': 'warning', 'width': 160},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#f0f0f0')

        # ===== ۷. Maximize =====
        win.update_idletasks()
        win.state('zoomed')

        pwd_entry.focus()



    def toggle_user_active(self):
        selected = self.users_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select a user first!')
            return

        real_id = self.users_tree.item(selected[0])['tags'][0]
        values = self.users_tree.item(selected[0])['values']
        username = values[1]
        current_status = values[4]

        if hasattr(self.root, 'current_user') and self.root.current_user:
            if str(self.root.current_user['id']) == str(real_id):
                messagebox.showerror('Error', '❌ You cannot disable your own account!')
                return

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT COUNT(*) FROM users
            WHERE role = 'admin' AND is_active = 1
        ''')
        active_admins = cursor.fetchone()[0]

        cursor.execute('SELECT role FROM users WHERE id = ?', (real_id,))
        user_role = cursor.fetchone()[0]
        conn.close()

        if user_role == 'admin' and active_admins <= 1 and 'Active' in current_status:
            messagebox.showerror('Error', '❌ Cannot disable the last active Admin!')
            return

        new_status = 0 if 'Active' in current_status else 1
        action = 'disable' if new_status == 0 else 'enable'

        if not messagebox.askyesno('Confirm',
            f'{action.capitalize()} user "{username}"?'):
            return

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('UPDATE users SET is_active = ? WHERE id = ?',
                       (new_status, real_id))
        conn.commit()
        conn.close()

        messagebox.showinfo('Success',
            f'✅ User "{username}" has been {action}d!')
        log_action(
            'UPDATE',
            entity_type='user',
            entity_id=real_id,
            details=f'User "{username}" {action}d (status={new_status})',
            user=getattr(self.root, 'current_user', None)
        )
        self.load_users()

    def delete_user(self):
        selected = self.users_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select at least one user first!')
            return

        count = len(selected)

        # چک کردن اینکه خودش رو حذف نکنه
        for item in selected:
            real_id = self.users_tree.item(item)['tags'][0]
            if hasattr(self.root, 'current_user') and self.root.current_user:
                if str(self.root.current_user['id']) == str(real_id):
                    messagebox.showerror('Error', '❌ You cannot delete your own account!')
                    return

        if count == 1:
            values = self.users_tree.item(selected[0])['values']
            confirm_msg = f'Delete user "{values[1]}"?\n\nThis action cannot be undone!'
        else:
            confirm_msg = f'Delete {count} user(s)?\n\nThis action cannot be undone!'

        if not messagebox.askyesno('Confirm', confirm_msg):
            return

        try:
            conn = get_connection()
            cursor = conn.cursor()

            deleted = 0
            for item in selected:
                real_id = self.users_tree.item(item)['tags'][0]

                # چک آخرین ادمین
                cursor.execute('SELECT role FROM users WHERE id = ?', (real_id,))
                row = cursor.fetchone()
                if row and row[0] == 'admin':
                    cursor.execute("SELECT COUNT(*) FROM users WHERE role = 'admin'")
                    if cursor.fetchone()[0] <= 1:
                        continue  # رد کردن آخرین ادمین

                cursor.execute('DELETE FROM users WHERE id = ?', (real_id,))
                deleted += 1

            conn.commit()
            conn.close()

            messagebox.showinfo('Success', f'✅ {deleted} user(s) deleted!')
            self.load_users()

        except Exception as e:
            messagebox.showerror('Error', f'❌ Error: {e}')

    def show_audit_log(self):
        """نمایش لاگ سیستم (بهبود یافته برای حالت مینیمم)"""
        win = tk.Toplevel(self.root)
        win.title('🔐 Audit Log - System Activity')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        # ===== ۱. هدر =====
        header = tk.Frame(win, bg='#34495e', height=55)
        header.pack(fill='x', side='top')
        header.pack_propagate(False)
        tk.Label(header, text='🔐 Audit Log - System Activity',
                 font=('Arial', 15, 'bold'), bg='#34495e',
                 fg='white').pack(pady=12)

        # ===== ۲. دکمه‌های پایین (ثابت) =====
        btn_frame = tk.Frame(win, bg='#ecf0f1')
        btn_frame.pack(side='bottom', fill='x', pady=12)

        # ===== ۳. بدنه اصلی (بدون اسکرول افقی کل صفحه) =====
        content_frame = tk.Frame(win, bg='#ecf0f1')
        content_frame.pack(fill='both', expand=True, padx=20, pady=10)

        # ===== فیلترها =====
        filter_frame = tk.Frame(content_frame, bg='#ecf0f1')
        filter_frame.pack(fill='x', pady=(0, 10))

        # ردیف ۱: Search + Action + Entity
        row1 = tk.Frame(filter_frame, bg='#ecf0f1')
        row1.pack(fill='x', pady=5)

        tk.Label(row1, text='🔍 Search:', font=('Arial', 10, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=(0, 5))

        search_var = tk.StringVar()
        search_entry = tk.Entry(row1, textvariable=search_var,
                                font=('Arial', 10), width=25,
                                bg='white', relief='solid', bd=1)
        search_entry.pack(side='left', ipady=4, padx=(0, 15))

        # Action filter
        tk.Label(row1, text='Action:', font=('Arial', 10, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=(5, 5))

        action_var = tk.StringVar(value='All')
        action_combo = ttk.Combobox(row1, textvariable=action_var,
                                    values=['All', 'CREATE', 'UPDATE', 'DELETE',
                                            'LOGIN', 'LOGOUT', 'PRINT', 'BACKUP', 'RESTORE'],
                                    state='readonly', width=12, font=('Arial', 10))
        action_combo.pack(side='left', padx=(0, 15))

        # Entity filter
        tk.Label(row1, text='Entity:', font=('Arial', 10, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=(5, 5))

        entity_var = tk.StringVar(value='All')
        entity_combo = ttk.Combobox(row1, textvariable=entity_var,
                                    values=['All', 'patient', 'doctor', 'appointment',
                                            'payment', 'medicine', 'medical_record',
                                            'prescription', 'user', 'database', 'sale',
                                            'lab_request', 'lab_test'],
                                    state='readonly', width=14, font=('Arial', 10))
        entity_combo.pack(side='left', padx=(0, 10))

        # ردیف ۲: Date filter + Count
        row2 = tk.Frame(filter_frame, bg='#ecf0f1')
        row2.pack(fill='x', pady=5)

        tk.Label(row2, text='📅 Date:', font=('Arial', 10, 'bold'),
                 bg='#ecf0f1').pack(side='left', padx=(0, 5))

        date_filter_var, date_from_var, date_to_var, df_container, _ = create_date_filter(
            row2, '#ecf0f1', lambda: load_logs()
        )
        df_container.pack(side='left')

        count_var = tk.StringVar(value='')
        tk.Label(row2, textvariable=count_var,
                 font=('Arial', 10, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(side='right')

        # ===== ۴. جدول با اسکرول داخلی (نکته کلیدی) =====
        tree_container = tk.Frame(content_frame, bg='#ecf0f1')
        tree_container.pack(fill='both', expand=True)

        # یک Canvas مخصوص جدول می‌سازیم تا فقط جدول اسکرول افقی بخورد
        tree_canvas = tk.Canvas(tree_container, bg='#ecf0f1', highlightthickness=0)
        tree_v_scroll = ttk.Scrollbar(tree_container, orient='vertical', command=tree_canvas.yview)
        tree_h_scroll = ttk.Scrollbar(tree_container, orient='horizontal', command=tree_canvas.xview)

        # فریم داخلی که جدول را نگه می‌دارد
        tree_inner = tk.Frame(tree_canvas, bg='#ecf0f1')
        tree_canvas.configure(yscrollcommand=tree_v_scroll.set, xscrollcommand=tree_h_scroll.set)

        # ساخت پنجره داخل Canvas
        tree_window = tree_canvas.create_window((0, 0), window=tree_inner, anchor='nw')

        # چیدمان گرید
        tree_canvas.grid(row=0, column=0, sticky='nsew')
        tree_v_scroll.grid(row=0, column=1, sticky='ns')
        tree_h_scroll.grid(row=1, column=0, sticky='ew')

        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        # ===== ساخت Treeview =====
        cols = ('ID', 'Timestamp', 'User', 'Role', 'Action', 'Entity', 'ID', 'Details')

        # عرض‌های ثابت و بزرگتر برای ستون‌ها تا در حالت کوچک فشرده نشوند
        widths = {'ID': 60, 'Timestamp': 150, 'User': 120, 'Role': 100,
                  'Action': 120, 'Entity': 140, 'ID': 60, 'Details': 500}

        tree = ttk.Treeview(tree_inner, columns=cols, show='headings', height=20)

        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=widths[c], anchor='w' if c in ('Details', 'User', 'Entity') else 'center',
                        stretch=False)

        tree.pack(fill='both', expand=True)

        # ===== تابع مدیریت اسکرول جدول =====
        def update_tree_scroll(*args):
            tree_canvas.configure(scrollregion=tree_canvas.bbox('all'))
            # عرض جدول را با عرض Canvas هماهنگ می‌کنیم
            tree_canvas.itemconfig(tree_window, width=tree_canvas.winfo_width())

            # اگر محتوای جدول از پنجره عریض‌تر بود، اسکرول افقی نمایش داده شود
            bbox = tree_canvas.bbox('all')
            if bbox:
                content_width = bbox[2] - bbox[0]
                canvas_width = tree_canvas.winfo_width()
                if content_width > canvas_width + 5:
                    tree_h_scroll.grid(row=1, column=0, sticky='ew')
                else:
                    tree_h_scroll.grid_remove()

        tree.bind('<Configure>', update_tree_scroll)
        tree_canvas.bind('<Configure>', update_tree_scroll)

        # اسکرول با ماوس
        def on_mousewheel(event):
            tree_canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        tree_canvas.bind_all('<MouseWheel>', on_mousewheel)

        # رنگ‌ها
        tree.tag_configure('CREATE', background='#d5f4e6')
        tree.tag_configure('UPDATE', background='#fff3cd')
        tree.tag_configure('DELETE', background='#f8d7da')
        tree.tag_configure('LOGIN', background='#d1ecf1')
        tree.tag_configure('LOGOUT', background='#d1ecf1')
        tree.tag_configure('PRINT', background='#e2e3e5')
        tree.tag_configure('BACKUP', background='#cce5ff')
        tree.tag_configure('RESTORE', background='#ffe5cc')

        def load_logs():
            from_date, to_date = get_date_range(
                date_filter_var.get(),
                date_from_var.get().strip(),
                date_to_var.get().strip()
            )

            search = search_var.get().strip()
            action = action_var.get()
            entity = entity_var.get()

            conn = get_connection()
            cursor = conn.cursor()

            query = 'SELECT id, timestamp, username, user_role, action, entity_type, entity_id, details FROM audit_log WHERE 1=1'
            params = []

            if search:
                query += ' AND (username LIKE ? OR details LIKE ? OR action LIKE ?)'
                pattern = f'%{search}%'
                params.extend([pattern, pattern, pattern])

            if action != 'All':
                query += ' AND action = ?'
                params.append(action)

            if entity != 'All':
                query += ' AND entity_type = ?'
                params.append(entity)

            if from_date:
                query += ' AND DATE(timestamp) >= ?'
                params.append(from_date)

            if to_date:
                query += ' AND DATE(timestamp) <= ?'
                params.append(to_date)

            query += ' ORDER BY timestamp DESC LIMIT 1000'

            cursor.execute(query, params)
            logs = cursor.fetchall()
            conn.close()

            for item in tree.get_children():
                tree.delete(item)

            for log in logs:
                log_id, timestamp, username, role, action_val, entity_type, entity_id, details = log

                icon = {
                    'CREATE': '➕',
                    'UPDATE': '✏️',
                    'DELETE': '🗑️',
                    'LOGIN': '🔓',
                    'LOGOUT': '🔒',
                    'PRINT': '🖨️',
                    'BACKUP': '💾',
                    'RESTORE': '♻️',
                }.get(action_val, '📌')

                tree.insert('', 'end', values=(
                    log_id,
                    timestamp or '—',
                    username or '—',
                    role or '—',
                    f'{icon} {action_val}',
                    entity_type or '—',
                    entity_id or '—',
                    details or '—'
                ), tags=(action_val,))

            count_var.set(f'📋 {len(logs)} log entries')
            win.after(100, update_tree_scroll)  # بعد از لود، اسکرول چک شود

        # Bind فیلترها
        search_var.trace_add('write', lambda *args: load_logs())
        action_combo.bind('<<ComboboxSelected>>', lambda e: load_logs())
        entity_combo.bind('<<ComboboxSelected>>', lambda e: load_logs())

        # ===== Export (جایگزین بخش قبلی در show_audit_log) =====
        def export_log():
            from tkinter import filedialog
            from datetime import datetime

            if not tree.get_children():
                messagebox.showwarning('Warning', '📭 No data to export!')
                return

            file_path = filedialog.asksaveasfilename(
                defaultextension='.xlsx',
                filetypes=[('Excel files', '*.xlsx')],
                initialfile=f'audit_log_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
            )

            if not file_path:
                return

            # آماده‌سازی داده‌ها برای تابع export_to_excel
            headers = ['ID', 'Timestamp', 'User', 'Role', 'Action', 'Entity', 'Entity ID', 'Details']
            rows = []
            for item in tree.get_children():
                vals = tree.item(item)['values']
                # فقط مقادیر بدون تگ‌های رنگی
                rows.append(tuple(vals))

            # استفاده از تابع اکسل‌ساز پروژه
            success, result = export_to_excel(
                headers, rows, file_path,
                sheet_name='Audit Log',
                title='System Audit Log'
            )

            if success:
                messagebox.showinfo('✅ Export Successful', f'Excel file saved to:\n{result}')
                try:
                    os.startfile(result) # باز کردن فایل به صورت خودکار
                except:
                    pass
            else:
                messagebox.showerror('Export Error', f'❌ {result}')

        def clear_old_logs():
            if not messagebox.askyesno('Confirm',
                                       'Delete audit logs older than 6 months?\n\nThis action cannot be undone!'):
                return
            try:
                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute("DELETE FROM audit_log WHERE timestamp < DATE('now', '-6 months')")
                deleted = cursor.rowcount
                conn.commit()
                conn.close()

                log_action('DELETE', entity_type='audit_log',
                           details=f'Cleaned {deleted} old log entries',
                           user=getattr(self.root, 'current_user', None))

                messagebox.showinfo('Success', f'✅ Deleted {deleted} old entries')
                load_logs()
            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        # ===== دکمه‌ها =====
        create_popup_buttons(btn_frame, [
            {'text': '🔄 Refresh', 'command': load_logs, 'color': 'primary', 'width': 140},
            {'text': '📊 Export CSV', 'command': export_log, 'color': 'info', 'width': 160},
            {'text': '🗑️ Clean Old', 'command': clear_old_logs, 'color': 'warning', 'width': 160},
            {'text': '❌ Close', 'command': win.destroy, 'color': 'danger', 'width': 120},
        ], bg='#ecf0f1')

        # ===== بارگذاری =====
        load_logs()

        # ===== Maximize =====
        win.update_idletasks()
        win.state('zoomed')


    def show_clinic_statistics(self):
        """داشبورد آمار کلینیک (واکنش‌گرا و بدون اسکرول در حالت فول اسکرین)"""
        win = tk.Toplevel(self.root)
        win.title('📈 Clinic Statistics')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین ثابت) =====
        btn_container = tk.Frame(win, bg='#ecf0f1')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه اصلی با اسکرول‌بار =====
        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=5, pady=5)

        # استفاده از Canvas برای اسکرول
        canvas = tk.Canvas(body_container, bg='#ecf0f1', highlightthickness=0)

        # اسکرول عمودی (سمت راست)
        v_scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        # اسکرول افقی (پایین)
        h_scrollbar = ttk.Scrollbar(body_container, orient='horizontal', command=canvas.xview)

        scrollable = tk.Frame(canvas, bg='#ecf0f1')

        # اتصال اسکرول به Canvas
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')

        # چیدمان با Grid
        canvas.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')

        body_container.grid_rowconfigure(0, weight=1)
        body_container.grid_columnconfigure(0, weight=1)

        # ===== 🌟 جادوی اسکرول (تنها زمانی که لازم باشد نمایش داده می‌شود) =====
        def check_scrollbars(event=None):
            # به‌روزرسانی ناحیه اسکرول
            canvas.configure(scrollregion=canvas.bbox('all'))

            # چک کردن ارتفاع (برای اسکرول عمودی)
            bbox = canvas.bbox('all')
            if bbox:
                content_height = bbox[3] - bbox[1]
                canvas_height = canvas.winfo_height()
                if content_height > canvas_height + 5:
                    v_scrollbar.grid(row=0, column=1, sticky='ns')
                else:
                    v_scrollbar.grid_remove()  # مخفی کردن

                # چک کردن عرض (برای اسکرول افقی)
                content_width = bbox[2] - bbox[0]
                canvas_width = canvas.winfo_width()
                if content_width > canvas_width + 5:
                    h_scrollbar.grid(row=1, column=0, sticky='ew')
                else:
                    h_scrollbar.grid_remove()  # مخفی کردن

        canvas.bind('<Configure>', check_scrollbars)
        scrollable.bind('<Configure>', check_scrollbars)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. یک فریم داخلی برای وسط‌چین کردن =====
        # این فریم باعث می‌شود در حالت فول اسکرین همه چیز وسط باشد
        center_frame = tk.Frame(scrollable, bg='#ecf0f1')
        center_frame.pack(fill='both', expand=True, padx=20, pady=10)

        # ===== ۴. عنوان =====
        tk.Label(center_frame, text='📈 Clinic Statistics',
                 font=('Arial', 16, 'bold'), bg='#ecf0f1',
                 fg='#2c3e50').pack(pady=15)

        # ===== ۵. خواندن داده‌ها =====
        conn = get_connection()
        cursor = conn.cursor()

        cursor.execute('SELECT COUNT(*) FROM patients')
        total_patients = cursor.fetchone()[0]

        cursor.execute('SELECT COUNT(*) FROM doctors')
        total_doctors = cursor.fetchone()[0]

        cursor.execute('SELECT COUNT(*) FROM medical_records')
        total_visits = cursor.fetchone()[0]

        cursor.execute('SELECT COUNT(*) FROM prescriptions')
        total_rx = cursor.fetchone()[0]

        cursor.execute("SELECT COALESCE(SUM(final_amount), 0) FROM payments WHERE payment_status = 'Paid'")
        total_revenue = cursor.fetchone()[0]

        # Top Medicines
        cursor.execute('''
            SELECT m.name, COUNT(pr.id) as cnt
            FROM prescriptions pr
            JOIN medicines m ON pr.medicine_id = m.id
            GROUP BY m.id
            ORDER BY cnt DESC
            LIMIT 5
        ''')
        top_meds = cursor.fetchall()

        # Top Doctors
        cursor.execute('''
            SELECT d.first_name || ' ' || d.last_name, COUNT(mr.id) as cnt
            FROM medical_records mr
            JOIN doctors d ON mr.doctor_id = d.id
            GROUP BY d.id
            ORDER BY cnt DESC
            LIMIT 5
        ''')
        top_docs = cursor.fetchall()

        conn.close()

        # ===== ۶. کارت‌های آمار =====
        cards_row = tk.Frame(center_frame, bg='#ecf0f1')
        cards_row.pack(pady=10)

        def make_card(parent, icon, title, value, color):
            # استفاده از width و height ثابت برای جلوگیری از بهم ریختگی
            card = tk.Frame(parent, bg=color, width=200, height=110)
            card.pack(side='left', padx=8)
            card.pack_propagate(False)  # جلوگیری از تغییر سایز توسط محتوا

            tk.Label(card, text=icon, font=('Arial', 24),
                     bg=color, fg='white').pack(pady=(10, 0))
            tk.Label(card, text=str(value), font=('Arial', 20, 'bold'),
                     bg=color, fg='white').pack()
            tk.Label(card, text=title, font=('Arial', 10),
                     bg=color, fg='white').pack()

        make_card(cards_row, '👥', 'Patients', total_patients, '#3498db')
        make_card(cards_row, '👨‍⚕️', 'Doctors', total_doctors, '#9b59b6')
        make_card(cards_row, '📋', 'Visits', total_visits, '#27ae60')
        make_card(cards_row, '💊', 'Prescriptions', total_rx, '#e74c3c')
        make_card(cards_row, '💰', 'Revenue', f'{total_revenue:,.0f}', '#f39c12')

        # ===== ۷. کادرهای Top Lists =====
        top_row = tk.Frame(center_frame, bg='#ecf0f1')
        top_row.pack(pady=15)

        # --- کادر Top Medicines (قرمز) ---
        # عرض 450 و ارتفاع 320 ثابت می‌کنیم تا کادرها هم‌اندازه باشند
        med_frame = tk.Frame(top_row, bg='white', relief='solid', bd=1, width=420, height=320)
        med_frame.pack(side='left', padx=10)
        med_frame.pack_propagate(False)

        med_header = tk.Frame(med_frame, bg='#c0392b')
        med_header.pack(fill='x')
        med_title = f'🏆 Top {len(top_meds)} Medicines' if top_meds else '🏆 Top Medicines'
        tk.Label(med_header, text=med_title,
                 font=('Arial', 11, 'bold'), bg='#c0392b',
                 fg='white', pady=6).pack()

        med_content = tk.Frame(med_frame, bg='white')
        med_content.pack(fill='both', expand=True, padx=10, pady=8)

        if top_meds:
            for i, (name, cnt) in enumerate(top_meds, 1):
                row = tk.Frame(med_content, bg='white')
                row.pack(fill='x', pady=3)

                colors = ['#f1c40f', '#95a5a6', '#e67e22', '#3498db', '#1abc9c']
                color = colors[i - 1] if i <= 5 else '#95a5a6'

                tk.Label(row, text=f' {i} ', font=('Arial', 10, 'bold'),
                         bg=color, fg='white', width=2).pack(side='left', padx=(0, 8))

                tk.Label(row, text=name[:25], font=('Arial', 10),
                         bg='white', anchor='w').pack(side='left', fill='x', expand=True)
                tk.Label(row, text=f'💊 {cnt}', font=('Arial', 10, 'bold'),
                         bg='white', fg='#c0392b', width=8, anchor='e').pack(side='right')
        else:
            tk.Label(med_content, text='📭 No data yet',
                     font=('Arial', 10, 'italic'), bg='white', fg='#95a5a6').pack(pady=20)

        # --- کادر Top Doctors (بنفش) ---
        doc_frame = tk.Frame(top_row, bg='white', relief='solid', bd=1, width=420, height=320)
        doc_frame.pack(side='left', padx=10)
        doc_frame.pack_propagate(False)

        doc_header = tk.Frame(doc_frame, bg='#8e44ad')
        doc_header.pack(fill='x')
        doc_title = f'🏆 Top {len(top_docs)} Doctors' if top_docs else '🏆 Top Doctors'
        tk.Label(doc_header, text=doc_title,
                 font=('Arial', 11, 'bold'), bg='#8e44ad',
                 fg='white', pady=6).pack()

        doc_content = tk.Frame(doc_frame, bg='white')
        doc_content.pack(fill='both', expand=True, padx=10, pady=8)

        if top_docs:
            for i, (name, cnt) in enumerate(top_docs, 1):
                row = tk.Frame(doc_content, bg='white')
                row.pack(fill='x', pady=3)

                colors = ['#f1c40f', '#95a5a6', '#e67e22', '#3498db', '#1abc9c']
                color = colors[i - 1] if i <= 5 else '#95a5a6'

                tk.Label(row, text=f' {i} ', font=('Arial', 10, 'bold'),
                         bg=color, fg='white', width=2).pack(side='left', padx=(0, 8))

                doctor_clean = name.replace('Dr.', '').replace('dr.', '').strip()
                tk.Label(row, text=f'Dr. {doctor_clean}'[:25], font=('Arial', 10),
                         bg='white', anchor='w').pack(side='left', fill='x', expand=True)
                tk.Label(row, text=f'👁️ {cnt}', font=('Arial', 10, 'bold'),
                         bg='white', fg='#8e44ad', width=8, anchor='e').pack(side='right')
        else:
            tk.Label(doc_content, text='📭 No data yet',
                     font=('Arial', 10, 'italic'), bg='white', fg='#95a5a6').pack(pady=20)

        # ===== ۸. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '🔄 Refresh', 'command': lambda: [win.destroy(), self.show_clinic_statistics()], 'color': 'primary',
             'width': 140},
            {'text': '❌ Close', 'command': win.destroy, 'color': 'danger', 'width': 120},
        ], bg='#ecf0f1')

        # ===== ۹. Maximize =====
        win.update_idletasks()
        win.state('zoomed')



    def load_medicines(self):
        """بارگذاری لیست داروها"""
        search_text = ''
        if hasattr(self, 'medicine_search_var'):
            search_text = self.medicine_search_var.get().strip()

        stock_filter = 'All'
        if hasattr(self, 'medicine_stock_filter'):
            stock_filter = self.medicine_stock_filter.get()

        conn = get_connection()
        cursor = conn.cursor()

        query = '''
            SELECT id, name, category, price, stock_quantity, min_stock, expiry_date,
                   manufacturer
            FROM medicines
            WHERE 1=1
        '''
        params = []

        if search_text:
            search_pattern = f'%{search_text}%'
            query += '''
                AND (name LIKE ? OR category LIKE ? OR manufacturer LIKE ?)
            '''
            params.extend([search_pattern] * 3)

        if stock_filter == 'In Stock':
            query += ' AND stock_quantity >= min_stock'
        elif stock_filter == 'Low Stock':
            query += ' AND stock_quantity > 0 AND stock_quantity < min_stock'
        elif stock_filter == 'Out of Stock':
            query += ' AND stock_quantity <= 0'

        query += ' ORDER BY name'

        cursor.execute(query, params)
        meds = cursor.fetchall()
        conn.close()

        for item in self.medicines_tree.get_children():
            self.medicines_tree.delete(item)

        low_stock_count = 0
        out_of_stock_count = 0

        for index, m in enumerate(meds, 1):
            real_id = m[0]
            stock = m[4]
            min_stock = m[5]

            if stock < 0:
                stock_display = f'❌ {stock}'
                out_of_stock_count += 1
            elif stock == 0:
                stock_display = f'❌ {stock}'
                out_of_stock_count += 1
            elif stock < min_stock:
                stock_display = f'⚠️ {stock}'
                low_stock_count += 1
            else:
                stock_display = f'✅ {stock}'

            self.medicines_tree.insert(
                '', 'end',
                values=(index, m[1], m[2], f'{m[3]:,.0f}', stock_display,
                        min_stock, m[6] or 'N/A'),
                tags=(str(real_id),)
            )

        if hasattr(self, 'medicine_count_var'):
            if search_text or stock_filter != 'All':
                self.medicine_count_var.set(f'🔍 {len(meds)} result(s)')
            else:
                self.medicine_count_var.set(f'💊 {len(meds)} medicine(s)')

        if hasattr(self, 'medicine_alert_var'):
            alerts = []
            if low_stock_count > 0:
                alerts.append(f'⚠️ {low_stock_count} low')
            if out_of_stock_count > 0:
                alerts.append(f'❌ {out_of_stock_count} out')
            self.medicine_alert_var.set('  '.join(alerts) if alerts else '')

    def sell_medicine_window(self):
        """فرم فروش آزاد دارو (بدون نیاز به بیمار)"""
        win = tk.Toplevel(self.root)
        win.title('💰 Sell Medicine (Direct)')
        win.configure(bg='#f0f0f0')
        win.resizable(True, True)

        btn_container = tk.Frame(win, bg='#f0f0f0')
        btn_container.pack(side='bottom', fill='x', pady=15)

        body_container = tk.Frame(win, bg='#f0f0f0')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f0f0f0')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        tk.Label(scrollable, text='💰 Sell Medicine (Direct)',
                 font=('Arial', 16, 'bold'), bg='#f0f0f0',
                 fg='#2c3e50').pack(pady=15)

        # Medicine Selection
        med_frame = tk.Frame(scrollable, bg='#f0f0f0')
        med_frame.pack(fill='x', pady=5, padx=40)

        tk.Label(med_frame, text='💊 Select Medicine:',
                 font=('Arial', 11, 'bold'), bg='#f0f0f0',
                 width=15, anchor='w').pack(side='left')

        conn = get_connection()
        cursor = conn.cursor()

        # ✅ گرفتن لیست داروها از دیتابیس (این همان خطی است که خطا می‌داد)
        cursor.execute('SELECT id, name, price, stock_quantity FROM medicines WHERE stock_quantity > 0 ORDER BY name')
        meds = cursor.fetchall()
        conn.close()

        if not meds:
            messagebox.showerror('Error', '❌ No medicines in stock!')
            win.destroy()
            return

        med_options = [f'{m[0]} - {m[1]} ({m[2]:,.0f}) - Stock: {m[3]}' for m in meds]
        med_var = tk.StringVar(value=med_options[0])
        ttk.Combobox(med_frame, textvariable=med_var, values=med_options,
                     width=45, state='readonly').pack(side='left', padx=5)

        qty_var = tk.StringVar(value='1')
        tk.Spinbox(med_frame, from_=1, to=1000, textvariable=qty_var,
                   width=8, font=('Arial', 11)).pack(side='left', padx=5)

        # Cart
        tk.Label(scrollable, text='🛒 Cart:',
                 font=('Arial', 12, 'bold'), bg='#f0f0f0',
                 fg='#2c3e50', anchor='w').pack(anchor='w', padx=40, pady=(15, 5))

        cart_container = tk.Frame(scrollable, bg='white', relief='solid', bd=1)
        cart_container.pack(fill='x', padx=40, pady=5)

        cart_cols = ('#', 'Medicine', 'Qty', 'Unit Price', 'Subtotal')
        cart_tree = ttk.Treeview(cart_container, columns=cart_cols, show='headings', height=6)
        for col in cart_cols:
            cart_tree.heading(col, text=col)
        cart_tree.column('#', width=40, anchor='center')
        cart_tree.column('Medicine', width=250, anchor='w')
        cart_tree.column('Qty', width=70, anchor='center')
        cart_tree.column('Unit Price', width=120, anchor='e')
        cart_tree.column('Subtotal', width=130, anchor='e')
        cart_tree.pack(fill='x', padx=5, pady=5)

        total_frame = tk.Frame(scrollable, bg='#f8f9fa', relief='solid', bd=2)
        total_frame.pack(fill='x', padx=40, pady=10)

        total_var = tk.StringVar(value='0 IRR')
        tk.Label(total_frame, text='💰 TOTAL:',
                 font=('Arial', 14, 'bold'), bg='#f8f9fa').pack(side='left', padx=15, pady=10)
        tk.Label(total_frame, textvariable=total_var,
                 font=('Arial', 14, 'bold'), bg='#f8f9fa', fg='#27ae60').pack(side='left')

        cart = []

        def update_cart_display():
            for item in cart_tree.get_children():
                cart_tree.delete(item)

            for idx, item in enumerate(cart, 1):
                cart_tree.insert('', 'end', values=(
                    idx, item['name'], item['qty'],
                    f'{item["price"]:,.0f}', f'{item["total"]:,.0f}'
                ))

            total = sum(item['total'] for item in cart)
            total_var.set(f'{total:,.0f} IRR')

        def add_to_cart():
            try:
                med_id = int(med_var.get().split(' - ')[0])
                qty = int(qty_var.get())
                if qty <= 0:
                    messagebox.showerror('Error', '❌ Quantity must be positive!')
                    return

                med = next((m for m in meds if m[0] == med_id), None)
                if not med:
                    messagebox.showerror('Error', '❌ Medicine not found!')
                    return

                already_in_cart = sum(item['qty'] for item in cart if item['id'] == med_id)
                if qty + already_in_cart > med[3]:
                    messagebox.showerror('Error', f'❌ Only {med[3] - already_in_cart} available!')
                    return

                cart.append({
                    'id': med_id, 'name': med[1], 'qty': qty,
                    'price': med[2], 'total': med[2] * qty
                })
                update_cart_display()
                qty_var.set('1')

            except ValueError:
                messagebox.showerror('Error', '❌ Invalid input!')

        def remove_selected():
            selected = cart_tree.selection()
            if not selected:
                messagebox.showwarning('Warning', '⚠️ Select an item first!')
                return
            idx = int(cart_tree.item(selected[0])['values'][0]) - 1
            if 0 <= idx < len(cart):
                cart.pop(idx)
                update_cart_display()

        def confirm_sale():
            if not cart:
                messagebox.showerror('Error', '❌ Cart is empty!')
                return

            total = sum(item['total'] for item in cart)
            if not messagebox.askyesno('Confirm', f'Sell {len(cart)} item(s)?\nTotal: {total:,.0f} IRR'):
                return

            try:
                import random
                from datetime import datetime
                receipt = f'INV-{datetime.now().strftime("%Y%m%d%H%M%S")}-{random.randint(100, 999)}'

                conn = get_connection()
                cursor = conn.cursor()

                # گرفتن اولین بیمار به عنوان بیمار پیش‌فرض برای فروش آزاد
                cursor.execute("SELECT id FROM patients LIMIT 1")
                row = cursor.fetchone()
                default_patient_id = row[0] if row else None

                # 1. ثبت در جدول پرداخت‌های داروخانه
                cursor.execute('''
                    INSERT INTO pharmacy_payments
                    (receipt_number, patient_id, amount, discount, final_amount,
                     payment_method, payment_status, created_by)
                    VALUES (?, ?, ?, ?, ?, ?, 'Paid', ?)
                ''', (receipt, default_patient_id, total, 0, total, 'Cash',
                      getattr(self.root, 'current_user', {}).get('username', 'admin')))

                # 2. ثبت فروش
                for item in cart:
                    cursor.execute('''
                        INSERT INTO medicine_sales
                        (medicine_id, patient_id, quantity, unit_price, total_price, invoice_id)
                        VALUES (?, ?, ?, ?, ?, ?)
                    ''', (item['id'], default_patient_id, item['qty'],
                          item['price'], item['total'], receipt))

                    cursor.execute('''
                        UPDATE medicines SET stock_quantity = stock_quantity - ?
                        WHERE id = ?
                    ''', (item['qty'], item['id']))

                conn.commit()
                conn.close()

                self.load_medicines()
                self.load_pharmacy_stats()
                win.destroy()

                messagebox.showinfo('✅ Success', f'✅ Sale complete!\n\nReceipt: {receipt}\nTotal: {total:,.0f} IRR')

            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        create_popup_buttons(btn_container, [
            {'text': '➕ Add to Cart', 'command': add_to_cart, 'color': 'primary', 'width': 160},
            {'text': '🗑️ Remove Selected', 'command': remove_selected, 'color': 'warning', 'width': 180},
            {'text': '✅ Confirm Sale', 'command': confirm_sale, 'color': 'success', 'width': 180},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 120},
        ], bg='#f0f0f0')

        win.update_idletasks()
        win.state('zoomed')
        update_cart_display()


    def add_stock_window(self):
        """فرم افزودن موجودی دارو"""
        selected = self.medicines_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select a medicine first!')
            return

        real_id = self.medicines_tree.item(selected[0])['tags'][0]
        values = self.medicines_tree.item(selected[0])['values']
        med_name = values[1]

        win = tk.Toplevel(self.root)
        win.title('📦 Add Stock')
        win.configure(bg='#f0f0f0')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین) =====
        btn_container = tk.Frame(win, bg='#f0f0f0')
        btn_container.pack(side='bottom', fill='x', pady=15)

        # ===== ۲. بدنه Scrollable =====
        body_container = tk.Frame(win, bg='#f0f0f0')
        body_container.pack(fill='both', expand=True, padx=10, pady=5)

        canvas = tk.Canvas(body_container, bg='#f0f0f0', highlightthickness=0)
        scrollbar = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#f0f0f0')

        scrollable.bind('<Configure>', lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas_window = canvas.create_window((0, 0), window=scrollable, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.bind('<Configure>', lambda e: canvas.itemconfig(canvas_window, width=e.width))

        canvas.pack(side='left', fill='both', expand=True)
        auto_scrollbar(canvas, scrollbar, scrollable)

        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mousewheel)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text='📦 Add Stock',
                 font=('Arial', 16, 'bold'), bg='#f0f0f0',
                 fg='#2c3e50').pack(pady=15)

        # ===== ۴. اطلاعات دارو =====
        info_frame = tk.Frame(scrollable, bg='#f8f9fa', relief='solid', bd=1)
        info_frame.pack(fill='x', padx=40, pady=10)

        tk.Label(info_frame, text=f'💊 Medicine: {med_name}',
                 font=('Arial', 12, 'bold'), bg='#f8f9fa',
                 fg='#2c3e50', anchor='w').pack(anchor='w', padx=15, pady=(10, 5))

        tk.Label(info_frame, text=f'📊 Current Stock: {values[4]}',
                 font=('Arial', 11), bg='#f8f9fa',
                 fg='#7f8c8d', anchor='w').pack(anchor='w', padx=15, pady=(3, 10))

        # ===== ۵. فرم =====
        form_frame = tk.Frame(scrollable, bg='#f0f0f0')
        form_frame.pack(pady=10, padx=40, fill='x')

        tk.Label(form_frame, text='📦 Quantity to add:',
                 font=('Arial', 11, 'bold'), bg='#f0f0f0',
                 anchor='w').pack(anchor='w', pady=(0, 5))

        qty_var = tk.StringVar(value='50')
        qty_entry = tk.Entry(form_frame, textvariable=qty_var,
                             font=('Arial', 12), width=20,
                             justify='center')
        qty_entry.pack(fill='x', ipady=6, pady=(0, 10))

        tk.Label(form_frame, text='💡 Enter a positive number',
                 font=('Arial', 8, 'italic'), bg='#f0f0f0',
                 fg='#7f8c8d', anchor='w').pack(anchor='w')

        # ===== ۶. تابع ذخیره =====
        def do_add():
            try:
                qty = int(qty_var.get())
                if qty <= 0:
                    messagebox.showerror('Error', '❌ Quantity must be positive!')
                    return

                conn = get_connection()
                cursor = conn.cursor()
                cursor.execute('''
                    UPDATE medicines SET stock_quantity = stock_quantity + ?
                    WHERE id = ?
                ''', (qty, real_id))
                conn.commit()
                conn.close()

                messagebox.showinfo('Success', f'✅ Added {qty} units!')

                try:
                    log_action('UPDATE', entity_type='medicine',
                               entity_id=real_id,
                               details=f'Added {qty} units to "{med_name}" stock',
                               user=getattr(self.root, 'current_user', None))
                except:
                    pass

                win.destroy()
                self.load_medicines()

            except ValueError:
                messagebox.showerror('Error', '❌ Please enter a number!')
            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        # ===== ۷. دکمه‌های گِرد =====
        create_popup_buttons(btn_container, [
            {'text': '📦 Add Stock', 'command': do_add, 'color': 'success', 'width': 160},
            {'text': '❌ Cancel', 'command': win.destroy, 'color': 'danger', 'width': 140},
        ], bg='#f0f0f0')

        # ===== ۸. Maximize =====
        win.update_idletasks()
        win.state('zoomed')

        qty_entry.focus()


    def low_stock_window(self):
        """نمایش داروهای کم‌موجود"""
        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT name, stock_quantity, min_stock
            FROM medicines
            WHERE stock_quantity < min_stock
            ORDER BY stock_quantity
        ''')
        low = cursor.fetchall()
        conn.close()

        if not low:
            messagebox.showinfo('✅ All Good', 'All medicines have sufficient stock!')
            return

        win = tk.Toplevel(self.root)
        win.title('⚠️ Low Stock Alert')
        win.geometry('500x400')
        win.configure(bg='#ecf0f1')

        win.update_idletasks()
        w, h = 500, 400
        x = (win.winfo_screenwidth() // 2) - (w // 2)
        y = (win.winfo_screenheight() // 2) - (h // 2)
        win.geometry(f'{w}x{h}+{x}+{y}')

        tk.Label(win, text='⚠️ Low Stock Alert', font=('Arial', 14, 'bold'),
                 bg='#ecf0f1', fg='#e67e22').pack(pady=10)

        text = tk.Text(win, font=('Courier New', 11), bg='white',
                       padx=15, pady=15, wrap='word')
        text.pack(fill='both', expand=True, padx=15, pady=10)

        for name, stock, min_stock in low:
            if stock < 0:
                text.insert('end', f'❌ {name}\n   Stock: {stock} (negative!)\n')
            elif stock == 0:
                text.insert('end', f'🔴 {name}\n   Stock: 0 | Min: {min_stock}\n')
            else:
                text.insert('end', f'⚠️  {name}\n   Stock: {stock} | Min: {min_stock}\n')
            text.insert('end', '-' * 40 + '\n')

        text.config(state='disabled')

        tk.Button(win, text='❌ Close', command=win.destroy,
                  bg='#e74c3c', fg='white', font=('Arial', 11, 'bold'),
                  padx=25, pady=8).pack(pady=10)

    def pharmacy_sales_report(self):
        """گزارش فروش روزانه داروخانه"""
        from datetime import datetime
        date_str = simpledialog.askstring('Sales Report',
                                          'Date (YYYY-MM-DD) or leave empty for today:')

        if date_str is None:
            return
        if not date_str.strip():
            date_str = datetime.now().strftime('%Y-%m-%d')

        try:
            datetime.strptime(date_str, '%Y-%m-%d')
        except ValueError:
            messagebox.showerror('Error', '❌ Invalid date! Use YYYY-MM-DD')
            return

        conn = get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT COUNT(*), SUM(quantity), SUM(total_price)
            FROM medicine_sales WHERE DATE(sale_date) = ?
        ''', (date_str,))
        result = cursor.fetchone()

        cursor.execute('''
            SELECT m.name, SUM(s.quantity), SUM(s.total_price)
            FROM medicine_sales s
            JOIN medicines m ON s.medicine_id = m.id
            WHERE DATE(s.sale_date) = ?
            GROUP BY m.id ORDER BY SUM(s.total_price) DESC LIMIT 10
        ''', (date_str,))
        top_sales = cursor.fetchall()
        conn.close()

        msg = f'📊 Sales Report - {date_str}\n\n'
        msg += f'Total Sales: {result[0]}\n'
        msg += f'Total Items: {result[1] or 0}\n'
        msg += f'Total Revenue: {result[2] or 0:,.0f} IRR\n'

        if top_sales:
            msg += '\n🔥 Top Sales:\n'
            for name, qty, total in top_sales:
                msg += f'  {name}: {qty} units ({total:,.0f} IRR)\n'

        messagebox.showinfo('📊 Sales Report', msg)

    def delete_medicine(self):
        """حذف دارو"""
        selected = self.medicines_tree.selection()
        if not selected:
            messagebox.showwarning('Warning', '⚠️ Please select a medicine!')
            return

        count = len(selected)

        if count == 1:
            values = self.medicines_tree.item(selected[0])['values']
            med_name = values[1]
            confirm_msg = f'Delete "{med_name}"?'
        else:
            confirm_msg = f'Delete {count} medicine(s)?'

        if not messagebox.askyesno('Confirm', confirm_msg):
            return

        try:
            conn = get_connection()
            cursor = conn.cursor()

            deleted = 0
            for item in selected:
                real_id = self.medicines_tree.item(item)['tags'][0]
                try:
                    cursor.execute('DELETE FROM medicines WHERE id = ?', (real_id,))
                    deleted += 1
                except sqlite3.IntegrityError:
                    pass

            conn.commit()
            conn.close()

            messagebox.showinfo('Success', f'✅ {deleted} medicine(s) deleted!')
            self.load_medicines()

        except Exception as e:
            messagebox.showerror('Error', f'❌ {e}')

    def pharmacy_full_report_window(self):
        """پنجره گزارش جامع داروخانه (روزانه/هفتگی/ماهانه/سالانه + نمودار)"""
        from datetime import datetime, timedelta

        win = tk.Toplevel(self.root)
        win.title('📊 Pharmacy Report Center')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        # ===== ۱. دکمه‌ها (پایین ثابت) =====
        btn_container = tk.Frame(win, bg='#ecf0f1')
        btn_container.pack(side='bottom', fill='x', pady=12)

        # ===== ۲. بدنه اصلی با اسکرول عمودی =====
        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=5, pady=5)

        canvas = tk.Canvas(body_container, bg='#ecf0f1', highlightthickness=0)
        v_scroll = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)

        scrollable = tk.Frame(canvas, bg='#ecf0f1')
        canvas.configure(yscrollcommand=v_scroll.set)
        canvas_win = canvas.create_window((0, 0), window=scrollable, anchor='nw')

        canvas.pack(side='left', fill='both', expand=True)
        v_scroll.pack(side='right', fill='y')

        def check_scroll(*args):
            canvas.configure(scrollregion=canvas.bbox('all'))
            bbox = canvas.bbox('all')
            if bbox:
                ch = bbox[3] - bbox[1]
                if ch > canvas.winfo_height() + 5:
                    if not v_scroll.winfo_ismapped(): v_scroll.pack(side='right', fill='y')
                else:
                    if v_scroll.winfo_ismapped(): v_scroll.pack_forget()

        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_win, width=event.width)
            check_scroll()

        scrollable.bind('<Configure>', check_scroll)
        canvas.bind('<Configure>', _on_canvas_configure)

        def on_mw(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mw)

        # ===== ۳. عنوان =====
        tk.Label(scrollable, text='📊 Pharmacy Report Center',
                 font=('Arial', 18, 'bold'), bg='#ecf0f1', fg='#2c3e50').pack(pady=15)

        # ===== ۴. انتخاب دوره =====
        period_frame = tk.Frame(scrollable, bg='#ecf0f1')
        period_frame.pack(pady=5)

        tk.Label(period_frame, text='📅 Period:', font=('Arial', 12, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(side='left', padx=(0, 10))

        period_var = tk.StringVar(value='Daily')
        periods = [('Daily', '📅 Daily'), ('Weekly', '📆 Weekly'),
                   ('Monthly', '🗓️ Monthly'), ('Yearly', '📅 Yearly')]

        for val, lbl in periods:
            tk.Radiobutton(period_frame, text=lbl, variable=period_var, value=val,
                           bg='#ecf0f1', font=('Arial', 10, 'bold'),
                           command=lambda: load_report(),
                           activebackground='#ecf0f1').pack(side='left', padx=5)

        range_var = tk.StringVar(value='')
        tk.Label(scrollable, textvariable=range_var, font=('Arial', 11, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(pady=5)

        # ===== ۵. کارت‌های آمار =====
        cards_outer = tk.Frame(scrollable, bg='#ecf0f1')
        cards_outer.pack(fill='x', pady=10)  # پر کردن عرض

        cards_inner = tk.Frame(cards_outer, bg='#ecf0f1')
        cards_inner.pack()  # وسط‌چین کردن کارت‌ها

        stat_updaters = {}
        CARD_COLOR = '#3b4a6b'

        card_configs = [
            ('total', '💰 Total Revenue'),
            ('direct', '🛒 Direct Sales'),
            ('dispense', '📋 Dispensed'),
            ('items', '📦 Items Sold'),
            ('count', '🧾 Trans.')
        ]

        for key, title in card_configs:
            canvas_card, updater = create_rounded_stat_card(
                cards_inner, title, '0', bg_color=CARD_COLOR, width=160, height=85
            )
            canvas_card.pack(side='left', padx=5)
            stat_updaters[key] = updater

        # ===== ۶. Notebook برای نمودارها و لیست =====
        style = ttk.Style()
        style.theme_use('clam')
        notebook = ttk.Notebook(scrollable)
        notebook.pack(fill='both', expand=True, padx=20, pady=10)

        # --- Tab 1: Revenue Trend ---
        tab_trend = tk.Frame(notebook, bg='white')
        notebook.add(tab_trend, text='📈 Revenue Trend')
        trend_chart_frame = tk.Frame(tab_trend, bg='white')
        trend_chart_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # --- Tab 2: Stock Consumption ---
        tab_consumption = tk.Frame(notebook, bg='white')
        notebook.add(tab_consumption, text='📊 Stock Consumption')
        consumption_chart_frame = tk.Frame(tab_consumption, bg='white')
        consumption_chart_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # --- Tab 3: Detailed List ---
        tab_detail = tk.Frame(notebook, bg='white')
        notebook.add(tab_detail, text='📋 Detailed List')

        detail_cols = ('Date', 'Type', 'Medicine', 'Qty', 'Unit Price', 'Total')
        detail_frame = tk.Frame(tab_detail, bg='white')
        detail_frame.pack(fill='both', expand=True, padx=10, pady=10)
        detail_tree = ttk.Treeview(detail_frame, columns=detail_cols, show='headings', height=15)
        for c, w in zip(detail_cols, [120, 120, 250, 70, 120, 130]):
            detail_tree.heading(c, text=c)
            detail_tree.column(c, width=w, anchor='center')
        d_scroll = ttk.Scrollbar(detail_frame, orient='vertical', command=detail_tree.yview)
        detail_tree.configure(yscrollcommand=d_scroll.set)
        detail_tree.pack(side='left', fill='both', expand=True)
        d_scroll.pack(side='right', fill='y')

        # ===== ۷. داده‌های سراسری =====
        report_data = {
            'period': 'Daily',
            'from_date': None, 'to_date': None,
            'stats': {},
            'trend': [],
            'consumption': [],
            'detail': [],
        }

        def get_period_range(period):
            today = datetime.now().date()
            if period == 'Daily':
                return today.strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d'), 'Today'
            elif period == 'Weekly':
                return (today - timedelta(days=6)).strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d'), 'Last 7 Days'
            elif period == 'Monthly':
                return today.replace(day=1).strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d'), 'This Month'
            else:
                return today.replace(month=1, day=1).strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d'), 'This Year'

        # ===== ۸. تابع بارگذاری داده‌ها =====
        def load_report():
            from_date, to_date, label = get_period_range(period_var.get())
            report_data['period'] = period_var.get()
            report_data['from_date'] = from_date
            report_data['to_date'] = to_date
            range_var.set(f'📅 {label}: {from_date} → {to_date}')

            try:
                conn = get_connection()
                cursor = conn.cursor()

                # 1. کل درآمد (از جدول pharmacy_payments)
                cursor.execute('''
                    SELECT COALESCE(SUM(final_amount),0), COUNT(*) 
                    FROM pharmacy_payments WHERE DATE(payment_date, 'localtime') BETWEEN ? AND ?
                ''', (from_date, to_date))
                total_rev, total_cnt = cursor.fetchone()

                # 2. فروش آزاد (از جدول medicine_sales)
                cursor.execute('''
                    SELECT COALESCE(SUM(total_price),0), COALESCE(SUM(quantity),0) 
                    FROM medicine_sales WHERE DATE(sale_date, 'localtime') BETWEEN ? AND ?
                ''', (from_date, to_date))
                direct_rev, direct_qty = cursor.fetchone()

                # 3. محاسبه درآمد نسخه‌ها (کل - فروش آزاد)
                disp_rev = max(0, total_rev - direct_rev)
                disp_cnt = total_cnt  # تعداد تراکنش‌ها

                # 4. روند (Trend)
                period = period_var.get()
                time_format = "%Y-%m" if period == "Yearly" else "%Y-%m-%d"

                cursor.execute(f'''
                    SELECT strftime('{time_format}', sale_date, 'localtime') as d, SUM(total_price)
                    FROM medicine_sales WHERE DATE(sale_date, 'localtime') BETWEEN ? AND ?
                    GROUP BY d ORDER BY d
                ''', (from_date, to_date))
                trend_direct = dict(cursor.fetchall())

                cursor.execute(f'''
                    SELECT strftime('{time_format}', payment_date, 'localtime') as d, SUM(final_amount)
                    FROM pharmacy_payments WHERE DATE(payment_date, 'localtime') BETWEEN ? AND ?
                    GROUP BY d ORDER BY d
                ''', (from_date, to_date))
                trend_total = dict(cursor.fetchall())

                all_dates = sorted(set(list(trend_direct.keys()) + list(trend_total.keys())))
                trend = []
                for d in all_dates:
                    t_dir = trend_direct.get(d, 0)
                    t_tot = trend_total.get(d, 0)
                    t_disp = max(0, t_tot - t_dir)
                    trend.append({
                        'date': d,
                        'direct': t_dir,
                        'dispense': t_disp,
                        'total': t_tot
                    })

                # 5. مصرف دارو (Consumption) - فروش آزاد
                cursor.execute('''
                    SELECT m.name, SUM(ms.quantity) as sold, m.stock_quantity as remaining
                    FROM medicine_sales ms
                    JOIN medicines m ON ms.medicine_id = m.id
                    WHERE DATE(ms.sale_date, 'localtime') BETWEEN ? AND ?
                    GROUP BY m.id ORDER BY sold DESC LIMIT 10
                ''', (from_date, to_date))
                cons_sales = cursor.fetchall()

                # 6. مصرف دارو - نسخه‌ها
                cursor.execute('''
                    SELECT m.name, SUM(dm.quantity) as sold, m.stock_quantity as remaining
                    FROM dispensed_medicines dm
                    JOIN medicines m ON dm.medicine_id = m.id
                    JOIN prescriptions pr ON dm.prescription_id = pr.id
                    JOIN medical_records mr ON pr.record_id = mr.id
                    WHERE DATE(mr.record_date, 'localtime') BETWEEN ? AND ?
                    GROUP BY m.id ORDER BY sold DESC LIMIT 10
                ''', (from_date, to_date))
                cons_disp = cursor.fetchall()

                cons_dict = {}
                for name, sold, rem in list(cons_sales) + list(cons_disp):
                    if name in cons_dict:
                        cons_dict[name]['sold'] += sold
                    else:
                        cons_dict[name] = {'sold': sold, 'remaining': rem}
                consumption = sorted(
                    [{'name': n, 'sold': d['sold'], 'remaining': d['remaining']} for n, d in cons_dict.items()],
                    key=lambda x: x['sold'], reverse=True
                )[:10]

                # 7. جزئیات فروش
                cursor.execute('''
                    SELECT sale_date, 'Direct Sale', m.name, ms.quantity, ms.unit_price, ms.total_price
                    FROM medicine_sales ms JOIN medicines m ON ms.medicine_id = m.id
                    WHERE DATE(ms.sale_date, 'localtime') BETWEEN ? AND ?
                    UNION ALL
                    SELECT mr.record_date, 'Prescription', dm.medicine_name, dm.quantity, dm.unit_price, dm.total_price
                    FROM dispensed_medicines dm
                    JOIN prescriptions pr ON dm.prescription_id = pr.id
                    JOIN medical_records mr ON pr.record_id = mr.id
                    WHERE DATE(mr.record_date, 'localtime') BETWEEN ? AND ?
                    ORDER BY 1 DESC LIMIT 200
                ''', (from_date, to_date, from_date, to_date))
                detail = cursor.fetchall()

                conn.close()

                # ذخیره در دیکشنری برای استفاده در نمودارها و اکسل
                report_data['stats'] = {
                    'total': total_rev,
                    'direct': direct_rev,
                    'dispense': disp_rev,
                    'items': direct_qty,
                    'count': total_cnt,
                }
                report_data['trend'] = trend
                report_data['consumption'] = consumption
                report_data['detail'] = detail

                # آپدیت کارت‌های بالای صفحه
                stat_updaters['total'](f'{total_rev:,.0f}')
                stat_updaters['direct'](f'{direct_rev:,.0f}')
                stat_updaters['dispense'](f'{disp_rev:,.0f}')
                stat_updaters['items'](f'{direct_qty:,.0f}')
                stat_updaters['count'](f'{total_cnt}')

                # رسم نمودارها و جدول
                draw_trend_chart()
                draw_consumption_chart()

                for item in detail_tree.get_children():
                    detail_tree.delete(item)
                for row in detail:
                    detail_tree.insert('', 'end', values=(
                        row[0][:16] if row[0] else '—', row[1], row[2],
                        row[3], f'{row[4]:,.0f}', f'{row[5]:,.0f}'
                    ))

            except Exception as e:
                messagebox.showerror('Database Error', f'❌ Error loading report:\n\n{e}')

        # ===== ۹. نمودار روند =====
        def draw_trend_chart():
            for widget in trend_chart_frame.winfo_children():
                widget.destroy()
            if not MATPLOTLIB_AVAILABLE:
                tk.Label(trend_chart_frame, text='⚠️ matplotlib is not installed', font=('Arial', 12), bg='white',
                         fg='#e74c3c').pack(pady=50)
                return
            if not report_data['trend']:
                tk.Label(trend_chart_frame, text='📭 No data in this period', font=('Arial', 12, 'italic'), bg='white',
                         fg='#95a5a6').pack(pady=50)
                return
            try:
                fig = Figure(figsize=(10, 4.5), dpi=80)
                ax = fig.add_subplot(111)
                dates = [t['date'] for t in report_data['trend']]
                direct_vals = [t['direct'] for t in report_data['trend']]
                disp_vals = [t['dispense'] for t in report_data['trend']]
                x = range(len(dates))
                width = 0.35
                bars1 = ax.bar([i - width / 2 for i in x], direct_vals, width, label='Direct Sales', color='#3498db',
                               edgecolor='white', zorder=3)
                bars2 = ax.bar([i + width / 2 for i in x], disp_vals, width, label='Dispensed Rx', color='#9b59b6',
                               edgecolor='white', zorder=3)

                def format_money(val):
                    if val >= 1000000: return f'{val / 1000000:.1f}M'
                    if val >= 1000: return f'{val / 1000:.0f}K'
                    return f'{val:.0f}'

                for bars in [bars1, bars2]:
                    for bar in bars:
                        height = bar.get_height()
                        if height > 0:
                            ax.annotate(format_money(height), xy=(bar.get_x() + bar.get_width() / 2, height),
                                        xytext=(0, 3), textcoords="offset points", ha='center', va='bottom', fontsize=8,
                                        fontweight='bold', color='#2c3e50')
                ax.set_xticks(x)
                if len(dates) > 15:
                    step = max(1, len(dates) // 10)
                    ax.set_xticks(list(x)[::step])
                    ax.set_xticklabels([dates[i] for i in list(x)[::step]], rotation=45, ha='right', fontsize=9)
                else:
                    ax.set_xticklabels(dates, rotation=45, ha='right', fontsize=9)
                ax.set_title('Revenue Trend: Direct Sales vs Prescriptions', fontsize=12, fontweight='bold',
                             color='#2c3e50')
                ax.set_ylabel('Revenue (IRR)', fontsize=10)
                ax.grid(axis='y', alpha=0.3, linestyle='--', zorder=0)
                ax.legend(loc='upper left', fontsize=9)
                ax.spines['top'].set_visible(False)
                ax.spines['right'].set_visible(False)
                from matplotlib.ticker import FuncFormatter
                ax.yaxis.set_major_formatter(FuncFormatter(lambda v, t: format_money(v)))
                fig.tight_layout()
                chart_canvas = FigureCanvasTkAgg(fig, master=trend_chart_frame)
                chart_canvas.draw()
                chart_canvas.get_tk_widget().pack(fill='both', expand=True)
            except Exception as e:
                tk.Label(trend_chart_frame, text=f'⚠️ Chart error: {e}', font=('Arial', 10), bg='white',
                         fg='#e74c3c').pack(pady=30)

        # ===== ۱۰. نمودار مصرف =====
        def draw_consumption_chart():
            for widget in consumption_chart_frame.winfo_children():
                widget.destroy()
            if not MATPLOTLIB_AVAILABLE:
                tk.Label(consumption_chart_frame, text='⚠️ matplotlib is not installed', font=('Arial', 12), bg='white',
                         fg='#e74c3c').pack(pady=50)
                return
            if not report_data['consumption']:
                tk.Label(consumption_chart_frame, text='📭 No medicine consumed in this period',
                         font=('Arial', 12, 'italic'), bg='white', fg='#95a5a6').pack(pady=50)
                return
            try:
                fig = Figure(figsize=(10, 4.5), dpi=80)
                ax = fig.add_subplot(111)
                names = [c['name'][:20] for c in report_data['consumption']]
                sold_vals = [c['sold'] for c in report_data['consumption']]
                remaining_vals = [c['remaining'] for c in report_data['consumption']]
                y_pos = range(len(names))
                ax.barh(y_pos, remaining_vals, color='#27ae60', label='Remaining Stock', alpha=0.8)
                ax.barh(y_pos, sold_vals, left=remaining_vals, color='#e74c3c', label='Sold in Period', alpha=0.9)
                ax.set_yticks(y_pos)
                ax.set_yticklabels(names, fontsize=9)
                ax.invert_yaxis()
                ax.set_title('Medicine Stock Movement\n(Green = Remaining | Red = Consumed in Period)', fontsize=11,
                             fontweight='bold')
                ax.set_xlabel('Quantity', fontsize=10)
                ax.legend(loc='lower right', fontsize=9)
                ax.grid(axis='x', alpha=0.3, linestyle='--')
                ax.spines['top'].set_visible(False)
                ax.spines['right'].set_visible(False)
                for i, (s, r) in enumerate(zip(sold_vals, remaining_vals)):
                    if s > 0: ax.text(r + s / 2, i, f'{int(s)}', ha='center', va='center', fontsize=8,
                                      fontweight='bold', color='white')
                    if r > 0: ax.text(r / 2, i, f'{int(r)}', ha='center', va='center', fontsize=8, fontweight='bold',
                                      color='white')
                fig.tight_layout()
                chart_canvas = FigureCanvasTkAgg(fig, master=consumption_chart_frame)
                chart_canvas.draw()
                chart_canvas.get_tk_widget().pack(fill='both', expand=True)
            except Exception as e:
                tk.Label(consumption_chart_frame, text=f'⚠️ Chart error: {e}', font=('Arial', 10), bg='white',
                         fg='#e74c3c').pack(pady=30)

        # ===== ۱۱. چاپ گزارش =====
        def print_report():
            from datetime import datetime
            d = report_data
            lines = []
            lines.append('=' * 70)
            lines.append(' ' * 15 + 'PHARMACY SALES REPORT')
            lines.append('=' * 70)
            lines.append(f'  Period: {d["period"]}')
            lines.append(f'  Range: {d["from_date"]} to {d["to_date"]}')
            lines.append(f'  Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
            lines.append('')
            lines.append('-' * 70)
            lines.append('  SUMMARY')
            lines.append('-' * 70)
            lines.append(f'  Total Revenue:     {d["stats"]["total"]:>20,.0f} IRR')
            lines.append(f'  Direct Sales:      {d["stats"]["direct"]:>20,.0f} IRR')
            lines.append(f'  Dispensed Rx:      {d["stats"]["dispense"]:>20,.0f} IRR')
            lines.append(f'  Items Sold:        {d["stats"]["items"]:>20}')
            lines.append(f'  Transactions:      {d["stats"]["count"]:>20}')
            lines.append('')
            lines.append('-' * 70)
            lines.append('  TOP CONSUMED MEDICINES')
            lines.append('-' * 70)
            for i, c in enumerate(d['consumption'], 1):
                lines.append(f'  {i}. {c["name"][:40]:<40} Sold: {c["sold"]:<5}  Remaining: {c["remaining"]}')
            lines.append('')
            lines.append('=' * 70)

            success, pdf_path = print_text_file('\n'.join(lines), f'pharmacy_report_{d["period"]}.txt', open_pdf=False)
            if not success:
                messagebox.showerror('Error', f'❌ PDF error: {pdf_path}')
                return
            if WIN32_AVAILABLE:
                printer = get_printer_for('report')
                if printer and not is_virtual_printer(printer):
                    ok, msg = print_pdf_to_printer(pdf_path, printer)
                    if ok:
                        messagebox.showinfo('🖨️ Sent', f'Sent to: {printer}')
                        return
            try:
                os.startfile(pdf_path)
                messagebox.showinfo('📄 PDF Ready', f'File: {pdf_path}')
            except:
                pass

        # ===== ۱۲. خروجی اکسل =====
        def export_excel():
            from tkinter import filedialog
            from datetime import datetime
            try:
                from openpyxl import Workbook
                from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                from openpyxl.chart import BarChart, LineChart, Reference
                from openpyxl.chart.label import DataLabelList
            except ImportError:
                messagebox.showerror('Error', '❌ openpyxl not installed')
                return
            file_path = filedialog.asksaveasfilename(
                defaultextension='.xlsx',
                filetypes=[('Excel files', '*.xlsx')],
                initialfile=f'pharmacy_report_{report_data["period"]}_{datetime.now().strftime("%Y%m%d")}.xlsx'
            )
            if not file_path: return
            try:
                d = report_data
                wb = Workbook()
                ws1 = wb.active
                ws1.title = 'Summary'
                ws1.merge_cells('A1:C1')
                ws1['A1'] = f'Pharmacy Report - {d["period"]}'
                ws1['A1'].font = Font(size=16, bold=True, color='FFFFFF')
                ws1['A1'].fill = PatternFill('solid', fgColor='2C3E50')
                ws1['A1'].alignment = Alignment(horizontal='center', vertical='center')
                ws1.row_dimensions[1].height = 30
                ws1['A2'] = f'Range: {d["from_date"]} to {d["to_date"]}'
                ws1['A2'].font = Font(italic=True)
                row = 4
                ws1.cell(row=row, column=1, value='Metric').font = Font(bold=True, color='FFFFFF')
                ws1.cell(row=row, column=2, value='Value').font = Font(bold=True, color='FFFFFF')
                ws1.cell(row=row, column=1).fill = PatternFill('solid', fgColor='3498DB')
                ws1.cell(row=row, column=2).fill = PatternFill('solid', fgColor='3498DB')
                summary_rows = [
                    ('Total Revenue (IRR)', d['stats']['total']),
                    ('Direct Sales (IRR)', d['stats']['direct']),
                    ('Dispensed Revenue (IRR)', d['stats']['dispense']),
                    ('Items Sold', d['stats']['items']),
                    ('Transactions', d['stats']['count']),
                ]
                for i, (k, v) in enumerate(summary_rows, 1):
                    ws1.cell(row=row + i, column=1, value=k).font = Font(bold=True)
                    ws1.cell(row=row + i, column=2, value=v)
                    ws1.cell(row=row + i, column=2).number_format = '#,##0'
                ws1.column_dimensions['A'].width = 30
                ws1.column_dimensions['B'].width = 20

                ws2 = wb.create_sheet('Consumption')
                ws2['A1'] = 'Medicine'
                ws2['B1'] = 'Sold in Period'
                ws2['C1'] = 'Remaining Stock'
                for cell in ['A1', 'B1', 'C1']:
                    ws2[cell].font = Font(bold=True, color='FFFFFF')
                    ws2[cell].fill = PatternFill('solid', fgColor='2C3E50')
                for i, c in enumerate(d['consumption'], 2):
                    ws2.cell(row=i, column=1, value=c['name'])
                    ws2.cell(row=i, column=2, value=c['sold'])
                    ws2.cell(row=i, column=3, value=c['remaining'])
                ws2.column_dimensions['A'].width = 30
                ws2.column_dimensions['B'].width = 15
                ws2.column_dimensions['C'].width = 15
                if d['consumption']:
                    chart1 = BarChart()
                    chart1.type = "bar"
                    chart1.style = 10
                    chart1.title = "Medicine Stock Movement"
                    chart1.y_axis.title = "Medicine"
                    chart1.x_axis.title = "Quantity"
                    data = Reference(ws2, min_col=2, max_col=3, min_row=1, max_row=len(d['consumption']) + 1)
                    cats = Reference(ws2, min_col=1, min_row=2, max_row=len(d['consumption']) + 1)
                    chart1.add_data(data, titles_from_data=True)
                    chart1.set_categories(cats)
                    chart1.height = 12
                    chart1.width = 22
                    chart1.dataLabels = DataLabelList()
                    chart1.dataLabels.showVal = True
                    ws2.add_chart(chart1, "E2")

                ws3 = wb.create_sheet('Trend')
                ws3['A1'] = 'Date'
                ws3['B1'] = 'Direct Sales'
                ws3['C1'] = 'Dispensed'
                ws3['D1'] = 'Total'
                for cell in ['A1', 'B1', 'C1', 'D1']:
                    ws3[cell].font = Font(bold=True, color='FFFFFF')
                    ws3[cell].fill = PatternFill('solid', fgColor='2C3E50')
                for i, t in enumerate(d['trend'], 2):
                    ws3.cell(row=i, column=1, value=t['date'])
                    ws3.cell(row=i, column=2, value=t['direct'])
                    ws3.cell(row=i, column=3, value=t['dispense'])
                    ws3.cell(row=i, column=4, value=t['total'])
                for col in ['A', 'B', 'C', 'D']:
                    ws3.column_dimensions[col].width = 18
                if d['trend']:
                    chart2 = BarChart()
                    chart2.type = "col"
                    chart2.style = 12
                    chart2.title = "Revenue Trend (Direct vs Dispensed)"
                    chart2.y_axis.title = "Revenue (IRR)"
                    chart2.x_axis.title = "Date"
                    data = Reference(ws3, min_col=2, max_col=3, min_row=1, max_row=len(d['trend']) + 1)
                    cats = Reference(ws3, min_col=1, min_row=2, max_row=len(d['trend']) + 1)
                    chart2.add_data(data, titles_from_data=True)
                    chart2.set_categories(cats)
                    chart2.height = 10
                    chart2.width = 22
                    chart2.dataLabels = DataLabelList()
                    chart2.dataLabels.showVal = True
                    ws3.add_chart(chart2, "F2")

                ws4 = wb.create_sheet('Details')
                headers = ['Date', 'Type', 'Medicine', 'Qty', 'Unit Price', 'Total']
                for i, h in enumerate(headers, 1):
                    cell = ws4.cell(row=1, column=i, value=h)
                    cell.font = Font(bold=True, color='FFFFFF')
                    cell.fill = PatternFill('solid', fgColor='2C3E50')
                for i, row_data in enumerate(d['detail'], 2):
                    ws4.cell(row=i, column=1, value=row_data[0])
                    ws4.cell(row=i, column=2, value=row_data[1])
                    ws4.cell(row=i, column=3, value=row_data[2])
                    ws4.cell(row=i, column=4, value=row_data[3])
                    ws4.cell(row=i, column=5, value=row_data[4])
                    ws4.cell(row=i, column=6, value=row_data[5])
                for col in ['A', 'B', 'C', 'D', 'E', 'F']:
                    ws4.column_dimensions[col].width = 18

                wb.save(file_path)
                messagebox.showinfo('✅ Success', f'Excel saved with colorful charts!\n\n{file_path}')
                try:
                    os.startfile(file_path)
                except:
                    pass
            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        # ===== ۱۳. دکمه‌ها =====
        create_popup_buttons(btn_container, [
            {'text': '🔄 Refresh', 'command': load_report, 'color': 'primary', 'width': 140},
            {'text': '🖨️ Print Report', 'command': print_report, 'color': 'secondary', 'width': 180},
            {'text': '📊 Export Excel', 'command': export_excel, 'color': 'success', 'width': 180},
            {'text': '❌ Close', 'command': win.destroy, 'color': 'danger', 'width': 120},
        ], bg='#ecf0f1')

        # ===== ۱۴. اجرای اولیه =====
        win.update_idletasks()
        win.state('zoomed')
        load_report()


    def load_lab_stats(self):
        if not hasattr(self, 'lab_stat_filter_var'): return

        from_date, to_date = get_date_range(
            self.lab_stat_filter_var.get(),
            self.lab_stat_from_var.get().strip(),
            self.lab_stat_to_var.get().strip()
        )

        conn = get_connection()
        cursor = conn.cursor()

        date_cond = ""
        params = []
        if from_date:
            date_cond += " AND DATE(request_date, 'localtime') >= ?"
            params.append(from_date)
        if to_date:
            date_cond += " AND DATE(request_date, 'localtime') <= ?"
            params.append(to_date)

        # 1. درآمد آزمایش‌های آزاد (بدون پرونده)
        cursor.execute(f"""
            SELECT COALESCE(SUM(final_amount), 0) FROM lab_requests 
            WHERE record_id IS NULL AND payment_status = 'Paid' {date_cond}
        """, params)
        direct_rev = cursor.fetchone()[0]

        # 2. درآمد آزمایش‌های نسخه‌ای (دارای پرونده)
        cursor.execute(f"""
            SELECT COALESCE(SUM(final_amount), 0) FROM lab_requests 
            WHERE record_id IS NOT NULL AND payment_status = 'Paid' {date_cond}
        """, params)
        rx_rev = cursor.fetchone()[0]

        # 3. تعداد در انتظار پرداخت
        cursor.execute(f"""
            SELECT COUNT(*) FROM lab_requests 
            WHERE (payment_status IS NULL OR payment_status = 'Pending') AND status != 'Cancelled' {date_cond}
        """, params)
        pending_tests = cursor.fetchone()[0]

        conn.close()

        total_rev = direct_rev + rx_rev

        self.lab_stat_cards['total_revenue'](f'{total_rev:,.0f}')
        self.lab_stat_cards['direct_revenue'](f'{direct_rev:,.0f}')
        self.lab_stat_cards['rx_revenue'](f'{rx_rev:,.0f}')
        self.lab_stat_cards['pending_tests'](f'{pending_tests:,.0f}')


    def export_lab_stats_excel(self):
        """خروجی اکسل سریع از آمار آزمایشگاه"""
        from tkinter import filedialog
        from datetime import datetime

        from_date, to_date = get_date_range(
            self.lab_stat_filter_var.get(),
            self.lab_stat_from_var.get().strip(),
            self.lab_stat_to_var.get().strip()
        )

        file_path = filedialog.asksaveasfilename(
            defaultextension='.xlsx',
            filetypes=[('Excel files', '*.xlsx')],
            initialfile=f'lab_stats_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )
        if not file_path: return

        try:
            conn = get_connection()
            cursor = conn.cursor()

            date_cond = ""
            params = []
            if from_date:
                date_cond += " AND DATE(lr.request_date) >= ?"
                params.append(from_date)
            if to_date:
                date_cond += " AND DATE(lr.request_date) <= ?"
                params.append(to_date)

            cursor.execute(f'''
                SELECT lt.name, COUNT(lr.id) as cnt, COALESCE(SUM(lr.final_amount), 0) as rev
                FROM lab_requests lr
                JOIN lab_tests lt ON lr.test_id = lt.id
                WHERE 1=1 {date_cond}
                GROUP BY lt.id ORDER BY cnt DESC
            ''', params)
            data = cursor.fetchall()
            conn.close()

            headers = ['Test Name', 'Total Requests', 'Total Revenue (IRR)']
            rows = [(row[0], row[1], row[2]) for row in data]
            title = f"Lab Stats Report ({from_date or 'All'} to {to_date or 'Now'})"

            success, msg = export_to_excel(headers, rows, file_path, sheet_name='Lab Stats', title=title)

            if success:
                messagebox.showinfo('✅ Export Successful', f'Excel saved to:\n{msg}')
                try:
                    os.startfile(msg)
                except:
                    pass
            else:
                messagebox.showerror('Export Error', f'❌ {msg}')
        except Exception as e:
            messagebox.showerror('Export Error', f'❌ {e}')

    def lab_full_report_window(self):
        """پنجره گزارش جامع آزمایشگاه"""
        from datetime import datetime, timedelta

        win = tk.Toplevel(self.root)
        win.title('📊 Laboratory Report Center')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        btn_container = tk.Frame(win, bg='#ecf0f1')
        btn_container.pack(side='bottom', fill='x', pady=12)

        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=5, pady=5)

        canvas = tk.Canvas(body_container, bg='#ecf0f1', highlightthickness=0)
        v_scroll = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#ecf0f1')
        canvas.configure(yscrollcommand=v_scroll.set)
        canvas_win = canvas.create_window((0, 0), window=scrollable, anchor='nw')

        canvas.pack(side='left', fill='both', expand=True)
        v_scroll.pack(side='right', fill='y')

        def check_scroll(*args):
            canvas.configure(scrollregion=canvas.bbox('all'))
            bbox = canvas.bbox('all')
            if bbox:
                ch = bbox[3] - bbox[1]
                if ch > canvas.winfo_height() + 5:
                    if not v_scroll.winfo_ismapped(): v_scroll.pack(side='right', fill='y')
                else:
                    if v_scroll.winfo_ismapped(): v_scroll.pack_forget()

        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_win, width=event.width)
            check_scroll()

        scrollable.bind('<Configure>', check_scroll)
        canvas.bind('<Configure>', _on_canvas_configure)

        def on_mw(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mw)

        # ===== عنوان =====
        tk.Label(scrollable, text='📊 Laboratory Report Center',
                 font=('Arial', 18, 'bold'), bg='#ecf0f1', fg='#2c3e50').pack(pady=15)

        # ===== انتخاب دوره =====
        period_frame = tk.Frame(scrollable, bg='#ecf0f1')
        period_frame.pack(pady=5)

        tk.Label(period_frame, text='📅 Period:', font=('Arial', 12, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(side='left', padx=(0, 10))

        period_var = tk.StringVar(value='Daily')
        periods = [('Daily', '📅 Daily'), ('Weekly', '📆 Weekly'), ('Monthly', '🗓️ Monthly'), ('Yearly', '📅 Yearly')]

        for val, lbl in periods:
            tk.Radiobutton(period_frame, text=lbl, variable=period_var, value=val,
                           bg='#ecf0f1', font=('Arial', 10, 'bold'),
                           command=lambda: load_report(), activebackground='#ecf0f1').pack(side='left', padx=5)

        range_var = tk.StringVar(value='')
        tk.Label(scrollable, textvariable=range_var, font=('Arial', 11, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(pady=5)

        # ===== کارت‌های آمار =====
        cards_outer = tk.Frame(scrollable, bg='#ecf0f1')
        cards_outer.pack(fill='x', pady=10)

        cards_inner = tk.Frame(cards_outer, bg='#ecf0f1')
        cards_inner.pack()

        stat_updaters = {}
        CARD_COLOR = '#3b4a6b'

        for key, title in [
            ('total', '💰 Total Revenue'),
            ('count', '🧪 Total Tests'),
            ('paid', '✅ Paid'),
            ('pending', '⏳ Pending')
        ]:
            canvas_card, updater = create_rounded_stat_card(cards_inner, title, '0', bg_color=CARD_COLOR, width=180,
                                                            height=85)
            canvas_card.pack(side='left', padx=5)
            stat_updaters[key] = updater

        # ===== Notebook =====
        style = ttk.Style()
        style.theme_use('clam')
        notebook = ttk.Notebook(scrollable)
        notebook.pack(fill='both', expand=True, padx=20, pady=10)

        # Tab 1: Revenue Trend
        tab_trend = tk.Frame(notebook, bg='white')
        notebook.add(tab_trend, text='📈 Revenue Trend')
        trend_chart_frame = tk.Frame(tab_trend, bg='white')
        trend_chart_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Tab 2: Top Tests
        tab_tests = tk.Frame(notebook, bg='white')
        notebook.add(tab_tests, text='🧪 Top Tests')
        tests_chart_frame = tk.Frame(tab_tests, bg='white')
        tests_chart_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Tab 3: Details
        tab_detail = tk.Frame(notebook, bg='white')
        notebook.add(tab_detail, text='📋 Detailed List')

        detail_cols = ('Date', 'Patient', 'Test', 'Status', 'Amount')
        detail_frame = tk.Frame(tab_detail, bg='white')
        detail_frame.pack(fill='both', expand=True, padx=10, pady=10)
        detail_tree = ttk.Treeview(detail_frame, columns=detail_cols, show='headings', height=15)
        for c, w in zip(detail_cols, [120, 200, 250, 100, 120]):
            detail_tree.heading(c, text=c)
            detail_tree.column(c, width=w, anchor='center')
        d_scroll = ttk.Scrollbar(detail_frame, orient='vertical', command=detail_tree.yview)
        detail_tree.configure(yscrollcommand=d_scroll.set)
        detail_tree.pack(side='left', fill='both', expand=True)
        d_scroll.pack(side='right', fill='y')

        report_data = {'period': 'Daily', 'from_date': None, 'to_date': None, 'stats': {}, 'trend': [], 'top_tests': [],
                       'detail': []}

        def get_period_range(period):
            today = datetime.now().date()
            if period == 'Daily':
                return today.strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d'), 'Today'
            elif period == 'Weekly':
                return (today - timedelta(days=6)).strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d'), 'Last 7 Days'
            elif period == 'Monthly':
                return today.replace(day=1).strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d'), 'This Month'
            else:
                return today.replace(month=1, day=1).strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d'), 'This Year'

        def load_report():
            from_date, to_date, label = get_period_range(period_var.get())
            report_data['period'] = period_var.get()
            report_data['from_date'], report_data['to_date'] = from_date, to_date
            range_var.set(f'📅 {label}: {from_date} → {to_date}')

            try:
                conn = get_connection()
                cursor = conn.cursor()

                # 1. درآمد کل
                cursor.execute("""
                    SELECT COALESCE(SUM(final_amount),0), COUNT(*) 
                    FROM lab_payments
                    WHERE DATE(payment_date, 'localtime') BETWEEN ? AND ?
                """, (from_date, to_date))
                rev, paid_cnt = cursor.fetchone()

                # 2. تعداد کل آزمایشات
                cursor.execute("""
                    SELECT COUNT(*) FROM lab_requests 
                    WHERE DATE(request_date, 'localtime') BETWEEN ? AND ?
                """, (from_date, to_date))
                total_tests = cursor.fetchone()[0]

                # 3. تعداد آزمایشات در انتظار پرداخت
                cursor.execute("""
                    SELECT COUNT(*) FROM lab_requests 
                    WHERE status != 'Cancelled' 
                    AND (payment_status IS NULL OR payment_status = 'Pending') 
                    AND DATE(request_date, 'localtime') BETWEEN ? AND ?
                """, (from_date, to_date))
                pending_cnt = cursor.fetchone()[0]

                # 4. روند درآمد
                period = period_var.get()
                time_format = "%Y-%m" if period == "Yearly" else "%Y-%m-%d"

                cursor.execute(f"""
                    SELECT strftime('{time_format}', payment_date, 'localtime') as d, SUM(final_amount)
                    FROM lab_payments 
                    WHERE DATE(payment_date, 'localtime') BETWEEN ? AND ?
                    GROUP BY d ORDER BY d
                """, (from_date, to_date))
                trend = [{'date': k, 'total': v} for k, v in cursor.fetchall()]

                # 5. برترین آزمایشات
                cursor.execute("""
                    SELECT lt.name, COUNT(lr.id) as cnt
                    FROM lab_requests lr 
                    JOIN lab_tests lt ON lr.test_id = lt.id
                    WHERE DATE(lr.request_date, 'localtime') BETWEEN ? AND ?
                    GROUP BY lt.id ORDER BY cnt DESC LIMIT 10
                """, (from_date, to_date))
                top_tests = cursor.fetchall()

                # 6. جزئیات
                cursor.execute("""
                    SELECT lr.request_date, p.first_name || ' ' || p.last_name, lt.name, lr.status, lr.final_amount
                    FROM lab_requests lr
                    JOIN patients p ON lr.patient_id = p.id
                    JOIN lab_tests lt ON lr.test_id = lt.id
                    WHERE DATE(lr.request_date, 'localtime') BETWEEN ? AND ?
                    ORDER BY lr.request_date DESC LIMIT 200
                """, (from_date, to_date))
                detail = cursor.fetchall()

                conn.close()

                report_data['stats'] = {'total': rev, 'count': total_tests, 'paid': paid_cnt, 'pending': pending_cnt}
                report_data['trend'] = trend
                report_data['top_tests'] = top_tests
                report_data['detail'] = detail

                stat_updaters['total'](f'{rev:,.0f}')
                stat_updaters['count'](f'{total_tests}')
                stat_updaters['paid'](f'{paid_cnt}')
                stat_updaters['pending'](f'{pending_cnt}')

                draw_trend_chart()
                draw_top_tests_chart()

                for item in detail_tree.get_children(): detail_tree.delete(item)
                for row in detail:
                    detail_tree.insert('', 'end', values=(row[0][:16] if row[0] else '—', row[1], row[2], row[3],
                                                          f'{row[4] or 0:,.0f}'))
            except Exception as e:
                messagebox.showerror('Database Error', f'❌ Error: {e}')


        def draw_trend_chart():
            for w in trend_chart_frame.winfo_children(): w.destroy()
            if not MATPLOTLIB_AVAILABLE or not report_data['trend']: return
            try:
                fig = Figure(figsize=(10, 4.5), dpi=80)
                ax = fig.add_subplot(111)
                dates = [t['date'] for t in report_data['trend']]
                vals = [t['total'] for t in report_data['trend']]
                x = range(len(dates))
                bars = ax.bar(x, vals, color='#e74c3c', edgecolor='white', zorder=3)

                def fmt(v):
                    return f'{v / 1000000:.1f}M' if v >= 1000000 else f'{v / 1000:.0f}K' if v >= 1000 else f'{v:.0f}'

                for bar in bars:
                    h = bar.get_height()
                    if h > 0: ax.annotate(fmt(h), xy=(bar.get_x() + bar.get_width() / 2, h), xytext=(0, 3),
                                          textcoords="offset points", ha='center', va='bottom', fontsize=8,
                                          fontweight='bold')

                ax.set_xticks(x)
                if len(dates) > 15:
                    step = max(1, len(dates) // 10)
                    ax.set_xticks(list(x)[::step]);
                    ax.set_xticklabels([dates[i] for i in list(x)[::step]], rotation=45, ha='right')
                else:
                    ax.set_xticklabels(dates, rotation=45, ha='right')

                ax.set_title('Laboratory Revenue Trend', fontsize=12, fontweight='bold', color='#2c3e50')
                ax.set_ylabel('Revenue (IRR)')
                ax.grid(axis='y', alpha=0.3, linestyle='--', zorder=0)
                from matplotlib.ticker import FuncFormatter
                ax.yaxis.set_major_formatter(FuncFormatter(lambda v, t: fmt(v)))
                fig.tight_layout()
                c = FigureCanvasTkAgg(fig, master=trend_chart_frame);
                c.draw();
                c.get_tk_widget().pack(fill='both', expand=True)
            except Exception as e:
                pass

        def draw_top_tests_chart():
            for w in tests_chart_frame.winfo_children(): w.destroy()
            if not MATPLOTLIB_AVAILABLE or not report_data['top_tests']: return
            try:
                fig = Figure(figsize=(10, 4.5), dpi=80)
                ax = fig.add_subplot(111)
                names = [t[0][:25] for t in report_data['top_tests']][::-1]
                vals = [t[1] for t in report_data['top_tests']][::-1]
                y_pos = range(len(names))
                bars = ax.barh(y_pos, vals, color='#f39c12', edgecolor='white', zorder=3)
                ax.set_yticks(y_pos);
                ax.set_yticklabels(names, fontsize=9)
                for bar in bars:
                    w = bar.get_width()
                    ax.text(w, bar.get_y() + bar.get_height() / 2, f' {int(w)}', va='center', fontsize=9,
                            fontweight='bold')
                ax.set_title('Top Requested Lab Tests', fontsize=12, fontweight='bold', color='#2c3e50')
                ax.set_xlabel('Number of Requests')
                ax.grid(axis='x', alpha=0.3, linestyle='--', zorder=0)
                fig.tight_layout()
                c = FigureCanvasTkAgg(fig, master=tests_chart_frame);
                c.draw();
                c.get_tk_widget().pack(fill='both', expand=True)
            except Exception as e:
                pass

        def export_excel():
            from tkinter import filedialog
            try:
                from openpyxl import Workbook; from openpyxl.chart import BarChart, Reference
            except:
                return

            fp = filedialog.asksaveasfilename(defaultextension='.xlsx',
                                              initialfile=f'lab_report_{report_data["period"]}.xlsx')
            if not fp: return
            try:
                d = report_data;
                wb = Workbook()

                ws1 = wb.active;
                ws1.title = 'Trend'
                ws1.append(['Date', 'Revenue'])
                for t in d['trend']: ws1.append([t['date'], t['total']])
                if d['trend']:
                    c1 = BarChart();
                    c1.title = "Lab Revenue";
                    c1.type = "col";
                    c1.style = 10
                    data = Reference(ws1, min_col=2, min_row=1, max_row=len(d['trend']) + 1)
                    cats = Reference(ws1, min_col=1, min_row=2, max_row=len(d['trend']) + 1)
                    c1.add_data(data, titles_from_data=True);
                    c1.set_categories(cats)
                    ws1.add_chart(c1, "D2")

                ws2 = wb.create_sheet('Top Tests')
                ws2.append(['Test Name', 'Requests'])
                for t in d['top_tests']: ws2.append([t[0], t[1]])
                if d['top_tests']:
                    c2 = BarChart();
                    c2.title = "Top Tests";
                    c2.type = "bar";
                    c2.style = 11
                    data = Reference(ws2, min_col=2, min_row=1, max_row=len(d['top_tests']) + 1)
                    cats = Reference(ws2, min_col=1, min_row=2, max_row=len(d['top_tests']) + 1)
                    c2.add_data(data, titles_from_data=True);
                    c2.set_categories(cats)
                    ws2.add_chart(c2, "D2")

                ws3 = wb.create_sheet('Details')
                ws3.append(['Date', 'Patient', 'Test', 'Status', 'Amount'])
                for row in d['detail']: ws3.append(list(row))

                wb.save(fp);
                os.startfile(fp)
            except Exception as e:
                messagebox.showerror('Error', str(e))

        create_popup_buttons(btn_container, [
            {'text': '🔄 Refresh', 'command': load_report, 'color': 'primary', 'width': 140},
            {'text': '📊 Export Excel', 'command': export_excel, 'color': 'success', 'width': 180},
            {'text': '❌ Close', 'command': win.destroy, 'color': 'danger', 'width': 120},
        ], bg='#ecf0f1')

        win.update_idletasks()
        win.state('zoomed')
        load_report()

    def visit_full_report_window(self):
        """پنجره گزارش جامع مالی پرداخت‌ها (تفکیک ویزیت و پرداخت مستقیم)"""
        from datetime import datetime, timedelta

        win = tk.Toplevel(self.root)
        win.title('📊 Total Payments & Visits Report')
        win.configure(bg='#ecf0f1')
        win.resizable(True, True)

        btn_container = tk.Frame(win, bg='#ecf0f1')
        btn_container.pack(side='bottom', fill='x', pady=12)

        body_container = tk.Frame(win, bg='#ecf0f1')
        body_container.pack(fill='both', expand=True, padx=5, pady=5)

        canvas = tk.Canvas(body_container, bg='#ecf0f1', highlightthickness=0)
        v_scroll = ttk.Scrollbar(body_container, orient='vertical', command=canvas.yview)
        scrollable = tk.Frame(canvas, bg='#ecf0f1')
        canvas.configure(yscrollcommand=v_scroll.set)
        canvas_win = canvas.create_window((0, 0), window=scrollable, anchor='nw')

        canvas.pack(side='left', fill='both', expand=True)
        v_scroll.pack(side='right', fill='y')

        def check_scroll(*args):
            canvas.configure(scrollregion=canvas.bbox('all'))
            bbox = canvas.bbox('all')
            if bbox:
                ch = bbox[3] - bbox[1]
                if ch > canvas.winfo_height() + 5:
                    if not v_scroll.winfo_ismapped(): v_scroll.pack(side='right', fill='y')
                else:
                    if v_scroll.winfo_ismapped(): v_scroll.pack_forget()

        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_win, width=event.width)
            check_scroll()

        scrollable.bind('<Configure>', check_scroll)
        canvas.bind('<Configure>', _on_canvas_configure)

        def on_mw(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        canvas.bind_all('<MouseWheel>', on_mw)

        tk.Label(scrollable, text='📊 Total Payments & Visits Report',
                 font=('Arial', 18, 'bold'), bg='#ecf0f1', fg='#2c3e50').pack(pady=15)

        period_frame = tk.Frame(scrollable, bg='#ecf0f1')
        period_frame.pack(pady=5)

        tk.Label(period_frame, text='📅 Period:', font=('Arial', 12, 'bold'),
                 bg='#ecf0f1', fg='#2c3e50').pack(side='left', padx=(0, 10))

        period_var = tk.StringVar(value='Daily')
        periods = [('Daily', '📅 Daily'), ('Weekly', '📆 Weekly'), ('Monthly', '🗓️ Monthly'), ('Yearly', '📅 Yearly')]

        for val, lbl in periods:
            tk.Radiobutton(period_frame, text=lbl, variable=period_var, value=val,
                           bg='#ecf0f1', font=('Arial', 10, 'bold'),
                           command=lambda: load_report(), activebackground='#ecf0f1').pack(side='left', padx=5)

        range_var = tk.StringVar(value='')
        tk.Label(scrollable, textvariable=range_var, font=('Arial', 11, 'italic'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(pady=5)

        # ===== کارت‌های آمار (5 کارت) =====
        cards_outer = tk.Frame(scrollable, bg='#ecf0f1')
        cards_outer.pack(fill='x', pady=10)

        cards_inner = tk.Frame(cards_outer, bg='#ecf0f1')
        cards_inner.pack()

        stat_updaters = {}
        CARD_COLOR = '#3b4a6b'

        card_configs = [
            ('total', '💰 Total Revenue'),
            ('visits', '🩺 Medical Visits'),
            ('direct', '💳 Direct Payments'),
            ('cash', '💵 Cash'),
            ('card', '💳 Card/Ins')
        ]

        for key, title in card_configs:
            canvas_card, updater = create_rounded_stat_card(
                cards_inner, title, '0', bg_color=CARD_COLOR, width=170, height=85
            )
            canvas_card.pack(side='left', padx=4)
            stat_updaters[key] = updater

        count_label_var = tk.StringVar(value='')
        tk.Label(cards_outer, textvariable=count_label_var, font=('Arial', 10, 'bold'),
                 bg='#ecf0f1', fg='#7f8c8d').pack(pady=(8, 0))

        # ===== Notebook =====
        style = ttk.Style()
        style.theme_use('clam')
        notebook = ttk.Notebook(scrollable)
        notebook.pack(fill='both', expand=True, padx=20, pady=10)

        tab_trend = tk.Frame(notebook, bg='white')
        notebook.add(tab_trend, text='📈 Revenue Trend')
        trend_chart_frame = tk.Frame(tab_trend, bg='white')
        trend_chart_frame.pack(fill='both', expand=True, padx=10, pady=10)

        tab_methods = tk.Frame(notebook, bg='white')
        notebook.add(tab_methods, text='💳 Payment Methods')
        methods_chart_frame = tk.Frame(tab_methods, bg='white')
        methods_chart_frame.pack(fill='both', expand=True, padx=10, pady=10)

        tab_detail = tk.Frame(notebook, bg='white')
        notebook.add(tab_detail, text='📋 Detailed List')

        detail_cols = ('Date', 'Receipt', 'Type', 'Patient', 'Amount', 'Discount', 'Final', 'Method')
        detail_frame = tk.Frame(tab_detail, bg='white')
        detail_frame.pack(fill='both', expand=True, padx=10, pady=10)
        detail_tree = ttk.Treeview(detail_frame, columns=detail_cols, show='headings', height=15)
        for c, w in zip(detail_cols, [130, 160, 100, 180, 90, 90, 90, 80]):
            detail_tree.heading(c, text=c)
            detail_tree.column(c, width=w, anchor='center')
        d_scroll = ttk.Scrollbar(detail_frame, orient='vertical', command=detail_tree.yview)
        detail_tree.configure(yscrollcommand=d_scroll.set)
        detail_tree.pack(side='left', fill='both', expand=True)
        d_scroll.pack(side='right', fill='y')

        report_data = {
            'period': 'Daily',
            'from_date': None, 'to_date': None,
            'stats': {}, 'trend': [], 'methods': {}, 'detail': [],
        }

        def get_period_range(period):
            today = datetime.now().date()
            if period == 'Daily':
                return today.strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d'), 'Today'
            elif period == 'Weekly':
                return (today - timedelta(days=6)).strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d'), 'Last 7 Days'
            elif period == 'Monthly':
                return today.replace(day=1).strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d'), 'This Month'
            else:
                return today.replace(month=1, day=1).strftime('%Y-%m-%d'), today.strftime('%Y-%m-%d'), 'This Year'

        # ===== تابع بارگذاری داده‌ها =====
        def load_report():
            from_date, to_date, label = get_period_range(period_var.get())
            report_data['period'] = period_var.get()
            report_data['from_date'] = from_date
            report_data['to_date'] = to_date
            range_var.set(f'📅 {label}: {from_date} → {to_date}')

            try:
                conn = get_connection()
                cursor = conn.cursor()

                # 1. آمار کل
                cursor.execute('''
                    SELECT COALESCE(SUM(final_amount), 0), COUNT(*)
                    FROM payments WHERE DATE(payment_date) BETWEEN ? AND ?
                ''', (from_date, to_date))
                total_rev, total_cnt = cursor.fetchone()

                # 2. ویزیت پزشکی (VIS-)
                cursor.execute('''
                    SELECT COALESCE(SUM(final_amount), 0), COUNT(*)
                    FROM payments WHERE receipt_number LIKE 'VIS-%' AND DATE(payment_date) BETWEEN ? AND ?
                ''', (from_date, to_date))
                visit_rev, visit_cnt = cursor.fetchone()

                # 3. پرداخت آزاد (RCP-)
                cursor.execute('''
                    SELECT COALESCE(SUM(final_amount), 0), COUNT(*)
                    FROM payments WHERE receipt_number LIKE 'RCP-%' AND DATE(payment_date) BETWEEN ? AND ?
                ''', (from_date, to_date))
                direct_rev, direct_cnt = cursor.fetchone()

                # ✅ چاپ در ترمینال برای تست
                print(f"DEBUG REPORT: Found {direct_cnt} Direct Payments with total {direct_rev}")

                # 4. نقدی و کارت
                cursor.execute('''
                    SELECT 
                        COALESCE(SUM(CASE WHEN payment_method = 'Cash' THEN final_amount ELSE 0 END), 0),
                        COALESCE(SUM(CASE WHEN payment_method != 'Cash' THEN final_amount ELSE 0 END), 0)
                    FROM payments WHERE DATE(payment_date) BETWEEN ? AND ?
                ''', (from_date, to_date))
                cash_rev, card_rev = cursor.fetchone()

                # 5. روند (Trend)
                period = period_var.get()
                time_format = "%Y-%m" if period == "Yearly" else "%Y-%m-%d"
                cursor.execute(f'''
                    SELECT strftime('{time_format}', payment_date) as d, 
                           SUM(CASE WHEN receipt_number LIKE 'VIS-%' THEN final_amount ELSE 0 END) as vis,
                           SUM(CASE WHEN receipt_number LIKE 'RCP-%' THEN final_amount ELSE 0 END) as rcp,
                           SUM(final_amount) as total
                    FROM payments WHERE DATE(payment_date) BETWEEN ? AND ?
                    GROUP BY d ORDER BY d
                ''', (from_date, to_date))
                trend = [{'date': k, 'visit': v, 'direct': r, 'total': t} for k, v, r, t in cursor.fetchall()]

                # 6. روش‌های پرداخت
                cursor.execute('''
                    SELECT payment_method, SUM(final_amount)
                    FROM payments WHERE DATE(payment_date) BETWEEN ? AND ?
                    GROUP BY payment_method
                ''', (from_date, to_date))
                methods = {m: v for m, v in cursor.fetchall()}

                # 7. جزئیات
                cursor.execute('''
                    SELECT p.payment_date, p.receipt_number, COALESCE(pt.first_name || ' ' || pt.last_name, '— (deleted)'), 
                           p.amount, p.discount, p.final_amount, p.payment_method
                    FROM payments p
                    LEFT JOIN patients pt ON p.patient_id = pt.id
                    WHERE DATE(p.payment_date) BETWEEN ? AND ?
                    ORDER BY p.payment_date DESC LIMIT 200
                ''', (from_date, to_date))
                detail = cursor.fetchall()

                conn.close()

                report_data['stats'] = {
                    'total': total_rev, 'total_cnt': total_cnt,
                    'visit': visit_rev, 'visit_cnt': visit_cnt,
                    'direct': direct_rev, 'direct_cnt': direct_cnt,
                    'cash': cash_rev, 'card': card_rev,
                }
                report_data['trend'] = trend
                report_data['methods'] = methods
                report_data['detail'] = detail

                stat_updaters['total'](f'{total_rev:,.0f}')
                stat_updaters['visits'](f'{visit_rev:,.0f}')
                stat_updaters['direct'](f'{direct_rev:,.0f}')
                stat_updaters['cash'](f'{cash_rev:,.0f}')
                stat_updaters['card'](f'{card_rev:,.0f}')

                count_label_var.set(
                    f"📊 {total_cnt} Total  |  🩺 {visit_cnt} Medical Visits  |  💳 {direct_cnt} Direct Payments"
                )

                draw_trend_chart()
                draw_methods_chart()

                for item in detail_tree.get_children():
                    detail_tree.delete(item)
                for row in detail:
                    receipt = row[1] or ''
                    if receipt.startswith('VIS-'):
                        ptype = '🩺 Visit'
                    elif receipt.startswith('RCP-'):
                        ptype = '💳 Direct'
                    else:
                        ptype = '—'

                    detail_tree.insert('', 'end', values=(
                        row[0][:16] if row[0] else '—', receipt, ptype, row[2],
                        f'{row[3]:,.0f}', f'{row[4]:,.0f}', f'{row[5]:,.0f}', row[6]
                    ))
            except Exception as e:
                messagebox.showerror('Database Error', f'❌ Error loading report:\n\n{e}')


        def draw_trend_chart():
            for w in trend_chart_frame.winfo_children(): w.destroy()
            if not MATPLOTLIB_AVAILABLE or not report_data['trend']:
                tk.Label(trend_chart_frame, text='📭 No data', font=('Arial', 12, 'italic'),
                         bg='white', fg='#95a5a6').pack(pady=50)
                return
            try:
                fig = Figure(figsize=(10, 4.5), dpi=80)
                ax = fig.add_subplot(111)
                dates = [t['date'] for t in report_data['trend']]
                vis_vals = [t['visit'] for t in report_data['trend']]
                rcp_vals = [t['direct'] for t in report_data['trend']]

                x = range(len(dates))
                width = 0.35
                bars1 = ax.bar([i - width / 2 for i in x], vis_vals, width, label='🩺 Medical Visits', color='#27ae60',
                               edgecolor='white', zorder=3)
                bars2 = ax.bar([i + width / 2 for i in x], rcp_vals, width, label='💳 Direct Payments', color='#3498db',
                               edgecolor='white', zorder=3)

                def fmt(v):
                    return f'{v / 1000000:.1f}M' if v >= 1000000 else f'{v / 1000:.0f}K' if v >= 1000 else f'{v:.0f}'

                for bars in [bars1, bars2]:
                    for bar in bars:
                        h = bar.get_height()
                        if h > 0:
                            ax.annotate(fmt(h), xy=(bar.get_x() + bar.get_width() / 2, h),
                                        xytext=(0, 3), textcoords="offset points",
                                        ha='center', va='bottom', fontsize=8, fontweight='bold')

                ax.set_xticks(x)
                if len(dates) > 15:
                    step = max(1, len(dates) // 10)
                    ax.set_xticks(list(x)[::step])
                    ax.set_xticklabels([dates[i] for i in list(x)[::step]], rotation=45, ha='right')
                else:
                    ax.set_xticklabels(dates, rotation=45, ha='right')

                ax.set_title('Revenue Trend: Visits vs Direct Payments', fontsize=12, fontweight='bold',
                             color='#2c3e50')
                ax.set_ylabel('Revenue (IRR)')
                ax.grid(axis='y', alpha=0.3, linestyle='--', zorder=0)
                ax.legend(loc='upper left', fontsize=9)
                ax.spines['top'].set_visible(False)
                ax.spines['right'].set_visible(False)
                from matplotlib.ticker import FuncFormatter
                ax.yaxis.set_major_formatter(FuncFormatter(lambda v, t: fmt(v)))
                fig.tight_layout()
                c = FigureCanvasTkAgg(fig, master=trend_chart_frame);
                c.draw();
                c.get_tk_widget().pack(fill='both', expand=True)
            except Exception as e:
                pass

        def draw_methods_chart():
            for w in methods_chart_frame.winfo_children(): w.destroy()
            if not MATPLOTLIB_AVAILABLE or not report_data['methods']:
                tk.Label(methods_chart_frame, text='📭 No data', font=('Arial', 12, 'italic'),
                         bg='white', fg='#95a5a6').pack(pady=50)
                return
            try:
                fig = Figure(figsize=(8, 5), dpi=80)
                ax = fig.add_subplot(111)
                labels = list(report_data['methods'].keys())
                sizes = list(report_data['methods'].values())
                colors = ['#27ae60', '#3498db', '#9b59b6', '#f39c12', '#e74c3c']
                wedges, texts, autotexts = ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', startangle=90,
                                                  wedgeprops=dict(width=0.4, edgecolor='w'))
                for text in texts: text.set_fontsize(11)
                for autotext in autotexts: autotext.set_fontsize(10); autotext.set_fontweight('bold')
                ax.axis('equal')
                ax.set_title('Payment Methods Breakdown', fontsize=12, fontweight='bold', color='#2c3e50')
                fig.tight_layout()
                c = FigureCanvasTkAgg(fig, master=methods_chart_frame);
                c.draw();
                c.get_tk_widget().pack(fill='both', expand=True)
            except Exception as e:
                pass

        def print_report():
            from datetime import datetime
            d = report_data
            lines = []
            lines.append('=' * 70)
            lines.append(' ' * 10 + 'TOTAL PAYMENTS & VISITS REPORT')
            lines.append('=' * 70)
            lines.append(f'  Period: {d["period"]}')
            lines.append(f'  Range: {d["from_date"]} to {d["to_date"]}')
            lines.append(f'  Generated: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
            lines.append('')
            lines.append('-' * 70)
            lines.append('  OVERALL')
            lines.append('-' * 70)
            lines.append(f'  Total Revenue:      {d["stats"]["total"]:>18,.0f} IRR')
            lines.append(f'  Total Trans:        {d["stats"]["total_cnt"]:>18}')
            lines.append('')
            lines.append('-' * 70)
            lines.append('  MEDICAL VISITS (VIS-)')
            lines.append('-' * 70)
            lines.append(f'  Amount:             {d["stats"]["visit"]:>18,.0f} IRR')
            lines.append(f'  Count:              {d["stats"]["visit_cnt"]:>18}')
            lines.append('')
            lines.append('-' * 70)
            lines.append('  DIRECT PAYMENTS (RCP-)')
            lines.append('-' * 70)
            lines.append(f'  Amount:             {d["stats"]["direct"]:>18,.0f} IRR')
            lines.append(f'  Count:              {d["stats"]["direct_cnt"]:>18}')
            lines.append('')
            lines.append('-' * 70)
            lines.append('  PAYMENT METHODS')
            lines.append('-' * 70)
            lines.append(f'  Cash:               {d["stats"]["cash"]:>18,.0f} IRR')
            lines.append(f'  Card/Insurance:     {d["stats"]["card"]:>18,.0f} IRR')
            lines.append('')
            lines.append('=' * 70)

            success, pdf_path = print_text_file('\n'.join(lines), f'payments_report_{d["period"]}.txt', open_pdf=False)
            if not success:
                messagebox.showerror('Error', f'❌ PDF error: {pdf_path}')
                return
            if WIN32_AVAILABLE:
                printer = get_printer_for('report')
                if printer and not is_virtual_printer(printer):
                    ok, msg = print_pdf_to_printer(pdf_path, printer)
                    if ok:
                        messagebox.showinfo('🖨️ Sent', f'Sent to: {printer}')
                        return
            try:
                os.startfile(pdf_path)
                messagebox.showinfo('📄 PDF Ready', f'File: {pdf_path}')
            except:
                pass

        def export_excel():
            from tkinter import filedialog
            from datetime import datetime
            try:
                from openpyxl import Workbook
                from openpyxl.styles import Font, PatternFill, Alignment
                from openpyxl.chart import BarChart, PieChart, Reference
                from openpyxl.chart.label import DataLabelList
            except ImportError:
                messagebox.showerror('Error', '❌ openpyxl not installed')
                return
            file_path = filedialog.asksaveasfilename(
                defaultextension='.xlsx',
                filetypes=[('Excel files', '*.xlsx')],
                initialfile=f'payments_report_{report_data["period"]}_{datetime.now().strftime("%Y%m%d")}.xlsx'
            )
            if not file_path: return
            try:
                d = report_data
                wb = Workbook()

                # ===== Sheet 1: Summary =====
                ws1 = wb.active
                ws1.title = 'Summary'
                ws1.merge_cells('A1:C1')
                ws1['A1'] = f'Payments & Visits Report - {d["period"]}'
                ws1['A1'].font = Font(size=16, bold=True, color='FFFFFF')
                ws1['A1'].fill = PatternFill('solid', fgColor='2C3E50')
                ws1['A1'].alignment = Alignment(horizontal='center', vertical='center')
                ws1.row_dimensions[1].height = 30
                ws1['A2'] = f'Range: {d["from_date"]} to {d["to_date"]}'
                ws1['A2'].font = Font(italic=True)

                row = 4
                headers = ['Category', 'Count', 'Amount (IRR)']
                for i, h in enumerate(headers, 1):
                    cell = ws1.cell(row=row, column=i, value=h)
                    cell.font = Font(bold=True, color='FFFFFF')
                    cell.fill = PatternFill('solid', fgColor='3498DB')

                data_rows = [
                    ('🩺 Medical Visits', d['stats']['visit_cnt'], d['stats']['visit']),
                    ('💳 Direct Payments', d['stats']['direct_cnt'], d['stats']['direct']),
                    ('', '', ''),
                    ('💰 TOTAL', d['stats']['total_cnt'], d['stats']['total']),
                    ('', '', ''),
                    ('💵 Cash', '', d['stats']['cash']),
                    ('💳 Card/Insurance', '', d['stats']['card']),
                ]
                for i, (cat, cnt, amt) in enumerate(data_rows, 1):
                    r = row + i
                    ws1.cell(row=r, column=1, value=cat).font = Font(bold=(cat == '💰 TOTAL'))
                    ws1.cell(row=r, column=2, value=cnt)
                    cell_amt = ws1.cell(row=r, column=3, value=amt)
                    cell_amt.number_format = '#,##0'
                    if cat == '💰 TOTAL':
                        for c in range(1, 4):
                            ws1.cell(row=r, column=c).fill = PatternFill('solid', fgColor='FFF3CD')
                            ws1.cell(row=r, column=c).font = Font(bold=True)

                ws1.column_dimensions['A'].width = 25
                ws1.column_dimensions['B'].width = 12
                ws1.column_dimensions['C'].width = 20

                # ===== Sheet 2: Trend =====
                ws2 = wb.create_sheet('Trend')
                headers2 = ['Date', 'Medical Visits', 'Direct Payments', 'Total']
                for i, h in enumerate(headers2, 1):
                    cell = ws2.cell(row=1, column=i, value=h)
                    cell.font = Font(bold=True, color='FFFFFF')
                    cell.fill = PatternFill('solid', fgColor='2C3E50')
                for i, t in enumerate(d['trend'], 2):
                    ws2.cell(row=i, column=1, value=t['date'])
                    ws2.cell(row=i, column=2, value=t['visit']).number_format = '#,##0'
                    ws2.cell(row=i, column=3, value=t['direct']).number_format = '#,##0'
                    ws2.cell(row=i, column=4, value=t['total']).number_format = '#,##0'
                for col in ['A', 'B', 'C', 'D']:
                    ws2.column_dimensions[col].width = 20
                if d['trend']:
                    chart1 = BarChart()
                    chart1.type = "col"
                    chart1.grouping = "clustered"
                    chart1.style = 10
                    chart1.title = "Revenue Trend (Visits vs Direct)"
                    chart1.y_axis.title = "Revenue (IRR)"
                    data = Reference(ws2, min_col=2, max_col=3, min_row=1, max_row=len(d['trend']) + 1)
                    cats = Reference(ws2, min_col=1, min_row=2, max_row=len(d['trend']) + 1)
                    chart1.add_data(data, titles_from_data=True)
                    chart1.set_categories(cats)
                    chart1.height = 10
                    chart1.width = 22
                    chart1.dataLabels = DataLabelList()
                    chart1.dataLabels.showVal = True
                    ws2.add_chart(chart1, "F2")

                # ===== Sheet 3: Methods =====
                ws3 = wb.create_sheet('Methods')
                ws3['A1'] = 'Method'
                ws3['B1'] = 'Amount'
                for cell in ['A1', 'B1']:
                    ws3[cell].font = Font(bold=True, color='FFFFFF')
                    ws3[cell].fill = PatternFill('solid', fgColor='2C3E50')
                r = 2
                for k, v in d['methods'].items():
                    ws3.cell(row=r, column=1, value=k)
                    ws3.cell(row=r, column=2, value=v).number_format = '#,##0'
                    r += 1
                if d['methods']:
                    chart2 = PieChart()
                    chart2.title = "Payment Methods"
                    data = Reference(ws3, min_col=2, min_row=1, max_row=r - 1)
                    cats = Reference(ws3, min_col=1, min_row=2, max_row=r - 1)
                    chart2.add_data(data, titles_from_data=True)
                    chart2.set_categories(cats)
                    chart2.height = 8
                    chart2.width = 14
                    ws3.add_chart(chart2, "D2")

                # ===== Sheet 4: Details =====
                ws4 = wb.create_sheet('Details')
                headers4 = ['Date', 'Receipt', 'Type', 'Patient', 'Amount', 'Discount', 'Final', 'Method']
                for i, h in enumerate(headers4, 1):
                    cell = ws4.cell(row=1, column=i, value=h)
                    cell.font = Font(bold=True, color='FFFFFF')
                    cell.fill = PatternFill('solid', fgColor='2C3E50')
                for i, row_data in enumerate(d['detail'], 2):
                    receipt = row_data[1] or ''
                    if receipt.startswith('VIS-'):
                        ptype = 'Medical Visit'
                    elif receipt.startswith('RCP-'):
                        ptype = 'Direct Payment'
                    else:
                        ptype = '—'
                    ws4.cell(row=i, column=1, value=row_data[0])
                    ws4.cell(row=i, column=2, value=receipt)
                    ws4.cell(row=i, column=3, value=ptype)
                    ws4.cell(row=i, column=4, value=row_data[2])
                    ws4.cell(row=i, column=5, value=row_data[3]).number_format = '#,##0'
                    ws4.cell(row=i, column=6, value=row_data[4]).number_format = '#,##0'
                    ws4.cell(row=i, column=7, value=row_data[5]).number_format = '#,##0'
                    ws4.cell(row=i, column=8, value=row_data[6])
                for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']:
                    ws4.column_dimensions[col].width = 18

                wb.save(file_path)
                messagebox.showinfo('✅ Success', f'Excel saved!\n\n{file_path}')
                try:
                    os.startfile(file_path)
                except:
                    pass
            except Exception as e:
                messagebox.showerror('Error', f'❌ {e}')

        create_popup_buttons(btn_container, [
            {'text': '🔄 Refresh', 'command': load_report, 'color': 'primary', 'width': 140},
            {'text': '🖨️ Print Report', 'command': print_report, 'color': 'secondary', 'width': 180},
            {'text': '📊 Export Excel', 'command': export_excel, 'color': 'success', 'width': 180},
            {'text': '❌ Close', 'command': win.destroy, 'color': 'danger', 'width': 120},
        ], bg='#ecf0f1')

        win.update_idletasks()
        win.state('zoomed')
        load_report()






# ==========================================
# Run Application
# ==========================================


if __name__ == '__main__':
    import sys
    import os

    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
        os.chdir(application_path)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))

    if application_path not in sys.path:
        sys.path.insert(0, application_path)

    from login_window import LoginWindow

    login_root = tk.Tk()

    def open_main_app():
        user_info = login_window.logged_in_user
        login_root.destroy()

        main_root = tk.Tk()
        main_root.current_user = user_info

        # ===== فول اسکرین =====
        main_root.state('zoomed')  # برای ویندوز: تمام صفحه
        # یا:
        # main_root.attributes('-fullscreen', True)  # فول اسکرین کامل

        app = HospitalApp(main_root)

        title_text = f'🏥 Hospital Management System - {user_info["full_name"]} ({user_info["role"].upper()})'
        main_root.title(title_text)

        main_root.mainloop()

    login_window = LoginWindow(login_root, open_main_app)
    login_root.mainloop()