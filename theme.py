"""مدیریت تم (رنگ‌ها)"""

# ===== تم روشن =====
LIGHT_THEME = {
    'bg': '#f5f7fa',           # پس‌زمینه‌ی اصلی
    'bg_secondary': '#ffffff',  # پس‌زمینه‌ی کارت‌ها
    'bg_tertiary': '#f8f9fa',   # پس‌زمینه‌ی محتوا
    'fg': '#2c3e50',           # متن اصلی
    'fg_secondary': '#7f8c8d',  # متن ثانویه
    'header_bg': '#3b4a6b',     # هدر
    'header_fg': '#ffffff',     # متن هدر
    'border': '#bdc3c7',        # خطوط
    'entry_bg': '#ffffff',      # کادر ورودی
    'entry_fg': '#2c3e50',
    'tree_bg': '#ffffff',       # جدول
    'tree_fg': '#2c3e50',
    'tree_alt': '#f8f9fa',      # ردیف‌های جایگزین
    'tree_selected': '#3498db', # ردیف انتخاب‌شده
    'notebook_bg': '#eef1f7',
    'notebook_tab': '#d5dae5',
    'notebook_tab_fg': '#3b4a6b',
}

# ===== تم تیره =====
DARK_THEME = {
    'bg': '#1e1e2e',           # پس‌زمینه‌ی اصلی
    'bg_secondary': '#2a2a3e',  # پس‌زمینه‌ی کارت‌ها
    'bg_tertiary': '#252535',   # پس‌زمینه‌ی محتوا
    'fg': '#e0e0e0',           # متن اصلی
    'fg_secondary': '#95a5a6',  # متن ثانویه
    'header_bg': '#16213e',     # هدر
    'header_fg': '#ffffff',     # متن هدر
    'border': '#44475a',        # خطوط
    'entry_bg': '#2a2a3e',      # کادر ورودی
    'entry_fg': '#e0e0e0',
    'tree_bg': '#252535',       # جدول
    'tree_fg': '#e0e0e0',
    'tree_alt': '#2a2a3e',      # ردیف‌های جایگزین
    'tree_selected': '#3498db',
    'notebook_bg': '#1e1e2e',
    'notebook_tab': '#2a2a3e',
    'notebook_tab_fg': '#95a5a6',
}


def get_theme(mode='light'):
    """گرفتن تم بر اساس حالت"""
    if mode == 'dark':
        return DARK_THEME
    return LIGHT_THEME