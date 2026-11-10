import tkinter as tk
from tkinter import messagebox
import sqlite3
import hashlib
import secrets

DB_NAME = 'hospital.db'


def get_connection():
    return sqlite3.connect(DB_NAME)


def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()


# ==========================================
# Rounded Button Helper
# ==========================================
def create_rounded_button(parent, text, command, bg_color, hover_color,
                          text_color='white', width=340, height=50,
                          corner_radius=25, font=('Arial', 12, 'bold'),
                          bg_parent='white', pack_kwargs=None):
    """ساخت دکمه گِرد با استفاده از Canvas"""
    if pack_kwargs is None:
        pack_kwargs = {'pady': 10, 'fill': 'x'}

    canvas = tk.Canvas(parent, width=width, height=height, bg=bg_parent, highlightthickness=0)
    canvas.pack(**pack_kwargs)

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
        canvas.config(cursor='hand2')

    def on_leave(e):
        canvas.itemconfig(rect, fill=bg_color, outline=bg_color)

    def on_click(e):
        command()

    canvas.tag_bind('btn', '<Enter>', on_enter)
    canvas.tag_bind('btn', '<Leave>', on_leave)
    canvas.tag_bind('btn', '<Button-1>', on_click)
    canvas.config(cursor='hand2')

    return canvas


class LoginWindow:
    """Login window for the hospital system"""

    def __init__(self, root, on_success):
        self.root = root
        self.on_success = on_success
        self.logged_in_user = None

        self.root.title('🏥 Hospital System - Login')
        self.root.geometry('500x600')
        self.root.configure(bg='#2c3e50')
        self.root.resizable(False, False)

        # وسط‌چین
        self.root.update_idletasks()
        w, h = 500, 600
        x = (self.root.winfo_screenwidth() // 2) - (w // 2)
        y = (self.root.winfo_screenheight() // 2) - (h // 2)
        self.root.geometry(f'{w}x{h}+{x}+{y}')

        self.create_widgets()

    def create_widgets(self):
        # کارت سفید وسط
        card = tk.Frame(self.root, bg='white', relief='flat', bd=0)
        card.place(relx=0.5, rely=0.5, anchor='center', width=420, height=540)

        # آیکون
        tk.Label(card, text='🏥', font=('Arial', 60),
                 bg='white').pack(pady=(25, 10))

        # عنوان
        tk.Label(card, text='Hospital Management System',
                 font=('Arial', 14, 'bold'), bg='white', fg='#2c3e50').pack()

        tk.Label(card, text='Please login to continue',
                 font=('Arial', 10), bg='white', fg='#7f8c8d').pack(pady=(5, 15))

        # خط جداکننده
        tk.Frame(card, bg='#ecf0f1', height=1).pack(fill='x', padx=40, pady=10)

        # ==== Username ====
        tk.Label(card, text='👤 Username:', font=('Arial', 11, 'bold'),
                 bg='white', anchor='w').pack(anchor='w', padx=40, pady=(10, 5))

        self.username_entry = tk.Entry(card, font=('Arial', 12),
                                       bg='#f8f9fa', relief='solid', bd=1)
        self.username_entry.pack(padx=40, pady=5, fill='x', ipady=8)

        # ==== Password ====
        tk.Label(card, text='🔒 Password:', font=('Arial', 11, 'bold'),
                 bg='white', anchor='w').pack(anchor='w', padx=40, pady=(10, 5))

        self.password_entry = tk.Entry(card, font=('Arial', 12), show='●',
                                       bg='#f8f9fa', relief='solid', bd=1)
        self.password_entry.pack(padx=40, pady=5, fill='x', ipady=8)

        # ==== Show Password Checkbox ====
        self.show_pwd_var = tk.BooleanVar(value=False)
        tk.Checkbutton(card, text='Show password', variable=self.show_pwd_var,
                       command=self.toggle_password, bg='white',
                       font=('Arial', 9)).pack(anchor='w', padx=40, pady=(5, 10))

        # ==== Login Button (استفاده از دکمه گِرد) ====
        btn_frame = tk.Frame(card, bg='white')
        btn_frame.pack(padx=40, pady=(5, 5), fill='x')

        create_rounded_button(
            btn_frame,
            text='🔐 LOGIN',
            command=self.do_login,
            bg_color='#27ae60',
            hover_color='#229954',
            width=340, height=50,
            corner_radius=25,
            font=('Arial', 13, 'bold'),
            bg_parent='white',
            pack_kwargs={'fill': 'x'}  # پر کردن عرض
        )

        # ==== Forgot Password Link ====
        forgot_label = tk.Label(card, text='🔑 Forgot your password?',
                                font=('Arial', 10, 'underline'),
                                bg='white', fg='#3498db', cursor='hand2')
        forgot_label.pack(pady=(10, 10))
        forgot_label.bind('<Button-1>', lambda e: self.open_forgot_password_window())

        # ==== Status Label ====
        self.status_var = tk.StringVar(value='')
        self.status_label = tk.Label(card, textvariable=self.status_var,
                                     font=('Arial', 10, 'italic'),
                                     bg='white', fg='#e74c3c',
                                     wraplength=340, justify='center')
        self.status_label.pack(pady=(5, 10))

        # ==== Info ====
        tk.Label(card, text='Default credentials shown in terminal',
                 font=('Arial', 8, 'italic'), bg='white', fg='#95a5a6').pack(pady=(5, 15))

        # Bind Enter key
        self.username_entry.bind('<Return>', lambda e: self.password_entry.focus())
        self.password_entry.bind('<Return>', lambda e: self.do_login())

        # Focus on username
        self.username_entry.focus()

    def toggle_password(self):
        """Show or hide password"""
        if self.show_pwd_var.get():
            self.password_entry.config(show='')
        else:
            self.password_entry.config(show='●')

    def do_login(self):
        """Attempt login with precise error messages"""
        username = self.username_entry.get().strip()
        password = self.password_entry.get()

        # پاک کردن پیام قبلی
        self.status_var.set('')

        # ===== اعتبارسنجی ورودی =====
        if not username and not password:
            self.status_var.set('❌ Please enter username and password!')
            self.username_entry.focus()
            return

        if not username:
            self.status_var.set('❌ Please enter username!')
            self.username_entry.focus()
            return

        if not password:
            self.status_var.set('❌ Please enter password!')
            self.password_entry.focus()
            return

        # ===== چک دیتابیس =====
        conn = get_connection()
        cursor = conn.cursor()

        # اول چک یوزرنیم
        cursor.execute('''
            SELECT id, username, full_name, role, is_active, password_hash
            FROM users
            WHERE username = ?
        ''', (username,))
        user = cursor.fetchone()

        # کاربر پیدا نشد
        if not user:
            self.status_var.set(f'❌ Username "{username}" not found!')
            self.status_label.config(fg='#e74c3c')
            self.username_entry.select_range(0, 'end')
            self.username_entry.focus()
            conn.close()
            return

        # چک فعال بودن
        if user[4] != 1:
            self.status_var.set('❌ Account is disabled! Contact admin.')
            self.status_label.config(fg='#e74c3c')
            conn.close()
            return

        # چک پسورد
        if user[5] != hash_password(password):
            self.status_var.set('❌ Incorrect password! Try again.')
            self.status_label.config(fg='#e74c3c')
            self.password_entry.delete(0, 'end')
            self.password_entry.focus()
            conn.close()
            return

        # ===== موفق =====
        cursor.execute('UPDATE users SET last_login = CURRENT_TIMESTAMP WHERE id = ?',
                       (user[0],))
        conn.commit()
        conn.close()

        self.logged_in_user = {
            'id': user[0],
            'username': user[1],
            'full_name': user[2],
            'role': user[3],
        }

        try:
            from hospital_gui_full import log_action
            log_action(
                'LOGIN',
                entity_type='user',
                entity_id=user[0],
                details=f'User "{user[1]}" logged in',
                user={'id': user[0], 'username': user[1], 'role': user[3]}
            )
        except Exception as e:
            print(f"Log action error: {e}")

        self.status_var.set(f'✅ Welcome, {user[2]}!')
        self.status_label.config(fg='#27ae60')

        self.root.after(800, self.on_success)

    # ==========================================
    # Forgot Password Window
    # ==========================================

    def open_forgot_password_window(self):
        """Open the forgot password / recovery window"""
        win = tk.Toplevel(self.root)
        win.title('🔑 Password Recovery')
        win.geometry('450x700')
        win.configure(bg='white')
        win.resizable(False, False)

        win.update_idletasks()
        w, h = 450, 700
        x = (win.winfo_screenwidth() // 2) - (w // 2)
        y = (win.winfo_screenheight() // 2) - (h // 2)
        win.geometry(f'{w}x{h}+{x}+{y}')

        # آیکون
        tk.Label(win, text='🔑', font=('Arial', 40),
                 bg='white').pack(pady=(20, 5))

        tk.Label(win, text='Password Recovery',
                 font=('Arial', 14, 'bold'), bg='white',
                 fg='#2c3e50').pack()

        tk.Label(win, text='Enter your username and recovery code',
                 font=('Arial', 9), bg='white', fg='#7f8c8d').pack(pady=(5, 15))

        tk.Frame(win, bg='#ecf0f1', height=1).pack(fill='x', padx=40, pady=5)

        # ===== Username =====
        tk.Label(win, text='👤 Username:', font=('Arial', 10, 'bold'),
                 bg='white', anchor='w').pack(anchor='w', padx=40, pady=(15, 3))
        username_entry = tk.Entry(win, font=('Arial', 11),
                                  bg='#f8f9fa', relief='solid', bd=1)
        username_entry.pack(padx=40, pady=3, fill='x', ipady=6)
        username_entry.insert(0, self.username_entry.get().strip())  # پیش‌پر کردن

        # ===== Recovery Code =====
        tk.Label(win, text='🆘 Recovery Code:', font=('Arial', 10, 'bold'),
                 bg='white', anchor='w').pack(anchor='w', padx=40, pady=(10, 3))
        code_entry = tk.Entry(win, font=('Arial', 11),
                              bg='#f8f9fa', relief='solid', bd=1)
        code_entry.pack(padx=40, pady=3, fill='x', ipady=6)

        tk.Label(win, text='💡 Format: 8 characters (e.g., A3F9B2C1)',
                 font=('Arial', 8, 'italic'), bg='white',
                 fg='#95a5a6').pack(anchor='w', padx=40)

        # ===== New Password =====
        tk.Label(win, text='🔒 New Password:', font=('Arial', 10, 'bold'),
                 bg='white', anchor='w').pack(anchor='w', padx=40, pady=(10, 3))
        new_pwd_entry = tk.Entry(win, font=('Arial', 11), show='●',
                                 bg='#f8f9fa', relief='solid', bd=1)
        new_pwd_entry.pack(padx=40, pady=3, fill='x', ipady=6)

        # ===== Confirm Password =====
        tk.Label(win, text='🔒 Confirm New Password:', font=('Arial', 10, 'bold'),
                 bg='white', anchor='w').pack(anchor='w', padx=40, pady=(10, 3))
        confirm_entry = tk.Entry(win, font=('Arial', 11), show='●',
                                 bg='#f8f9fa', relief='solid', bd=1)
        confirm_entry.pack(padx=40, pady=3, fill='x', ipady=6)

        # ===== Status =====
        status_var = tk.StringVar(value='')
        status_label = tk.Label(win, textvariable=status_var,
                                font=('Arial', 9, 'italic'),
                                bg='white', fg='#e74c3c',
                                wraplength=370, justify='center')
        status_label.pack(pady=(10, 5))

        # ===== Reset Button =====
        def do_reset():
            status_var.set('')
            username = username_entry.get().strip()
            code = code_entry.get().strip().upper()
            new_pwd = new_pwd_entry.get()
            confirm = confirm_entry.get()

            # اعتبارسنجی
            if not username:
                status_var.set('❌ Please enter username!')
                return
            if not code:
                status_var.set('❌ Please enter recovery code!')
                return
            if len(new_pwd) < 6:
                status_var.set('❌ Password must be at least 6 characters!')
                return
            if new_pwd != confirm:
                status_var.set('❌ Passwords do not match!')
                return

            # چک دیتابیس
            conn = get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, recovery_code FROM users
                WHERE username = ?
            ''', (username,))
            user = cursor.fetchone()

            if not user:
                status_var.set(f'❌ Username "{username}" not found!')
                conn.close()
                return

            if not user[1]:
                status_var.set('❌ No recovery code set! Contact admin.')
                conn.close()
                return

            if user[1] != code:
                status_var.set('❌ Invalid recovery code!')
                conn.close()
                return

            # ===== همه چی درسته → رمز رو عوض کن =====
            new_hash = hash_password(new_pwd)
            # کد بازیابی جدید هم بساز (برای امنیت بیشتر)
            new_code = secrets.token_hex(4).upper()

            cursor.execute('''
                UPDATE users
                SET password_hash = ?, recovery_code = ?
                WHERE id = ?
            ''', (new_hash, new_code, user[0]))
            conn.commit()
            conn.close()

            messagebox.showinfo(
                '✅ Success',
                f'Password reset successfully!\n\n'
                f'👤 Username: {username}\n'
                f'🔑 New Password: {new_pwd}\n'
                f'🆘 New Recovery Code: {new_code}\n\n'
                f'⚠️ Save the new Recovery Code somewhere safe!\n'
                f'You will need it if you forget the password again.'
            )

            # برگرد به صفحه لاگین و پیش‌پر کن
            win.destroy()
            self.username_entry.delete(0, 'end')
            self.username_entry.insert(0, username)
            self.password_entry.delete(0, 'end')
            self.password_entry.focus()
            self.status_var.set('✅ Password reset! Please login with new password.')
            self.status_label.config(fg='#27ae60')

        # ==== Reset Button (گِرد) ====
        btn_reset_frame = tk.Frame(win, bg='white')
        btn_reset_frame.pack(padx=40, pady=(10, 5), fill='x')

        create_rounded_button(
            btn_reset_frame,
            text='🔑 Reset Password',
            command=do_reset,
            bg_color='#3498db',
            hover_color='#2980b9',
            width=370, height=45,
            corner_radius=20,
            font=('Arial', 11, 'bold'),
            bg_parent='white',
            pack_kwargs={'fill': 'x'}
        )

        # ==== Cancel Button (گِرد) ====
        btn_cancel_frame = tk.Frame(win, bg='white')
        btn_cancel_frame.pack(padx=40, pady=(5, 10), fill='x')

        create_rounded_button(
            btn_cancel_frame,
            text='❌ Cancel',
            command=win.destroy,
            bg_color='#e74c3c',
            hover_color='#c0392b',
            width=370, height=45,
            corner_radius=20,
            font=('Arial', 11, 'bold'),
            bg_parent='white',
            pack_kwargs={'fill': 'x'}
        )

        # Bind Enter
        code_entry.bind('<Return>', lambda e: new_pwd_entry.focus())
        new_pwd_entry.bind('<Return>', lambda e: confirm_entry.focus())
        confirm_entry.bind('<Return>', lambda e: do_reset())

        code_entry.focus()