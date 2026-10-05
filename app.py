import sys
import os
import sqlite3
import shutil
from datetime import datetime

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QStackedWidget, QTableWidget,
    QTableWidgetItem, QHeaderView, QMessageBox, QComboBox, QSpinBox,
    QDoubleSpinBox, QDateEdit, QFormLayout, QGroupBox, QFileDialog,
    QFrame, QSizePolicy
)
from PyQt6.QtCore import Qt, QDate
from PyQt6.QtGui import QFont, QIcon, QColor

# کتابخانه‌های اختیاری جهت خروجی Excel، PDF و نمودار
try:
    import openpyxl
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

try:
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False


# ==========================================
# 1. مدیریت پایگاه داده (SQLite Database)
# ==========================================
class DatabaseManager:
    def __init__(self, db_name="imart_studio.db"):
        self.db_name = db_name
        self.init_db()

    def get_connection(self):
        return sqlite3.connect(self.db_name)

    def init_db(self):
        conn = self.get_connection()
        cursor = conn.cursor()

        # تنظیمات برنامه (رمز عبور)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        ''')
        cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('password', '1234')")

        # پروژه‌های عروس و داماد
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS wedding_contracts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT,
                groom_name TEXT,
                bride_name TEXT,
                phone TEXT,
                event_date TEXT,
                total_amount REAL,
                discount REAL,
                deposit_1 REAL,
                remaining_balance REAL,
                details TEXT,
                created_at TEXT
            )
        ''')

        # دریافتی‌ها و اقساط پروژه‌ها
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS project_payments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                contract_id INTEGER,
                amount REAL,
                payment_date TEXT,
                bank_card TEXT,
                description TEXT
            )
        ''')

        # سایر پروژه‌ها (تبلیغاتی، بیوتی، تولد)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS other_projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT,
                category TEXT,
                client_name TEXT,
                phone TEXT,
                event_date TEXT,
                camera_count INTEGER,
                total_amount REAL,
                received_amount REAL,
                remaining_amount REAL,
                details TEXT
            )
        ''')

        # هزینه‌ها و دستمزد نیروها
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS expenses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT,
                category TEXT,
                role TEXT,
                person_name TEXT,
                amount REAL,
                expense_date TEXT,
                description TEXT
            )
        ''')

        # انبار تجهیزات
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS inventory (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_name TEXT,
                category TEXT,
                total_quantity INTEGER,
                in_use_quantity INTEGER,
                description TEXT
            )
        ''')

        # کارت‌های بانکی
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS bank_cards (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                bank_name TEXT,
                card_number TEXT,
                sheba TEXT,
                owner_name TEXT
            )
        ''')

        conn.commit()
        conn.close()

    def check_password(self, pwd):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM settings WHERE key='password'")
        res = cursor.fetchone()
        conn.close()
        return res and res[0] == pwd

    def set_password(self, new_pwd):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("UPDATE settings SET value=? WHERE key='password'", (new_pwd,))
        conn.commit()
        conn.close()


# ==========================================
# 2. ویجت صفحه مشکی ورود (Embedded Dark Overlay)
# ==========================================
class OverlayLoginWidget(QWidget):
    def __init__(self, db, on_login_success):
        super().__init__()
        self.db = db
        self.on_login_success = on_login_success
        self.init_ui()

    def init_ui(self):
        # استایل مشکی شیک و مینیمال برای لایه ورود
        self.setStyleSheet("""
            QWidget {
                background-color: #0b0c10;
                color: #ffffff;
            }
            QGroupBox {
                border: 2px solid #45a29e;
                border-radius: 12px;
                background-color: #1f2833;
                margin-top: 10px;
                font-size: 14px;
            }
            QLabel {
                color: #c5c6c7;
                font-size: 15px;
            }
            QLineEdit {
                background-color: #0b0c10;
                border: 1px solid #45a29e;
                border-radius: 6px;
                padding: 10px;
                color: #66fcf1;
                font-size: 18px;
            }
            QLineEdit:focus {
                border: 2px solid #66fcf1;
            }
            QPushButton {
                background-color: #45a29e;
                color: #1f2833;
                font-weight: bold;
                font-size: 16px;
                border: none;
                border-radius: 6px;
                padding: 12px;
            }
            QPushButton:hover {
                background-color: #66fcf1;
            }
        """)

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        box = QGroupBox("ورود به سیستم حسابداری IMART STUDIO")
        box.setFixedWidth(420)
        box_layout = QVBoxLayout()
        box_layout.setSpacing(20)

        title = QLabel("🔒 ورود به برنامه")
        title.setFont(QFont("B Yekan", 18, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        box_layout.addWidget(title)

        desc = QLabel("جهت دسترسی به اطلاعات مالی، رمز عبور را وارد کنید:")
        desc.setWordWrap(True)
        desc.setAlignment(Qt.AlignmentFlag.AlignCenter)
        box_layout.addWidget(desc)

        self.txt_pass = QLineEdit()
        self.txt_pass.setEchoMode(QLineEdit.EchoMode.Password)
        self.txt_pass.setPlaceholderText("رمز عبور (پیش‌فرض: 1234)")
        self.txt_pass.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.txt_pass.returnPressed.connect(self.verify)
        box_layout.addWidget(self.txt_pass)

        btn_login = QPushButton("تأیید و ورود")
        btn_login.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_login.clicked.connect(self.verify)
        box_layout.addWidget(btn_login)

        box.setLayout(box_layout)
        layout.addWidget(box)
        self.setLayout(layout)

    def verify(self):
        pwd = self.txt_pass.text().strip()
        if self.db.check_password(pwd):
            self.txt_pass.clear()
            self.on_login_success()
        else:
            QMessageBox.critical(self, "خطا", "رمز عبور اشتباه است!")


# ==========================================
# 3. پنجره اصلی نرم‌افزار (MainWindow)
# ==========================================
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = DatabaseManager()

        self.setWindowTitle("IMART STUDIO - سیستم مدیریت مالی و حسابداری")
        self.resize(1200, 750)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)

        # ویجت مرکزی
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)

        # لایه‌بندی اصلی برنامه
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0, 0, 0, 0)

        # 1. ساخت بخش اصلی برنامه (شامل منوها و صفحات)
        self.app_content_widget = QWidget()
        self.setup_app_ui()

        # 2. ساخت لایه مشکی ورود
        self.login_overlay = OverlayLoginWidget(self.db, self.unlock_app)

        # افزودن هر دو به لایه اصلی
        self.main_layout.addWidget(self.login_overlay)
        self.main_layout.addWidget(self.app_content_widget)

        # در ابتدا فقط صفحه مشکی نمایش داده می‌شود
        self.app_content_widget.hide()
        self.login_overlay.show()

    def unlock_app(self):
        """پس از ورود رمز درست، صفحه مشکی پنهان و برنامه اصلی فعال می‌شود"""
        self.login_overlay.hide()
        self.app_content_widget.show()
        self.refresh_dashboard()

    def setup_app_ui(self):
        """ایجاد ظاهر و بخش‌های اصلی برنامه"""
        layout = QHBoxLayout(self.app_content_widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # استایل تم تاریک اختصاصی آی‌مارت استودیو
        self.app_content_widget.setStyleSheet("""
            QWidget {
                background-color: #1e1e2e;
                color: #cdd6f4;
                font-family: 'B Yekan', 'Tahoma', sans-serif;
                font-size: 13px;
            }
            QFrame#Sidebar {
                background-color: #181825;
                border-left: 1px solid #313244;
            }
            QPushButton.NavBtn {
                background-color: transparent;
                color: #cdd6f4;
                text-align: right;
                padding: 12px 18px;
                border: none;
                border-radius: 8px;
                font-size: 14px;
            }
            QPushButton.NavBtn:hover {
                background-color: #313244;
                color: #89b4fa;
            }
            QPushButton.NavBtn:checked {
                background-color: #89b4fa;
                color: #11111b;
                font-weight: bold;
            }
            QTableWidget {
                background-color: #181825;
                gridline-color: #313244;
                border: 1px solid #313244;
                border-radius: 8px;
            }
            QHeaderView::section {
                background-color: #313244;
                color: #cdd6f4;
                padding: 8px;
                font-weight: bold;
                border: none;
            }
            QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QDateEdit {
                background-color: #313244;
                border: 1px solid #45475a;
                border-radius: 6px;
                padding: 6px 10px;
                color: #cdd6f4;
            }
            QPushButton.ActionBtn {
                background-color: #89b4fa;
                color: #11111b;
                border-radius: 6px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton.ActionBtn:hover {
                background-color: #b4befe;
            }
            QGroupBox {
                border: 1px solid #45475a;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 10px;
                font-weight: bold;
            }
        """)

        # منوی کناری (Sidebar)
        sidebar = QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setFixedWidth(220)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(10, 20, 10, 20)

        brand_label = QLabel("IMART STUDIO")
        brand_label.setFont(QFont("B Yekan", 16, QFont.Weight.Bold))
        brand_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        brand_label.setStyleSheet("color: #89b4fa; margin-bottom: 20px;")
        sidebar_layout.addWidget(brand_label)

        self.btn_dash = QPushButton("📊 داشبورد مالی")
        self.btn_wedding = QPushButton("💍 پروژه‌های عروس و داماد")
        self.btn_other = QPushButton("🎬 سایر پروژه‌ها")
        self.btn_expense = QPushButton("💸 هزینه‌ها و دستمزدها")
        self.btn_inventory = QPushButton("🎥 انبار تجهیزات")
        self.btn_cards = QPushButton("💳 کارت‌های بانکی")
        self.btn_settings = QPushButton("⚙️ تنظیمات و پشتیبان")

        self.nav_buttons = [
            self.btn_dash, self.btn_wedding, self.btn_other,
            self.btn_expense, self.btn_inventory, self.btn_cards, self.btn_settings
        ]

        for idx, btn in enumerate(self.nav_buttons):
            btn.setCheckable(True)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setProperty("class", "NavBtn")
            btn.clicked.connect(lambda checked, i=idx: self.switch_page(i))
            sidebar_layout.addWidget(btn)

        sidebar_layout.addStretch()

        btn_lock = QPushButton("🔒 قفل مجدد برنامه")
        btn_lock.setStyleSheet("background-color: #f38ba8; color: #11111b; font-weight: bold; padding: 8px; border-radius: 6px;")
        btn_lock.clicked.connect(self.lock_app)
        sidebar_layout.addWidget(btn_lock)

        layout.addWidget(sidebar)

        # بخش اصلی پشته‌ای (Stacked Widget)
        self.pages = QStackedWidget()
        layout.addWidget(self.pages)

        # ساخت صفحات
        self.page_dash = self.create_dash_page()
        self.page_wedding = self.create_wedding_page()
        self.page_other = self.create_other_page()
        self.page_expense = self.create_expense_page()
        self.page_inventory = self.create_inventory_page()
        self.page_cards = self.create_cards_page()
        self.page_settings = self.create_settings_page()

        self.pages.addWidget(self.page_dash)
        self.pages.addWidget(self.page_wedding)
        self.pages.addWidget(self.page_other)
        self.pages.addWidget(self.page_expense)
        self.pages.addWidget(self.page_inventory)
        self.pages.addWidget(self.page_cards)
        self.pages.addWidget(self.page_settings)

        self.switch_page(0)

    def switch_page(self, index):
        for btn in self.nav_buttons:
            btn.setChecked(False)
        self.nav_buttons[index].setChecked(True)
        self.pages.setCurrentIndex(index)

    def lock_app(self):
        """قفل کردن مجدد برنامه و نمایش صفحه مشکی"""
        self.app_content_widget.hide()
        self.login_overlay.show()

    # ==========================================
    # 4. ایجاد صفحات نرم‌افزار
    # ==========================================
    def create_dash_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)

        header = QLabel("داشبورد خلاصه وضعیت مالی")
        header.setFont(QFont("B Yekan", 16, QFont.Weight.Bold))
        layout.addWidget(header)

        # کارت‌های آمار
        cards_layout = QHBoxLayout()

        self.lbl_income = QLabel("درآمد کل: 0 تومان")
        self.lbl_expense = QLabel("مجموع هزینه‌ها: 0 تومان")
        self.lbl_profit = QLabel("سود خالص: 0 تومان")

        for lbl, color in zip([self.lbl_income, self.lbl_expense, self.lbl_profit], ["#a6e3a1", "#f38ba8", "#89b4fa"]):
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setFont(QFont("B Yekan", 13, QFont.Weight.Bold))
            lbl.setStyleSheet(f"background-color: #313244; color: {color}; border-radius: 10px; padding: 20px;")
            cards_layout.addWidget(lbl)

        layout.addLayout(cards_layout)

        # بخش نمودار
        if HAS_MATPLOTLIB:
            self.fig, self.ax = plt.subplots(figsize=(5, 3))
            self.fig.patch.set_facecolor('#1e1e2e')
            self.canvas = FigureCanvas(self.fig)
            layout.addWidget(self.canvas)
        else:
            layout.addWidget(QLabel("کتابخانه Matplotlib جهت نمایش نمودار نصب نیست."))

        return page

    def refresh_dashboard(self):
        conn = self.db.get_connection()
        cursor = conn.cursor()

        # محاسبه مجموع ورودی‌ها
        cursor.execute("SELECT SUM(received_amount) FROM other_projects")
        r1 = cursor.fetchone()[0] or 0

        cursor.execute("SELECT SUM(deposit_1) FROM wedding_contracts")
        r2 = cursor.fetchone()[0] or 0

        cursor.execute("SELECT SUM(amount) FROM project_payments")
        r3 = cursor.fetchone()[0] or 0

        total_income = r1 + r2 + r3

        # محاسبه هزینه‌ها
        cursor.execute("SELECT SUM(amount) FROM expenses")
        total_expense = cursor.fetchone()[0] or 0

        profit = total_income - total_expense

        self.lbl_income.setText(f"درآمد کل:\n{total_income:,.0f} تومان")
        self.lbl_expense.setText(f"مجموع هزینه‌ها:\n{total_expense:,.0f} تومان")
        self.lbl_profit.setText(f"سود خالص:\n{profit:,.0f} تومان")

        if HAS_MATPLOTLIB:
            self.ax.clear()
            self.ax.set_facecolor('#1e1e2e')
            if total_income > 0 or total_expense > 0:
                labels = ['درآمد', 'هزینه']
                sizes = [total_income, total_expense]
                colors = ['#a6e3a1', '#f38ba8']
                self.ax.pie(sizes, labels=labels, colors=colors, autopct='%1.1f%%', textprops={'color': 'white'})
            else:
                self.ax.text(0.5, 0.5, 'داده‌ای موجود نیست', color='white', ha='center', va='center')
            self.canvas.draw()

        conn.close()

    def create_wedding_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)

        header = QLabel("مدیریت قراردادهای عروس و داماد")
        header.setFont(QFont("B Yekan", 14, QFont.Weight.Bold))
        layout.addWidget(header)

        # فرم ثبت قرارداد
        form = QFormLayout()
        self.w_title = QLineEdit()
        self.w_groom = QLineEdit()
        self.w_bride = QLineEdit()
        self.w_phone = QLineEdit()
        self.w_date = QDateEdit(QDate.currentDate())
        self.w_total = QDoubleSpinBox()
        self.w_total.setMaximum(10000000000)
        self.w_discount = QDoubleSpinBox()
        self.w_discount.setMaximum(1000000000)
        self.w_deposit = QDoubleSpinBox()
        self.w_deposit.setMaximum(1000000000)

        form.addRow("عنوان قرارداد:", self.w_title)
        form.addRow("نام داماد:", self.w_groom)
        form.addRow("نام عروس:", self.w_bride)
        form.addRow("شماره تماس:", self.w_phone)
        form.addRow("تاریخ مراسم:", self.w_date)
        form.addRow("مبلغ کل (تومان):", self.w_total)
        form.addRow("تخفیف (تومان):", self.w_discount)
        form.addRow("بیعانه اول (تومان):", self.w_deposit)

        btn_add = QPushButton("ثبت قرارداد جدید")
        btn_add.setProperty("class", "ActionBtn")
        btn_add.clicked.connect(self.save_wedding_contract)

        layout.addLayout(form)
        layout.addWidget(btn_add)

        # جدول قراردادها
        self.table_wedding = QTableWidget(0, 7)
        self.table_wedding.setHorizontalHeaderLabels(["ID", "عنوان", "داماد", "عروس", "مبلغ کل", "مانده حساب", "تاریخ"])
        self.table_wedding.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table_wedding)

        self.load_wedding_contracts()
        return page

    def save_wedding_contract(self):
        title = self.w_title.text()
        groom = self.w_groom.text()
        bride = self.w_bride.text()
        phone = self.w_phone.text()
        date = self.w_date.date().toString("yyyy-MM-dd")
        total = self.w_total.value()
        discount = self.w_discount.value()
        deposit = self.w_deposit.value()
        remaining = total - discount - deposit

        conn = self.db.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO wedding_contracts (title, groom_name, bride_name, phone, event_date, total_amount, discount, deposit_1, remaining_balance, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (title, groom, bride, phone, date, total, discount, deposit, remaining, datetime.now().strftime("%Y-%m-%d")))
        conn.commit()
        conn.close()

        QMessageBox.information(self, "موفقیت", "قرارداد با موفقیت ثبت شد.")
        self.load_wedding_contracts()

    def load_wedding_contracts(self):
        self.table_wedding.setRowCount(0)
        conn = self.db.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id, title, groom_name, bride_name, total_amount, remaining_balance, event_date FROM wedding_contracts")
        for row_idx, row_data in enumerate(cursor.fetchall()):
            self.table_wedding.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                self.table_wedding.setItem(row_idx, col_idx, QTableWidgetItem(str(value)))
        conn.close()

    def create_other_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)

        header = QLabel("پروژه‌های تبلیغاتی، بیوتی و تولد")
        header.setFont(QFont("B Yekan", 14, QFont.Weight.Bold))
        layout.addWidget(header)

        form = QFormLayout()
        self.o_title = QLineEdit()
        self.o_cat = QComboBox()
        self.o_cat.addItems(["تبلیغاتی", "بیوتی", "تولد", "سایر"])
        self.o_client = QLineEdit()
        self.o_cameras = QSpinBox()
        self.o_total = QDoubleSpinBox()
        self.o_total.setMaximum(10000000000)
        self.o_received = QDoubleSpinBox()
        self.o_received.setMaximum(10000000000)

        form.addRow("عنوان پروژه:", self.o_title)
        form.addRow("دسته بندی:", self.o_cat)
        form.addRow("نام کارفرما:", self.o_client)
        form.addRow("تعداد دوربین استفاده شده:", self.o_cameras)
        form.addRow("مبلغ کل (تومان):", self.o_total)
        form.addRow("مبلغ دریافتی (تومان):", self.o_received)

        btn_add = QPushButton("ثبت پروژه")
        btn_add.setProperty("class", "ActionBtn")
        btn_add.clicked.connect(self.save_other_project)

        layout.addLayout(form)
        layout.addWidget(btn_add)

        self.table_other = QTableWidget(0, 6)
        self.table_other.setHorizontalHeaderLabels(["ID", "عنوان", "دسته", "کارفرما", "دوربین‌ها", "مانده"])
        self.table_other.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table_other)

        self.load_other_projects()
        return page

    def save_other_project(self):
        title = self.o_title.text()
        cat = self.o_cat.currentText()
        client = self.o_client.text()
        cams = self.o_cameras.value()
        total = self.o_total.value()
        rec = self.o_received.value()
        rem = total - rec

        conn = self.db.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO other_projects (title, category, client_name, camera_count, total_amount, received_amount, remaining_amount)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (title, cat, client, cams, total, rec, rem))

        # به‌روزرسانی انبار دوربین‌ها
        if cams > 0:
            cursor.execute("UPDATE inventory SET in_use_quantity = in_use_quantity + ? WHERE category='دوربین'", (cams,))

        conn.commit()
        conn.close()

        QMessageBox.information(self, "موفقیت", "پروژه ثبت شد.")
        self.load_other_projects()

    def load_other_projects(self):
        self.table_other.setRowCount(0)
        conn = self.db.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id, title, category, client_name, camera_count, remaining_amount FROM other_projects")
        for row_idx, row_data in enumerate(cursor.fetchall()):
            self.table_other.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                self.table_other.setItem(row_idx, col_idx, QTableWidgetItem(str(value)))
        conn.close()

    def create_expense_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)

        header = QLabel("ثبت هزینه‌ها و دستمزد نیروها")
        header.setFont(QFont("B Yekan", 14, QFont.Weight.Bold))
        layout.addWidget(header)

        form = QFormLayout()
        self.ex_title = QLineEdit()
        self.ex_role = QComboBox()
        self.ex_role.addItems(["تدوینگر", "عکاس", "فیلمبردار", "کرین‌کار", "هلی‌شات‌کار", "سایر"])
        self.ex_person = QLineEdit()
        self.ex_amount = QDoubleSpinBox()
        self.ex_amount.setMaximum(1000000000)

        form.addRow("عنوان هزینه:", self.ex_title)
        form.addRow("نقش / سمت:", self.ex_role)
        form.addRow("نام شخص / گیرنده:", self.ex_person)
        form.addRow("مبلغ (تومان):", self.ex_amount)

        btn_add = QPushButton("ثبت هزینه")
        btn_add.setProperty("class", "ActionBtn")
        btn_add.clicked.connect(self.save_expense)

        layout.addLayout(form)
        layout.addWidget(btn_add)

        self.table_expense = QTableWidget(0, 5)
        self.table_expense.setHorizontalHeaderLabels(["ID", "عنوان", "نقش", "نام", "مبلغ"])
        self.table_expense.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table_expense)

        self.load_expenses()
        return page

    def save_expense(self):
        title = self.ex_title.text()
        role = self.ex_role.currentText()
        person = self.ex_person.text()
        amount = self.ex_amount.value()

        conn = self.db.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO expenses (title, role, person_name, amount, expense_date)
            VALUES (?, ?, ?, ?, ?)
        ''', (title, role, person, amount, datetime.now().strftime("%Y-%m-%d")))
        conn.commit()
        conn.close()

        QMessageBox.information(self, "موفقیت", "هزینه ثبت شد.")
        self.load_expenses()

    def load_expenses(self):
        self.table_expense.setRowCount(0)
        conn = self.db.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id, title, role, person_name, amount FROM expenses")
        for row_idx, row_data in enumerate(cursor.fetchall()):
            self.table_expense.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                self.table_expense.setItem(row_idx, col_idx, QTableWidgetItem(str(value)))
        conn.close()

    def create_inventory_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)

        header = QLabel("انبار تجهیزات آی‌مارت استودیو")
        header.setFont(QFont("B Yekan", 14, QFont.Weight.Bold))
        layout.addWidget(header)

        form = QFormLayout()
        self.inv_name = QLineEdit()
        self.inv_cat = QComboBox()
        self.inv_cat.addItems(["دوربین", "لنز", "نورپردازی", "صدا", "پایه/استابلایزر", "سایر"])
        self.inv_qty = QSpinBox()

        form.addRow("نام تجهیزات:", self.inv_name)
        form.addRow("دسته‌بندی:", self.inv_cat)
        form.addRow("تعداد کل:", self.inv_qty)

        btn_add = QPushButton("افزودن به انبار")
        btn_add.setProperty("class", "ActionBtn")
        btn_add.clicked.connect(self.save_inventory)

        layout.addLayout(form)
        layout.addWidget(btn_add)

        self.table_inv = QTableWidget(0, 5)
        self.table_inv.setHorizontalHeaderLabels(["ID", "نام قطعه", "دسته", "تعداد کل", "در حال استفاده"])
        self.table_inv.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table_inv)

        self.load_inventory()
        return page

    def save_inventory(self):
        name = self.inv_name.text()
        cat = self.inv_cat.currentText()
        qty = self.inv_qty.value()

        conn = self.db.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO inventory (item_name, category, total_quantity, in_use_quantity)
            VALUES (?, ?, ?, 0)
        ''', (name, cat, qty))
        conn.commit()
        conn.close()

        QMessageBox.information(self, "موفقیت", "تجهیزات ثبت شد.")
        self.load_inventory()

    def load_inventory(self):
        self.table_inv.setRowCount(0)
        conn = self.db.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id, item_name, category, total_quantity, in_use_quantity FROM inventory")
        for row_idx, row_data in enumerate(cursor.fetchall()):
            self.table_inv.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                self.table_inv.setItem(row_idx, col_idx, QTableWidgetItem(str(value)))
        conn.close()

    def create_cards_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)

        header = QLabel("مدیریت کارت‌های بانکی جهت واریزی")
        header.setFont(QFont("B Yekan", 14, QFont.Weight.Bold))
        layout.addWidget(header)

        form = QFormLayout()
        self.card_bank = QLineEdit()
        self.card_num = QLineEdit()
        self.card_sheba = QLineEdit()
        self.card_owner = QLineEdit()

        form.addRow("نام بانک:", self.card_bank)
        form.addRow("شماره کارت:", self.card_num)
        form.addRow("شماره شبا:", self.card_sheba)
        form.addRow("نام صاحب حساب:", self.card_owner)

        btn_add = QPushButton("ثبت کارت")
        btn_add.setProperty("class", "ActionBtn")
        btn_add.clicked.connect(self.save_card)

        layout.addLayout(form)
        layout.addWidget(btn_add)

        self.table_cards = QTableWidget(0, 5)
        self.table_cards.setHorizontalHeaderLabels(["ID", "بانک", "شماره کارت", "شبا", "صاحب حساب"])
        self.table_cards.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table_cards)

        self.load_cards()
        return page

    def save_card(self):
        bank = self.card_bank.text()
        num = self.card_num.text()
        sheba = self.card_sheba.text()
        owner = self.card_owner.text()

        conn = self.db.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO bank_cards (bank_name, card_number, sheba, owner_name)
            VALUES (?, ?, ?, ?)
        ''', (bank, num, sheba, owner))
        conn.commit()
        conn.close()

        QMessageBox.information(self, "موفقیت", "کارت بانکی ثبت شد.")
        self.load_cards()

    def load_cards(self):
        self.table_cards.setRowCount(0)
        conn = self.db.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id, bank_name, card_number, sheba, owner_name FROM bank_cards")
        for row_idx, row_data in enumerate(cursor.fetchall()):
            self.table_cards.insertRow(row_idx)
            for col_idx, value in enumerate(row_data):
                self.table_cards.setItem(row_idx, col_idx, QTableWidgetItem(str(value)))
        conn.close()

    def create_settings_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)

        header = QLabel("تنظیمات و پشتیبان‌گیری")
        header.setFont(QFont("B Yekan", 14, QFont.Weight.Bold))
        layout.addWidget(header)

        # تغییر رمز عبور
        box_pwd = QGroupBox("تغییر رمز عبور ورود")
        pwd_layout = QFormLayout()
        self.txt_new_pwd = QLineEdit()
        self.txt_new_pwd.setEchoMode(QLineEdit.EchoMode.Password)
        btn_pwd = QPushButton("تغییر رمز")
        btn_pwd.setProperty("class", "ActionBtn")
        btn_pwd.clicked.connect(self.change_password)

        pwd_layout.addRow("رمز جدید:", self.txt_new_pwd)
        pwd_layout.addRow("", btn_pwd)
        box_pwd.setLayout(pwd_layout)
        layout.addWidget(box_pwd)

        # پشتیبان‌گیری
        box_backup = QGroupBox("پشتیبان‌گیری از دیتابیس")
        backup_layout = QHBoxLayout()
        btn_backup = QPushButton("💾 ایجاد نسخه پشتیبان (Backup)")
        btn_backup.setProperty("class", "ActionBtn")
        btn_backup.clicked.connect(self.make_backup)

        backup_layout.addWidget(btn_backup)
        box_backup.setLayout(backup_layout)
        layout.addWidget(box_backup)

        layout.addStretch()
        return page

    def change_password(self):
        new_pwd = self.txt_new_pwd.text().strip()
        if new_pwd:
            self.db.set_password(new_pwd)
            QMessageBox.information(self, "موفقیت", "رمز عبور با موفقیت تغییر یافت.")
            self.txt_new_pwd.clear()
        else:
            QMessageBox.warning(self, "خطا", "رمز عبور نمی‌تواند خالی باشد.")

    def make_backup(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "ذخیره فایل پشتیبان", "backup_imart.db", "Database Files (*.db)")
        if file_path:
            shutil.copy("imart_studio.db", file_path)
            QMessageBox.information(self, "موفقیت", "پشتیبان‌گیری با موفقیت انجام شد.")

    def closeEvent(self, event):
        """پشتیبان‌گیری خودکار هنگام بستن برنامه"""
        try:
            if not os.path.exists("auto_backups"):
                os.makedirs("auto_backups")
            filename = f"auto_backups/backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
            shutil.copy("imart_studio.db", filename)
        except Exception:
            pass
        event.accept()


# ==========================================
# 5. نقطه شروع اجرا (Main)
# ==========================================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
