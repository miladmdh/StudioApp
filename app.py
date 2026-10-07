import sys
import os
import shutil
import sqlite3
import jdatetime
from datetime import datetime

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QComboBox, QPushButton, QTableWidget, QTableWidgetItem,
    QHeaderView, QMessageBox, QGroupBox, QSpinBox, QFormLayout, QFileDialog,
    QTabWidget, QCheckBox, QInputDialog, QDialog, QStackedWidget, QFrame,
    QTextEdit, QScrollArea, QTreeWidget, QTreeWidgetItem, QDateEdit,
    QProgressBar, QGraphicsDropShadowEffect
)
from PyQt6.QtCore import Qt, QDate, QTimer
from PyQt6.QtGui import QFont, QIcon, QColor, QFontDatabase, QTextDocument, QPixmap, QPainter, QLinearGradient, QBrush
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog

import openpyxl
from openpyxl.styles import Font as XLFont, PatternFill, Alignment
import matplotlib.pyplot as plt

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# تنظیم ID جهت آیکون ویندوز
try:
    import ctypes
    myappid = 'imartstudio.accounting.v8.0'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass

DB_NAME = "studio_accounting.db"

# ⚠️ مقدار پیش‌فرض فونت (بعد از ساخت QApplication نهایی می‌شود)
APP_FONT_FAMILY = "Tahoma"


def setup_fonts():
    """تشخیص فونت مناسب فارسی - باید بعد از ساخت QApplication صدا زده شود"""
    font_family = "B Yekan"
    available_fonts = QFontDatabase.families()
    if "B Yekan" not in available_fonts and "B Nazanin" in available_fonts:
        font_family = "B Nazanin"
    elif "B Yekan" not in available_fonts and "B Nazanin" not in available_fonts:
        font_family = "Tahoma"
    return font_family


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def format_number(val):
    try:
        clean_val = str(val).replace(',', '').strip()
        if not clean_val:
            return ""
        return f"{int(clean_val):,}"
    except ValueError:
        return str(val)


def parse_number(val_str):
    try:
        return int(str(val_str).replace(',', '').strip())
    except ValueError:
        return 0


def ask_security_password(parent):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT value FROM settings WHERE key='app_password'")
    row = cursor.fetchone()
    saved_pass = row[0] if row else "123"
    conn.close()

    entered_pass, ok = QInputDialog.getText(
        parent, "تایید امنیتی", "لطفاً رمز عبور را وارد کنید:", QLineEdit.EchoMode.Password
    )
    if ok and entered_pass == saved_pass:
        return True
    elif ok:
        QMessageBox.critical(parent, "خطا", "رمز عبور نادرست است!")
        return False
    return False


def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    cursor.execute('''CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)''')
    cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('app_password', '123')")

    cursor.execute('''CREATE TABLE IF NOT EXISTS persons (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, role TEXT NOT NULL)''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS transactions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trans_type TEXT NOT NULL,
            category TEXT NOT NULL,
            person_id INTEGER,
            amount INTEGER NOT NULL,
            year INTEGER NOT NULL,
            month INTEGER NOT NULL,
            day INTEGER NOT NULL,
            description TEXT,
            FOREIGN KEY (person_id) REFERENCES persons(id)
        )
    ''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS expenses (id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT, category TEXT, amount INTEGER, date_str TEXT, description TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS bank_cards (id INTEGER PRIMARY KEY AUTOINCREMENT, bank_name TEXT NOT NULL, card_number TEXT, sheba_number TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS item_prices (item_name TEXT PRIMARY KEY, price INTEGER DEFAULT 0, is_package INTEGER DEFAULT 0, parent_package TEXT)''')

    default_items = [
        ("کلیپ فرمالیته", 15000000, 0, ""), ("کلیپ باغ", 8000000, 0, ""),
        ("یک دوربین", 5000000, 0, ""), ("دو دوربین", 9000000, 0, ""),
        ("کرین", 6000000, 0, ""), ("هلی شات", 7000000, 0, ""),
        ("عکاس مجلس", 4000000, 0, ""), ("پکیج طلایی VIP", 35000000, 1, "")
    ]
    for item, price, is_pkg, parent in default_items:
        cursor.execute("INSERT OR IGNORE INTO item_prices (item_name, price, is_package, parent_package) VALUES (?, ?, ?, ?)", (item, price, is_pkg, parent))

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS wedding_contracts (
            id INTEGER PRIMARY KEY AUTOINCREMENT, groom_name TEXT, bride_name TEXT, groom_phone TEXT, bride_phone TEXT,
            contract_date TEXT, ceremony_date TEXT, selected_items TEXT, total_amount INTEGER, discount INTEGER,
            paid_amount INTEGER, bank_id INTEGER, is_settled INTEGER DEFAULT 0, description TEXT
        )
    ''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS wedding_deposits (id INTEGER PRIMARY KEY AUTOINCREMENT, contract_id INTEGER, amount INTEGER, deposit_date TEXT, bank_name TEXT, FOREIGN KEY (contract_id) REFERENCES wedding_contracts(id))''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS commercial_projects (id INTEGER PRIMARY KEY AUTOINCREMENT, title TEXT, project_type TEXT, camera_count INTEGER, total_amount INTEGER, paid_amount INTEGER, project_date TEXT, description TEXT)''')
    cursor.execute('''CREATE TABLE IF NOT EXISTS inventory (item_name TEXT PRIMARY KEY, total_count INTEGER DEFAULT 0, used_count INTEGER DEFAULT 0)''')

    for item, _, _, _ in default_items:
        cursor.execute("INSERT OR IGNORE INTO inventory (item_name, total_count, used_count) VALUES (?, 10, 0)", (item,))

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS checks (
            id INTEGER PRIMARY KEY AUTOINCREMENT, check_number TEXT, bank_name TEXT, amount INTEGER,
            due_date TEXT, is_passed INTEGER DEFAULT 0, description TEXT
        )
    ''')

    conn.commit()
    conn.close()


# --- پنجره بارگذاری (Loading Screen) ---
class LoadingScreen(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("IMART STUDIO - در حال بارگذاری...")
        self.setFixedSize(550, 340)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        # کادر اصلی با گرادیانت
        self.container = QFrame(self)
        self.container.setGeometry(0, 0, 550, 340)
        self.container.setStyleSheet("""
            QFrame {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                    stop:0 #1F4E78, stop:0.5 #2c3e50, stop:1 #34495e);
                border-radius: 18px;
                border: 2px solid #2980b9;
            }
        """)

        # سایه زیر پنجره
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(30)
        shadow.setColor(QColor(0, 0, 0, 180))
        shadow.setOffset(0, 6)
        self.container.setGraphicsEffect(shadow)

        layout = QVBoxLayout(self.container)
        layout.setContentsMargins(35, 30, 35, 30)
        layout.setSpacing(15)

        # لوگو / نام استودیو
        self.lbl_logo = QLabel("🎬 IMART STUDIO")
        self.lbl_logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_logo.setStyleSheet("color: white; font-size: 26pt; font-weight: bold; background: transparent;")
        layout.addWidget(self.lbl_logo)

        # زیرعنوان
        self.lbl_sub = QLabel("سیستم جامع مدیریت مالی و حسابداری - نسخه ۸.۰")
        self.lbl_sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_sub.setStyleSheet("color: #aed6f1; font-size: 11pt; background: transparent;")
        layout.addWidget(self.lbl_sub)

        layout.addSpacing(10)

        # درصد
        self.lbl_percent = QLabel("0%")
        self.lbl_percent.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_percent.setStyleSheet("color: #f1c40f; font-size: 32pt; font-weight: bold; background: transparent;")
        layout.addWidget(self.lbl_percent)

        # نوار پیشرفت
        self.progress = QProgressBar()
        self.progress.setFixedHeight(20)
        self.progress.setTextVisible(False)
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setStyleSheet("""
            QProgressBar {
                background-color: rgba(255, 255, 255, 30);
                border-radius: 10px;
                border: 1px solid rgba(255, 255, 255, 60);
            }
            QProgressBar::chunk {
                background-color: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #27ae60, stop:0.5 #2ecc71, stop:1 #1abc9c);
                border-radius: 10px;
            }
        """)
        layout.addWidget(self.progress)

        # متن مرحله فعلی
        self.lbl_status = QLabel("در حال آماده‌سازی...")
        self.lbl_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_status.setStyleSheet("color: white; font-size: 11pt; background: transparent;")
        layout.addWidget(self.lbl_status)

        layout.addStretch()

        # فوتر
        self.lbl_footer = QLabel("📞 09173736618")
        self.lbl_footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_footer.setStyleSheet("color: #bdc3c7; font-size: 9pt; background: transparent;")
        layout.addWidget(self.lbl_footer)

        # Layout خارجی برای جاگذاری container
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(self.container)

        self.center_on_screen()

    def center_on_screen(self):
        screen = QApplication.primaryScreen().geometry()
        x = (screen.width() - self.width()) // 2
        y = (screen.height() - self.height()) // 2
        self.move(x, y)

    def set_progress(self, value, status_text=""):
        self.progress.setValue(value)
        self.lbl_percent.setText(f"{value}%")
        if status_text:
            self.lbl_status.setText(status_text)
        QApplication.processEvents()


# --- پنجره مدیریت فاکتور و پکیج‌ها ---
class ManageItemsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("مدیریت و افزودن موارد و پکیج‌های فاکتور")
        self.resize(550, 500)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.setFont(QFont(APP_FONT_FAMILY, 10))

        layout = QVBoxLayout()
        form_layout = QFormLayout()

        self.txt_item_name = QLineEdit()
        self.txt_item_price = QLineEdit()
        self.txt_item_price.textChanged.connect(lambda t: self.txt_item_price.setText(format_number(t)))

        self.chk_is_package = QCheckBox("این مورد یک پکیج اصلی است")
        self.chk_is_package.stateChanged.connect(self.toggle_package_mode)

        self.combo_parent_package = QComboBox()
        self.load_packages_combo()

        form_layout.addRow("نام مورد / پکیج:", self.txt_item_name)
        form_layout.addRow("قیمت (تومان):", self.txt_item_price)
        form_layout.addRow("", self.chk_is_package)
        form_layout.addRow("انتخاب پکیج مادر (اختیاری):", self.combo_parent_package)

        btn_add = QPushButton("افزودن / بروزرسانی قیمت پیش‌فرض")
        btn_add.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; padding: 6px;")
        btn_add.clicked.connect(self.add_or_update_item)
        form_layout.addRow(btn_add)
        layout.addLayout(form_layout)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["نام آیتم / زیرمجموعه", "قیمت پیش‌فرض (تومان)"])
        layout.addWidget(self.tree)

        btn_delete = QPushButton("حذف مورد انتخاب‌شده (با رمز عبور)")
        btn_delete.setStyleSheet("background-color: #c0392b; color: white; font-weight: bold; padding: 6px;")
        btn_delete.clicked.connect(self.delete_item)
        layout.addWidget(btn_delete)

        self.load_tree_items()
        self.setLayout(layout)

    def toggle_package_mode(self, state):
        self.combo_parent_package.setEnabled(not self.chk_is_package.isChecked())

    def load_packages_combo(self):
        self.combo_parent_package.clear()
        self.combo_parent_package.addItem("--- بدون پکیج مادر (آیتم مستقل) ---", "")
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT item_name FROM item_prices WHERE is_package=1")
        for (pkg_name,) in cursor.fetchall():
            self.combo_parent_package.addItem(pkg_name, pkg_name)
        conn.close()

    def load_tree_items(self):
        self.tree.clear()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT item_name, price, is_package, parent_package FROM item_prices")
        rows = cursor.fetchall()
        conn.close()

        packages = [r for r in rows if r[2] == 1]
        singles = [r for r in rows if r[2] == 0]

        for pkg in packages:
            pkg_item = QTreeWidgetItem(self.tree, [pkg[0], f"{pkg[1]:,}"])
            pkg_item.setFont(0, QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
            sub_items = [s for s in singles if s[3] == pkg[0]]
            for sub in sub_items:
                QTreeWidgetItem(pkg_item, [f"  └ {sub[0]}", f"{sub[1]:,}"])

        independent = [s for s in singles if not s[3]]
        for ind in independent:
            QTreeWidgetItem(self.tree, [ind[0], f"{ind[1]:,}"])

        self.tree.expandAll()

    def add_or_update_item(self):
        if not ask_security_password(self):
            return

        name = self.txt_item_name.text().strip()
        price = parse_number(self.txt_item_price.text())
        is_pkg = 1 if self.chk_is_package.isChecked() else 0
        parent = self.combo_parent_package.currentData() if not is_pkg else ""

        if not name:
            QMessageBox.warning(self, "خطا", "لطفاً نام را وارد کنید.")
            return

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO item_prices (item_name, price, is_package, parent_package)
            VALUES (?, ?, ?, ?) ON CONFLICT(item_name) DO UPDATE SET price=excluded.price, is_package=excluded.is_package, parent_package=excluded.parent_package
        ''', (name, price, is_pkg, parent))
        cursor.execute("INSERT OR IGNORE INTO inventory (item_name, total_count, used_count) VALUES (?, 10, 0)", (name,))
        conn.commit()
        conn.close()

        self.txt_item_name.clear()
        self.txt_item_price.clear()
        self.load_packages_combo()
        self.load_tree_items()

    def delete_item(self):
        selected = self.tree.currentItem()
        if not selected:
            QMessageBox.warning(self, "خطا", "لطفاً یک آیتم را انتخاب کنید.")
            return

        item_name = selected.text(0).replace("  └ ", "").strip()
        if not ask_security_password(self):
            return

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM item_prices WHERE item_name=?", (item_name,))
        cursor.execute("DELETE FROM inventory WHERE item_name=?", (item_name,))
        conn.commit()
        conn.close()

        self.load_packages_combo()
        self.load_tree_items()


# --- پنجره مدیریت بیعانه‌ها ---
class DepositsDialog(QDialog):
    def __init__(self, contract_id, parent=None):
        super().__init__(parent)
        self.contract_id = contract_id
        self.setWindowTitle(f"مدیریت بیعانه‌ها و پرداخت‌های قرارداد #{contract_id}")
        self.resize(500, 350)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.setFont(QFont(APP_FONT_FAMILY, 10))

        layout = QVBoxLayout()
        form_layout = QFormLayout()

        self.txt_amount = QLineEdit()
        self.txt_amount.textChanged.connect(lambda t: self.txt_amount.setText(format_number(t)))
        self.txt_date = QLineEdit(jdatetime.date.today().strftime("%Y/%m/%d"))
        self.combo_bank = QComboBox()
        self.load_banks()

        form_layout.addRow("مبلغ پرداخت (تومان):", self.txt_amount)
        form_layout.addRow("تاریخ دریافت:", self.txt_date)
        form_layout.addRow("بانک واریزی:", self.combo_bank)

        btn_add = QPushButton("ثبت بیعانه / پرداختی جدید")
        btn_add.setStyleSheet("background-color: #2980b9; color: white; font-weight: bold; padding: 5px;")
        btn_add.clicked.connect(self.add_deposit)
        form_layout.addRow(btn_add)
        layout.addLayout(form_layout)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["ردیف", "مبلغ (تومان)", "تاریخ", "بانک"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table)

        self.load_deposits()
        self.setLayout(layout)

    def load_banks(self):
        self.combo_bank.clear()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT bank_name FROM bank_cards")
        for (b_name,) in cursor.fetchall():
            self.combo_bank.addItem(b_name)
        conn.close()

    def load_deposits(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, amount, deposit_date, bank_name FROM wedding_deposits WHERE contract_id=?", (self.contract_id,))
        rows = cursor.fetchall()
        conn.close()

        self.table.setRowCount(0)
        for r_idx, (d_id, amount, d_date, bank) in enumerate(rows):
            self.table.insertRow(r_idx)
            self.table.setItem(r_idx, 0, QTableWidgetItem(str(d_id)))
            self.table.setItem(r_idx, 1, QTableWidgetItem(f"{amount:,}"))
            self.table.setItem(r_idx, 2, QTableWidgetItem(d_date))
            self.table.setItem(r_idx, 3, QTableWidgetItem(bank if bank else "-"))

    def add_deposit(self):
        amount = parse_number(self.txt_amount.text())
        if amount <= 0:
            QMessageBox.warning(self, "خطا", "لطفاً مبلغ معتبر وارد کنید.")
            return

        d_date = self.txt_date.text().strip()
        bank = self.combo_bank.currentText()

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO wedding_deposits (contract_id, amount, deposit_date, bank_name) VALUES (?, ?, ?, ?)", (self.contract_id, amount, d_date, bank))
        cursor.execute("SELECT SUM(amount) FROM wedding_deposits WHERE contract_id=?", (self.contract_id,))
        total_paid = cursor.fetchone()[0] or 0
        cursor.execute("SELECT total_amount, discount FROM wedding_contracts WHERE id=?", (self.contract_id,))
        result = cursor.fetchone()
        if result:
            tot, disc = result
            is_settled = 1 if (tot - disc - total_paid) <= 0 else 0
            cursor.execute("UPDATE wedding_contracts SET paid_amount=?, is_settled=? WHERE id=?", (total_paid, is_settled, self.contract_id))

        conn.commit()
        conn.close()

        self.txt_amount.clear()
        self.load_deposits()
        QMessageBox.information(self, "موفقیت", "بیعانه با موفقیت ثبت شد.")


# --- پنجره جزئیات قرارداد ---
class ContractDetailsDialog(QDialog):
    def __init__(self, contract_id, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"جزئیات کامل قرارداد ردیف #{contract_id}")
        self.resize(600, 450)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.setFont(QFont(APP_FONT_FAMILY, 10))

        layout = QVBoxLayout()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT groom_name, bride_name, groom_phone, bride_phone, contract_date, ceremony_date, selected_items, total_amount, discount, paid_amount, description FROM wedding_contracts WHERE id=?", (contract_id,))
        row = cursor.fetchone()
        conn.close()

        if row:
            info_text = f"<b>زوجین:</b> {row[0]} و {row[1]} | <b>تلفن داماد:</b> {row[2]} | <b>تلفن عروس:</b> {row[3]}<br>"
            info_text += f"<b>تاریخ قرارداد:</b> {row[4]} | <b>تاریخ مراسم:</b> {row[5]}<br>"
            info_text += f"<b>توضیحات:</b> {row[10] if row[10] else 'ندارد'}"
            lbl_info = QLabel(info_text)
            lbl_info.setStyleSheet("background-color: #f8f9fa; padding: 10px; border-radius: 5px;")
            layout.addWidget(lbl_info)

            lbl_items = QLabel("<b>مفاد و خدمات انتخاب شده:</b>")
            layout.addWidget(lbl_items)

            table = QTableWidget()
            table.setColumnCount(1)
            table.setHorizontalHeaderLabels(["عنوان خدمت / پکیج"])
            table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

            items = row[6].split(',') if row[6] else []
            table.setRowCount(len(items))
            for idx, item in enumerate(items):
                table.setItem(idx, 0, QTableWidgetItem(item))
            layout.addWidget(table)

            remain = row[7] - row[9]
            summary = f"<b>جمع کل:</b> {row[7]:,} تومان | <b>تخفیف:</b> {row[8]:,} تومان | <b>پرداختی:</b> {row[9]:,} تومان | <b>مانده:</b> {remain:,} تومان"
            lbl_sum = QLabel(summary)
            lbl_sum.setStyleSheet("font-size: 11pt; color: #2c3e50; font-weight: bold; margin-top: 5px;")
            layout.addWidget(lbl_sum)

        self.setLayout(layout)


# --- کلاس اصلی برنامه ---
class StudioAccountingApp(QMainWindow):
    ROLES = ["تدوینگر", "عکاس", "فیلمبردار", "هلی شات و FPV کار", "اوپراتور کرین"]
    PROJECT_TYPES = ["عروسی", "عقد", "تولد", "تبلیغاتی", "قبض و کرایه"]
    PERSIAN_MONTHS = ["فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور", "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند"]

    def __init__(self):
        super().__init__()
        self.setWindowTitle("نرم‌افزار مدیریت مالی و حسابداری IMART STUDIO - نسخه ۸.۰")
        self.resize(1300, 900)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)

        icon_path = resource_path("Accounting.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.setFont(QFont(APP_FONT_FAMILY, 10))

        # ساخت استک و صفحات
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        self.dashboard_screen = QWidget()
        self.stack.addWidget(self.dashboard_screen)

        self.main_app_screen = QWidget()
        self.stack.addWidget(self.main_app_screen)

        # اطمینان از اینکه داشبورد به عنوان صفحه پیش‌فرض نمایش داده شود
        self.stack.setCurrentWidget(self.dashboard_screen)

    def prompt_login(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM settings WHERE key='app_password'")
        row = cursor.fetchone()
        saved_pass = row[0] if row else "123"
        conn.close()

        entered_pass, ok = QInputDialog.getText(
            None, "ورود به سیستم IMART STUDIO", "لطفاً رمز عبور را وارد کنید:", QLineEdit.EchoMode.Password
        )
        if ok and entered_pass == saved_pass:
            return True
        elif ok:
            QMessageBox.critical(None, "خطا", "رمز عبور اشتباه است!")
            return False
        return False

    def setup_dashboard_ui(self):
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title = QLabel("سیستم جامع مدیریت مالی و حسابداری IMART STUDIO (نسخه ۸.۰)")
        title.setFont(QFont(APP_FONT_FAMILY, 16, QFont.Weight.Bold))
        title.setStyleSheet("color: #2c3e50; margin-bottom: 25px;")
        layout.addWidget(title, alignment=Qt.AlignmentFlag.AlignCenter)

        grid_layout = QHBoxLayout()
        buttons = [
            ("پروژه‌های عروس و داماد\n🔒 (قراردادها)", 0, "#2980b9"),
            ("تبلیغاتی / بیوتی / تولدی\n🔒 (پروژه‌ها)", 1, "#27ae60"),
            ("بخش کارکنان و حقوق\n🔒 (پرسنل)", 2, "#8e44ad"),
            ("مدیریت هزینه‌ها\n🔒 (قبوض و...) ", 3, "#c0392b"),
            ("گزارش مالی جامع\n(نمودارها)", 4, "#d35400"),
            ("جستجوی پیشرفته\n(کلی)", 5, "#16a085"),
            ("انبار تجهیزات\n(تجهیزات)", 6, "#2c3e50"),
            ("کارت‌های بانکی\n(بانک‌ها)", 7, "#1abc9c"),
            ("مدیریت چک‌ها\n(چک‌ها)", 8, "#e67e22"),
            ("رسید وجه\n(خدمات)", 9, "#7f8c8d"),
            ("سال کاری\n(بستن حساب)", 10, "#95a5a6"),
            ("درباره برنامه\n(توسعه‌دهنده)", 11, "#34495e")
        ]

        for text, idx, color in buttons:
            btn = QPushButton(text)
            btn.setFixedSize(130, 130)
            btn.setFont(QFont(APP_FONT_FAMILY, 9, QFont.Weight.Bold))
            btn.setStyleSheet(f"QPushButton {{ background-color: {color}; color: white; border-radius: 10px; padding: 5px; }} QPushButton:hover {{ background-color: #34495e; }}")
            btn.clicked.connect(lambda _, i=idx: self.open_tab_index(i))
            grid_layout.addWidget(btn)

        layout.addLayout(grid_layout)
        self.dashboard_screen.setLayout(layout)

    def open_tab_index(self, index):
        if hasattr(self, 'tabs'):
            self.tabs.setCurrentIndex(index)
            self.stack.setCurrentWidget(self.main_app_screen)

    def setup_main_app_ui(self):
        main_layout = QVBoxLayout()

        top_bar = QHBoxLayout()
        btn_dash = QPushButton("بازگشت به منوی اصلی")
        btn_dash.setStyleSheet("background-color: #34495e; color: white; font-weight: bold; padding: 6px;")
        btn_dash.clicked.connect(lambda: self.stack.setCurrentWidget(self.dashboard_screen))

        btn_backup = QPushButton("پشتیبان‌گیری (Backup)")
        btn_backup.clicked.connect(self.backup_db)
        btn_restore = QPushButton("بازیابی بک‌آپ (Restore)")
        btn_restore.clicked.connect(self.restore_db)
        btn_change_pass = QPushButton("تغییر رمز عبور")
        btn_change_pass.clicked.connect(self.change_password)

        top_bar.addWidget(btn_dash)
        top_bar.addWidget(btn_backup)
        top_bar.addWidget(btn_restore)
        top_bar.addWidget(btn_change_pass)
        top_bar.addStretch()
        main_layout.addLayout(top_bar)

        self.tabs = QTabWidget()
        self.tabs.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))

        self.tab_wedding = QWidget()
        self.tab_commercial = QWidget()
        self.tab_staff = QWidget()
        self.tab_expenses = QWidget()
        self.tab_reports = QWidget()
        self.tab_search = QWidget()
        self.tab_inventory = QWidget()
        self.tab_banks = QWidget()
        self.tab_checks = QWidget()
        self.tab_receipt = QWidget()
        self.tab_workyear = QWidget()
        self.tab_about = QWidget()

        self.tabs.addTab(self.tab_wedding, "پروژه‌های عروس و داماد")
        self.tabs.addTab(self.tab_commercial, "تبلیغاتی / بیوتی")
        self.tabs.addTab(self.tab_staff, "کارکنان و پرسنل")
        self.tabs.addTab(self.tab_expenses, "مدیریت هزینه‌ها")
        self.tabs.addTab(self.tab_reports, "گزارش مالی جامع")
        self.tabs.addTab(self.tab_search, "جستجوی پیشرفته")
        self.tabs.addTab(self.tab_inventory, "انبار تجهیزات")
        self.tabs.addTab(self.tab_banks, "کارت‌های بانکی")
        self.tabs.addTab(self.tab_checks, "مدیریت چک‌ها")
        self.tabs.addTab(self.tab_receipt, "صدور رسید وجه")
        self.tabs.addTab(self.tab_workyear, "سال کاری")
        self.tabs.addTab(self.tab_about, "درباره برنامه")

        main_layout.addWidget(self.tabs)

        self.setup_wedding_tab()
        self.setup_commercial_tab()
        self.setup_staff_tab()
        self.setup_expenses_tab()
        self.setup_reports_tab()
        self.setup_search_tab()
        self.setup_inventory_tab()
        self.setup_banks_tab()
        self.setup_checks_tab()
        self.setup_receipt_tab()
        self.setup_workyear_tab()
        self.setup_about_tab()

        footer = QLabel("IMART STUDIO - Phone: 09173736618 | نسخه ۸.۰")
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        footer.setStyleSheet("color: #2c3e50; margin-top: 5px;")
        main_layout.addWidget(footer)

        self.main_app_screen.setLayout(main_layout)

    # --- زبانه ۱: عروس و داماد ---
    def setup_wedding_tab(self):
        layout = QHBoxLayout()
        form_box = QGroupBox("ثبت قرارداد جدید عروس و داماد")
        form_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        form_layout = QFormLayout()

        self.w_groom = QLineEdit()
        self.w_bride = QLineEdit()
        self.w_groom_phone = QLineEdit()
        self.w_bride_phone = QLineEdit()
        self.w_contract_date = QLineEdit(jdatetime.date.today().strftime("%Y/%m/%d"))
        self.w_ceremony_date = QLineEdit(jdatetime.date.today().strftime("%Y/%m/%d"))

        form_layout.addRow("نام داماد *:", self.w_groom)
        form_layout.addRow("نام عروس *:", self.w_bride)
        form_layout.addRow("تلفن داماد:", self.w_groom_phone)
        form_layout.addRow("تلفن عروس:", self.w_bride_phone)
        form_layout.addRow("تاریخ قرارداد:", self.w_contract_date)
        form_layout.addRow("تاریخ مراسم:", self.w_ceremony_date)

        btn_manage_items = QPushButton("مدیریت / افزودن پکیج و موارد")
        btn_manage_items.setStyleSheet("background-color: #e67e22; color: white; font-weight: bold;")
        btn_manage_items.clicked.connect(self.open_manage_items)
        form_layout.addRow(btn_manage_items)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFixedHeight(180)

        self.items_container = QWidget()
        self.items_vbox = QVBoxLayout()
        self.items_container.setLayout(self.items_vbox)
        scroll_area.setWidget(self.items_container)

        items_box = QGroupBox("مفاد فاکتور (انتخاب موارد)")
        items_box_layout = QVBoxLayout()
        items_box_layout.addWidget(scroll_area)
        items_box.setLayout(items_box_layout)
        form_layout.addRow(items_box)

        self.w_lbl_total = QLabel("جمع کل: ۰ تومان")
        self.w_lbl_total.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        form_layout.addRow(self.w_lbl_total)

        self.w_discount = QLineEdit("0")
        self.w_discount.textChanged.connect(lambda t: self.on_discount_changed(t))
        self.w_first_deposit = QLineEdit("0")
        self.w_first_deposit.textChanged.connect(lambda t: self.on_deposit_changed(t))

        form_layout.addRow("تخفیف (تومان):", self.w_discount)
        form_layout.addRow("بیعانه اول (تومان):", self.w_first_deposit)

        self.w_bank_combo = QComboBox()
        self.load_bank_combo()
        form_layout.addRow("بانک واریزی بیعانه:", self.w_bank_combo)

        self.w_lbl_remain = QLabel("باقی‌مانده حساب: ۰ تومان")
        self.w_lbl_remain.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        self.w_lbl_remain.setStyleSheet("color: #c0392b;")
        form_layout.addRow(self.w_lbl_remain)

        self.w_desc = QTextEdit()
        self.w_desc.setFixedHeight(60)
        self.w_desc.setPlaceholderText("توضیحات کامل قرارداد...")
        form_layout.addRow("توضیحات قرارداد:", self.w_desc)

        self.item_checkboxes = {}
        self.item_price_inputs = {}
        self.load_item_checkboxes()

        self.btn_save_wedding = QPushButton("ثبت نهایی قرارداد")
        self.btn_save_wedding.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        self.btn_save_wedding.setStyleSheet("background-color: #2980b9; color: white; padding: 8px;")
        self.btn_save_wedding.clicked.connect(self.save_wedding_contract)
        form_layout.addRow(self.btn_save_wedding)

        form_box.setLayout(form_layout)
        layout.addWidget(form_box, 1)

        table_box = QGroupBox("لیست قراردادها (جهت مشاهده جزئیات دوبار کلیک کنید)")
        table_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        table_layout = QVBoxLayout()

        self.w_table = QTableWidget()
        self.w_table.setColumnCount(13)
        self.w_table.setHorizontalHeaderLabels([
            "ردیف", "زوجین", "تلفن داماد", "تلفن عروس", "تاریخ مراسم", "جمع کل", "تخفیف", "دریافتی", "مانده", "وضعیت", "بیعانه‌ها", "PDF / چاپ", "حذف"
        ])
        self.w_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.w_table.doubleClicked.connect(self.on_wedding_double_click)
        table_layout.addWidget(self.w_table)
        table_box.setLayout(table_layout)
        layout.addWidget(table_box, 2)

        self.tab_wedding.setLayout(layout)
        self.load_wedding_contracts()

    def open_manage_items(self):
        dlg = ManageItemsDialog(self)
        dlg.exec()
        self.load_item_checkboxes()

    def load_item_checkboxes(self):
        for i in reversed(range(self.items_vbox.count())):
            widget = self.items_vbox.itemAt(i).widget()
            if widget:
                widget.setParent(None)

        self.item_checkboxes = {}
        self.item_price_inputs = {}

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT item_name, price FROM item_prices")
        rows = cursor.fetchall()
        conn.close()

        for item_name, price in rows:
            row_w = QWidget()
            row_l = QHBoxLayout()
            row_l.setContentsMargins(0, 0, 0, 0)

            cb = QCheckBox(item_name)
            cb.stateChanged.connect(self.calc_wedding_total)

            txt_p = QLineEdit(f"{price:,}")
            txt_p.setFixedWidth(90)
            txt_p.textChanged.connect(lambda t, p=txt_p: p.setText(format_number(t)))
            txt_p.textChanged.connect(self.calc_wedding_total)

            row_l.addWidget(cb)
            row_l.addWidget(QLabel("قیمت:"))
            row_l.addWidget(txt_p)
            row_w.setLayout(row_l)

            self.items_vbox.addWidget(row_w)
            self.item_checkboxes[item_name] = cb
            self.item_price_inputs[item_name] = txt_p

        self.calc_wedding_total()

    def on_discount_changed(self, text):
        formatted = format_number(text)
        if formatted != text:
            self.w_discount.setText(formatted)
        self.calc_wedding_total()

    def on_deposit_changed(self, text):
        formatted = format_number(text)
        if formatted != text:
            self.w_first_deposit.setText(formatted)
        self.calc_wedding_total()

    def calc_wedding_total(self):
        if not hasattr(self, 'w_discount') or not hasattr(self, 'w_first_deposit'):
            return

        selected_sum = 0
        for item_name, cb in self.item_checkboxes.items():
            if cb.isChecked():
                selected_sum += parse_number(self.item_price_inputs[item_name].text())

        discount = parse_number(self.w_discount.text())
        deposit = parse_number(self.w_first_deposit.text())

        final_total = max(0, selected_sum - discount)
        remain = max(0, final_total - deposit)

        self.w_lbl_total.setText(f"جمع کل: {final_total:,} تومان")
        self.w_lbl_remain.setText(f"باقی‌مانده حساب: {remain:,} تومان")

    def load_bank_combo(self):
        self.w_bank_combo.clear()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, bank_name FROM bank_cards")
        for b_id, b_name in cursor.fetchall():
            self.w_bank_combo.addItem(b_name, b_id)
        conn.close()

    def save_wedding_contract(self):
        groom = self.w_groom.text().strip()
        bride = self.w_bride.text().strip()

        if not groom or not bride:
            QMessageBox.warning(self, "خطا", "لطفاً نام عروس و داماد را وارد کنید.")
            return

        selected_items = [item for item, cb in self.item_checkboxes.items() if cb.isChecked()]
        if not selected_items:
            QMessageBox.warning(self, "خطا", "تا زمانی که حداقل یک مورد از فاکتور انتخاب نشود، امکان ثبت وجود ندارد.")
            return

        selected_sum = sum(parse_number(self.item_price_inputs[i].text()) for i in selected_items)
        discount = parse_number(self.w_discount.text())
        total_amount = max(0, selected_sum - discount)
        first_deposit = parse_number(self.w_first_deposit.text())
        bank_id = self.w_bank_combo.currentData()
        bank_name = self.w_bank_combo.currentText()
        desc = self.w_desc.toPlainText().strip()

        remain = total_amount - first_deposit
        is_settled = 1 if remain <= 0 else 0

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO wedding_contracts 
            (groom_name, bride_name, groom_phone, bride_phone, contract_date, ceremony_date, selected_items, total_amount, discount, paid_amount, bank_id, is_settled, description)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (groom, bride, self.w_groom_phone.text(), self.w_bride_phone.text(),
              self.w_contract_date.text(), self.w_ceremony_date.text(),
              ",".join(selected_items), total_amount, discount, first_deposit, bank_id, is_settled, desc))

        contract_id = cursor.lastrowid
        if first_deposit > 0:
            cursor.execute('''
                INSERT INTO wedding_deposits (contract_id, amount, deposit_date, bank_name)
                VALUES (?, ?, ?, ?)
            ''', (contract_id, first_deposit, self.w_contract_date.text(), bank_name))

        for item in selected_items:
            cursor.execute("UPDATE inventory SET used_count = used_count + 1 WHERE item_name = ?", (item,))

        conn.commit()
        conn.close()

        QMessageBox.information(self, "موفقیت", "قرارداد با موفقیت ثبت شد.")
        self.load_wedding_contracts()
        self.load_inventory()

    def load_wedding_contracts(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT id, groom_name, bride_name, groom_phone, bride_phone, ceremony_date, total_amount, discount, paid_amount, is_settled
            FROM wedding_contracts ORDER BY id DESC
        ''')
        rows = cursor.fetchall()
        conn.close()

        self.w_table.setRowCount(0)
        for r_idx, row in enumerate(rows):
            c_id, groom, bride, g_phone, b_phone, cer_date, total, discount, paid, settled = row
            remain = total - paid

            self.w_table.insertRow(r_idx)
            self.w_table.setItem(r_idx, 0, QTableWidgetItem(str(c_id)))
            self.w_table.setItem(r_idx, 1, QTableWidgetItem(f"{groom} و {bride}"))
            self.w_table.setItem(r_idx, 2, QTableWidgetItem(g_phone if g_phone else "-"))
            self.w_table.setItem(r_idx, 3, QTableWidgetItem(b_phone if b_phone else "-"))
            self.w_table.setItem(r_idx, 4, QTableWidgetItem(cer_date))
            self.w_table.setItem(r_idx, 5, QTableWidgetItem(f"{total:,}"))
            self.w_table.setItem(r_idx, 6, QTableWidgetItem(f"{discount:,}"))
            self.w_table.setItem(r_idx, 7, QTableWidgetItem(f"{paid:,}"))
            self.w_table.setItem(r_idx, 8, QTableWidgetItem(f"{remain:,}"))

            item_settled = QTableWidgetItem("✅ تسویه کامل" if settled else "⏳ در جریان")
            if settled:
                item_settled.setForeground(QColor("green"))
            self.w_table.setItem(r_idx, 9, item_settled)

            btn_dep = QPushButton("بیعانه‌ها")
            btn_dep.clicked.connect(lambda _, cid=c_id: self.open_deposits_dialog(cid))
            self.w_table.setCellWidget(r_idx, 10, btn_dep)

            print_widget = QWidget()
            print_layout = QHBoxLayout()
            print_layout.setContentsMargins(0, 0, 0, 0)
            btn_pdf = QPushButton("PDF")
            btn_pdf.clicked.connect(lambda _, cid=c_id: self.export_wedding_pdf(cid))
            btn_direct_print = QPushButton("🖨️ پرینت")
            btn_direct_print.clicked.connect(lambda _, cid=c_id: self.direct_print_wedding(cid))
            print_layout.addWidget(btn_pdf)
            print_layout.addWidget(btn_direct_print)
            print_widget.setLayout(print_layout)
            self.w_table.setCellWidget(r_idx, 11, print_widget)

            btn_del = QPushButton("حذف")
            btn_del.setStyleSheet("background-color: #e74c3c; color: white;")
            btn_del.clicked.connect(lambda _, cid=c_id: self.delete_wedding_contract(cid))
            self.w_table.setCellWidget(r_idx, 12, btn_del)

        self.w_table.resizeRowsToContents()

    def delete_wedding_contract(self, contract_id):
        if not ask_security_password(self):
            return
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM wedding_contracts WHERE id=?", (contract_id,))
        cursor.execute("DELETE FROM wedding_deposits WHERE contract_id=?", (contract_id,))
        conn.commit()
        conn.close()
        self.load_wedding_contracts()

    def on_wedding_double_click(self, index):
        row = index.row()
        contract_id = int(self.w_table.item(row, 0).text())
        dlg = ContractDetailsDialog(contract_id, self)
        dlg.exec()

    def open_deposits_dialog(self, contract_id):
        dlg = DepositsDialog(contract_id, self)
        dlg.exec()
        self.load_wedding_contracts()

    def generate_pdf_html(self, contract_id):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT groom_name, bride_name, groom_phone, bride_phone, contract_date, ceremony_date, selected_items, total_amount, discount, paid_amount, description FROM wedding_contracts WHERE id=?", (contract_id,))
        c = cursor.fetchone()
        conn.close()

        if not c:
            return "", ""

        filename_default = f"فاکتور_{c[0]}_{c[1]}_{c[4].replace('/', '-')}.pdf"
        items_list = c[6].split(',') if c[6] else []
        items_rows = "".join([f"<tr><td style='border:1px solid #ccc; padding:8px; text-align:center;'>{idx+1}</td><td style='border:1px solid #ccc; padding:8px;'>{item}</td></tr>" for idx, item in enumerate(items_list)])

        remain = c[7] - c[9]
        html_content = f"""
        <div dir="rtl" style="font-family: '{APP_FONT_FAMILY}', 'Tahoma'; padding: 20px;">
            <h2 style="text-align: center; color: #1F4E78; margin-bottom: 5px;">IMART STUDIO</h2>
            <h3 style="text-align: center; margin-top: 0;">فاکتور رسمی قرارداد خدمات فیلمبرداری و عکاسی</h3>
            <hr>
            <table style="width: 100%; margin-bottom: 15px; border-collapse: collapse;">
                <tr>
                    <td style="padding: 5px;"><b>نام داماد و عروس:</b> {c[0]} و {c[1]}</td>
                    <td style="padding: 5px; text-align: left;"><b>تاریخ قرارداد:</b> {c[4]}</td>
                </tr>
                <tr>
                    <td style="padding: 5px;"><b>تلفن تماس:</b> {c[2]} / {c[3]}</td>
                    <td style="padding: 5px; text-align: left;"><b>تاریخ مراسم:</b> {c[5]}</td>
                </tr>
            </table>

            <h4 style="margin-bottom: 5px;">مفاد و خدمات تعهد شده:</h4>
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 15px;">
                <thead>
                    <tr style="background-color: #f2f2f2;">
                        <th style="border:1px solid #ccc; padding:8px; width: 10%;">ردیف</th>
                        <th style="border:1px solid #ccc; padding:8px;">عنوان خدمت / پکیج</th>
                    </tr>
                </thead>
                <tbody>
                    {items_rows}
                </tbody>
            </table>

            <table style="width: 100%; border-collapse: collapse; margin-top: 15px;">
                <tr style="background-color: #f9f9f9;">
                    <td style="border:1px solid #ccc; padding:8px;"><b>مبلغ کل قرارداد:</b> {c[7]:,} تومان</td>
                    <td style="border:1px solid #ccc; padding:8px;"><b>تخفیف:</b> {c[8]:,} تومان</td>
                </tr>
                <tr style="background-color: #f9f9f9;">
                    <td style="border:1px solid #ccc; padding:8px;"><b>مجموع پرداختی (بیعانه‌ها):</b> {c[9]:,} تومان</td>
                    <td style="border:1px solid #ccc; padding:8px; color: red;"><b>باقی‌مانده حساب:</b> {remain:,} تومان</td>
                </tr>
            </table>

            <p style="margin-top: 15px;"><b>توضیحات:</b> {c[10] if c[10] else 'موردی ثبت نشده است.'}</p>
            <br><br>
            <table style="width: 100%; text-align: center; margin-top: 30px;">
                <tr>
                    <td>امضاء مشتری (عروس و داماد)</td>
                    <td>امضاء و مهر IMART STUDIO</td>
                </tr>
            </table>
        </div>
        """
        return html_content, filename_default

    def export_wedding_pdf(self, contract_id):
        html_content, filename_default = self.generate_pdf_html(contract_id)
        if not html_content:
            return

        file_path, _ = QFileDialog.getSaveFileName(self, "ذخیره فاکتور PDF", filename_default, "PDF Files (*.pdf)")
        if not file_path:
            return

        doc = QTextDocument()
        doc.setHtml(html_content)
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
        printer.setOutputFileName(file_path)
        doc.print_(printer)

        QMessageBox.information(self, "موفقیت", "فاکتور PDF با موفقیت ایجاد گردید.")

    def direct_print_wedding(self, contract_id):
        html_content, _ = self.generate_pdf_html(contract_id)
        if not html_content:
            return

        doc = QTextDocument()
        doc.setHtml(html_content)
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        dialog = QPrintDialog(printer, self)
        if dialog.exec() == QPrintDialog.DialogCode.Accepted:
            doc.print_(printer)

    # --- زبانه ۲: تبلیغاتی و بیوتی ---
    def setup_commercial_tab(self):
        layout = QHBoxLayout()
        form_box = QGroupBox("ثبت پروژه جدید تبلیغاتی / بیوتی")
        form_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        form_layout = QFormLayout()

        self.c_title = QLineEdit()
        self.c_type = QComboBox()
        self.c_type.addItems(["تبلیغاتی", "بیوتی", "تولدی"])
        self.c_cameras = QSpinBox()
        self.c_cameras.setValue(1)
        self.c_amount = QLineEdit()
        self.c_amount.textChanged.connect(lambda t: self.c_amount.setText(format_number(t)))
        self.c_paid = QLineEdit()
        self.c_paid.textChanged.connect(lambda t: self.c_paid.setText(format_number(t)))
        self.c_date = QLineEdit(jdatetime.date.today().strftime("%Y/%m/%d"))
        self.c_desc = QTextEdit()
        self.c_desc.setFixedHeight(60)

        form_layout.addRow("عنوان پروژه *:", self.c_title)
        form_layout.addRow("نوع پروژه:", self.c_type)
        form_layout.addRow("تعداد دوربین:", self.c_cameras)
        form_layout.addRow("مبلغ کل (تومان) *:", self.c_amount)
        form_layout.addRow("مبلغ پرداختی (تومان):", self.c_paid)
        form_layout.addRow("تاریخ پروژه:", self.c_date)
        form_layout.addRow("توضیحات کامل:", self.c_desc)

        btn_save = QPushButton("ثبت پروژه")
        btn_save.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        btn_save.setStyleSheet("background-color: #27ae60; color: white; padding: 6px;")
        btn_save.clicked.connect(self.save_commercial_project)
        form_layout.addRow(btn_save)

        form_box.setLayout(form_layout)
        layout.addWidget(form_box, 1)

        self.c_table = QTableWidget()
        self.c_table.setColumnCount(9)
        self.c_table.setHorizontalHeaderLabels(["ردیف", "عنوان", "نوع", "دوربین", "مبلغ کل", "پرداختی", "تاریخ", "توضیحات", "حذف"])
        self.c_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.c_table, 2)

        self.tab_commercial.setLayout(layout)
        self.load_commercial_projects()

    def save_commercial_project(self):
        title = self.c_title.text().strip()
        amount = parse_number(self.c_amount.text())
        if not title or amount <= 0:
            QMessageBox.warning(self, "خطا", "لطفاً عنوان و مبلغ کل را به درستی وارد کنید.")
            return

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO commercial_projects (title, project_type, camera_count, total_amount, paid_amount, project_date, description)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (title, self.c_type.currentText(), self.c_cameras.value(),
              amount, parse_number(self.c_paid.text()), self.c_date.text(), self.c_desc.toPlainText()))

        conn.commit()
        conn.close()
        QMessageBox.information(self, "موفقیت", "پروژه ثبت شد.")
        self.load_commercial_projects()

    def load_commercial_projects(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, title, project_type, camera_count, total_amount, paid_amount, project_date, description FROM commercial_projects ORDER BY id DESC")
        rows = cursor.fetchall()
        conn.close()

        self.c_table.setRowCount(0)
        for r_idx, row in enumerate(rows):
            self.c_table.insertRow(r_idx)
            for c_idx, val in enumerate(row):
                val_str = f"{val:,}" if c_idx in [4, 5] and isinstance(val, int) else str(val)
                self.c_table.setItem(r_idx, c_idx, QTableWidgetItem(val_str))

            btn_del = QPushButton("حذف")
            btn_del.setStyleSheet("background-color: #e74c3c; color: white;")
            btn_del.clicked.connect(lambda _, pid=row[0]: self.delete_commercial_project(pid))
            self.c_table.setCellWidget(r_idx, 8, btn_del)

    def delete_commercial_project(self, p_id):
        if not ask_security_password(self):
            return
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM commercial_projects WHERE id=?", (p_id,))
        conn.commit()
        conn.close()
        self.load_commercial_projects()

    # --- زبانه ۳: بخش پرسنل و کارکنان ---
    def setup_staff_tab(self):
        layout = QVBoxLayout()
        forms_layout = QHBoxLayout()

        person_box = QGroupBox("تعریف نیروی جدید")
        person_form = QFormLayout()
        self.txt_person_name = QLineEdit()
        self.combo_person_role = QComboBox()
        self.combo_person_role.addItems(self.ROLES)
        btn_add_person = QPushButton("ثبت فرد جدید")
        btn_add_person.clicked.connect(self.add_person)

        person_form.addRow("نام و نام خانوادگی *:", self.txt_person_name)
        person_form.addRow("تخصص / نقش:", self.combo_person_role)
        person_form.addRow(btn_add_person)
        person_box.setLayout(person_form)

        trans_box = QGroupBox("ثبت حقوق و پرداختی به کارکنان")
        trans_form = QFormLayout()

        self.combo_staff_persons = QComboBox()
        self.txt_staff_amount = QLineEdit()
        self.txt_staff_amount.textChanged.connect(lambda t: self.txt_staff_amount.setText(format_number(t)))
        self.txt_staff_desc = QLineEdit()

        btn_add_trans = QPushButton("ثبت پرداخت حقوق")
        btn_add_trans.setStyleSheet("background-color: #8e44ad; color: white; font-weight: bold;")
        btn_add_trans.clicked.connect(self.add_staff_payment)

        trans_form.addRow("انتخاب فرد پرسنل *:", self.combo_staff_persons)
        trans_form.addRow("مبلغ پرداخت (تومان) *:", self.txt_staff_amount)
        trans_form.addRow("بابت / توضیحات:", self.txt_staff_desc)
        trans_form.addRow(btn_add_trans)
        trans_box.setLayout(trans_form)

        forms_layout.addWidget(person_box, 1)
        forms_layout.addWidget(trans_box, 2)
        layout.addLayout(forms_layout)

        self.staff_table = QTableWidget()
        self.staff_table.setColumnCount(6)
        self.staff_table.setHorizontalHeaderLabels(["ردیف", "نام فرد", "نقش/تخصص", "مبلغ پرداخت (تومان)", "توضیحات", "حذف"])
        self.staff_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.staff_table)

        self.tab_staff.setLayout(layout)
        self.load_persons_combo()
        self.load_staff_table()

    def add_person(self):
        name = self.txt_person_name.text().strip()
        if not name:
            QMessageBox.warning(self, "خطا", "لطفاً نام فرد را وارد کنید.")
            return

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO persons (name, role) VALUES (?, ?)", (name, self.combo_person_role.currentText()))
        conn.commit()
        conn.close()

        self.txt_person_name.clear()
        self.load_persons_combo()

    def load_persons_combo(self):
        self.combo_staff_persons.clear()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, name, role FROM persons")
        for p_id, name, role in cursor.fetchall():
            self.combo_staff_persons.addItem(f"{name} ({role})", p_id)
        conn.close()

    def add_staff_payment(self):
        person_id = self.combo_staff_persons.currentData()
        amount = parse_number(self.txt_staff_amount.text())
        if not person_id or amount <= 0:
            QMessageBox.warning(self, "خطا", "لطفاً فرد و مبلغ معتبر وارد کنید.")
            return

        today = jdatetime.date.today()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO transactions (trans_type, category, person_id, amount, year, month, day, description)
            VALUES ('expense', 'حقوق کارکنان', ?, ?, ?, ?, ?, ?)
        ''', (person_id, amount, today.year, today.month, today.day, self.txt_staff_desc.text()))
        conn.commit()
        conn.close()

        self.txt_staff_amount.clear()
        self.txt_staff_desc.clear()
        self.load_staff_table()

    def load_staff_table(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT t.id, p.name, p.role, t.amount, t.description
            FROM transactions t
            JOIN persons p ON t.person_id = p.id
            ORDER BY t.id DESC
        ''')
        rows = cursor.fetchall()
        conn.close()

        self.staff_table.setRowCount(0)
        for r_idx, row in enumerate(rows):
            self.staff_table.insertRow(r_idx)
            self.staff_table.setItem(r_idx, 0, QTableWidgetItem(str(row[0])))
            self.staff_table.setItem(r_idx, 1, QTableWidgetItem(row[1]))
            self.staff_table.setItem(r_idx, 2, QTableWidgetItem(row[2]))
            self.staff_table.setItem(r_idx, 3, QTableWidgetItem(f"{row[3]:,}"))
            self.staff_table.setItem(r_idx, 4, QTableWidgetItem(row[4] if row[4] else "-"))

            btn_del = QPushButton("حذف")
            btn_del.setStyleSheet("background-color: #e74c3c; color: white;")
            btn_del.clicked.connect(lambda _, tid=row[0]: self.delete_staff_trans(tid))
            self.staff_table.setCellWidget(r_idx, 5, btn_del)

    def delete_staff_trans(self, tid):
        if not ask_security_password(self):
            return
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM transactions WHERE id=?", (tid,))
        conn.commit()
        conn.close()
        self.load_staff_table()

    # --- زبانه ۴: مدیریت هزینه‌ها ---
    def setup_expenses_tab(self):
        layout = QHBoxLayout()
        form_box = QGroupBox("ثبت هزینه جدید (قبوض، کرایه و...)")
        form_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        form_layout = QFormLayout()

        self.exp_title = QLineEdit()
        self.exp_cat = QComboBox()
        self.exp_cat.addItems(["قبوض (آب، برق، گاز، تلفن)", "کرایه آتلیه / دفتر", "خرید تجهیزات و مصرفی", "تبلیغات و بازاریابی", "سایر هزینه‌ها"])
        self.exp_amount = QLineEdit()
        self.exp_amount.textChanged.connect(lambda t: self.exp_amount.setText(format_number(t)))
        self.exp_date = QLineEdit(jdatetime.date.today().strftime("%Y/%m/%d"))
        self.exp_desc = QTextEdit()
        self.exp_desc.setFixedHeight(60)

        form_layout.addRow("عنوان هزینه *:", self.exp_title)
        form_layout.addRow("دسته‌بندی:", self.exp_cat)
        form_layout.addRow("مبلغ (تومان) *:", self.exp_amount)
        form_layout.addRow("تاریخ هزینه:", self.exp_date)
        form_layout.addRow("توضیحات:", self.exp_desc)

        btn_save = QPushButton("ثبت هزینه")
        btn_save.setStyleSheet("background-color: #c0392b; color: white; font-weight: bold; padding: 6px;")
        btn_save.clicked.connect(self.save_expense)
        form_layout.addRow(btn_save)

        form_box.setLayout(form_layout)
        layout.addWidget(form_box, 1)

        self.exp_table = QTableWidget()
        self.exp_table.setColumnCount(7)
        self.exp_table.setHorizontalHeaderLabels(["ردیف", "عنوان", "دسته‌بندی", "مبلغ (تومان)", "تاریخ", "توضیحات", "حذف"])
        self.exp_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.exp_table, 2)

        self.tab_expenses.setLayout(layout)
        self.load_expenses_table()

    def save_expense(self):
        title = self.exp_title.text().strip()
        amount = parse_number(self.exp_amount.text())
        if not title or amount <= 0:
            QMessageBox.warning(self, "خطا", "لطفاً عنوان و مبلغ را وارد کنید.")
            return

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO expenses (title, category, amount, date_str, description) VALUES (?, ?, ?, ?, ?)",
                       (title, self.exp_cat.currentText(), amount, self.exp_date.text(), self.exp_desc.toPlainText()))
        conn.commit()
        conn.close()

        self.exp_title.clear()
        self.exp_amount.clear()
        self.load_expenses_table()

    def load_expenses_table(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, title, category, amount, date_str, description FROM expenses ORDER BY id DESC")
        rows = cursor.fetchall()
        conn.close()

        self.exp_table.setRowCount(0)
        for r_idx, row in enumerate(rows):
            self.exp_table.insertRow(r_idx)
            self.exp_table.setItem(r_idx, 0, QTableWidgetItem(str(row[0])))
            self.exp_table.setItem(r_idx, 1, QTableWidgetItem(row[1]))
            self.exp_table.setItem(r_idx, 2, QTableWidgetItem(row[2]))
            self.exp_table.setItem(r_idx, 3, QTableWidgetItem(f"{row[3]:,}"))
            self.exp_table.setItem(r_idx, 4, QTableWidgetItem(row[4]))
            self.exp_table.setItem(r_idx, 5, QTableWidgetItem(row[5] if row[5] else "-"))

            btn_del = QPushButton("حذف")
            btn_del.setStyleSheet("background-color: #e74c3c; color: white;")
            btn_del.clicked.connect(lambda _, eid=row[0]: self.delete_expense(eid))
            self.exp_table.setCellWidget(r_idx, 6, btn_del)

    def delete_expense(self, eid):
        if not ask_security_password(self):
            return
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM expenses WHERE id=?", (eid,))
        conn.commit()
        conn.close()
        self.load_expenses_table()

    # --- زبانه ۵: گزارش مالی جامع ---
    def setup_reports_tab(self):
        layout = QVBoxLayout()

        top_layout = QHBoxLayout()
        self.combo_report_type = QComboBox()
        self.combo_report_type.addItems(["گزارش ماهانه", "گزارش سالانه", "گزارش کل"])
        top_layout.addWidget(QLabel("نوع گزارش:"))
        top_layout.addWidget(self.combo_report_type)

        btn_calc = QPushButton("محاسبه گزارش")
        btn_calc.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold;")
        btn_calc.clicked.connect(self.calculate_financial_report)
        top_layout.addWidget(btn_calc)

        btn_chart = QPushButton("نمایش نمودار مقایسه‌ای")
        btn_chart.setStyleSheet("background-color: #2980b9; color: white; font-weight: bold;")
        btn_chart.clicked.connect(self.show_financial_chart)
        top_layout.addWidget(btn_chart)

        layout.addLayout(top_layout)

        self.lbl_rep_summary = QLabel("ورودی کل: ۰ تومان | خروجی کل (هزینه‌ها و حقوق): ۰ تومان | سود خالص: ۰ تومان")
        self.lbl_rep_summary.setStyleSheet("font-size: 12pt; font-weight: bold; background-color: #e8f8f5; padding: 10px; border-radius: 5px;")
        layout.addWidget(self.lbl_rep_summary)

        self.rep_table = QTableWidget()
        self.rep_table.setColumnCount(5)
        self.rep_table.setHorizontalHeaderLabels(["ردیف", "بخش", "نوع تراکنش", "مبلغ (تومان)", "تاریخ / شرح"])
        self.rep_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.rep_table)

        self.tab_reports.setLayout(layout)

    def calculate_financial_report(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        cursor.execute("SELECT SUM(total_amount) FROM wedding_contracts")
        w_in = cursor.fetchone()[0] or 0

        cursor.execute("SELECT SUM(total_amount) FROM commercial_projects")
        c_in = cursor.fetchone()[0] or 0

        cursor.execute("SELECT SUM(amount) FROM expenses")
        exp_out = cursor.fetchone()[0] or 0

        cursor.execute("SELECT SUM(amount) FROM transactions WHERE trans_type='expense'")
        staff_out = cursor.fetchone()[0] or 0

        conn.close()

        total_in = w_in + c_in
        total_out = exp_out + staff_out
        profit = total_in - total_out

        self.lbl_rep_summary.setText(f"ورودی کل: {total_in:,} تومان | خروجی کل (هزینه‌ها و حقوق): {total_out:,} تومان | سود خالص: {profit:,} تومان")

        self.rep_table.setRowCount(0)
        records = [
            ("قراردادهای عروس و داماد", "درآمد (ورودی)", w_in, "مجموع ورودی‌ها"),
            ("پروژه‌های تبلیغاتی", "درآمد (ورودی)", c_in, "مجموع ورودی‌ها"),
            ("هزینه‌های جاری و قبوض", "هزینه (خروجی)", exp_out, "مجموع هزینه‌ها"),
            ("حقوق و پرداختی پرسنل", "هزینه (خروجی)", staff_out, "مجموع پرداختی‌ها")
        ]
        for idx, rec in enumerate(records):
            self.rep_table.insertRow(idx)
            self.rep_table.setItem(idx, 0, QTableWidgetItem(str(idx + 1)))
            self.rep_table.setItem(idx, 1, QTableWidgetItem(rec[0]))
            self.rep_table.setItem(idx, 2, QTableWidgetItem(rec[1]))
            self.rep_table.setItem(idx, 3, QTableWidgetItem(f"{rec[2]:,}"))
            self.rep_table.setItem(idx, 4, QTableWidgetItem(rec[3]))

    def show_financial_chart(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT SUM(total_amount) FROM wedding_contracts")
        w_in = cursor.fetchone()[0] or 0
        cursor.execute("SELECT SUM(total_amount) FROM commercial_projects")
        c_in = cursor.fetchone()[0] or 0
        cursor.execute("SELECT SUM(amount) FROM expenses")
        exp_out = cursor.fetchone()[0] or 0
        cursor.execute("SELECT SUM(amount) FROM transactions WHERE trans_type='expense'")
        staff_out = cursor.fetchone()[0] or 0
        conn.close()

        labels = ['درآمد عروسی', 'درآمد تبلیغاتی', 'هزینه‌ها/قبوض', 'حقوق پرسنل']
        amounts = [w_in, c_in, exp_out, staff_out]

        plt.figure(figsize=(8, 5))
        plt.bar(labels, amounts, color=['#27ae60', '#2980b9', '#e74c3c', '#8e44ad'])
        plt.title("نمودار مقایسه‌ای درآمدها و هزینه‌های IMART STUDIO")
        plt.ylabel("مبلغ (تومان)")
        plt.show()

    # --- زبانه ۶: جستجوی پیشرفته ---
    def setup_search_tab(self):
        layout = QVBoxLayout()
        search_bar = QHBoxLayout()

        self.txt_search_keyword = QLineEdit()
        self.txt_search_keyword.setPlaceholderText("جستجو بر اساس نام شخص، تاریخ یا عنوان...")
        btn_search = QPushButton("جستجو")
        btn_search.setStyleSheet("background-color: #16a085; color: white; font-weight: bold;")
        btn_search.clicked.connect(self.perform_search)

        search_bar.addWidget(QLabel("عبارت جستجو:"))
        search_bar.addWidget(self.txt_search_keyword)
        search_bar.addWidget(btn_search)
        layout.addLayout(search_bar)

        self.search_table = QTableWidget()
        self.search_table.setColumnCount(5)
        self.search_table.setHorizontalHeaderLabels(["ردیف", "بخش", "نام / عنوان", "تاریخ", "مبلغ / شرح"])
        self.search_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.search_table)

        self.tab_search.setLayout(layout)

    def perform_search(self):
        kw = f"%{self.txt_search_keyword.text().strip()}%"
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        results = []
        cursor.execute("SELECT id, 'قرارداد عروسی', groom_name || ' و ' || bride_name, ceremony_date, total_amount FROM wedding_contracts WHERE groom_name LIKE ? OR bride_name LIKE ? OR ceremony_date LIKE ?", (kw, kw, kw))
        results.extend(cursor.fetchall())

        cursor.execute("SELECT id, 'پروژه تبلیغاتی', title, project_date, total_amount FROM commercial_projects WHERE title LIKE ? OR project_date LIKE ?", (kw, kw))
        results.extend(cursor.fetchall())

        cursor.execute("SELECT id, 'هزینه‌ها', title, date_str, amount FROM expenses WHERE title LIKE ? OR date_str LIKE ?", (kw, kw))
        results.extend(cursor.fetchall())

        conn.close()

        self.search_table.setRowCount(0)
        for r_idx, row in enumerate(results):
            self.search_table.insertRow(r_idx)
            self.search_table.setItem(r_idx, 0, QTableWidgetItem(str(row[0])))
            self.search_table.setItem(r_idx, 1, QTableWidgetItem(row[1]))
            self.search_table.setItem(r_idx, 2, QTableWidgetItem(row[2]))
            self.search_table.setItem(r_idx, 3, QTableWidgetItem(str(row[3])))
            self.search_table.setItem(r_idx, 4, QTableWidgetItem(f"{row[4]:,} تومان"))

    # --- زبانه ۷: انبار تجهیزات ---
    def setup_inventory_tab(self):
        layout = QVBoxLayout()
        self.inv_table = QTableWidget()
        self.inv_table.setColumnCount(4)
        self.inv_table.setHorizontalHeaderLabels(["نام آیتم / پکیج", "تعداد کل", "استفاده شده در پروژه‌ها", "موجودی باقی‌مانده"])
        self.inv_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.inv_table)
        self.tab_inventory.setLayout(layout)
        self.load_inventory()

    def load_inventory(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT item_name, total_count, used_count FROM inventory")
        rows = cursor.fetchall()
        conn.close()

        self.inv_table.setRowCount(0)
        for r_idx, (name, total, used) in enumerate(rows):
            remain = total - used
            self.inv_table.insertRow(r_idx)
            self.inv_table.setItem(r_idx, 0, QTableWidgetItem(name))
            self.inv_table.setItem(r_idx, 1, QTableWidgetItem(str(total)))
            self.inv_table.setItem(r_idx, 2, QTableWidgetItem(str(used)))
            self.inv_table.setItem(r_idx, 3, QTableWidgetItem(str(remain)))

    # --- زبانه ۸: کارت‌های بانکی ---
    def setup_banks_tab(self):
        layout = QHBoxLayout()
        form_box = QGroupBox("افزودن حساب بانکی")
        form_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        form_layout = QFormLayout()

        self.b_name = QLineEdit()
        self.b_card = QLineEdit()
        self.b_sheba = QLineEdit()

        form_layout.addRow("نام بانک *:", self.b_name)
        form_layout.addRow("شماره کارت:", self.b_card)
        form_layout.addRow("شماره شبا:", self.b_sheba)

        btn_save = QPushButton("ذخیره کارت")
        btn_save.setStyleSheet("background-color: #16a085; color: white; padding: 6px;")
        btn_save.clicked.connect(self.save_bank_card)
        form_layout.addRow(btn_save)

        form_box.setLayout(form_layout)
        layout.addWidget(form_box, 1)

        self.b_table = QTableWidget()
        self.b_table.setColumnCount(3)
        self.b_table.setHorizontalHeaderLabels(["نام بانک", "شماره کارت", "شماره شبا"])
        self.b_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.b_table, 2)

        self.tab_banks.setLayout(layout)
        self.load_bank_cards()

    def save_bank_card(self):
        if not self.b_name.text().strip():
            return
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO bank_cards (bank_name, card_number, sheba_number) VALUES (?, ?, ?)",
                       (self.b_name.text(), self.b_card.text(), self.b_sheba.text()))
        conn.commit()
        conn.close()
        self.load_bank_cards()
        self.load_bank_combo()

    def load_bank_cards(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT bank_name, card_number, sheba_number FROM bank_cards")
        rows = cursor.fetchall()
        conn.close()

        self.b_table.setRowCount(0)
        for r_idx, row in enumerate(rows):
            self.b_table.insertRow(r_idx)
            for c_idx, val in enumerate(row):
                self.b_table.setItem(r_idx, c_idx, QTableWidgetItem(str(val)))

    # --- زبانه ۹: مدیریت چک‌ها ---
    def setup_checks_tab(self):
        layout = QHBoxLayout()
        form_box = QGroupBox("ثبت چک جدید")
        form_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        form_layout = QFormLayout()

        self.chk_num = QLineEdit()
        self.chk_bank = QLineEdit()
        self.chk_amount = QLineEdit()
        self.chk_amount.textChanged.connect(lambda t: self.chk_amount.setText(format_number(t)))
        self.chk_date = QLineEdit(jdatetime.date.today().strftime("%Y/%m/%d"))
        self.chk_desc = QLineEdit()

        form_layout.addRow("شماره چک *:", self.chk_num)
        form_layout.addRow("بانک صادرکننده:", self.chk_bank)
        form_layout.addRow("مبلغ چک (تومان) *:", self.chk_amount)
        form_layout.addRow("تاریخ سررسید (شمسی):", self.chk_date)
        form_layout.addRow("توضیحات:", self.chk_desc)

        btn_save = QPushButton("ثبت چک")
        btn_save.setStyleSheet("background-color: #2980b9; color: white; padding: 6px;")
        btn_save.clicked.connect(self.save_check)
        form_layout.addRow(btn_save)

        form_box.setLayout(form_layout)
        layout.addWidget(form_box, 1)

        self.chk_table = QTableWidget()
        self.chk_table.setColumnCount(7)
        self.chk_table.setHorizontalHeaderLabels(["ردیف", "شماره چک", "بانک", "مبلغ (تومان)", "سررسید", "وضعیت پاس شدن", "حذف"])
        self.chk_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.chk_table, 2)

        self.tab_checks.setLayout(layout)
        self.load_checks_table()

    def save_check(self):
        num = self.chk_num.text().strip()
        amount = parse_number(self.chk_amount.text())
        if not num or amount <= 0:
            QMessageBox.warning(self, "خطا", "اطلاعات چک را کامل کنید.")
            return

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO checks (check_number, bank_name, amount, due_date, description) VALUES (?, ?, ?, ?, ?)",
                       (num, self.chk_bank.text(), amount, self.chk_date.text(), self.chk_desc.text()))
        conn.commit()
        conn.close()

        self.load_checks_table()

    def load_checks_table(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, check_number, bank_name, amount, due_date, is_passed FROM checks ORDER BY id DESC")
        rows = cursor.fetchall()
        conn.close()

        self.chk_table.setRowCount(0)
        for r_idx, row in enumerate(rows):
            c_id, num, bank, amount, due, is_passed = row
            self.chk_table.insertRow(r_idx)
            self.chk_table.setItem(r_idx, 0, QTableWidgetItem(str(c_id)))
            self.chk_table.setItem(r_idx, 1, QTableWidgetItem(num))
            self.chk_table.setItem(r_idx, 2, QTableWidgetItem(bank))
            self.chk_table.setItem(r_idx, 3, QTableWidgetItem(f"{amount:,}"))
            self.chk_table.setItem(r_idx, 4, QTableWidgetItem(due))

            btn_pass = QPushButton("✅ پاس شده" if is_passed else "⏳ مانده (پاس‌کردن)")
            btn_pass.setStyleSheet("background-color: green; color: white;" if is_passed else "background-color: orange; color: white;")
            btn_pass.clicked.connect(lambda _, cid=c_id, state=is_passed: self.toggle_check_pass(cid, state))
            self.chk_table.setCellWidget(r_idx, 5, btn_pass)

            btn_del = QPushButton("حذف")
            btn_del.setStyleSheet("background-color: #e74c3c; color: white;")
            btn_del.clicked.connect(lambda _, cid=c_id: self.delete_check(cid))
            self.chk_table.setCellWidget(r_idx, 6, btn_del)

    def toggle_check_pass(self, cid, current_state):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("UPDATE checks SET is_passed=? WHERE id=?", (0 if current_state else 1, cid))
        conn.commit()
        conn.close()
        self.load_checks_table()

    def delete_check(self, cid):
        if not ask_security_password(self):
            return
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM checks WHERE id=?", (cid,))
        conn.commit()
        conn.close()
        self.load_checks_table()

    # --- زبانه ۱۰: صدور رسید وجه ---
    def setup_receipt_tab(self):
        layout = QVBoxLayout()
        form_box = QGroupBox("صدور رسید بیعانه / وجه دریافتی")
        form_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        form_layout = QFormLayout()

        self.rc_name = QLineEdit()
        self.rc_amount = QLineEdit()
        self.rc_amount.textChanged.connect(lambda t: self.rc_amount.setText(format_number(t)))
        self.rc_for = QLineEdit("خدمات آتلیه و فیلمبرداری")
        self.rc_date = QLineEdit(jdatetime.date.today().strftime("%Y/%m/%d"))

        form_layout.addRow("مبلغ به دریافت از آقا / خانم *:", self.rc_name)
        form_layout.addRow("مبلغ بیعانه (تومان) *:", self.rc_amount)
        form_layout.addRow("بابت خدمات:", self.rc_for)
        form_layout.addRow("تاریخ:", self.rc_date)

        btn_print = QPushButton("🖨️ چاپ و صدور رسید وجه")
        btn_print.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; padding: 8px;")
        btn_print.clicked.connect(self.print_receipt)
        form_layout.addRow(btn_print)

        form_box.setLayout(form_layout)
        layout.addWidget(form_box)
        layout.addStretch()
        self.tab_receipt.setLayout(layout)

    def print_receipt(self):
        name = self.rc_name.text().strip()
        amount = parse_number(self.rc_amount.text())
        if not name or amount <= 0:
            QMessageBox.warning(self, "خطا", "لطفاً نام پرداخت‌کننده و مبلغ را به درستی وارد کنید.")
            return

        html_content = f"""
        <div dir="rtl" style="font-family: '{APP_FONT_FAMILY}', 'Tahoma'; padding: 30px; border: 2px solid #2c3e50; border-radius: 10px;">
            <h2 style="text-align: center; color: #1F4E78; margin-bottom: 5px;">IMART STUDIO</h2>
            <h3 style="text-align: center; margin-top: 0;">رسید دریافتی وجه / بیعانه</h3>
            <hr>
            <p style="font-size: 14pt; line-height: 2;">
                بدین‌وسیله گواهی می‌شود مبلغ <b>{amount:,} تومان</b> 
                از جناب آقای / سرکار خانم <b>{name}</b> 
                بابت <b>{self.rc_for.text()}</b> در تاریخ <b>{self.rc_date.text()}</b> دریافت گردید.
            </p>
            <br><br><br>
            <table style="width: 100%; text-align: center; font-size: 12pt;">
                <tr>
                    <td>امضاء دریافت‌کننده (IMART STUDIO)</td>
                    <td>امضاء و مهر پرداخت‌کننده</td>
                </tr>
            </table>
        </div>
        """

        doc = QTextDocument()
        doc.setHtml(html_content)
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        dialog = QPrintDialog(printer, self)
        if dialog.exec() == QPrintDialog.DialogCode.Accepted:
            doc.print_(printer)

    # --- زبانه ۱۱: سال کاری ---
    def setup_workyear_tab(self):
        layout = QVBoxLayout()
        box = QGroupBox("مدیریت سال کاری و بستن حساب‌ها")
        box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        vbox = QVBoxLayout()

        lbl_info = QLabel("بستن سال کاری تمامی اطلاعات قراردادها، هزینه‌ها و پروژه‌ها را به یک فایل پشتیبان با نام سال منتقل کرده و جدول‌ها را جهت شروع سال کاری جدید خالی می‌کند.")
        lbl_info.setWordWrap(True)
        vbox.addWidget(lbl_info)

        btn_close_year = QPushButton("🔒 بستن سال کاری و انتقال به سال جدید")
        btn_close_year.setStyleSheet("background-color: #c0392b; color: white; font-weight: bold; padding: 10px;")
        btn_close_year.clicked.connect(self.close_work_year)
        vbox.addWidget(btn_close_year)

        box.setLayout(vbox)
        layout.addWidget(box)
        layout.addStretch()
        self.tab_workyear.setLayout(layout)

    def close_work_year(self):
        if not ask_security_password(self):
            return

        reply = QMessageBox.question(self, "تایید نهایی", "آیا مطمئن هستید که می‌خواهید سال کاری را ببندید؟ از دیتابیس فعلی بک‌آپ گرفته خواهد شد.", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            today_str = jdatetime.date.today().strftime("%Y_%m_%d")
            backup_name = f"studio_accounting_archive_{today_str}.db"
            shutil.copy(DB_NAME, backup_name)

            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM wedding_contracts")
            cursor.execute("DELETE FROM wedding_deposits")
            cursor.execute("DELETE FROM commercial_projects")
            cursor.execute("DELETE FROM expenses")
            cursor.execute("DELETE FROM transactions")
            cursor.execute("DELETE FROM checks")
            cursor.execute("UPDATE inventory SET used_count=0")
            conn.commit()
            conn.close()

            QMessageBox.information(self, "موفقیت", f"سال کاری با موفقیت بسته شد. فایل آرشیو با نام {backup_name} ذخیره گردید.")
            self.load_wedding_contracts()
            self.load_commercial_projects()
            self.load_expenses_table()
            self.load_staff_table()
            self.load_inventory()

    # --- زبانه ۱۲: درباره برنامه ---
    def setup_about_tab(self):
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        card = QGroupBox("درباره نرم‌افزار IMART STUDIO")
        card_layout = QVBoxLayout()

        lbl_title = QLabel("نرم‌افزار جامع مدیریت مالی و حسابداری IMART STUDIO")
        lbl_title.setFont(QFont(APP_FONT_FAMILY, 14, QFont.Weight.Bold))
        lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)

        lbl_desc = QLabel(
            "نسخه: ۸.۰\n"
            "طراحی و توسعه اختصاصی جهت آتلیه‌ها، استودیوهای فیلمبرداری و پروژه‌های تولید محتوا.\n\n"
            "📞 شماره تماس پشتیبانی: 09173736618\n"
            "تمامی حقوق این نرم‌افزار محفوظ می‌باشد."
        )
        lbl_desc.setFont(QFont(APP_FONT_FAMILY, 11))
        lbl_desc.setAlignment(Qt.AlignmentFlag.AlignCenter)

        card_layout.addWidget(lbl_title)
        card_layout.addWidget(lbl_desc)
        card.setLayout(card_layout)

        layout.addWidget(card)
        self.tab_about.setLayout(layout)

    # --- متدهای عمومی بک‌آپ، ریستور و تغییر رمز ---
    def backup_db(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "ذخیره فایل پشتیبان", f"backup_{jdatetime.date.today().strftime('%Y_%m_%d')}.db", "Database Files (*.db)")
        if file_path:
            shutil.copy(DB_NAME, file_path)
            QMessageBox.information(self, "موفقیت", "پشتیبان‌گیری با موفقیت انجام شد.")

    def restore_db(self):
        if not ask_security_password(self):
            return
        file_path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل پشتیبان", "", "Database Files (*.db)")
        if file_path:
            shutil.copy(file_path, DB_NAME)
            QMessageBox.information(self, "موفقیت", "پایگاه داده با موفقیت بازیابی شد. برنامه را دوباره اجرا کنید.")

    def change_password(self):
        if not ask_security_password(self):
            return
        new_pass, ok = QInputDialog.getText(self, "تغییر رمز عبور", "رمز عبور جدید را وارد کنید:", QLineEdit.EchoMode.Password)
        if ok and new_pass.strip():
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("UPDATE settings SET value=? WHERE key='app_password'", (new_pass.strip(),))
            conn.commit()
            conn.close()
            QMessageBox.information(self, "موفقیت", "رمز عبور جدید با موفقیت ثبت گردید.")


# --- نقطه‌ی ورود و اجرای برنامه ---
if __name__ == '__main__':
    app = QApplication(sys.argv)

    # ⚠️ تشخیص فونت باید بعد از ساخت QApplication انجام شود
    APP_FONT_FAMILY = setup_fonts()
    app.setFont(QFont(APP_FONT_FAMILY, 10))

    # ─── نمایش پنجره بارگذاری ───
    loading = LoadingScreen()
    loading.show()

    # لیست مراحل بارگذاری (متن، درصد)
    steps = [
        ("راه‌اندازی موتور برنامه...", 5),
        ("بارگذاری فونت‌های فارسی...", 12),
        ("اتصال به پایگاه داده SQLite...", 22),
        ("ایجاد جداول در صورت نیاز...", 35),
        ("بارگذاری تنظیمات امنیتی...", 45),
        ("آماده‌سازی رابط کاربری...", 58),
        ("بارگذاری تب‌های برنامه...", 72),
        ("بارگذاری قراردادها و پروژه‌ها...", 85),
        ("بارگذاری انبار و کارت‌های بانکی...", 93),
        ("آماده‌سازی نهایی...", 98),
        ("ورود به سیستم...", 100),
    ]

    # پنجره اصلی برنامه (هنوز نمایش داده نمی‌شود)
    main_win = StudioAccountingApp()

    # اجرای پله‌ای مراحل با QTimer
    current_step = {"index": 0}

    def finish_loading():
        loading.close()
        if main_win.prompt_login():
            main_win.stack.setCurrentWidget(main_win.dashboard_screen)
            main_win.show()
        else:
            app.quit()

    def run_step():
        if current_step["index"] < len(steps):
            text, percent = steps[current_step["index"]]
            loading.set_progress(percent, text)

            # عملیات واقعی هر مرحله
            if current_step["index"] == 2:
                init_db()
            elif current_step["index"] == 5:
                main_win.setup_dashboard_ui()
            elif current_step["index"] == 6:
                main_win.setup_main_app_ui()

            current_step["index"] += 1
            QTimer.singleShot(300, run_step)
        else:
            QTimer.singleShot(400, finish_loading)

    # شروع اجرای مراحل
    QTimer.singleShot(400, run_step)

    sys.exit(app.exec())