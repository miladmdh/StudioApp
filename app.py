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
    QProgressBar, QGraphicsDropShadowEffect, QListWidget, QListWidgetItem
)
from PyQt6.QtCore import Qt, QDate, QTimer
from PyQt6.QtGui import QFont, QIcon, QColor, QFontDatabase, QTextDocument, QPixmap, QPainter, QLinearGradient, QBrush, QPageSize
from PyQt6.QtPrintSupport import QPrinter, QPrintDialog

import matplotlib
matplotlib.use('QtAgg')
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

# تنظیم ID جهت آیکون ویندوز
try:
    import ctypes
    myappid = 'imartstudio.accounting.v9.0'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass

DB_NAME = "studio_accounting.db"
APP_VERSION = "9.0"
DEVELOPER_NAME = "میلاد محمدحسینی"

# ⚠️ مقدار پیش‌فرض فونت (بعد از ساخت QApplication نهایی می‌شود)
APP_FONT_FAMILY = "Tahoma"
INVOICE_FONT_FAMILY = "Tahoma"


def setup_fonts():
    """تشخیص فونت مناسب فارسی - باید بعد از ساخت QApplication صدا زده شود"""
    global INVOICE_FONT_FAMILY
    font_family = "Tahoma"
    available_fonts = QFontDatabase.families()

    # لاگ کن که چه فونت‌هایی در دسترس هست
    has_nazanin = "B Nazanin" in available_fonts
    has_yekan = "B Yekan" in available_fonts

    print(f"[FONT] B Nazanin installed: {has_nazanin}")
    print(f"[FONT] B Yekan installed: {has_yekan}")

    if has_yekan:
        font_family = "B Yekan"
    elif has_nazanin:
        font_family = "B Nazanin"
    else:
        font_family = "Tahoma"

    if has_nazanin:
        INVOICE_FONT_FAMILY = "B Nazanin"
    elif has_yekan:
        INVOICE_FONT_FAMILY = "B Yekan"
    else:
        INVOICE_FONT_FAMILY = "Tahoma"

    print(f"[FONT] Using for UI: {font_family}")
    print(f"[FONT] Using for invoice: {INVOICE_FONT_FAMILY}")
    return font_family


def setup_matplotlib_font():
    """پیدا کردن فونت فارسی برای matplotlib"""
    font_paths = [
        "C:/Windows/Fonts/BNAZANIN.TTF",
        "C:/Windows/Fonts/bnazanin.ttf",
        "C:/Windows/Fonts/BYekan.ttf",
        "C:/Windows/Fonts/byekan.ttf",
        "C:/Windows/Fonts/tahoma.ttf",
        "C:/Windows/Fonts/arial.ttf",
    ]
    for fp in font_paths:
        if os.path.exists(fp):
            try:
                fm.fontManager.addfont(fp)
                prop = fm.FontProperties(fname=fp)
                plt.rcParams['font.family'] = prop.get_name()
                plt.rcParams['axes.unicode_minus'] = False
                print(f"[MATPLOTLIB FONT] Using: {fp}")
                return True
            except Exception as e:
                print(f"[MATPLOTLIB FONT] Failed {fp}: {e}")
                continue
    plt.rcParams['axes.unicode_minus'] = False
    return False


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


def number_to_persian_words(num):
    """تبدیل عدد به حروف فارسی"""
    try:
        num = int(num)
    except Exception:
        return ""
    if num == 0:
        return "صفر"
    if num < 0:
        return "منفی " + number_to_persian_words(-num)

    yekan = ["", "یک", "دو", "سه", "چهار", "پنج", "شش", "هفت", "هشت", "نه"]
    dahgan = ["", "", "بیست", "سی", "چهل", "پنجاه", "شصت", "هفتاد", "هشتاد", "نود"]
    dah = ["ده", "یازده", "دوازده", "سیزده", "چهارده", "پانزده", "شانزده", "هفده", "هجده", "نوزده"]
    sadgan = ["", "صد", "دویست", "سیصد", "چهارصد", "پانصد", "ششصد", "هفتصد", "هشتصد", "نهصد"]
    scale = ["", " هزار", " میلیون", " میلیارد", " تریلیون"]

    def three_digit_to_words(n):
        result = []
        s = n // 100
        rem = n % 100
        if s > 0:
            result.append(sadgan[s])
        if rem >= 10 and rem < 20:
            result.append(dah[rem - 10])
        else:
            d = rem // 10
            y = rem % 10
            if d > 0:
                result.append(dahgan[d])
            if y > 0:
                result.append(yekan[y])
        return " و ".join(result)

    parts = []
    idx = 0
    while num > 0:
        chunk = num % 1000
        if chunk > 0:
            parts.insert(0, three_digit_to_words(chunk) + scale[idx])
        num //= 1000
        idx += 1

    return " و ".join(parts)


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


def get_setting(key, default=""):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("SELECT value FROM settings WHERE key=?", (key,))
    row = cursor.fetchone()
    conn.close()
    return row[0] if row else default


def set_setting(key, value):
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    cursor.execute("INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)", (key, str(value)))
    conn.commit()
    conn.close()


def get_current_year():
    """سال کاری فعلی"""
    return int(get_setting("current_work_year", jdatetime.date.today().year))


def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    cursor.execute('''CREATE TABLE IF NOT EXISTS settings (key TEXT PRIMARY KEY, value TEXT)''')
    cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('app_password', '123')")
    cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('studio_name', 'IMART STUDIO')")
    cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('studio_phone', '09173736618')")
    cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('studio_address', '')")
    cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('current_work_year', ?)", (str(jdatetime.date.today().year),))

    cursor.execute('''CREATE TABLE IF NOT EXISTS persons (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT,
        name TEXT NOT NULL,
        role TEXT NOT NULL,
        phone TEXT,
        work_year INTEGER
    )''')

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
            work_year INTEGER,
            FOREIGN KEY (person_id) REFERENCES persons(id)
        )
    ''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS expenses (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT, title TEXT, category TEXT, amount INTEGER,
        date_str TEXT, description TEXT, work_year INTEGER
    )''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS bank_cards (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        bank_name TEXT NOT NULL, card_number TEXT, sheba_number TEXT
    )''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS item_prices (
        item_name TEXT PRIMARY KEY,
        code TEXT, price INTEGER DEFAULT 0,
        is_package INTEGER DEFAULT 0, parent_package TEXT
    )''')

    default_items = [
        ("کلیپ فرمالیته", 15000000, 0, ""), ("کلیپ باغ", 8000000, 0, ""),
        ("یک دوربین", 5000000, 0, ""), ("دو دوربین", 9000000, 0, ""),
        ("کرین", 6000000, 0, ""), ("هلی شات", 7000000, 0, ""),
        ("عکاس مجلس", 4000000, 0, ""), ("پکیج طلایی VIP", 35000000, 1, "")
    ]
    for idx, (item, price, is_pkg, parent) in enumerate(default_items):
        code = f"SRV-{1000 + idx}"
        cursor.execute(
            "INSERT OR IGNORE INTO item_prices (item_name, code, price, is_package, parent_package) VALUES (?, ?, ?, ?, ?)",
            (item, code, price, is_pkg, parent)
        )

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS wedding_contracts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT,
            groom_name TEXT, bride_name TEXT, groom_phone TEXT, bride_phone TEXT,
            contract_date TEXT, ceremony_date TEXT, selected_items TEXT,
            total_amount INTEGER, discount INTEGER, paid_amount INTEGER,
            bank_id INTEGER, is_settled INTEGER DEFAULT 0, description TEXT,
            work_year INTEGER
        )
    ''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS wedding_deposits (
        id INTEGER PRIMARY KEY AUTOINCREMENT, contract_id INTEGER, amount INTEGER,
        deposit_date TEXT, bank_name TEXT, work_year INTEGER,
        FOREIGN KEY (contract_id) REFERENCES wedding_contracts(id)
    )''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS commercial_projects (
        id INTEGER PRIMARY KEY AUTOINCREMENT, code TEXT, title TEXT, project_type TEXT,
        camera_count INTEGER, total_amount INTEGER, paid_amount INTEGER,
        project_date TEXT, description TEXT, client_name TEXT, client_phone TEXT,
        work_year INTEGER
    )''')

    cursor.execute('''CREATE TABLE IF NOT EXISTS inventory (
        item_name TEXT PRIMARY KEY, total_count INTEGER DEFAULT 0, used_count INTEGER DEFAULT 0
    )''')

    for item, _, _, _ in default_items:
        cursor.execute("INSERT OR IGNORE INTO inventory (item_name, total_count, used_count) VALUES (?, 10, 0)", (item,))

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS checks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            check_number TEXT, bank_name TEXT, amount INTEGER,
            due_date TEXT, is_passed INTEGER DEFAULT 0, description TEXT,
            issuer_name TEXT, check_type TEXT DEFAULT 'دریافتی',
            work_year INTEGER
        )
    ''')

    conn.commit()
    conn.close()


# ============================================================
# ======================  LoadingScreen  =====================
# ============================================================
class LoadingScreen(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("IMART STUDIO - در حال بارگذاری...")
        self.setFixedSize(550, 340)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

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

        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(30)
        shadow.setColor(QColor(0, 0, 0, 180))
        shadow.setOffset(0, 6)
        self.container.setGraphicsEffect(shadow)

        layout = QVBoxLayout(self.container)
        layout.setContentsMargins(35, 30, 35, 30)
        layout.setSpacing(15)

        self.lbl_logo = QLabel("🎬 IMART STUDIO")
        self.lbl_logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_logo.setStyleSheet("color: white; font-size: 26pt; font-weight: bold; background: transparent;")
        layout.addWidget(self.lbl_logo)

        self.lbl_sub = QLabel(f"سیستم جامع مدیریت مالی و حسابداری - نسخه {APP_VERSION}")
        self.lbl_sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_sub.setStyleSheet("color: #aed6f1; font-size: 11pt; background: transparent;")
        layout.addWidget(self.lbl_sub)

        layout.addSpacing(10)

        self.lbl_percent = QLabel("0%")
        self.lbl_percent.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_percent.setStyleSheet("color: #f1c40f; font-size: 32pt; font-weight: bold; background: transparent;")
        layout.addWidget(self.lbl_percent)

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

        self.lbl_status = QLabel("در حال آماده‌سازی...")
        self.lbl_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_status.setStyleSheet("color: white; font-size: 11pt; background: transparent;")
        layout.addWidget(self.lbl_status)

        layout.addStretch()

        self.lbl_footer = QLabel("📞 09173736618")
        self.lbl_footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_footer.setStyleSheet("color: #bdc3c7; font-size: 9pt; background: transparent;")
        layout.addWidget(self.lbl_footer)

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


# ============================================================
# =============  ManageItemsDialog (پکیج‌ها)  ================
# ============================================================
class ManageItemsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("مدیریت پکیج‌ها، زیرمجموعه‌ها و کدهای کالا")
        self.resize(750, 620)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.setFont(QFont(APP_FONT_FAMILY, 10))

        layout = QVBoxLayout()
        form_layout = QFormLayout()

        self.txt_item_code = QLineEdit()
        self.txt_item_name = QLineEdit()
        self.txt_item_price = QLineEdit()
        self.txt_item_price.textChanged.connect(lambda t: self.txt_item_price.setText(format_number(t)))

        self.chk_is_package = QCheckBox("این مورد یک پکیج اصلی است")
        self.chk_is_package.stateChanged.connect(self.toggle_package_mode)

        self.combo_parent_package = QComboBox()
        self.load_packages_combo()

        form_layout.addRow("کد کالا / خدمات:", self.txt_item_code)
        form_layout.addRow("نام مورد / پکیج:", self.txt_item_name)
        form_layout.addRow("قیمت (تومان):", self.txt_item_price)
        form_layout.addRow("", self.chk_is_package)
        form_layout.addRow("پکیج مادر (برای زیرمجموعه):", self.combo_parent_package)

        btn_row = QHBoxLayout()
        btn_add = QPushButton("افزودن / بروزرسانی")
        btn_add.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; padding: 6px;")
        btn_add.clicked.connect(self.add_or_update_item)

        btn_auto_code = QPushButton("تولید خودکار کد")
        btn_auto_code.setStyleSheet("background-color: #2980b9; color: white; padding: 6px;")
        btn_auto_code.clicked.connect(self.generate_auto_code)

        btn_row.addWidget(btn_add)
        btn_row.addWidget(btn_auto_code)
        form_layout.addRow(btn_row)
        layout.addLayout(form_layout)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabels(["کد", "نام آیتم / زیرمجموعه", "قیمت (تومان)", "نوع"])
        self.tree.setColumnWidth(0, 100)
        self.tree.setColumnWidth(2, 150)
        layout.addWidget(self.tree)

        bottom_row = QHBoxLayout()
        btn_delete = QPushButton("حذف مورد انتخاب‌شده")
        btn_delete.setStyleSheet("background-color: #c0392b; color: white; font-weight: bold; padding: 6px;")
        btn_delete.clicked.connect(self.delete_item)

        btn_edit = QPushButton("ویرایش (تغییر قیمت/نام)")
        btn_edit.setStyleSheet("background-color: #f39c12; color: white; font-weight: bold; padding: 6px;")
        btn_edit.clicked.connect(self.edit_selected_item)

        bottom_row.addWidget(btn_edit)
        bottom_row.addWidget(btn_delete)
        layout.addLayout(bottom_row)

        self.load_tree_items()
        self.setLayout(layout)

    def generate_auto_code(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM item_prices")
        count = cursor.fetchone()[0]
        conn.close()
        self.txt_item_code.setText(f"SRV-{1000 + count + 1}")

    def toggle_package_mode(self, state):
        self.combo_parent_package.setEnabled(not self.chk_is_package.isChecked())
        if self.chk_is_package.isChecked():
            self.combo_parent_package.setCurrentIndex(0)

    def load_packages_combo(self):
        self.combo_parent_package.clear()
        self.combo_parent_package.addItem("--- بدون پکیج مادر ---", "")
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
        cursor.execute("SELECT item_name, code, price, is_package, parent_package FROM item_prices")
        rows = cursor.fetchall()
        conn.close()

        packages = [r for r in rows if r[3] == 1]
        singles = [r for r in rows if r[3] == 0]

        for pkg in packages:
            pkg_item = QTreeWidgetItem(self.tree, [
                pkg[1] or "-", pkg[0], f"{pkg[2]:,}", "پکیج اصلی"
            ])
            pkg_item.setFont(1, QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
            pkg_item.setForeground(1, QColor("#c0392b"))

            sub_items = [s for s in singles if s[4] == pkg[0]]
            for sub in sub_items:
                QTreeWidgetItem(pkg_item, [
                    sub[1] or "-", f"  └ {sub[0]}", f"{sub[2]:,}", "زیرمجموعه"
                ])

        independent = [s for s in singles if not s[4]]
        for ind in independent:
            QTreeWidgetItem(self.tree, [
                ind[1] or "-", ind[0], f"{ind[2]:,}", "مستقل"
            ])

        self.tree.expandAll()

    def add_or_update_item(self):
        if not ask_security_password(self):
            return

        code = self.txt_item_code.text().strip()
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
            INSERT INTO item_prices (item_name, code, price, is_package, parent_package)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(item_name) DO UPDATE SET
                code=excluded.code, price=excluded.price,
                is_package=excluded.is_package, parent_package=excluded.parent_package
        ''', (name, code, price, is_pkg, parent))
        cursor.execute("INSERT OR IGNORE INTO inventory (item_name, total_count, used_count) VALUES (?, 10, 0)", (name,))
        conn.commit()
        conn.close()

        self.txt_item_code.clear()
        self.txt_item_name.clear()
        self.txt_item_price.clear()
        self.chk_is_package.setChecked(False)
        self.load_packages_combo()
        self.load_tree_items()

    def edit_selected_item(self):
        selected = self.tree.currentItem()
        if not selected:
            QMessageBox.warning(self, "خطا", "لطفاً یک آیتم را انتخاب کنید.")
            return

        code = selected.text(0)
        name = selected.text(1).replace("  └ ", "").strip()
        price_str = selected.text(2).replace(",", "").strip()

        if not ask_security_password(self):
            return

        new_code, ok1 = QInputDialog.getText(self, "ویرایش کد", "کد کالا/خدمات:", text=code if code != "-" else "")
        if not ok1:
            return
        new_name, ok2 = QInputDialog.getText(self, "ویرایش نام", "نام آیتم:", text=name)
        if not ok2:
            return
        new_price, ok3 = QInputDialog.getText(self, "ویرایش قیمت", "قیمت (تومان):", text=price_str)
        if not ok3:
            return

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("UPDATE item_prices SET code=?, item_name=?, price=? WHERE item_name=?",
                       (new_code, new_name, parse_number(new_price), name))
        cursor.execute("UPDATE inventory SET item_name=? WHERE item_name=?", (new_name, name))
        conn.commit()
        conn.close()

        self.load_packages_combo()
        self.load_tree_items()

    def delete_item(self):
        selected = self.tree.currentItem()
        if not selected:
            QMessageBox.warning(self, "خطا", "لطفاً یک آیتم را انتخاب کنید.")
            return

        item_name = selected.text(1).replace("  └ ", "").strip()
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


# ============================================================
# ==================  DepositsDialog  ========================
# ============================================================
class DepositsDialog(QDialog):
    def __init__(self, contract_id, parent=None):
        super().__init__(parent)
        self.contract_id = contract_id
        self.setWindowTitle(f"مدیریت بیعانه‌ها و پرداخت‌های قرارداد #{contract_id}")
        self.resize(600, 400)
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
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["ردیف", "مبلغ (تومان)", "تاریخ", "بانک", "حذف"])
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

            btn_del = QPushButton("حذف")
            btn_del.setStyleSheet("background-color: #e74c3c; color: white;")
            btn_del.clicked.connect(lambda _, did=d_id: self.delete_deposit(did))
            self.table.setCellWidget(r_idx, 4, btn_del)

    def delete_deposit(self, d_id):
        if not ask_security_password(self):
            return
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM wedding_deposits WHERE id=?", (d_id,))
        cursor.execute("SELECT SUM(amount) FROM wedding_deposits WHERE contract_id=?", (self.contract_id,))
        total_paid = cursor.fetchone()[0] or 0
        cursor.execute("SELECT total_amount, discount FROM wedding_contracts WHERE id=?", (self.contract_id,))
        result = cursor.fetchone()
        if result:
            tot, disc = result
            is_settled = 1 if (tot - disc - total_paid) <= 0 else 0
            cursor.execute("UPDATE wedding_contracts SET paid_amount=?, is_settled=? WHERE id=?",
                           (total_paid, is_settled, self.contract_id))
        conn.commit()
        conn.close()
        self.load_deposits()

    def add_deposit(self):
        amount = parse_number(self.txt_amount.text())
        if amount <= 0:
            QMessageBox.warning(self, "خطا", "لطفاً مبلغ معتبر وارد کنید.")
            return

        d_date = self.txt_date.text().strip()
        bank = self.combo_bank.currentText()

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO wedding_deposits (contract_id, amount, deposit_date, bank_name, work_year) VALUES (?, ?, ?, ?, ?)",
                       (self.contract_id, amount, d_date, bank, get_current_year()))
        cursor.execute("SELECT SUM(amount) FROM wedding_deposits WHERE contract_id=?", (self.contract_id,))
        total_paid = cursor.fetchone()[0] or 0
        cursor.execute("SELECT total_amount, discount FROM wedding_contracts WHERE id=?", (self.contract_id,))
        result = cursor.fetchone()
        if result:
            tot, disc = result
            is_settled = 1 if (tot - disc - total_paid) <= 0 else 0
            cursor.execute("UPDATE wedding_contracts SET paid_amount=?, is_settled=? WHERE id=?",
                           (total_paid, is_settled, self.contract_id))

        conn.commit()
        conn.close()

        self.txt_amount.clear()
        self.load_deposits()
        QMessageBox.information(self, "موفقیت", "بیعانه با موفقیت ثبت شد.")
        # ============================================================
# ================  ContractDetailsDialog  ==================
# ============================================================
class ContractDetailsDialog(QDialog):
    def __init__(self, contract_id, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"جزئیات کامل قرارداد #{contract_id}")
        self.resize(700, 550)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.setFont(QFont(APP_FONT_FAMILY, 10))

        layout = QVBoxLayout()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT groom_name, bride_name, groom_phone, bride_phone, contract_date, ceremony_date, selected_items, total_amount, discount, paid_amount, description FROM wedding_contracts WHERE id=?", (contract_id,))
        row = cursor.fetchone()
        conn.close()

        if row:
            info_text = f"<b>زوجین:</b> {row[0]} و {row[1]}<br>"
            info_text += f"<b>تلفن داماد:</b> {row[2] or '-'} | <b>تلفن عروس:</b> {row[3] or '-'}<br>"
            info_text += f"<b>تاریخ قرارداد:</b> {row[4]} | <b>تاریخ مراسم:</b> {row[5]}<br>"
            info_text += f"<b>توضیحات:</b> {row[10] if row[10] else 'ندارد'}"
            lbl_info = QLabel(info_text)
            lbl_info.setStyleSheet("background-color: #f8f9fa; padding: 12px; border-radius: 5px; font-size: 11pt;")
            lbl_info.setWordWrap(True)
            layout.addWidget(lbl_info)

            lbl_items = QLabel("<b>مفاد و خدمات انتخاب شده:</b>")
            layout.addWidget(lbl_items)

            table = QTableWidget()
            table.setColumnCount(2)
            table.setHorizontalHeaderLabels(["ردیف", "عنوان خدمت / پکیج"])
            table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
            table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)

            items = row[6].split(',') if row[6] else []
            table.setRowCount(len(items))
            for idx, item in enumerate(items):
                table.setItem(idx, 0, QTableWidgetItem(str(idx + 1)))
                table.setItem(idx, 1, QTableWidgetItem(item))
            layout.addWidget(table)

            remain = row[7] - row[9]
            summary = (
                f"<b>جمع کل:</b> {row[7]:,} تومان &nbsp;|&nbsp; "
                f"<b>تخفیف:</b> {row[8]:,} تومان &nbsp;|&nbsp; "
                f"<b>پرداختی:</b> {row[9]:,} تومان &nbsp;|&nbsp; "
                f"<b style='color:#c0392b;'>مانده:</b> {remain:,} تومان"
            )
            lbl_sum = QLabel(summary)
            lbl_sum.setStyleSheet("font-size: 11pt; background-color: #fff9e6; padding: 10px; border-radius: 5px; border: 1px solid #f39c12;")
            layout.addWidget(lbl_sum)

        self.setLayout(layout)


# ============================================================
# ================  EditContractDialog  =====================
# ============================================================
class EditContractDialog(QDialog):
    def __init__(self, contract_id, parent=None):
        super().__init__(parent)
        self.contract_id = contract_id
        self.setWindowTitle(f"ویرایش قرارداد #{contract_id}")
        self.resize(650, 700)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.setFont(QFont(APP_FONT_FAMILY, 10))

        layout = QVBoxLayout()
        form_layout = QFormLayout()

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("""SELECT groom_name, bride_name, groom_phone, bride_phone,
                          contract_date, ceremony_date, selected_items, total_amount,
                          discount, description FROM wedding_contracts WHERE id=?""", (contract_id,))
        row = cursor.fetchone()

        cursor.execute("SELECT item_name, code, price FROM item_prices")
        all_items = cursor.fetchall()
        conn.close()

        self.groom = QLineEdit(row[0] if row else "")
        self.bride = QLineEdit(row[1] if row else "")
        self.groom_phone = QLineEdit(row[2] if row else "")
        self.bride_phone = QLineEdit(row[3] if row else "")
        self.contract_date = QLineEdit(row[4] if row else "")
        self.ceremony_date = QLineEdit(row[5] if row else "")
        self.discount = QLineEdit(f"{row[8]:,}" if row else "0")
        self.discount.textChanged.connect(lambda t: self.discount.setText(format_number(t)))
        self.desc = QTextEdit(row[9] if row else "")
        self.desc.setFixedHeight(70)

        form_layout.addRow("نام داماد:", self.groom)
        form_layout.addRow("نام عروس:", self.bride)
        form_layout.addRow("تلفن داماد:", self.groom_phone)
        form_layout.addRow("تلفن عروس:", self.bride_phone)
        form_layout.addRow("تاریخ قرارداد:", self.contract_date)
        form_layout.addRow("تاریخ مراسم:", self.ceremony_date)
        form_layout.addRow("تخفیف (تومان):", self.discount)

        selected = row[6].split(',') if row and row[6] else []
        items_box = QGroupBox("مفاد فاکتور (تیک بزنید)")
        items_box_layout = QVBoxLayout()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFixedHeight(200)

        inner = QWidget()
        inner_layout = QVBoxLayout()
        self.item_checks = {}
        self.item_prices = {}
        for item_name, code, price in all_items:
            row_w = QWidget()
            row_l = QHBoxLayout()
            row_l.setContentsMargins(0, 0, 0, 0)
            cb = QCheckBox(f"[{code or '-'}] {item_name}")
            cb.setChecked(item_name in selected)
            txt = QLineEdit(f"{price:,}")
            txt.setFixedWidth(110)
            txt.textChanged.connect(lambda t, p=txt: p.setText(format_number(t)))
            self.item_checks[item_name] = cb
            self.item_prices[item_name] = txt
            row_l.addWidget(cb)
            row_l.addStretch()
            row_l.addWidget(QLabel("قیمت:"))
            row_l.addWidget(txt)
            row_w.setLayout(row_l)
            inner_layout.addWidget(row_w)
        inner.setLayout(inner_layout)
        scroll.setWidget(inner)
        items_box_layout.addWidget(scroll)
        items_box.setLayout(items_box_layout)
        form_layout.addRow(items_box)

        form_layout.addRow("توضیحات:", self.desc)
        layout.addLayout(form_layout)

        btn_row = QHBoxLayout()
        btn_save = QPushButton("💾 ذخیره تغییرات")
        btn_save.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; padding: 8px;")
        btn_save.clicked.connect(self.save_changes)

        btn_cancel = QPushButton("لغو")
        btn_cancel.clicked.connect(self.reject)

        btn_row.addWidget(btn_save)
        btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row)

        self.setLayout(layout)

    def save_changes(self):
        if not ask_security_password(self):
            return

        selected_items = []
        total = 0
        for item_name, cb in self.item_checks.items():
            if cb.isChecked():
                selected_items.append(item_name)
                total += parse_number(self.item_prices[item_name].text())

        discount = parse_number(self.discount.text())
        final_total = max(0, total - discount)

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("""UPDATE wedding_contracts SET
            groom_name=?, bride_name=?, groom_phone=?, bride_phone=?,
            contract_date=?, ceremony_date=?, selected_items=?,
            total_amount=?, discount=?, description=?
            WHERE id=?""",
            (self.groom.text(), self.bride.text(), self.groom_phone.text(),
             self.bride_phone.text(), self.contract_date.text(), self.ceremony_date.text(),
             ",".join(selected_items), final_total, discount,
             self.desc.toPlainText(), self.contract_id))
        conn.commit()
        conn.close()

        QMessageBox.information(self, "موفقیت", "تغییرات با موفقیت ذخیره شد.")
        self.accept()


# ============================================================
# =============  EditCommercialDialog  ======================
# ============================================================
class EditCommercialDialog(QDialog):
    def __init__(self, project_id, parent=None):
        super().__init__(parent)
        self.project_id = project_id
        self.setWindowTitle(f"ویرایش پروژه #{project_id}")
        self.resize(500, 500)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.setFont(QFont(APP_FONT_FAMILY, 10))

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("""SELECT title, project_type, camera_count, total_amount,
                          paid_amount, project_date, description, client_name, client_phone
                          FROM commercial_projects WHERE id=?""", (project_id,))
        row = cursor.fetchone()
        conn.close()

        layout = QVBoxLayout()
        form = QFormLayout()

        self.title = QLineEdit(row[0] if row else "")
        self.ptype = QComboBox()
        self.ptype.addItems(["تبلیغاتی", "بیوتی", "تولدی", "عروسی", "عقد"])
        if row:
            self.ptype.setCurrentText(row[1])
        self.cameras = QSpinBox()
        self.cameras.setRange(1, 20)
        self.cameras.setValue(row[2] if row else 1)
        self.amount = QLineEdit(f"{row[3]:,}" if row else "0")
        self.amount.textChanged.connect(lambda t: self.amount.setText(format_number(t)))
        self.paid = QLineEdit(f"{row[4]:,}" if row else "0")
        self.paid.textChanged.connect(lambda t: self.paid.setText(format_number(t)))
        self.date = QLineEdit(row[5] if row else jdatetime.date.today().strftime("%Y/%m/%d"))
        self.client_name = QLineEdit(row[7] if row else "")
        self.client_phone = QLineEdit(row[8] if row else "")
        self.desc = QTextEdit(row[6] if row else "")
        self.desc.setFixedHeight(80)

        form.addRow("عنوان پروژه:", self.title)
        form.addRow("نوع پروژه:", self.ptype)
        form.addRow("تعداد دوربین:", self.cameras)
        form.addRow("مبلغ کل:", self.amount)
        form.addRow("پرداختی:", self.paid)
        form.addRow("تاریخ:", self.date)
        form.addRow("نام مشتری:", self.client_name)
        form.addRow("تلفن مشتری:", self.client_phone)
        form.addRow("توضیحات:", self.desc)
        layout.addLayout(form)

        btn_row = QHBoxLayout()
        btn_save = QPushButton("💾 ذخیره تغییرات")
        btn_save.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; padding: 8px;")
        btn_save.clicked.connect(self.save_changes)
        btn_cancel = QPushButton("لغو")
        btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_save)
        btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row)

        self.setLayout(layout)

    def save_changes(self):
        if not ask_security_password(self):
            return

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("""UPDATE commercial_projects SET
            title=?, project_type=?, camera_count=?, total_amount=?,
            paid_amount=?, project_date=?, description=?,
            client_name=?, client_phone=?
            WHERE id=?""",
            (self.title.text(), self.ptype.currentText(), self.cameras.value(),
             parse_number(self.amount.text()), parse_number(self.paid.text()),
             self.date.text(), self.desc.toPlainText(),
             self.client_name.text(), self.client_phone.text(), self.project_id))
        conn.commit()
        conn.close()

        QMessageBox.information(self, "موفقیت", "تغییرات ذخیره شد.")
        self.accept()


# ============================================================
# ====================  InvoiceBuilder  =====================
# ============================================================
class InvoiceBuilder:
    """سازنده فاکتور گرافیکی شبیه اکسل آفیس باز"""

    @staticmethod
    def build_wedding_invoice(contract_id, with_details=True):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("""SELECT id, code, groom_name, bride_name, groom_phone, bride_phone,
                          contract_date, ceremony_date, selected_items, total_amount,
                          discount, paid_amount, description
                          FROM wedding_contracts WHERE id=?""", (contract_id,))
        c = cursor.fetchone()

        cursor.execute("SELECT amount, deposit_date, bank_name FROM wedding_deposits WHERE contract_id=? ORDER BY id", (contract_id,))
        deposits = cursor.fetchall()

        cursor.execute("SELECT check_number, bank_name, amount, due_date, is_passed, issuer_name, check_type FROM checks WHERE work_year=?", (get_current_year(),))
        checks = cursor.fetchall()
        conn.close()

        if not c:
            return "", ""

        studio_name = get_setting("studio_name", "IMART STUDIO")
        studio_phone = get_setting("studio_phone", "09173736618")
        now = jdatetime.datetime.now()
        date_str = now.strftime("%Y/%m/%d")
        time_str = now.strftime("%H:%M")
        invoice_code = f"INV-{c[0]:05d}"

        items_list = c[8].split(',') if c[8] else []

        # جمع‌آوری اطلاعات قیمت هر آیتم
        items_rows = ""
        total_before_discount = 0
        for idx, item in enumerate(items_list):
            conn2 = sqlite3.connect(DB_NAME)
            cur2 = conn2.cursor()
            cur2.execute("SELECT code, price FROM item_prices WHERE item_name=?", (item,))
            found = cur2.fetchone()
            conn2.close()
            code = found[0] if found else "-"
            price = found[1] if found else 0
            total_before_discount += price
            items_rows += f"""
            <tr>
                <td style='border:1px solid #000; padding:6px; text-align:center;'>{idx+1}</td>
                <td style='border:1px solid #000; padding:6px; text-align:center;'>{code}</td>
                <td style='border:1px solid #000; padding:6px;'>{item}</td>
                <td style='border:1px solid #000; padding:6px; text-align:center;'>1</td>
                <td style='border:1px solid #000; padding:6px; text-align:center;'>{price:,}</td>
                <td style='border:1px solid #000; padding:6px; text-align:center;'>{price:,}</td>
                <td style='border:1px solid #000; padding:6px; text-align:center;'>-</td>
            </tr>"""

        deposits_rows = ""
        for idx, (amt, dt, bank) in enumerate(deposits):
            deposits_rows += f"""
            <tr>
                <td style='border:1px solid #000; padding:5px; text-align:center;'>{idx+1}</td>
                <td style='border:1px solid #000; padding:5px; text-align:center;'>{amt:,}</td>
                <td style='border:1px solid #000; padding:5px; text-align:center;'>{dt}</td>
                <td style='border:1px solid #000; padding:5px; text-align:center;'>{bank or '-'}</td>
            </tr>"""

        checks_rows = ""
        for chk in checks:
            status = "✅ پاس شده" if chk[4] else "⏳ پاس نشده"
            checks_rows += f"""
            <tr>
                <td style='border:1px solid #000; padding:5px; text-align:center;'>{chk[0]}</td>
                <td style='border:1px solid #000; padding:5px; text-align:center;'>{chk[1] or '-'}</td>
                <td style='border:1px solid #000; padding:5px; text-align:center;'>{chk[2]:,}</td>
                <td style='border:1px solid #000; padding:5px; text-align:center;'>{chk[3]}</td>
                <td style='border:1px solid #000; padding:5px; text-align:center;'>{chk[5] or '-'}</td>
                <td style='border:1px solid #000; padding:5px; text-align:center;'>{chk[6] or 'دریافتی'}</td>
                <td style='border:1px solid #000; padding:5px; text-align:center;'>{status}</td>
            </tr>"""

        remain = c[9] - c[10] - c[11]
        final_amount = c[9] - c[10]
        amount_words = number_to_persian_words(max(0, remain))

        html = f"""
        <div dir="rtl" style="font-family: '{INVOICE_FONT_FAMILY}', Tahoma; padding: 15px;">
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 12px;">
                <tr>
                    <td style="text-align: center; padding: 5px;">
                        <h1 style="margin: 0; font-size: 28pt; color: #1F4E78;">{studio_name}</h1>
                        <h3 style="margin: 3px 0; font-size: 14pt; color: #333;">فاکتور فروش</h3>
                    </td>
                    <td style="text-align: left; padding: 5px; width: 30%; vertical-align: top;">
                        <div style="font-size: 10pt;">تاریخ: {date_str}</div>
                        <div style="font-size: 10pt;">ساعت: {time_str}</div>
                        <div style="font-size: 10pt;">کد فاکتور: {invoice_code}</div>
                    </td>
                </tr>
            </table>
            <hr style="border: 2px solid #1F4E78; margin: 5px 0 15px 0;">

            <table style="width: 100%; margin-bottom: 12px; border-collapse: collapse;">
                <tr>
                    <td style="padding: 5px; font-size: 11pt;"><b>صورتحساب آقای / خانم:</b> {c[2]} و {c[3]}</td>
                    <td style="padding: 5px; font-size: 11pt;"><b>تلفن همراه:</b> {c[4] or '-'} / {c[5] or '-'}</td>
                </tr>
                <tr>
                    <td style="padding: 5px; font-size: 11pt;"><b>تاریخ قرارداد:</b> {c[6]}</td>
                    <td style="padding: 5px; font-size: 11pt;"><b>تاریخ مراسم:</b> {c[7]}</td>
                </tr>
            </table>

            <table style="width: 100%; border-collapse: collapse; margin-bottom: 12px; font-size: 10pt;">
                <thead>
                    <tr style="background-color: #D9E1F2;">
                        <th style="border:1px solid #000; padding:6px; width: 5%;">ردیف</th>
                        <th style="border:1px solid #000; padding:6px; width: 10%;">کد کالا</th>
                        <th style="border:1px solid #000; padding:6px;">نام خدمات</th>
                        <th style="border:1px solid #000; padding:6px; width: 7%;">تعداد</th>
                        <th style="border:1px solid #000; padding:6px; width: 12%;">بهای واحد</th>
                        <th style="border:1px solid #000; padding:6px; width: 13%;">مبلغ کل</th>
                        <th style="border:1px solid #000; padding:6px; width: 15%;">شرح خدمات</th>
                    </tr>
                </thead>
                <tbody>
                    {items_rows}
                </tbody>
            </table>

            <table style="width: 100%; border-collapse: collapse; margin-bottom: 12px; font-size: 10pt;">
                <tr>
                    <td style="border:1px solid #000; padding:8px; width: 70%; text-align: right;">
                        <b>مبلغ به حروف:</b> {amount_words} تومان
                    </td>
                    <td style="border:1px solid #000; padding:8px; background-color: #F2F2F2; width: 30%;">
                        <table style="width:100%; font-size:10pt;">
                            <tr><td>جمع کل:</td><td style="text-align:left;">{c[9]:,}</td></tr>
                            <tr><td>تخفیف:</td><td style="text-align:left;">{c[10]:,}</td></tr>
                            <tr><td><b>جمع کل فاکتور:</b></td><td style="text-align:left;"><b>{final_amount:,}</b></td></tr>
                            <tr><td>بیعانه‌ها:</td><td style="text-align:left;">{c[11]:,}</td></tr>
                            <tr style="background-color:#FFE699;"><td><b>باقی‌مانده کل:</b></td><td style="text-align:left;"><b>{remain:,}</b></td></tr>
                        </table>
                    </td>
                </tr>
            </table>

            <h4 style="margin: 10px 0 5px 0;">بیعانه‌های پرداختی:</h4>
            <table style="width: 100%; border-collapse: collapse; font-size: 10pt; margin-bottom: 12px;">
                <thead>
                    <tr style="background-color: #D9E1F2;">
                        <th style="border:1px solid #000; padding:5px; width: 8%;">ردیف</th>
                        <th style="border:1px solid #000; padding:5px;">مبلغ</th>
                        <th style="border:1px solid #000; padding:5px;">تاریخ</th>
                        <th style="border:1px solid #000; padding:5px;">بانک</th>
                    </tr>
                </thead>
                <tbody>
                    {deposits_rows if deposits_rows else "<tr><td colspan='4' style='border:1px solid #000; padding:5px; text-align:center;'>بیعانه‌ای ثبت نشده</td></tr>"}
                </tbody>
            </table>

            <h4 style="margin: 10px 0 5px 0;">لیست چک‌های پاس شده و پاس نشده:</h4>
            <table style="width: 100%; border-collapse: collapse; font-size: 9pt; margin-bottom: 12px;">
                <thead>
                    <tr style="background-color: #D9E1F2;">
                        <th style="border:1px solid #000; padding:5px;">شماره چک</th>
                        <th style="border:1px solid #000; padding:5px;">بانک</th>
                        <th style="border:1px solid #000; padding:5px;">مبلغ</th>
                        <th style="border:1px solid #000; padding:5px;">سررسید</th>
                        <th style='border:1px solid #000; padding:5px;'>صادرکننده</th>
                        <th style='border:1px solid #000; padding:5px;'>نوع</th>
                        <th style='border:1px solid #000; padding:5px;'>وضعیت</th>
                    </tr>
                </thead>
                <tbody>
                    {checks_rows if checks_rows else "<tr><td colspan='7' style='border:1px solid #000; padding:5px; text-align:center;'>چکی ثبت نشده</td></tr>"}
                </tbody>
            </table>

            <p style="font-size: 10pt;"><b>توضیحات:</b> {c[12] if c[12] else 'ندارد'}</p>

            <table style="width: 100%; text-align: center; margin-top: 30px; font-size: 11pt;">
                <tr>
                    <td style="width: 50%; padding: 15px;">مهر و امضای مشتری</td>
                    <td style="width: 50%; padding: 15px;">مهر و امضای استودیو</td>
                </tr>
            </table>

            <p style="text-align: center; margin-top: 20px; font-size: 9pt; color: #666;">
                📞 {studio_phone} | {studio_name}
            </p>
        </div>
        """
        filename = f"فاکتور_{c[2]}_{c[3]}_{invoice_code}.pdf"
        return html, filename

    @staticmethod
    def build_commercial_invoice(project_id):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("""SELECT id, code, title, project_type, camera_count, total_amount,
                          paid_amount, project_date, description, client_name, client_phone
                          FROM commercial_projects WHERE id=?""", (project_id,))
        c = cursor.fetchone()
        conn.close()
        if not c:
            return "", ""

        studio_name = get_setting("studio_name", "IMART STUDIO")
        studio_phone = get_setting("studio_phone", "09173736618")
        now = jdatetime.datetime.now()
        invoice_code = f"CIV-{c[0]:05d}"
        remain = c[5] - c[6]
        amount_words = number_to_persian_words(max(0, remain))

        html = f"""
        <div dir="rtl" style="font-family: '{INVOICE_FONT_FAMILY}', Tahoma; padding: 15px;">
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 12px;">
                <tr>
                    <td style="text-align: center; padding: 5px;">
                        <h1 style="margin: 0; font-size: 28pt; color: #1F4E78;">{studio_name}</h1>
                        <h3 style="margin: 3px 0; font-size: 14pt; color: #333;">فاکتور پروژه {c[3]}</h3>
                    </td>
                    <td style="text-align: left; padding: 5px; width: 30%; vertical-align: top;">
                        <div style="font-size: 10pt;">تاریخ: {now.strftime('%Y/%m/%d')}</div>
                        <div style="font-size: 10pt;">ساعت: {now.strftime('%H:%M')}</div>
                        <div style="font-size: 10pt;">کد فاکتور: {invoice_code}</div>
                    </td>
                </tr>
            </table>
            <hr style="border: 2px solid #1F4E78; margin: 5px 0 15px 0;">

            <table style="width: 100%; margin-bottom: 12px; border-collapse: collapse;">
                <tr>
                    <td style="padding: 5px; font-size: 11pt;"><b>صورتحساب آقای / خانم:</b> {c[9] or '-'}</td>
                    <td style="padding: 5px; font-size: 11pt;"><b>تلفن همراه:</b> {c[10] or '-'}</td>
                </tr>
                <tr>
                    <td style="padding: 5px; font-size: 11pt;"><b>تاریخ پروژه:</b> {c[7]}</td>
                    <td style="padding: 5px; font-size: 11pt;"><b>تعداد دوربین:</b> {c[4]}</td>
                </tr>
            </table>

            <table style="width: 100%; border-collapse: collapse; margin-bottom: 12px; font-size: 10pt;">
                <thead>
                    <tr style="background-color: #D9E1F2;">
                        <th style="border:1px solid #000; padding:6px;">ردیف</th>
                        <th style="border:1px solid #000; padding:6px;">عنوان خدمات</th>
                        <th style="border:1px solid #000; padding:6px;">تعداد</th>
                        <th style="border:1px solid #000; padding:6px;">مبلغ کل</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td style="border:1px solid #000; padding:6px; text-align:center;">1</td>
                        <td style="border:1px solid #000; padding:6px;">{c[2]}</td>
                        <td style="border:1px solid #000; padding:6px; text-align:center;">1</td>
                        <td style="border:1px solid #000; padding:6px; text-align:center;">{c[5]:,}</td>
                    </tr>
                </tbody>
            </table>

            <table style="width: 100%; border-collapse: collapse; margin-bottom: 12px; font-size: 10pt;">
                <tr>
                    <td style="border:1px solid #000; padding:8px; text-align: right;">
                        <b>مبلغ به حروف:</b> {amount_words} تومان
                    </td>
                    <td style="border:1px solid #000; padding:8px; background-color:#F2F2F2;">
                        <table style="width:100%; font-size:10pt;">
                            <tr><td>جمع کل:</td><td style="text-align:left;">{c[5]:,}</td></tr>
                            <tr><td>پرداختی:</td><td style="text-align:left;">{c[6]:,}</td></tr>
                            <tr style="background-color:#FFE699;"><td><b>باقی‌مانده:</b></td><td style="text-align:left;"><b>{remain:,}</b></td></tr>
                        </table>
                    </td>
                </tr>
            </table>

            <p style="font-size: 10pt;"><b>توضیحات:</b> {c[8] if c[8] else 'ندارد'}</p>

            <table style="width: 100%; text-align: center; margin-top: 30px; font-size: 11pt;">
                <tr>
                    <td style="width: 50%; padding: 15px;">مهر و امضای مشتری</td>
                    <td style="width: 50%; padding: 15px;">مهر و امضای استودیو</td>
                </tr>
            </table>

            <p style="text-align: center; margin-top: 20px; font-size: 9pt; color: #666;">📞 {studio_phone}</p>
        </div>
        """
        filename = f"فاکتور_{c[2]}_{invoice_code}.pdf"
        return html, filename

    @staticmethod
    def build_receipt(name, amount, for_service, date_str):
        studio_name = get_setting("studio_name", "IMART STUDIO")
        studio_phone = get_setting("studio_phone", "09173736618")
        now = jdatetime.datetime.now()
        amount_words = number_to_persian_words(amount)

        html = f"""
        <div dir="rtl" style="font-family: '{INVOICE_FONT_FAMILY}', Tahoma; padding: 25px;">
            <table style="width: 100%; border-collapse: collapse; margin-bottom: 12px;">
                <tr>
                    <td style="text-align: center;">
                        <h1 style="margin: 0; font-size: 30pt; color: #1F4E78;">{studio_name}</h1>
                        <h3 style="margin: 5px 0; font-size: 16pt;">رسید وجه</h3>
                    </td>
                    <td style="text-align: left; width: 30%; vertical-align: top;">
                        <div style="font-size: 11pt;">تاریخ: {date_str}</div>
                        <div style="font-size: 11pt;">ساعت: {now.strftime('%H:%M')}</div>
                    </td>
                </tr>
            </table>
            <hr style="border: 2px solid #1F4E78; margin: 10px 0 25px 0;">

            <p style="font-size: 13pt; line-height: 2.2;">
                مبلغ: <b>{amount:,} تومان</b> ({amount_words} تومان)<br>
                از آقا / خانم: <b>{name}</b><br>
                بابت: <b>{for_service}</b><br>
                به حسابداری تحویل گردید.
            </p>

            <p style="font-size: 11pt; color: #c0392b; margin-top: 25px;">
                <b>توجه:</b> بیعانه پرداختی پس داده نمی‌شود.
            </p>

            <table style="width: 100%; margin-top: 80px; font-size: 12pt;">
                <tr>
                    <td style="text-align: center; width: 50%;">امضای حسابداری</td>
                    <td style="text-align: center; width: 50%;">مهر استودیو</td>
                </tr>
            </table>

            <p style="text-align: center; margin-top: 40px; font-size: 9pt; color: #666;">📞 {studio_phone}</p>
        </div>
        """
        return html

    @staticmethod
    def build_staff_report(person_id, period="ماهانه", year=None, month=None):
        """گزارش ریز کارکرد یک فرد"""
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT code, name, role, phone FROM persons WHERE id=?", (person_id,))
        p = cursor.fetchone()
        if not p:
            conn.close()
            return "", ""

        work_year = year or get_current_year()

        cursor.execute("""SELECT amount, year, month, day, description, category
                          FROM transactions WHERE person_id=? AND work_year=?
                          ORDER BY year DESC, month DESC, day DESC""",
                       (person_id, work_year))
        trans = cursor.fetchall()
        conn.close()

        total_salary = sum(t[0] for t in trans)

        trans_rows = ""
        for idx, t in enumerate(trans):
            trans_rows += f"""
            <tr>
                <td style='border:1px solid #000; padding:5px; text-align:center;'>{idx+1}</td>
                <td style='border:1px solid #000; padding:5px; text-align:center;'>{t[1]}/{t[2]:02d}/{t[3]:02d}</td>
                <td style='border:1px solid #000; padding:5px;'>{t[5]}</td>
                <td style='border:1px solid #000; padding:5px;'>{t[4] or '-'}</td>
                <td style='border:1px solid #000; padding:5px; text-align:center;'>{t[0]:,}</td>
            </tr>"""

        studio_name = get_setting("studio_name", "IMART STUDIO")
        now = jdatetime.datetime.now()

        html = f"""
        <div dir="rtl" style="font-family: '{INVOICE_FONT_FAMILY}', Tahoma; padding: 15px;">
            <table style="width: 100%; border-collapse: collapse;">
                <tr>
                    <td style="text-align: center;">
                        <h1 style="margin: 0; font-size: 26pt; color: #1F4E78;">{studio_name}</h1>
                        <h3 style="margin: 5px 0; font-size: 14pt;">گزارش ریز کارکرد پرسنل</h3>
                    </td>
                    <td style="text-align: left; width: 25%; vertical-align: top;">
                        <div style="font-size: 10pt;">تاریخ: {now.strftime('%Y/%m/%d')}</div>
                        <div style="font-size: 10pt;">دوره: {period}</div>
                        <div style="font-size: 10pt;">سال: {work_year}</div>
                    </td>
                </tr>
            </table>
            <hr style="border: 2px solid #1F4E78; margin: 10px 0 20px 0;">

            <table style="width: 100%; border-collapse: collapse; margin-bottom: 15px;">
                <tr>
                    <td style="padding: 5px; font-size: 12pt;"><b>کد پرسنل:</b> {p[0] or '-'}</td>
                    <td style="padding: 5px; font-size: 12pt;"><b>نام:</b> {p[1]}</td>
                </tr>
                <tr>
                    <td style="padding: 5px; font-size: 12pt;"><b>نقش:</b> {p[2]}</td>
                    <td style="padding: 5px; font-size: 12pt;"><b>تلفن:</b> {p[3] or '-'}</td>
                </tr>
            </table>

            <h4>ریز پرداختی‌ها ({period}):</h4>
            <table style="width: 100%; border-collapse: collapse; font-size: 10pt;">
                <thead>
                    <tr style="background-color: #D9E1F2;">
                        <th style="border:1px solid #000; padding:6px;">ردیف</th>
                        <th style="border:1px solid #000; padding:6px;">تاریخ</th>
                        <th style="border:1px solid #000; padding:6px;">دسته</th>
                        <th style="border:1px solid #000; padding:6px;">شرح</th>
                        <th style="border:1px solid #000; padding:6px;">مبلغ (تومان)</th>
                    </tr>
                </thead>
                <tbody>
                    {trans_rows if trans_rows else "<tr><td colspan='5' style='border:1px solid #000; padding:5px; text-align:center;'>تراکنشی یافت نشد</td></tr>"}
                </tbody>
            </table>

            <table style="width: 100%; margin-top: 15px;">
                <tr>
                    <td style="border:1px solid #000; padding:10px; background-color:#FFE699; font-size: 12pt; text-align: center;">
                        <b>جمع کل پرداختی به این فرد:</b> {total_salary:,} تومان
                    </td>
                </tr>
            </table>
            <p style="text-align: center; font-size: 9pt; color: #666; margin-top: 20px;">
                مهر و امضای مدیر / حسابداری
            </p>
        </div>
        """
        filename = f"ریز_کارکرد_{p[1]}_{period}.pdf"
        return html, filename
        # ============================================================
# ================  StudioAccountingApp  ====================
# ============================================================
class StudioAccountingApp(QMainWindow):
    ROLES = ["تدوینگر", "عکاس", "فیلمبردار", "هلی شات و FPV کار", "اوپراتور کرین"]
    PROJECT_TYPES = ["عروسی", "عقد", "تولد", "تبلیغاتی", "قبض و کرایه"]
    PERSIAN_MONTHS = ["فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
                      "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند"]

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"نرم‌افزار مدیریت مالی و حسابداری IMART STUDIO - نسخه {APP_VERSION}")
        self.resize(1400, 900)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)

        icon_path = resource_path("Accounting.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.setFont(QFont(APP_FONT_FAMILY, 10))

        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        self.dashboard_screen = QWidget()
        self.stack.addWidget(self.dashboard_screen)

        self.main_app_screen = QWidget()
        self.stack.addWidget(self.main_app_screen)

        self.stack.setCurrentWidget(self.dashboard_screen)

    def prompt_login(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM settings WHERE key='app_password'")
        row = cursor.fetchone()
        saved_pass = row[0] if row else "123"
        conn.close()

        entered_pass, ok = QInputDialog.getText(
            None, "ورود به سیستم IMART STUDIO", "لطفاً رمز عبور را وارد کنید:",
            QLineEdit.EchoMode.Password
        )
        if ok and entered_pass == saved_pass:
            return True
        elif ok:
            QMessageBox.critical(None, "خطا", "رمز عبور اشتباه است!")
            return False
        return False

    def check_check_alerts(self):
        """اخطار ۳ روز مانده به سررسید چک‌ها"""
        try:
            today = jdatetime.date.today()
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("""SELECT check_number, bank_name, amount, due_date,
                              issuer_name, check_type
                              FROM checks WHERE is_passed=0""")
            rows = cursor.fetchall()
            conn.close()

            alerts = []
            for chk in rows:
                try:
                    parts = chk[3].split('/')
                    due = jdatetime.date(int(parts[0]), int(parts[1]), int(parts[2]))
                    days_left = (due - today).days
                    if 0 <= days_left <= 3:
                        alerts.append((chk, days_left))
                    elif days_left < 0:
                        alerts.append((chk, days_left))
                except Exception:
                    pass

            if alerts:
                msg = "<h3 style='color:#c0392b;'>⚠️ اخطار سررسید چک‌ها</h3><hr>"
                for chk, days in alerts:
                    if days < 0:
                        status = f"<span style='color:#c0392b;'><b>گذشته ({abs(days)} روز)</b></span>"
                    elif days == 0:
                        status = "<span style='color:#e67e22;'><b>امروز!</b></span>"
                    else:
                        status = f"<span style='color:#f39c12;'><b>{days} روز مانده</b></span>"

                    msg += f"""
                    <p style='font-size:11pt;'>
                        <b>شماره چک:</b> {chk[0]} | <b>بانک:</b> {chk[1] or '-'}<br>
                        <b>مبلغ:</b> {chk[2]:,} تومان | <b>سررسید:</b> {chk[3]} ({status})<br>
                        <b>صادرکننده:</b> {chk[4] or '-'} | <b>نوع:</b> {chk[5] or 'دریافتی'}
                    </p><hr>
                    """
                box = QMessageBox(self)
                box.setWindowTitle("اخطار چک")
                box.setTextFormat(Qt.TextFormat.RichText)
                box.setText(msg)
                box.setIcon(QMessageBox.Icon.Warning)
                box.exec()
        except Exception as e:
            print("Error in check alerts:", e)

    def setup_dashboard_ui(self):
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title = QLabel(f"سیستم جامع مدیریت مالی و حسابداری IMART STUDIO (نسخه {APP_VERSION})")
        title.setFont(QFont(APP_FONT_FAMILY, 16, QFont.Weight.Bold))
        title.setStyleSheet("color: #2c3e50; margin-bottom: 15px;")
        layout.addWidget(title, alignment=Qt.AlignmentFlag.AlignCenter)

        # نوار انتخاب سال کاری
        year_bar = QHBoxLayout()
        year_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_year = QLabel("📅 سال کاری فعلی:")
        lbl_year.setFont(QFont(APP_FONT_FAMILY, 12, QFont.Weight.Bold))
        year_bar.addWidget(lbl_year)

        self.combo_work_year = QComboBox()
        self.combo_work_year.setFont(QFont(APP_FONT_FAMILY, 12, QFont.Weight.Bold))
        self.combo_work_year.setFixedWidth(120)
        self.load_work_years()
        current_year = get_current_year()
        idx = self.combo_work_year.findText(str(current_year))
        if idx >= 0:
            self.combo_work_year.setCurrentIndex(idx)
        self.combo_work_year.currentTextChanged.connect(self.switch_work_year)
        year_bar.addWidget(self.combo_work_year)

        layout.addLayout(year_bar)
        layout.addSpacing(15)

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
            ("صدور رسید وجه\n(خدمات)", 9, "#7f8c8d"),
            ("تنظیمات و سال کاری\n(سیستم)", 10, "#95a5a6"),
            ("درباره برنامه\n(توسعه‌دهنده)", 11, "#34495e")
        ]

        for text, idx, color in buttons:
            btn = QPushButton(text)
            btn.setFixedSize(140, 130)
            btn.setFont(QFont(APP_FONT_FAMILY, 9, QFont.Weight.Bold))
            btn.setStyleSheet(
                f"QPushButton {{ background-color: {color}; color: white; "
                f"border-radius: 10px; padding: 5px; }} "
                f"QPushButton:hover {{ background-color: #34495e; }}"
            )
            btn.clicked.connect(lambda _, i=idx: self.open_tab_index(i))
            grid_layout.addWidget(btn)

        layout.addLayout(grid_layout)
        self.dashboard_screen.setLayout(layout)

    def load_work_years(self):
        """بارگذاری لیست سال‌های کاری - با محافظ"""
        if not hasattr(self, 'combo_work_year'):
            return
        self.combo_work_year.blockSignals(True)
        self.combo_work_year.clear()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        years = set()
        years.add(jdatetime.date.today().year)
        try:
            years.add(get_current_year())
        except Exception:
            pass

        for table in ["wedding_contracts", "commercial_projects", "expenses",
                      "transactions", "checks", "wedding_deposits", "persons"]:
            try:
                cursor.execute(f"SELECT DISTINCT work_year FROM {table} WHERE work_year IS NOT NULL")
                for (y,) in cursor.fetchall():
                    if y:
                        years.add(int(y))
            except Exception:
                pass

        conn.close()

        for y in sorted(years, reverse=True):
            self.combo_work_year.addItem(str(y))

        self.combo_work_year.blockSignals(False)

    def switch_work_year(self, year_str):
        """تغییر سال با محافظ"""
        try:
            year = int(year_str)
        except Exception:
            return
        set_setting("current_work_year", year)

        # ⚠️ فقط اگه پنجره اصلی ساخته شده، رفرش کن
        if hasattr(self, 'tabs') and hasattr(self, 'w_table'):
            try:
                self.load_wedding_contracts()
                self.load_commercial_projects()
                self.load_expenses_table()
                self.load_staff_table()
                self.load_inventory()
                self.load_checks_table()
            except Exception as e:
                print("Refresh error:", e)

        if hasattr(self, 'combo_settings_year'):
            try:
                self.load_settings_years()
            except Exception:
                pass

        if hasattr(self, 'tabs'):
            QMessageBox.information(self, "تغییر سال", f"سال کاری به {year} تغییر یافت.")

    def open_tab_index(self, index):
        if hasattr(self, 'tabs'):
            self.tabs.setCurrentIndex(index)
            self.stack.setCurrentWidget(self.main_app_screen)

    def setup_main_app_ui(self):
        main_layout = QVBoxLayout()

        top_bar = QHBoxLayout()
        btn_dash = QPushButton("🏠 بازگشت به منوی اصلی")
        btn_dash.setStyleSheet("background-color: #34495e; color: white; font-weight: bold; padding: 8px;")
        btn_dash.clicked.connect(lambda: self.stack.setCurrentWidget(self.dashboard_screen))

        btn_edit_titles = QPushButton("✏️ تغییر عنوان‌های پیش‌فرض")
        btn_edit_titles.setStyleSheet("background-color: #e67e22; color: white; font-weight: bold; padding: 8px;")
        btn_edit_titles.clicked.connect(self.edit_default_titles)

        btn_backup = QPushButton("💾 پشتیبان‌گیری")
        btn_backup.clicked.connect(self.backup_db)

        btn_restore = QPushButton("♻️ بازیابی بک‌آپ")
        btn_restore.clicked.connect(self.restore_db)

        btn_change_pass = QPushButton("🔑 تغییر رمز")
        btn_change_pass.clicked.connect(self.change_password)

        top_bar.addWidget(btn_dash)
        top_bar.addWidget(btn_edit_titles)
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
        self.tabs.addTab(self.tab_workyear, "تنظیمات و سال کاری")
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

        footer = QLabel(f"IMART STUDIO - Phone: 09173736618 | نسخه {APP_VERSION} | طراح: {DEVELOPER_NAME}")
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        footer.setStyleSheet("color: #2c3e50; margin-top: 5px;")
        main_layout.addWidget(footer)

        self.main_app_screen.setLayout(main_layout)

    def edit_default_titles(self):
        if not ask_security_password(self):
            return
        dlg = QDialog(self)
        dlg.setWindowTitle("تغییر عنوان‌های پیش‌فرض")
        dlg.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        dlg.resize(500, 400)
        dlg.setFont(QFont(APP_FONT_FAMILY, 10))

        v = QVBoxLayout()
        form = QFormLayout()

        txt_studio = QLineEdit(get_setting("studio_name", "IMART STUDIO"))
        txt_phone = QLineEdit(get_setting("studio_phone", "09173736618"))
        txt_address = QLineEdit(get_setting("studio_address", ""))

        form.addRow("نام استودیو:", txt_studio)
        form.addRow("تلفن:", txt_phone)
        form.addRow("آدرس:", txt_address)

        v.addLayout(form)

        info = QLabel("<b>راهنما:</b> نام استودیو و تلفن در فاکتورها و رسیدها نمایش داده می‌شود.")
        info.setStyleSheet("background-color: #e8f8f5; padding: 10px; border-radius: 5px;")
        info.setWordWrap(True)
        v.addWidget(info)

        def save_titles():
            if not ask_security_password(self):
                return
            set_setting("studio_name", txt_studio.text())
            set_setting("studio_phone", txt_phone.text())
            set_setting("studio_address", txt_address.text())
            QMessageBox.information(dlg, "موفقیت", "عنوان‌ها به‌روزرسانی شد.")
            dlg.accept()

        btn = QPushButton("💾 ذخیره")
        btn.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; padding: 8px;")
        btn.clicked.connect(save_titles)
        v.addWidget(btn)

        dlg.setLayout(v)
        dlg.exec()

    # ============================================================
    # ==============  تب ۱: عروس و داماد  =======================
    # ============================================================
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

        form_layout.addRow("نام داماد <span style='color:red;'>*</span>:", self.w_groom)
        form_layout.addRow("نام عروس <span style='color:red;'>*</span>:", self.w_bride)
        form_layout.addRow("تلفن داماد:", self.w_groom_phone)
        form_layout.addRow("تلفن عروس:", self.w_bride_phone)
        form_layout.addRow("تاریخ قرارداد:", self.w_contract_date)
        form_layout.addRow("تاریخ مراسم:", self.w_ceremony_date)

        btn_manage_items = QPushButton("📦 مدیریت / افزودن پکیج و موارد")
        btn_manage_items.setStyleSheet("background-color: #e67e22; color: white; font-weight: bold; padding: 6px;")
        btn_manage_items.clicked.connect(self.open_manage_items)
        form_layout.addRow(btn_manage_items)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFixedHeight(200)

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
        self.w_lbl_total.setStyleSheet("color: #27ae60; font-size: 12pt;")
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
        self.w_lbl_remain.setFont(QFont(APP_FONT_FAMILY, 11, QFont.Weight.Bold))
        self.w_lbl_remain.setStyleSheet("color: #c0392b; font-size: 12pt;")
        form_layout.addRow(self.w_lbl_remain)

        self.w_desc = QTextEdit()
        self.w_desc.setFixedHeight(60)
        self.w_desc.setPlaceholderText("توضیحات کامل قرارداد...")
        form_layout.addRow("توضیحات قرارداد:", self.w_desc)

        self.item_checkboxes = {}
        self.item_price_inputs = {}
        self.load_item_checkboxes()

        self.btn_save_wedding = QPushButton("💾 ثبت نهایی قرارداد")
        self.btn_save_wedding.setFont(QFont(APP_FONT_FAMILY, 11, QFont.Weight.Bold))
        self.btn_save_wedding.setStyleSheet("background-color: #2980b9; color: white; padding: 10px;")
        self.btn_save_wedding.clicked.connect(self.save_wedding_contract)
        form_layout.addRow(self.btn_save_wedding)

        form_box.setLayout(form_layout)
        layout.addWidget(form_box, 1)

        table_box = QGroupBox("لیست قراردادها (برای جزئیات دوبار کلیک کنید)")
        table_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        table_layout = QVBoxLayout()

        self.w_table = QTableWidget()
        self.w_table.setColumnCount(14)
        self.w_table.setHorizontalHeaderLabels([
            "کد", "زوجین", "تلفن داماد", "تلفن عروس", "تاریخ مراسم",
            "جمع کل", "تخفیف", "دریافتی", "مانده", "وضعیت",
            "بیعانه‌ها", "چاپ / PDF", "ویرایش", "حذف"
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
        cursor.execute("SELECT item_name, code, price, is_package, parent_package FROM item_prices ORDER BY is_package DESC, parent_package, item_name")
        rows = cursor.fetchall()
        conn.close()

        packages = [r for r in rows if r[3] == 1]
        singles = [r for r in rows if r[3] == 0]

        def add_row(item_name, code, price, indent=False):
            row_w = QWidget()
            row_l = QHBoxLayout()
            row_l.setContentsMargins(20 if indent else 0, 0, 0, 0)

            cb = QCheckBox(f"{'└ ' if indent else ''}[{code or '-'}] {item_name}")
            cb.stateChanged.connect(self.calc_wedding_total)
            cb.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold if not indent else QFont.Weight.Normal))

            txt_p = QLineEdit(f"{price:,}")
            txt_p.setFixedWidth(100)
            txt_p.textChanged.connect(lambda t, p=txt_p: p.setText(format_number(t)))
            txt_p.textChanged.connect(self.calc_wedding_total)

            row_l.addWidget(cb)
            row_l.addStretch()
            row_l.addWidget(QLabel("قیمت:"))
            row_l.addWidget(txt_p)
            row_w.setLayout(row_l)

            self.items_vbox.addWidget(row_w)
            self.item_checkboxes[item_name] = cb
            self.item_price_inputs[item_name] = txt_p

        for pkg in packages:
            add_row(pkg[0], pkg[1], pkg[2], indent=False)
            sub_items = [s for s in singles if s[4] == pkg[0]]
            for sub in sub_items:
                add_row(sub[0], sub[1], sub[2], indent=True)

        independent = [s for s in singles if not s[4]]
        for ind in independent:
            add_row(ind[0], ind[1], ind[2], indent=False)

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
            QMessageBox.warning(self, "خطا", "حداقل یک مورد از فاکتور را انتخاب کنید.")
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
        work_year = get_current_year()

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        cursor.execute("SELECT COUNT(*) FROM wedding_contracts")
        count = cursor.fetchone()[0]
        code = f"W-{work_year}-{count + 1:04d}"

        cursor.execute('''
            INSERT INTO wedding_contracts 
            (code, groom_name, bride_name, groom_phone, bride_phone, contract_date,
             ceremony_date, selected_items, total_amount, discount, paid_amount,
             bank_id, is_settled, description, work_year)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (code, groom, bride, self.w_groom_phone.text(), self.w_bride_phone.text(),
              self.w_contract_date.text(), self.w_ceremony_date.text(),
              ",".join(selected_items), total_amount, discount, first_deposit,
              bank_id, is_settled, desc, work_year))

        contract_id = cursor.lastrowid
        if first_deposit > 0:
            cursor.execute('''
                INSERT INTO wedding_deposits (contract_id, amount, deposit_date, bank_name, work_year)
                VALUES (?, ?, ?, ?, ?)
            ''', (contract_id, first_deposit, self.w_contract_date.text(), bank_name, work_year))

        for item in selected_items:
            cursor.execute("UPDATE inventory SET used_count = used_count + 1 WHERE item_name = ?", (item,))

        conn.commit()
        conn.close()

        QMessageBox.information(self, "موفقیت", f"قرارداد با کد {code} ثبت شد.")
        self.w_groom.clear()
        self.w_bride.clear()
        self.w_groom_phone.clear()
        self.w_bride_phone.clear()
        self.w_discount.setText("0")
        self.w_first_deposit.setText("0")
        self.w_desc.clear()
        for cb in self.item_checkboxes.values():
            cb.setChecked(False)

        self.load_wedding_contracts()
        self.load_inventory()
        self.load_work_years()

    def load_wedding_contracts(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        work_year = get_current_year()
        cursor.execute('''
            SELECT id, code, groom_name, bride_name, groom_phone, bride_phone,
                   ceremony_date, total_amount, discount, paid_amount, is_settled
            FROM wedding_contracts WHERE work_year=? ORDER BY id DESC
        ''', (work_year,))
        rows = cursor.fetchall()
        conn.close()

        self.w_table.setRowCount(0)
        for r_idx, row in enumerate(rows):
            (c_id, code, groom, bride, g_phone, b_phone,
             cer_date, total, discount, paid, settled) = row
            remain = total - paid

            self.w_table.insertRow(r_idx)
            self.w_table.setItem(r_idx, 0, QTableWidgetItem(code or str(c_id)))
            self.w_table.setItem(r_idx, 1, QTableWidgetItem(f"{groom} و {bride}"))
            self.w_table.setItem(r_idx, 2, QTableWidgetItem(g_phone if g_phone else "-"))
            self.w_table.setItem(r_idx, 3, QTableWidgetItem(b_phone if b_phone else "-"))
            self.w_table.setItem(r_idx, 4, QTableWidgetItem(cer_date))
            self.w_table.setItem(r_idx, 5, QTableWidgetItem(f"{total:,}"))
            self.w_table.setItem(r_idx, 6, QTableWidgetItem(f"{discount:,}"))
            self.w_table.setItem(r_idx, 7, QTableWidgetItem(f"{paid:,}"))
            self.w_table.setItem(r_idx, 8, QTableWidgetItem(f"{remain:,}"))

            item_settled = QTableWidgetItem("✅ تسویه" if settled else "⏳ در جریان")
            if settled:
                item_settled.setForeground(QColor("green"))
            else:
                item_settled.setForeground(QColor("#e67e22"))
            self.w_table.setItem(r_idx, 9, item_settled)

            btn_dep = QPushButton("💰 بیعانه‌ها")
            btn_dep.clicked.connect(lambda _, cid=c_id: self.open_deposits_dialog(cid))
            self.w_table.setCellWidget(r_idx, 10, btn_dep)

            print_widget = QWidget()
            print_layout = QHBoxLayout()
            print_layout.setContentsMargins(0, 0, 0, 0)
            print_layout.setSpacing(2)
            btn_pdf = QPushButton("PDF")
            btn_pdf.setStyleSheet("background-color: #c0392b; color: white; padding: 4px;")
            btn_pdf.clicked.connect(lambda _, cid=c_id: self.export_wedding_pdf(cid))
            btn_chap = QPushButton("🖨️ چاپ")
            btn_chap.setStyleSheet("background-color: #2980b9; color: white; padding: 4px;")
            btn_chap.clicked.connect(lambda _, cid=c_id: self.print_wedding(cid))
            print_layout.addWidget(btn_pdf)
            print_layout.addWidget(btn_chap)
            print_widget.setLayout(print_layout)
            self.w_table.setCellWidget(r_idx, 11, print_widget)

            btn_edit = QPushButton("✏️")
            btn_edit.setStyleSheet("background-color: #f39c12; color: white;")
            btn_edit.clicked.connect(lambda _, cid=c_id: self.edit_wedding_contract(cid))
            self.w_table.setCellWidget(r_idx, 12, btn_edit)

            btn_del = QPushButton("🗑️")
            btn_del.setStyleSheet("background-color: #e74c3c; color: white;")
            btn_del.clicked.connect(lambda _, cid=c_id: self.delete_wedding_contract(cid))
            self.w_table.setCellWidget(r_idx, 13, btn_del)

        self.w_table.resizeRowsToContents()

    def edit_wedding_contract(self, contract_id):
        dlg = EditContractDialog(contract_id, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.load_wedding_contracts()

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
        code = self.w_table.item(row, 0).text()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM wedding_contracts WHERE code=?", (code,))
        r = cursor.fetchone()
        conn.close()
        if r:
            dlg = ContractDetailsDialog(r[0], self)
            dlg.exec()

    def open_deposits_dialog(self, contract_id):
        dlg = DepositsDialog(contract_id, self)
        dlg.exec()
        self.load_wedding_contracts()

    def export_wedding_pdf(self, contract_id):
        html, filename = InvoiceBuilder.build_wedding_invoice(contract_id, with_details=True)
        if not html:
            return
        file_path, _ = QFileDialog.getSaveFileName(
            self, "ذخیره فاکتور PDF", filename, "PDF Files (*.pdf)"
        )
        if not file_path:
            return
        try:
            doc = QTextDocument()
            doc.setHtml(html)
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
            printer.setOutputFileName(file_path)
            printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
            doc.print_(printer)
            QMessageBox.information(self, "موفقیت", "فاکتور PDF با موفقیت ساخته شد.")
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در ساخت PDF:\n{e}")

    def print_wedding(self, contract_id):
        html, _ = InvoiceBuilder.build_wedding_invoice(contract_id, with_details=True)
        if not html:
            return
        try:
            doc = QTextDocument()
            doc.setHtml(html)
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
            dialog = QPrintDialog(printer, self)
            if dialog.exec() == QPrintDialog.DialogCode.Accepted:
                doc.print_(printer)
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا در چاپ:\n{e}")

    # ============================================================
    # ==============  تب ۲: تبلیغاتی / بیوتی  ===================
    # ============================================================
    def setup_commercial_tab(self):
        layout = QHBoxLayout()
        form_box = QGroupBox("ثبت پروژه جدید تبلیغاتی / بیوتی / تولدی")
        form_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        form_layout = QFormLayout()

        self.c_title = QLineEdit()
        self.c_type = QComboBox()
        self.c_type.addItems(["تبلیغاتی", "بیوتی", "تولدی", "عروسی", "عقد"])
        self.c_client_name = QLineEdit()
        self.c_client_phone = QLineEdit()
        self.c_cameras = QSpinBox()
        self.c_cameras.setRange(1, 20)
        self.c_cameras.setValue(1)
        self.c_amount = QLineEdit()
        self.c_amount.textChanged.connect(lambda t: self.c_amount.setText(format_number(t)))
        self.c_paid = QLineEdit()
        self.c_paid.textChanged.connect(lambda t: self.c_paid.setText(format_number(t)))
        self.c_date = QLineEdit(jdatetime.date.today().strftime("%Y/%m/%d"))
        self.c_desc = QTextEdit()
        self.c_desc.setFixedHeight(60)

        form_layout.addRow("عنوان پروژه <span style='color:red;'>*</span>:", self.c_title)
        form_layout.addRow("نوع پروژه:", self.c_type)
        form_layout.addRow("نام مشتری:", self.c_client_name)
        form_layout.addRow("تلفن مشتری:", self.c_client_phone)
        form_layout.addRow("تعداد دوربین:", self.c_cameras)
        form_layout.addRow("مبلغ کل (تومان) <span style='color:red;'>*</span>:", self.c_amount)
        form_layout.addRow("مبلغ پرداختی (تومان):", self.c_paid)
        form_layout.addRow("تاریخ پروژه:", self.c_date)
        form_layout.addRow("توضیحات:", self.c_desc)

        btn_save = QPushButton("💾 ثبت پروژه")
        btn_save.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        btn_save.setStyleSheet("background-color: #27ae60; color: white; padding: 8px;")
        btn_save.clicked.connect(self.save_commercial_project)
        form_layout.addRow(btn_save)

        form_box.setLayout(form_layout)
        layout.addWidget(form_box, 1)

        table_box = QGroupBox("لیست پروژه‌ها")
        table_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        t_layout = QVBoxLayout()

        self.c_table = QTableWidget()
        self.c_table.setColumnCount(11)
        self.c_table.setHorizontalHeaderLabels([
            "کد", "عنوان", "نوع", "مشتری", "دوربین", "مبلغ کل",
            "پرداختی", "مانده", "تاریخ", "چاپ/PDF", "عملیات"
        ])
        self.c_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.c_table.doubleClicked.connect(self.on_commercial_double_click)
        t_layout.addWidget(self.c_table)
        table_box.setLayout(t_layout)
        layout.addWidget(table_box, 2)

        self.tab_commercial.setLayout(layout)
        self.load_commercial_projects()

    def save_commercial_project(self):
        title = self.c_title.text().strip()
        amount = parse_number(self.c_amount.text())
        if not title or amount <= 0:
            QMessageBox.warning(self, "خطا", "لطفاً عنوان و مبلغ کل را وارد کنید.")
            return

        work_year = get_current_year()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        cursor.execute("SELECT COUNT(*) FROM commercial_projects")
        count = cursor.fetchone()[0]
        code = f"C-{work_year}-{count + 1:04d}"

        cursor.execute('''
            INSERT INTO commercial_projects 
            (code, title, project_type, camera_count, total_amount, paid_amount,
             project_date, description, client_name, client_phone, work_year)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (code, title, self.c_type.currentText(), self.c_cameras.value(),
              amount, parse_number(self.c_paid.text()), self.c_date.text(),
              self.c_desc.toPlainText(), self.c_client_name.text(),
              self.c_client_phone.text(), work_year))

        conn.commit()
        conn.close()
        QMessageBox.information(self, "موفقیت", f"پروژه با کد {code} ثبت شد.")

        self.c_title.clear()
        self.c_client_name.clear()
        self.c_client_phone.clear()
        self.c_amount.clear()
        self.c_paid.clear()
        self.c_desc.clear()

        self.load_commercial_projects()
        self.load_work_years()

    def load_commercial_projects(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        work_year = get_current_year()
        cursor.execute("""SELECT id, code, title, project_type, camera_count,
                          total_amount, paid_amount, project_date, description,
                          client_name, client_phone
                          FROM commercial_projects WHERE work_year=? ORDER BY id DESC""",
                       (work_year,))
        rows = cursor.fetchall()
        conn.close()

        self.c_table.setRowCount(0)
        for r_idx, row in enumerate(rows):
            (p_id, code, title, ptype, cameras, total, paid, date, desc,
             client, phone) = row
            remain = total - paid

            self.c_table.insertRow(r_idx)
            self.c_table.setItem(r_idx, 0, QTableWidgetItem(code or str(p_id)))
            self.c_table.setItem(r_idx, 1, QTableWidgetItem(title))
            self.c_table.setItem(r_idx, 2, QTableWidgetItem(ptype))
            self.c_table.setItem(r_idx, 3, QTableWidgetItem(f"{client or '-'}\n{phone or ''}"))
            self.c_table.setItem(r_idx, 4, QTableWidgetItem(str(cameras)))
            self.c_table.setItem(r_idx, 5, QTableWidgetItem(f"{total:,}"))
            self.c_table.setItem(r_idx, 6, QTableWidgetItem(f"{paid:,}"))
            remain_item = QTableWidgetItem(f"{remain:,}")
            if remain > 0:
                remain_item.setForeground(QColor("#c0392b"))
            self.c_table.setItem(r_idx, 7, remain_item)
            self.c_table.setItem(r_idx, 8, QTableWidgetItem(date))

            pw = QWidget()
            pl = QHBoxLayout()
            pl.setContentsMargins(0, 0, 0, 0)
            pl.setSpacing(2)
            btn_pdf = QPushButton("PDF")
            btn_pdf.setStyleSheet("background-color: #c0392b; color: white; padding: 4px;")
            btn_pdf.clicked.connect(lambda _, pid=p_id: self.export_commercial_pdf(pid))
            btn_chap = QPushButton("🖨️")
            btn_chap.setStyleSheet("background-color: #2980b9; color: white; padding: 4px;")
            btn_chap.clicked.connect(lambda _, pid=p_id: self.print_commercial(pid))
            pl.addWidget(btn_pdf)
            pl.addWidget(btn_chap)
            pw.setLayout(pl)
            self.c_table.setCellWidget(r_idx, 9, pw)

            aw = QWidget()
            al = QHBoxLayout()
            al.setContentsMargins(0, 0, 0, 0)
            al.setSpacing(2)
            btn_edit = QPushButton("✏️")
            btn_edit.setStyleSheet("background-color: #f39c12; color: white;")
            btn_edit.clicked.connect(lambda _, pid=p_id: self.edit_commercial_project(pid))
            btn_del = QPushButton("🗑️")
            btn_del.setStyleSheet("background-color: #e74c3c; color: white;")
            btn_del.clicked.connect(lambda _, pid=p_id: self.delete_commercial_project(pid))
            al.addWidget(btn_edit)
            al.addWidget(btn_del)
            aw.setLayout(al)
            self.c_table.setCellWidget(r_idx, 10, aw)

    def edit_commercial_project(self, p_id):
        dlg = EditCommercialDialog(p_id, self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            self.load_commercial_projects()

    def on_commercial_double_click(self, index):
        row = index.row()
        code = self.c_table.item(row, 0).text()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id FROM commercial_projects WHERE code=?", (code,))
        r = cursor.fetchone()
        conn.close()
        if r:
            dlg = EditCommercialDialog(r[0], self)
            if dlg.exec() == QDialog.DialogCode.Accepted:
                self.load_commercial_projects()

    def export_commercial_pdf(self, p_id):
        html, filename = InvoiceBuilder.build_commercial_invoice(p_id)
        if not html:
            return
        file_path, _ = QFileDialog.getSaveFileName(
            self, "ذخیره فاکتور PDF", filename, "PDF Files (*.pdf)"
        )
        if not file_path:
            return
        try:
            doc = QTextDocument()
            doc.setHtml(html)
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
            printer.setOutputFileName(file_path)
            printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
            doc.print_(printer)
            QMessageBox.information(self, "موفقیت", "PDF ساخته شد.")
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا:\n{e}")

    def print_commercial(self, p_id):
        html, _ = InvoiceBuilder.build_commercial_invoice(p_id)
        if not html:
            return
        try:
            doc = QTextDocument()
            doc.setHtml(html)
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
            dialog = QPrintDialog(printer, self)
            if dialog.exec() == QPrintDialog.DialogCode.Accepted:
                doc.print_(printer)
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا:\n{e}")

    def delete_commercial_project(self, p_id):
        if not ask_security_password(self):
            return
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM commercial_projects WHERE id=?", (p_id,))
        conn.commit()
        conn.close()
        self.load_commercial_projects()

    # ============================================================
    # ==============  تب ۳: پرسنل و کارکنان  ====================
    # ============================================================
    def setup_staff_tab(self):
        layout = QVBoxLayout()
        forms_layout = QHBoxLayout()

        person_box = QGroupBox("تعریف نیروی جدید")
        person_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        person_form = QFormLayout()
        self.txt_person_code = QLineEdit()
        self.txt_person_name = QLineEdit()
        self.txt_person_phone = QLineEdit()
        self.combo_person_role = QComboBox()
        self.combo_person_role.addItems(self.ROLES)
        btn_add_person = QPushButton("➕ ثبت فرد جدید")
        btn_add_person.setStyleSheet("background-color: #8e44ad; color: white; font-weight: bold;")
        btn_add_person.clicked.connect(self.add_person)

        person_form.addRow("کد پرسنل:", self.txt_person_code)
        person_form.addRow("نام و نام خانوادگی <span style='color:red;'>*</span>:", self.txt_person_name)
        person_form.addRow("تخصص / نقش:", self.combo_person_role)
        person_form.addRow("تلفن همراه:", self.txt_person_phone)
        person_form.addRow(btn_add_person)
        person_box.setLayout(person_form)

        trans_box = QGroupBox("ثبت حقوق و پرداختی")
        trans_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        trans_form = QFormLayout()

        self.combo_staff_persons = QComboBox()
        self.txt_staff_amount = QLineEdit()
        self.txt_staff_amount.textChanged.connect(lambda t: self.txt_staff_amount.setText(format_number(t)))
        self.txt_staff_desc = QLineEdit()

        btn_add_trans = QPushButton("💵 ثبت پرداخت حقوق")
        btn_add_trans.setStyleSheet("background-color: #8e44ad; color: white; font-weight: bold;")
        btn_add_trans.clicked.connect(self.add_staff_payment)

        trans_form.addRow("انتخاب فرد <span style='color:red;'>*</span>:", self.combo_staff_persons)
        trans_form.addRow("مبلغ (تومان) <span style='color:red;'>*</span>:", self.txt_staff_amount)
        trans_form.addRow("بابت / توضیحات:", self.txt_staff_desc)
        trans_form.addRow(btn_add_trans)
        trans_box.setLayout(trans_form)

        forms_layout.addWidget(person_box, 1)
        forms_layout.addWidget(trans_box, 2)
        layout.addLayout(forms_layout)

        report_box = QGroupBox("گزارش ریز کارکرد (هفتگی / ماهانه / سالانه)")
        report_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        rep_layout = QHBoxLayout()

        self.combo_report_person = QComboBox()
        self.combo_report_period = QComboBox()
        self.combo_report_period.addItems(["هفتگی", "ماهانه", "سالانه", "کل"])

        btn_show = QPushButton("🔍 نمایش گزارش")
        btn_show.setStyleSheet("background-color: #2980b9; color: white; font-weight: bold;")
        btn_show.clicked.connect(self.show_staff_report)

        btn_rep_pdf = QPushButton("📄 PDF")
        btn_rep_pdf.setStyleSheet("background-color: #c0392b; color: white; font-weight: bold;")
        btn_rep_pdf.clicked.connect(self.export_staff_report_pdf)

        btn_rep_print = QPushButton("🖨️ چاپ")
        btn_rep_print.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold;")
        btn_rep_print.clicked.connect(self.print_staff_report)

        rep_layout.addWidget(QLabel("فرد:"))
        rep_layout.addWidget(self.combo_report_person, 2)
        rep_layout.addWidget(QLabel("دوره:"))
        rep_layout.addWidget(self.combo_report_period, 1)
        rep_layout.addWidget(btn_show)
        rep_layout.addWidget(btn_rep_pdf)
        rep_layout.addWidget(btn_rep_print)
        rep_layout.addStretch()
        report_box.setLayout(rep_layout)
        layout.addWidget(report_box)

        self.staff_table = QTableWidget()
        self.staff_table.setColumnCount(7)
        self.staff_table.setHorizontalHeaderLabels([
            "کد", "نام فرد", "نقش/تخصص", "مبلغ پرداخت (تومان)",
            "توضیحات", "تاریخ", "حذف"
        ])
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

        code = self.txt_person_code.text().strip()
        work_year = get_current_year()

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        if not code:
            cursor.execute("SELECT COUNT(*) FROM persons")
            count = cursor.fetchone()[0]
            code = f"P-{work_year}-{count + 1:04d}"

        cursor.execute("""INSERT INTO persons (code, name, role, phone, work_year)
                          VALUES (?, ?, ?, ?, ?)""",
                       (code, name, self.combo_person_role.currentText(),
                        self.txt_person_phone.text(), work_year))
        conn.commit()
        conn.close()

        self.txt_person_code.clear()
        self.txt_person_name.clear()
        self.txt_person_phone.clear()
        self.load_persons_combo()
        QMessageBox.information(self, "موفقیت", f"فرد با کد {code} ثبت شد.")

    def load_persons_combo(self):
        self.combo_staff_persons.clear()
        self.combo_report_person.clear()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, code, name, role FROM persons ORDER BY id DESC")
        for p_id, code, name, role in cursor.fetchall():
            label = f"[{code or '-'}] {name} ({role})"
            self.combo_staff_persons.addItem(label, p_id)
            self.combo_report_person.addItem(label, p_id)
        conn.close()

    def add_staff_payment(self):
        person_id = self.combo_staff_persons.currentData()
        amount = parse_number(self.txt_staff_amount.text())
        if not person_id or amount <= 0:
            QMessageBox.warning(self, "خطا", "لطفاً فرد و مبلغ معتبر وارد کنید.")
            return

        today = jdatetime.date.today()
        work_year = get_current_year()

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO transactions 
            (trans_type, category, person_id, amount, year, month, day, description, work_year)
            VALUES ('expense', 'حقوق کارکنان', ?, ?, ?, ?, ?, ?, ?)
        ''', (person_id, amount, today.year, today.month, today.day,
              self.txt_staff_desc.text(), work_year))
        conn.commit()
        conn.close()

        self.txt_staff_amount.clear()
        self.txt_staff_desc.clear()
        self.load_staff_table()

    def load_staff_table(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        work_year = get_current_year()
        cursor.execute('''
            SELECT t.id, p.code, p.name, p.role, t.amount, t.description,
                   t.year, t.month, t.day
            FROM transactions t
            JOIN persons p ON t.person_id = p.id
            WHERE t.work_year=?
            ORDER BY t.id DESC
        ''', (work_year,))
        rows = cursor.fetchall()
        conn.close()

        self.staff_table.setRowCount(0)
        for r_idx, row in enumerate(rows):
            (tid, code, name, role, amount, desc, yr, mo, dy) = row
            self.staff_table.insertRow(r_idx)
            self.staff_table.setItem(r_idx, 0, QTableWidgetItem(code or "-"))
            self.staff_table.setItem(r_idx, 1, QTableWidgetItem(name))
            self.staff_table.setItem(r_idx, 2, QTableWidgetItem(role))
            self.staff_table.setItem(r_idx, 3, QTableWidgetItem(f"{amount:,}"))
            self.staff_table.setItem(r_idx, 4, QTableWidgetItem(desc if desc else "-"))
            self.staff_table.setItem(r_idx, 5, QTableWidgetItem(f"{yr}/{mo:02d}/{dy:02d}"))

            btn_del = QPushButton("🗑️")
            btn_del.setStyleSheet("background-color: #e74c3c; color: white;")
            btn_del.clicked.connect(lambda _, x=tid: self.delete_staff_trans(x))
            self.staff_table.setCellWidget(r_idx, 6, btn_del)

    def delete_staff_trans(self, tid):
        if not ask_security_password(self):
            return
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM transactions WHERE id=?", (tid,))
        conn.commit()
        conn.close()
        self.load_staff_table()

    def show_staff_report(self):
        person_id = self.combo_report_person.currentData()
        if not person_id:
            QMessageBox.warning(self, "خطا", "لطفاً یک فرد انتخاب کنید.")
            return

        period = self.combo_report_period.currentText()
        work_year = get_current_year()

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("""SELECT amount, year, month, day, description, category
                          FROM transactions WHERE person_id=? AND work_year=?
                          ORDER BY year DESC, month DESC, day DESC""",
                       (person_id, work_year))
        all_trans = cursor.fetchall()
        conn.close()

        today = jdatetime.date.today()
        filtered = []
        for t in all_trans:
            try:
                tdate = jdatetime.date(t[1], t[2], t[3])
                diff = (today - tdate).days
                if period == "هفتگی" and 0 <= diff <= 7:
                    filtered.append(t)
                elif period == "ماهانه" and 0 <= diff <= 30:
                    filtered.append(t)
                elif period == "سالانه":
                    filtered.append(t)
                elif period == "کل":
                    filtered.append(t)
            except Exception:
                filtered.append(t)

        if not filtered:
            QMessageBox.information(self, "گزارش", "تراکنشی در این بازه یافت نشد.")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle(f"ریز کارکرد - {period}")
        dlg.resize(700, 500)
        dlg.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        dlg.setFont(QFont(APP_FONT_FAMILY, 10))
        v = QVBoxLayout()

        total = sum(t[0] for t in filtered)
        lbl = QLabel(f"<b>جمع کل پرداختی در {period}:</b> {total:,} تومان")
        lbl.setStyleSheet("background-color: #e8f8f5; padding: 10px; border-radius: 5px; font-size: 12pt;")
        v.addWidget(lbl)

        table = QTableWidget()
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["تاریخ", "دسته", "شرح", "مبلغ (تومان)"])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.setRowCount(len(filtered))
        for idx, t in enumerate(filtered):
            table.setItem(idx, 0, QTableWidgetItem(f"{t[1]}/{t[2]:02d}/{t[3]:02d}"))
            table.setItem(idx, 1, QTableWidgetItem(t[5]))
            table.setItem(idx, 2, QTableWidgetItem(t[4] if t[4] else "-"))
            table.setItem(idx, 3, QTableWidgetItem(f"{t[0]:,}"))
        v.addWidget(table)

        dlg.setLayout(v)
        dlg.exec()

    def export_staff_report_pdf(self):
        person_id = self.combo_report_person.currentData()
        if not person_id:
            QMessageBox.warning(self, "خطا", "لطفاً یک فرد انتخاب کنید.")
            return
        period = self.combo_report_period.currentText()
        html, filename = InvoiceBuilder.build_staff_report(
            person_id, period=period, year=get_current_year()
        )
        if not html:
            return
        file_path, _ = QFileDialog.getSaveFileName(self, "ذخیره PDF", filename, "PDF Files (*.pdf)")
        if not file_path:
            return
        try:
            doc = QTextDocument()
            doc.setHtml(html)
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
            printer.setOutputFileName(file_path)
            printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
            doc.print_(printer)
            QMessageBox.information(self, "موفقیت", "PDF ساخته شد.")
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا:\n{e}")

    def print_staff_report(self):
        person_id = self.combo_report_person.currentData()
        if not person_id:
            QMessageBox.warning(self, "خطا", "لطفاً یک فرد انتخاب کنید.")
            return
        period = self.combo_report_period.currentText()
        html, _ = InvoiceBuilder.build_staff_report(
            person_id, period=period, year=get_current_year()
        )
        if not html:
            return
        try:
            doc = QTextDocument()
            doc.setHtml(html)
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
            dialog = QPrintDialog(printer, self)
            if dialog.exec() == QPrintDialog.DialogCode.Accepted:
                doc.print_(printer)
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا:\n{e}")

    # ============================================================
    # ==============  تب ۴: مدیریت هزینه‌ها  ====================
    # ============================================================
    def setup_expenses_tab(self):
        layout = QHBoxLayout()
        form_box = QGroupBox("ثبت هزینه جدید")
        form_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        form_layout = QFormLayout()

        self.exp_code = QLineEdit()
        self.exp_title = QLineEdit()
        self.exp_cat = QComboBox()
        self.exp_cat.addItems([
            "قبوض (آب، برق، گاز، تلفن)", "کرایه آتلیه / دفتر",
            "خرید تجهیزات و مصرفی", "تبلیغات و بازاریابی",
            "حقوق و دستمزد", "سایر هزینه‌ها"
        ])
        self.exp_amount = QLineEdit()
        self.exp_amount.textChanged.connect(lambda t: self.exp_amount.setText(format_number(t)))
        self.exp_date = QLineEdit(jdatetime.date.today().strftime("%Y/%m/%d"))
        self.exp_desc = QTextEdit()
        self.exp_desc.setFixedHeight(60)

        form_layout.addRow("کد هزینه:", self.exp_code)
        form_layout.addRow("عنوان هزینه <span style='color:red;'>*</span>:", self.exp_title)
        form_layout.addRow("دسته‌بندی:", self.exp_cat)
        form_layout.addRow("مبلغ (تومان) <span style='color:red;'>*</span>:", self.exp_amount)
        form_layout.addRow("تاریخ هزینه:", self.exp_date)
        form_layout.addRow("توضیحات:", self.exp_desc)

        btn_save = QPushButton("💾 ثبت هزینه")
        btn_save.setStyleSheet("background-color: #c0392b; color: white; font-weight: bold; padding: 8px;")
        btn_save.clicked.connect(self.save_expense)
        form_layout.addRow(btn_save)

        form_box.setLayout(form_layout)
        layout.addWidget(form_box, 1)

        table_box = QGroupBox("لیست هزینه‌ها")
        table_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        t_layout = QVBoxLayout()

        self.exp_table = QTableWidget()
        self.exp_table.setColumnCount(8)
        self.exp_table.setHorizontalHeaderLabels([
            "کد", "عنوان", "دسته‌بندی", "مبلغ (تومان)",
            "تاریخ", "توضیحات", "چاپ", "حذف"
        ])
        self.exp_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        t_layout.addWidget(self.exp_table)
        table_box.setLayout(t_layout)
        layout.addWidget(table_box, 2)

        self.tab_expenses.setLayout(layout)
        self.load_expenses_table()

    def save_expense(self):
        title = self.exp_title.text().strip()
        amount = parse_number(self.exp_amount.text())
        if not title or amount <= 0:
            QMessageBox.warning(self, "خطا", "لطفاً عنوان و مبلغ را وارد کنید.")
            return

        work_year = get_current_year()
        code = self.exp_code.text().strip()

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        if not code:
            cursor.execute("SELECT COUNT(*) FROM expenses")
            count = cursor.fetchone()[0]
            code = f"E-{work_year}-{count + 1:04d}"

        cursor.execute('''INSERT INTO expenses 
            (code, title, category, amount, date_str, description, work_year)
            VALUES (?, ?, ?, ?, ?, ?, ?)''',
            (code, title, self.exp_cat.currentText(), amount,
             self.exp_date.text(), self.exp_desc.toPlainText(), work_year))
        conn.commit()
        conn.close()

        self.exp_code.clear()
        self.exp_title.clear()
        self.exp_amount.clear()
        self.exp_desc.clear()
        self.load_expenses_table()

    def load_expenses_table(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        work_year = get_current_year()
        cursor.execute("""SELECT id, code, title, category, amount, date_str, description
                          FROM expenses WHERE work_year=? ORDER BY id DESC""",
                       (work_year,))
        rows = cursor.fetchall()
        conn.close()

        self.exp_table.setRowCount(0)
        for r_idx, row in enumerate(rows):
            (eid, code, title, cat, amount, date_str, desc) = row
            self.exp_table.insertRow(r_idx)
            self.exp_table.setItem(r_idx, 0, QTableWidgetItem(code or "-"))
            self.exp_table.setItem(r_idx, 1, QTableWidgetItem(title))
            self.exp_table.setItem(r_idx, 2, QTableWidgetItem(cat))
            self.exp_table.setItem(r_idx, 3, QTableWidgetItem(f"{amount:,}"))
            self.exp_table.setItem(r_idx, 4, QTableWidgetItem(date_str))
            self.exp_table.setItem(r_idx, 5, QTableWidgetItem(desc if desc else "-"))

            btn_print = QPushButton("🖨️")
            btn_print.setStyleSheet("background-color: #2980b9; color: white;")
            btn_print.clicked.connect(lambda _, x=eid: self.print_expense(x))
            self.exp_table.setCellWidget(r_idx, 6, btn_print)

            btn_del = QPushButton("🗑️")
            btn_del.setStyleSheet("background-color: #e74c3c; color: white;")
            btn_del.clicked.connect(lambda _, x=eid: self.delete_expense(x))
            self.exp_table.setCellWidget(r_idx, 7, btn_del)

    def print_expense(self, eid):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT code, title, category, amount, date_str, description FROM expenses WHERE id=?", (eid,))
        r = cursor.fetchone()
        conn.close()
        if not r:
            return

        studio_name = get_setting("studio_name", "IMART STUDIO")
        now = jdatetime.datetime.now()
        html = f"""
        <div dir="rtl" style="font-family: '{INVOICE_FONT_FAMILY}', Tahoma; padding: 20px;">
            <h1 style="text-align: center; color: #1F4E78; font-size: 26pt; margin: 0;">{studio_name}</h1>
            <h3 style="text-align: center; margin: 5px 0;">رسید هزینه</h3>
            <hr style="border: 2px solid #1F4E78;">
            <p style="font-size: 12pt; line-height: 2;">
                <b>کد:</b> {r[0]}<br>
                <b>عنوان:</b> {r[1]}<br>
                <b>دسته:</b> {r[2]}<br>
                <b>مبلغ:</b> {r[3]:,} تومان ({number_to_persian_words(r[3])} تومان)<br>
                <b>تاریخ:</b> {r[4]}<br>
                <b>توضیحات:</b> {r[5] or '-'}
            </p>
            <br><br>
            <p style="text-align: left;">تاریخ چاپ: {now.strftime('%Y/%m/%d %H:%M')}</p>
            <br><br>
            <p style="text-align: center;">مهر و امضای حسابداری</p>
        </div>
        """
        doc = QTextDocument()
        doc.setHtml(html)
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
        d = QPrintDialog(printer, self)
        if d.exec() == QPrintDialog.DialogCode.Accepted:
            doc.print_(printer)

    def delete_expense(self, eid):
        if not ask_security_password(self):
            return
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM expenses WHERE id=?", (eid,))
        conn.commit()
        conn.close()
        self.load_expenses_table()
            # ============================================================
    # ==============  تب ۵: گزارش مالی جامع  ====================
    # ============================================================
    def setup_reports_tab(self):
        layout = QVBoxLayout()

        top_layout = QHBoxLayout()
        self.combo_report_type = QComboBox()
        self.combo_report_type.addItems(["گزارش کل", "گزارش ماهانه", "گزارش سالانه", "گزارش هر فرد"])
        top_layout.addWidget(QLabel("نوع گزارش:"))
        top_layout.addWidget(self.combo_report_type)

        top_layout.addWidget(QLabel("فرد (برای گزارش فردی):"))
        self.combo_report_person_fin = QComboBox()
        self.load_report_persons_combo()
        top_layout.addWidget(self.combo_report_person_fin)

        btn_calc = QPushButton("🧮 محاسبه گزارش")
        btn_calc.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; padding: 6px;")
        btn_calc.clicked.connect(self.calculate_financial_report)
        top_layout.addWidget(btn_calc)

        btn_chart = QPushButton("📊 نمایش نمودار")
        btn_chart.setStyleSheet("background-color: #2980b9; color: white; font-weight: bold; padding: 6px;")
        btn_chart.clicked.connect(self.show_financial_chart)
        top_layout.addWidget(btn_chart)

        btn_print_rep = QPushButton("🖨️ چاپ گزارش")
        btn_print_rep.setStyleSheet("background-color: #8e44ad; color: white; font-weight: bold; padding: 6px;")
        btn_print_rep.clicked.connect(self.print_financial_report)
        top_layout.addWidget(btn_print_rep)

        layout.addLayout(top_layout)

        self.lbl_rep_summary = QLabel("ورودی کل: ۰ تومان | خروجی کل: ۰ تومان | سود خالص: ۰ تومان")
        self.lbl_rep_summary.setStyleSheet(
            "font-size: 12pt; font-weight: bold; background-color: #e8f8f5; "
            "padding: 12px; border-radius: 5px; border: 1px solid #16a085;"
        )
        layout.addWidget(self.lbl_rep_summary)

        self.rep_table = QTableWidget()
        self.rep_table.setColumnCount(5)
        self.rep_table.setHorizontalHeaderLabels([
            "ردیف", "بخش", "نوع تراکنش", "مبلغ (تومان)", "تاریخ / شرح"
        ])
        self.rep_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.rep_table)

        self.tab_reports.setLayout(layout)

    def load_report_persons_combo(self):
        self.combo_report_person_fin.clear()
        self.combo_report_person_fin.addItem("--- همه ---", None)
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, code, name, role FROM persons ORDER BY name")
        for p_id, code, name, role in cursor.fetchall():
            self.combo_report_person_fin.addItem(f"[{code or '-'}] {name} ({role})", p_id)
        conn.close()

    def calculate_financial_report(self):
        work_year = get_current_year()
        report_type = self.combo_report_type.currentText()

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        month_filter = ""
        if report_type == "گزارش ماهانه":
            today = jdatetime.date.today()
            month_filter = f" AND contract_date LIKE '{today.year}/{today.month:02d}%'"

        cursor.execute(f"SELECT SUM(total_amount) FROM wedding_contracts WHERE work_year=?{month_filter}", (work_year,))
        w_in = cursor.fetchone()[0] or 0

        cursor.execute(f"SELECT SUM(total_amount) FROM commercial_projects WHERE work_year=?{month_filter}", (work_year,))
        c_in = cursor.fetchone()[0] or 0

        cursor.execute(f"SELECT SUM(amount) FROM expenses WHERE work_year=?{month_filter}", (work_year,))
        exp_out = cursor.fetchone()[0] or 0

        cursor.execute(f"SELECT SUM(amount) FROM transactions WHERE work_year=?{month_filter}", (work_year,))
        staff_out = cursor.fetchone()[0] or 0

        total_in = w_in + c_in
        total_out = exp_out + staff_out
        profit = total_in - total_out

        self.lbl_rep_summary.setText(
            f"ورودی کل: {total_in:,} تومان | خروجی کل: {total_out:,} تومان | "
            f"سود خالص: <span style='color:{'green' if profit >= 0 else 'red'};'>{profit:,} تومان</span>"
        )

        self.rep_table.setRowCount(0)
        records = [
            ("قراردادهای عروس و داماد", "درآمد (ورودی)", w_in, "مجموع ورودی‌ها"),
            ("پروژه‌های تبلیغاتی / بیوتی / تولدی", "درآمد (ورودی)", c_in, "مجموع ورودی‌ها"),
            ("هزینه‌های جاری و قبوض", "هزینه (خروجی)", exp_out, "مجموع هزینه‌ها"),
            ("حقوق و پرداختی پرسنل", "هزینه (خروجی)", staff_out, "مجموع پرداختی‌ها")
        ]

        person_id = self.combo_report_person_fin.currentData()
        if report_type == "گزارش هر فرد" and person_id:
            cursor.execute("""SELECT name, role FROM persons WHERE id=?""", (person_id,))
            p = cursor.fetchone()
            if p:
                cursor.execute("""SELECT amount, year, month, day, description
                                  FROM transactions WHERE person_id=? AND work_year=?
                                  ORDER BY year DESC, month DESC, day DESC""",
                               (person_id, work_year))
                trans = cursor.fetchall()

                cursor.execute("SELECT SUM(amount) FROM transactions WHERE person_id=? AND work_year=?",
                               (person_id, work_year))
                total_person = cursor.fetchone()[0] or 0

                self.lbl_rep_summary.setText(
                    f"<b>گزارش فردی:</b> {p[0]} ({p[1]}) | "
                    f"<b>جمع کل پرداختی:</b> {total_person:,} تومان | "
                    f"<b>تعداد تراکنش:</b> {len(trans)}"
                )

                records = []
                for idx, t in enumerate(trans):
                    records.append((
                        f"پرداختی {t[1]}/{t[2]:02d}/{t[3]:02d}",
                        "حقوق کارکنان",
                        t[0],
                        t[4] if t[4] else "-"
                    ))

        conn.close()

        for idx, rec in enumerate(records):
            self.rep_table.insertRow(idx)
            self.rep_table.setItem(idx, 0, QTableWidgetItem(str(idx + 1)))
            self.rep_table.setItem(idx, 1, QTableWidgetItem(str(rec[0])))
            self.rep_table.setItem(idx, 2, QTableWidgetItem(str(rec[1])))
            self.rep_table.setItem(idx, 3, QTableWidgetItem(f"{rec[2]:,}"))
            self.rep_table.setItem(idx, 4, QTableWidgetItem(str(rec[3])))

    def show_financial_chart(self):
        # ✅ استفاده از تابع متمرکز فونت
        setup_matplotlib_font()

        work_year = get_current_year()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT SUM(total_amount) FROM wedding_contracts WHERE work_year=?", (work_year,))
        w_in = cursor.fetchone()[0] or 0
        cursor.execute("SELECT SUM(total_amount) FROM commercial_projects WHERE work_year=?", (work_year,))
        c_in = cursor.fetchone()[0] or 0
        cursor.execute("SELECT SUM(amount) FROM expenses WHERE work_year=?", (work_year,))
        exp_out = cursor.fetchone()[0] or 0
        cursor.execute("SELECT SUM(amount) FROM transactions WHERE work_year=?", (work_year,))
        staff_out = cursor.fetchone()[0] or 0
        conn.close()

        labels = ['درآمد عروسی', 'درآمد تبلیغاتی', 'هزینه‌ها', 'حقوق پرسنل']
        amounts = [w_in, c_in, exp_out, staff_out]

        fig, ax = plt.subplots(figsize=(10, 6), facecolor='#f8f9fa')
        colors = ['#27ae60', '#2980b9', '#e74c3c', '#8e44ad']
        bars = ax.bar(labels, amounts, color=colors, edgecolor='#2c3e50', linewidth=1.5)

        for bar, amount in zip(bars, amounts):
            height = bar.get_height()
            ax.text(
                bar.get_x() + bar.get_width() / 2, height,
                f'{amount:,}\nتومان',
                ha='center', va='bottom',
                fontsize=10, fontweight='bold', color='#2c3e50'
            )

        ax.set_title(f'نمودار درآمد و هزینه - سال {work_year}',
                     fontsize=14, fontweight='bold', pad=20)
        ax.set_ylabel('مبلغ (تومان)', fontsize=11, fontweight='bold')
        ax.grid(axis='y', alpha=0.3, linestyle='--')
        ax.tick_params(axis='x', labelsize=11)
        ax.tick_params(axis='y', labelsize=10)

        ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'{int(x):,}'))

        plt.tight_layout()
        plt.show()

    def print_financial_report(self):
        work_year = get_current_year()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT SUM(total_amount) FROM wedding_contracts WHERE work_year=?", (work_year,))
        w_in = cursor.fetchone()[0] or 0
        cursor.execute("SELECT SUM(total_amount) FROM commercial_projects WHERE work_year=?", (work_year,))
        c_in = cursor.fetchone()[0] or 0
        cursor.execute("SELECT SUM(amount) FROM expenses WHERE work_year=?", (work_year,))
        exp_out = cursor.fetchone()[0] or 0
        cursor.execute("SELECT SUM(amount) FROM transactions WHERE work_year=?", (work_year,))
        staff_out = cursor.fetchone()[0] or 0
        conn.close()

        total_in = w_in + c_in
        total_out = exp_out + staff_out
        profit = total_in - total_out

        studio_name = get_setting("studio_name", "IMART STUDIO")
        now = jdatetime.datetime.now()

        html = f"""
        <div dir="rtl" style="font-family: '{INVOICE_FONT_FAMILY}', Tahoma; padding: 20px;">
            <table style="width:100%; border-collapse:collapse;">
                <tr>
                    <td style="text-align:center;">
                        <h1 style="color:#1F4E78; font-size:26pt; margin:0;">{studio_name}</h1>
                        <h3 style="margin:5px 0;">گزارش مالی جامع - سال {work_year}</h3>
                    </td>
                    <td style="text-align:left; width:25%;">
                        <div>تاریخ چاپ: {now.strftime('%Y/%m/%d')}</div>
                        <div>ساعت: {now.strftime('%H:%M')}</div>
                    </td>
                </tr>
            </table>
            <hr style="border:2px solid #1F4E78;">

            <table style="width:100%; border-collapse:collapse; margin-top:20px;">
                <thead>
                    <tr style="background-color:#D9E1F2;">
                        <th style="border:1px solid #000; padding:8px;">ردیف</th>
                        <th style="border:1px solid #000; padding:8px;">بخش</th>
                        <th style="border:1px solid #000; padding:8px;">نوع</th>
                        <th style="border:1px solid #000; padding:8px;">مبلغ (تومان)</th>
                    </tr>
                </thead>
                <tbody>
                    <tr><td style="border:1px solid #000; padding:6px; text-align:center;">1</td>
                        <td style="border:1px solid #000; padding:6px;">قراردادهای عروس و داماد</td>
                        <td style="border:1px solid #000; padding:6px; text-align:center;">درآمد</td>
                        <td style="border:1px solid #000; padding:6px; text-align:center;">{w_in:,}</td></tr>
                    <tr><td style="border:1px solid #000; padding:6px; text-align:center;">2</td>
                        <td style="border:1px solid #000; padding:6px;">پروژه‌های تبلیغاتی/بیوتی</td>
                        <td style="border:1px solid #000; padding:6px; text-align:center;">درآمد</td>
                        <td style="border:1px solid #000; padding:6px; text-align:center;">{c_in:,}</td></tr>
                    <tr><td style="border:1px solid #000; padding:6px; text-align:center;">3</td>
                        <td style="border:1px solid #000; padding:6px;">هزینه‌های جاری و قبوض</td>
                        <td style="border:1px solid #000; padding:6px; text-align:center;">هزینه</td>
                        <td style="border:1px solid #000; padding:6px; text-align:center;">{exp_out:,}</td></tr>
                    <tr><td style="border:1px solid #000; padding:6px; text-align:center;">4</td>
                        <td style="border:1px solid #000; padding:6px;">حقوق و پرداختی پرسنل</td>
                        <td style="border:1px solid #000; padding:6px; text-align:center;">هزینه</td>
                        <td style="border:1px solid #000; padding:6px; text-align:center;">{staff_out:,}</td></tr>
                </tbody>
            </table>

            <table style="width:100%; margin-top:20px; border-collapse:collapse;">
                <tr>
                    <td style="border:1px solid #000; padding:10px; background-color:#E8F8F5;">
                        <b>ورودی کل:</b> {total_in:,} تومان
                    </td>
                    <td style="border:1px solid #000; padding:10px; background-color:#FDEDEC;">
                        <b>خروجی کل:</b> {total_out:,} تومان
                    </td>
                    <td style="border:1px solid #000; padding:10px; background-color:{'#E8F8F5' if profit >= 0 else '#FDEDEC'};">
                        <b>سود خالص:</b> {profit:,} تومان
                    </td>
                </tr>
            </table>

            <p style="text-align:center; margin-top:50px; font-size:11pt;">مهر و امضای مدیر / حسابداری</p>
        </div>
        """

        doc = QTextDocument()
        doc.setHtml(html)
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
        d = QPrintDialog(printer, self)
        if d.exec() == QPrintDialog.DialogCode.Accepted:
            doc.print_(printer)

    # ============================================================
    # ==============  تب ۶: جستجوی پیشرفته  =====================
    # ============================================================
    def setup_search_tab(self):
        layout = QVBoxLayout()

        search_bar = QHBoxLayout()

        self.txt_search_keyword = QLineEdit()
        self.txt_search_keyword.setPlaceholderText("جستجو بر اساس نام، کد، تاریخ، عنوان...")
        self.txt_search_keyword.returnPressed.connect(self.perform_search)

        self.chk_search_by_date = QCheckBox("جستجو بر اساس تاریخ خاص")
        self.chk_search_by_date.stateChanged.connect(self.toggle_date_search)

        self.date_search = QDateEdit()
        self.date_search.setCalendarPopup(True)
        self.date_search.setDisplayFormat("yyyy/MM/dd")
        self.date_search.setEnabled(False)
        self.date_search.setDate(QDate.currentDate())

        btn_search = QPushButton("🔍 جستجو")
        btn_search.setStyleSheet("background-color: #16a085; color: white; font-weight: bold; padding: 6px;")
        btn_search.clicked.connect(self.perform_search)

        btn_clear = QPushButton("🧹 پاک کردن")
        btn_clear.setStyleSheet("background-color: #95a5a6; color: white; font-weight: bold; padding: 6px;")
        btn_clear.clicked.connect(self.clear_search)

        search_bar.addWidget(QLabel("عبارت:"))
        search_bar.addWidget(self.txt_search_keyword, 2)
        search_bar.addWidget(self.chk_search_by_date)
        search_bar.addWidget(self.date_search)
        search_bar.addWidget(btn_search)
        search_bar.addWidget(btn_clear)
        layout.addLayout(search_bar)

        self.search_table = QTableWidget()
        self.search_table.setColumnCount(7)
        self.search_table.setHorizontalHeaderLabels([
            "نوع", "کد", "عنوان / نام", "تاریخ", "مبلغ کل", "پرداختی", "مانده"
        ])
        self.search_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.search_table.doubleClicked.connect(self.on_search_double_click)
        layout.addWidget(self.search_table)

        detail_box = QGroupBox("جزئیات کامل (برای مشاهده، روی نتیجه دوبار کلیک کنید)")
        detail_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        dl = QVBoxLayout()
        self.lbl_search_detail = QLabel("نتیجه‌ای انتخاب نشده است.")
        self.lbl_search_detail.setWordWrap(True)
        self.lbl_search_detail.setStyleSheet(
            "background-color: #f8f9fa; padding: 12px; border-radius: 5px; "
            "border: 1px solid #bdc3c7; font-size: 11pt;"
        )
        self.lbl_search_detail.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        dl.addWidget(self.lbl_search_detail)
        detail_box.setLayout(dl)
        layout.addWidget(detail_box)

        self.tab_search.setLayout(layout)

    def toggle_date_search(self):
        self.date_search.setEnabled(self.chk_search_by_date.isChecked())
        if self.chk_search_by_date.isChecked():
            self.txt_search_keyword.setEnabled(False)
        else:
            self.txt_search_keyword.setEnabled(True)

    def clear_search(self):
        self.txt_search_keyword.clear()
        self.chk_search_by_date.setChecked(False)
        self.date_search.setEnabled(False)
        self.txt_search_keyword.setEnabled(True)
        self.search_table.setRowCount(0)
        self.lbl_search_detail.setText("نتیجه‌ای انتخاب نشده است.")

    def perform_search(self):
        self.search_table.setRowCount(0)
        self.lbl_search_detail.setText("در حال جستجو...")

        results = []

        if self.chk_search_by_date.isChecked():
            date_str = self.date_search.date().toString("yyyy/MM/dd")
            kw_pattern = f"%{date_str}%"
            use_kw = False
        else:
            kw = self.txt_search_keyword.text().strip()
            if not kw:
                QMessageBox.warning(self, "خطا", "عبارت جستجو را وارد کنید.")
                return
            kw_pattern = f"%{kw}%"
            use_kw = True

        work_year = get_current_year()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        # قراردادهای عروس
        if use_kw:
            cursor.execute("""SELECT id, code, groom_name, bride_name, ceremony_date,
                              total_amount, paid_amount
                              FROM wedding_contracts 
                              WHERE work_year=? AND (
                                groom_name LIKE ? OR bride_name LIKE ? OR 
                                code LIKE ? OR ceremony_date LIKE ? OR
                                groom_phone LIKE ? OR bride_phone LIKE ?
                              )""",
                           (work_year, kw_pattern, kw_pattern, kw_pattern,
                            kw_pattern, kw_pattern, kw_pattern))
        else:
            cursor.execute("""SELECT id, code, groom_name, bride_name, ceremony_date,
                              total_amount, paid_amount
                              FROM wedding_contracts 
                              WHERE work_year=? AND ceremony_date LIKE ?""",
                           (work_year, kw_pattern))
        for r in cursor.fetchall():
            results.append({
                "type": "قرارداد عروسی",
                "id": r[0], "code": r[1], "title": f"{r[2]} و {r[3]}",
                "date": r[4], "total": r[5], "paid": r[6],
                "table": "wedding_contracts"
            })

        # پروژه‌های تبلیغاتی
        if use_kw:
            cursor.execute("""SELECT id, code, title, client_name, project_date,
                              total_amount, paid_amount
                              FROM commercial_projects 
                              WHERE work_year=? AND (
                                title LIKE ? OR client_name LIKE ? OR
                                code LIKE ? OR project_date LIKE ? OR
                                client_phone LIKE ?
                              )""",
                           (work_year, kw_pattern, kw_pattern, kw_pattern,
                            kw_pattern, kw_pattern))
        else:
            cursor.execute("""SELECT id, code, title, client_name, project_date,
                              total_amount, paid_amount
                              FROM commercial_projects 
                              WHERE work_year=? AND project_date LIKE ?""",
                           (work_year, kw_pattern))
        for r in cursor.fetchall():
            results.append({
                "type": "پروژه تبلیغاتی",
                "id": r[0], "code": r[1], "title": r[2],
                "date": r[4], "total": r[5], "paid": r[6],
                "table": "commercial_projects"
            })

        # هزینه‌ها
        if use_kw:
            cursor.execute("""SELECT id, code, title, category, date_str, amount
                              FROM expenses 
                              WHERE work_year=? AND (
                                title LIKE ? OR category LIKE ? OR
                                code LIKE ? OR date_str LIKE ?
                              )""",
                           (work_year, kw_pattern, kw_pattern,
                            kw_pattern, kw_pattern))
        else:
            cursor.execute("""SELECT id, code, title, category, date_str, amount
                              FROM expenses WHERE work_year=? AND date_str LIKE ?""",
                           (work_year, kw_pattern))
        for r in cursor.fetchall():
            results.append({
                "type": "هزینه",
                "id": r[0], "code": r[1], "title": r[2],
                "date": r[4], "total": r[5], "paid": r[5],
                "table": "expenses"
            })

        # چک‌ها
        if use_kw:
            cursor.execute("""SELECT id, check_number, bank_name, issuer_name, due_date, amount, is_passed
                              FROM checks 
                              WHERE work_year=? AND (
                                check_number LIKE ? OR bank_name LIKE ? OR
                                issuer_name LIKE ? OR due_date LIKE ?
                              )""",
                           (work_year, kw_pattern, kw_pattern,
                            kw_pattern, kw_pattern))
        else:
            cursor.execute("""SELECT id, check_number, bank_name, issuer_name, due_date, amount, is_passed
                              FROM checks WHERE work_year=? AND due_date LIKE ?""",
                           (work_year, kw_pattern))
        for r in cursor.fetchall():
            results.append({
                "type": "چک",
                "id": r[0], "code": r[1], "title": f"{r[2]} - {r[3] or '-'}",
                "date": r[4], "total": r[5], "paid": r[5] if r[6] else 0,
                "table": "checks"
            })

        # پرسنل
        if use_kw:
            cursor.execute("""SELECT id, code, name, role, phone
                              FROM persons 
                              WHERE name LIKE ? OR code LIKE ? OR role LIKE ? OR phone LIKE ?""",
                           (kw_pattern, kw_pattern, kw_pattern, kw_pattern))
            for r in cursor.fetchall():
                results.append({
                    "type": "پرسنل",
                    "id": r[0], "code": r[1], "title": f"{r[2]} ({r[3]})",
                    "date": "-", "total": 0, "paid": 0,
                    "table": "persons"
                })

        # پرداختی پرسنل
        if use_kw:
            cursor.execute("""SELECT t.id, p.code, p.name, t.description, 
                                     t.year, t.month, t.day, t.amount
                              FROM transactions t
                              JOIN persons p ON t.person_id = p.id
                              WHERE t.work_year=? AND (
                                p.name LIKE ? OR p.code LIKE ? OR t.description LIKE ?
                              )""",
                           (work_year, kw_pattern, kw_pattern, kw_pattern))
            for r in cursor.fetchall():
                results.append({
                    "type": "پرداختی پرسنل",
                    "id": r[0], "code": r[1], "title": f"پرداختی به {r[2]}",
                    "date": f"{r[4]}/{r[5]:02d}/{r[6]:02d}",
                    "total": r[7], "paid": r[7],
                    "table": "transactions"
                })

        conn.close()

        self.search_table.setRowCount(0)
        for r_idx, res in enumerate(results):
            self.search_table.insertRow(r_idx)
            self.search_table.setItem(r_idx, 0, QTableWidgetItem(res["type"]))
            self.search_table.setItem(r_idx, 1, QTableWidgetItem(str(res["code"] or "-")))
            self.search_table.setItem(r_idx, 2, QTableWidgetItem(res["title"]))
            self.search_table.setItem(r_idx, 3, QTableWidgetItem(str(res["date"])))
            self.search_table.setItem(r_idx, 4, QTableWidgetItem(f"{res['total']:,}" if res['total'] else "-"))
            remain = res['total'] - res['paid']
            self.search_table.setItem(r_idx, 5, QTableWidgetItem(f"{res['paid']:,}" if res['paid'] else "-"))
            self.search_table.setItem(r_idx, 6, QTableWidgetItem(f"{remain:,}" if res['total'] else "-"))

            self.search_table.item(r_idx, 0).setData(Qt.ItemDataRole.UserRole, res)

        # ✅ پاک کردن خودکار کادر جستجو
        self.txt_search_keyword.clear()
        self.lbl_search_detail.setText(f"<b>{len(results)}</b> نتیجه یافت شد. برای مشاهده جزئیات روی نتیجه دوبار کلیک کنید.")

    def on_search_double_click(self, index):
        row = index.row()
        item = self.search_table.item(row, 0)
        if not item:
            return
        res = item.data(Qt.ItemDataRole.UserRole)
        if not res:
            return

        detail_html = f"<h3 style='color:#2980b9;'>جزئیات: {res['type']}</h3><hr>"
        detail_html += f"<p><b>کد:</b> {res['code']}<br>"
        detail_html += f"<b>عنوان:</b> {res['title']}<br>"
        detail_html += f"<b>تاریخ:</b> {res['date']}</p>"

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        if res["table"] == "wedding_contracts":
            cursor.execute("""SELECT groom_phone, bride_phone, contract_date, 
                                     ceremony_date, selected_items, description,
                                     discount, is_settled
                                     FROM wedding_contracts WHERE id=?""", (res["id"],))
            r = cursor.fetchone()
            if r:
                detail_html += f"<p><b>تلفن داماد:</b> {r[0] or '-'} | <b>تلفن عروس:</b> {r[1] or '-'}<br>"
                detail_html += f"<b>تاریخ قرارداد:</b> {r[2]} | <b>تاریخ مراسم:</b> {r[3]}<br>"
                detail_html += f"<b>تخفیف:</b> {r[6]:,} تومان | <b>وضعیت:</b> {'✅ تسویه' if r[7] else '⏳ در جریان'}</p>"
                detail_html += f"<p><b>توضیحات:</b> {r[5] or '-'}</p>"
                items = r[4].split(',') if r[4] else []
                detail_html += "<p><b>خدمات انتخاب شده:</b></p><ul>"
                for it in items:
                    detail_html += f"<li>{it}</li>"
                detail_html += "</ul>"

                cursor.execute("SELECT amount, deposit_date, bank_name FROM wedding_deposits WHERE contract_id=?", (res["id"],))
                deps = cursor.fetchall()
                if deps:
                    detail_html += "<p><b>بیعانه‌ها:</b></p><ul>"
                    for d in deps:
                        detail_html += f"<li>{d[0]:,} تومان - {d[1]} - {d[2] or '-'}</li>"
                    detail_html += "</ul>"

        elif res["table"] == "commercial_projects":
            cursor.execute("""SELECT project_type, camera_count, description,
                                     client_name, client_phone, project_date
                                     FROM commercial_projects WHERE id=?""", (res["id"],))
            r = cursor.fetchone()
            if r:
                detail_html += f"<p><b>نوع:</b> {r[0]} | <b>دوربین:</b> {r[1]}<br>"
                detail_html += f"<b>مشتری:</b> {r[3] or '-'} | <b>تلفن:</b> {r[4] or '-'}<br>"
                detail_html += f"<b>تاریخ پروژه:</b> {r[5]}</p>"
                detail_html += f"<p><b>توضیحات:</b> {r[2] or '-'}</p>"

        elif res["table"] == "expenses":
            cursor.execute("SELECT category, description FROM expenses WHERE id=?", (res["id"],))
            r = cursor.fetchone()
            if r:
                detail_html += f"<p><b>دسته:</b> {r[0]}<br><b>توضیحات:</b> {r[1] or '-'}</p>"

        elif res["table"] == "checks":
            cursor.execute("SELECT check_type, description FROM checks WHERE id=?", (res["id"],))
            r = cursor.fetchone()
            if r:
                detail_html += f"<p><b>نوع چک:</b> {r[0] or 'دریافتی'}<br><b>توضیحات:</b> {r[1] or '-'}</p>"

        elif res["table"] == "persons":
            cursor.execute("""SELECT code, name, role, phone FROM persons WHERE id=?""", (res["id"],))
            r = cursor.fetchone()
            if r:
                detail_html += f"<p><b>کد پرسنل:</b> {r[0]}<br><b>نام:</b> {r[1]}<br>"
                detail_html += f"<b>نقش:</b> {r[2]}<br><b>تلفن:</b> {r[3] or '-'}</p>"
                cursor.execute("""SELECT SUM(amount) FROM transactions 
                                  WHERE person_id=? AND work_year=?""",
                               (res["id"], get_current_year()))
                total = cursor.fetchone()[0] or 0
                detail_html += f"<p><b>جمع پرداختی در این سال:</b> {total:,} تومان</p>"

        conn.close()

        detail_html += f"<hr><p><b>جمع کل:</b> {res['total']:,} تومان | "
        detail_html += f"<b>پرداختی:</b> {res['paid']:,} تومان | "
        remain = res['total'] - res['paid']
        color = '#c0392b' if remain > 0 else 'green'
        detail_html += f"<b style='color:{color};'>مانده:</b> {remain:,} تومان</p>"

        self.lbl_search_detail.setText(detail_html)

    # ============================================================
    # ==============  تب ۷: انبار تجهیزات  =====================
    # ============================================================
    def setup_inventory_tab(self):
        layout = QVBoxLayout()

        top = QHBoxLayout()
        btn_refresh = QPushButton("🔄 بروزرسانی")
        btn_refresh.setStyleSheet("background-color: #3498db; color: white; font-weight: bold; padding: 6px;")
        btn_refresh.clicked.connect(self.load_inventory)

        btn_print = QPushButton("🖨️ چاپ لیست انبار")
        btn_print.setStyleSheet("background-color: #2980b9; color: white; font-weight: bold; padding: 6px;")
        btn_print.clicked.connect(self.print_inventory)

        top.addWidget(btn_refresh)
        top.addWidget(btn_print)
        top.addStretch()
        layout.addLayout(top)

        self.inv_table = QTableWidget()
        self.inv_table.setColumnCount(5)
        self.inv_table.setHorizontalHeaderLabels([
            "کد کالا", "نام آیتم / پکیج", "تعداد کل",
            "استفاده شده", "موجودی"
        ])
        self.inv_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.inv_table)

        self.tab_inventory.setLayout(layout)
        self.load_inventory()

    def load_inventory(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("""SELECT i.item_name, ip.code, i.total_count, i.used_count
                          FROM inventory i
                          LEFT JOIN item_prices ip ON i.item_name = ip.item_name
                          ORDER BY i.item_name""")
        rows = cursor.fetchall()
        conn.close()

        self.inv_table.setRowCount(0)
        for r_idx, (name, code, total, used) in enumerate(rows):
            remain = total - used
            self.inv_table.insertRow(r_idx)
            self.inv_table.setItem(r_idx, 0, QTableWidgetItem(code or "-"))
            self.inv_table.setItem(r_idx, 1, QTableWidgetItem(name))
            self.inv_table.setItem(r_idx, 2, QTableWidgetItem(str(total)))
            self.inv_table.setItem(r_idx, 3, QTableWidgetItem(str(used)))
            remain_item = QTableWidgetItem(str(remain))
            if remain < 3:
                remain_item.setForeground(QColor("#c0392b"))
            self.inv_table.setItem(r_idx, 4, remain_item)

    def print_inventory(self):
        studio_name = get_setting("studio_name", "IMART STUDIO")
        now = jdatetime.datetime.now()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("""SELECT i.item_name, ip.code, i.total_count, i.used_count
                          FROM inventory i
                          LEFT JOIN item_prices ip ON i.item_name = ip.item_name
                          ORDER BY i.item_name""")
        rows = cursor.fetchall()
        conn.close()

        rows_html = ""
        for idx, (name, code, total, used) in enumerate(rows):
            remain = total - used
            rows_html += f"""<tr>
                <td style="border:1px solid #000; padding:6px; text-align:center;">{idx+1}</td>
                <td style="border:1px solid #000; padding:6px; text-align:center;">{code or '-'}</td>
                <td style="border:1px solid #000; padding:6px;">{name}</td>
                <td style="border:1px solid #000; padding:6px; text-align:center;">{total}</td>
                <td style="border:1px solid #000; padding:6px; text-align:center;">{used}</td>
                <td style="border:1px solid #000; padding:6px; text-align:center; color:{'red' if remain < 3 else 'black'};">{remain}</td>
            </tr>"""

        html = f"""
        <div dir="rtl" style="font-family: '{INVOICE_FONT_FAMILY}', Tahoma; padding: 20px;">
            <h1 style="text-align:center; color:#1F4E78; font-size:24pt; margin:0;">{studio_name}</h1>
            <h3 style="text-align:center; margin:5px 0;">لیست انبار تجهیزات</h3>
            <p style="text-align:left;">تاریخ چاپ: {now.strftime('%Y/%m/%d %H:%M')}</p>
            <hr style="border:2px solid #1F4E78;">
            <table style="width:100%; border-collapse:collapse;">
                <thead>
                    <tr style="background-color:#D9E1F2;">
                        <th style="border:1px solid #000; padding:6px;">ردیف</th>
                        <th style="border:1px solid #000; padding:6px;">کد</th>
                        <th style="border:1px solid #000; padding:6px;">نام</th>
                        <th style="border:1px solid #000; padding:6px;">تعداد کل</th>
                        <th style="border:1px solid #000; padding:6px;">استفاده شده</th>
                        <th style="border:1px solid #000; padding:6px;">موجودی</th>
                    </tr>
                </thead>
                <tbody>{rows_html}</tbody>
            </table>
        </div>
        """
        doc = QTextDocument()
        doc.setHtml(html)
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
        d = QPrintDialog(printer, self)
        if d.exec() == QPrintDialog.DialogCode.Accepted:
            doc.print_(printer)

    # ============================================================
    # ==============  تب ۸: کارت‌های بانکی  ====================
    # ============================================================
    def setup_banks_tab(self):
        layout = QHBoxLayout()
        form_box = QGroupBox("افزودن حساب بانکی")
        form_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        form_layout = QFormLayout()

        self.b_name = QLineEdit()
        self.b_card = QLineEdit()
        self.b_sheba = QLineEdit()

        form_layout.addRow("نام بانک <span style='color:red;'>*</span>:", self.b_name)
        form_layout.addRow("شماره کارت:", self.b_card)
        form_layout.addRow("شماره شبا:", self.b_sheba)

        btn_save = QPushButton("💾 ذخیره کارت")
        btn_save.setStyleSheet("background-color: #16a085; color: white; font-weight: bold; padding: 6px;")
        btn_save.clicked.connect(self.save_bank_card)
        form_layout.addRow(btn_save)

        form_box.setLayout(form_layout)
        layout.addWidget(form_box, 1)

        self.b_table = QTableWidget()
        self.b_table.setColumnCount(4)
        self.b_table.setHorizontalHeaderLabels(["نام بانک", "شماره کارت", "شماره شبا", "حذف"])
        self.b_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.b_table, 2)

        self.tab_banks.setLayout(layout)
        self.load_bank_cards()

    def save_bank_card(self):
        if not self.b_name.text().strip():
            QMessageBox.warning(self, "خطا", "نام بانک را وارد کنید.")
            return
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO bank_cards (bank_name, card_number, sheba_number) VALUES (?, ?, ?)",
                       (self.b_name.text(), self.b_card.text(), self.b_sheba.text()))
        conn.commit()
        conn.close()
        self.b_name.clear()
        self.b_card.clear()
        self.b_sheba.clear()
        self.load_bank_cards()
        self.load_bank_combo()

    def load_bank_cards(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, bank_name, card_number, sheba_number FROM bank_cards")
        rows = cursor.fetchall()
        conn.close()

        self.b_table.setRowCount(0)
        for r_idx, (bid, name, card, sheba) in enumerate(rows):
            self.b_table.insertRow(r_idx)
            self.b_table.setItem(r_idx, 0, QTableWidgetItem(name))
            self.b_table.setItem(r_idx, 1, QTableWidgetItem(str(card)))
            self.b_table.setItem(r_idx, 2, QTableWidgetItem(str(sheba)))
            btn_del = QPushButton("🗑️")
            btn_del.setStyleSheet("background-color: #e74c3c; color: white;")
            btn_del.clicked.connect(lambda _, x=bid: self.delete_bank_card(x))
            self.b_table.setCellWidget(r_idx, 3, btn_del)

    def delete_bank_card(self, bid):
        if not ask_security_password(self):
            return
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM bank_cards WHERE id=?", (bid,))
        conn.commit()
        conn.close()
        self.load_bank_cards()
        self.load_bank_combo()

    # ============================================================
    # ==============  تب ۹: مدیریت چک‌ها  =====================
    # ============================================================
    def setup_checks_tab(self):
        layout = QHBoxLayout()
        form_box = QGroupBox("ثبت چک جدید")
        form_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        form_layout = QFormLayout()

        self.chk_num = QLineEdit()
        self.chk_bank = QLineEdit()
        self.chk_issuer = QLineEdit()
        self.chk_type = QComboBox()
        self.chk_type.addItems(["دریافتی", "صادراتی"])
        self.chk_amount = QLineEdit()
        self.chk_amount.textChanged.connect(lambda t: self.chk_amount.setText(format_number(t)))
        self.chk_date = QLineEdit(jdatetime.date.today().strftime("%Y/%m/%d"))
        self.chk_desc = QLineEdit()

        form_layout.addRow("شماره چک <span style='color:red;'>*</span>:", self.chk_num)
        form_layout.addRow("بانک:", self.chk_bank)
        form_layout.addRow("نام صادرکننده:", self.chk_issuer)
        form_layout.addRow("نوع چک:", self.chk_type)
        form_layout.addRow("مبلغ (تومان) <span style='color:red;'>*</span>:", self.chk_amount)
        form_layout.addRow("تاریخ سررسید:", self.chk_date)
        form_layout.addRow("توضیحات:", self.chk_desc)

        btn_save = QPushButton("💾 ثبت چک")
        btn_save.setStyleSheet("background-color: #2980b9; color: white; font-weight: bold; padding: 6px;")
        btn_save.clicked.connect(self.save_check)
        form_layout.addRow(btn_save)

        form_box.setLayout(form_layout)
        layout.addWidget(form_box, 1)

        table_box = QGroupBox("لیست چک‌ها (با روزشمار)")
        table_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        tl = QVBoxLayout()

        self.chk_table = QTableWidget()
        self.chk_table.setColumnCount(10)
        self.chk_table.setHorizontalHeaderLabels([
            "شماره", "بانک", "صادرکننده", "نوع", "مبلغ",
            "سررسید", "روزشمار", "وضعیت", "چاپ", "حذف"
        ])
        self.chk_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        tl.addWidget(self.chk_table)
        table_box.setLayout(tl)
        layout.addWidget(table_box, 2)

        self.tab_checks.setLayout(layout)
        self.load_checks_table()

    def save_check(self):
        num = self.chk_num.text().strip()
        amount = parse_number(self.chk_amount.text())
        if not num or amount <= 0:
            QMessageBox.warning(self, "خطا", "اطلاعات چک را کامل کنید.")
            return

        work_year = get_current_year()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("""INSERT INTO checks 
            (check_number, bank_name, amount, due_date, description,
             issuer_name, check_type, work_year)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (num, self.chk_bank.text(), amount, self.chk_date.text(),
             self.chk_desc.text(), self.chk_issuer.text(),
             self.chk_type.currentText(), work_year))
        conn.commit()
        conn.close()

        self.chk_num.clear()
        self.chk_bank.clear()
        self.chk_issuer.clear()
        self.chk_amount.clear()
        self.chk_desc.clear()
        self.load_checks_table()

    def load_checks_table(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        work_year = get_current_year()
        cursor.execute("""SELECT id, check_number, bank_name, amount, due_date,
                          is_passed, issuer_name, check_type, description
                          FROM checks WHERE work_year=? ORDER BY due_date""",
                       (work_year,))
        rows = cursor.fetchall()
        conn.close()

        today = jdatetime.date.today()
        self.chk_table.setRowCount(0)

        for r_idx, row in enumerate(rows):
            (cid, num, bank, amount, due, is_passed,
             issuer, ctype, desc) = row

            try:
                parts = due.split('/')
                due_date = jdatetime.date(int(parts[0]), int(parts[1]), int(parts[2]))
                days_left = (due_date - today).days
                if is_passed:
                    countdown = "✅ پاس شده"
                elif days_left < 0:
                    countdown = f"🔴 {abs(days_left)} روز گذشته"
                elif days_left == 0:
                    countdown = "🟠 امروز!"
                elif days_left <= 3:
                    countdown = f"🟡 {days_left} روز مانده"
                else:
                    countdown = f"🟢 {days_left} روز مانده"
            except Exception:
                countdown = "-"

            self.chk_table.insertRow(r_idx)
            self.chk_table.setItem(r_idx, 0, QTableWidgetItem(num))
            self.chk_table.setItem(r_idx, 1, QTableWidgetItem(bank or "-"))
            self.chk_table.setItem(r_idx, 2, QTableWidgetItem(issuer or "-"))

            type_item = QTableWidgetItem(ctype or "دریافتی")
            if ctype == "دریافتی":
                type_item.setForeground(QColor("green"))
            else:
                type_item.setForeground(QColor("#c0392b"))
            self.chk_table.setItem(r_idx, 3, type_item)

            self.chk_table.setItem(r_idx, 4, QTableWidgetItem(f"{amount:,}"))
            self.chk_table.setItem(r_idx, 5, QTableWidgetItem(due))
            self.chk_table.setItem(r_idx, 6, QTableWidgetItem(countdown))

            btn_pass = QPushButton("✅ پاس شده" if is_passed else "⏳ پاس‌کردن")
            if is_passed:
                btn_pass.setStyleSheet("background-color: green; color: white;")
            else:
                btn_pass.setStyleSheet("background-color: orange; color: white;")
            btn_pass.clicked.connect(lambda _, x=cid, s=is_passed: self.toggle_check_pass(x, s))
            self.chk_table.setCellWidget(r_idx, 7, btn_pass)

            btn_print = QPushButton("🖨️")
            btn_print.setStyleSheet("background-color: #2980b9; color: white;")
            btn_print.clicked.connect(lambda _, x=cid: self.print_check(x))
            self.chk_table.setCellWidget(r_idx, 8, btn_print)

            btn_del = QPushButton("🗑️")
            btn_del.setStyleSheet("background-color: #e74c3c; color: white;")
            btn_del.clicked.connect(lambda _, x=cid: self.delete_check(x))
            self.chk_table.setCellWidget(r_idx, 9, btn_del)

    def toggle_check_pass(self, cid, current_state):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("UPDATE checks SET is_passed=? WHERE id=?", (0 if current_state else 1, cid))
        conn.commit()
        conn.close()
        self.load_checks_table()

    def print_check(self, cid):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("""SELECT check_number, bank_name, amount, due_date,
                          is_passed, issuer_name, check_type, description
                          FROM checks WHERE id=?""", (cid,))
        r = cursor.fetchone()
        conn.close()
        if not r:
            return

        studio_name = get_setting("studio_name", "IMART STUDIO")
        now = jdatetime.datetime.now()
        status = "پاس شده ✅" if r[4] else "پاس نشده ⏳"

        html = f"""
        <div dir="rtl" style="font-family: '{INVOICE_FONT_FAMILY}', Tahoma; padding: 20px;">
            <h1 style="text-align:center; color:#1F4E78; font-size:26pt; margin:0;">{studio_name}</h1>
            <h3 style="text-align:center; margin:5px 0;">رسید چک</h3>
            <hr style="border:2px solid #1F4E78;">
            <table style="width:100%; border-collapse:collapse; font-size:12pt; margin-top:20px;">
                <tr><td style="border:1px solid #000; padding:10px; width:35%; background-color:#F2F2F2;"><b>شماره چک:</b></td>
                    <td style="border:1px solid #000; padding:10px;">{r[0]}</td></tr>
                <tr><td style="border:1px solid #000; padding:10px; background-color:#F2F2F2;"><b>بانک:</b></td>
                    <td style="border:1px solid #000; padding:10px;">{r[1] or '-'}</td></tr>
                <tr><td style="border:1px solid #000; padding:10px; background-color:#F2F2F2;"><b>نام صادرکننده:</b></td>
                    <td style="border:1px solid #000; padding:10px;">{r[5] or '-'}</td></tr>
                <tr><td style="border:1px solid #000; padding:10px; background-color:#F2F2F2;"><b>نوع چک:</b></td>
                    <td style="border:1px solid #000; padding:10px;">{r[6] or 'دریافتی'}</td></tr>
                <tr><td style="border:1px solid #000; padding:10px; background-color:#F2F2F2;"><b>مبلغ:</b></td>
                    <td style="border:1px solid #000; padding:10px;"><b>{r[2]:,} تومان</b> ({number_to_persian_words(r[2])} تومان)</td></tr>
                <tr><td style="border:1px solid #000; padding:10px; background-color:#F2F2F2;"><b>تاریخ سررسید:</b></td>
                    <td style="border:1px solid #000; padding:10px;">{r[3]}</td></tr>
                <tr><td style="border:1px solid #000; padding:10px; background-color:#F2F2F2;"><b>وضعیت:</b></td>
                    <td style="border:1px solid #000; padding:10px;">{status}</td></tr>
                <tr><td style="border:1px solid #000; padding:10px; background-color:#F2F2F2;"><b>توضیحات:</b></td>
                    <td style="border:1px solid #000; padding:10px;">{r[7] or '-'}</td></tr>
            </table>
            <p style="text-align:left; margin-top:30px;">تاریخ چاپ: {now.strftime('%Y/%m/%d %H:%M')}</p>
            <p style="text-align:center; margin-top:50px;">مهر و امضای حسابداری</p>
        </div>
        """
        doc = QTextDocument()
        doc.setHtml(html)
        printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
        d = QPrintDialog(printer, self)
        if d.exec() == QPrintDialog.DialogCode.Accepted:
            doc.print_(printer)

    def delete_check(self, cid):
        if not ask_security_password(self):
            return
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM checks WHERE id=?", (cid,))
        conn.commit()
        conn.close()
        self.load_checks_table()

    # ============================================================
    # ==============  تب ۱۰: صدور رسید وجه  ====================
    # ============================================================
    def setup_receipt_tab(self):
        layout = QVBoxLayout()
        form_box = QGroupBox("صدور رسید وجه / بیعانه")
        form_box.setFont(QFont(APP_FONT_FAMILY, 11, QFont.Weight.Bold))
        form_layout = QFormLayout()

        self.rc_name = QLineEdit()
        self.rc_amount = QLineEdit()
        self.rc_amount.textChanged.connect(lambda t: self.rc_amount.setText(format_number(t)))
        self.rc_for = QLineEdit("خدمات فیلم و عکس")
        self.rc_date = QLineEdit(jdatetime.date.today().strftime("%Y/%m/%d"))

        form_layout.addRow("مبلغ دریافتی از آقا / خانم <span style='color:red;'>*</span>:", self.rc_name)
        form_layout.addRow("مبلغ (تومان) <span style='color:red;'>*</span>:", self.rc_amount)
        form_layout.addRow("بابت خدمات:", self.rc_for)
        form_layout.addRow("تاریخ:", self.rc_date)

        btn_row = QHBoxLayout()
        btn_print = QPushButton("🖨️ چاپ رسید")
        btn_print.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; padding: 10px; font-size: 11pt;")
        btn_print.clicked.connect(self.print_receipt)

        btn_pdf = QPushButton("📄 ذخیره PDF")
        btn_pdf.setStyleSheet("background-color: #c0392b; color: white; font-weight: bold; padding: 10px; font-size: 11pt;")
        btn_pdf.clicked.connect(self.export_receipt_pdf)

        btn_row.addWidget(btn_print)
        btn_row.addWidget(btn_pdf)
        form_layout.addRow(btn_row)

        form_box.setLayout(form_layout)
        layout.addWidget(form_box)

        preview_box = QGroupBox("پیش‌نمایش")
        preview_box.setFont(QFont(APP_FONT_FAMILY, 10, QFont.Weight.Bold))
        pv = QVBoxLayout()
        self.rc_preview = QTextEdit()
        self.rc_preview.setReadOnly(True)
        self.rc_preview.setStyleSheet("background-color: #f8f9fa; font-size: 11pt;")
        self.rc_preview.setFixedHeight(300)
        pv.addWidget(self.rc_preview)
        preview_box.setLayout(pv)
        layout.addWidget(preview_box)

        layout.addStretch()
        self.tab_receipt.setLayout(layout)

    def get_receipt_html(self):
        name = self.rc_name.text().strip()
        amount = parse_number(self.rc_amount.text())
        if not name or amount <= 0:
            return None
        return InvoiceBuilder.build_receipt(
            name, amount, self.rc_for.text(), self.rc_date.text()
        )

    def print_receipt(self):
        html = self.get_receipt_html()
        if not html:
            QMessageBox.warning(self, "خطا", "لطفاً نام پرداخت‌کننده و مبلغ را وارد کنید.")
            return
        try:
            doc = QTextDocument()
            doc.setHtml(html)
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
            d = QPrintDialog(printer, self)
            if d.exec() == QPrintDialog.DialogCode.Accepted:
                doc.print_(printer)
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا:\n{e}")

    def export_receipt_pdf(self):
        html = self.get_receipt_html()
        if not html:
            QMessageBox.warning(self, "خطا", "لطفاً نام پرداخت‌کننده و مبلغ را وارد کنید.")
            return
        name = self.rc_name.text().strip()
        file_path, _ = QFileDialog.getSaveFileName(
            self, "ذخیره رسید PDF",
            f"رسید_{name}_{jdatetime.date.today().strftime('%Y%m%d')}.pdf",
            "PDF Files (*.pdf)"
        )
        if not file_path:
            return
        try:
            doc = QTextDocument()
            doc.setHtml(html)
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
            printer.setOutputFileName(file_path)
            printer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
            doc.print_(printer)
            QMessageBox.information(self, "موفقیت", "رسید PDF ذخیره شد.")
        except Exception as e:
            QMessageBox.critical(self, "خطا", f"خطا:\n{e}")

    # ============================================================
    # ==============  تب ۱۱: تنظیمات و سال کاری  ===============
    # ============================================================
    def setup_workyear_tab(self):
        layout = QVBoxLayout()

        year_box = QGroupBox("مدیریت سال‌های کاری")
        year_box.setFont(QFont(APP_FONT_FAMILY, 11, QFont.Weight.Bold))
        yb = QVBoxLayout()

        year_row = QHBoxLayout()
        year_row.addWidget(QLabel("سال کاری فعلی:"))
        self.combo_settings_year = QComboBox()
        self.combo_settings_year.setFont(QFont(APP_FONT_FAMILY, 11, QFont.Weight.Bold))
        self.combo_settings_year.setFixedWidth(120)
        year_row.addWidget(self.combo_settings_year)

        btn_switch = QPushButton("🔀 تغییر سال")
        btn_switch.setStyleSheet("background-color: #2980b9; color: white; font-weight: bold; padding: 6px;")
        btn_switch.clicked.connect(self.switch_year_from_settings)
        year_row.addWidget(btn_switch)

        btn_new_year = QPushButton("➕ شروع سال کاری جدید")
        btn_new_year.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; padding: 6px;")
        btn_new_year.clicked.connect(self.start_new_work_year)
        year_row.addWidget(btn_new_year)

        year_row.addStretch()
        yb.addLayout(year_row)

        info = QLabel(
            "💡 با تغییر سال کاری، تمام تب‌ها به طور خودکار فقط اطلاعات آن سال را نمایش می‌دهند. "
            "شروع سال جدید به معنی پاک کردن اطلاعات سال قبل نیست — همه سال‌ها محفوظ می‌مانند."
        )
        info.setWordWrap(True)
        info.setStyleSheet("background-color: #e8f8f5; padding: 10px; border-radius: 5px;")
        yb.addWidget(info)

        year_box.setLayout(yb)
        layout.addWidget(year_box)

        close_box = QGroupBox("بستن سال کاری (آرشیو)")
        close_box.setFont(QFont(APP_FONT_FAMILY, 11, QFont.Weight.Bold))
        cb = QVBoxLayout()

        warn = QLabel(
            "⚠️ <b>اخطار:</b> بستن سال کاری باعث خالی شدن جداول قراردادها، پروژه‌ها، هزینه‌ها و "
            "پرداختی‌ها می‌شود. از دیتابیس بک‌آپ گرفته می‌شود ولی این عمل بازگشت‌پذیر نیست."
        )
        warn.setWordWrap(True)
        warn.setStyleSheet("background-color: #fdedec; padding: 10px; border-radius: 5px; border: 1px solid #e74c3c;")
        cb.addWidget(warn)

        btn_close_year = QPushButton("🔒 بستن سال کاری (آرشیو + خالی کردن)")
        btn_close_year.setStyleSheet("background-color: #c0392b; color: white; font-weight: bold; padding: 12px; font-size: 11pt;")
        btn_close_year.clicked.connect(self.close_work_year)
        cb.addWidget(btn_close_year)

        close_box.setLayout(cb)
        layout.addWidget(close_box)

        settings_box = QGroupBox("تنظیمات عمومی")
        settings_box.setFont(QFont(APP_FONT_FAMILY, 11, QFont.Weight.Bold))
        sb = QVBoxLayout()

        info2 = QLabel("📌 برای تغییر عنوان‌های پیش‌فرض استودیو از دکمه بالای صفحه استفاده کنید.")
        info2.setStyleSheet("padding: 8px;")
        sb.addWidget(info2)

        settings_box.setLayout(sb)
        layout.addWidget(settings_box)

        layout.addStretch()
        self.tab_workyear.setLayout(layout)
        self.load_settings_years()

    def load_settings_years(self):
        if not hasattr(self, 'combo_settings_year'):
            return
        self.combo_settings_year.blockSignals(True)
        self.combo_settings_year.clear()
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        years = set()
        years.add(jdatetime.date.today().year)
        try:
            years.add(get_current_year())
        except Exception:
            pass

        for table in ["wedding_contracts", "commercial_projects", "expenses",
                      "transactions", "checks", "persons"]:
            try:
                cursor.execute(f"SELECT DISTINCT work_year FROM {table} WHERE work_year IS NOT NULL")
                for (y,) in cursor.fetchall():
                    if y:
                        years.add(int(y))
            except Exception:
                pass

        conn.close()

        for y in sorted(years, reverse=True):
            self.combo_settings_year.addItem(str(y))

        current = get_current_year()
        idx = self.combo_settings_year.findText(str(current))
        if idx >= 0:
            self.combo_settings_year.setCurrentIndex(idx)
        self.combo_settings_year.blockSignals(False)

    def switch_year_from_settings(self):
        try:
            new_year = int(self.combo_settings_year.currentText())
        except Exception:
            return
        set_setting("current_work_year", new_year)

        if hasattr(self, 'w_table'):
            self.load_wedding_contracts()
            self.load_commercial_projects()
            self.load_expenses_table()
            self.load_staff_table()
            self.load_inventory()
            self.load_checks_table()

        if hasattr(self, 'combo_work_year'):
            self.load_work_years()

        QMessageBox.information(self, "تغییر سال", f"سال کاری به {new_year} تغییر یافت.")

    def start_new_work_year(self):
        if not ask_security_password(self):
            return
        new_year, ok = QInputDialog.getInt(
            self, "سال کاری جدید", "سال جدید را وارد کنید:",
            get_current_year() + 1, 1300, 1500
        )
        if not ok:
            return
        set_setting("current_work_year", new_year)
        self.load_work_years()
        self.load_settings_years()
        if hasattr(self, 'w_table'):
            self.load_wedding_contracts()
            self.load_commercial_projects()
            self.load_expenses_table()
            self.load_staff_table()
            self.load_inventory()
            self.load_checks_table()
        QMessageBox.information(self, "موفقیت", f"سال کاری جدید {new_year} شروع شد. تمام داده‌ها محفوظ هستند.")

    def close_work_year(self):
        if not ask_security_password(self):
            return

        reply = QMessageBox.question(
            self, "تایید نهایی",
            "آیا مطمئن هستید که می‌خواهید سال کاری فعلی را ببندید؟\n\n"
            "⚠️ اطلاعات جاری پاک می‌شود (ولی بک‌آپ گرفته می‌شود).",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            today_str = jdatetime.date.today().strftime("%Y_%m_%d")
            work_year = get_current_year()
            backup_name = f"studio_archive_{work_year}_{today_str}.db"
            shutil.copy(DB_NAME, backup_name)

            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM wedding_contracts WHERE work_year=?", (work_year,))
            cursor.execute("DELETE FROM wedding_deposits WHERE work_year=?", (work_year,))
            cursor.execute("DELETE FROM commercial_projects WHERE work_year=?", (work_year,))
            cursor.execute("DELETE FROM expenses WHERE work_year=?", (work_year,))
            cursor.execute("DELETE FROM transactions WHERE work_year=?", (work_year,))
            cursor.execute("DELETE FROM checks WHERE work_year=?", (work_year,))
            cursor.execute("UPDATE inventory SET used_count=0")
            conn.commit()
            conn.close()

            QMessageBox.information(
                self, "موفقیت",
                f"سال کاری {work_year} بسته شد.\n\nفایل آرشیو: {backup_name}"
            )
            if hasattr(self, 'w_table'):
                self.load_wedding_contracts()
                self.load_commercial_projects()
                self.load_expenses_table()
                self.load_staff_table()
                self.load_inventory()
                self.load_checks_table()

    # ============================================================
    # ==============  تب ۱۲: درباره برنامه  ====================
    # ============================================================
    def setup_about_tab(self):
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        card = QGroupBox("درباره نرم‌افزار IMART STUDIO")
        card.setFont(QFont(APP_FONT_FAMILY, 12, QFont.Weight.Bold))
        card_layout = QVBoxLayout()
        card_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        lbl_title = QLabel("🎬 IMART STUDIO")
        lbl_title.setFont(QFont(APP_FONT_FAMILY, 22, QFont.Weight.Bold))
        lbl_title.setStyleSheet("color: #1F4E78;")
        lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)

        lbl_version = QLabel(f"نسخه {APP_VERSION}")
        lbl_version.setFont(QFont(APP_FONT_FAMILY, 12, QFont.Weight.Bold))
        lbl_version.setStyleSheet("color: #7f8c8d;")
        lbl_version.setAlignment(Qt.AlignmentFlag.AlignCenter)

        lbl_desc = QLabel(
            "نرم‌افزار جامع مدیریت مالی، حسابداری و پرسنلی\n"
            "مخصوص آتلیه‌ها، استودیوهای فیلمبرداری و پروژه‌های تولید محتوا\n\n"
            f"👨‍💻 طراح و توسعه‌دهنده: {DEVELOPER_NAME}\n"
            "📞 شماره تماس پشتیبانی: 09173736618\n\n"
            "تمامی حقوق این نرم‌افزار محفوظ می‌باشد."
        )
        lbl_desc.setFont(QFont(APP_FONT_FAMILY, 11))
        lbl_desc.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_desc.setStyleSheet("line-height: 2;")

        card_layout.addWidget(lbl_title)
        card_layout.addSpacing(10)
        card_layout.addWidget(lbl_version)
        card_layout.addSpacing(20)
        card_layout.addWidget(lbl_desc)

        card.setLayout(card_layout)
        card.setFixedWidth(600)

        layout.addWidget(card, alignment=Qt.AlignmentFlag.AlignCenter)
        self.tab_about.setLayout(layout)

    # ============================================================
    # ==============  متدهای عمومی  ============================
    # ============================================================
    def backup_db(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "ذخیره فایل پشتیبان",
            f"backup_{jdatetime.date.today().strftime('%Y_%m_%d')}.db",
            "Database Files (*.db)"
        )
        if file_path:
            shutil.copy(DB_NAME, file_path)
            QMessageBox.information(self, "موفقیت", "پشتیبان‌گیری با موفقیت انجام شد.")

    def restore_db(self):
        if not ask_security_password(self):
            return
        file_path, _ = QFileDialog.getOpenFileName(
            self, "انتخاب فایل پشتیبان", "", "Database Files (*.db)"
        )
        if file_path:
            shutil.copy(file_path, DB_NAME)
            QMessageBox.information(
                self, "موفقیت",
                "پایگاه داده با موفقیت بازیابی شد. برنامه را دوباره اجرا کنید."
            )

    def change_password(self):
        if not ask_security_password(self):
            return
        new_pass, ok = QInputDialog.getText(
            self, "تغییر رمز عبور", "رمز عبور جدید را وارد کنید:",
            QLineEdit.EchoMode.Password
        )
        if ok and new_pass.strip():
            set_setting("app_password", new_pass.strip())
            QMessageBox.information(self, "موفقیت", "رمز عبور جدید با موفقیت ثبت گردید.")

    def closeEvent(self, event):
        """بک‌آپ اتوماتیک قبل از خروج"""
        try:
            if os.path.exists(DB_NAME):
                backup_dir = "auto_backups"
                if not os.path.exists(backup_dir):
                    os.makedirs(backup_dir)
                ts = jdatetime.datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
                backup_path = os.path.join(backup_dir, f"auto_backup_{ts}.db")
                shutil.copy(DB_NAME, backup_path)

                backups = sorted([
                    f for f in os.listdir(backup_dir) if f.startswith("auto_backup_")
                ])
                if len(backups) > 10:
                    for old in backups[:-10]:
                        try:
                            os.remove(os.path.join(backup_dir, old))
                        except Exception:
                            pass
        except Exception as e:
            print("Auto backup error:", e)

        event.accept()


# ============================================================
# ====================  نقطه ورود  ==========================
# ============================================================
if __name__ == '__main__':
    app = QApplication(sys.argv)

    APP_FONT_FAMILY = setup_fonts()
    app.setFont(QFont(APP_FONT_FAMILY, 10))

    # آماده‌سازی فونت نمودار
    setup_matplotlib_font()

    loading = LoadingScreen()
    loading.show()

    steps = [
        ("راه‌اندازی موتور برنامه...", 5),
        ("بارگذاری فونت‌های فارسی...", 12),
        ("اتصال به پایگاه داده SQLite...", 22),
        ("ایجاد جداول و کدهای کالا...", 35),
        ("بارگذاری تنظیمات امنیتی...", 45),
        ("آماده‌سازی داشبورد...", 55),
        ("بارگذاری تب‌های برنامه...", 70),
        ("بارگذاری قراردادها و پروژه‌ها...", 82),
        ("بارگذاری انبار، بانک و چک‌ها...", 92),
        ("آماده‌سازی نهایی...", 98),
        ("ورود به سیستم...", 100),
    ]

    main_win = StudioAccountingApp()
    current_step = {"index": 0}

    def finish_loading():
        loading.close()
        if main_win.prompt_login():
            main_win.stack.setCurrentWidget(main_win.dashboard_screen)
            main_win.show()
            QTimer.singleShot(2000, main_win.check_check_alerts)
        else:
            app.quit()

    def run_step():
        if current_step["index"] < len(steps):
            text, percent = steps[current_step["index"]]
            loading.set_progress(percent, text)

            if current_step["index"] == 2:
                init_db()
            elif current_step["index"] == 5:
                main_win.setup_dashboard_ui()
            elif current_step["index"] == 6:
                main_win.setup_main_app_ui()

            current_step["index"] += 1
            QTimer.singleShot(250, run_step)
        else:
            QTimer.singleShot(400, finish_loading)

    QTimer.singleShot(400, run_step)

    sys.exit(app.exec())