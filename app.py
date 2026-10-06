import sys
import os
import shutil
import sqlite3
import jdatetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QComboBox, QPushButton, QTableWidget, QTableWidgetItem,
    QHeaderView, QMessageBox, QGroupBox, QSpinBox, QFormLayout, QFileDialog,
    QTabWidget, QCheckBox, QInputDialog, QDialog, QStackedWidget, QFrame
)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont, QIcon

import openpyxl
from openpyxl.styles import Font as XLFont, PatternFill, Alignment
import matplotlib.pyplot as plt

from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

# تنظیم ID برای شناسایی آیکون روی دسکتاپ و تسک‌بار ویندوز
try:
    import ctypes
    myappid = 'imartstudio.accounting.v4.0'
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass

DB_NAME = "studio_accounting.db"

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

def init_db():
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    ''')
    cursor.execute("INSERT OR IGNORE INTO settings (key, value) VALUES ('app_password', '123')")
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS persons (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            role TEXT NOT NULL
        )
    ''')
    
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

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS bank_cards (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bank_name TEXT NOT NULL,
            card_number TEXT,
            sheba_number TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS item_prices (
            item_name TEXT PRIMARY KEY,
            price INTEGER DEFAULT 0
        )
    ''')

    default_items = [
        "کلیپ فرمالیته", "کلیپ باغ", "کلیپ پارسیان", "شمال", "خارج از کشور",
        "یک دوربین", "دو دوربین", "کرین", "هلی شات", "FPV", "عکاس مجلس"
    ]
    for item in default_items:
        cursor.execute("INSERT OR IGNORE INTO item_prices (item_name, price) VALUES (?, 0)", (item,))

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS wedding_contracts (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            groom_name TEXT,
            bride_name TEXT,
            groom_phone TEXT,
            bride_phone TEXT,
            contract_date TEXT,
            ceremony_date TEXT,
            selected_items TEXT,
            total_amount INTEGER,
            discount INTEGER,
            paid_amount INTEGER,
            bank_id INTEGER,
            is_settled INTEGER DEFAULT 0
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS wedding_deposits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            contract_id INTEGER,
            amount INTEGER,
            deposit_date TEXT,
            bank_name TEXT,
            FOREIGN KEY (contract_id) REFERENCES wedding_contracts(id)
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS commercial_projects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT,
            project_type TEXT,
            camera_count INTEGER,
            total_amount INTEGER,
            paid_amount INTEGER,
            project_date TEXT,
            description TEXT
        )
    ''')

    cursor.execute('''
        CREATE TABLE IF NOT EXISTS inventory (
            item_name TEXT PRIMARY KEY,
            total_count INTEGER DEFAULT 0,
            used_count INTEGER DEFAULT 0
        )
    ''')
    for item in default_items:
        cursor.execute("INSERT OR IGNORE INTO inventory (item_name, total_count, used_count) VALUES (?, 10, 0)", (item,))

    conn.commit()
    conn.close()

class ManageItemsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("مدیریت و افزودن موارد فاکتور")
        self.resize(450, 400)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.setFont(QFont("B Yekan", 10))

        layout = QVBoxLayout()
        
        form_layout = QFormLayout()
        self.txt_item_name = QLineEdit()
        self.txt_item_price = QLineEdit()
        self.txt_item_price.textChanged.connect(self.on_price_changed)

        form_layout.addRow("نام مورد جدید:", self.txt_item_name)
        form_layout.addRow("قیمت اولیه (تومان):", self.txt_item_price)
        
        btn_add = QPushButton("افزودن / بروزرسانی")
        btn_add.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; padding: 5px;")
        btn_add.clicked.connect(self.add_or_update_item)
        form_layout.addRow(btn_add)
        
        layout.addLayout(form_layout)

        self.table = QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["نام مورد", "قیمت (تومان)"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        layout.addWidget(self.table)

        self.load_items()
        self.setLayout(layout)

    def on_price_changed(self, text):
        formatted = format_number(text)
        if formatted != text:
            self.txt_item_price.setText(formatted)

    def load_items(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT item_name, price FROM item_prices")
        rows = cursor.fetchall()
        conn.close()

        self.table.setRowCount(0)
        for r_idx, (name, price) in enumerate(rows):
            self.table.insertRow(r_idx)
            self.table.setItem(r_idx, 0, QTableWidgetItem(name))
            self.table.setItem(r_idx, 1, QTableWidgetItem(f"{price:,}"))

    def add_or_update_item(self):
        name = self.txt_item_name.text().strip()
        price = parse_number(self.txt_item_price.text())
        if not name:
            QMessageBox.warning(self, "خطا", "لطفاً نام مورد را وارد کنید.")
            return

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO item_prices (item_name, price) VALUES (?, ?) ON CONFLICT(item_name) DO UPDATE SET price=excluded.price", (name, price))
        cursor.execute("INSERT OR IGNORE INTO inventory (item_name, total_count, used_count) VALUES (?, 10, 0)", (name,))
        conn.commit()
        conn.close()

        self.txt_item_name.clear()
        self.txt_item_price.clear()
        self.load_items()

class DepositsDialog(QDialog):
    def __init__(self, contract_id, parent=None):
        super().__init__(parent)
        self.contract_id = contract_id
        self.setWindowTitle(f"مدیریت بیعانه‌ها و پرداخت‌های قرارداد #{contract_id}")
        self.resize(500, 350)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.setFont(QFont("B Yekan", 10))

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
        self.table.setHorizontalHeaderLabels(["ID", "مبلغ (تومان)", "تاریخ", "بانک"])
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
        cursor.execute("INSERT INTO wedding_deposits (contract_id, amount, deposit_date, bank_name) VALUES (?, ?, ?, ?)",
                       (self.contract_id, amount, d_date, bank))
        
        cursor.execute("SELECT SUM(amount) FROM wedding_deposits WHERE contract_id=?", (self.contract_id,))
        total_paid = cursor.fetchone()[0] or 0
        
        cursor.execute("SELECT total_amount, discount FROM wedding_contracts WHERE id=?", (self.contract_id,))
        tot, disc = cursor.fetchone()
        
        is_settled = 1 if (tot - disc - total_paid) <= 0 else 0
        cursor.execute("UPDATE wedding_contracts SET paid_amount=?, is_settled=? WHERE id=?", (total_paid, is_settled, self.contract_id))

        conn.commit()
        conn.close()

        self.txt_amount.clear()
        self.load_deposits()
        QMessageBox.information(self, "موفقیت", "بیعانه با موفقیت ثبت شد.")

class StudioAccountingApp(QMainWindow):
    ROLES = ["تدوینگر", "عکاس", "فیلمبردار", "هلی شات و FPV کار", "اوپراتور کرین"]
    PROJECT_TYPES = ["عروسی", "عقد", "تولد", "تبلیغاتی", "قبض و کرایه"]
    PERSIAN_MONTHS = [
        "فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
        "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند"
    ]

    def __init__(self):
        super().__init__()
        self.setWindowTitle("نرم‌افزار مدیریت مالی - IMART STUDIO v4.0")
        self.resize(1250, 850)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)

        icon_path = resource_path("Accounting.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.setFont(QFont("B Yekan", 10))
        init_db()

        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        self.dashboard_screen = QWidget()
        self.stack.addWidget(self.dashboard_screen)

        self.main_app_screen = QWidget()
        self.stack.addWidget(self.main_app_screen)

        self.setup_dashboard_ui()
        self.setup_main_app_ui()
        self.stack.setCurrentWidget(self.dashboard_screen)

    def prompt_login(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM settings WHERE key='app_password'")
        saved_pass = cursor.fetchone()[0]
        conn.close()

        entered_pass, ok = QInputDialog.getText(
            self, "ورود به نرم‌افزار IMART STUDIO", "لطفاً رمز عبور را وارد کنید:", QLineEdit.EchoMode.Password
        )

        if not (ok and entered_pass == saved_pass):
            QMessageBox.critical(self, "خطا", "رمز عبور اشتباه است!")
            sys.exit()

    def setup_dashboard_ui(self):
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)

        title = QLabel("سیستم جامع مدیریت مالی و حسابداری IMART STUDIO")
        title.setFont(QFont("B Yekan", 16, QFont.Weight.Bold))
        title.setStyleSheet("color: #2c3e50; margin-bottom: 30px;")
        layout.addWidget(title, alignment=Qt.AlignmentFlag.AlignCenter)

        grid_layout = QHBoxLayout()

        btn_w = QPushButton("پروژه‌های عروس و داماد\n🔒 (ثبت قراردادها)")
        btn_c = QPushButton("تبلیغاتی / بیوتی / تولدی\n🔒 (پروژه‌های استودیو)")
        btn_s = QPushButton("بخش کارکنان و هزینه‌ها\n🔒 (حسابداری جامع)")
        btn_i = QPushButton("انبار تجهیزات\n(موجودی و وضعیت)")
        btn_b = QPushButton("کارت‌های بانکی\n(مدیریت حساب‌ها)")

        buttons = [
            (btn_w, 0, "#2980b9"),
            (btn_c, 1, "#27ae60"),
            (btn_s, 2, "#8e44ad"),
            (btn_i, 3, "#d35400"),
            (btn_b, 4, "#16a085")
        ]

        for btn, idx, color in buttons:
            btn.setFixedSize(220, 140)
            btn.setFont(QFont("B Yekan", 12, QFont.Weight.Bold))
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {color};
                    color: white;
                    border-radius: 12px;
                    padding: 10px;
                }}
                QPushButton:hover {{
                    background-color: #34495e;
                }}
            """)
            btn.clicked.connect(lambda _, i=idx: self.open_tab_index(i))
            grid_layout.addWidget(btn)

        layout.addLayout(grid_layout)
        self.dashboard_screen.setLayout(layout)

    def open_tab_index(self, index):
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
        self.tabs.setFont(QFont("B Yekan", 11, QFont.Weight.Bold))

        self.tab_wedding = QWidget()
        self.tab_commercial = QWidget()
        self.tab_staff = QWidget()
        self.tab_inventory = QWidget()
        self.tab_banks = QWidget()

        self.tabs.addTab(self.tab_wedding, "پروژه‌های عروس و داماد")
        self.tabs.addTab(self.tab_commercial, "تبلیغاتی / بیوتی / تولدی")
        self.tabs.addTab(self.tab_staff, "بخش کارکنان و هزینه‌ها")
        self.tabs.addTab(self.tab_inventory, "انبار تجهیزات")
        self.tabs.addTab(self.tab_banks, "کارت‌های بانکی")

        main_layout.addWidget(self.tabs)

        self.setup_wedding_tab()
        self.setup_commercial_tab()
        self.setup_staff_tab()
        self.setup_inventory_tab()
        self.setup_banks_tab()

        footer = QLabel("IMART STUDIO - Phone: 09173736618")
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer.setFont(QFont("B Nazanin", 11, QFont.Weight.Bold))
        footer.setStyleSheet("color: #2c3e50; margin-top: 5px;")
        main_layout.addWidget(footer)

        self.main_app_screen.setLayout(main_layout)

    # --- زبانه ۱: عروس و داماد ---
    def setup_wedding_tab(self):
        layout = QHBoxLayout()
        form_box = QGroupBox("ثبت قرارداد جدید عروس و داماد")
        form_box.setFont(QFont("B Yekan", 10, QFont.Weight.Bold))
        form_layout = QFormLayout()

        self.w_groom = QLineEdit()
        self.w_bride = QLineEdit()
        self.w_groom_phone = QLineEdit()
        self.w_bride_phone = QLineEdit()
        self.w_contract_date = QLineEdit(jdatetime.date.today().strftime("%Y/%m/%d"))
        self.w_ceremony_date = QLineEdit(jdatetime.date.today().strftime("%Y/%m/%d"))

        form_layout.addRow("نام داماد:", self.w_groom)
        form_layout.addRow("نام عروس:", self.w_bride)
        form_layout.addRow("تلفن داماد:", self.w_groom_phone)
        form_layout.addRow("تلفن عروس:", self.w_bride_phone)
        form_layout.addRow("تاریخ قرارداد:", self.w_contract_date)
        form_layout.addRow("تاریخ مراسم:", self.w_ceremony_date)

        btn_manage_items = QPushButton("مدیریت / افزودن موارد فاکتور")
        btn_manage_items.setStyleSheet("background-color: #e67e22; color: white; font-weight: bold;")
        btn_manage_items.clicked.connect(self.open_manage_items)
        form_layout.addRow(btn_manage_items)

        self.items_container = QWidget()
        self.items_vbox = QVBoxLayout()
        self.items_container.setLayout(self.items_vbox)
        
        items_box = QGroupBox("موارد فاکتور (تیک بزنید)")
        items_box_layout = QVBoxLayout()
        items_box_layout.addWidget(self.items_container)
        items_box.setLayout(items_box_layout)
        form_layout.addRow(items_box)

        self.load_item_checkboxes()

        self.w_lbl_total = QLabel("جمع کل: ۰ تومان")
        self.w_lbl_total.setFont(QFont("B Yekan", 10, QFont.Weight.Bold))
        self.w_lbl_total.setStyleSheet("color: #2c3e50;")
        form_layout.addRow(self.w_lbl_total)

        self.w_discount = QLineEdit("0")
        self.w_discount.textChanged.connect(lambda t: self.on_discount_changed(t))

        self.w_first_deposit = QLineEdit("0")
        self.w_first_deposit.textChanged.connect(lambda t: self.on_deposit_changed(t))

        form_layout.addRow("تخفیف (تومان):", self.w_discount)
        form_layout.addRow("بیعانه اول (تومان):", self.w_first_deposit)

        self.w_bank_combo = QComboBox()
        self.load_bank_combo()
        form_layout.addRow("بانک واریزی بیعانه اول:", self.w_bank_combo)

        self.w_lbl_remain = QLabel("باقی‌مانده حساب: ۰ تومان")
        self.w_lbl_remain.setFont(QFont("B Yekan", 10, QFont.Weight.Bold))
        self.w_lbl_remain.setStyleSheet("color: #c0392b;")
        form_layout.addRow(self.w_lbl_remain)

        btn_save = QPushButton("ثبت نهایی قرارداد")
        btn_save.setFont(QFont("B Yekan", 10, QFont.Weight.Bold))
        btn_save.setStyleSheet("background-color: #2980b9; color: white; padding: 8px;")
        btn_save.clicked.connect(self.save_wedding_contract)
        form_layout.addRow(btn_save)

        form_box.setLayout(form_layout)
        layout.addWidget(form_box, 1)

        table_box = QGroupBox("لیست قراردادها")
        table_box.setFont(QFont("B Yekan", 10, QFont.Weight.Bold))
        table_layout = QVBoxLayout()

        self.w_table = QTableWidget()
        self.w_table.setColumnCount(11)
        self.w_table.setHorizontalHeaderLabels([
            "ID", "زوجین", "تلفن داماد", "تلفن عروس", "تاریخ مراسم", "جمع کل", "تخفیف", "کل دریافتی", "مانده", "بیعانه‌ها", "PDF / پرینت"
        ])
        self.w_table.setWordWrap(True)
        self.w_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
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
            self.items_vbox.itemAt(i).widget().setParent(None)

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
            cb.setFont(QFont("B Nazanin", 10))
            cb.stateChanged.connect(self.calc_wedding_total)

            txt_p = QLineEdit(f"{price:,}")
            txt_p.setFixedWidth(100)
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
        selected_sum = 0
        for item_name, cb in self.item_checkboxes.items():
            if cb.isChecked():
                price = parse_number(self.item_price_inputs[item_name].text())
                selected_sum += price

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
            QMessageBox.warning(self, "خطا", "نام عروس و داماد را وارد کنید.")
            return

        selected_items = [item for item, cb in self.item_checkboxes.items() if cb.isChecked()]
        if not selected_items:
            QMessageBox.warning(self, "خطا", "تا قبل از ورود و انتخاب موارد فاکتور اجازه ثبت داده نمی‌شود.")
            return

        selected_sum = sum(parse_number(self.item_price_inputs[i].text()) for i in selected_items)
        discount = parse_number(self.w_discount.text())
        total_amount = max(0, selected_sum - discount)

        first_deposit = parse_number(self.w_first_deposit.text())
        bank_id = self.w_bank_combo.currentData()
        bank_name = self.w_bank_combo.currentText()

        remain = total_amount - first_deposit
        is_settled = 1 if remain <= 0 else 0

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        cursor.execute('''
            INSERT INTO wedding_contracts 
            (groom_name, bride_name, groom_phone, bride_phone, contract_date, ceremony_date, selected_items, total_amount, discount, paid_amount, bank_id, is_settled)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (groom, bride, self.w_groom_phone.text(), self.w_bride_phone.text(),
              self.w_contract_date.text(), self.w_ceremony_date.text(),
              ",".join(selected_items), total_amount, discount, first_deposit, bank_id, is_settled))

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

            btn_dep = QPushButton("مدیریت بیعانه‌ها")
            btn_dep.clicked.connect(lambda _, cid=c_id: self.open_deposits_dialog(cid))
            self.w_table.setCellWidget(r_idx, 9, btn_dep)

            btn_pdf = QPushButton("چاپ PDF")
            btn_pdf.clicked.connect(lambda _, cid=c_id: self.export_wedding_pdf(cid))
            self.w_table.setCellWidget(r_idx, 10, btn_pdf)

        self.w_table.resizeRowsToContents()

    def open_deposits_dialog(self, contract_id):
        dlg = DepositsDialog(contract_id, self)
        dlg.exec()
        self.load_wedding_contracts()

    def export_wedding_pdf(self, contract_id):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT groom_name, bride_name, groom_phone, bride_phone, contract_date, ceremony_date, selected_items, total_amount, discount, paid_amount FROM wedding_contracts WHERE id=?", (contract_id,))
        c = cursor.fetchone()
        conn.close()

        if not c: return
        file_path, _ = QFileDialog.getSaveFileName(self, "ذخیره فاکتور PDF", f"Factore_Wedding_{contract_id}.pdf", "PDF Files (*.pdf)")
        if not file_path: return

        pdf = canvas.Canvas(file_path, pagesize=letter)
        pdf.setFont("Helvetica-Bold", 18)
        pdf.drawString(220, 750, "IMART STUDIO")
        pdf.setFont("Helvetica", 11)
        pdf.drawString(225, 732, "Tel: 09173736618")
        pdf.setFont("Helvetica", 10)
        pdf.drawString(200, 715, f"Date: {jdatetime.date.today().strftime('%Y/%m/%d')}")
        pdf.line(50, 700, 550, 700)

        pdf.setFont("Helvetica", 12)
        pdf.drawString(50, 670, f"Groom & Bride: {c[0]} & {c[1]}")
        pdf.drawString(50, 650, f"Groom Phone: {c[2]} | Bride Phone: {c[3]}")
        pdf.drawString(50, 630, f"Contract Date: {c[4]} | Ceremony Date: {c[5]}")
        pdf.drawString(50, 610, f"Selected Services: {c[6]}")
        pdf.line(50, 590, 550, 590)

        pdf.drawString(50, 560, f"Total Amount: {c[7]:,} Tomans")
        pdf.drawString(50, 540, f"Discount: {c[8]:,} Tomans")
        pdf.drawString(50, 520, f"Paid Deposits: {c[9]:,} Tomans")
        remain = c[7] - c[9]
        pdf.drawString(50, 500, f"Remaining: {remain:,} Tomans")

        pdf.save()
        QMessageBox.information(self, "موفقیت", "فاکتور PDF با موفقیت ذخیره شد.")

    # --- زبانه ۲: تبلیغاتی و تولدی ---
    def setup_commercial_tab(self):
        layout = QHBoxLayout()
        form_box = QGroupBox("ثبت پروژه جدید")
        form_box.setFont(QFont("B Yekan", 10, QFont.Weight.Bold))
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
        self.c_desc = QLineEdit()

        form_layout.addRow("عنوان پروژه:", self.c_title)
        form_layout.addRow("نوع پروژه:", self.c_type)
        form_layout.addRow("تعداد دوربین:", self.c_cameras)
        form_layout.addRow("مبلغ کل (تومان):", self.c_amount)
        form_layout.addRow("مبلغ پرداختی (تومان):", self.c_paid)
        form_layout.addRow("تاریخ پروژه:", self.c_date)
        form_layout.addRow("توضیحات:", self.c_desc)

        btn_save = QPushButton("ثبت پروژه")
        btn_save.setFont(QFont("B Yekan", 10, QFont.Weight.Bold))
        btn_save.setStyleSheet("background-color: #27ae60; color: white; padding: 6px;")
        btn_save.clicked.connect(self.save_commercial_project)
        form_layout.addRow(btn_save)

        form_box.setLayout(form_layout)
        layout.addWidget(form_box, 1)

        self.c_table = QTableWidget()
        self.c_table.setColumnCount(8)
        self.c_table.setHorizontalHeaderLabels(["ID", "عنوان", "نوع", "تعداد دوربین", "مبلغ کل", "پرداختی", "تاریخ", "توضیحات"])
        self.c_table.setWordWrap(True)
        self.c_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        layout.addWidget(self.c_table, 2)

        self.tab_commercial.setLayout(layout)
        self.load_commercial_projects()

    def save_commercial_project(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO commercial_projects (title, project_type, camera_count, total_amount, paid_amount, project_date, description)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        ''', (self.c_title.text(), self.c_type.currentText(), self.c_cameras.value(),
              parse_number(self.c_amount.text()), parse_number(self.c_paid.text()), self.c_date.text(), self.c_desc.text()))
        
        cursor.execute("UPDATE inventory SET used_count = used_count + ? WHERE item_name LIKE '%دوربین%'", (self.c_cameras.value(),))

        conn.commit()
        conn.close()
        QMessageBox.information(self, "موفقیت", "پروژه ثبت شد.")
        self.load_commercial_projects()
        self.load_inventory()

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
        self.c_table.resizeRowsToContents()

    # --- زبانه ۳: بخش کارکنان ---
    def setup_staff_tab(self):
        main_layout = QVBoxLayout()

        summary_group = QGroupBox("خلاصه وضعیت مالی ماه انتخاب‌شده")
        summary_layout = QHBoxLayout()

        self.lbl_income = QLabel("درآمد کل (ورودی): ۰ تومان")
        self.lbl_expense = QLabel("مجموع هزینه‌ها (خروجی): ۰ تومان")
        self.lbl_profit = QLabel("سود خالص: ۰ تومان")

        for lbl in [self.lbl_income, self.lbl_expense, self.lbl_profit]:
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setStyleSheet("font-size: 13px; font-weight: bold; padding: 10px; border: 1px solid #bdc3c7; background-color: #f8f9fa; border-radius: 6px;")
            summary_layout.addWidget(lbl)

        summary_group.setLayout(summary_layout)
        main_layout.addWidget(summary_group)

        top_bar = QHBoxLayout()
        date_group = QGroupBox("انتخاب تاریخ شمسی")
        date_layout = QHBoxLayout()

        today = jdatetime.date.today()

        date_layout.addWidget(QLabel("روز:"))
        self.spin_day = QSpinBox()
        self.spin_day.setRange(1, 31)
        self.spin_day.setValue(today.day)
        date_layout.addWidget(self.spin_day)

        date_layout.addWidget(QLabel("ماه:"))
        self.combo_month = QComboBox()
        self.combo_month.addItems(self.PERSIAN_MONTHS)
        self.combo_month.setCurrentIndex(today.month - 1)
        self.combo_month.currentIndexChanged.connect(self.load_summary_and_table)
        date_layout.addWidget(self.combo_month)

        date_layout.addWidget(QLabel("سال:"))
        self.spin_year = QSpinBox()
        self.spin_year.setRange(1390, 1450)
        self.spin_year.setValue(today.year)
        self.spin_year.valueChanged.connect(self.load_summary_and_table)
        date_layout.addWidget(self.spin_year)

        date_group.setLayout(date_layout)
        top_bar.addWidget(date_group, 2)

        reports_group = QGroupBox("گزارش‌‌گیری")
        reports_layout = QHBoxLayout()

        btn_excel = QPushButton("خروجی اکسل (Excel)")
        btn_excel.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; padding: 8px;")
        btn_excel.clicked.connect(self.export_to_excel)

        btn_chart = QPushButton("نمایش نمودار هزینه‌ها")
        btn_chart.setStyleSheet("background-color: #2980b9; color: white; font-weight: bold; padding: 8px;")
        btn_chart.clicked.connect(self.show_cost_chart)

        reports_layout.addWidget(btn_excel)
        reports_layout.addWidget(btn_chart)
        reports_group.setLayout(reports_layout)
        top_bar.addWidget(reports_group, 1)

        main_layout.addLayout(top_bar)

        forms_layout = QHBoxLayout()

        person_box = QGroupBox("تعریف نیروی جدید")
        person_form = QFormLayout()
        self.txt_person_name = QLineEdit()
        self.combo_person_role = QComboBox()
        self.combo_person_role.addItems(self.ROLES)
        btn_add_person = QPushButton("ثبت فرد جدید")
        btn_add_person.clicked.connect(self.add_person)

        person_form.addRow("نام و نام خانوادگی:", self.txt_person_name)
        person_form.addRow("تخصص / نقش:", self.combo_person_role)
        person_form.addRow(btn_add_person)
        person_box.setLayout(person_form)

        trans_box = QGroupBox("ثبت تراکنش مالی")
        trans_form = QFormLayout()

        self.combo_trans_type = QComboBox()
        self.combo_trans_type.addItems(["پرداختی/هزینه (خروجی)", "درآمد پروژه (ورودی)"])
        self.combo_trans_type.currentIndexChanged.connect(self.toggle_trans_type_fields)

        self.combo_category = QComboBox()
        self.combo_category.addItems(self.ROLES + self.PROJECT_TYPES)
        self.combo_category.currentIndexChanged.connect(self.update_persons_dropdown)

        self.combo_persons = QComboBox()
        self.txt_amount = QLineEdit()
        self.txt_amount.setPlaceholderText("مبلغ به تومان")
        self.txt_amount.textChanged.connect(lambda t: self.txt_amount.setText(format_number(t)))
        self.txt_desc = QLineEdit()
        self.txt_desc.setPlaceholderText("توضیحات پروژه...")

        btn_add_trans = QPushButton("ثبت تراکنش")
        btn_add_trans.setStyleSheet("background-color: #16a085; color: white; font-weight: bold;")
        btn_add_trans.clicked.connect(self.add_transaction)

        trans_form.addRow("نوع تراکنش:", self.combo_trans_type)
        trans_form.addRow("بخش / نقش:", self.combo_category)
        trans_form.addRow("انتخاب فرد:", self.combo_persons)
        trans_form.addRow("مبلغ (تومان):", self.txt_amount)
        trans_form.addRow("توضیحات:", self.txt_desc)
        trans_form.addRow(btn_add_trans)
        trans_box.setLayout(trans_form)

        forms_layout.addWidget(person_box, 1)
        forms_layout.addWidget(trans_box, 2)
        main_layout.addLayout(forms_layout)

        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels([
            "ID", "نوع", "تاریخ", "بخش/دسته‌‌بندی", "فرد مربوطه", "مبلغ (تومان)", "توضیحات", "حذف"
        ])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        main_layout.addWidget(self.table)

        self.tab_staff.setLayout(main_layout)

        self.load_persons_combo()
        self.load_summary_and_table()
        self.toggle_trans_type_fields()

    def add_person(self):
        name = self.txt_person_name.text().strip()
        role = self.combo_person_role.currentText()
        if not name:
            QMessageBox.warning(self, "خطا", "لطفاً نام فرد را وارد کنید.")
            return

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO persons (name, role) VALUES (?, ?)", (name, role))
        conn.commit()
        conn.close()

        QMessageBox.information(self, "موفقیت", f"فرد '{name}' اضافه شد.")
        self.txt_person_name.clear()
        self.load_persons_combo()

    def update_persons_dropdown(self):
        category = self.combo_category.currentText()
        self.combo_persons.clear()
        self.combo_persons.addItem("--- بدون انتخاب / متفرقه ---", None)

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, name FROM persons WHERE role = ?", (category,))
        rows = cursor.fetchall()
        conn.close()

        for person_id, name in rows:
            self.combo_persons.addItem(name, person_id)

    def load_persons_combo(self):
        self.update_persons_dropdown()

    def toggle_trans_type_fields(self):
        is_expense = self.combo_trans_type.currentIndex() == 0
        self.combo_persons.setEnabled(is_expense)

    def add_transaction(self):
        is_expense = self.combo_trans_type.currentIndex() == 0
        trans_type = 'expense' if is_expense else 'income'
        category = self.combo_category.currentText()
        person_id = self.combo_persons.currentData() if is_expense else None

        amount = parse_number(self.txt_amount.text())
        description = self.txt_desc.text().strip()

        if amount <= 0:
            QMessageBox.warning(self, "خطا", "لطفاً مبلغ معتبر وارد کنید.")
            return

        year = self.spin_year.value()
        month = self.combo_month.currentIndex() + 1
        day = self.spin_day.value()

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO transactions (trans_type, category, person_id, amount, year, month, day, description)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (trans_type, category, person_id, amount, year, month, day, description))
        conn.commit()
        conn.close()

        self.txt_amount.clear()
        self.txt_desc.clear()
        self.load_summary_and_table()

    def load_summary_and_table(self):
        year = self.spin_year.value()
        month = self.combo_month.currentIndex() + 1

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()

        cursor.execute('''
            SELECT t.id, t.trans_type, t.day, t.category, p.name, t.amount, t.description
            FROM transactions t
            LEFT JOIN persons p ON t.person_id = p.id
            WHERE t.year = ? AND t.month = ?
            ORDER BY t.day DESC, t.id DESC
        ''', (year, month))
        rows = cursor.fetchall()

        cursor.execute("SELECT SUM(amount) FROM transactions WHERE year=? AND month=? AND trans_type='income'", (year, month))
        total_income = cursor.fetchone()[0] or 0

        cursor.execute("SELECT SUM(amount) FROM transactions WHERE year=? AND month=? AND trans_type='expense'", (year, month))
        total_expense = cursor.fetchone()[0] or 0

        profit = total_income - total_expense

        self.lbl_income.setText(f"درآمد کل (ورودی): {total_income:,} تومان")
        self.lbl_expense.setText(f"مجموع هزینه‌ها (خروجی/نیروها): {total_expense:,} تومان")
        self.lbl_profit.setText(f"سود خالص: {profit:,} تومان")

        if profit >= 0:
            self.lbl_profit.setStyleSheet("font-size: 13px; font-weight: bold; padding: 10px; border: 1px solid green; background-color: #e8f8f5; color: #27ae60; border-radius: 6px;")
        else:
            self.lbl_profit.setStyleSheet("font-size: 13px; font-weight: bold; padding: 10px; border: 1px solid red; background-color: #fadbd8; color: #c0392b; border-radius: 6px;")

        self.table.setRowCount(0)
        for row_idx, row in enumerate(rows):
            trans_id, trans_type, day, category, person_name, amount, desc = row

            self.table.insertRow(row_idx)
            self.table.setItem(row_idx, 0, QTableWidgetItem(str(trans_id)))

            type_str = "درآمد (ورودی)" if trans_type == 'income' else "هزینه (خروجی)"
            self.table.setItem(row_idx, 1, QTableWidgetItem(type_str))

            date_str = f"{year}/{month:02d}/{day:02d}"
            self.table.setItem(row_idx, 2, QTableWidgetItem(date_str))
            self.table.setItem(row_idx, 3, QTableWidgetItem(category))
            self.table.setItem(row_idx, 4, QTableWidgetItem(person_name if person_name else "-"))
            self.table.setItem(row_idx, 5, QTableWidgetItem(f"{amount:,}"))
            self.table.setItem(row_idx, 6, QTableWidgetItem(desc if desc else "-"))

            btn_delete = QPushButton("حذف")
            btn_delete.setStyleSheet("background-color: #e74c3c; color: white;")
            btn_delete.clicked.connect(lambda _, tid=trans_id: self.delete_transaction(tid))
            self.table.setCellWidget(row_idx, 7, btn_delete)

        conn.close()

    def delete_transaction(self, trans_id):
        reply = QMessageBox.question(self, "تأیید حذف", "آیا از حذف این تراکنش مطمئن هستید؟",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM transactions WHERE id = ?", (trans_id,))
            conn.commit()
            conn.close()
            self.load_summary_and_table()

    def export_to_excel(self):
        year = self.spin_year.value()
        month = self.combo_month.currentIndex() + 1
        month_name = self.PERSIAN_MONTHS[month - 1]

        file_path, _ = QFileDialog.getSaveFileName(self, "ذخیره فایل اکسل", f"گزارش_{month_name}_{year}.xlsx", "Excel Files (*.xlsx)")
        if not file_path:
            return

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT t.id, t.trans_type, t.year || '/' || t.month || '/' || t.day, t.category, p.name, t.amount, t.description
            FROM transactions t
            LEFT JOIN persons p ON t.person_id = p.id
            WHERE t.year = ? AND t.month = ?
            ORDER BY t.day ASC
        ''', (year, month))
        rows = cursor.fetchall()
        conn.close()

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = f"گزارش {month_name}"
        ws.views.sheetView[0].rightToLeft = True

        headers = ["شناسه", "نوع تراکنش", "تاریخ", "دسته‌بندی/نقش", "فرد مربوطه", "مبلغ (تومان)", "توضیحات"]
        ws.append(headers)

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = XLFont(name="Tahoma", size=11, bold=True, color="FFFFFF")

        for col_idx in range(1, 8):
            cell = ws.cell(row=1, column=col_idx)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        for row in rows:
            r_list = list(row)
            r_list[1] = "درآمد (ورودی)" if r_list[1] == 'income' else "هزینه (خروجی)"
            r_list[4] = r_list[4] if r_list[4] else "-"
            r_list[6] = r_list[6] if r_list[6] else "-"
            ws.append(r_list)

        for r in range(2, len(rows) + 2):
            for c in range(1, 8):
                cell = ws.cell(row=r, column=c)
                cell.font = XLFont(name="Tahoma", size=10)
                if c == 6:
                    cell.number_format = '#,##0'

        wb.save(file_path)
        QMessageBox.information(self, "موفقیت", "فایل اکسل با موفقیت ذخیره شد.")

    def show_cost_chart(self):
        year = self.spin_year.value()
        month = self.combo_month.currentIndex() + 1

        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('''
            SELECT category, SUM(amount)
            FROM transactions
            WHERE year = ? AND month = ? AND trans_type = 'expense'
            GROUP BY category
        ''', (year, month))
        data = cursor.fetchall()
        conn.close()

        if not data:
            QMessageBox.information(self, "اطلاع", "هیچ هزینه‌ای برای این ماه ثبت نشده است.")
            return

        categories = [d[0] for d in data]
        amounts = [d[1] for d in data]

        plt.figure(figsize=(7, 7))
        plt.pie(amounts, labels=categories, autopct='%1.1f%%', startangle=140)
        plt.title(f"سهم هزینه‌ها و پرداخت‌های ماه {self.PERSIAN_MONTHS[month-1]} {year}")
        plt.show()

    # --- زبانه ۴: انبار تجهیزات ---
    def setup_inventory_tab(self):
        layout = QVBoxLayout()
        self.inv_table = QTableWidget()
        self.inv_table.setColumnCount(4)
        self.inv_table.setHorizontalHeaderLabels(["نام تجهیزات", "تعداد کل", "در حال استفاده (پروژه‌ها)", "موجودی باقی‌مانده"])
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

    # --- زبانه ۵: کارت‌های بانکی ---
    def setup_banks_tab(self):
        layout = QHBoxLayout()
        form_box = QGroupBox("افزودن حساب بانکی")
        form_box.setFont(QFont("B Yekan", 10, QFont.Weight.Bold))
        form_layout = QFormLayout()

        self.b_name = QLineEdit()
        self.b_card = QLineEdit()
        self.b_sheba = QLineEdit()

        form_layout.addRow("نام بانک:", self.b_name)
        form_layout.addRow("شماره کارت:", self.b_card)
        form_layout.addRow("شماره شبا:", self.b_sheba)

        btn_save = QPushButton("ذخیره کارت")
        btn_save.setFont(QFont("B Yekan", 10, QFont.Weight.Bold))
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

    # --- عمومی و پشتیبان‌گیری ---
    def backup_db(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "ذخیره فایل پشتیبان", "studio_backup.db", "Database Files (*.db)")
        if file_path:
            shutil.copyfile(DB_NAME, file_path)
            QMessageBox.information(self, "پشتیبان‌گیری", "پشتیبان‌گیری با موفقیت انجام شد.")

    def restore_db(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل پشتیبان", "", "Database Files (*.db)")
        if file_path:
            shutil.copyfile(file_path, DB_NAME)
            QMessageBox.information(self, "بازیابی", "اطلاعات بازیابی شد. برنامه مجدداً اجرا می‌شود.")

    def change_password(self):
        new_pass, ok = QInputDialog.getText(self, "تغییر رمز عبور", "رمز عبور جدید را وارد کنید:")
        if ok and new_pass:
            conn = sqlite3.connect(DB_NAME)
            cursor = conn.cursor()
            cursor.execute("UPDATE settings SET value=? WHERE key='app_password'", (new_pass,))
            conn.commit()
            conn.close()
            QMessageBox.information(self, "موفقیت", "رمز عبور تغییر یافت.")

    def closeEvent(self, event):
        reply = QMessageBox.question(self, "پشتیبان‌گیری خودکار", "آیا مایلید قبل از خروج فایل بک‌آپ ذخیره شود؟",
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.backup_db()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = StudioAccountingApp()
    window.show()
    QTimer.singleShot(100, window.prompt_login)
    sys.exit(app.exec())
