import sys
import os
import shutil
import sqlite3
import jdatetime
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QComboBox, QPushButton, QTableWidget, QTableWidgetItem,
    QHeaderView, QMessageBox, QGroupBox, QSpinBox, QFormLayout, QFileDialog,
    QTabWidget, QCheckBox, QInputDialog
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QIcon

import openpyxl
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas

DB_NAME = "studio_accounting.db"

def resource_path(relative_path):
    """ دریافت مسیر فایل‌ها برای اجرا در حالت عادی یا بعد از PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

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
            deposit INTEGER,
            discount INTEGER,
            paid_amount INTEGER,
            bank_id INTEGER,
            is_settled INTEGER DEFAULT 0
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
            total_count INTEGER DEFAULT 0
        )
    ''')
    for item in default_items:
        cursor.execute("INSERT OR IGNORE INTO inventory (item_name, total_count) VALUES (?, 10)", (item,))

    conn.commit()
    conn.close()

class StudioAccountingApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("نرم‌افزار مدیریت مالی - IMART STUDIO v3.0")
        self.resize(1200, 850)
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        
        # تنظیم آیکون پنجره
        icon_path = resource_path("Accounting.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.setFont(QFont("B Yekan", 10))
        self.authenticated_tabs = set() # ثبت زبانه‌های تایید شده با رمز
        init_db()
        self.init_ui()

    def verify_password(self):
        """ متد احراز هویت با رمز عبور """
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT value FROM settings WHERE key='app_password'")
        saved_pass = cursor.fetchone()[0]
        conn.close()

        entered_pass, ok = QInputDialog.getText(
            self, "ورود بخش حفاظت‌شده", "لطفاً رمز عبور را وارد کنید:", QLineEdit.EchoMode.Password
        )
        if ok and entered_pass == saved_pass:
            QMessageBox.information(self, "موفقیت", "رمز صحیح است، خوش آمدید!")
            return True
        elif ok:
            QMessageBox.critical(self, "خطا", "رمز اشتباه است، دوباره تلاش کنید")
            return False
        return False

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)

        # منوی بالا
        top_bar = QHBoxLayout()
        btn_backup = QPushButton("پشتیبان‌گیری (Backup)")
        btn_backup.clicked.connect(self.backup_db)
        btn_restore = QPushButton("بازیابی بک‌آپ (Restore)")
        btn_restore.clicked.connect(self.restore_db)
        btn_change_pass = QPushButton("تغییر رمز عبور")
        btn_change_pass.clicked.connect(self.change_password)
        btn_about = QPushButton("درباره برنامه")
        btn_about.clicked.connect(self.show_about)

        for btn in [btn_backup, btn_restore, btn_change_pass, btn_about]:
            btn.setFont(QFont("B Yekan", 9, QFont.Weight.Bold))
            btn.setStyleSheet("padding: 5px 10px;")
            top_bar.addWidget(btn)
        top_bar.addStretch()
        main_layout.addLayout(top_bar)

        # زبانه اصلی (Tabs)
        self.tabs = QTabWidget()
        self.tabs.setFont(QFont("B Yekan", 10, QFont.Weight.Bold))
        
        self.tab_wedding = QWidget()
        self.tab_commercial = QWidget()
        self.tab_staff = QWidget()
        self.tab_inventory = QWidget()
        self.tab_banks = QWidget()

        self.tabs.addTab(self.tab_wedding, "پروژه‌های عروس و داماد 🔒")
        self.tabs.addTab(self.tab_commercial, "تبلیغاتی / بیوتی / تولدی 🔒")
        self.tabs.addTab(self.tab_staff, "بخش کارکنان و هزینه‌ها 🔒")
        self.tabs.addTab(self.tab_inventory, "انبار تجهیزات")
        self.tabs.addTab(self.tab_banks, "کارت‌های بانکی")

        # چک کردن رمز هنگام تعویض زبانه
        self.tabs.currentChanged.connect(self.on_tab_change)

        main_layout.addWidget(self.tabs)

        self.setup_wedding_tab()
        self.setup_commercial_tab()
        self.setup_staff_tab()
        self.setup_inventory_tab()
        self.setup_banks_tab()

        # فوتر زیر برنامه
        footer = QLabel("IMART STUDIO - Phone: 09173736618")
        footer.setAlignment(Qt.AlignmentFlag.AlignCenter)
        footer.setFont(QFont("B Nazanin", 10, QFont.Weight.Bold))
        footer.setStyleSheet("color: #2c3e50; margin-top: 5px;")
        main_layout.addWidget(footer)

    def on_tab_change(self, index):
        # زبانه‌های 0، 1 و 2 نیاز به رمز دارند
        if index in [0, 1, 2] and index not in self.authenticated_tabs:
            if self.verify_password():
                self.authenticated_tabs.add(index)
            else:
                # اگر رمز اشتباه بود برگرد به زبانه عمومی (انبار)
                self.tabs.setCurrentIndex(3)

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

        for w in [self.w_groom, self.w_bride, self.w_groom_phone, self.w_bride_phone, self.w_contract_date, self.w_ceremony_date]:
            w.setFont(QFont("B Nazanin", 10))

        form_layout.addRow("نام داماد:", self.w_groom)
        form_layout.addRow("نام عروس:", self.w_bride)
        form_layout.addRow("تلفن داماد:", self.w_groom_phone)
        form_layout.addRow("تلفن عروس:", self.w_bride_phone)
        form_layout.addRow("تاریخ قرارداد:", self.w_contract_date)
        form_layout.addRow("تاریخ مراسم:", self.w_ceremony_date)

        # چک باکس موارد فاکتور
        self.item_checkboxes = {}
        items_box = QGroupBox("موارد فاکتور")
        items_grid = QVBoxLayout()
        
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT item_name, price FROM item_prices")
        prices = dict(cursor.fetchall())
        conn.close()

        for item, price in prices.items():
            cb = QCheckBox(f"{item} ({price:,} تومان)")
            cb.setFont(QFont("B Nazanin", 10))
            cb.stateChanged.connect(self.calc_wedding_total)
            self.item_checkboxes[item] = cb
            items_grid.addWidget(cb)
        
        items_box.setLayout(items_grid)
        form_layout.addRow(items_box)

        self.w_lbl_total = QLabel("جمع کل: ۰ تومان")
        self.w_lbl_total.setFont(QFont("B Yekan", 10, QFont.Weight.Bold))
        self.w_lbl_total.setStyleSheet("color: #2c3e50;")
        form_layout.addRow(self.w_lbl_total)

        self.w_deposit = QLineEdit("0")
        self.w_discount = QLineEdit("0")
        self.w_paid = QLineEdit("0")
        for w in [self.w_deposit, self.w_discount, self.w_paid]:
            w.setFont(QFont("B Nazanin", 10))
            w.textChanged.connect(self.calc_wedding_total)

        form_layout.addRow("مبلغ بیعانه (تومان):", self.w_deposit)
        form_layout.addRow("تخفیف (تومان):", self.w_discount)
        form_layout.addRow("مبلغ پرداختی بعدی (تومان):", self.w_paid)

        self.w_bank_combo = QComboBox()
        self.w_bank_combo.setFont(QFont("B Nazanin", 10))
        self.load_bank_combo()
        form_layout.addRow("بانک واریزی:", self.w_bank_combo)

        self.w_lbl_remain = QLabel("مانده حساب: ۰ تومان")
        self.w_lbl_remain.setFont(QFont("B Yekan", 10, QFont.Weight.Bold))
        self.w_lbl_remain.setStyleSheet("color: #c0392b;")
        form_layout.addRow(self.w_lbl_remain)

        btn_save = QPushButton("ثبت قرارداد")
        btn_save.setFont(QFont("B Yekan", 10, QFont.Weight.Bold))
        btn_save.setStyleSheet("background-color: #2980b9; color: white; padding: 6px;")
        btn_save.clicked.connect(self.save_wedding_contract)
        form_layout.addRow(btn_save)

        form_box.setLayout(form_layout)
        layout.addWidget(form_box, 1)

        # جدول قراردادها
        table_box = QGroupBox("لیست قراردادهای ثبت‌شده")
        table_box.setFont(QFont("B Yekan", 10, QFont.Weight.Bold))
        table_layout = QVBoxLayout()
        
        self.w_table = QTableWidget()
        self.w_table.setFont(QFont("B Nazanin", 10))
        self.w_table.setColumnCount(10)
        self.w_table.setHorizontalHeaderLabels([
            "ID", "زوجین", "تماس", "تاریخ مراسم", "جمع کل", "تخفیف", "پرداختی", "مانده", "وضعیت", "پرینت PDF"
        ])
        table_layout.addWidget(self.w_table)
        table_box.setLayout(table_layout)
        layout.addWidget(table_box, 2)

        self.tab_wedding.setLayout(layout)
        self.load_wedding_contracts()

    def calc_wedding_total(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT item_name, price FROM item_prices")
        prices = dict(cursor.fetchall())
        conn.close()

        total = sum(prices.get(item, 0) for item, cb in self.item_checkboxes.items() if cb.isChecked())

        try:
            deposit = int(self.w_deposit.text() or 0)
            discount = int(self.w_discount.text() or 0)
            paid = int(self.w_paid.text() or 0)
        except ValueError:
            deposit, discount, paid = 0, 0, 0

        final_total = total - discount
        remain = final_total - (deposit + paid)

        self.w_lbl_total.setText(f"جمع کل: {final_total:,} تومان")
        self.w_lbl_remain.setText(f"مانده حساب: {remain:,} تومان")

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
        
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT item_name, price FROM item_prices")
        prices = dict(cursor.fetchall())
        
        total = sum(prices.get(i, 0) for i in selected_items)
        deposit = int(self.w_deposit.text() or 0)
        discount = int(self.w_discount.text() or 0)
        paid = int(self.w_paid.text() or 0)
        bank_id = self.w_bank_combo.currentData()
        
        remain = (total - discount) - (deposit + paid)
        is_settled = 1 if remain <= 0 else 0

        cursor.execute('''
            INSERT INTO wedding_contracts 
            (groom_name, bride_name, groom_phone, bride_phone, contract_date, ceremony_date, selected_items, total_amount, deposit, discount, paid_amount, bank_id, is_settled)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (groom, bride, self.w_groom_phone.text(), self.w_bride_phone.text(),
              self.w_contract_date.text(), self.w_ceremony_date.text(),
              ",".join(selected_items), total, deposit, discount, paid, bank_id, is_settled))
        conn.commit()
        conn.close()

        QMessageBox.information(self, "موفقیت", "قرارداد عروسی با موفقیت ثبت شد.")
        self.load_wedding_contracts()

    def load_wedding_contracts(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, groom_name, bride_name, groom_phone, ceremony_date, total_amount, discount, deposit + paid_amount, is_settled FROM wedding_contracts ORDER BY id DESC")
        rows = cursor.fetchall()
        conn.close()

        self.w_table.setRowCount(0)
        for r_idx, row in enumerate(rows):
            c_id, groom, bride, phone, cer_date, total, discount, paid, settled = row
            remain = (total - discount) - paid

            self.w_table.insertRow(r_idx)
            self.w_table.setItem(r_idx, 0, QTableWidgetItem(str(c_id)))
            self.w_table.setItem(r_idx, 1, QTableWidgetItem(f"{groom} و {bride}"))
            self.w_table.setItem(r_idx, 2, QTableWidgetItem(phone))
            self.w_table.setItem(r_idx, 3, QTableWidgetItem(cer_date))
            self.w_table.setItem(r_idx, 4, QTableWidgetItem(f"{total:,}"))
            
            disc_item = QTableWidgetItem(f"{discount:,}")
            disc_item.setForeground(Qt.GlobalColor.red)
            self.w_table.setItem(r_idx, 5, disc_item)
            
            self.w_table.setItem(r_idx, 6, QTableWidgetItem(f"{paid:,}"))
            self.w_table.setItem(r_idx, 7, QTableWidgetItem(f"{remain:,}"))

            status_item = QTableWidgetItem("✔ تسویه کامل" if settled else "در حال تسویه")
            if settled:
                status_item.setForeground(Qt.GlobalColor.green)
            self.w_table.setItem(r_idx, 8, status_item)

            btn_pdf = QPushButton("چاپ PDF")
            btn_pdf.setFont(QFont("B Yekan", 9))
            btn_pdf.clicked.connect(lambda _, cid=c_id: self.export_wedding_pdf(cid))
            self.w_table.setCellWidget(r_idx, 9, btn_pdf)

    def export_wedding_pdf(self, contract_id):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT groom_name, bride_name, contract_date, ceremony_date, selected_items, total_amount, discount, deposit, paid_amount FROM wedding_contracts WHERE id=?", (contract_id,))
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
        pdf.drawString(50, 650, f"Contract Date: {c[2]}  |  Ceremony Date: {c[3]}")
        pdf.drawString(50, 630, f"Selected Services: {c[4]}")
        pdf.line(50, 610, 550, 610)

        pdf.drawString(50, 580, f"Total Amount: {c[5]:,} Tomans")
        pdf.drawString(50, 560, f"Discount: {c[6]:,} Tomans")
        pdf.drawString(50, 540, f"Deposit & Paid: {c[7] + c[8]:,} Tomans")
        remain = (c[5] - c[6]) - (c[7] + c[8])
        pdf.drawString(50, 520, f"Remaining: {remain:,} Tomans")

        pdf.save()
        QMessageBox.information(self, "موفقیت", "فاکتور PDF ذخیره شد.")

    # --- زبانه ۲: تبلیغاتی و تولدی ---
    def setup_commercial_tab(self):
        layout = QHBoxLayout()
        form_box = QGroupBox("ثبت پروژه تبلیغاتی / تولدی")
        form_box.setFont(QFont("B Yekan", 10, QFont.Weight.Bold))
        form_layout = QFormLayout()

        self.c_title = QLineEdit()
        self.c_type = QComboBox()
        self.c_type.addItems(["تبلیغاتی", "بیوتی", "تولدی"])
        self.c_cameras = QSpinBox()
        self.c_cameras.setValue(1)
        self.c_amount = QLineEdit()
        self.c_paid = QLineEdit()
        self.c_date = QLineEdit(jdatetime.date.today().strftime("%Y/%m/%d"))

        for w in [self.c_title, self.c_type, self.c_cameras, self.c_amount, self.c_paid, self.c_date]:
            w.setFont(QFont("B Nazanin", 10))

        form_layout.addRow("عنوان پروژه:", self.c_title)
        form_layout.addRow("نوع پروژه:", self.c_type)
        form_layout.addRow("تعداد دوربین:", self.c_cameras)
        form_layout.addRow("مبلغ کل (تومان):", self.c_amount)
        form_layout.addRow("مبلغ پرداختی (تومان):", self.c_paid)
        form_layout.addRow("تاریخ پروژه:", self.c_date)

        btn_save = QPushButton("ثبت پروژه")
        btn_save.setFont(QFont("B Yekan", 10, QFont.Weight.Bold))
        btn_save.setStyleSheet("background-color: #27ae60; color: white; padding: 6px;")
        btn_save.clicked.connect(self.save_commercial_project)
        form_layout.addRow(btn_save)

        form_box.setLayout(form_layout)
        layout.addWidget(form_box, 1)

        self.c_table = QTableWidget()
        self.c_table.setFont(QFont("B Nazanin", 10))
        self.c_table.setColumnCount(7)
        self.c_table.setHorizontalHeaderLabels(["ID", "عنوان", "نوع", "تعداد دوربین", "مبلغ", "تاریخ", "چاپ PDF"])
        layout.addWidget(self.c_table, 2)

        self.tab_commercial.setLayout(layout)
        self.load_commercial_projects()

    def save_commercial_project(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO commercial_projects (title, project_type, camera_count, total_amount, paid_amount, project_date)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', (self.c_title.text(), self.c_type.currentText(), self.c_cameras.value(),
              int(self.c_amount.text() or 0), int(self.c_paid.text() or 0), self.c_date.text()))
        conn.commit()
        conn.close()
        QMessageBox.information(self, "موفقیت", "پروژه ثبت شد.")
        self.load_commercial_projects()

    def load_commercial_projects(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT id, title, project_type, camera_count, total_amount, project_date FROM commercial_projects")
        rows = cursor.fetchall()
        conn.close()

        self.c_table.setRowCount(0)
        for r_idx, row in enumerate(rows):
            self.c_table.insertRow(r_idx)
            for c_idx, val in enumerate(row):
                self.c_table.setItem(r_idx, c_idx, QTableWidgetItem(str(val)))
            btn_pdf = QPushButton("چاپ PDF")
            btn_pdf.setFont(QFont("B Yekan", 9))
            self.c_table.setCellWidget(r_idx, 6, btn_pdf)

    # --- زبانه ۳: کارکنان ---
    def setup_staff_tab(self):
        layout = QVBoxLayout()
        lbl = QLabel("بخش مدیریت کارکنان و هزینه‌ها")
        lbl.setFont(QFont("B Yekan", 12, QFont.Weight.Bold))
        layout.addWidget(lbl)
        self.tab_staff.setLayout(layout)

    # --- زبانه ۴: انبار تجهیزات ---
    def setup_inventory_tab(self):
        layout = QVBoxLayout()
        self.inv_table = QTableWidget()
        self.inv_table.setFont(QFont("B Nazanin", 10))
        self.inv_table.setColumnCount(4)
        self.inv_table.setHorizontalHeaderLabels(["نام تجهیزات", "تعداد کل", "در حال استفاده", "باقی‌مانده"])
        layout.addWidget(self.inv_table)
        self.tab_inventory.setLayout(layout)
        self.load_inventory()

    def load_inventory(self):
        conn = sqlite3.connect(DB_NAME)
        cursor = conn.cursor()
        cursor.execute("SELECT item_name, total_count FROM inventory")
        rows = cursor.fetchall()
        conn.close()

        self.inv_table.setRowCount(0)
        for r_idx, (name, total) in enumerate(rows):
            used = 0
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

        for w in [self.b_name, self.b_card, self.b_sheba]:
            w.setFont(QFont("B Nazanin", 10))

        form_layout.addRow("نام بانک:", self.b_name)
        form_layout.addRow("شماره کارت:", self.b_card)
        form_layout.addRow("شماره شبا:", self.b_sheba)

        btn_save = QPushButton("ذخیره کارت")
        btn_save.setFont(QFont("B Yekan", 10, QFont.Weight.Bold))
        btn_save.clicked.connect(self.save_bank_card)
        form_layout.addRow(btn_save)

        form_box.setLayout(form_layout)
        layout.addWidget(form_box, 1)

        self.b_table = QTableWidget()
        self.b_table.setFont(QFont("B Nazanin", 10))
        self.b_table.setColumnCount(3)
        self.b_table.setHorizontalHeaderLabels(["نام بانک", "شماره کارت", "شماره شبا"])
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

    # --- امکانات عمومی ---
    def backup_db(self):
        if self.verify_password():
            file_path, _ = QFileDialog.getSaveFileName(self, "ذخیره فایل پشتیبان", "studio_backup.db", "Database Files (*.db)")
            if file_path:
                shutil.copyfile(DB_NAME, file_path)
                QMessageBox.information(self, "پشتیبان‌گیری", "فایل پشتیبان با موفقیت ذخیره شد.")

    def restore_db(self):
        if self.verify_password():
            file_path, _ = QFileDialog.getOpenFileName(self, "انتخاب فایل پشتیبان", "", "Database Files (*.db)")
            if file_path:
                shutil.copyfile(file_path, DB_NAME)
                QMessageBox.information(self, "بازیابی", "اطلاعات با موفقیت بازیابی شد. برنامه را مجدداً باز کنید.")

    def change_password(self):
        if self.verify_password():
            new_pass, ok = QInputDialog.getText(self, "تغییر رمز عبور", "رمز عبور جدید را وارد کنید:")
            if ok and new_pass:
                conn = sqlite3.connect(DB_NAME)
                cursor = conn.cursor()
                cursor.execute("UPDATE settings SET value=? WHERE key='app_password'", (new_pass,))
                conn.commit()
                conn.close()
                QMessageBox.information(self, "موفقیت", "رمز عبور با موفقیت تغییر کرد.")

    def show_about(self):
        msg = """
        <b>نرم‌افزار مدیریت مالی ایمارت استودیو</b><br>
        <b>نسخه:</b> 3.0<br>
        <b>طراح و توسعه‌دهنده:</b> میلاد محمدحسینی<br>
        <b>تلفن تماس:</b> 09173736618<br><br>
        <i>کلیه حقوق این نرم‌افزار متعلق به ایمارت استودیو می‌باشد.</i>
        """
        QMessageBox.about(self, "درباره برنامه", msg)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = StudioAccountingApp()
    window.show()
    sys.exit(app.exec())
