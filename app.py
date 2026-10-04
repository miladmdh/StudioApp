{\rtf1\ansi\ansicpg1252\cocoartf2907
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\paperw11900\paperh16840\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 import sys\
import sqlite3\
import jdatetime\
from PyQt6.QtWidgets import (\
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,\
    QLabel, QLineEdit, QComboBox, QPushButton, QTableWidget, QTableWidgetItem,\
    QHeaderView, QMessageBox, QGroupBox, QSpinBox, QFormLayout, QFileDialog\
)\
from PyQt6.QtCore import Qt\
from PyQt6.QtGui import QFont\
\
import openpyxl\
from openpyxl.styles import Font as XLFont, PatternFill, Alignment, Border, Side\
from openpyxl.utils import get_column_letter\
\
import matplotlib.pyplot as plt\
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas\
\
# --- Database Setup ---\
DB_NAME = "studio_accounting.db"\
\
def init_db():\
    conn = sqlite3.connect(DB_NAME)\
    cursor = conn.cursor()\
    cursor.execute('''\
        CREATE TABLE IF NOT EXISTS persons (\
            id INTEGER PRIMARY KEY AUTOINCREMENT,\
            name TEXT NOT NULL,\
            role TEXT NOT NULL\
        )\
    ''')\
    cursor.execute('''\
        CREATE TABLE IF NOT EXISTS transactions (\
            id INTEGER PRIMARY KEY AUTOINCREMENT,\
            trans_type TEXT NOT NULL,\
            category TEXT NOT NULL,\
            person_id INTEGER,\
            amount INTEGER NOT NULL,\
            year INTEGER NOT NULL,\
            month INTEGER NOT NULL,\
            day INTEGER NOT NULL,\
            description TEXT,\
            FOREIGN KEY (person_id) REFERENCES persons(id)\
        )\
    ''')\
    conn.commit()\
    conn.close()\
\
class StudioAccountingApp(QMainWindow):\
    ROLES = ["\uc0\u1578 \u1583 \u1608 \u1740 \u1606 \u1711 \u1585 ", "\u1593 \u1705 \u1575 \u1587 ", "\u1601 \u1740 \u1604 \u1605 \u1576 \u1585 \u1583 \u1575 \u1585 ", "\u1607 \u1604 \u1740  \u1588 \u1575 \u1578  \u1608  FPV \u1705 \u1575 \u1585 ", "\u1575 \u1608 \u1662 \u1585 \u1575 \u1578 \u1608 \u1585  \u1705 \u1585 \u1740 \u1606 "]\
    PROJECT_TYPES = ["\uc0\u1593 \u1585 \u1608 \u1587 \u1740 ", "\u1593 \u1602 \u1583 ", "\u1578 \u1608 \u1604 \u1583 ", "\u1578 \u1576 \u1604 \u1740 \u1594 \u1575 \u1578 \u1740 ", "\u1602 \u1576 \u1590  \u1608  \u1705 \u1585 \u1575 \u1740 \u1607 "]\
    PERSIAN_MONTHS = [\
        "\uc0\u1601 \u1585 \u1608 \u1585 \u1583 \u1740 \u1606 ", "\u1575 \u1585 \u1583 \u1740 \u1576 \u1607 \u1588 \u1578 ", "\u1582 \u1585 \u1583 \u1575 \u1583 ", "\u1578 \u1740 \u1585 ", "\u1605 \u1585 \u1583 \u1575 \u1583 ", "\u1588 \u1607 \u1585 \u1740 \u1608 \u1585 ",\
        "\uc0\u1605 \u1607 \u1585 ", "\u1570 \u1576 \u1575 \u1606 ", "\u1570 \u1584 \u1585 ", "\u1583 \u1740 ", "\u1576 \u1607 \u1605 \u1606 ", "\u1575 \u1587 \u1601 \u1606 \u1583 "\
    ]\
\
    def __init__(self):\
        super().__init__()\
        self.setWindowTitle("\uc0\u1606 \u1585 \u1605 \u8204 \u1575 \u1601 \u1586 \u1575 \u1585  \u1605 \u1583 \u1740 \u1585 \u1740 \u1578  \u1605 \u1575 \u1604 \u1740  \u1570 \u1578 \u1604 \u1740 \u1607  \u1608  \u1575 \u1587 \u1578 \u1608 \u1583 \u1740 \u1608  \u1601 \u1740 \u1604 \u1605 \u8204 \u1576 \u1585 \u1583 \u1575 \u1585 \u1740 ")\
        self.resize(1150, 800)\
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)\
        \
        # \uc0\u1578 \u1606 \u1592 \u1740 \u1605  \u1601 \u1608 \u1606 \u1578  \u1575 \u1587 \u1578 \u1575 \u1606 \u1583 \u1575 \u1585 \u1583  \u1601 \u1575 \u1585 \u1587 \u1740  \u1576 \u1585 \u1575 \u1740  \u1580 \u1604 \u1608 \u1711 \u1740 \u1585 \u1740  \u1575 \u1586  \u1580 \u1583 \u1575  \u1588 \u1583 \u1606  \u1581 \u1585 \u1608 \u1601 \
        app_font = QFont("Tahoma", 10)\
        self.setFont(app_font)\
\
        init_db()\
        self.init_ui()\
        self.load_persons_combo()\
        self.load_summary_and_table()\
\
    def init_ui(self):\
        main_widget = QWidget()\
        self.setCentralWidget(main_widget)\
        main_layout = QVBoxLayout()\
        main_widget.setLayout(main_layout)\
\
        # 1. \uc0\u1582 \u1604 \u1575 \u1589 \u1607  \u1605 \u1575 \u1604 \u1740 \
        summary_group = QGroupBox("\uc0\u1582 \u1604 \u1575 \u1589 \u1607  \u1608 \u1590 \u1593 \u1740 \u1578  \u1605 \u1575 \u1604 \u1740  \u1605 \u1575 \u1607  \u1575 \u1606 \u1578 \u1582 \u1575 \u1576 \u8204 \u1588 \u1583 \u1607 ")\
        summary_layout = QHBoxLayout()\
        \
        self.lbl_income = QLabel("\uc0\u1583 \u1585 \u1570 \u1605 \u1583  \u1705 \u1604  (\u1608 \u1585 \u1608 \u1583 \u1740 ): \u1776  \u1578 \u1608 \u1605 \u1575 \u1606 ")\
        self.lbl_expense = QLabel("\uc0\u1605 \u1580 \u1605 \u1608 \u1593  \u1607 \u1586 \u1740 \u1606 \u1607 \u8204 \u1607 \u1575  (\u1582 \u1585 \u1608 \u1580 \u1740 ): \u1776  \u1578 \u1608 \u1605 \u1575 \u1606 ")\
        self.lbl_profit = QLabel("\uc0\u1587 \u1608 \u1583  \u1582 \u1575 \u1604 \u1589 : \u1776  \u1578 \u1608 \u1605 \u1575 \u1606 ")\
\
        for lbl in [self.lbl_income, self.lbl_expense, self.lbl_profit]:\
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)\
            lbl.setStyleSheet("font-size: 13px; font-weight: bold; padding: 10px; border: 1px solid #bdc3c7; background-color: #f8f9fa; border-radius: 6px;")\
            summary_layout.addWidget(lbl)\
\
        summary_group.setLayout(summary_layout)\
        main_layout.addWidget(summary_group)\
\
        # 2. \uc0\u1578 \u1575 \u1585 \u1740 \u1582  \u1588 \u1605 \u1587 \u1740  \u1608  \u1583 \u1705 \u1605 \u1607 \u8204 \u1607 \u1575 \u1740  \u1711 \u1586 \u1575 \u1585 \u1588 \u8204 \u1711 \u1740 \u1585 \u1740 \
        top_bar = QHBoxLayout()\
        \
        date_group = QGroupBox("\uc0\u1575 \u1606 \u1578 \u1582 \u1575 \u1576  \u1578 \u1575 \u1585 \u1740 \u1582  \u1588 \u1605 \u1587 \u1740 ")\
        date_layout = QHBoxLayout()\
\
        today = jdatetime.date.today()\
\
        date_layout.addWidget(QLabel("\uc0\u1585 \u1608 \u1586 :"))\
        self.spin_day = QSpinBox()\
        self.spin_day.setRange(1, 31)\
        self.spin_day.setValue(today.day)\
        date_layout.addWidget(self.spin_day)\
\
        date_layout.addWidget(QLabel("\uc0\u1605 \u1575 \u1607 :"))\
        self.combo_month = QComboBox()\
        self.combo_month.addItems(self.PERSIAN_MONTHS)\
        self.combo_month.setCurrentIndex(today.month - 1)\
        self.combo_month.currentIndexChanged.connect(self.load_summary_and_table)\
        date_layout.addWidget(self.combo_month)\
\
        date_layout.addWidget(QLabel("\uc0\u1587 \u1575 \u1604 :"))\
        self.spin_year = QSpinBox()\
        self.spin_year.setRange(1390, 1450)\
        self.spin_year.setValue(today.year)\
        self.spin_year.valueChanged.connect(self.load_summary_and_table)\
        date_layout.addWidget(self.spin_year)\
\
        date_group.setLayout(date_layout)\
        top_bar.addWidget(date_group, 2)\
\
        # \uc0\u1583 \u1705 \u1605 \u1607 \u8204 \u1607 \u1575 \u1740  \u1582 \u1585 \u1608 \u1580 \u1740 \
        reports_group = QGroupBox("\uc0\u1711 \u1586 \u1575 \u1585 \u1588 \u8204 \u1711 \u1740 \u1585 \u1740 ")\
        reports_layout = QHBoxLayout()\
\
        btn_excel = QPushButton("\uc0\u1582 \u1585 \u1608 \u1580 \u1740  \u1575 \u1705 \u1587 \u1604  (Excel)")\
        btn_excel.setStyleSheet("background-color: #27ae60; color: white; font-weight: bold; padding: 8px;")\
        btn_excel.clicked.connect(self.export_to_excel)\
\
        btn_chart = QPushButton("\uc0\u1606 \u1605 \u1575 \u1740 \u1588  \u1606 \u1605 \u1608 \u1583 \u1575 \u1585  \u1607 \u1586 \u1740 \u1606 \u1607 \u8204 \u1607 \u1575 ")\
        btn_chart.setStyleSheet("background-color: #2980b9; color: white; font-weight: bold; padding: 8px;")\
        btn_chart.clicked.connect(self.show_cost_chart)\
\
        reports_layout.addWidget(btn_excel)\
        reports_layout.addWidget(btn_chart)\
        reports_group.setLayout(reports_layout)\
        top_bar.addWidget(reports_group, 1)\
\
        main_layout.addLayout(top_bar)\
\
        # 3. \uc0\u1601 \u1585 \u1605 \u8204 \u1607 \u1575 \u1740  \u1608 \u1585 \u1608 \u1583  \u1575 \u1591 \u1604 \u1575 \u1593 \u1575 \u1578 \
        forms_layout = QHBoxLayout()\
\
        # \uc0\u1578 \u1593 \u1585 \u1740 \u1601  \u1606 \u1740 \u1585 \u1608 \
        person_box = QGroupBox("\uc0\u1578 \u1593 \u1585 \u1740 \u1601  \u1606 \u1740 \u1585 \u1608 \u1740  \u1580 \u1583 \u1740 \u1583 ")\
        person_form = QFormLayout()\
        self.txt_person_name = QLineEdit()\
        self.combo_person_role = QComboBox()\
        self.combo_person_role.addItems(self.ROLES)\
        btn_add_person = QPushButton("\uc0\u1579 \u1576 \u1578  \u1601 \u1585 \u1583  \u1580 \u1583 \u1740 \u1583 ")\
        btn_add_person.clicked.connect(self.add_person)\
\
        person_form.addRow("\uc0\u1606 \u1575 \u1605  \u1608  \u1606 \u1575 \u1605  \u1582 \u1575 \u1606 \u1608 \u1575 \u1583 \u1711 \u1740 :", self.txt_person_name)\
        person_form.addRow("\uc0\u1578 \u1582 \u1589 \u1589  / \u1606 \u1602 \u1588 :", self.combo_person_role)\
        person_form.addRow(btn_add_person)\
        person_box.setLayout(person_form)\
\
        # \uc0\u1579 \u1576 \u1578  \u1578 \u1585 \u1575 \u1705 \u1606 \u1588 \
        trans_box = QGroupBox("\uc0\u1579 \u1576 \u1578  \u1578 \u1585 \u1575 \u1705 \u1606 \u1588  \u1605 \u1575 \u1604 \u1740 ")\
        trans_form = QFormLayout()\
\
        self.combo_trans_type = QComboBox()\
        self.combo_trans_type.addItems(["\uc0\u1662 \u1585 \u1583 \u1575 \u1582 \u1578 \u1740 /\u1607 \u1586 \u1740 \u1606 \u1607  (\u1582 \u1585 \u1608 \u1580 \u1740 )", "\u1583 \u1585 \u1570 \u1605 \u1583  \u1662 \u1585 \u1608 \u1688 \u1607  (\u1608 \u1585 \u1608 \u1583 \u1740 )"])\
        self.combo_trans_type.currentIndexChanged.connect(self.toggle_trans_type_fields)\
\
        self.combo_category = QComboBox()\
        self.combo_category.addItems(self.ROLES + self.PROJECT_TYPES)\
        self.combo_category.currentIndexChanged.connect(self.update_persons_dropdown)\
\
        self.combo_persons = QComboBox()\
        self.txt_amount = QLineEdit()\
        self.txt_amount.setPlaceholderText("\uc0\u1605 \u1576 \u1604 \u1594  \u1576 \u1607  \u1578 \u1608 \u1605 \u1575 \u1606 ")\
        self.txt_desc = QLineEdit()\
        self.txt_desc.setPlaceholderText("\uc0\u1578 \u1608 \u1590 \u1740 \u1581 \u1575 \u1578  \u1662 \u1585 \u1608 \u1688 \u1607 ...")\
\
        btn_add_trans = QPushButton("\uc0\u1579 \u1576 \u1578  \u1578 \u1585 \u1575 \u1705 \u1606 \u1588 ")\
        btn_add_trans.setStyleSheet("background-color: #16a085; color: white; font-weight: bold;")\
        btn_add_trans.clicked.connect(self.add_transaction)\
\
        trans_form.addRow("\uc0\u1606 \u1608 \u1593  \u1578 \u1585 \u1575 \u1705 \u1606 \u1588 :", self.combo_trans_type)\
        trans_form.addRow("\uc0\u1576 \u1582 \u1588  / \u1606 \u1602 \u1588 :", self.combo_category)\
        trans_form.addRow("\uc0\u1575 \u1606 \u1578 \u1582 \u1575 \u1576  \u1601 \u1585 \u1583 :", self.combo_persons)\
        trans_form.addRow("\uc0\u1605 \u1576 \u1604 \u1594  (\u1578 \u1608 \u1605 \u1575 \u1606 ):", self.txt_amount)\
        trans_form.addRow("\uc0\u1578 \u1608 \u1590 \u1740 \u1581 \u1575 \u1578 :", self.txt_desc)\
        trans_form.addRow(btn_add_trans)\
        trans_box.setLayout(trans_form)\
\
        forms_layout.addWidget(person_box, 1)\
        forms_layout.addWidget(trans_box, 2)\
        main_layout.addLayout(forms_layout)\
\
        # 4. \uc0\u1580 \u1583 \u1608 \u1604 \
        self.table = QTableWidget()\
        self.table.setColumnCount(8)\
        self.table.setHorizontalHeaderLabels([\
            "ID", "\uc0\u1606 \u1608 \u1593 ", "\u1578 \u1575 \u1585 \u1740 \u1582 ", "\u1576 \u1582 \u1588 /\u1583 \u1587 \u1578 \u1607 \u8204 \u1576 \u1606 \u1583 \u1740 ", "\u1601 \u1585 \u1583  \u1605 \u1585 \u1576 \u1608 \u1591 \u1607 ", "\u1605 \u1576 \u1604 \u1594  (\u1578 \u1608 \u1605 \u1575 \u1606 )", "\u1578 \u1608 \u1590 \u1740 \u1581 \u1575 \u1578 ", "\u1581 \u1584 \u1601 "\
        ])\
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)\
        main_layout.addWidget(self.table)\
\
        self.toggle_trans_type_fields()\
\
    # --- \uc0\u1578 \u1608 \u1575 \u1576 \u1593  \u1605 \u1606 \u1591 \u1602  \u1576 \u1585 \u1606 \u1575 \u1605 \u1607  ---\
    def add_person(self):\
        name = self.txt_person_name.text().strip()\
        role = self.combo_person_role.currentText()\
        if not name:\
            QMessageBox.warning(self, "\uc0\u1582 \u1591 \u1575 ", "\u1604 \u1591 \u1601 \u1575 \u1611  \u1606 \u1575 \u1605  \u1601 \u1585 \u1583  \u1585 \u1575  \u1608 \u1575 \u1585 \u1583  \u1705 \u1606 \u1740 \u1583 .")\
            return\
\
        conn = sqlite3.connect(DB_NAME)\
        cursor = conn.cursor()\
        cursor.execute("INSERT INTO persons (name, role) VALUES (?, ?)", (name, role))\
        conn.commit()\
        conn.close()\
\
        QMessageBox.information(self, "\uc0\u1605 \u1608 \u1601 \u1602 \u1740 \u1578 ", f"\u1601 \u1585 \u1583  '\{name\}' \u1575 \u1590 \u1575 \u1601 \u1607  \u1588 \u1583 .")\
        self.txt_person_name.clear()\
        self.load_persons_combo()\
\
    def update_persons_dropdown(self):\
        category = self.combo_category.currentText()\
        self.combo_persons.clear()\
        self.combo_persons.addItem("--- \uc0\u1576 \u1583 \u1608 \u1606  \u1575 \u1606 \u1578 \u1582 \u1575 \u1576  / \u1605 \u1578 \u1601 \u1585 \u1602 \u1607  ---", None)\
\
        conn = sqlite3.connect(DB_NAME)\
        cursor = conn.cursor()\
        cursor.execute("SELECT id, name FROM persons WHERE role = ?", (category,))\
        rows = cursor.fetchall()\
        conn.close()\
\
        for person_id, name in rows:\
            self.combo_persons.addItem(name, person_id)\
\
    def load_persons_combo(self):\
        self.update_persons_dropdown()\
\
    def toggle_trans_type_fields(self):\
        is_expense = self.combo_trans_type.currentIndex() == 0\
        self.combo_persons.setEnabled(is_expense)\
\
    def add_transaction(self):\
        is_expense = self.combo_trans_type.currentIndex() == 0\
        trans_type = 'expense' if is_expense else 'income'\
        category = self.combo_category.currentText()\
        person_id = self.combo_persons.currentData() if is_expense else None\
        \
        amount_str = self.txt_amount.text().strip()\
        description = self.txt_desc.text().strip()\
\
        if not amount_str.isdigit():\
            QMessageBox.warning(self, "\uc0\u1582 \u1591 \u1575 ", "\u1604 \u1591 \u1601 \u1575 \u1611  \u1605 \u1576 \u1604 \u1594  \u1585 \u1575  \u1576 \u1607  \u1593 \u1583 \u1583  (\u1578 \u1608 \u1605 \u1575 \u1606 ) \u1608 \u1575 \u1585 \u1583  \u1705 \u1606 \u1740 \u1583 .")\
            return\
\
        amount = int(amount_str)\
        year = self.spin_year.value()\
        month = self.combo_month.currentIndex() + 1\
        day = self.spin_day.value()\
\
        conn = sqlite3.connect(DB_NAME)\
        cursor = conn.cursor()\
        cursor.execute('''\
            INSERT INTO transactions (trans_type, category, person_id, amount, year, month, day, description)\
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)\
        ''', (trans_type, category, person_id, amount, year, month, day, description))\
        conn.commit()\
        conn.close()\
\
        self.txt_amount.clear()\
        self.txt_desc.clear()\
        self.load_summary_and_table()\
\
    def load_summary_and_table(self):\
        year = self.spin_year.value()\
        month = self.combo_month.currentIndex() + 1\
\
        conn = sqlite3.connect(DB_NAME)\
        cursor = conn.cursor()\
\
        cursor.execute('''\
            SELECT t.id, t.trans_type, t.day, t.category, p.name, t.amount, t.description\
            FROM transactions t\
            LEFT JOIN persons p ON t.person_id = p.id\
            WHERE t.year = ? AND t.month = ?\
            ORDER BY t.day DESC, t.id DESC\
        ''', (year, month))\
        rows = cursor.fetchall()\
\
        cursor.execute("SELECT SUM(amount) FROM transactions WHERE year=? AND month=? AND trans_type='income'", (year, month))\
        total_income = cursor.fetchone()[0] or 0\
\
        cursor.execute("SELECT SUM(amount) FROM transactions WHERE year=? AND month=? AND trans_type='expense'", (year, month))\
        total_expense = cursor.fetchone()[0] or 0\
\
        profit = total_income - total_expense\
\
        self.lbl_income.setText(f"\uc0\u1583 \u1585 \u1570 \u1605 \u1583  \u1705 \u1604  (\u1608 \u1585 \u1608 \u1583 \u1740 ): \{total_income:,\} \u1578 \u1608 \u1605 \u1575 \u1606 ")\
        self.lbl_expense.setText(f"\uc0\u1605 \u1580 \u1605 \u1608 \u1593  \u1607 \u1586 \u1740 \u1606 \u1607 \u8204 \u1607 \u1575  (\u1582 \u1585 \u1608 \u1580 \u1740 /\u1606 \u1740 \u1585 \u1608 \u1607 \u1575 ): \{total_expense:,\} \u1578 \u1608 \u1605 \u1575 \u1606 ")\
        self.lbl_profit.setText(f"\uc0\u1587 \u1608 \u1583  \u1582 \u1575 \u1604 \u1589 : \{profit:,\} \u1578 \u1608 \u1605 \u1575 \u1606 ")\
\
        if profit >= 0:\
            self.lbl_profit.setStyleSheet("font-size: 13px; font-weight: bold; padding: 10px; border: 1px solid green; background-color: #e8f8f5; color: #27ae60; border-radius: 6px;")\
        else:\
            self.lbl_profit.setStyleSheet("font-size: 13px; font-weight: bold; padding: 10px; border: 1px solid red; background-color: #fadbd8; color: #c0392b; border-radius: 6px;")\
\
        self.table.setRowCount(0)\
        for row_idx, row in enumerate(rows):\
            trans_id, trans_type, day, category, person_name, amount, desc = row\
            \
            self.table.insertRow(row_idx)\
            self.table.setItem(row_idx, 0, QTableWidgetItem(str(trans_id)))\
            \
            type_str = "\uc0\u1583 \u1585 \u1570 \u1605 \u1583  (\u1608 \u1585 \u1608 \u1583 \u1740 )" if trans_type == 'income' else "\u1607 \u1586 \u1740 \u1606 \u1607  (\u1582 \u1585 \u1608 \u1580 \u1740 )"\
            self.table.setItem(row_idx, 1, QTableWidgetItem(type_str))\
            \
            date_str = f"\{year\}/\{month:02d\}/\{day:02d\}"\
            self.table.setItem(row_idx, 2, QTableWidgetItem(date_str))\
            self.table.setItem(row_idx, 3, QTableWidgetItem(category))\
            self.table.setItem(row_idx, 4, QTableWidgetItem(person_name if person_name else "-"))\
            self.table.setItem(row_idx, 5, QTableWidgetItem(f"\{amount:,\}"))\
            self.table.setItem(row_idx, 6, QTableWidgetItem(desc if desc else "-"))\
\
            btn_delete = QPushButton("\uc0\u1581 \u1584 \u1601 ")\
            btn_delete.setStyleSheet("background-color: #e74c3c; color: white;")\
            btn_delete.clicked.connect(lambda _, tid=trans_id: self.delete_transaction(tid))\
            self.table.setCellWidget(row_idx, 7, btn_delete)\
\
        conn.close()\
\
    def delete_transaction(self, trans_id):\
        reply = QMessageBox.question(self, "\uc0\u1578 \u1571 \u1740 \u1740 \u1583  \u1581 \u1584 \u1601 ", "\u1570 \u1740 \u1575  \u1575 \u1586  \u1581 \u1584 \u1601  \u1575 \u1740 \u1606  \u1578 \u1585 \u1575 \u1705 \u1606 \u1588  \u1605 \u1591 \u1605 \u1574 \u1606  \u1607 \u1587 \u1578 \u1740 \u1583 \u1567 ",\
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)\
        if reply == QMessageBox.StandardButton.Yes:\
            conn = sqlite3.connect(DB_NAME)\
            cursor = conn.cursor()\
            cursor.execute("DELETE FROM transactions WHERE id = ?", (trans_id,))\
            conn.commit()\
            conn.close()\
            self.load_summary_and_table()\
\
    # --- \uc0\u1582 \u1585 \u1608 \u1580 \u1740  \u1575 \u1705 \u1587 \u1604  ---\
    def export_to_excel(self):\
        year = self.spin_year.value()\
        month = self.combo_month.currentIndex() + 1\
        month_name = self.PERSIAN_MONTHS[month - 1]\
\
        file_path, _ = QFileDialog.getSaveFileName(self, "\uc0\u1584 \u1582 \u1740 \u1585 \u1607  \u1601 \u1575 \u1740 \u1604  \u1575 \u1705 \u1587 \u1604 ", f"\u1711 \u1586 \u1575 \u1585 \u1588 _\{month_name\}_\{year\}.xlsx", "Excel Files (*.xlsx)")\
        if not file_path:\
            return\
\
        conn = sqlite3.connect(DB_NAME)\
        cursor = conn.cursor()\
        cursor.execute('''\
            SELECT t.id, t.trans_type, t.year || '/' || t.month || '/' || t.day, t.category, p.name, t.amount, t.description\
            FROM transactions t\
            LEFT JOIN persons p ON t.person_id = p.id\
            WHERE t.year = ? AND t.month = ?\
            ORDER BY t.day ASC\
        ''', (year, month))\
        rows = cursor.fetchall()\
        conn.close()\
\
        wb = openpyxl.Workbook()\
        ws = wb.active\
        ws.title = f"\uc0\u1711 \u1586 \u1575 \u1585 \u1588  \{month_name\}"\
        ws.views.sheetView[0].rightToLeft = True\
\
        # \uc0\u1578 \u1740 \u1578 \u1585  \u1608  \u1575 \u1587 \u1578 \u1575 \u1740 \u1604 \
        headers = ["\uc0\u1588 \u1606 \u1575 \u1587 \u1607 ", "\u1606 \u1608 \u1593  \u1578 \u1585 \u1575 \u1705 \u1606 \u1588 ", "\u1578 \u1575 \u1585 \u1740 \u1582 ", "\u1583 \u1587 \u1578 \u1607 \u8204 \u1576 \u1606 \u1583 \u1740 /\u1606 \u1602 \u1588 ", "\u1601 \u1585 \u1583  \u1605 \u1585 \u1576 \u1608 \u1591 \u1607 ", "\u1605 \u1576 \u1604 \u1594  (\u1578 \u1608 \u1605 \u1575 \u1606 )", "\u1578 \u1608 \u1590 \u1740 \u1581 \u1575 \u1578 "]\
        ws.append(headers)\
\
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")\
        header_font = XLFont(name="Tahoma", size=11, bold=True, color="FFFFFF")\
\
        for col_idx in range(1, 8):\
            cell = ws.cell(row=1, column=col_idx)\
            cell.fill = header_fill\
            cell.font = header_font\
            cell.alignment = Alignment(horizontal="center", vertical="center")\
\
        for row in rows:\
            r_list = list(row)\
            r_list[1] = "\uc0\u1583 \u1585 \u1570 \u1605 \u1583  (\u1608 \u1585 \u1608 \u1583 \u1740 )" if r_list[1] == 'income' else "\u1607 \u1586 \u1740 \u1606 \u1607  (\u1582 \u1585 \u1608 \u1580 \u1740 )"\
            r_list[4] = r_list[4] if r_list[4] else "-"\
            r_list[6] = r_list[6] if r_list[6] else "-"\
            ws.append(r_list)\
\
        # \uc0\u1575 \u1587 \u1578 \u1575 \u1740 \u1604 \u8204 \u8204 \u1583 \u1607 \u1740  \u1587 \u1604 \u1608 \u1604 \u8204 \u1607 \u1575  \u1608  \u1587 \u1607  \u1585 \u1602 \u1605  \u1580 \u1583 \u1575  \u1705 \u1585 \u1583 \u1606 \
        for r in range(2, len(rows) + 2):\
            for c in range(1, 8):\
                cell = ws.cell(row=r, column=c)\
                cell.font = XLFont(name="Tahoma", size=10)\
                if c == 6:\
                    cell.number_format = '#,##0'\
\
        wb.save(file_path)\
        QMessageBox.information(self, "\uc0\u1605 \u1608 \u1601 \u1602 \u1740 \u1578 ", "\u1601 \u1575 \u1740 \u1604  \u1575 \u1705 \u1587 \u1604  \u1576 \u1575  \u1605 \u1608 \u1601 \u1602 \u1740 \u1578  \u1584 \u1582 \u1740 \u1585 \u1607  \u1588 \u1583 .")\
\
    # --- \uc0\u1606 \u1605 \u1608 \u1583 \u1575 \u1585  \u1578 \u1601 \u1705 \u1740 \u1705 \u1740  \u1607 \u1586 \u1740 \u1606 \u1607 \u8204 \u1607 \u1575  ---\
    def show_cost_chart(self):\
        year = self.spin_year.value()\
        month = self.combo_month.currentIndex() + 1\
\
        conn = sqlite3.connect(DB_NAME)\
        cursor = conn.cursor()\
        cursor.execute('''\
            SELECT category, SUM(amount)\
            FROM transactions\
            WHERE year = ? AND month = ? AND trans_type = 'expense'\
            GROUP BY category\
        ''', (year, month))\
        data = cursor.fetchall()\
        conn.close()\
\
        if not data:\
            QMessageBox.information(self, "\uc0\u1575 \u1591 \u1604 \u1575 \u1593 ", "\u1607 \u1740 \u1670  \u1607 \u1586 \u1740 \u1606 \u1607 \u8204 \u1575 \u1740  \u1576 \u1585 \u1575 \u1740  \u1575 \u1740 \u1606  \u1605 \u1575 \u1607  \u1579 \u1576 \u1578  \u1606 \u1588 \u1583 \u1607  \u1575 \u1587 \u1578 .")\
            return\
\
        categories = [d[0] for d in data]\
        amounts = [d[1] for d in data]\
\
        plt.figure(figsize=(7, 7))\
        plt.pie(amounts, labels=categories, autopct='%1.1f%%', startangle=140)\
        plt.title(f"\uc0\u1587 \u1607 \u1605  \u1607 \u1586 \u1740 \u1606 \u1607 \u8204 \u1607 \u1575  \u1608  \u1662 \u1585 \u1583 \u1575 \u1582 \u1578 \u8204 \u1607 \u1575 \u1740  \u1605 \u1575 \u1607  \{self.PERSIAN_MONTHS[month-1]\} \{year\}")\
        plt.show()\
\
if __name__ == "__main__":\
    app = QApplication(sys.argv)\
    window = StudioAccountingApp()\
    window.show()\
    sys.exit(app.exec())\
}