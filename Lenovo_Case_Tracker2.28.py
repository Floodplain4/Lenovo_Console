
import csv
import os
import re
import sys
import shutil
import ctypes
from datetime import datetime
from typing import List, Optional, Tuple, Set

try:
    import pyperclip
except ImportError:
    pyperclip = None

from PySide6.QtCore import Qt, QSettings
from PySide6.QtGui import QAction, QColor, QBrush, QIcon, QTextCursor
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QComboBox,
    QDialog,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)


APP_NAME = "Lenovo Case Tracker"
APP_VERSION = "v2.2.8"
LOG_FILE = "lcd_log.csv"
BACKUP_DIR = "backups"
ICON_FILE = "lenovo_case_tracker_icon.ico"

PART_OPTIONS = [
    "Top lid",
    "Hinges",
    "Bezel",
    "LCD",
    "Keyboard",
    "Motherboard",
]

DISPLAY_COLUMNS = ["Work Order", "Serial Number", "Status", "Parts", "Notes", "Timestamp"]
STATUS_OPTIONS = ["Ordered", "Pending", "Replaced", "Returned", "Complete"]
PART_FILTER_OPTIONS = ["All"] + PART_OPTIONS + ["Other"]

STATUS_COLORS = {
    "Ordered": "#17324d",
    "Pending": "#4a3716",
    "Replaced": "#173d31",
    "Returned": "#4d1f29",
    "Complete": "#30363d",
}

PART_KEYWORDS = {
    "Top lid": [r"\btop lid\b", r"\blid\b", r"\bback cover\b"],
    "Hinges": [r"\bhinge\b", r"\bhinges\b"],
    "Bezel": [r"\bbezel\b"],
    "LCD": [r"\blcd\b", r"\bscreen\b", r"\bdisplay\b"],
    "Keyboard": [r"\bkeyboard\b", r"\bkeys?\b"],
    "Motherboard": [r"\bmotherboard\b", r"\bmainboard\b", r"\bsystem board\b"],
}


def resource_path(relative_path: str) -> str:
    """
    Resolve resource paths for normal execution and PyInstaller one-file builds.
    """
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)


def current_timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def initialize_log(log_file: str) -> None:
    if not os.path.exists(log_file):
        with open(log_file, "w", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow(["Work Order", "Serial Number", "Status", "Notes", "Timestamp"])


def ensure_backup_dir() -> None:
    os.makedirs(BACKUP_DIR, exist_ok=True)


def create_backup(source_file: str) -> Optional[str]:
    if not os.path.exists(source_file):
        return None
    ensure_backup_dir()
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(BACKUP_DIR, f"lcd_log_backup_{stamp}.csv")
    shutil.copy2(source_file, backup_path)
    return backup_path


def build_notes_field(selected_parts: List[str], other_part: str, user_notes: str) -> str:
    parts_text = ", ".join(selected_parts) if selected_parts else "None"
    other_text = other_part.strip() if other_part.strip() else "None"
    notes_text = user_notes.strip() if user_notes.strip() else ""
    return f"Parts: {parts_text} | Other: {other_text} | Notes: {notes_text}"


def parse_notes_field(notes_value: str) -> Tuple[List[str], str, str]:
    if not notes_value:
        return [], "", ""

    pattern = r"^Parts:\s*(.*?)\s*\|\s*Other:\s*(.*?)\s*\|\s*Notes:\s*(.*)$"
    match = re.match(pattern, notes_value, re.DOTALL)
    if match:
        parts_raw = match.group(1).strip()
        other_raw = match.group(2).strip()
        notes_raw = match.group(3).strip()

        selected_parts = []
        if parts_raw and parts_raw.lower() != "none":
            selected_parts = [p.strip() for p in parts_raw.split(",") if p.strip()]

        other_text = "" if other_raw.lower() == "none" else other_raw
        return selected_parts, other_text, notes_raw

    return [], "", notes_value.strip()


def build_parts_display(selected_parts: List[str], other_part: str) -> str:
    parts = list(selected_parts)
    if other_part.strip():
        parts.append(f"Other: {other_part.strip()}")
    return ", ".join(parts) if parts else ""


def normalize_csv_row(row: List[str]) -> Optional[List[str]]:
    if not row:
        return None

    row = [str(cell).strip() for cell in row]
    if not any(row):
        return None
    if len(row) < 4:
        return None

    work_order = row[0]
    serial_number = row[1]
    status = row[2] if row[2] else "Ordered"
    notes = row[3] if len(row) >= 4 else ""
    timestamp = row[4] if len(row) >= 5 and row[4] else current_timestamp()

    if not work_order and not serial_number:
        return None
    if status not in STATUS_OPTIONS:
        status = status if status else "Ordered"

    return [work_order, serial_number, status, notes, timestamp]


def csv_row_to_display_row(row: List[str]) -> Optional[List[str]]:
    if len(row) < 5:
        return None
    selected_parts, other_part, user_notes = parse_notes_field(row[3])
    return [
        row[0],
        row[1],
        row[2],
        build_parts_display(selected_parts, other_part),
        user_notes,
        row[4],
    ]


def detect_parts_from_text(text: str) -> Set[str]:
    detected = set()
    lowered = text.lower()
    for part_name, patterns in PART_KEYWORDS.items():
        for pattern in patterns:
            if re.search(pattern, lowered, re.IGNORECASE):
                detected.add(part_name)
                break
    return detected


class SelectAllPlainTextEdit(QPlainTextEdit):
    """
    Read-only copy box that selects all text on focus and lets Tab move focus.
    """
    def __init__(self, text: str = "", parent=None):
        super().__init__(parent)
        self.setPlainText(text)
        self.setReadOnly(True)
        self.setTabChangesFocus(True)

    def focusInEvent(self, event):
        super().focusInEvent(event)
        self.selectAll()


class EditEntryDialog(QDialog):
    def __init__(self, parent: "MainWindow", original_row: List[str]) -> None:
        super().__init__(parent)
        self.parent_window = parent
        self.original_row = original_row
        self.setWindowTitle("Edit Entry")
        self.resize(760, 390)

        work_order, serial_number, status, notes_value, _timestamp = original_row
        selected_parts_existing, other_part_existing, user_notes_existing = parse_notes_field(notes_value)

        layout = QVBoxLayout(self)
        layout.setSpacing(6)

        form = QGridLayout()
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(10)

        self.work_order_edit = QLineEdit(work_order)
        self.serial_edit = QLineEdit(serial_number)
        self.status_combo = QComboBox()
        self.status_combo.addItems(STATUS_OPTIONS)
        self.status_combo.setCurrentText(status)
        self.other_edit = QLineEdit(other_part_existing)

        form.addWidget(QLabel("Work Order:"), 0, 0)
        form.addWidget(self.work_order_edit, 0, 1)
        form.addWidget(QLabel("Serial Number:"), 0, 2)
        form.addWidget(self.serial_edit, 0, 3)
        form.addWidget(QLabel("Status:"), 1, 0)
        form.addWidget(self.status_combo, 1, 1)
        form.addWidget(QLabel("Other:"), 1, 2)
        form.addWidget(self.other_edit, 1, 3)
        layout.addLayout(form)

        parts_group = QGroupBox("Parts")
        parts_layout = QGridLayout(parts_group)
        self.edit_part_buttons = {}
        for idx, part in enumerate(PART_OPTIONS):
            button = QPushButton(part)
            button.setCheckable(True)
            button.setChecked(part in selected_parts_existing)
            button.setMinimumHeight(46)
            button.setStyleSheet(self.parent_window.part_button_style(button.isChecked()))
            button.toggled.connect(lambda checked, b=button: b.setStyleSheet(self.parent_window.part_button_style(checked)))
            self.edit_part_buttons[part] = button
            parts_layout.addWidget(button, idx // 3, idx % 3)
        layout.addWidget(parts_group)

        notes_group = QGroupBox("Notes")
        notes_layout = QVBoxLayout(notes_group)
        self.notes_edit = QPlainTextEdit(user_notes_existing)
        self.notes_edit.setFixedHeight(60)
        notes_layout.addWidget(self.notes_edit)
        layout.addWidget(notes_group)

        button_row = QHBoxLayout()
        button_row.addStretch(1)
        save_button = QPushButton("Save")
        cancel_button = QPushButton("Cancel")
        save_button.clicked.connect(self.save_edits)
        cancel_button.clicked.connect(self.reject)
        button_row.addWidget(save_button)
        button_row.addWidget(cancel_button)
        layout.addLayout(button_row)

    def save_edits(self) -> None:
        new_work_order = self.work_order_edit.text().strip()
        new_serial = self.serial_edit.text().strip()
        new_status = self.status_combo.currentText().strip()
        new_other = self.other_edit.text().strip()
        new_notes_text = self.notes_edit.toPlainText().strip()
        selected_parts = [part for part, button in self.edit_part_buttons.items() if button.isChecked()]

        if not self.parent_window.validate_entry_fields(new_work_order, new_serial, new_status):
            return

        old_key = (self.original_row[0].strip(), self.original_row[1].strip())
        duplicate_decision = self.parent_window.handle_duplicate_decision(new_work_order, new_serial, exclude_key=old_key)
        if duplicate_decision == "cancel":
            return
        if duplicate_decision == "open":
            self.parent_window.select_existing_entry(new_work_order, new_serial)
            return

        new_notes = build_notes_field(selected_parts, new_other, new_notes_text)
        new_timestamp = current_timestamp()

        rows = self.parent_window.read_all_rows()
        updated_rows = [rows[0]]
        for row in rows[1:]:
            if len(row) >= 5 and row[0].strip() == old_key[0] and row[1].strip() == old_key[1]:
                updated_rows.append([new_work_order, new_serial, new_status, new_notes, new_timestamp])
            else:
                updated_rows.append(row)

        self.parent_window.write_all_rows(updated_rows)
        self.parent_window.display_log()
        self.parent_window.update_dashboard()
        self.parent_window.set_status_message(f"Updated entry {new_work_order} / {new_serial}.")
        self.accept()


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.log_file = LOG_FILE
        self.settings = QSettings("OpenAI", "LenovoCaseTracker")
        initialize_log(self.log_file)

        self.show_complete_entries = True
        self.sort_column: Optional[int] = None
        self.sort_descending = False

        self.setWindowTitle(f"{APP_NAME} {APP_VERSION}")
        self.resize(1500, 920)
        self.setMinimumSize(1280, 820)
        icon_path = resource_path(ICON_FILE)
        if os.path.exists(icon_path):
            icon = QIcon(icon_path)
            self.setWindowIcon(icon)
            QApplication.instance().setWindowIcon(icon)

        self.setStyleSheet("""
            QMainWindow, QWidget {
                background: #111827;
                color: #e5e7eb;
                font-family: "Segoe UI";
                font-size: 9pt;
            }
            QGroupBox {
                font-weight: 600;
                border: 1px solid #2d3748;
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 12px;
                background: #111827;
                color: #f3f4f6;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px 0 6px;
                color: #f9fafb;
            }
            QLineEdit, QComboBox, QPlainTextEdit {
                background: #1f2937;
                color: #f9fafb;
                border: 1px solid #374151;
                border-radius: 6px;
                padding: 4px;
                selection-background-color: #2563eb;
            }
            QComboBox QAbstractItemView {
                background: #1f2937;
                color: #f9fafb;
                border: 1px solid #374151;
                selection-background-color: #2563eb;
            }
            QPushButton {
                min-height: 24px;
                background: #1f2937;
                color: #f9fafb;
                border: 1px solid #374151;
                border-radius: 6px;
                padding: 3px 6px;
            }
            QPushButton:hover {
                background: #273449;
                border: 1px solid #4b5563;
            }
            QPushButton:pressed {
                background: #162132;
            }
            QTableWidget {
                font-size: 9pt;
                background: #0f172a;
                color: #e5e7eb;
                border: 1px solid #334155;
                gridline-color: #1f2937;
                selection-background-color: #1d4ed8;
                selection-color: #ffffff;
                alternate-background-color: #111827;
            }
            QHeaderView::section {
                background: #1f2937;
                color: #f9fafb;
                padding: 4px;
                border: 0px;
                border-right: 1px solid #374151;
                border-bottom: 1px solid #374151;
                font-weight: 600;
            }
            QLabel[role="statValue"] {
                font-size: 13pt;
                font-weight: 700;
                color: #f9fafb;
                background: transparent;
            }
            QLabel[role="statText"] {
                font-size: 8pt;
                color: #9ca3af;
                background: transparent;
            }
            QLabel[role="versionText"] {
                color: #9ca3af;
                font-size: 8pt;
                padding-left: 8px;
            }
            QFrame[card="true"] {
                background: #1f2937;
                border: 1px solid #374151;
                border-radius: 6px;
            }
            QFrame#stat_total { border-left: 3px solid #60a5fa; }
            QFrame#stat_ordered { border-left: 3px solid #3b82f6; }
            QFrame#stat_pending { border-left: 3px solid #f59e0b; }
            QFrame#stat_replaced { border-left: 3px solid #10b981; }
            QFrame#stat_returned { border-left: 3px solid #ef4444; }
            QFrame#stat_complete { border-left: 3px solid #9ca3af; }

            QMenuBar {
                background: #0f172a;
                color: #f9fafb;
            }
            QMenuBar::item:selected {
                background: #1f2937;
            }
            QMenu {
                background: #1f2937;
                color: #f9fafb;
                border: 1px solid #374151;
            }
            QMenu::item:selected {
                background: #2563eb;
            }
            QStatusBar {
                background: #0f172a;
                color: #d1d5db;
                border-top: 1px solid #374151;
            }
        """)

        self.build_menu_bar()

        central = QWidget()
        self.setCentralWidget(central)

        self.main_layout = QVBoxLayout(central)
        self.main_layout.setContentsMargins(12, 12, 12, 10)
        self.main_layout.setSpacing(8)

        self.build_overview()
        self.build_actions()
        self.build_table()
        self.build_manage_row()
        self.build_add_entry()
        self.build_status_bar()

        self.restore_settings()
        self.display_log()
        self.update_dashboard()
        self.work_order_edit.setFocus()

    def build_menu_bar(self) -> None:
        help_menu = self.menuBar().addMenu("Help")
        about_action = QAction("About", self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)

    def build_status_bar(self) -> None:
        status_bar = QStatusBar()
        self.setStatusBar(status_bar)
        self.status_version_label = QLabel(f"{APP_NAME} {APP_VERSION}")
        self.status_version_label.setProperty("role", "versionText")
        self.statusBar().addPermanentWidget(self.status_version_label)

    def set_status_message(self, message: str, timeout: int = 5000) -> None:
        self.statusBar().showMessage(message, timeout)

    def part_button_style(self, checked: bool) -> str:
        if checked:
            return """
                QPushButton {
                    background: #1d4ed8;
                    color: #ffffff;
                    border: 1px solid #60a5fa;
                    border-radius: 6px;
                    padding: 5px 8px;
                    font-weight: 600;
                }
                QPushButton:hover {
                    background: #2563eb;
                }
            """
        return """
            QPushButton {
                min-height: 24px;
                background: #1f2937;
                color: #f9fafb;
                border: 1px solid #374151;
                border-radius: 6px;
                padding: 5px 8px;
            }
            QPushButton:hover {
                background: #273449;
                border: 1px solid #4b5563;
            }
        """

    def build_overview(self) -> None:
        overview = QGroupBox("Overview")
        overview.setContentsMargins(6, 10, 6, 6)
        layout = QVBoxLayout(overview)
        layout.setSpacing(10)

        search_row = QHBoxLayout()
        search_row.addWidget(QLabel("Search:"))
        self.search_edit = QLineEdit()
        self.search_edit.returnPressed.connect(self.search_log)
        search_button = QPushButton("Search")
        search_button.setFixedSize(72, 26)
        search_button.clicked.connect(self.search_log)
        search_row.addWidget(self.search_edit, 1)
        search_row.addWidget(search_button)
        layout.addLayout(search_row)

        stats_row = QHBoxLayout()
        stats_row.setSpacing(8)
        self.stat_labels = {}
        for key in ["Total", "Ordered", "Pending", "Replaced", "Returned", "Complete"]:
            card = QFrame()
            card.setObjectName(f"stat_{key.lower()}")
            card.setProperty("card", True)
            card_layout = QVBoxLayout(card)
            card_layout.setContentsMargins(10, 8, 10, 8)
            card_layout.setSpacing(1)

            value = QLabel("0")
            value.setProperty("role", "statValue")
            text = QLabel(key)
            text.setProperty("role", "statText")
            value.setAlignment(Qt.AlignmentFlag.AlignCenter)
            text.setAlignment(Qt.AlignmentFlag.AlignCenter)

            card.setFixedWidth(78)
            card_layout.addWidget(value)
            card_layout.addWidget(text)
            stats_row.addWidget(card)
            self.stat_labels[key] = value

        stats_row.addStretch(1)
        layout.addLayout(stats_row)
        self.main_layout.addWidget(overview)

    def build_actions(self) -> None:
        actions = QGroupBox("Actions")
        actions.setContentsMargins(6, 10, 6, 6)
        row = QHBoxLayout(actions)
        row.setSpacing(8)

        refresh_button = QPushButton("Refresh Log")
        refresh_button.clicked.connect(self.handle_refresh)
        export_button = QPushButton("Export to CSV")
        export_button.clicked.connect(self.export_to_csv)
        import_button = QPushButton("Import from CSV")
        import_button.clicked.connect(self.import_from_csv)
        self.toggle_complete_button = QPushButton("Hide Complete Entries")
        self.toggle_complete_button.clicked.connect(self.toggle_complete_entries)
        script_button = QPushButton("LCD Script")
        script_button.clicked.connect(self.open_lcd_script_window)
        copy_summary_button = QPushButton("Copy Case Summary")
        copy_summary_button.clicked.connect(self.copy_case_summary)
        about_button = QPushButton("About")
        about_button.clicked.connect(self.show_about_dialog)

        for button in [
            refresh_button, export_button, import_button,
            self.toggle_complete_button, script_button, copy_summary_button, about_button
        ]:
            button.setFixedHeight(30)
            row.addWidget(button)

        row.addStretch(1)
        self.main_layout.addWidget(actions)

    def build_add_entry(self) -> None:
        group = QGroupBox("Add New Entry")
        group.setContentsMargins(6, 10, 6, 6)
        layout = QVBoxLayout(group)
        layout.setSpacing(8)

        top_grid = QGridLayout()
        top_grid.setHorizontalSpacing(10)
        top_grid.setVerticalSpacing(6)

        self.work_order_edit = QLineEdit()
        self.work_order_edit.setMinimumWidth(280)
        self.serial_edit = QLineEdit()
        self.serial_edit.setMinimumWidth(280)
        self.status_combo = QComboBox()
        self.status_combo.setMinimumWidth(180)
        self.status_combo.addItems(STATUS_OPTIONS)
        self.status_combo.setCurrentIndex(-1)
        self.other_edit = QLineEdit()
        self.other_edit.setMinimumWidth(280)

        top_grid.addWidget(QLabel("Work Order:"), 0, 0)
        top_grid.addWidget(self.work_order_edit, 0, 1)
        top_grid.addWidget(QLabel("Serial Number:"), 0, 2)
        top_grid.addWidget(self.serial_edit, 0, 3)
        top_grid.addWidget(QLabel("Status:"), 1, 0)
        top_grid.addWidget(self.status_combo, 1, 1)
        top_grid.addWidget(QLabel("Other:"), 1, 2)
        top_grid.addWidget(self.other_edit, 1, 3)
        top_grid.setColumnStretch(1, 1)
        top_grid.setColumnStretch(3, 1)
        layout.addLayout(top_grid)

        parts_group = QGroupBox("Parts")
        parts_layout = QGridLayout(parts_group)
        self.part_buttons = {}
        for idx, part in enumerate(PART_OPTIONS):
            button = QPushButton(part)
            button.setCheckable(True)
            button.setMinimumHeight(46)
            button.setStyleSheet(self.part_button_style(False))
            button.toggled.connect(lambda checked, b=button: b.setStyleSheet(self.part_button_style(checked)))
            self.part_buttons[part] = button
            parts_layout.addWidget(button, idx // 3, idx % 3)
        layout.addWidget(parts_group)

        notes_group = QGroupBox("Notes")
        notes_layout = QVBoxLayout(notes_group)
        self.notes_edit = QPlainTextEdit()
        self.notes_edit.setFixedHeight(56)
        notes_layout.addWidget(self.notes_edit)
        layout.addWidget(notes_group)

        button_row = QHBoxLayout()
        add_button = QPushButton("Add Entry")
        add_button.clicked.connect(self.handle_add_entry)
        clear_button = QPushButton("Clear Form")
        clear_button.clicked.connect(self.clear_add_entry_form)
        paste_button = QPushButton("Paste Lenovo Info")
        paste_button.clicked.connect(self.fill_from_clipboard)

        for button in (add_button, clear_button, paste_button):
            button.setFixedHeight(30)
            button_row.addWidget(button)
        button_row.addStretch(1)
        layout.addLayout(button_row)

        self.main_layout.addWidget(group)

    def build_manage_row(self) -> None:
        group = QGroupBox("Update / Manage Entries")
        group.setContentsMargins(6, 10, 6, 6)
        row = QHBoxLayout(group)
        row.setSpacing(8)

        row.addWidget(QLabel("New Status:"))
        self.update_status_combo = QComboBox()
        self.update_status_combo.setFixedWidth(140)
        self.update_status_combo.addItems(STATUS_OPTIONS)
        self.update_status_combo.setCurrentIndex(-1)

        update_button = QPushButton("Update Status")
        update_button.clicked.connect(self.handle_update_status)
        delete_button = QPushButton("Delete Entry")
        delete_button.clicked.connect(self.handle_delete_entry)
        edit_button = QPushButton("Edit Entry")
        edit_button.clicked.connect(self.handle_edit_entry)

        update_button.setFixedHeight(26)
        delete_button.setFixedHeight(26)
        edit_button.setFixedHeight(26)
        row.addWidget(self.update_status_combo)
        row.addWidget(update_button)
        row.addWidget(delete_button)
        row.addWidget(edit_button)
        row.addStretch(1)

        self.main_layout.addWidget(group)

    def build_table(self) -> None:
        group = QGroupBox("Log Entries")
        group.setContentsMargins(6, 10, 6, 6)
        layout = QVBoxLayout(group)

        filter_row = QHBoxLayout()
        filter_row.setSpacing(6)
        filter_row.addWidget(QLabel("Status Filter:"))
        self.status_filter_combo = QComboBox()
        self.status_filter_combo.setFixedWidth(96)
        self.status_filter_combo.addItems(["All"] + STATUS_OPTIONS)
        self.status_filter_combo.currentTextChanged.connect(self.on_filter_changed)
        filter_row.addWidget(self.status_filter_combo)

        filter_row.addWidget(QLabel("Part Filter:"))
        self.part_filter_combo = QComboBox()
        self.part_filter_combo.setFixedWidth(96)
        self.part_filter_combo.addItems(PART_FILTER_OPTIONS)
        self.part_filter_combo.currentTextChanged.connect(self.on_filter_changed)
        filter_row.addWidget(self.part_filter_combo)

        self.filtered_count_label = QLabel("Showing 0 entries")
        self.filtered_count_label.setStyleSheet("color: #9ca3af;")
        filter_row.addStretch(1)
        filter_row.addWidget(self.filtered_count_label)
        layout.addLayout(filter_row)

        self.table = QTableWidget(0, len(DISPLAY_COLUMNS))
        self.table.setHorizontalHeaderLabels(DISPLAY_COLUMNS)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.setSortingEnabled(False)
        self.table.setShowGrid(True)
        self.table.setMinimumHeight(200)
        self.table.setMaximumHeight(260)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        header.sectionClicked.connect(self.sort_treeview)

        self.table.itemDoubleClicked.connect(lambda _item: self.open_edit_window_for_selection())
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

        layout.addWidget(self.table)
        self.main_layout.addWidget(group, 1)

    def restore_settings(self) -> None:
        geometry = self.settings.value("geometry")
        if geometry:
            self.restoreGeometry(geometry)

        # Keep older saved window sizes from forcing a much larger layout.
        if self.width() > 1700 or self.height() > 1100:
            self.resize(1500, 920)

        self.show_complete_entries = self.settings.value("show_complete_entries", True, type=bool)
        self.toggle_complete_button.setText(
            "Show Complete Entries" if not self.show_complete_entries else "Hide Complete Entries"
        )

        status_filter = self.settings.value("status_filter", "All")
        part_filter = self.settings.value("part_filter", "All")
        if status_filter in ["All"] + STATUS_OPTIONS:
            self.status_filter_combo.setCurrentText(status_filter)
        if part_filter in PART_FILTER_OPTIONS:
            self.part_filter_combo.setCurrentText(part_filter)

    def save_settings(self) -> None:
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("show_complete_entries", self.show_complete_entries)
        self.settings.setValue("status_filter", self.status_filter_combo.currentText())
        self.settings.setValue("part_filter", self.part_filter_combo.currentText())

    def closeEvent(self, event) -> None:
        self.save_settings()
        super().closeEvent(event)

    def read_all_rows(self) -> List[List[str]]:
        initialize_log(self.log_file)
        with open(self.log_file, "r", newline="", encoding="utf-8") as file:
            return list(csv.reader(file))

    def write_all_rows(self, rows: List[List[str]]) -> None:
        with open(self.log_file, "w", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerows(rows)

    def validate_entry_fields(self, work_order: str, serial_number: str, status: str) -> bool:
        if len(work_order) > 10:
            QMessageBox.warning(self, "Input Error", "Work Order number must be 10 digits or less.")
            return False
        if len(serial_number) > 8:
            QMessageBox.warning(self, "Input Error", "Serial Number must be 8 digits or less.")
            return False
        if not work_order or not serial_number or not status:
            QMessageBox.warning(self, "Input Error", "Please fill in Work Order, Serial Number, and Status.")
            return False
        return True

    def entry_exists(self, work_order: str, serial_number: str, exclude_key=None) -> bool:
        rows = self.read_all_rows()
        for row in rows[1:]:
            if len(row) < 2:
                continue
            key = (row[0].strip(), row[1].strip())
            if exclude_key and key == exclude_key:
                continue
            if key == (work_order.strip(), serial_number.strip()):
                return True
        return False

    def handle_duplicate_decision(self, work_order: str, serial_number: str, exclude_key=None) -> str:
        if not self.entry_exists(work_order, serial_number, exclude_key=exclude_key):
            return "add"

        box = QMessageBox(self)
        box.setWindowTitle("Duplicate Entry")
        box.setIcon(QMessageBox.Icon.Warning)
        box.setText("An entry with this Work Order and Serial Number already exists.")
        box.setInformativeText("Choose what you want to do next.")
        open_button = box.addButton("Open Existing", QMessageBox.ButtonRole.ActionRole)
        add_button = box.addButton("Add Anyway", QMessageBox.ButtonRole.AcceptRole)
        box.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
        box.setDefaultButton(add_button)
        box.exec()

        clicked = box.clickedButton()
        if clicked == open_button:
            return "open"
        if clicked == add_button:
            return "add"
        return "cancel"

    def select_existing_entry(self, work_order: str, serial_number: str) -> None:
        self.status_filter_combo.setCurrentText("All")
        self.part_filter_combo.setCurrentText("All")
        self.display_log()

        for row in range(self.table.rowCount()):
            work_item = self.table.item(row, 0)
            serial_item = self.table.item(row, 1)
            if work_item and serial_item and work_item.text().strip() == work_order.strip() and serial_item.text().strip() == serial_number.strip():
                self.table.selectRow(row)
                self.table.scrollToItem(self.table.item(row, 0), QAbstractItemView.ScrollHint.PositionAtCenter)
                self.set_status_message(f"Selected existing entry {work_order} / {serial_number}.")
                return

    def selected_parts(self) -> List[str]:
        return [part for part, button in self.part_buttons.items() if button.isChecked()]

    def selected_row_keys(self) -> List[Tuple[str, str]]:
        keys = []
        for index in self.table.selectionModel().selectedRows():
            row = index.row()
            work_order_item = self.table.item(row, 0)
            serial_item = self.table.item(row, 1)
            if work_order_item and serial_item:
                keys.append((work_order_item.text().strip(), serial_item.text().strip()))
        return keys

    def clear_add_entry_form(self) -> None:
        self.work_order_edit.clear()
        self.serial_edit.clear()
        self.status_combo.setCurrentIndex(-1)
        self.other_edit.clear()
        self.notes_edit.clear()
        for button in self.part_buttons.values():
            button.setChecked(False)
        self.work_order_edit.setFocus()
        self.set_status_message("Form cleared.")

    def handle_add_entry(self) -> None:
        work_order = self.work_order_edit.text().strip()
        serial_number = self.serial_edit.text().strip()
        status = self.status_combo.currentText().strip()
        other_part = self.other_edit.text().strip()
        user_notes = self.notes_edit.toPlainText().strip()
        selected_parts = self.selected_parts()

        if not self.validate_entry_fields(work_order, serial_number, status):
            return

        duplicate_decision = self.handle_duplicate_decision(work_order, serial_number)
        if duplicate_decision == "cancel":
            return
        if duplicate_decision == "open":
            self.select_existing_entry(work_order, serial_number)
            return

        notes = build_notes_field(selected_parts, other_part, user_notes)
        with open(self.log_file, "a", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow([work_order, serial_number, status, notes, current_timestamp()])

        self.display_log()
        self.update_dashboard()
        self.clear_add_entry_form()
        self.set_status_message(f"Added entry {work_order} / {serial_number}.")

    def confirm_bulk_action(self, title: str, action_text: str, count: int) -> bool:
        result = QMessageBox.question(
            self,
            title,
            f"{action_text} {count} selected {'entry' if count == 1 else 'entries'}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return result == QMessageBox.StandardButton.Yes

    def handle_update_status(self) -> None:
        selected = self.selected_row_keys()
        if not selected:
            QMessageBox.warning(self, "Selection Error", "Please select at least one entry to update.")
            return

        new_status = self.update_status_combo.currentText().strip()
        if not new_status:
            QMessageBox.warning(self, "Input Error", "Please select a new status.")
            return

        if len(selected) > 1:
            if not self.confirm_bulk_action("Confirm Status Update", f"Update status to '{new_status}' for", len(selected)):
                return

        rows = self.read_all_rows()
        updated_rows = [rows[0]]
        timestamp = current_timestamp()
        key_set = set(selected)

        for row in rows[1:]:
            if len(row) >= 5 and (row[0].strip(), row[1].strip()) in key_set:
                row[2] = new_status
                row[4] = timestamp
            updated_rows.append(row)

        self.write_all_rows(updated_rows)
        self.display_log()
        self.update_dashboard()
        self.update_status_combo.setCurrentIndex(-1)
        self.set_status_message(f"Updated {len(selected)} entr{'y' if len(selected) == 1 else 'ies'} to {new_status}.")

    def handle_delete_entry(self) -> None:
        selected = self.selected_row_keys()
        if not selected:
            QMessageBox.warning(self, "Selection Error", "Please select at least one entry to delete.")
            return

        if not self.confirm_bulk_action("Confirm Deletion", "Delete", len(selected)):
            return

        rows = self.read_all_rows()
        key_set = set(selected)
        kept_rows = [rows[0]]
        for row in rows[1:]:
            if len(row) >= 5 and (row[0].strip(), row[1].strip()) not in key_set:
                kept_rows.append(row)

        self.write_all_rows(kept_rows)
        self.display_log()
        self.update_dashboard()
        self.set_status_message(f"Deleted {len(selected)} entr{'y' if len(selected) == 1 else 'ies'}.")

    def open_edit_window_for_selection(self) -> None:
        selected = self.selected_row_keys()
        if not selected:
            QMessageBox.warning(self, "Selection Error", "Please select an entry to edit.")
            return

        target_key = selected[0]
        rows = self.read_all_rows()
        for row in rows[1:]:
            if len(row) >= 5 and (row[0].strip(), row[1].strip()) == target_key:
                dialog = EditEntryDialog(self, row)
                dialog.exec()
                return

        QMessageBox.warning(self, "Data Error", "Could not locate the selected entry in the CSV.")

    def handle_edit_entry(self) -> None:
        self.open_edit_window_for_selection()

    def search_log(self) -> None:
        query = self.search_edit.text().strip().lower()
        if not query:
            return

        for row in range(self.table.rowCount()):
            values = []
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                values.append(item.text().lower() if item else "")
            searchable_text = " ".join(values)
            if query in searchable_text:
                self.table.selectRow(row)
                self.table.scrollToItem(self.table.item(row, 0), QAbstractItemView.ScrollHint.PositionAtCenter)
                self.set_status_message("Search match selected.")
                return

        QMessageBox.information(self, "No Match", "No matching entries found.")
        self.set_status_message("No search match found.")

    def copy_serial_number(self) -> None:
        selected = self.selected_row_keys()
        if not selected:
            QMessageBox.warning(self, "Selection Error", "Please select an entry to copy.")
            return
        QApplication.clipboard().setText(selected[0][1])
        QMessageBox.information(self, "Copied", f"Serial Number '{selected[0][1]}' copied to clipboard.")
        self.set_status_message("Serial number copied to clipboard.")

    def copy_work_order_number(self) -> None:
        selected = self.selected_row_keys()
        if not selected:
            QMessageBox.warning(self, "Selection Error", "Please select an entry to copy.")
            return
        QApplication.clipboard().setText(selected[0][0])
        QMessageBox.information(self, "Copied", f"Work Order '{selected[0][0]}' copied to clipboard.")
        self.set_status_message("Work order copied to clipboard.")

    def copy_case_summary(self) -> None:
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "Selection Error", "Please select at least one entry.")
            return

        summaries = []
        for index in selected_rows:
            row = index.row()
            values = [self.table.item(row, col).text() if self.table.item(row, col) else "" for col in range(6)]
            summary = f"WO {values[0]} | Serial {values[1]} | Parts: {values[3] or 'None'} | Status: {values[2]}"
            if values[4]:
                summary += f" | Notes: {values[4]}"
            summaries.append(summary)

        QApplication.clipboard().setText("\n".join(summaries))
        QMessageBox.information(self, "Copied", "Case summary copied to clipboard.")
        self.set_status_message(f"Copied case summary for {len(selected_rows)} entr{'y' if len(selected_rows) == 1 else 'ies'}.")

    def fill_from_clipboard(self) -> None:
        try:
            clipboard_text = ""
            if pyperclip is not None:
                clipboard_text = pyperclip.paste() or ""
            if not clipboard_text:
                clipboard_text = QApplication.clipboard().text() or ""

            if not clipboard_text:
                raise RuntimeError("Clipboard is empty or unavailable.")

            parsed = self.parse_lenovo_clipboard_minimal(clipboard_text)
            detected_parts = detect_parts_from_text(clipboard_text)

            if parsed["WO"]:
                self.work_order_edit.setText(parsed["WO"])
            if parsed["Serial"]:
                self.serial_edit.setText(parsed["Serial"])

            self.status_combo.setCurrentText("Ordered")
            self.other_edit.clear()
            self.notes_edit.clear()

            for part_name, button in self.part_buttons.items():
                button.setChecked(part_name in detected_parts)

            if detected_parts:
                QMessageBox.information(
                    self,
                    "Success",
                    "Fields populated from clipboard.\nDetected parts: " + ", ".join(sorted(detected_parts)),
                )
            else:
                QMessageBox.information(self, "Success", "Fields populated from clipboard.")

            self.set_status_message("Clipboard data imported into form.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to process clipboard: {str(e)}")

    def parse_lenovo_clipboard_minimal(self, clipboard_text: str):
        data = {"WO": None, "Serial": None}

        wo_match = re.search(r"WO:\s*(\d{8,10})", clipboard_text)
        if wo_match:
            data["WO"] = wo_match.group(1)

        serial_match = re.search(r"Serial\s*Number\s*[\s:\-]*\n([A-Z0-9]{7,10})", clipboard_text, re.IGNORECASE)
        if not serial_match:
            serial_match = re.search(r"Serial\s*Number\s*[:\s]*([A-Z0-9]{7,10})", clipboard_text, re.IGNORECASE)

        if serial_match:
            data["Serial"] = serial_match.group(1)

        return data

    def handle_refresh(self) -> None:
        self.display_log()
        self.update_dashboard()
        self.set_status_message("Log refreshed.")

    def export_to_csv(self) -> None:
        path, _ = QFileDialog.getSaveFileName(self, "Export CSV", "", "CSV files (*.csv)")
        if not path:
            return

        with open(path, "w", newline="", encoding="utf-8") as file:
            writer = csv.writer(file)
            writer.writerow(["Work Order", "Serial Number", "Status", "Notes", "Timestamp"])
            rows = self.read_all_rows()
            for row in rows[1:]:
                writer.writerow(row)

        QMessageBox.information(self, "Export Successful", f"Log exported successfully to {path}")
        self.set_status_message("CSV exported successfully.")

    def import_from_csv(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Import CSV", "", "CSV files (*.csv)")
        if not path:
            return

        try:
            backup_path = create_backup(self.log_file)

            with open(path, "r", newline="", encoding="utf-8-sig") as file:
                reader = csv.reader(file)
                rows = list(reader)

            if not rows:
                self.write_all_rows([["Work Order", "Serial Number", "Status", "Notes", "Timestamp"]])
                self.display_log()
                self.update_dashboard()
                msg = "Imported empty CSV. Log cleared to header only."
                if backup_path:
                    msg += f"\nBackup created: {backup_path}"
                QMessageBox.information(self, "Import Successful", msg)
                self.set_status_message("Imported empty CSV; current log backed up.")
                return

            start_index = 0
            first_row_joined = ",".join(cell.strip().lower() for cell in rows[0])
            if "work order" in first_row_joined and "serial number" in first_row_joined:
                start_index = 1

            imported_rows = []
            for row in rows[start_index:]:
                normalized = normalize_csv_row(row)
                if normalized:
                    imported_rows.append(normalized)

            final_rows = [["Work Order", "Serial Number", "Status", "Notes", "Timestamp"]]
            final_rows.extend(imported_rows)
            self.write_all_rows(final_rows)
            self.display_log()
            self.update_dashboard()

            msg = f"Imported {len(imported_rows)} {'entry' if len(imported_rows) == 1 else 'entries'}."
            if backup_path:
                msg += f"\nBackup created: {backup_path}"
            QMessageBox.information(self, "Import Successful", msg)
            self.set_status_message("CSV imported successfully. Backup created before import.")
        except Exception as e:
            QMessageBox.critical(self, "Import Error", f"Failed to import CSV:\n{str(e)}")

    def update_dashboard(self) -> None:
        total_entries = 0
        counts = {status: 0 for status in STATUS_OPTIONS}

        rows = self.read_all_rows()
        for row in rows[1:]:
            if len(row) >= 3:
                total_entries += 1
                if row[2] in counts:
                    counts[row[2]] += 1

        self.stat_labels["Total"].setText(str(total_entries))
        self.stat_labels["Ordered"].setText(str(counts["Ordered"]))
        self.stat_labels["Pending"].setText(str(counts["Pending"]))
        self.stat_labels["Replaced"].setText(str(counts["Replaced"]))
        self.stat_labels["Returned"].setText(str(counts["Returned"]))
        self.stat_labels["Complete"].setText(str(counts["Complete"]))

    def on_filter_changed(self) -> None:
        self.display_log()
        self.set_status_message("Filters updated.")

    def row_matches_filters(self, display_row: List[str]) -> bool:
        status_filter = self.status_filter_combo.currentText()
        part_filter = self.part_filter_combo.currentText()

        if status_filter != "All" and display_row[2] != status_filter:
            return False

        if part_filter != "All":
            parts_text = display_row[3]
            if part_filter == "Other":
                if "Other:" not in parts_text:
                    return False
            elif part_filter not in parts_text:
                return False

        return True

    def display_log(self) -> None:
        rows = self.read_all_rows()
        display_rows = []

        for row in rows[1:]:
            if len(row) >= 5:
                if not self.show_complete_entries and row[2] == "Complete":
                    continue
                display_row = csv_row_to_display_row(row)
                if display_row and self.row_matches_filters(display_row):
                    display_rows.append(display_row)

        self.table.setRowCount(len(display_rows))
        for row_idx, row_values in enumerate(display_rows):
            status = row_values[2]
            bg_color = QColor(STATUS_COLORS.get(status, "#0f172a"))

            for col_idx, value in enumerate(row_values):
                item = QTableWidgetItem(value)
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setBackground(QBrush(bg_color))
                self.table.setItem(row_idx, col_idx, item)

        self.filtered_count_label.setText(f"Showing {len(display_rows)} entries")

        if self.sort_column is not None:
            order = Qt.SortOrder.DescendingOrder if self.sort_descending else Qt.SortOrder.AscendingOrder
            self.table.sortItems(self.sort_column, order)

    def sort_treeview(self, column_index: int) -> None:
        if self.sort_column == column_index:
            self.sort_descending = not self.sort_descending
        else:
            self.sort_column = column_index
            self.sort_descending = False

        order = Qt.SortOrder.DescendingOrder if self.sort_descending else Qt.SortOrder.AscendingOrder
        self.table.sortItems(column_index, order)

    def toggle_complete_entries(self) -> None:
        self.show_complete_entries = not self.show_complete_entries
        self.toggle_complete_button.setText(
            "Show Complete Entries" if not self.show_complete_entries else "Hide Complete Entries"
        )
        self.display_log()
        self.set_status_message("Complete entry visibility updated.")

    def show_context_menu(self, position) -> None:
        index = self.table.indexAt(position)
        if index.isValid() and not self.table.selectionModel().isRowSelected(index.row(), index.parent()):
            self.table.selectRow(index.row())

        menu = QMenu(self)

        copy_serial_action = QAction("Copy Serial Number", self)
        copy_serial_action.triggered.connect(self.copy_serial_number)
        menu.addAction(copy_serial_action)

        copy_wo_action = QAction("Copy Work Order Number", self)
        copy_wo_action.triggered.connect(self.copy_work_order_number)
        menu.addAction(copy_wo_action)

        copy_summary_action = QAction("Copy Case Summary", self)
        copy_summary_action.triggered.connect(self.copy_case_summary)
        menu.addAction(copy_summary_action)

        menu.addSeparator()

        edit_action = QAction("Edit Entry", self)
        edit_action.triggered.connect(self.handle_edit_entry)
        menu.addAction(edit_action)

        delete_action = QAction("Delete Entry", self)
        delete_action.triggered.connect(self.handle_delete_entry)
        menu.addAction(delete_action)

        menu.exec(self.table.viewport().mapToGlobal(position))

    def show_about_dialog(self) -> None:
        text = (
            f"{APP_NAME} {APP_VERSION}\n\n"
            "A lightweight desktop utility for tracking Lenovo case parts, statuses, and notes.\n\n"
            f"CSV file: {os.path.abspath(self.log_file)}\n"
            f"Backups folder: {os.path.abspath(BACKUP_DIR)}"
        )
        QMessageBox.information(self, f"About {APP_NAME}", text)

    def open_lcd_script_window(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("LCD Script")
        dialog.resize(430, 200)
        layout = QVBoxLayout(dialog)

        text1 = SelectAllPlainTextEdit("Student dropped laptop which cracked the LCD.")
        text2 = SelectAllPlainTextEdit("Replaced LCD. Laptop boots and displays image correctly.")
        text1.setFixedHeight(52)
        text2.setFixedHeight(52)
        text1.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        text2.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

        close_button = QPushButton("Close")
        close_button.setFixedHeight(26)
        close_button.clicked.connect(dialog.accept)
        close_button.setAutoDefault(True)
        close_button.setDefault(True)
        close_button.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

        layout.addWidget(text1)
        layout.addWidget(text2)

        button_row = QHBoxLayout()
        button_row.addStretch(1)
        button_row.addWidget(close_button)
        layout.addLayout(button_row)

        QWidget.setTabOrder(text1, text2)
        QWidget.setTabOrder(text2, close_button)

        text1.setFocus()
        text1.selectAll()

        dialog.exec()


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setOrganizationName("OpenAI")
    app.setStyle("Fusion")

    icon_path = resource_path(ICON_FILE)
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
