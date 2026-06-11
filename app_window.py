from __future__ import annotations

from datetime import datetime

from PySide6.QtCore import QSize, Qt, Signal
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QListView,
    QMainWindow,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from runtime_paths import resource_path

APP_TITLE = "TransareApp"
MONTH_OPTIONS = [
    "ianuarie",
    "februarie",
    "martie",
    "aprilie",
    "mai",
    "iunie",
    "iulie",
    "august",
    "septembrie",
    "octombrie",
    "noiembrie",
    "decembrie",
]
TEMPLATE_MODE_PREDEFINED_LABEL = "Template predefinit"
TEMPLATE_MODE_PREVIOUS_SHEET_LABEL = "Copiază din sheet anterior"

COLOR_BG = "#F3F4F6"
COLOR_SURFACE = "#FFFFFF"
COLOR_SURFACE_ALT = "#F8FAFC"
COLOR_BORDER = "#D1D5DB"
COLOR_TEXT = "#111827"
COLOR_MUTED = "#6B7280"
COLOR_PRIMARY = "#2563EB"
COLOR_PRIMARY_HOVER = "#1D4ED8"
COLOR_PRIMARY_PRESSED = "#1E40AF"
COLOR_SUCCESS = "#166534"
COLOR_ERROR = "#B91C1C"
COLOR_WARNING = "#B45309"

APP_QSS = f"""
QWidget#root {{
    background: {COLOR_BG};
    color: {COLOR_TEXT};
}}
QFrame#panel {{
    background: {COLOR_SURFACE};
    border: 1px solid {COLOR_BORDER};
    border-radius: 8px;
}}
QLabel {{
    background: transparent;
    color: {COLOR_TEXT};
}}
QLabel#appTitle {{
    font-size: 22px;
    font-weight: 800;
}}
QLabel#sectionTitle {{
    color: {COLOR_MUTED};
    font-size: 12px;
    font-weight: 800;
    text-transform: uppercase;
}}
QLabel#fieldLabel {{
    color: {COLOR_TEXT};
    font-weight: 700;
}}
QLabel#statusBadge {{
    border-radius: 7px;
    color: white;
    font-size: 11px;
    font-weight: 800;
    padding: 5px 10px;
}}
QLineEdit, QComboBox {{
    background: white;
    border: 1px solid {COLOR_BORDER};
    border-radius: 6px;
    color: {COLOR_TEXT};
    min-height: 34px;
    padding: 0 10px;
    selection-background-color: {COLOR_PRIMARY};
    selection-color: white;
}}
QLineEdit {{
    padding-right: 28px;
}}
QLineEdit[readOnly="true"] {{
    background: #F9FAFB;
}}
QLineEdit:focus, QComboBox:focus {{
    border-color: {COLOR_PRIMARY};
}}
QLineEdit:disabled, QComboBox:disabled {{
    background: #F3F4F6;
    color: {COLOR_MUTED};
}}
QComboBox {{
    padding-right: 28px;
}}
QComboBox::drop-down {{
    border: 0;
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 26px;
}}
QComboBox::down-arrow {{
    width: 0;
    height: 0;
}}
QComboBox QAbstractItemView {{
    background: white;
    border: 1px solid {COLOR_BORDER};
    color: {COLOR_TEXT};
    selection-background-color: {COLOR_PRIMARY};
    selection-color: white;
    outline: 0;
}}
QPushButton {{
    background: {COLOR_SURFACE_ALT};
    border: 1px solid {COLOR_BORDER};
    border-radius: 6px;
    color: {COLOR_TEXT};
    min-height: 34px;
    padding: 0 14px;
    font-weight: 700;
}}
QPushButton:hover {{
    border-color: {COLOR_PRIMARY};
    background: #EFF6FF;
}}
QPushButton:pressed {{
    background: #DBEAFE;
}}
QPushButton#primaryButton {{
    background: {COLOR_PRIMARY};
    border-color: {COLOR_PRIMARY};
    color: white;
    min-height: 38px;
}}
QPushButton#primaryButton:hover {{
    background: {COLOR_PRIMARY_HOVER};
}}
QPushButton#primaryButton:pressed {{
    background: {COLOR_PRIMARY_PRESSED};
}}
QPushButton#secondaryButton {{
    background: white;
    color: {COLOR_TEXT};
}}
QPushButton:disabled {{
    background: #F3F4F6;
    border-color: #E5E7EB;
    color: #9CA3AF;
}}
QPlainTextEdit {{
    background: #F9FAFB;
    border: 1px solid {COLOR_BORDER};
    border-radius: 6px;
    color: {COLOR_TEXT};
    font-family: "Cascadia Mono", "Consolas", monospace;
    font-size: 12px;
    padding: 8px;
    selection-background-color: {COLOR_PRIMARY};
    selection-color: white;
}}
QProgressBar {{
    border: 1px solid {COLOR_BORDER};
    border-radius: 4px;
    background: #EEF2F7;
    color: {COLOR_TEXT};
    min-height: 10px;
    max-height: 10px;
}}
QProgressBar::chunk {{
    background: {COLOR_PRIMARY};
    border-radius: 4px;
}}
QDialog#alertDialog {{
    background: {COLOR_BG};
}}
QFrame#alertPanel {{
    background: {COLOR_SURFACE};
    border: 1px solid {COLOR_BORDER};
    border-radius: 8px;
}}
QLabel#alertSeverity {{
    color: {COLOR_MUTED};
    font-size: 12px;
    font-weight: 800;
    text-transform: uppercase;
}}
QLabel#alertMessage {{
    color: {COLOR_TEXT};
    font-size: 13px;
    font-weight: 700;
}}
QPushButton#warningButton {{
    background: {COLOR_WARNING};
    border-color: {COLOR_WARNING};
    color: white;
}}
QPushButton#warningButton:hover {{
    background: #92400E;
    border-color: #92400E;
}}
QPushButton#warningButton:pressed {{
    background: #78350F;
    border-color: #78350F;
}}
QPushButton#dangerButton {{
    background: {COLOR_ERROR};
    border-color: {COLOR_ERROR};
    color: white;
}}
QPushButton#dangerButton:hover {{
    background: #991B1B;
    border-color: #991B1B;
}}
QPushButton#dangerButton:pressed {{
    background: #7F1D1D;
    border-color: #7F1D1D;
}}
"""


class AlertDialog(QDialog):
    def __init__(
        self,
        parent: QWidget | None,
        message: str,
        kind: str,
        *,
        confirm: bool = False,
        title: str = APP_TITLE,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("alertDialog")
        self.setWindowTitle(title)
        self.setModal(True)
        self.setMinimumWidth(460)
        self.setStyleSheet(APP_QSS)
        if parent is not None and not parent.windowIcon().isNull():
            self.setWindowIcon(parent.windowIcon())

        accent = COLOR_ERROR if kind == "error" else COLOR_WARNING
        severity = "EROARE" if kind == "error" else "ATENȚIE"

        root = QVBoxLayout(self)
        root.setContentsMargins(14, 14, 14, 14)
        root.setSpacing(0)

        panel = QFrame(objectName="alertPanel")
        root.addWidget(panel)

        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(16, 15, 16, 14)
        panel_layout.setSpacing(14)

        content_row = QHBoxLayout()
        content_row.setSpacing(12)

        marker = QLabel("!")
        marker.setAlignment(Qt.AlignCenter)
        marker.setFixedSize(34, 34)
        marker.setStyleSheet(
            f"background: {accent}; color: white; border-radius: 17px; "
            "font-size: 20px; font-weight: 900;"
        )

        text_column = QVBoxLayout()
        text_column.setSpacing(5)
        severity_label = QLabel(severity)
        severity_label.setObjectName("alertSeverity")
        message_label = QLabel(message)
        message_label.setObjectName("alertMessage")
        message_label.setWordWrap(True)
        message_label.setTextFormat(Qt.PlainText)
        message_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        message_label.setMinimumWidth(360)
        message_label.setMaximumWidth(620)

        text_column.addWidget(severity_label)
        text_column.addWidget(message_label)
        content_row.addWidget(marker, 0, Qt.AlignTop)
        content_row.addLayout(text_column, 1)
        panel_layout.addLayout(content_row)

        button_row = QHBoxLayout()
        button_row.setSpacing(8)
        button_row.addStretch(1)

        if confirm:
            cancel_button = QPushButton("Anulează")
            cancel_button.setObjectName("secondaryButton")
            cancel_button.setMinimumWidth(104)
            cancel_button.clicked.connect(self.reject)
            cancel_button.setDefault(True)

            continue_button = QPushButton("Continuă")
            continue_button.setObjectName("warningButton")
            continue_button.setMinimumWidth(104)
            continue_button.clicked.connect(self.accept)

            button_row.addWidget(cancel_button)
            button_row.addWidget(continue_button)
        else:
            close_button = QPushButton("Închide")
            close_button.setObjectName("dangerButton" if kind == "error" else "warningButton")
            close_button.setMinimumWidth(104)
            close_button.clicked.connect(self.accept)
            close_button.setDefault(True)
            button_row.addWidget(close_button)

        panel_layout.addLayout(button_row)


def show_warning_dialog(parent: QWidget | None, message: str, title: str = APP_TITLE) -> None:
    AlertDialog(parent, message, "warning", title=title).exec()


def show_error_dialog(parent: QWidget | None, message: str, title: str = APP_TITLE) -> None:
    AlertDialog(parent, message, "error", title=title).exec()


def confirm_warning_dialog(parent: QWidget | None, message: str, title: str = APP_TITLE) -> bool:
    result = AlertDialog(parent, message, "warning", confirm=True, title=title).exec()
    return result == QDialog.DialogCode.Accepted


class MainWindow(QMainWindow):
    browse_source_requested = Signal()
    browse_target_requested = Signal()
    browse_attendance_requested = Signal()
    browse_output_requested = Signal()
    run_requested = Signal()
    open_output_requested = Signal()

    def __init__(self) -> None:
        super().__init__()
        self._is_busy = False
        self._last_output_exists = False
        self.setWindowTitle(APP_TITLE)
        icon_path = resource_path("letter-s2.png")
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        self.resize(1040, 720)
        self.setMinimumSize(900, 640)
        self.setStyleSheet(APP_QSS)
        self._build_ui()
        self.set_status("Alege fișierele.", "info")
        self.set_progress(0.0, "În așteptare")
        self._sync_template_sheet_state()
        self._sync_ready_state()

    def _build_ui(self) -> None:
        root = QWidget(objectName="root")
        self.setCentralWidget(root)

        outer = QVBoxLayout(root)
        outer.setContentsMargins(14, 12, 14, 12)
        outer.setSpacing(10)

        header = QHBoxLayout()
        header.setSpacing(10)
        title = QLabel(APP_TITLE)
        title.setObjectName("appTitle")
        self.status_badge_label = QLabel("INFO")
        self.status_badge_label.setObjectName("statusBadge")
        header.addWidget(title)
        header.addStretch(1)
        header.addWidget(self.status_badge_label)
        outer.addLayout(header)

        form_panel = QFrame(objectName="panel")
        form_layout = QGridLayout(form_panel)
        form_layout.setContentsMargins(16, 13, 16, 13)
        form_layout.setHorizontalSpacing(12)
        form_layout.setVerticalSpacing(8)
        form_layout.setColumnMinimumWidth(0, 150)
        form_layout.setColumnStretch(1, 1)

        self.source_path_edit = self._make_path_edit("Fișier sursă cu situația pe zile")
        self.target_path_edit = self._make_path_edit("Workbook salarii existent")
        self.attendance_path_edit = self._make_path_edit("Fișier Transatori+Detinuți pentru CO și mențiuni (opțional)")
        self.output_dir_edit = self._make_path_edit("Folder unde se salvează workbook-ul generat")
        self.target_month_combo = QComboBox()
        self.target_month_combo.addItems(MONTH_OPTIONS)
        self._configure_combo(self.target_month_combo)
        self.template_mode_combo = QComboBox()
        self.template_mode_combo.addItems(
            [TEMPLATE_MODE_PREDEFINED_LABEL, TEMPLATE_MODE_PREVIOUS_SHEET_LABEL]
        )
        self._configure_combo(self.template_mode_combo)
        self.template_sheet_combo = QComboBox()
        self._configure_combo(self.template_sheet_combo)
        self.sheet_combo = self.template_sheet_combo

        self.source_button = self._make_browse_button()
        self.target_button = self._make_browse_button()
        self.attendance_button = self._make_browse_button()
        self.output_button = self._make_browse_button()
        self.source_button.clicked.connect(self.browse_source_requested.emit)
        self.target_button.clicked.connect(self.browse_target_requested.emit)
        self.attendance_button.clicked.connect(self.browse_attendance_requested.emit)
        self.output_button.clicked.connect(self.browse_output_requested.emit)

        files_title = self._make_section_title("Fișiere", "file")
        settings_title = self._make_section_title("Setări", "settings")

        form_layout.addWidget(files_title, 0, 0, 1, 3)
        self._add_row(form_layout, 1, "Situație", self.source_path_edit, self.source_button)
        self._add_row(form_layout, 2, "Salarii", self.target_path_edit, self.target_button)
        self._add_row(form_layout, 3, "Transatori+Detinuți (opțional)", self.attendance_path_edit, self.attendance_button)
        self._add_row(form_layout, 4, "Output", self.output_dir_edit, self.output_button)
        form_layout.addWidget(settings_title, 5, 0, 1, 3)
        self._add_row(form_layout, 6, "Lună", self.target_month_combo, None)
        self._add_row(form_layout, 7, "Template", self.template_mode_combo, None)
        self.template_sheet_label = QLabel("Sheet template")
        self.template_sheet_label.setObjectName("fieldLabel")
        form_layout.addWidget(self.template_sheet_label, 8, 0)
        form_layout.addWidget(self.template_sheet_combo, 8, 1, 1, 2)
        outer.addWidget(form_panel)

        self.source_path_edit.textChanged.connect(self._sync_ready_state)
        self.target_path_edit.textChanged.connect(self._sync_ready_state)
        self.attendance_path_edit.textChanged.connect(self._sync_ready_state)
        self.output_dir_edit.textChanged.connect(self._sync_ready_state)
        self.target_month_combo.currentTextChanged.connect(self._sync_ready_state)
        self.template_sheet_combo.currentTextChanged.connect(self._sync_ready_state)
        self.template_mode_combo.currentTextChanged.connect(self._sync_template_sheet_state)
        self.template_mode_combo.currentTextChanged.connect(self._sync_ready_state)

        actions_panel = QFrame(objectName="panel")
        actions_layout = QVBoxLayout(actions_panel)
        actions_layout.setContentsMargins(16, 13, 16, 13)
        actions_layout.setSpacing(9)

        status_row = QHBoxLayout()
        self.status_label = QLabel()
        self.status_label.setWordWrap(True)
        self.progress_text_label = QLabel()
        self.progress_text_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        self.progress_text_label.setStyleSheet(f"color: {COLOR_MUTED}; font-weight: 700;")
        status_row.addWidget(self.status_label, 1)
        status_row.addWidget(self.progress_text_label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(False)

        buttons = QHBoxLayout()
        buttons.setSpacing(10)
        self.run_button = self._make_action_button("Generează", "play", primary=True)
        self.run_button.setObjectName("primaryButton")
        self.open_button = self._make_action_button("Deschide", "external-link", primary=False)
        self.open_button.setObjectName("secondaryButton")
        self.open_button.setEnabled(False)
        self.run_button.clicked.connect(self.run_requested.emit)
        self.open_button.clicked.connect(self.open_output_requested.emit)
        buttons.addWidget(self.run_button)
        buttons.addWidget(self.open_button)
        buttons.addStretch(1)

        actions_layout.addLayout(status_row)
        actions_layout.addWidget(self.progress_bar)
        actions_layout.addLayout(buttons)
        outer.addWidget(actions_panel)

        log_panel = QFrame(objectName="panel")
        log_layout = QVBoxLayout(log_panel)
        log_layout.setContentsMargins(16, 13, 16, 13)
        log_layout.setSpacing(8)
        log_title = self._make_section_title("Jurnal", "log")
        self.log_output = QPlainTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setMaximumBlockCount(600)
        log_layout.addWidget(log_title)
        log_layout.addWidget(self.log_output, 1)
        outer.addWidget(log_panel, 1)

    def _make_path_edit(self, placeholder: str) -> QLineEdit:
        edit = QLineEdit()
        edit.setPlaceholderText(placeholder)
        edit.setClearButtonEnabled(True)
        return edit

    def _icon(self, name: str) -> QIcon:
        icon_path = resource_path("assets", "icons", f"{name}.svg")
        return QIcon(str(icon_path)) if icon_path.exists() else QIcon()

    def _make_browse_button(self) -> QPushButton:
        button = QPushButton("Alege")
        button.setIcon(self._icon("folder"))
        button.setIconSize(QSize(16, 16))
        button.setMinimumWidth(92)
        button.setToolTip("Alege fișierul sau folderul")
        return button

    def _make_action_button(self, text: str, icon_name: str, primary: bool = False) -> QPushButton:
        button = QPushButton(text)
        button.setIcon(self._icon(icon_name))
        button.setIconSize(QSize(17, 17))
        button.setMinimumWidth(118 if primary else 108)
        return button

    def _make_section_title(self, text: str, icon_name: str) -> QWidget:
        wrapper = QWidget()
        layout = QHBoxLayout(wrapper)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        icon_label = QLabel()
        icon_label.setPixmap(self._icon(icon_name).pixmap(QSize(14, 14)))
        title_label = QLabel(text)
        title_label.setObjectName("sectionTitle")

        layout.addWidget(icon_label)
        layout.addWidget(title_label)
        layout.addStretch(1)
        return wrapper

    def _configure_combo(self, combo: QComboBox) -> None:
        combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        combo.setMaxVisibleItems(8)
        view = QListView()
        view.setUniformItemSizes(True)
        view.setMinimumHeight(36)
        view.setMaximumHeight(220)
        combo.setView(view)

    def _add_row(
        self,
        layout: QGridLayout,
        row: int,
        label: str,
        control: QWidget,
        button: QPushButton | None,
    ) -> None:
        title = QLabel(label)
        title.setObjectName("fieldLabel")
        layout.addWidget(title, row, 0)
        layout.addWidget(control, row, 1)
        if button is None:
            layout.addWidget(QWidget(), row, 2)
        else:
            layout.addWidget(button, row, 2)

    def form_values(self) -> dict[str, str]:
        return {
            "source_path": self.source_path_edit.text().strip(),
            "target_path": self.target_path_edit.text().strip(),
            "attendance_path": self.attendance_path_edit.text().strip(),
            "target_month_name": self.target_month_combo.currentText().strip(),
            "template_mode_label": self.template_mode_combo.currentText().strip(),
            "template_sheet_name": self.template_sheet_combo.currentText().strip(),
            "output_dir": self.output_dir_edit.text().strip(),
        }

    def set_source_path(self, path: str) -> None:
        self.source_path_edit.setText(path)
        self._sync_ready_state()

    def set_target_path(self, path: str) -> None:
        self.target_path_edit.setText(path)
        self._sync_ready_state()

    def set_attendance_path(self, path: str) -> None:
        self.attendance_path_edit.setText(path)
        self._sync_ready_state()

    def set_output_dir(self, path: str) -> None:
        self.output_dir_edit.setText(path)
        self._sync_ready_state()

    def set_template_sheet_options(self, sheet_names: list[str], selected: str = "") -> None:
        self.template_sheet_combo.clear()
        self.template_sheet_combo.addItems(sheet_names)
        if selected and selected in sheet_names:
            self.template_sheet_combo.setCurrentText(selected)
        elif sheet_names:
            self.template_sheet_combo.setCurrentIndex(0)
        self._sync_ready_state()

    def set_sheet_options(self, sheet_names: list[str], selected: str = "") -> None:
        self.set_template_sheet_options(sheet_names, selected)

    def set_target_month(self, month_name: str) -> None:
        if month_name and month_name in MONTH_OPTIONS:
            self.target_month_combo.setCurrentText(month_name)
        self._sync_ready_state()

    def set_source_month_hint(self, month_name: str | None) -> None:
        self.target_month_combo.setToolTip(f"Lună detectată: {month_name}" if month_name else "")

    def _sync_template_sheet_state(self) -> None:
        uses_template_sheet = (
            self.template_mode_combo.currentText() == TEMPLATE_MODE_PREVIOUS_SHEET_LABEL
        )
        self.template_sheet_label.setVisible(uses_template_sheet)
        self.template_sheet_combo.setVisible(uses_template_sheet)

    def _sync_ready_state(self) -> None:
        values = self.form_values()
        needs_template_sheet = (
            values["template_mode_label"] == TEMPLATE_MODE_PREVIOUS_SHEET_LABEL
        )
        required_ok = all(
            values[key]
            for key in ("source_path", "target_path", "target_month_name", "output_dir")
        )
        if needs_template_sheet:
            required_ok = required_ok and bool(values["template_sheet_name"])
        self.run_button.setEnabled((not self._is_busy) and required_ok)
        self.open_button.setEnabled((not self._is_busy) and self._last_output_exists)

        if self._is_busy:
            return
        if required_ok:
            self.set_status("Gata de rulare.", "success")
        else:
            self.set_status("Alege fișierele.", "info")

    def set_status(self, message: str, kind: str = "info") -> None:
        colors = {
            "info": COLOR_TEXT,
            "success": COLOR_SUCCESS,
            "warning": COLOR_WARNING,
            "error": COLOR_ERROR,
        }
        badges = {
            "info": "INFO",
            "success": "GATA",
            "warning": "ATENȚIE",
            "error": "EROARE",
        }
        badge_bg = {
            "info": COLOR_PRIMARY,
            "success": COLOR_SUCCESS,
            "warning": COLOR_WARNING,
            "error": COLOR_ERROR,
        }
        self.status_label.setText(message)
        self.status_label.setStyleSheet(f"font-weight: 700; color: {colors.get(kind, COLOR_TEXT)};")
        self.status_badge_label.setText(badges.get(kind, "INFO"))
        self.status_badge_label.setStyleSheet(
            f"background: {badge_bg.get(kind, COLOR_PRIMARY)}; color: white; "
            "font-weight: 800; padding: 4px 9px; border-radius: 8px;"
        )

    def set_progress(self, percent: float, text: str = "") -> None:
        value = max(0, min(100, int(round(percent))))
        self.progress_bar.setValue(value)
        self.progress_text_label.setText(f"{text} · {value}%" if text else f"{value}%")

    def append_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_output.appendPlainText(f"[{timestamp}] {message}")

    def set_busy(self, is_busy: bool, last_output_exists: bool = False) -> None:
        self._is_busy = is_busy
        self._last_output_exists = last_output_exists
        for control in (
            self.source_path_edit,
            self.target_path_edit,
            self.attendance_path_edit,
            self.output_dir_edit,
            self.target_month_combo,
            self.template_mode_combo,
            self.template_sheet_combo,
        ):
            control.setEnabled(not is_busy)
        self._sync_ready_state()

    def set_open_output_enabled(self, enabled: bool) -> None:
        self._last_output_exists = enabled
        self.open_button.setEnabled((not self._is_busy) and enabled)
