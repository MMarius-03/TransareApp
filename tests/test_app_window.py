from __future__ import annotations

from app_window import (
    APP_TITLE,
    TEMPLATE_MODE_PREDEFINED_LABEL,
    TEMPLATE_MODE_PREVIOUS_SHEET_LABEL,
    MainWindow,
)


def test_main_window_smoke(qapp) -> None:
    window = MainWindow()

    assert window.windowTitle() == APP_TITLE
    assert window.run_button.text() == "Generează"
    assert window.open_button.isEnabled() is False
    assert window.open_button.text() == "Deschide"
    assert not window.open_button.icon().isNull()
    assert window.run_button.isEnabled() is False
    assert not window.run_button.icon().isNull()
    assert window.source_button.text() == "Alege"
    assert window.source_button.isEnabled() is True
    assert not window.source_button.icon().isNull()
    assert window.attendance_button.text() == "Alege"
    assert window.attendance_button.isEnabled() is True
    assert window.template_mode_combo.currentText() == TEMPLATE_MODE_PREDEFINED_LABEL
    assert window.target_month_combo.currentText() == "ianuarie"
    assert window.target_month_combo.maxVisibleItems() == 8
    assert window.template_mode_combo.maxVisibleItems() == 8
    assert window.template_sheet_combo.isHidden() is True

    window.set_template_sheet_options(["martie", "februarie"], selected="martie")
    assert window.template_sheet_combo.currentText() == "martie"

    window.template_mode_combo.setCurrentText(TEMPLATE_MODE_PREVIOUS_SHEET_LABEL)
    assert window.template_sheet_combo.isHidden() is False

    window.set_target_month("aprilie")
    assert window.form_values()["target_month_name"] == "aprilie"
    window.set_source_path("source.xlsx")
    window.set_target_path("target.xlsx")
    window.set_attendance_path("transatori_detinuti.xlsx")
    window.set_output_dir("output")
    assert window.form_values()["attendance_path"] == "transatori_detinuti.xlsx"
    assert window.run_button.isEnabled() is True

    window.set_busy(True)
    assert window.run_button.isEnabled() is False

    window.set_busy(False, last_output_exists=True)
    assert window.run_button.isEnabled() is True
    assert window.open_button.isEnabled() is True
    window.close()
