from __future__ import annotations

import sys

from PySide6.QtGui import QFont
from PySide6.QtWidgets import QApplication

from app_controller import AppController
from app_window import APP_TITLE, MainWindow
from runtime_paths import default_base_dir


def main() -> int:
    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName(APP_TITLE)
    app.setFont(QFont("Segoe UI", 10))

    window = MainWindow()
    controller = AppController(window, base_dir=default_base_dir())
    window._controller = controller
    controller.initialize()
    window.show()
    return int(app.exec())


if __name__ == "__main__":
    raise SystemExit(main())
