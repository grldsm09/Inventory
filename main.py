import sys
from pathlib import Path

from PySide6.QtWidgets import QApplication, QMessageBox

from ui.main_window import MainWindow


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName("Inventory Generator")
    app.setOrganizationName("Inventory Generator")

    stylesheet = Path(__file__).resolve().parent / "ui" / "dark.qss"
    if not stylesheet.is_file():
        QMessageBox.critical(None, "Startup Error", f"Theme file not found:\n{stylesheet}")
        return 1
    app.setStyleSheet(stylesheet.read_text(encoding="utf-8"))

    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
