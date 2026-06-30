from pathlib import Path
import sys

from PySide6.QtWidgets import QApplication

from src.core.logger import configurar_logger
from src.ui.main_window import MainWindow


def base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def main() -> None:
    base = base_dir()
    configurar_logger(base)
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
