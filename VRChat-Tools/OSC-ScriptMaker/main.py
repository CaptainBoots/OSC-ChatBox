"""
main.py
───────
Entry point for OSC-ScriptMaker.
"""

import sys

from PySide6.QtWidgets import QApplication

VERSION = "0.2.0"


def main():
    app = QApplication(sys.argv)

    from ui import theme
    app.setStyleSheet(theme.qss())

    from ui.app import App
    window = App()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
