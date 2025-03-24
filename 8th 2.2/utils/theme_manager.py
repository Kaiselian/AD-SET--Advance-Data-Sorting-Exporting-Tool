from PyQt5.QtWidgets import QApplication, QMainWindow, QMenu, QAction
from PyQt5.QtCore import Qt
import qdarkstyle
import darkdetect

def get_system_theme():
    return "dark" if darkdetect.isDark() else "light"

def apply_theme(app, theme):
    if theme == "dark":
        app.setStyleSheet(qdarkstyle.load_stylesheet(qt_api='pyqt5'))
    else:
        app.setStyleSheet("")  # Reset to default light theme

class ThemeManager:
    def __init__(self, main_window):
        self.main_window = main_window
        self.app = QApplication.instance()  # Get the application instance
        self.current_theme = get_system_theme()
        apply_theme(self.app, self.current_theme)

        self.create_theme_menu()

    def create_theme_menu(self):
        menu_bar = self.main_window.menuBar()
        theme_menu = QMenu("Theme", self.main_window)

        dark_action = QAction("🌙 Dark", self.main_window)
        dark_action.triggered.connect(lambda: self.set_theme("dark"))
        theme_menu.addAction(dark_action)

        light_action = QAction("☀️ Light", self.main_window)
        light_action.triggered.connect(lambda: self.set_theme("light"))
        theme_menu.addAction(light_action)

        menu_bar.addMenu(theme_menu)

    def set_theme(self, theme):
        self.current_theme = theme
        apply_theme(self.app, theme)