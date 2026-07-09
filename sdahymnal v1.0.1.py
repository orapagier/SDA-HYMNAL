import os
import sys
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QListWidget, QLineEdit,
    QPushButton, QFileDialog, QMessageBox
)
from PySide6.QtGui import QIcon
from PySide6.QtCore import Qt, Signal
import pythoncom
import win32com.client
import shutil
import threading
import winreg


def get_base_path():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.realpath(__file__))


def get_office_version():
    try:
        office_base_key = "Software\\Microsoft\\Office"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, office_base_key) as office_key:
            index = 0
            versions = []
            while True:
                try:
                    subkey_name = winreg.EnumKey(office_key, index)
                    if subkey_name.replace(".", "").isdigit():
                        versions.append(subkey_name)
                    index += 1
                except OSError:
                    break
            if versions:
                # Sort versions (newest first) and pick highest
                latest_version = sorted(versions, key=lambda v: list(map(int, v.split("."))), reverse=True)[0]
                return latest_version
    except Exception as e:
        print(f"Error detecting Office version: {e}")
    return None


def is_folder_trusted(path):
    office_version = get_office_version()
    if not office_version:
        print("No Microsoft Office Powerpoint Detected. \n Please install and try again!")
        return False
    
    key_path = f"Software\\Microsoft\\Office\\{office_version}\\PowerPoint\\Security\\Trusted Locations"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            index = 0
            while True:
                try:
                    subkey_name = winreg.EnumKey(key, index)
                except OSError:
                    break
                index += 1
                if not subkey_name.lower().startswith("location"):
                    continue
                try:
                    with winreg.OpenKey(key, subkey_name) as subkey:
                        location_path, _ = winreg.QueryValueEx(subkey, "Path")
                        if os.path.normcase(location_path.rstrip("\\")) == os.path.normcase(path.rstrip("\\")):
                            return True  # Found matching trusted path
                except OSError:
                    continue
    except FileNotFoundError:
        pass
    except Exception as e:
        print(f"Error checking trusted location: {e}")
    return False


def add_trusted_location(path):
    office_version = get_office_version()
    if not office_version:
        print("No Microsoft Office Powerpoint Detected. \n Please install and try again!")
        return False

    key_path = f"Software\\Microsoft\\Office\\{office_version}\\PowerPoint\\Security\\Trusted Locations"
    try:
        key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
        try:
            existing_names = set()
            index = 0
            while True:
                try:
                    existing_names.add(winreg.EnumKey(key, index))
                except OSError:
                    break
                index += 1

            new_index = 0
            while f"Location{new_index}" in existing_names:
                new_index += 1

            with winreg.CreateKey(key, f"Location{new_index}") as new_location_key:
                winreg.SetValueEx(new_location_key, "Path", 0, winreg.REG_SZ, path + "\\")
                winreg.SetValueEx(new_location_key, "AllowSubfolders", 0, winreg.REG_DWORD, 1)
                winreg.SetValueEx(new_location_key, "Description", 0, winreg.REG_SZ, "Added by SDA Hymnal App")
            return True
        finally:
            winreg.CloseKey(key)

    except PermissionError:
        print("Permission denied.")
    except Exception as e:
        print(f"Error adding trusted location: {e}")
    return False

def check_and_offer_trust():
    app_folder = get_base_path()

    if not is_folder_trusted(app_folder):
        success = add_trusted_location(app_folder)
        if not success:
            return
        
def is_presenter_view_enabled(office_version):
    key_path = f"Software\\Microsoft\\Office\\{office_version}\\PowerPoint\\Options"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            value, regtype = winreg.QueryValueEx(key, "UsePresenterView")
            return value == 1
    except FileNotFoundError:
        # Key or value doesn't exist
        return False
    except Exception as e:
        print(f"Error reading Presenter View setting: {e}")
        return False

def force_presenter_view():
    office_version = get_office_version()
    if not office_version:
        return False

    key_path = f"Software\\Microsoft\\Office\\{office_version}\\PowerPoint\\Options"
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            current_value, _ = winreg.QueryValueEx(key, "UsePresenterView")
            if current_value == 1:
                return True
    except FileNotFoundError:
        pass
    except Exception:
        pass

    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path)
        winreg.SetValueEx(key, "UsePresenterView", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
        return True
    except Exception:
        pass
    return False

class HymnalApp(QWidget):
    error_occurred = Signal(str)

    def __init__(self):
        super().__init__()
        check_and_offer_trust()
        force_presenter_view()

        self.dir_path = get_base_path()
        self.hymns = []

        self.setWindowTitle("Seventh Day Adventist Hymnal")
        self.setFixedSize(640, 480)
        self.setWindowIcon(QIcon(os.path.join(self.dir_path, "_internal", "Data", "favicon.ico")))

        self.error_occurred.connect(self.show_error)

        self.init_ui()
        self.search_bar.setFocus()

    def show_error(self, message):
        QMessageBox.critical(self, "Error", message)

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        menu_layout = QHBoxLayout()

        add_btn = QPushButton("Add Hymns")
        add_btn.setCursor(Qt.PointingHandCursor)
        add_btn.clicked.connect(self.add_hymns)

        help_btn = QPushButton("Help")
        help_btn.setCursor(Qt.PointingHandCursor)
        help_btn.clicked.connect(self.show_help)

        about_btn = QPushButton("About")
        about_btn.setCursor(Qt.PointingHandCursor)
        about_btn.clicked.connect(self.show_about)

        clear_btn = QPushButton("Refresh")
        clear_btn.setCursor(Qt.PointingHandCursor)
        clear_btn.clicked.connect(self.refresh_btn)

        menu_layout.addWidget(add_btn)
        menu_layout.addWidget(help_btn)
        menu_layout.addWidget(about_btn)
        menu_layout.addWidget(clear_btn)
        menu_layout.addStretch()

        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Search Hymns...")
        self.search_bar.textChanged.connect(self.search_files)
        menu_layout.addWidget(self.search_bar)

        main_layout.addLayout(menu_layout)

        self.result_list = QListWidget()
        self.result_list.itemDoubleClicked.connect(self.open_selected)
        main_layout.addWidget(self.result_list)

        self.scan_files()
        self.search_files()

    def scan_files(self):
        allowed_extensions = (".pps", ".ppsx", ".ppt", ".pptx")
        self.hymns = []

        for root, dirs, files in os.walk(self.dir_path):
            for file in files:
                if file.lower().endswith(allowed_extensions):
                    name = os.path.splitext(file)[0]
                    self.hymns.append((name, os.path.join(root, file)))

        self.hymns.sort(key=lambda item: item[0].lower())

    def search_files(self):
        term = self.search_bar.text().lower()
        self.result_list.clear()

        matches = [name for name, _ in self.hymns if term in name.lower()]
        if matches:
            self.result_list.addItems(matches)
        else:
            self.result_list.addItem("No hymn found. Try another!")

        if term.strip() == "":
            self.result_list.scrollToTop()

    def open_selected(self):
        selected = self.result_list.currentItem()
        if not selected:
            return
        name = selected.text()
        for hymn_name, full_path in self.hymns:
            if hymn_name == name:
                operator_screen = self.screen()
                threading.Thread(target=self.launch_ppt, args=(full_path,), daemon=True).start()
                self.hide()
                self.show_control_window(operator_screen)
                return

    def launch_ppt(self, file_path):
        try:
            pythoncom.CoInitialize()
            ppt = win32com.client.Dispatch('PowerPoint.Application')
            ppt.Visible = True

            presentation = ppt.Presentations.Open(file_path, WithWindow=False)
            settings = presentation.SlideShowSettings
            settings.AdvanceMode = 1  # On click
            settings.ShowType = 1     # Presented by speaker (Presenter View-compatible)
            settings.Run()

            ppt.WindowState = 2  # Minimize PowerPoint
        except Exception as e:
            self.error_occurred.emit(f"Error launching presentation in Presenter View:\n{e}")
        finally:
            pythoncom.CoUninitialize()

        
    def refresh_btn(self):
        self.search_bar.clear()
        self.search_files()
        self.result_list.scrollToTop()
        self.toggle_focus()
        threading.Thread(target=self._quit_ppt, daemon=True).start()

    def _quit_ppt(self):
        try:
            pythoncom.CoInitialize()
            ppt = win32com.client.GetActiveObject("PowerPoint.Application")
            ppt.Quit()
        except Exception as e:
            print("PowerPoint not running or failed to quit:", e)
        finally:
            pythoncom.CoUninitialize()

    def toggle_focus(self):
        if self.search_bar.hasFocus():
            self.result_list.setFocus()
            if self.result_list.count() > 0:
                self.result_list.setCurrentRow(0)
        else:
            self.result_list.clearSelection()
            self.search_bar.setFocus()

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Shift:
            self.toggle_focus()
        elif event.key() == Qt.Key_Up:
            self.select_previous_result()
        elif event.key() == Qt.Key_Down:
            self.select_next_result()
        elif event.key() == Qt.Key_Return or event.key() == Qt.Key_Enter:
            self.open_selected()
        elif event.key() == Qt.Key_Escape:
            self.refresh_btn()

    def select_next_result(self):
        current = self.result_list.currentRow()
        if current < self.result_list.count() - 1:
            self.result_list.setCurrentRow(current + 1)

    def select_previous_result(self):
        current = self.result_list.currentRow()
        if current > 0:
            self.result_list.setCurrentRow(current - 1)

    def add_hymns(self):
        paths, _ = QFileDialog.getOpenFileNames(self, "Select Hymns", filter="PowerPoint Files (*.pps *.ppsx *.ppt *.pptx)")
        if paths:
            target = os.path.join(self.dir_path, "Data", "Added Hymns")
            os.makedirs(target, exist_ok=True)
            added = 0

            for path in paths:
                dest_path = os.path.join(target, os.path.basename(path))
                if not os.path.exists(dest_path):
                    try:
                        shutil.copy(path, dest_path)
                        added += 1
                    except Exception as e:
                        QMessageBox.critical(self, "Error", f"Could not copy {path}: {e}")
            QMessageBox.information(self, "Done", f"{added} new hymn(s) added!")
            self.scan_files()
            self.search_files()

    def show_help(self):
        QMessageBox.information(self, "Help", "Keyboard Shortcuts:\n\nDouble-click or Enter key to open a hymn. \nEsc key to close current hymn and refreshes the app. \nShift key to switch between search bar and hymns' list. \nEnter Hymn number or Hymn keyword in the search bar. \nAdd Hymns via the 'Add Hymns' button.")
        self.search_bar.setFocus()

    def show_about(self):
        QMessageBox.information(self, "About", "Seventh Day Adventist Church Hymnal \n\nDeveloper: Jelmar A. Orapa\nEmail: orapajelmar@gmail.com")
        self.search_bar.setFocus()

    def show_control_window(self, screen=None):
        self.control_window = ControlWindow(self.restore_main_window, screen)
        self.control_window.show()
        self.control_window.activateWindow()
        self.control_window.setFocus()

    def restore_main_window(self):
        self.show()
        if hasattr(self, 'control_window') and self.control_window:
            self.control_window.close()
        self.refresh_btn()

class ControlWindow(QWidget):
    def __init__(self, restore_callback, screen=None):
        super().__init__()
        self.restore_callback = restore_callback
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)

        restore_btn = QPushButton(" Seventh Day Adventist Hymnal ")
        restore_btn.setCursor(Qt.PointingHandCursor)
        restore_btn.setStyleSheet("""
            QPushButton {
                padding: 1px 1px;
                background-color: #357ABD;
                font-family: Times New Roman;
                color: white;
                border: none;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #285A8D;
            }
        """)
        restore_btn.clicked.connect(self.restore_callback)

        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(restore_btn)
        self.setLayout(layout)

        target_screen = screen or QApplication.primaryScreen()
        screen_geo = target_screen.availableGeometry()
        x = screen_geo.x() + (screen_geo.width() - restore_btn.sizeHint().width()) // 2
        y = screen_geo.y() + 2
        self.move(x, y)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Escape:
            self.restore_callback()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    hymnal_app = HymnalApp()
    hymnal_app.show()
    sys.exit(app.exec())
