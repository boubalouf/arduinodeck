import sys
import json
import threading
import subprocess
import keyboard
import socket
import os
import winreg
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QListWidget, QListWidgetItem,
    QPushButton, QHBoxLayout, QGridLayout, QInputDialog, QMenu,
    QSystemTrayIcon, QStyle
)
from PyQt5.QtCore import Qt, QMimeData, QSize
from PyQt5.QtGui import QDrag, QIcon

CONFIG_FILE = "config_streamdeck.json"
APP_NAME = "StreamDeckMaison"
LOCK_PORT = 65432  # Port utilisé pour empêcher les doubles instances

ACTION_OPEN_APP = "open_app"
ACTION_SHORTCUT = "shortcut"
ACTION_PLAY_PAUSE = "play_pause"

actions_proposees = [
    {"type": ACTION_OPEN_APP, "name": "Ouvrir une application", "icon": ""},
    {"type": ACTION_SHORTCUT, "name": "Simuler un raccourci", "icon": ""},
    {"type": ACTION_PLAY_PAUSE, "name": "Play/Pause média", "icon": ""},
]

def check_single_instance(port=LOCK_PORT):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    try:
        s.bind(("127.0.0.1", port))
    except OSError:
        print("Une autre instance est déjà en cours.")
        sys.exit()

def add_to_startup():
    exe_path = sys.executable
    key = winreg.HKEY_CURRENT_USER
    regpath = r"Software\Microsoft\Windows\CurrentVersion\Run"
    with winreg.OpenKey(key, regpath, 0, winreg.KEY_SET_VALUE) as reg:
        winreg.SetValueEx(reg, APP_NAME, 0, winreg.REG_SZ, f'"{exe_path}" silent')

class ActionListWidget(QListWidget):
    def __init__(self):
        super().__init__()
        self.setDragEnabled(True)
        self.setFixedWidth(180)
        self.populate()

    def populate(self):
        self.clear()
        for action in actions_proposees:
            item = QListWidgetItem(action["name"])
            item.setData(Qt.UserRole, json.dumps(action))
            self.addItem(item)

    def startDrag(self, supportedActions):
        item = self.currentItem()
        if not item:
            return
        mimeData = QMimeData()
        mimeData.setData("application/x-action", item.data(Qt.UserRole).encode())
        drag = QDrag(self)
        drag.setMimeData(mimeData)
        drag.exec_(Qt.MoveAction)

class StreamDeckButton(QPushButton):
    def __init__(self, index):
        super().__init__()
        self.index = index
        self.action = None
        self.setAcceptDrops(True)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.open_menu)

        self.setFixedSize(120, 120)
        self.setStyleSheet("""
            QPushButton {
                background-color: #1E1E1E;
                color: white;
                border-radius: 16px;
                font-size: 14px;
                font-weight: bold;
                border: 2px solid #333333;
                padding: 8px;
                text-align: center;
            }
            QPushButton:hover {
                background-color: #292929;
                border: 2px solid #00A2FF;
            }
            QPushButton:pressed {
                background-color: #005F99;
                border: 2px solid #007ACC;
            }
        """)

        self.update_button()

    def update_button(self):
        if self.action:
            self.setText(self.action["name"])
            self.setIconSize(QSize(48, 48))
        else:
            self.setText(f"Bouton {self.index+1}")
            self.setIcon(QIcon())

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat("application/x-action"):
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasFormat("application/x-action"):
            data = event.mimeData().data("application/x-action").data().decode()
            try:
                action = json.loads(data)
                if action["type"] == ACTION_OPEN_APP:
                    path, ok = QInputDialog.getText(self, "Chemin application", "Chemin complet du fichier .exe")
                    if not ok or not path.strip():
                        event.ignore()
                        return
                    action["path"] = path.strip()
                elif action["type"] == ACTION_SHORTCUT:
                    shortcut, ok = QInputDialog.getText(self, "Raccourci clavier", "Exemple: ctrl+alt+t")
                    if not ok or not shortcut.strip():
                        event.ignore()
                        return
                    action["shortcut"] = shortcut.strip()
                self.action = action
                self.update_button()
                self.window().save_all()
                event.accept()
            except Exception as e:
                print("Erreur drop:", e)
                event.ignore()
        else:
            event.ignore()

    def open_menu(self, pos):
        if not self.action:
            return
        menu = QMenu()
        modif = menu.addAction("Modifier")
        suppr = menu.addAction("Supprimer")
        action = menu.exec_(self.mapToGlobal(pos))
        if action == modif:
            self.modify_action()
        elif action == suppr:
            self.action = None
            self.update_button()
            self.window().save_all()

    def modify_action(self):
        if not self.action:
            return
        typ = self.action["type"]
        if typ == ACTION_OPEN_APP:
            path, ok = QInputDialog.getText(self, "Modifier chemin", "Chemin complet du fichier .exe", text=self.action.get("path", ""))
            if ok and path.strip():
                self.action["path"] = path.strip()
        elif typ == ACTION_SHORTCUT:
            shortcut, ok = QInputDialog.getText(self, "Modifier raccourci", "Tapez le raccourci", text=self.action.get("shortcut", ""))
            if ok and shortcut.strip():
                self.action["shortcut"] = shortcut.strip()
        self.update_button()
        self.window().save_all()

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Stream Deck Maison")
        self.resize(850, 520)
        self.setStyleSheet("QMainWindow { background-color: #121212; }")

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QHBoxLayout(main_widget)

        self.action_list = ActionListWidget()
        layout.addWidget(self.action_list)

        self.buttons_layout = QGridLayout()
        layout.addLayout(self.buttons_layout)

        self.buttons = []
        rows, cols = 3, 4
        for i in range(rows * cols):
            btn = StreamDeckButton(i)
            self.buttons_layout.addWidget(btn, i // cols, i % cols)
            self.buttons.append(btn)

        self.load_all()

        # Setup System Tray
        self.tray_icon = QSystemTrayIcon(self)
        icon = self.style().standardIcon(QStyle.SP_ComputerIcon)
        self.tray_icon.setIcon(icon)
        self.tray_icon.setToolTip("Stream Deck Maison")
        tray_menu = QMenu()
        restore_action = tray_menu.addAction("Ouvrir")
        quit_action = tray_menu.addAction("Quitter")
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.tray_activated)
        restore_action.triggered.connect(self.show_window)
        quit_action.triggered.connect(self.quit_app)
        self.tray_icon.show()

        self.listener_thread = threading.Thread(target=self.listen_keys, daemon=True)
        self.listener_thread.start()

    def save_all(self):
        data = [b.action if b.action else None for b in self.buttons]
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print("Erreur sauvegarde:", e)

    def load_all(self):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            for i, action in enumerate(data):
                if action:
                    self.buttons[i].action = action
                    self.buttons[i].update_button()
        except Exception:
            pass

    def listen_keys(self):
        def on_key(e):
            if e.event_type == 'down' and e.name.upper() in [
                'F13', 'F14', 'F15', 'F16', 'F17', 'F18', 'F19', 'F20', 'F21', 'F22', 'F23', 'F24'
            ]:
                index = int(e.name[1:]) - 13
                if 0 <= index < 12:
                    action = self.buttons[index].action
                    if action:
                        self.run_action(action)

        keyboard.hook(on_key)
        keyboard.wait()

    def run_action(self, action):
        try:
            typ = action["type"]
            if typ == ACTION_OPEN_APP:
                path = action.get("path")
                if path:
                    subprocess.Popen(path)
            elif typ == ACTION_SHORTCUT:
                shortcut = action.get("shortcut")
                if shortcut:
                    keyboard.press_and_release(shortcut)
            elif typ == ACTION_PLAY_PAUSE:
                keyboard.press_and_release("play/pause media")
        except Exception as e:
            print("Erreur exécution action :", e)

    def closeEvent(self, event):
        event.ignore()
        self.hide()
        self.tray_icon.showMessage(
            "ArduinoDeck",
            "L'application continue de tourner en arrière-plan.",
            QSystemTrayIcon.Information,
            2000
        )

    def tray_activated(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.show_window()

    def show_window(self):
        self.show()
        self.raise_()
        self.activateWindow()

    def quit_app(self):
        self.tray_icon.hide()
        QApplication.quit()

if __name__ == "__main__":
    check_single_instance()
    add_to_startup()

    app = QApplication(sys.argv)
    window = MainWindow()

    # Si "silent" est dans les arguments, ne pas afficher la fenêtre
    if "silent" not in sys.argv:
        window.show()

    sys.exit(app.exec_())
