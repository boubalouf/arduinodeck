import sys
import os

# Fix du répertoire de travail : garantit que l'app trouve ses fichiers même
# lorsqu'elle est lancée depuis le registre Windows (démarrage automatique)
if getattr(sys, 'frozen', False):
    os.chdir(os.path.dirname(sys.executable))
else:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
import ctypes as _ctypes
import json
import threading
import subprocess
import keyboard
from pynput import keyboard as pynput_keyboard
import socket
import webbrowser
import requests
from abc import ABC, abstractmethod
try:
    import winreg
except ImportError:
    winreg = None
import emoji
from ctypes import cast, POINTER
try:
    from comtypes import CoInitialize, CoUninitialize
except ImportError:
    CoInitialize = None
    CoUninitialize = None

class ctypes(ABC):
    """Concrete wrapper around the stdlib ctypes API used by this app.

    The original project uses ``ctypes.windll`` to call Win32 APIs, so we keep the
    familiar attribute access while exposing useful helper methods for library loading,
    function invocation and JSON conversion.
    """

    @property
    def windll(self):
        return _ctypes.windll

    @property
    def wintypes(self):
        return _ctypes.wintypes

    @abstractmethod
    def load_library(self, name):
        raise NotImplementedError("load_library() must be implemented")

    @abstractmethod
    def call_function(self, library_name, function_name, *args, restype=None, argtypes=None):
        raise NotImplementedError("call_function() must be implemented")

    @abstractmethod
    def to_json(self, value):
        raise NotImplementedError("to_json() must be implemented")

    def __getattr__(self, name):
        return getattr(_ctypes, name)

    def load_library(self, name):
        """Load a DLL and return the callable object exposing its exported functions."""
        return _ctypes.WinDLL(name)

    def call_function(self, library_name, function_name, *args, restype=None, argtypes=None):
        """Invoke a function from a DLL with optional type metadata."""
        dll = _ctypes.WinDLL(library_name)
        func = getattr(dll, function_name)
        if restype is not None:
            func.restype = restype
        if argtypes is not None:
            func.argtypes = argtypes
        return func(*args)

    def to_json(self, value):
        """Serialize Python data to JSON with UTF-8-friendly output."""
        return json.dumps(value, ensure_ascii=False, indent=2)

ctypes = ctypes()

try:
    if CoInitialize:
        CoInitialize()
except Exception:
    pass

# Hook global pour intercepter les exceptions non capturées et éviter les fermetures silencieuses
def handle_exception(exc_type, exc_value, exc_traceback):
    if issubclass(exc_type, KeyboardInterrupt):
        sys.__excepthook__(exc_type, exc_value, exc_traceback)
        return
    try:
        print("[CRITICAL ERROR]", exc_type, exc_value)
    except Exception:
        # En cas de problème lors du log, retomber sur le hook par défaut
        sys.__excepthook__(exc_type, exc_value, exc_traceback)

sys.excepthook = handle_exception

try:
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
except ImportError:
    AudioUtilities = None
    IAudioEndpointVolume = None
    ISimpleAudioVolume = None
from PyQt5.QtNetwork import QLocalServer, QLocalSocket
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QListWidget, QListWidgetItem, QLabel,
    QPushButton, QHBoxLayout, QVBoxLayout, QGridLayout, QInputDialog, QMenu, QFileDialog, 
    QComboBox, QSpinBox, QMessageBox,
    QSystemTrayIcon, QStyle, QFrame, QLineEdit, QSpacerItem, QSizePolicy, QToolButton, 
    QDialog, QScrollArea, QDesktopWidget, QTabWidget, QCheckBox, QStackedWidget,
    QGraphicsOpacityEffect
)
from PyQt5.QtCore import Qt, QMimeData, QSize, QTimer, QPropertyAnimation, QEasingCurve, pyqtSignal, QObject
from PyQt5.QtGui import QDrag, QIcon, QPixmap, QPainter, QFont, QColor

# Forcer l'affichage de l'icône personnalisée dans la barre des tâches Windows
myappid = 'ArduinoDeck'
try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass

def resource_path(relative_path):
    """ Retourne le chemin absolu vers la ressource, compatible dev et PyInstaller """
    try:
        # PyInstaller crée un dossier temporaire et stocke le chemin dans _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)

def get_app_dir():
    if getattr(sys, 'frozen', False):
        # Si l'application est compilée avec PyInstaller (.exe)
        return os.path.dirname(sys.executable)
    else:
        # Si l'application tourne depuis le script .py
        return os.path.dirname(os.path.abspath(__file__))

def get_microphone_volume_control():
    """Récupère l'interface de contrôle du volume du microphone principal de Windows."""
    try:
        if CoInitialize:
            try:
                CoInitialize()
            except Exception:
                pass
        if not AudioUtilities:
            return None
        mic = AudioUtilities.GetMicrophone()
        if not mic:
            return None
        if hasattr(mic, '_ctl'):
            return mic._ctl.QueryInterface(ISimpleAudioVolume)
        elif IAudioEndpointVolume:
            interface = mic.Activate(IAudioEndpointVolume._iid_, 7, None)
            return cast(interface, POINTER(IAudioEndpointVolume))
    except Exception as e:
        print(f"[ERROR Audio] Accès microphone: {e}")
    return None

def get_microphone_mute_state():
    """Retourne True si le micro est coupé, False s'il est actif, None si indisponible."""
    try:
        vol = get_microphone_volume_control()
        if vol is not None:
            return bool(vol.GetMute())
    except Exception as e:
        print(f"[ERROR Audio] Lecture état micro: {e}")
    return None


# Dossier de données utilisateur (lecture/écriture sûre, même depuis Program Files)
APPDATA_DIR = os.path.join(os.getenv('APPDATA', os.path.expanduser('~')), 'ArduinoDeck')
os.makedirs(APPDATA_DIR, exist_ok=True)

CONFIG_FILE_PATH = os.path.join(APPDATA_DIR, "config_streamdeck.json")
CONFIG_FILE = CONFIG_FILE_PATH
APP_NAME = "ArduinoDeck"
ICON_PATH = resource_path('icon.ico')
INSTANCE_ID = "ArduinoDeck_Unique_ID"
LOG_FILE = os.path.join(APPDATA_DIR, "error.log")

def log_error(msg):
    """Écrit un message d'erreur dans le fichier log situé dans APPDATA."""
    import traceback, datetime
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.datetime.now().isoformat()}] {msg}\n")
            if sys.exc_info()[0]:
                traceback.print_exc(file=f)
    except Exception:
        pass
    print(msg)

ACTION_OPEN_APP = "open_app"
ACTION_SHORTCUT = "shortcut"
ACTION_PLAY_PAUSE = "play_pause"
ACTION_VOL_UP = "vol_up"
ACTION_VOL_DOWN = "vol_down"
ACTION_MUTE = "mute"
ACTION_MUTE_MIC = "mute_mic"
ACTION_COPY = "copy"
ACTION_PASTE = "paste"
ACTION_OPEN_WEB = "open_web"
ACTION_PREV_TRACK = "prev_track"
ACTION_NEXT_TRACK = "next_track"
ACTION_HA = "home_assistant"

actions_proposees = {
    "Son": [
        {"type": ACTION_PLAY_PAUSE, "name": "Lecture / Pause", "icon": "⏯️"},
        {"type": ACTION_VOL_UP, "name": "Volume +", "icon": "🔊"},
        {"type": ACTION_VOL_DOWN, "name": "Volume -", "icon": "🔉"},
        {"type": ACTION_MUTE, "name": "Muer (Mute)", "icon": "🔇"},
        {"type": ACTION_MUTE_MIC, "name": "Muer Micro (Mute Mic)", "icon": "🎙️"},
    ],
    "Raccourcis / Application": [
        {"type": ACTION_COPY, "name": "Copier", "icon": "📄"},
        {"type": ACTION_PASTE, "name": "Coller", "icon": "📋"},
        {"type": ACTION_SHORTCUT, "name": "Raccourci Clavier", "icon": "⌨️"},
        {"type": ACTION_OPEN_APP, "name": "Ouvrir Application", "icon": "🚀"},
        {"type": ACTION_OPEN_WEB, "name": "Ouvrir Site Web", "icon": "🌐"},
        {"type": ACTION_PREV_TRACK, "name": "Chanson Précédente", "icon": "⏮️"},
        {"type": ACTION_NEXT_TRACK, "name": "Chanson Suivante", "icon": "⏭️"},
    ],
    "Home Assistant": [
        {"type": ACTION_HA, "name": "Home Assistant", "icon": "🏠"},
    ]
}


class MicMuteOverlay(QWidget):
    def __init__(self):
        super().__init__(None)
        self.setWindowFlags(
            Qt.WindowStaysOnTopHint | 
            Qt.FramelessWindowHint | 
            Qt.Tool | 
            Qt.WindowDoesNotAcceptFocus
        )
        self.setAttribute(Qt.WA_TranslucentBackground)
        
        self.layout = QHBoxLayout(self)
        self.label = QLabel("", self)
        self.label.setStyleSheet("""
            QLabel {
                background-color: rgba(10, 10, 10, 230);
                color: white;
                font-size: 16px;
                font-weight: bold;
                padding: 12px 24px;
                border-radius: 10px;
                border: 2px solid #555555;
            }
        """)
        self.layout.addWidget(self.label)
        
        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self.hide)

    def show_status(self, is_muted):
        if is_muted:
            self.label.setText("🎙️ Micro : COUPÉ")
            self.label.setStyleSheet(self.label.styleSheet() + "QLabel { color: #FF4444; border-color: #FF4444; }")
        else:
            self.label.setText("🎙️ Micro : ACTIF")
            self.label.setStyleSheet(self.label.styleSheet() + "QLabel { color: #44FF44; border-color: #44FF44; }")
            
        self.adjustSize()
        self.move(50, 50)
        self.show()
        self.raise_()
        self.timer.start(1500)


class EmojiDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Bibliothèque d'Emojis (FR)")
        self.setFixedSize(520, 680)
        self.setStyleSheet("background-color: #1c1c1c; color: white;")
        
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(15, 15, 15, 15)

        # Barre de recherche avec style moderne
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Rechercher un emoji (ex: fusée, rire, coeur)...")
        self.search_input.setStyleSheet("""
            QLineEdit {
                background: #000; border: 2px solid #333; padding: 12px; 
                border-radius: 10px; font-size: 15px; color: white;
            }
            QLineEdit:focus { border: 2px solid #00a2ff; }
        """)
        self.search_input.textChanged.connect(self.on_search_triggered)
        main_layout.addWidget(self.search_input)

        # Configuration du QListWidget pour une grille performante
        self.list_widget = QListWidget()
        self.list_widget.setViewMode(QListWidget.IconMode)
        self.list_widget.setIconSize(QSize(50, 50))
        self.list_widget.setResizeMode(QListWidget.Adjust)
        self.list_widget.setMovement(QListWidget.Static)
        self.list_widget.setSpacing(10)
        self.list_widget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.list_widget.setUniformItemSizes(True)
        self.list_widget.setStyleSheet("""
            QListWidget { border: none; background: transparent; outline: none; padding-top: 10px; }
            QListWidget::item { background: #262626; border-radius: 8px; }
            QListWidget::item:hover { background: #333333; border: 2px solid #00a2ff; }
        """)
        self.list_widget.itemClicked.connect(self.on_item_clicked)
        
        # Surveillance du scroll pour le Lazy Loading
        self.list_widget.verticalScrollBar().valueChanged.connect(self.check_scroll)
        main_layout.addWidget(self.list_widget)

        # Timer pour le debounce de la recherche (200ms)
        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.timeout.connect(self.perform_search)

        # Initialisation des données
        self.selected_emoji = None
        self.all_emojis_cache = []
        self.all_filtered_emojis = []
        self.current_index = 0
        self.batch_size = 40
        self.loading_batch = False

        # Pré-chargement des noms français
        self.prepare_emojis()
        self.perform_search() # Premier affichage

    def prepare_emojis(self):
        """Génère un cache des noms d'emojis en français."""
        for char, data in emoji.EMOJI_DATA.items():
            # Tentative demojize FR, sinon fallback alias ou EN
            name_fr = emoji.demojize(char, language='fr').strip(':').replace('_', ' ')
            if name_fr == char: # Fallback si pas de trad FR
                name_fr = data.get('en', '').strip(':').replace('_', ' ')
            
            self.all_emojis_cache.append({'e': char, 'n': name_fr.lower()})

    def on_search_triggered(self):
        self.search_timer.start(200)

    def perform_search(self):
        # Stopper tout chargement progressif en cours pour libérer les ressources
        self.loading_batch = False
        
        text = self.search_input.text().lower().strip()
        
        if not text:
            self.all_filtered_emojis = self.all_emojis_cache
        else:
            # Filtrage
            results = [i for i in self.all_emojis_cache if text in i['n']]
            # Priorité : commence par le mot recherché
            starts_with = [i for i in results if i['n'].startswith(text)]
            others = [i for i in results if not i['n'].startswith(text)]
            self.all_filtered_emojis = starts_with + others

        self.current_index = 0
        self.list_widget.clear()  # Nettoyage mémoire : les items précédents sont détruits
        self.load_more()

    def load_more(self):
        """Lance le chargement progressif du lot d'emojis."""
        if self.loading_batch or self.current_index >= len(self.all_filtered_emojis):
            return

        self.loading_batch = True
        self.items_to_load = self.batch_size
        self.list_widget.setUpdatesEnabled(False)  # Désactive le rendu pendant le lot
        self.add_item_step()

    def add_item_step(self):
        """Ajoute un emoji et planifie le suivant pour garder l'interface réactive."""
        if not self.loading_batch or self.items_to_load <= 0 or self.current_index >= len(self.all_filtered_emojis):
            self.loading_batch = False
            self.list_widget.setUpdatesEnabled(True)  # Réactive le rendu une fois le lot fini
            return

        item_data = self.all_filtered_emojis[self.current_index]
        list_item = QListWidgetItem(self.render_emoji(item_data['e']), "")
        list_item.setData(Qt.UserRole, item_data['e'])
        list_item.setSizeHint(QSize(80, 80))
        self.list_widget.addItem(list_item)

        self.current_index += 1
        self.items_to_load -= 1
        # Rappel asynchrone dans 1ms pour laisser l'UI respirer
        QTimer.singleShot(1, self.add_item_step)

    def check_scroll(self, value):
        """Déclenche le chargement quand on atteint 90% du scroll."""
        bar = self.list_widget.verticalScrollBar()
        if bar.maximum() > 0 and value > (bar.maximum() * 0.9):
            self.load_more()

    def render_emoji(self, char):
        """Génère une icône propre pour la liste."""
        pixmap = QPixmap(100, 100)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setFont(QFont("Segoe UI Emoji", 48))
        painter.drawText(pixmap.rect(), Qt.AlignCenter, char)
        painter.end()
        return QIcon(pixmap)

    def on_item_clicked(self, item):
        self.selected_emoji = item.data(Qt.UserRole)
        self.accept()

class KeyCaptureDialog(QDialog):
    """Boîte de dialogue simple pour capturer une touche unique."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Assignation")
        self.setFixedSize(300, 150)
        self.setStyleSheet("background-color: #1c1c1c; color: white;")
        layout = QVBoxLayout(self)
        self.label = QLabel("Appuyez sur la touche physique\nà assigner...")
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setStyleSheet("font-size: 14px; font-weight: bold;")
        layout.addWidget(self.label)
        self.captured_key = None

    def keyPressEvent(self, event):
        # On utilise keyboard pour obtenir le nom exact reconnu par la lib
        # ou on intercepte la touche via Qt
        key_code = event.key()
        if key_code == Qt.Key_Escape:
            self.reject()
            return
            
        # Capture de la touche via keyboard pour rester cohérent avec le listener
        self.captured_key = keyboard.read_event().name
        self.accept()

class CalibrationWizard(QDialog):
    """Fenêtre de configuration initiale pour lier touches physiques et grille."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Assistant de Calibration ArduinoDeck")
        self.setFixedSize(600, 500)
        self.setStyleSheet("background-color: #111111; color: white;")
        
        self.mapping = {} # index_grille: "nom_touche"
        
        self.main_layout = QVBoxLayout(self)
        self.stack = QStackedWidget()
        self.main_layout.addWidget(self.stack)
        
        # Page 1: Introduction
        self.intro_page = QWidget()
        self.setup_intro_page()
        self.stack.addWidget(self.intro_page)
        
        # Page 2: Grille de Calibration
        self.grid_page = QWidget()
        self.setup_grid_page()
        self.stack.addWidget(self.grid_page)

    def setup_intro_page(self):
        layout = QVBoxLayout(self.intro_page)
        layout.setContentsMargins(40, 40, 40, 40)
        layout.setSpacing(20)
        
        title = QLabel("Bienvenue sur ArduinoDeck")
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #00a2ff;")
        title.setAlignment(Qt.AlignCenter)
        
        desc = QLabel("Cet assistant va vous aider à configurer la liaison entre vos touches physiques (Arduino) et l'interface logicielle.\n\nPréparez votre boîtier et cliquez sur le bouton ci-dessous pour commencer.")
        desc.setStyleSheet("font-size: 14px; color: #b0b0b0;")
        desc.setWordWrap(True)
        desc.setAlignment(Qt.AlignCenter)
        
        start_btn = QPushButton("Commencer la configuration")
        start_btn.setCursor(Qt.PointingHandCursor)
        start_btn.setStyleSheet("""
            QPushButton { 
                background-color: #00a2ff; color: white; padding: 15px; 
                font-weight: bold; border-radius: 10px; font-size: 15px;
            }
            QPushButton:hover { background-color: #0088d6; }
        """)
        start_btn.clicked.connect(lambda: self.stack.setCurrentIndex(1))
        
        layout.addStretch()
        layout.addWidget(title)
        layout.addWidget(desc)
        layout.addSpacing(20)
        layout.addWidget(start_btn)
        layout.addStretch()

    def setup_grid_page(self):
        layout = QVBoxLayout(self.grid_page)
        title = QLabel("CALIBRATION DES TOUCHES")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: #00a2ff;")
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        layout.addWidget(QLabel("Cliquez sur un carré, puis appuyez sur la touche correspondante sur votre Arduino."))
        
        self.grid_layout = QGridLayout()
        self.btns = []
        for i in range(12):
            btn = QPushButton(f"Touche {i+1}\n(Non liée)")
            btn.setFixedSize(120, 100)
            btn.setStyleSheet("""
                QPushButton { 
                    background-color: #1c1c1c; border: 2px solid #333; 
                    border-radius: 10px; font-weight: bold; 
                }
                QPushButton:hover { border: 2px solid #00a2ff; }
            """)
            btn.clicked.connect(lambda checked, idx=i: self.capture_key(idx))
            self.grid_layout.addWidget(btn, i // 4, i % 4)
            self.btns.append(btn)
        
        layout.addLayout(self.grid_layout)
        
        self.finish_btn = QPushButton("Terminer la configuration")
        self.finish_btn.setStyleSheet("""
            QPushButton { 
                background-color: #00a2ff; padding: 15px; 
                font-weight: bold; border-radius: 5px; 
            }
        """)
        self.finish_btn.clicked.connect(self.save_and_close)
        layout.addWidget(self.finish_btn)

    def capture_key(self, index):
        cap = KeyCaptureDialog(self)
        if cap.exec_():
            key_name = cap.captured_key.upper()
            self.mapping[str(index)] = key_name
            self.btns[index].setText(f"Touche {index+1}\nLiaison: {key_name}")
            self.btns[index].setStyleSheet("background-color: #1c1c1c; border: 2px solid #00ff00; border-radius: 10px;")

    def save_and_close(self):
        if not self.mapping:
            QMessageBox.warning(self, "Attention", "Veuillez configurer au moins une touche.")
            return
            
        # Sauvegarde initiale du fichier
        data = {
            "buttons": [None] * 12,
            "settings": {
                "ha_url": "",
                "ha_token": "",
                "key_mapping": self.mapping
            }
        }
        with open(CONFIG_FILE_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=4)
        self.accept()

def check_config_exists():
    """Vérifie si une configuration valide existe déjà."""
    if not os.path.exists(CONFIG_FILE_PATH):
        return False
    try:
        with open(CONFIG_FILE_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
            buttons = data.get("buttons", [])
            # Si au moins un bouton n'est pas None, on considère que c'est configuré
            if any(b is not None for b in buttons):
                return True
            # Vérifier aussi si le mapping des touches existe
            if data.get("settings", {}).get("key_mapping"):
                return True
    except:
        return False
    return False

class HASettingsDialog(QDialog):
    def __init__(self, ha_url, ha_token, key_mapping, autostart, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Paramètres ArduinoDeck")
        self.setFixedSize(650, 580)
        self.mapping = key_mapping.copy()
        
        # Thème Moderne Noir Pur
        self.setStyleSheet("""
            QDialog { background-color: #000000; color: #e0e0e0; font-family: 'Segoe UI'; }
            QLabel { color: #b0b0b0; font-size: 13px; }
            QLineEdit { 
                background-color: #121212; border: 1px solid #333; 
                padding: 10px; border-radius: 6px; color: white;
            }
            QLineEdit:focus { border: 1px solid #00a2ff; }
            QPushButton { 
                background-color: #222; color: white; border-radius: 6px; 
                padding: 10px; font-weight: bold; 
            }
            QPushButton:hover { background-color: #444; }
            
            QTabWidget::pane { border: 1px solid #2a2a2a; top: -1px; background: #000000; border-radius: 8px; }
            QTabBar::tab { 
                background: #121212; color: #888; padding: 12px 25px; 
                border-top-left-radius: 8px; border-top-right-radius: 8px; margin-right: 2px;
            }
            QTabBar::tab:selected { background: #222222; color: #00a2ff; border-bottom: 2px solid #00a2ff; }
        """)

        self.main_layout = QVBoxLayout(self)
        self.tabs = QTabWidget()
        
        # --- ONGLET 1 : GÉNÉRAL (Home Assistant) ---
        self.tab_general = QWidget()
        gen_layout = QVBoxLayout(self.tab_general)
        gen_layout.setContentsMargins(30, 30, 30, 30)
        gen_layout.setSpacing(15)

        gen_layout.addWidget(QLabel("CONFIGURATION DOMOTIQUE"))
        
        gen_layout.addWidget(QLabel("URL du serveur Home Assistant"))
        self.url_input = QLineEdit(ha_url)
        self.url_input.setPlaceholderText("ex: http://192.168.1.100:8123")
        gen_layout.addWidget(self.url_input)
        
        gen_layout.addWidget(QLabel("Token d'accès longue durée (Bearer)"))
        self.token_input = QLineEdit(ha_token)
        self.token_input.setEchoMode(QLineEdit.Password)
        gen_layout.addWidget(self.token_input)
        
        test_row = QHBoxLayout()
        self.status_led = QLabel()
        self.status_led.setFixedSize(12, 12)
        self.set_status_color("#555") # Gris par défaut
        
        self.test_btn = QPushButton("Tester la connexion")
        self.test_btn.setCursor(Qt.PointingHandCursor)
        self.test_btn.clicked.connect(self.test_connection)
        
        test_row.addWidget(self.status_led)
        test_row.addWidget(QLabel("Statut du serveur"))
        test_row.addStretch()
        test_row.addWidget(self.test_btn)
        gen_layout.addLayout(test_row)
        gen_layout.addStretch()

        # --- ONGLET 2 : TOUCHES (Mapping) ---
        self.tab_keys = QWidget()
        keys_layout = QVBoxLayout(self.tab_keys)
        keys_layout.setContentsMargins(15, 15, 15, 15)
        
        # Grille 4x3 Matrix sans scroll
        self.grid_keys = QGridLayout()
        self.grid_keys.setSpacing(10)
        
        self.key_btns = []
        for i in range(12):
            key_name = self.mapping.get(str(i), "NON LIÉE")
            container = QFrame()
            container.setStyleSheet("background: #121212; border-radius: 8px; padding: 5px;")
            c_lay = QVBoxLayout(container)
            
            lbl = QLabel(f"Touche {i+1}")
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setStyleSheet("font-size: 11px; color: #666;")
            c_lay.addWidget(lbl)
            
            btn = QPushButton(key_name)
            btn.setStyleSheet("background: #1c1c1c; color: #00a2ff; border: 1px solid #333; font-size: 11px;")
            btn.clicked.connect(lambda checked, idx=i: self.reassign_key(idx))
            c_lay.addWidget(btn)
            
            # Placement Matrix 4 colonnes x 3 lignes
            self.grid_keys.addWidget(container, i // 4, i % 4)
            self.key_btns.append(btn)

        keys_layout.addLayout(self.grid_keys)
        keys_layout.addStretch()

        # --- ONGLET 3 : SYSTÈME ---
        self.tab_sys = QWidget()
        sys_layout = QVBoxLayout(self.tab_sys)
        sys_layout.setContentsMargins(30, 30, 30, 30)
        sys_layout.setSpacing(20)

        sys_layout.addWidget(QLabel("OPTIONS DE DÉMARRAGE"))
        self.autostart_cb = QCheckBox("Lancer automatiquement avec Windows")
        self.autostart_cb.setStyleSheet("color: white; font-size: 14px;")
        self.autostart_cb.setChecked(autostart)
        sys_layout.addWidget(self.autostart_cb)

        sys_layout.addWidget(QLabel("MAINTENANCE"))
        self.reset_btn = QPushButton("Réinitialiser les réglages d'usine")
        self.reset_btn.setStyleSheet(""" 
            QPushButton {
                background-color: transparent; border: 1px solid #441111; 
                color: #ff5555; font-size: 12px;
            }
            QPushButton:hover { background-color: #661111; }
        """)
        self.reset_btn.clicked.connect(self.reset_application)
        sys_layout.addWidget(self.reset_btn)
        sys_layout.addStretch()

        # Ajout des onglets
        self.tabs.addTab(self.tab_general, "Home Assistant")
        self.tabs.addTab(self.tab_keys, "Touches")
        self.tabs.addTab(self.tab_sys, "Système")
        self.main_layout.addWidget(self.tabs)

        # Boutons Action Bas de page
        bottom_row = QHBoxLayout()
        self.save_all_btn = QPushButton("ENREGISTRER")
        self.save_all_btn.setStyleSheet("""
            QPushButton { background-color: #00a2ff; padding: 12px 30px; border-radius: 20px; }
            QPushButton:hover { background-color: #0088d6; }
        """)
        self.save_all_btn.clicked.connect(self.accept)
        
        self.close_btn = QPushButton("Annuler")
        self.close_btn.setStyleSheet("background: transparent; color: #888;")
        self.close_btn.clicked.connect(self.reject)
        
        bottom_row.addStretch()
        bottom_row.addWidget(self.close_btn)
        bottom_row.addWidget(self.save_all_btn)
        self.main_layout.addLayout(bottom_row)

    def set_status_color(self, color):
        self.status_led.setStyleSheet(f"background-color: {color}; border-radius: 6px;")

    def reassign_key(self, index):
        cap = KeyCaptureDialog(self)
        if cap.exec_():
            key_name = cap.captured_key.upper()
            self.mapping[str(index)] = key_name
            self.key_btns[index].setText(key_name)
            self.key_btns[index].setStyleSheet("background: #2a2a2a; color: #00ff00; border: 1px solid #00ff00;")

    def test_connection(self):
        url = self.url_input.text().strip().rstrip('/')
        token = self.token_input.text().strip()
        try:
            headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
            response = requests.get(f"{url}/api/", headers=headers, timeout=5)
            if response.status_code == 200:
                self.set_status_color("#00ff00") # Vert
            else:
                self.set_status_color("#ff4444") # Rouge
        except Exception as e:
            self.set_status_color("#ff4444")

    def reset_application(self):
        reply = QMessageBox.question(self, 'Confirmation de réinitialisation',
                                     "Êtes-vous sûr de vouloir réinitialiser l'application ?\n\n"
                                     "Toutes les actions configurées et les réglages seront définitivement supprimés.",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            try:
                if os.path.exists(CONFIG_FILE_PATH):
                    os.remove(CONFIG_FILE_PATH)
                # Redémarrage propre de l'application
                os.execl(sys.executable, sys.executable, *sys.argv)
            except Exception as e:
                QMessageBox.critical(self, "Erreur", f"Échec de la réinitialisation : {str(e)}")

    def get_values(self):
        return self.url_input.text(), self.token_input.text(), self.mapping, self.autostart_cb.isChecked()

class DraggableListWidget(QListWidget):
    """QListWidget personnalisé pour corriger le bug de la barre noire lors du Drag."""
    def sizeHint(self):
        """Calcule la hauteur réelle nécessaire pour afficher tous les items sans scroll interne."""
        if self.count() == 0:
            return QSize(super().sizeHint().width(), 0)
        # Utilisation d'une hauteur fixe par ligne pour la sidebar (45px est idéal pour icône + texte)
        row_height = 45
        return QSize(super().sizeHint().width(), (self.count() * row_height) + 5)

    def startDrag(self, supportedActions):
        item = self.currentItem()
        if not item:
            return

        # Récupération des données JSON de l'action
        data = item.data(Qt.UserRole)
        mime_data = QMimeData()
        mime_data.setData("application/x-action", data.encode())

        drag = QDrag(self)
        drag.setMimeData(mime_data)

        # Création d'un pixmap simple pour le curseur (évite la barre noire)
        pixmap = QPixmap(40, 40)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setFont(QFont("Segoe UI Emoji", 20))
        painter.drawText(pixmap.rect(), Qt.AlignCenter, item.text().split()[0]) # L'emoji
        painter.end()
        
        drag.setPixmap(pixmap)
        drag.setHotSpot(pixmap.rect().center())
        drag.exec_(Qt.MoveAction)

class CollapsibleSection(QWidget):
    """Widget d'accordéon avec animation de hauteur."""
    def __init__(self, title, actions, parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(0)

        # Bouton d'en-tête (Catégorie)
        self.toggle_button = QPushButton(f"▼  {title}")
        self.toggle_button.setCheckable(True)
        self.toggle_button.setChecked(True)
        self.toggle_button.setStyleSheet("""
            QPushButton {
                background-color: #222222;
                color: #00a2ff;
                border: none;
                border-bottom: 1px solid #333333;
                text-align: left;
                padding: 12px;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton:hover { background-color: #2a2a2a; }
        """)
        self.toggle_button.clicked.connect(self.on_toggle)

        # Liste des actions (Contenu)
        self.content_list = DraggableListWidget()
        self.content_list.setDragEnabled(True)
        # Politique de taille : ne doit pas s'étendre verticalement pour ne pas écraser les autres
        self.content_list.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.content_list.setMinimumHeight(0) 
        self.content_list.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.content_list.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.content_list.setStyleSheet("""
            QListWidget {
                background-color: transparent;
                border: none;
                color: white;
                font-size: 13px;
                outline: none;
                padding: 5px;
            }
            QListWidget::item { padding: 10px; border-radius: 5px; margin: 2px; color: white; background: transparent; }
            QListWidget::item:hover { background-color: #2a2a2a; color: white; }
        """)
        
        for action in actions:
            item = QListWidgetItem(f"{action['icon']}  {action['name']}")
            item.setData(Qt.UserRole, json.dumps(action))
            self.content_list.addItem(item)

        # Ajustement de la hauteur initiale basée sur le contenu réel
        self.content_list.setMaximumHeight(self.content_list.sizeHint().height())

        self.layout.addWidget(self.toggle_button)
        self.layout.addWidget(self.content_list)

        # Animation
        self.animation = QPropertyAnimation(self.content_list, b"maximumHeight")
        self.animation.setDuration(300)
        self.animation.setEasingCurve(QEasingCurve.InOutQuart)

    def on_toggle(self):
        checked = self.toggle_button.isChecked()
        self.toggle_button.setText(f"{'▼' if checked else '▶'}  {self.toggle_button.text()[3:]}")
        
        # Recalcul dynamique de la hauteur cible pour l'animation
        target_height = self.content_list.sizeHint().height()
        start = self.content_list.maximumHeight()
        end = target_height if checked else 0
        
        self.animation.stop()
        self.animation.setStartValue(start)
        self.animation.setEndValue(end)
        # On force la mise à jour de la géométrie du parent pour le scrollbar
        self.animation.finished.connect(self.updateGeometry)
        self.animation.start()

class ActionListWidget(QScrollArea):
    """Barre latérale contenant les sections accordéon."""
    def __init__(self):
        super().__init__()
        self.setFixedWidth(220)
        self.setWidgetResizable(True)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        
        # Style sombre global pour la ScrollArea et ses barres
        self.setStyleSheet("""
            QScrollArea { 
                border: none; 
                background-color: transparent; 
                border-left: 1px solid #2a2a2a; 
            }
            QScrollBar:vertical {
                border: none;
                background: #111111;
                width: 4px;
                margin: 0px;
            }
            QScrollBar::handle:vertical { background: #333333; border-radius: 2px; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0px; }
        """)
        
        self.container = QWidget()
        self.container.setStyleSheet("background-color: #111111;")
        self.main_layout = QVBoxLayout(self.container)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)
        self.main_layout.setAlignment(Qt.AlignTop) # Bloque les éléments en haut

        self.populate()
        self.setWidget(self.container)

    def populate(self):
        # Nettoyage propre de la mémoire et des widgets existants
        while self.main_layout.count():
            item = self.main_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        for category, actions in actions_proposees.items():
            section = CollapsibleSection(category, actions)
            self.main_layout.addWidget(section)

        # Restauration de l'affichage et du scroll
        self.container.adjustSize()
        self.container.update()
        self.repaint()

class StreamDeckButton(QPushButton):
    def __init__(self, index):
        super().__init__()
        self.index = index
        self.action = None
        self.is_selected = False
        self.setAcceptDrops(True)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.open_menu)

        self.setFixedSize(120, 120)
        self.refresh_style()
        self.update_button()

    def refresh_style(self, is_muted=None):
        border_color = "#00a2ff" if self.is_selected else "#2a2a2a"
        text_color = "white"
        if self.action and self.action.get("type") == ACTION_MUTE_MIC:
            if is_muted is True:
                border_color = "#ff4444" if not self.is_selected else "#ff7777"
                text_color = "#ff5555"
            elif is_muted is False:
                border_color = "#00cc66" if not self.is_selected else "#33ff88"
                text_color = "#00ff88"

        self.setStyleSheet(f"""
            QPushButton {{
                background-color: #1c1c1c;
                color: {text_color};
                border-radius: 20px;
                font-size: 11px;
                font-weight: bold;
                border: 2px solid {border_color};
                text-align: bottom center;
                padding-bottom: 12px;
            }}
            QPushButton:hover {{
                border: 2px solid #00a2ff;
            }}
            QPushButton:pressed {{
                background-color: #005F99;
                border: 2px solid #007ACC;
            }}
        """)

    def update_button(self, is_muted=None):
        if self.action and self.action.get("type") == ACTION_MUTE_MIC:
            if is_muted is None:
                is_muted = get_microphone_mute_state()
            self.refresh_style(is_muted=is_muted)
            custom_icon = self.action.get("icon")
            if is_muted is True:
                self.setText("Micro OFF")
                icon_char = "🔇" if not custom_icon or custom_icon == "🎙️" else custom_icon
                self.setIcon(self.create_text_icon(icon_char, color="#ff5555"))
            else:
                self.setText("Micro ON")
                icon_char = "🎙️" if not custom_icon else custom_icon
                self.setIcon(self.create_text_icon(icon_char, color="#00ff88"))
            self.setIconSize(QSize(60, 60))
            return

        self.setText("") # On ne veut plus de texte sur le bouton
        if self.action:
            if self.action.get("icon"):
                self.setIcon(self.create_text_icon(self.action["icon"]))
            else:
                self.setIcon(self.create_text_icon("+", color="#444444"))
            self.setIconSize(QSize(80, 80))
        else:
            self.setIcon(QIcon())
        self.refresh_style()

    def create_text_icon(self, text, color="white"):
        pixmap = QPixmap(100, 100)
        pixmap.fill(Qt.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.Antialiasing)
        painter.setPen(QColor(color))
        painter.setFont(QFont("Segoe UI Emoji", 50))
        painter.drawText(pixmap.rect(), Qt.AlignCenter, text)
        painter.end()
        return QIcon(pixmap)

    def mousePressEvent(self, event):
        self.window().select_button(self)
        super().mousePressEvent(event)

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
                # Initialisation des valeurs par défaut si nécessaire
                if action["type"] == ACTION_OPEN_APP:
                    action["path"] = action.get("path", "")
                elif action["type"] == ACTION_SHORTCUT:
                    action["shortcut"] = action.get("shortcut", "")
                elif action["type"] == ACTION_HA:
                    action["entity_id"] = action.get("entity_id", "")
                    action["ha_value"] = action.get("ha_value", 0)
                
                self.action = action
                self.update_button()
                self.window().select_button(self)
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
            self.window().delete_selected_action()
            self.window().save_all()

    def modify_action(self):
        if not self.action:
            return
        # Mise à jour graphique
        self.update_button()
        self.window().save_all()

class KeyListenerThread(QObject):
    key_pressed_signal = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.listener = None

    def start_listening(self):
        def on_press(key):
            try:
                # Try attribute 'name' (pynput Key) or virtual key 'vk'
                if hasattr(key, 'name') and key.name:
                    key_name = key.name.lower()
                    if key_name in [f"f{i}" for i in range(13, 25)]:
                        self.key_pressed_signal.emit(key_name)
                elif hasattr(key, 'vk') and isinstance(key.vk, int) and 124 <= key.vk <= 135:
                    f_num = key.vk - 123
                    self.key_pressed_signal.emit(f"f{f_num}")
            except Exception as e:
                print(f"[ERROR KeyHook] {e}")

        try:
            self.listener = pynput_keyboard.Listener(on_press=on_press)
            self.listener.start()
        except Exception as e:
            print(f"[ERROR KeyHook] Failed to start listener: {e}")


class MainWindow(QMainWindow):
    mic_toggle_signal = pyqtSignal(bool)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("ArduinoDeck - Configuration")
        self.setWindowIcon(QIcon(ICON_PATH))
        self.resize(1000, 700)
        self.setStyleSheet("QMainWindow { background-color: #111111; }")

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        self.main_layout = QVBoxLayout(main_widget)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(20)

        # --- BARRE SUPÉRIEURE (HEADER) ---
        header_layout = QHBoxLayout()
        header_layout.setContentsMargins(5, 0, 5, 0)
        
        app_title = QLabel("ArduinoDeck")
        app_title.setStyleSheet("color: white; font-size: 26px; font-weight: bold; font-family: 'Segoe UI';")
        header_layout.addWidget(app_title)
        
        header_layout.addStretch()
        
        self.ha_settings_btn = QToolButton()
        self.ha_settings_btn.setText("⚙️ Paramètres")
        self.ha_settings_btn.setCursor(Qt.PointingHandCursor)
        self.ha_settings_btn.setStyleSheet("""
            QToolButton { 
                background-color: #00a2ff; color: white; 
                border: none; padding: 10px 20px; 
                font-weight: bold; border-radius: 8px;
                font-size: 13px;
            }
            QToolButton:hover { background-color: #0088d6; }
            QToolButton:pressed { background-color: #0072b3; }
        """)
        self.ha_settings_btn.clicked.connect(self.open_ha_settings)
        header_layout.addWidget(self.ha_settings_btn)
        
        self.main_layout.addLayout(header_layout)

        # Zone supérieure (Grille + Sidebar)
        top_container = QHBoxLayout()
        
        # Centrage de la grille
        grid_container = QVBoxLayout()
        self.buttons_layout = QGridLayout()
        self.buttons_layout.setSpacing(15)
        
        grid_wrapper = QWidget()
        grid_wrapper.setLayout(self.buttons_layout)
        grid_container.addStretch()
        grid_container.addWidget(grid_wrapper, 0, Qt.AlignCenter)
        grid_container.addStretch()

        top_container.addLayout(grid_container, 1)

        self.action_list = ActionListWidget()
        top_container.addWidget(self.action_list)

        self.main_layout.addLayout(top_container, 3)

        # Zone de configuration (Bas de page)
        self.selected_button = None
        self.setup_config_panel()

        self.buttons = []
        rows, cols = 3, 4
        for i in range(rows * cols):
            btn = StreamDeckButton(i)
            self.buttons_layout.addWidget(btn, i // cols, i % cols)
            self.buttons.append(btn)

        self.ha_url = ""
        self.ha_token = ""
        self.autostart = False
        self.key_mapping = {} # Mapping TouchePhysique -> IndexGrille
        self.ha_entities_cache = {} # Dictionnaire { "Nom de l'appareil": ["entité1", "entité2"] }
        self.load_all()

        # Overlay OSD préparé en avance (aucun appel COM ici)
        self.tray_icon = None

        # Démarrer le listener global de touches (pynput) pour F13-F24 (Arduino en mode HID clavier)
        self.key_listener = KeyListenerThread()
        self.key_listener.key_pressed_signal.connect(self.handle_key_trigger)

        # Initialisation différée : on attend 1 seconde que le bureau Windows soit prêt
        QTimer.singleShot(1000, self.init_system_tray)
        QTimer.singleShot(1000, self.key_listener.start_listening)

        # Timer de synchronisation de l'état du microphone
        self.mic_sync_timer = QTimer(self)
        self.mic_sync_timer.timeout.connect(self.update_mic_buttons)
        self.mic_sync_timer.start(2000)

        # Overlay OSD pour les notifications de statut du micro
        self.mic_overlay = MicMuteOverlay()
        
        # Connexion du signal pour un appel thread-safe de l'interface
        self.mic_toggle_signal.connect(self.handle_mic_overlay_gui)

        # Démarrage différé du service audio (2 s pour laisser Windows Audio se stabiliser)
        QTimer.singleShot(2000, self.init_audio_service)

    def handle_mic_overlay_gui(self, is_muted):
        # Cette méthode est exécutée dans le THREAD PRINCIPAL (GUI)
        if hasattr(self, 'mic_overlay'):
            self.mic_overlay.show_status(is_muted)

    def init_system_tray(self):
        """Initialise l'icône System Tray avec fallback icône système.
        Appelée 1 seconde après le démarrage via QTimer pour laisser Windows prêt."""
        try:
            if not QSystemTrayIcon.isSystemTrayAvailable():
                log_error("[WARN] System Tray non disponible sur ce bureau.")
                return

            self.tray_icon = QSystemTrayIcon(self)

            # Charge l'icône ou utilise une icône système si le fichier est introuvable
            if os.path.exists(ICON_PATH):
                self.tray_icon.setIcon(QIcon(ICON_PATH))
            else:
                self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))

            self.tray_icon.setToolTip("ArduinoDeck")

            tray_menu = QMenu()
            show_action = tray_menu.addAction("Afficher / Ouvrir")
            show_action.triggered.connect(self.show_and_restore)
            tray_menu.addSeparator()
            quit_action = tray_menu.addAction("Quitter")
            quit_action.triggered.connect(self.quit_app)

            self.tray_icon.setContextMenu(tray_menu)
            self.tray_icon.activated.connect(self.tray_activated)
            self.tray_icon.show()
            print("[INFO] System Tray initialisé.")
        except Exception as e:
            log_error(f"[CRITICAL] Erreur init System Tray: {e}")

    def init_audio_service(self):
        """Lecture initiale de l'état du micro, 2 secondes après le démarrage.
        Exécutée sur le thread principal (appelée via QTimer.singleShot)."""
        try:
            if CoInitialize:
                try:
                    CoInitialize()
                except Exception:
                    pass
            self.update_mic_buttons()
        except Exception as e:
            log_error(f"[WARN] Impossible de lire l'état audio au démarrage: {e}")


    def show_and_restore(self):
        """Affiche et restaure la fenêtre principale depuis le Tray."""
        self.show()
        self.setWindowState(self.windowState() & ~Qt.WindowMinimized | Qt.WindowActive)
        self.activateWindow()

    def handle_key_trigger(self, key_name):
        """Called from KeyListenerThread via signal when a F13-F24 key is pressed.
        Recherche la touche dans le mapping et déclenche l'action associée.
        """
        try:
            lookup = key_name.upper()
            for idx_str, assigned_key in self.key_mapping.items():
                if str(assigned_key).upper() == lookup:
                    action = self.buttons[int(idx_str)].action
                    if action:
                        self.run_action(action)
                        # Si action micro, afficher l'overlay avec l'état courant
                        if action.get('type') == ACTION_MUTE_MIC:
                            try:
                                is_muted = get_microphone_mute_state()
                                self.mic_overlay.show_status(is_muted)
                            except Exception:
                                pass
                    break
        except Exception as e:
            print(f"[ERROR KeyTrigger] {e}")

    # Serial port logic removed - Arduino is HID keyboard; global key hook used instead

    def setup_config_panel(self):
        self.config_panel = QFrame()
        self.config_panel.setFixedHeight(180)
        self.config_panel.setStyleSheet("""
            QFrame {
                background-color: #1c1c1c;
                border-radius: 15px;
                border: 1px solid #2a2a2a;
            }
            QLabel { color: #888888; font-weight: bold; border: none; }
            QLineEdit {
                background-color: #000000;
                border: 1px solid #333333;
                color: white;
                padding: 8px;
                border-radius: 5px;
            }
            QToolButton {
                background-color: #333333;
                border: none;
                color: white;
                border-radius: 5px;
                padding: 5px;
            }
            QToolButton:hover { background-color: #444444; }
        """)
        
        config_layout = QVBoxLayout(self.config_panel)
        self.config_title = QLabel("CONFIGURATION DE L'ACTION")
        config_layout.addWidget(self.config_title)

        form_layout = QHBoxLayout()
        
        col1 = QVBoxLayout()
        col1.addWidget(QLabel("Titre"))
        self.name_input = QLineEdit()
        self.name_input.textChanged.connect(self.apply_config_changes)
        col1.addWidget(self.name_input)
        
        col2 = QVBoxLayout()
        self.path_label = QLabel("Action / Chemin")
        col2.addWidget(self.path_label)
        self.path_input = QLineEdit()
        self.path_input.textChanged.connect(self.apply_config_changes)
        col2.addWidget(self.path_input)

        # Widgets spécifiques Home Assistant
        self.ha_device_combo = QComboBox()
        self.ha_device_combo.setStyleSheet("""
            QComboBox {
                background-color: #000;
                color: #00a2ff;
                padding: 5px;
                border: 1px solid #333;
                border-radius: 5px;
            }
            QComboBox::drop-down {
                border: none;
                width: 25px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #00a2ff;
            }
            QAbstractItemView {
                background-color: #1c1c1c;
                color: white;
                selection-background-color: #00a2ff;
            }
        """)
        self.ha_device_combo.currentIndexChanged.connect(self.update_ha_entity_list)
        col2.addWidget(self.ha_device_combo)
        self.ha_device_combo.hide()

        self.ha_entity_combo = QComboBox()
        self.ha_entity_combo.setStyleSheet("background-color: #000; color: white; padding: 5px;")
        self.ha_entity_combo.currentIndexChanged.connect(self.update_ha_controls)
        col2.addWidget(self.ha_entity_combo)
        self.ha_entity_combo.hide()

        self.ha_refresh_btn = QPushButton("🔄 Rafraîchir les entités")
        self.ha_refresh_btn.clicked.connect(self.fetch_ha_entities)
        col2.addWidget(self.ha_refresh_btn)
        self.ha_refresh_btn.hide()

        # Contrôles de commandes HA (Service & Valeur)
        self.ha_service_combo = QComboBox()
        self.ha_service_combo.setStyleSheet("background-color: #000; color: white; padding: 5px;")
        self.ha_service_combo.currentIndexChanged.connect(self.apply_config_changes)

        self.ha_value_spin = QSpinBox()
        self.ha_value_spin.setRange(0, 100)
        self.ha_value_spin.setSuffix(" %")
        self.ha_value_spin.valueChanged.connect(self.apply_config_changes)

        col3 = QVBoxLayout()
        col3.addWidget(QLabel("Icône"))
        icon_row = QHBoxLayout()
        
        self.emoji_btn = QToolButton()
        self.emoji_btn.setText("😊")
        self.emoji_btn.setToolTip("Emoji")
        self.emoji_btn.setFixedSize(40, 35)
        self.emoji_btn.clicked.connect(self.open_emoji_picker)
        
        self.clear_icon_btn = QToolButton()
        self.clear_icon_btn.setText("✨")
        self.clear_icon_btn.setToolTip("Effacer l'icône")
        self.clear_icon_btn.setFixedSize(40, 35)
        self.clear_icon_btn.clicked.connect(self.clear_selected_icon)

        icon_row.addWidget(self.emoji_btn)
        icon_row.addWidget(self.clear_icon_btn)
        col3.addLayout(icon_row)

        self.ha_val_col = QVBoxLayout()
        self.ha_service_label = QLabel("Commande")
        self.ha_val_label = QLabel("Luminosité (%)")
        self.ha_val_col.addWidget(self.ha_service_label)
        self.ha_val_col.addWidget(self.ha_service_combo)
        self.ha_val_col.addWidget(self.ha_val_label)
        self.ha_val_col.addWidget(self.ha_value_spin)
        
        form_layout.addLayout(col1)
        form_layout.addLayout(col2)
        form_layout.addLayout(col3)
        form_layout.addLayout(self.ha_val_col)
        
        # Bouton Supprimer
        self.delete_btn = QToolButton()
        self.delete_btn.setText("🗑️")
        self.delete_btn.setStyleSheet("background-color: #441111; color: #ff5555; font-size: 18px;")
        self.delete_btn.setFixedSize(45, 45)
        self.delete_btn.clicked.connect(self.delete_selected_action)
        form_layout.addWidget(self.delete_btn, 0, Qt.AlignBottom)

        config_layout.addLayout(form_layout)
        config_layout.addStretch()

        self.main_layout.addWidget(self.config_panel)
        self.config_panel.hide()

    def select_button(self, btn):
        # Désélectionner le précédent
        if self.selected_button:
            self.selected_button.is_selected = False
            self.selected_button.refresh_style()

        self.selected_button = btn
        btn.is_selected = True
        btn.refresh_style()

        # Blocage global des signaux du panneau pour éviter les boucles
        self.config_panel.blockSignals(True)
        self.action_list.blockSignals(True)

        if btn.action:
            self.config_panel.show()
            self.name_input.blockSignals(True)
            self.path_input.blockSignals(True)
            self.ha_device_combo.blockSignals(True)
            self.ha_entity_combo.blockSignals(True)
            self.ha_service_combo.blockSignals(True)
            self.ha_value_spin.blockSignals(True)
            
            self.name_input.setText(btn.action.get("name", ""))
            typ = btn.action.get("type")
            if typ == ACTION_OPEN_APP:
                self.path_label.show()
                self.path_input.show()
                self.path_label.setText("Chemin (.exe)")
                self.path_input.setText(btn.action.get("path", ""))
            elif typ == ACTION_SHORTCUT:
                self.path_label.show()
                self.path_input.show()
                self.path_label.setText("Raccourci (ex: ctrl+c)")
                self.path_input.setText(btn.action.get("shortcut", ""))
                self.ha_device_combo.hide()
                self.ha_entity_combo.hide()
                self.ha_refresh_btn.hide()
                self.ha_service_label.hide()
                self.ha_service_combo.hide()
                self.ha_val_label.hide()
                self.ha_value_spin.hide()
            elif typ == ACTION_OPEN_WEB:
                self.path_label.show()
                self.path_input.show()
                self.path_label.setText("URL (ex: https://...)")
                self.path_input.setText(btn.action.get("url", ""))
                self.ha_device_combo.hide()
                self.ha_entity_combo.hide()
                self.ha_refresh_btn.hide()
                self.ha_service_label.hide()
                self.ha_service_combo.hide()
                self.ha_val_label.hide()
                self.ha_value_spin.hide()
            elif typ == ACTION_HA:
                self.path_label.setText("Appareil et Entité")
                self.path_input.hide()
                self.ha_device_combo.show()
                self.ha_refresh_btn.show()
                
                # Masquer l'entité tant qu'aucun domaine n'est choisi
                self.ha_entity_combo.hide()
                self.ha_service_label.hide()
                self.ha_service_combo.hide()
                self.ha_val_label.hide()
                self.ha_value_spin.hide()

                self.ha_device_combo.clear()
                device_names = sorted(self.ha_entities_cache.keys())
                if not device_names: 
                    self.ha_device_combo.addItem("Chargement...")
                else:
                    self.ha_device_combo.addItems(device_names)
                
                current_entity = btn.action.get("entity_id", "")
                current_device = ""
                for dev, ents in self.ha_entities_cache.items():
                    if current_entity in ents:
                        current_device = dev
                        break
                
                idx = self.ha_device_combo.findText(current_device)
                if idx >= 0:
                    self.ha_device_combo.setCurrentIndex(idx)
                
                self.update_ha_entity_list()
                # Sélectionner l'entité actuelle dans la liste dépendante
                ent_idx = self.ha_entity_combo.findText(current_entity)
                if ent_idx >= 0: 
                    self.ha_entity_combo.setCurrentIndex(ent_idx)
                self.update_ha_controls()

                self.ha_service_combo.setCurrentText(btn.action.get("ha_service", "toggle"))

                self.ha_value_spin.setValue(btn.action.get("ha_value", 0))
            else:
                # Cacher les champs inutiles pour les actions système fixes
                self.path_label.hide()
                self.path_input.hide()
                self.ha_device_combo.hide()
                self.ha_entity_combo.hide()
                self.ha_refresh_btn.hide()
                self.ha_service_label.hide()
                self.ha_service_combo.hide()
                self.ha_val_label.hide()
                self.ha_value_spin.hide()

            # Libération des signaux
            self.name_input.blockSignals(False)
            self.path_input.blockSignals(False)
            self.ha_device_combo.blockSignals(False)
            self.ha_entity_combo.blockSignals(False)
            self.ha_service_combo.blockSignals(False)
            self.ha_value_spin.blockSignals(False)
        else:
            self.config_panel.hide()

        self.config_panel.blockSignals(False)
        self.action_list.blockSignals(False)
        self.config_panel.update()

    def apply_config_changes(self):
        if not self.selected_button or not self.selected_button.action:
            return
        
        # Empêche les redessinements intempestifs pendant la modification
        self.config_panel.blockSignals(True)
        
        self.selected_button.action["name"] = self.name_input.text()
        typ = self.selected_button.action.get("type")
        if typ == ACTION_OPEN_APP:
            self.selected_button.action["path"] = self.path_input.text()
        elif typ == ACTION_SHORTCUT:
            self.selected_button.action["shortcut"] = self.path_input.text()
        elif typ == ACTION_OPEN_WEB:
            self.selected_button.action["url"] = self.path_input.text()
        elif typ == ACTION_HA:
            self.selected_button.action["entity_id"] = self.ha_entity_combo.currentText()
            self.selected_button.action["ha_service"] = self.ha_service_combo.currentData()
            self.selected_button.action["ha_value"] = self.ha_value_spin.value()
            
        self.selected_button.update_button()
        self.save_all()
        
        self.config_panel.blockSignals(False)
        self.config_panel.repaint()

    def update_ha_entity_list(self):
        domain = self.ha_device_combo.currentText()
        self.ha_entity_combo.blockSignals(True)
        self.ha_entity_combo.clear()
        if domain in self.ha_entities_cache:
            self.ha_entity_combo.addItems(sorted(self.ha_entities_cache[domain]))
            self.ha_entity_combo.show()
        else:
            self.ha_entity_combo.hide()
            self.ha_service_label.hide()
            self.ha_service_combo.hide()
            self.ha_val_label.hide()
            self.ha_value_spin.hide()
        self.ha_entity_combo.blockSignals(False)
        self.update_ha_controls()

    def update_ha_controls(self):
        """Affiche les options selon le type d'entité (Light, Switch, etc)."""
        entity_id = self.ha_entity_combo.currentText()
        if not entity_id:
            self.ha_service_label.hide()
            self.ha_service_combo.hide()
            self.ha_val_label.hide()
            self.ha_value_spin.hide()
            return

        self.ha_service_combo.blockSignals(True)
        self.ha_service_combo.clear()
        services = [("On/Off", "toggle"), ("Allumer", "turn_on"), ("Éteindre", "turn_off")]
        for label, val in services:
            self.ha_service_combo.addItem(label, val)
        
        self.ha_service_label.show()
        self.ha_service_combo.show()

        if entity_id.startswith("light"):
            self.ha_val_label.setText("Luminosité (0-100%)")
            self.ha_val_label.show()
            self.ha_value_spin.show()
        else:
            self.ha_val_label.hide()
            self.ha_value_spin.hide()
        
        self.ha_service_combo.blockSignals(False)
        self.apply_config_changes()

    def open_ha_settings(self):
        dialog = HASettingsDialog(self.ha_url, self.ha_token, self.key_mapping, self.autostart, self)
        if dialog.exec_():
            url, token, mapping, autostart = dialog.get_values()
            self.ha_url = url
            self.ha_token = token
            self.key_mapping = mapping # Mise à jour immédiate en mémoire
            self.autostart = autostart
            self.set_autostart(autostart)
            self.save_all()
            self.fetch_ha_entities()

    def set_autostart(self, enabled):
        """Gère l'inscription de l'application dans le registre Windows."""
        if not winreg: return
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE)
            if enabled:
                app_path = f'"{sys.executable}" "{os.path.abspath(__file__)}" silent'
                winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, app_path)
            else:
                try:
                    winreg.DeleteValue(key, APP_NAME)
                except FileNotFoundError:
                    pass
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Erreur configuration Autostart: {e}")

    def fetch_ha_entities(self):
        if not self.ha_url or not self.ha_token: return
        try:
            url = self.ha_url.strip().rstrip('/')
            headers = {"Authorization": f"Bearer {self.ha_token}", "Content-Type": "application/json"}
            
            # Utilisation de /api/states pour éviter les erreurs 404 des registres de config
            response = requests.get(f"{url}/api/states", headers=headers, timeout=5)
            if response.status_code != 200:
                print(f"Erreur API HA States: {response.status_code}")
                return

            states = response.json()
            new_cache = {}
            for item in states:
                eid = item['entity_id']
                if '.' not in eid: continue
                
                # Heuristique de regroupement par "Appareil" basée sur le nom d'entité
                # (Extrait le préfixe avant le premier underscore pour identifier le matériel)
                domain, obj_id = eid.split('.', 1)
                device_prefix = obj_id.split('_')[0].title() if '_' in obj_id else domain.title()
                
                if device_prefix not in new_cache:
                    new_cache[device_prefix] = []
                new_cache[device_prefix].append(eid)

            self.ha_entities_cache = new_cache
            if self.selected_button and self.selected_button.action and self.selected_button.action.get("type") == ACTION_HA:
                self.select_button(self.selected_button)
        except Exception as e:
            print("Erreur fetch HA:", e)

    def open_emoji_picker(self):
        if not self.selected_button or not self.selected_button.action:
            return
        dialog = EmojiDialog(self)

        # Affichage temporaire pour calculer la taille réelle
        dialog.show()
        main_geo = self.geometry()
        
        # Calcul du centre relatif à la fenêtre parente
        # On utilise frameGeometry pour inclure les bordures de fenêtre
        x = main_geo.x() + (main_geo.width() - dialog.frameGeometry().width()) // 2
        y = main_geo.y() + (main_geo.height() - dialog.height()) // 2
        dialog.move(x, y)
        
        if dialog.exec_():
            self.selected_button.action["icon"] = dialog.selected_emoji
            self.selected_button.update_button()
            self.save_all()

    def clear_selected_icon(self):
        if self.selected_button and self.selected_button.action:
            self.selected_button.action["icon"] = ""
            self.selected_button.update_button()
            self.save_all()

    def delete_selected_action(self):
        if self.selected_button:
            self.selected_button.action = None
            self.selected_button.is_selected = False
            self.selected_button.update_button()
            self.selected_button.refresh_style()
            self.config_panel.hide()
            self.save_all()

    def save_all(self):
        # Vérification de sécurité pour éviter d'écraser avec des données vides
        if not hasattr(self, 'buttons') or len(self.buttons) == 0:
            return

        # Structure de sauvegarde incluant les paramètres globaux
        data = {
            "buttons": [b.action if b.action else None for b in self.buttons],
            "settings": {
                "ha_url": self.ha_url,
                "ha_token": self.ha_token,
                "key_mapping": self.key_mapping,
                "autostart": self.autostart
            }
        }
        try:
            with open(CONFIG_FILE_PATH, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print("Erreur sauvegarde:", e)

    def load_all(self):
        try:
            if not os.path.exists(CONFIG_FILE_PATH): return
            with open(CONFIG_FILE_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)

            # Gestion du format dictionnaire (nouveau) ou liste (ancien)
            if isinstance(data, dict):
                buttons_data = data.get("buttons", [])
                settings = data.get("settings", {})
                self.ha_url = settings.get("ha_url", "")
                self.ha_token = settings.get("ha_token", "")
                self.key_mapping = settings.get("key_mapping", {})
                self.autostart = settings.get("autostart", False)
                QTimer.singleShot(1000, self.fetch_ha_entities) # Charger les entités au démarrage
            else:
                buttons_data = data

            for i, action in enumerate(buttons_data):
                if i < len(self.buttons) and action:
                    self.buttons[i].action = action
                    self.buttons[i].update_button()
        except Exception as e:
            print("Erreur chargement:", e)

    def listen_keys(self):
        try:
            if CoInitialize:
                CoInitialize()
        except Exception:
            pass
            
        def on_key(e):
            if e.event_type == 'down':
                key_name = e.name.upper()
                # Recherche de l'index associé à cette touche dans le mapping
                for idx_str, assigned_key in self.key_mapping.items():
                    if assigned_key == key_name:
                        action = self.buttons[int(idx_str)].action
                        if action: self.run_action(action)
                        break

        keyboard.hook(on_key)
        keyboard.wait()
        
        try:
            if CoUninitialize:
                CoUninitialize()
        except Exception:
            pass

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
            elif typ == ACTION_VOL_UP:
                keyboard.press_and_release("volume up")
            elif typ == ACTION_VOL_DOWN:
                keyboard.press_and_release("volume down")
            elif typ == ACTION_MUTE:
                keyboard.press_and_release("volume mute")
            elif typ == ACTION_MUTE_MIC:
                self.toggle_microphone_mute()
            elif typ == ACTION_COPY:
                keyboard.press_and_release("ctrl+c")
            elif typ == ACTION_PASTE:
                keyboard.press_and_release("ctrl+v")
            elif typ == ACTION_OPEN_WEB:
                url = action.get("url")
                if url:
                    webbrowser.open(url)
            elif typ == ACTION_PREV_TRACK:
                keyboard.press_and_release("previous track")
            elif typ == ACTION_NEXT_TRACK:
                keyboard.press_and_release("next track")
            elif typ == ACTION_HA:
                # Lecture stricte des paramètres depuis le fichier config avant chaque requête
                try:
                    with open(CONFIG_FILE_PATH, "r", encoding="utf-8") as f:
                        config_data = json.load(f)
                        ha_settings = config_data.get("settings", {})
                        ha_url = ha_settings.get("ha_url", "").strip().rstrip('/')
                        ha_token = ha_settings.get("ha_token", "").strip()
                except:
                    ha_url, ha_token = self.ha_url.strip().rstrip('/'), self.ha_token.strip()

                entity_id = action.get("entity_id")
                service = action.get("ha_service", "toggle")
                val_pct = action.get("ha_value", 0)
                domain = entity_id.split('.')[0] if entity_id else "homeassistant"
                
                if entity_id and ha_url and ha_token:
                    headers = {"Authorization": f"Bearer {ha_token}", "Content-Type": "application/json"}
                    payload = {"entity_id": entity_id}

                    if domain == "light" and service in ["turn_on", "toggle"]:
                        # Conversion 0-100% -> 0-255
                        payload["brightness"] = int(val_pct * 2.55)

                    requests.post(f"{ha_url}/api/services/{domain}/{service}", headers=headers, json=payload, timeout=5)
        except Exception as e:
            print("Erreur exécution action :", e)

    def toggle_microphone_mute(self):
        """Bascule l'état (Mute / Unmute) du microphone principal de Windows."""
        try:
            if CoInitialize:
                try:
                    CoInitialize()
                except Exception:
                    pass
            vol = get_microphone_volume_control()
            if vol is not None:
                is_muted = bool(vol.GetMute())
                new_state = not is_muted
                vol.SetMute(new_state, None)
                
                is_muted = vol.GetMute()
                print(f"[DEBUG] Micro muted: {is_muted}")
                
                # APPEL DU SIGNAL VERS LE THREAD PRINCIPAL (GUI)
                self.mic_toggle_signal.emit(bool(is_muted))
                return new_state
            else:
                print("[ERROR Audio] Impossible d'accéder au périphérique microphone.")
        except Exception as e:
            print(f"[ERROR Audio] toggle_microphone_mute: {e}")
        return None

    def update_mic_buttons(self, is_muted=None):
        """Met à jour l'apparence des boutons Mute Micro sans bloquer l'UI."""
        try:
            mic_btns = [b for b in self.buttons if b.action and b.action.get("type") == ACTION_MUTE_MIC]
            if not mic_btns:
                return
            if is_muted is None:
                is_muted = get_microphone_mute_state()
            for btn in mic_btns:
                btn.update_button(is_muted=is_muted)
        except Exception as e:
            print(f"[ERROR Audio] update_mic_buttons: {e}")

    def closeEvent(self, event):
        event.ignore()
        self.hide()

    def tray_activated(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.show_and_restore()

    def show_window(self):
        self.show()
        self.raise_()
        self.activateWindow()

    def quit_app(self):
        if self.tray_icon:
            self.tray_icon.hide()
        QApplication.quit()

def handle_new_connection(server, window):
    """Gère l'appel d'une nouvelle instance essayant de démarrer."""
    socket = server.nextPendingConnection()
    if socket and socket.waitForReadyRead(500):
        data = socket.readAll().data().decode()
        if data == "SHOW":
            window.show()
            window.raise_()
            window.activateWindow()
    socket.close()

if __name__ == '__main__':
    import sys
    try:
        from comtypes import CoInitialize, CoUninitialize
        # Initialisation de COM pour Windows avant le lancement de Qt
        try:
            CoInitialize()
        except Exception:
            pass
    except ImportError:
        CoUninitialize = None

    app = QApplication(sys.argv)

    # Vérification de l'instance unique via QLocalSocket
    test_socket = QLocalSocket()
    test_socket.connectToServer(INSTANCE_ID)
    if test_socket.waitForConnected(500):
        # Une instance tourne déjà, on lui envoie le signal de réveil et on quitte
        test_socket.write(b"SHOW")
        test_socket.waitForBytesWritten(500)
        sys.exit(0)

    # Si aucune instance n'existe, on prépare le serveur pour écouter les prochaines
    server = QLocalServer()
    server.removeServer(INSTANCE_ID) # Nettoyage si crash précédent
    server.listen(INSTANCE_ID)

    # --- LOGIQUE WIZARD / STARTUP ---
    if not check_config_exists():
        wizard = CalibrationWizard()
        if not wizard.exec_():
            sys.exit(0) # Quitter si l'utilisateur annule le wizard

    # Initialisation de la fenêtre principale (Garde une référence stricte de l'instance)
    window = MainWindow()

    # Connexion du signal pour réveiller cette instance via le serveur local
    server.newConnection.connect(lambda: handle_new_connection(server, window))

    # Gestion du mode silencieux (réduit dans la barre des tâches) ou affichage normal
    if "silent" in sys.argv:
        window.hide()
    else:
        window.show()

    exit_code = app.exec_()

    try:
        if CoUninitialize:
            CoUninitialize()
    except Exception:
        pass

    sys.exit(exit_code)
