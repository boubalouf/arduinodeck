import sys
import ctypes
import json
import threading
import subprocess
import keyboard
import socket
import os
import emoji
from PyQt5.QtNetwork import QLocalServer, QLocalSocket
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QListWidget, QListWidgetItem, QLabel,
    QPushButton, QHBoxLayout, QVBoxLayout, QGridLayout, QInputDialog, QMenu, QFileDialog,
    QSystemTrayIcon, QStyle, QFrame, QLineEdit, QSpacerItem, QSizePolicy, QToolButton,
    QDialog, QScrollArea, QDesktopWidget
)
from PyQt5.QtCore import Qt, QMimeData, QSize, QTimer, QPropertyAnimation, QEasingCurve
from PyQt5.QtGui import QDrag, QIcon, QPixmap, QPainter, QFont, QColor

# Forcer l'affichage de l'icône personnalisée dans la barre des tâches Windows
myappid = 'com.mon.streamdeck.arduino.1.0'
try:
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass

CONFIG_FILE = "config_streamdeck.json"
APP_NAME = "StreamDeckMaison"
ICON_PATH = os.path.join(os.path.dirname(__file__), 'icon.ico')
INSTANCE_ID = "ArduinoDeck_Unique_ID"

ACTION_OPEN_APP = "open_app"
ACTION_SHORTCUT = "shortcut"
ACTION_PLAY_PAUSE = "play_pause"
ACTION_VOL_UP = "vol_up"
ACTION_VOL_DOWN = "vol_down"
ACTION_MUTE = "mute"
ACTION_COPY = "copy"
ACTION_PASTE = "paste"

actions_proposees = {
    "Raccourcis Système": [
        {"type": ACTION_PLAY_PAUSE, "name": "Lecture / Pause", "icon": "⏯️"},
        {"type": ACTION_VOL_UP, "name": "Volume +", "icon": "🔊"},
        {"type": ACTION_VOL_DOWN, "name": "Volume -", "icon": "🔉"},
        {"type": ACTION_MUTE, "name": "Muer (Mute)", "icon": "🔇"},
        {"type": ACTION_COPY, "name": "Copier", "icon": "📄"},
        {"type": ACTION_PASTE, "name": "Coller", "icon": "📋"},
        {"type": ACTION_SHORTCUT, "name": "Raccourci Clavier", "icon": "⌨️"},
        {"type": ACTION_OPEN_APP, "name": "Ouvrir Application", "icon": "🚀"},
    ]
}


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

class DraggableListWidget(QListWidget):
    """QListWidget personnalisé pour corriger le bug de la barre noire lors du Drag."""
    def sizeHint(self):
        """Calcule la hauteur réelle nécessaire pour afficher tous les items sans scroll interne."""
        h = 0
        for i in range(self.count()):
            # Si sizeHintForRow n'est pas défini, on utilise une valeur de secours généreuse (48px)
            h += self.sizeHintForRow(i) if self.sizeHintForRow(i) > 0 else 48
        return QSize(super().sizeHint().width(), h + 10)

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
                background-color: #111111; 
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
        self.main_layout.setContentsMargins(0, 0, 0, 20)
        self.main_layout.setSpacing(0)
        self.main_layout.setAlignment(Qt.AlignTop) # Bloque les éléments en haut

        self.populate()
        self.main_layout.addStretch(1) # Pousse les catégories vers le haut
        self.setWidget(self.container)

    def populate(self):
        for category, actions in actions_proposees.items():
            section = CollapsibleSection(category, actions)
            self.main_layout.addWidget(section)

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

    def refresh_style(self):
        border_color = "#00a2ff" if self.is_selected else "#2a2a2a"
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: #1c1c1c;
                color: white;
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

    def update_button(self):
        self.setText("") # On ne veut plus de texte sur le bouton
        if self.action:
            if self.action.get("icon"):
                self.setIcon(self.create_text_icon(self.action["icon"]))
            else:
                self.setIcon(self.create_text_icon("+", color="#444444"))
            self.setIconSize(QSize(80, 80))
        else:
            self.setIcon(QIcon())

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

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ArduinoDeck")
        self.setWindowIcon(QIcon(ICON_PATH))
        self.resize(1000, 700)
        self.setStyleSheet("QMainWindow { background-color: #111111; }")

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        self.main_layout = QVBoxLayout(main_widget)
        self.main_layout.setContentsMargins(20, 20, 20, 20)
        self.main_layout.setSpacing(20)

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

        self.load_all()

        # Setup System Tray
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon(ICON_PATH))
        self.tray_icon.setToolTip("ArduinoDeck")
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
        
        form_layout.addLayout(col1)
        form_layout.addLayout(col2)
        form_layout.addLayout(col3)
        
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

        if btn.action:
            self.config_panel.show()
            # Bloquer les signaux pour éviter une boucle lors du remplissage
            self.name_input.blockSignals(True)
            self.path_input.blockSignals(True)
            
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
            else:
                # Cacher les champs inutiles pour les actions système fixes
                self.path_label.hide()
                self.path_input.hide()

            self.name_input.blockSignals(False)
            self.path_input.blockSignals(False)
        else:
            self.config_panel.hide()

    def apply_config_changes(self):
        if not self.selected_button or not self.selected_button.action:
            return
        
        self.selected_button.action["name"] = self.name_input.text()
        typ = self.selected_button.action.get("type")
        if typ == ACTION_OPEN_APP:
            self.selected_button.action["path"] = self.path_input.text()
        elif typ == ACTION_SHORTCUT:
            self.selected_button.action["shortcut"] = self.path_input.text()
            
        self.selected_button.update_button()
        self.save_all()

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
            elif typ == ACTION_VOL_UP:
                keyboard.press_and_release("volume up")
            elif typ == ACTION_VOL_DOWN:
                keyboard.press_and_release("volume down")
            elif typ == ACTION_MUTE:
                keyboard.press_and_release("volume mute")
            elif typ == ACTION_COPY:
                keyboard.press_and_release("ctrl+c")
            elif typ == ACTION_PASTE:
                keyboard.press_and_release("ctrl+v")
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

if __name__ == "__main__":
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

    # Initialisation de la fenêtre principale
    window = MainWindow()

    # Connexion du signal pour réveiller cette instance via le serveur local
    server.newConnection.connect(lambda: handle_new_connection(server, window))

    # Gestion du mode silencieux (réduit dans la barre des tâches) ou affichage normal
    if "silent" in sys.argv:
        window.hide()
    else:
        window.show()

    sys.exit(app.exec_())
