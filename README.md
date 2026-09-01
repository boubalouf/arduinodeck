# 🚀 ArduinoDeck

**ArduinoDeck** est un contrôleur de macros et de raccourcis puissant développé en Python avec **PyQt5**. Conçu à l'origine pour fonctionner avec un Arduino émulant un clavier HID sur les touches de fonction étendues (**F13 à F24**), il constitue une alternative logicielle complète, fluide et personnalisable au Stream Deck d'Elgato.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python)
![PyQt5](https://img.shields.io/badge/Framework-PyQt5-green?logo=qt)
![Windows](https://img.shields.io/badge/Platform-Windows-0078D4?logo=windows)

---

## ✨ Fonctionnalités

### 🎙️ Gestion Audio & Overlay OSD
* **Mute / Unmute du Microphone :** Contrôle instantané du micro principal sous Windows via `Pycaw`.
* **Overlay visuel à l'écran :** Notification contextuelle dynamique en haut à gauche pour vérifier l'état du micro sans altérer le focus de vos jeux ou logiciels.

### 🎛️ Écoute Clavier & Grille
* **Support natif F13 à F24 :** Écoute globale en arrière-plan via `Pynput` des touches envoyées par votre Arduino (HID).
* **12 touches configurables (3x4) :** Réorganisation rapide des boutons par **Drag & Drop**.
* **Personnalisation complète :** Modification du titre, de l'icône (Emoji) et des actions associées.

### 🔍 Bibliothèque d'Emojis Intégrée
* Recherche intelligente en **Français**.
* Système de **Scroll Infini** (Lazy Loading) pour une fluidité maximale.
* Rendu haute définition des emojis pour une interface moderne.

### ⚙️ Automatisation & Système
* **Actions système :** Contrôle du volume général, gestion multimédia (Play/Pause, Suivant/Précédent), raccourcis presse-papiers.
* **Lancement d'applications :** Exécution directe de fichiers `.exe`, raccourcis ou dossiers.
* **Auto-Start & Mode Tray :** Option de lancement au démarrage de Windows et réduction discrète dans la zone de notification.

---

## 🛠️ Installation & Démarrage

### 1. Prérequis
* Windows 10/11
* Python 3.10 ou version ultérieure

### 2. Cloner le projet
```bash
git clone [https://github.com/votre-utilisateur/arduinodeck.git](https://github.com/votre-utilisateur/arduinodeck.git)
cd arduinodeck
