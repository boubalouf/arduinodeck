# 🚀 ArduinoDeck

**ArduinoDeck** est un contrôleur de macros et de raccourcis puissant développé en Python avec **PyQt5**. Conçu à l'origine pour fonctionner avec un Arduino émulant les touches de fonction étendues (**F13 à F24**), il peut être utilisé comme une alternative logicielle complète au Stream Deck d'Elgato.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python)
![PyQt5](https://img.shields.io/badge/Framework-PyQt5-green?logo=qt)
![Windows](https://img.shields.io/badge/Platform-Windows-0078D4?logo=windows)

---

## ✨ Fonctionnalités

### 🎛️ Gestion de la Grille
* **12 touches configurables** (3x4) par simple **Drag & Drop**.
* Support natif des touches **F13 à F24** pour déclencher les actions.
* Personnalisation complète : titre, icône (Emoji) et commande.

### 🔍 Bibliothèque d'Emojis Intégrée
* Recherche intelligente en **Français**.
* Système de **Scroll Infini** (Lazy Loading) pour une fluidité maximale.
* Rendu haute définition des emojis pour un look moderne.

### ⚙️ Automatisation & Système
* **Actions prédéfinies :** Gestion du volume, contrôle multimédia (Play/Pause), Copier/Coller.
* **Lancement d'Apps :** Ouvrez n'importe quel fichier `.exe` ou dossier.
* **Auto-Start :** Option pour lancer l'application au démarrage de Windows.
* **Mode Tray :** Se réduit dans la zone de notification pour rester actif en arrière-plan.

## 🛠️ Installation

### 1. Prérequis
Assurez-vous d'avoir Python 3.10 ou plus récent installé.

### 2. Cloner le projet
```bash
git clone [https://github.com/votre-utilisateur/arduinodeck.git](https://github.com/votre-utilisateur/arduinodeck.git)
cd arduinodeck
