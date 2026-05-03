# 🚀 ArduinoDeck

**ArduinoDeck** est un contrôleur de macros et de raccourcis puissant développé en Python avec **PyQt5**. Conçu à l'origine pour fonctionner avec un Arduino émulant les touches de fonction étendues (**F13 à F24**), il peut être utilisé comme une alternative logicielle complète au Stream Deck d'Elgato.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python)
![PyQt5](https://img.shields.io/badge/Framework-PyQt5-green?logo=qt)
![Windows](https://img.shields.io/badge/Platform-Windows-0078D4?logo=windows)

---

## ✨ Fonctionnalités

### 🎛️ Assistant de Configuration (Wizard)
* **Guide d'accueil interactif :** Prise en main pas-à-pas lors du premier lancement.
* **Calibration intuitive :** Configuration rapide et assistée des assignations de touches.

### 🎛️ Gestion de la Grille
* **12 touches configurables** (4 colonnes x 3 lignes) par simple **Drag & Drop**.
* Support natif des touches **F13 à F24** pour déclencher les actions.
* Catégories dédiées : **Son**, **Raccourcis / Applications**, et **Home Assistant**.

### ⚙️ Panneau de Paramètres Moderne
* Nouvelle interface épurée avec **thème sombre** et sections distinctes (Général, Touches, Système).
* Sauvegarde en temps réel vers le fichier `config_streamdeck.json`.
* Testez et configurez vos connexions à Home Assistant en toute simplicité.

### 🔍 Bibliothèque d'Emojis Intégrée
* Recherche intelligente en **Français**.
* Système de **Scroll Infini** pour une fluidité maximale.

---

## 🛠️ Installation

### 1. Prérequis
Assurez-vous d'avoir Python 3.10 ou plus récent installé.

### 2. Cloner le projet
```bash
git clone [https://github.com/boubalouf/arduinodeck.git](https://github.com/boubalouf/arduinodeck.git)
cd arduinodeck
