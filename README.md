# Portalia - Simulateur de Salaire

**Portalia** est une application simple qui vous permet de **simuler un salaire** en fonction de votre statut, régime fiscal et autres paramètres.  

---

## Ce qu’il faut installer (une seule fois)

Pour faire fonctionner l’application, il faut **deux petits outils gratuits** :

### 1. ✅ **Python**
C’est un logiciel qui permet de faire fonctionner une partie de l’application.

- Télécharger ici : [https://www.python.org/downloads/](https://www.python.org/downloads/)
- Cliquez sur **Download Python**  
- Lors de l’installation, **très important** :  
  👉 **Cochez la case "Add Python to PATH"** avant de cliquer sur "Installer"

---

### 2. ✅ **Node.js**
Il permet d’ouvrir la partie visible du simulateur.

- Télécharger ici : [https://nodejs.org/fr](https://nodejs.org/fr)
- Cliquez sur le bouton **vert** "LTS"  
- Installer le fichier téléchargé (laisser les options par défaut)

> ✅ Ces deux outils ne sont à installer qu’une seule fois !

---

## 📁 Installer Portalia (le simulateur)

### 1. Créer un dossier
- Allez sur votre **Bureau**
- Faites **clic droit** > **Nouveau** > **Dossier**
- Donnez-lui un nom (ex : `Portalia`)

---

### 2. Ouvrir l’invite de commande
- Cliquez sur la **loupe** en bas à gauche (ou sur le bouton Démarrer)
- Tapez `cmd` et ouvrez l’**Invite de commande**

---

### 3. Télécharger Portalia

1. Dans la fenêtre noire qui s’ouvre, tapez (en adaptant à votre nom de dossier) :

   ```bash
   cd Desktop\Portalia
   ```

2. Tapez ensuite :

   ```bash
   git clone https://github.com/Adrien2MN/Portalia.git
   cd Portalia
   ```

> ❗ Si un message d’erreur s’affiche disant que `git` n’est pas reconnu, dites-le nous, on vous aidera à l’installer.

---

## ▶️ Lancer l'application

### 🖥️ Étape 1 : Ouvrir la partie visible (le site web)

1. Dans l’invite de commande (toujours dans le dossier `portalia`), tapez :

   ```bash
   cd portalia
   npm install
   ng serve
   ```

2. Une fois terminé, ouvrez votre navigateur (Chrome, Edge, etc.)  
   et allez à cette adresse :  
   👉 [http://localhost:4200](http://localhost:4200)

---

### ⚙️ Étape 2 : Démarrer la partie calculs (le cerveau du site)

1. Ouvrez **une nouvelle** fenêtre d’invite de commande (comme tout à l’heure)
2. Tapez les lignes suivantes :

   ```bash
   cd Desktop\Portalia\Portalia\portalia
   pip install -r requirements.txt
   python -m uvicorn main:app --reload
   ```

3. Le simulateur est maintenant **entièrement actif** ! 🎉  
   Vous pouvez continuer à l’utiliser dans votre navigateur.

---

⚠️ **Note** : Si une erreur se produit lors de l'installation des dépendances Python, essayez de commenter la dernière ligne du fichier `requirements.txt`.

## Technologies utilisées
- **Frontend** : Angular, TypeScript, HTML, CSS
- **Backend** : FastAPI, Python
- **Base de données** : (à préciser si applicable)

## Contribution
1. Forker le projet
2. Créer une branche : 
    ```bash
    git checkout -b feature-nom
    ```
3. Apporter vos modifications et commit : 
    ```bash
    git commit -m "Ajout d'une nouvelle fonctionnalité"
    ```
4. Pousser la branche : 
    ```bash
    git push origin feature-nom
    ```
5. Ouvrir une Pull Request

## Auteur
Projet développé par l'équipe Portalia 🚀
