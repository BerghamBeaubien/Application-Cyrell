# Guide de l'Application Cyrell

## Introduction
Bienvenue dans le guide de l'application Cyrell ! Ce document vous aidera à comprendre et utiliser facilement l'application, même si vous êtes débutant. L'application est conçue pour automatiser certaines opérations sur Solid Edge et simplifier la gestion des fichiers DXF, STEP et DFT.

Des **illustrations** seront ajoutées à certains endroits pour clarifier les instructions.

---

## Structure de l'Application
L'application est divisée en **deux onglets principaux** :

1. **Automatisation des Opérations sur Solid Edge** : Génération de fichiers DFT, conversion en DXF/STEP, extraction de dimensions.
2. **Validation des Fichiers et Quantités** : Vérification de la cohérence des fichiers DXF/STEP avec un fichier Excel.

![Capture d'écran de l'interface principale](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/CapturePrincipale.png)

---

## **Onglet 1 : Automatisation des Opérations sur Solid Edge**
Cet onglet automatise la gestion des fichiers Solid Edge. Avant de commencer, **sélectionnez un répertoire** contenant les fichiers à traiter.

### **1. Ouvrir les fichiers sélectionnés**
![Bouton Ouvrir Fichiers Choisis](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/btnOpenSelFiles.png)
Le bouton **"Ouvrir Fichiers Choisis"** permet d'ouvrir rapidement les fichiers choisis dans Solid Edge.

- Sélectionnez des fichiers.
- Cliquez sur **"Ouvrir Fichiers Choisis"**.
- Les fichiers s'ouvrent dans Solid Edge.

### **2. Traitement des Fichiers DXF**
![Bouton Annoter DXF (Tag)](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/btnTagDxf.png)
Le bouton **"Annoter DXF (Tag)"** ajoute une annotation avec le nom du fichier dans les fichiers DXF choisis. 

- Choisissez les fichiers à traiter.
- Cliquez sur **"Annoter DXF (Tag)"**.
- Les fichiers traités sont annotés et sauvegardés

### **3. Exporter les dimensions**
![Bouton Exporter dimensions](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/btnExportDim.png)
Le bouton **"Exporter Dimensions"** extrait les dimensions des fichiers sélectionnés et les enregistre dans un fichier Excel.

- Sélectionnez des fichiers **DXF, PAR, PSM**.
- Cliquez sur **"Exporter Dimensions"**.
- Un fichier **Excel (Dimensions-Deplie.xlsx)** est créé avec les dimensions des fichiers.

### **4. Créer Flat Pattern**
![Bouton Créer Flat Pattern](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/btnFlatPattern.png)
Le bouton **"Créer Flat Pattern"** crée un Flat Pattern des fichiers sélectionnés et les enregistre.

- Sélectionnez des fichiers **PAR/PSM**.
- Cliquez sur **"Créer Flat Pattern"**.
- Choisissez le mode de sélection de face.
  - Mode Manuel: À chaque pièce vous devez cliquer sur la face et l'arête 
  - Mode Automatique: La face la plus loin du centre de la pièce sera utilisée pour faire le déplié
![Formulaire Sélection du mode](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/btnFlatForm.png)
- Les fichiers sont dépliés.


### **5. Sauvegarder DXF/STEP**
![Bouton Sauvegarder DXF & Step](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/btnSaveDxfStep.png)
Le bouton **"Sauvegarder DXF/STEP"** convertit les fichiers sélectionnés en formats DXF **(si la pièce est dépliée)** et STEP.

- Sélectionnez des fichiers **PAR/PSM**.
- Cliquez sur **"Sauvegarder DXF/STEP"**.
- Choisissez le dossier de destination et cochez les options spéciales (voir Explications Détaillées pour plus d’info).
- ![Formulaire Sauvegarder DXF & Step](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/btnDxfStepForm.png)
- Les fichiers sont convertis en DXF et STEP.

### **6. Créer un fichier DFT**
![Bouton Générer Dessins (DFT)](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/btnDft.png)
Le bouton **"Créer un fichier DFT"** génère un fichier de dessin (DFT) à partir des fichiers sélectionnés.

- Sélectionnez des fichiers **PAR, PSM ou ASM**.
- Cliquez sur **"Créer un fichier DFT"**.
- Choisissez un des 2 modes et cochez les paramètres voulus
![Fomulaire DFT Standard](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/standardDft.png)
![Fomulaire DFT Custom](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/customDft.png)
- Un fichier DFT est généré avec des vues automatiques et, si applicable, une nomenclature et une table de pliage.

### **7. Paramètres**  
![Pic bouton Parametres](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/btnParam.png)  
![Pic Parametres](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/params.png)

Ces paramètres permettent de contrôler certaines actions lors du traitement des fichiers.  

#### **Tag DXF**  
- Lorsque l'option est activée, après l'annotation des DXF, tous les fichiers traités seront affichés dans Solid Edge. *(Activé par défaut)*  

#### **Extraire Dimensions**  
- Si une pièce n'est pas dépliée, ses mesures retournées seront à 0. Cette option contrôle l'affichage d'un message pour chaque pièce non dépliée. *(Activé par défaut)*  
- Garde les fichiers traités (*seulement PSM et PAR*) ouverts dans Solid Edge. *(Désactivé par défaut)*  

### **8. Autres Fonctionnalités**

#### **Fermer Solid Edge**
- Ferme toutes les instances de Solid Edge en cours d'exécution.

#### **Select All**
- Sélectionne tous les fichiers affichés.

#### **Mode Sombre/Clair**
- Ce bouton caché (encerclé en rouge) permet de passer entre les 2 modes.
![Pic NightMode](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/NightMode.png)

#### **Gestion des extensions**
- **Clic droit sur une extension** : la coche seule.
- **Clic droit à nouveau** : toutes les extensions sont cochées.
![Pic ext bar](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/extensionBar.png)

---

## **Onglet 2 : Validation des Fichiers et Quantités**
Cet onglet permet de comparer les fichiers avec un document Excel pour assurer leur cohérence.
![Pic Panel Excel](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/capturePanelXl.png)

### **1. Vérifier les dimensions**
- Assurez-vous que les chemins vers le fichier Excel et les fichiers DXF sont définis.
- Cliquez sur **"Vérifier les dimensions"**.
- Un rapport s'affiche indiquant les incohérences.

![Bouton Vérifier dimensions](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/btnXlDim.png)

### **2. Vérifier le nombre de pièces**
- Assurez-vous que les chemins vers le fichier Excel, DXF et STEP sont définis.
- Cliquez sur **"Vérifier le nombre de pièces"**.
- Un rapport s'affiche indiquant les incohérences.

![Bouton Vérifier quantites](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/btnXlQte.png)

---
## **Explications Détaillées des Fonctionnalités Solid Edge**

### Bouton : Ouvrir les fichiers choisis
Ce bouton effectue les actions suivantes :

1. **Vérification des fichiers sélectionnés** :
   - Vérifie qu'au moins un fichier est sélectionné.
   - Affiche un message d'erreur si aucun fichier n'est sélectionné.

2. **Connexion à Solid Edge** :
   - Se connecte à Solid Edge en mode silencieux (sans afficher l'interface utilisateur).
   - Si Solid Edge n'est pas déjà ouvert, une nouvelle instance est démarrée.

3. **Traitement des fichiers** :
   - Pour chaque fichier sélectionné :
     - Si le fichier est un fichier STEP/STP, il est ouvert en utilisant un modèle spécifique (`Normal.asm`).
     - Si le fichier est un fichier DXF ou autre, il est ouvert directement dans Solid Edge.
   - Le nom du document est conservé pour correspondre au nom du fichier d'origine.

4. **Gestion des erreurs** :
   - En cas d'erreur pendant l'ouverture des fichiers, un message d'erreur est affiché avec les détails.

5. **Nettoyage et fin de l'opération** :
   - Une fois les fichiers ouverts, Solid Edge devient visible.
   - Un message de confirmation est affiché pour indiquer que les fichiers ont été traités avec succès.

### Bouton : Traiter les fichiers DXF
Ce bouton effectue les actions suivantes :

1. **Vérification des fichiers sélectionnés** :
   - Vérifie qu'au moins un fichier DXF est sélectionné.
   - Affiche un message d'erreur si aucun fichier n'est sélectionné.

2. **Connexion à Solid Edge** :
   - Se connecte à Solid Edge en mode silencieux (sans afficher l'interface utilisateur).
   - Désactive les alertes de Solid Edge pour éviter les interruptions.

3. **Traitement des fichiers** :
   - Pour chaque fichier DXF sélectionné :
     - Ouvre le fichier dans Solid Edge.
     - Ajoute une annotation (callout) en fonction des dimensions du fichier.
     - Sauvegarde le fichier avec les modifications.
     - Supprime le fichier d'origine si nécessaire.

4. **Gestion des erreurs** :
   - Affiche un message d'erreur en cas de problème pendant le traitement.

5. **Nettoyage et fin de l'opération** :
   - Ferme Solid Edge (sauf si l'utilisateur a choisi de le garder ouvert).
   - Affiche un message de confirmation une fois le traitement terminé.

### Bouton : Exporter dimensions
Ce bouton effectue les actions suivantes :

1. **Vérification des fichiers sélectionnés** :
   - Vérifie qu'au moins un fichier est sélectionné.
   - Affiche un message d'erreur si aucun fichier n'est sélectionné.

2. **Initialisation d'Excel** :
   - Lance une instance d'Excel pour créer un fichier de sortie.
   - Ajoute une feuille de calcul avec des en-têtes prédéfinis : **File Name**, **Width (inches)**, **Height (inches)**, **Thickness (inches)**.

3. **Traitement des fichiers** :
   - Pour chaque fichier sélectionné :
     - **Fichiers DXF** : Extrait les dimensions (largeur et hauteur) à l'aide de `DxfDimensionExtractor`.
     - **Fichiers PAR/PSM** : Ouvre le fichier dans Solid Edge, extrait les dimensions du modèle déplié (Flat Pattern) et les convertit en pouces.
     - Les dimensions sont écrites dans le fichier Excel.

4. **Gestion des erreurs** :
   - En cas d'erreur pendant le traitement, un message d'erreur est affiché avec les détails.
   - Si un fichier PAR/PSM n'a pas de modèle déplié, un message d'avertissement peut être affiché (selon les paramètres).

5. **Sauvegarde du fichier Excel** :
   - Le fichier Excel est sauvegardé sous le nom `Dimensions-Deplie.xlsx` dans le répertoire.
   - Si un fichier portant le même nom existe déjà, un numéro est ajouté pour éviter les conflits (par exemple, `Dimensions-Deplie (1).xlsx`).

6. **Nettoyage et fin de l'opération** :
   - Ferme Solid Edge (sauf si l'utilisateur a choisi de le garder ouvert).
   - Affiche un message de confirmation une fois l'exportation terminée.

### Bouton : Sauvegarder DXF/STEP
Ce bouton effectue les actions suivantes :

1. **Vérification des fichiers sélectionnés** :
   - Vérifie qu'au moins un fichier PAR ou PSM est sélectionné.
   - Affiche un message d'erreur si aucun fichier n'est sélectionné.

2. **Sélection du répertoire** :
   - Ouvre une fenêtre pour choisir le répertoire de sauvegarde commun pour les fichiers DXF et STEP.
   - Permet également d'activer les options :
     - **Tag DXF** : pour ajouter des annotations aux fichiers DXF.
     - **Changer le nom** : pour modifier le nom des fichiers générés.
     - **Fabbrica** : pour remplacer les "B" par des "P" dans le nom et supprimer "_Default_As Machined".

3. **Traitement des fichiers** :
   - Pour chaque fichier sélectionné :
     - Ouvre le fichier dans Solid Edge.
     - Vérifie si le fichier contient un modèle déplié (Flat Pattern).
     - Si un modèle déplié est trouvé, le fichier est sauvegardé en DXF et STEP dans le même répertoire.
     - Si l'option **Tag DXF** est activée, une annotation est ajoutée au fichier DXF.

4. **Gestion des erreurs** :
   - En cas d'erreur pendant le traitement, un message d'erreur est affiché avec les détails.
   - Si un fichier ne contient pas de modèle déplié, un message d'avertissement est affiché et seul le fichier STEP est généré.

5. **Nettoyage et fin de l'opération** :
   - Ferme Solid Edge (sauf si l'utilisateur choisit de le garder ouvert).
   - Affiche un message de confirmation une fois l'opération terminée.

### Bouton : Créer un fichier DFT
Ce bouton effectue les actions suivantes :

1. **Vérification des fichiers sélectionnés** :
   - Vérifie qu'au moins un fichier PAR, PSM ou ASM est sélectionné.
   - Affiche un message d'erreur si aucun fichier n'est sélectionné.

2a. **Création du fichier DFT** Cas #1 (Mode Standard):
   - Ouvre Solid Edge et crée un nouveau document DFT basé sur un modèle prédéfini (`Normal.dft`).
   - Pour chaque fichier sélectionné:
     - Ajoute une nouvelle feuille dans le document DFT.
     - Crée des vues (vue de face, vue de côté, vue de dessous) à partir du fichier PAR, PSM ou ASM.
     - Si le fichier est un assemblage (ASM), une nomenclature est générée automatiquement et une fenêtre s'ouvre affichant toutes les pièces contenues dans l'assemblage.
        - L'utilisateur peut cocher/décocher les pièces qu'il souhaite inclure dans le fichier DFT:
        - Un clic droit sur une pièce permet de **modifier son nom**. 
        ![Pic liste pieces asm](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/reName.png)
        ![Pic rename](https://github.com/BerghamBeaubien/Application-Cyrell/blob/main/Resources/actualRename.png)
        - **Attention** : une pièce ne peut pas avoir le même nom que l'assemblage.
     - Si le fichier est une pièce en tôle (PAR ou PSM), une table de pliage est ajoutée.

2b. **Création du fichier DFT** Cas #2 (Mode Avancé):
   - Ouvre Solid Edge et crée un nouveau document DFT basé sur un modèle prédéfini (`Normal.dft`).
   - Pour chaque fichier sélectionné (Dans la même page du document DFT):
     - Crée la vue du déplié, la table de plie et la table de nomenclature (dépendamment des paramètres cochés).
     - Crée un espace entre les différentes pièces pour que l'information soit lisible.
   - Les fichiers ASM ne sont pas traités

3. **Gestion des assemblages** :
   - Si l'option **Dessins individuels pour les assemblages** est activée, une fenêtre permet de sélectionner les composants à inclure dans des feuilles séparées.
   - Les composants sélectionnés sont traités individuellement, avec leurs propres vues et tables de pliage.

4. **Gestion des erreurs** :
   - En cas d'erreur pendant le traitement, un message d'erreur est affiché avec les détails.
   - Si un fichier ne contient pas de modèle déplié (pour les tables de pliage), un message d'avertissement est affiché.

5. **Nettoyage et fin de l'opération** :
   - Le fichier DFT est sauvegardé avec le nom `Dessins dft`.
   - Solid Edge reste ouvert pour permettre à l'utilisateur de visualiser ou de modifier le fichier DFT.

## **Explications Détaillées des Fonctionnalités Excel QC**

### Bouton : Vérifier les dimensions
Ce bouton effectue les actions suivantes :

1. **Vérification des chemins** :
   - Vérifie que le fichier Excel et le répertoire des fichiers DXF sont spécifiés.
   - Affiche un message d'erreur si l'un des chemins est manquant ou invalide.

2. **Lecture des données Excel** :
   - Lit les données de la feuille "PROJET" dans le fichier Excel.
   - Extrait les informations suivantes pour chaque ligne :
     - **TAG** : Identifiant unique de la pièce.
     - **Quantité** : Quantité de la pièce.
     - **Largeur** : Largeur de la pièce.
     - **Hauteur** : Hauteur de la pièce.

3. **Validation des fichiers DXF** :
   - Compare les fichiers DXF présents dans le répertoire spécifié avec les tags extraits du fichier Excel.
   - Vérifie les dimensions (largeur et hauteur) des fichiers DXF par rapport aux valeurs Excel.

4. **Génération du rapport** :
   - **Fichiers manquants** : Liste les fichiers DXF manquants par rapport aux tags Excel.
   - **Fichiers supplémentaires** : Liste les fichiers DXF qui n'ont pas de correspondance dans le fichier Excel.
   - **Incohérences de dimensions** : Signale les différences entre les dimensions Excel et DXF.
   - **Statut global** : Indique si la vérification est réussie ou non.

5. **Affichage des détails** :
   - Un tableau détaillé montre pour chaque tag :
     - La correspondance avec un fichier DXF.
     - La correspondance des dimensions (largeur et hauteur).
     - Le statut (OK, erreur ou avertissement).

### Bouton : Vérifier le nombre de pièces
Ce bouton effectue les actions suivantes :

1. **Vérification des chemins** :
   - Vérifie que le fichier Excel, le répertoire des fichiers DXF et le répertoire des fichiers STEP sont spécifiés.
   - Affiche un message d'erreur si l'un des chemins est manquant ou invalide.

2. **Lecture des données Excel** :
   - Lit les données de la feuille "PROJET" dans le fichier Excel.
   - Extrait les informations suivantes pour chaque ligne :
     - **TAG** : Identifiant unique de la pièce.
     - **Quantité** : Quantité de la pièce.

3. **Validation des fichiers DXF et STEP** :
   - Compare les fichiers DXF et STEP présents dans les répertoires spécifiés avec les tags extraits du fichier Excel.
   - Vérifie la présence des fichiers DXF et STEP pour chaque tag.

4. **Génération du rapport** :
   - **Fichiers manquants** : Liste les fichiers DXF et STEP manquants par rapport aux tags Excel.
   - **Fichiers supplémentaires** : Liste les fichiers DXF et STEP qui n'ont pas de correspondance dans le fichier Excel.
   - **Statut global** : Indique si la vérification est réussie ou non.

5. **Affichage des détails** :
   - Un tableau détaillé montre pour chaque tag :
     - La correspondance avec un fichier DXF.
     - La correspondance avec un fichier STEP.
     - Le statut (OK, erreur ou avertissement).


---

## **Conclusion**
Ce guide vous aide à utiliser l’application Cyrell efficacement. Si vous avez des questions, n’hésitez pas à contacter Mouad pour tout problème ou suggestions.

