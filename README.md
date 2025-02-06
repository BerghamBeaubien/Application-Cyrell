# Guide de l'Application Cyrell

## Introduction
Cette application est conçue pour automatiser certaines opérations sur Solid Edge. Elle est divisée en deux onglets principaux, chacun offrant des fonctionnalités spécifiques. Ce guide explique comment utiliser l'application de manière efficace.

---

## Structure de l'Application

L'application est divisée en **deux onglets principaux**, chacun offrant des fonctionnalités spécifiques pour automatiser et simplifier les tâches liées à Solid Edge et à la gestion des fichiers.

---

### **Onglet 1 : Automatisation des Opérations sur Solid Edge**
Cet onglet permet d'automatiser certaines opérations sur Solid Edge, notamment la création de dessins (DFT), la gestion des fichiers DXF/STEP, et l'extraction des dimensions.

#### Fonctionnalités Principales :
1. **Créer un fichier DFT** :
   - Génère un fichier DFT à partir des fichiers PAR, PSM ou ASM sélectionnés.
   - Ajoute des vues (vue de face, vue de côté, vue de dessous) et des nomenclatures automatiques.
   - Inclut des tables de pliage pour les pièces en tôle.

2. **Sauvegarder DXF/STEP** :
   - Convertit les fichiers PAR/PSM en fichiers DXF et STEP.
   - Permet d'ajouter des annotations aux fichiers DXF (option **Tag DXF**).

3. **Exporter les dimensions** :
   - Exporte les dimensions des fichiers DXF, PAR et PSM vers un fichier Excel.
   - Compare les dimensions des fichiers DXF avec les valeurs Excel et signale les incohérences.

4. **Ouvrir les fichiers sélectionnés** :
   - Ouvre les fichiers PAR, PSM ou ASM sélectionnés dans Solid Edge.
   - Utilise des modèles prédéfinis pour garantir une configuration cohérente.

---

### **Onglet 2 : Validation des Fichiers et Quantités**
Cet onglet permet de vérifier la cohérence entre les fichiers Excel, les fichiers DXF et les fichiers STEP, ainsi que de valider les quantités et les dimensions.

#### Fonctionnalités Principales :
1. **Vérifier les dimensions** :
   - Compare les dimensions des fichiers DXF avec les valeurs spécifiées dans le fichier Excel.
   - Signale les fichiers manquants, les fichiers supplémentaires et les incohérences de dimensions.

2. **Vérifier le nombre de pièces** :
   - Vérifie que les fichiers DXF et STEP correspondent aux tags Excel.
   - Valide les quantités spécifiées dans le fichier Excel.
   - Signale les fichiers manquants, les fichiers supplémentaires et les incohérences de quantités.

---

## Onglet 1 : Automatisation des Opérations sur Solid Edge

Cet onglet permet d'automatiser certaines opérations sur des fichiers Solid Edge. Avant de commencer, l'utilisateur doit sélectionner un répertoire contenant les fichiers à traiter.

### Bouton : Traiter les fichiers DXF
Ce bouton permet de traiter les fichiers DXF sélectionnés en placant un tag(annontation) à l'intérieur de la pièce :

1. **Sélectionnez les fichiers** : Dans la liste, sélectionnez un ou plusieurs fichiers DXF.
2. **Cliquez sur le bouton** : Utilisez le bouton **Traiter les fichiers DXF** pour lancer l'opération.
3. **Attendez la confirmation** : Une fois le traitement terminé, un message s'affiche pour confirmer la fin des opérations.

![Image du bouton](C:\Users\mouad.khalladi\source\repos\firstCSMacro\firstCSMacro\Resources\btnTagDxf.png) 

### Bouton : Ouvrir les fichiers sélectionnés
Ce bouton permet d'ouvrir les fichiers sélectionnés dans Solid Edge. Voici comment l'utiliser :

1. **Sélectionnez les fichiers** : Dans la liste, sélectionnez un ou plusieurs fichiers (DXF, STEP, etc.).
2. **Cliquez sur le bouton** : Utilisez le bouton **Ouvrir fichiers choisis** pour lancer l'opération.
3. **Attendez la confirmation** : Une fois les fichiers ouverts, un message s'affiche pour confirmer la fin des opérations.

![Image du bouton](C:\Users\mouad.khalladi\source\repos\firstCSMacro\firstCSMacro\Resources\btnOpenSelFiles.png)

### Bouton : Exporter dimensions
Ce bouton permet d'exporter les dimensions des fichiers sélectionnés (DXF, PAR, PSM) vers un fichier Excel. Voici comment l'utiliser :

1. **Sélectionnez les fichiers** : Dans la liste, sélectionnez un ou plusieurs fichiers (DXF, PAR, PSM).
2. **Cliquez sur le bouton** : Utilisez le bouton Exporter dimensions pour lancer l'exportation.
3. **Vérifiez le fichier Excel** : Un fichier Excel nommé Dimensions-Deplie.xlsx est créé dans le répertoire sélectionné, contenant les dimensions des fichiers.

![Image du bouton](C:\Users\mouad.khalladi\source\repos\firstCSMacro\firstCSMacro\Resources\btnExportDim.png)

### Bouton : Sauvegarder DXF/STEP
Ce bouton permet de sauvegarder les fichiers PAR/PSM sélectionnés en fichiers DXF et STEP. Voici comment l'utiliser :

1. **Sélectionnez les fichiers** : Dans la liste, sélectionnez un ou plusieurs fichiers PAR ou PSM.
2. **Cliquez sur le bouton** : Utilisez le bouton **Sauvegarder DXF/STEP** pour lancer l'opération.
3. **Confirmez les répertoires** : Une fenêtre s'ouvre pour choisir les répertoires de sauvegarde des fichiers DXF et STEP.
4. **Vérifiez les fichiers générés** : Les fichiers DXF et STEP sont sauvegardés dans les répertoires spécifiés.

![Image du bouton](C:\Users\mouad.khalladi\source\repos\firstCSMacro\firstCSMacro\Resources\btnSaveDxfStep.png)

### Bouton : Créer un fichier DFT
Ce bouton permet de créer un fichier DFT (dessin) à partir des fichiers PAR, PSM ou ASM sélectionnés. Voici comment l'utiliser :

1. **Sélectionnez les fichiers** : Dans la liste, sélectionnez un ou plusieurs fichiers PAR, PSM ou ASM.
2. **Cliquez sur le bouton** : Utilisez le bouton **Créer un fichier DFT** pour lancer l'opération.
3. **Vérifiez le fichier DFT** : Un fichier DFT est généré avec des vues, des nomenclatures (si applicable) et des tables de pliage (si applicable).

![Image du bouton](C:\Users\mouad.khalladi\source\repos\firstCSMacro\firstCSMacro\Resources\btnDft.png)

## Autres Fonctionnalités et Astuces

### Bouton : Fermer Solid Edge
Ce bouton va fermer l'application Solid Edge sans sauvegarder son contenu.
Ça arrive que Solid Edge est en cours d'execution et que vous ne le voyez pas. Ce bouton va s'assure qu'aucune instance du programme est en cours d'execution.

### Bouton : Select All
Ce bouton va choisir tout les fichiers affichés dans la liste.

### Barre D'extension
Par défaut, toutes les extensions sont cochées.
Vous pouvez faire **CLIQUE DROIT** avec la souris sur une extension et elle sera l'unique cochée.
Si vous réappuyez sur **CLIQUE DROIT** toutes les extensions vont rendevenir cochés.

---

## Onglet 2 : [Nom de l'onglet]

### Fonctionnalités Principales
- **Fonction 1** : Description de la fonction.
- **Fonction 2** : Description de la fonction.
- **Fonction 3** : Description de la fonction.

### Instructions Simples

#### Bouton : Vérifier les dimensions
Ce bouton permet de vérifier la valeurs longueur et largeur du déplié entre les fichiers Excel et les fichiers DXF. Voici comment l'utiliser :

1. **Sélectionnez les fichiers** : Assurez-vous que les chemins vers le fichier Excel et le répertoire des fichiers DXF sont correctement définis.
2. **Cliquez sur le bouton** : Utilisez le bouton **Vérifier les dimensions** pour lancer l'opération.
3. **Consultez le rapport** : Un rapport détaillé s'affiche, montrant les correspondances, les erreurs et les avertissements.

![Image du bouton] <!-- changer le chemin -->

#### Bouton : Vérifier le nombre de pièces
Ce bouton permet de vérifier la cohérence entre les fichiers Excel, les fichiers DXF et les fichiers STEP. Voici comment l'utiliser :

1. **Sélectionnez les fichiers** : Assurez-vous que les chemins vers le fichier Excel, le(s) répertoire(s) des fichiers DXF et STEP sont correctement définis.
2. **Cliquez sur le bouton** : Utilisez le bouton **Vérifier le nombre de pièces** pour lancer l'opération.
3. **Consultez le rapport** : Un rapport détaillé s'affiche, montrant les correspondances, les erreurs et les avertissements.

![Image du bouton] <!-- changer le chemin -->

---

## Details Fonctions Solid Edge

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

2. **Sélection des répertoires** :
   - Ouvre une fenêtre pour choisir les répertoires de sauvegarde des fichiers DXF et STEP.
   - Permet également d'activer l'option **Tag DXF** pour ajouter des annotations aux fichiers DXF.

3. **Traitement des fichiers** :
   - Pour chaque fichier sélectionné :
     - Ouvre le fichier dans Solid Edge.
     - Vérifie si le fichier contient un modèle déplié (Flat Pattern).
     - Si un modèle déplié est trouvé, le fichier est sauvegardé en DXF et STEP.
     - Si l'option **Tag DXF** est activée, une annotation est ajoutée au fichier DXF.

4. **Gestion des erreurs** :
   - En cas d'erreur pendant le traitement, un message d'erreur est affiché avec les détails.
   - Si un fichier ne contient pas de modèle déplié, un message d'avertissement est affiché.

5. **Nettoyage et fin de l'opération** :
   - Ferme Solid Edge (sauf si l'utilisateur choisit de le garder ouvert).
   - Affiche un message de confirmation une fois l'opération terminée.

### Bouton : Créer un fichier DFT
Ce bouton effectue les actions suivantes :

1. **Vérification des fichiers sélectionnés** :
   - Vérifie qu'au moins un fichier PAR, PSM ou ASM est sélectionné.
   - Affiche un message d'erreur si aucun fichier n'est sélectionné.

2. **Création du fichier DFT** :
   - Ouvre Solid Edge et crée un nouveau document DFT basé sur un modèle prédéfini (`Normal.dft`).
   - Pour chaque fichier sélectionné :
     - Ajoute une nouvelle feuille dans le document DFT.
     - Crée des vues (vue de face, vue de côté, vue de dessous) à partir du fichier PAR, PSM ou ASM.
     - Si le fichier est un assemblage (ASM), une nomenclature est générée automatiquement.
     - Si le fichier est une pièce en tôle (PAR ou PSM), une table de pliage est ajoutée.

3. **Gestion des assemblages** :
   - Si l'option **Dessins individuels pour les assemblages** est activée, une fenêtre permet de sélectionner les composants à inclure dans des feuilles séparées.
   - Les composants sélectionnés sont traités individuellement, avec leurs propres vues et tables de pliage.

4. **Gestion des erreurs** :
   - En cas d'erreur pendant le traitement, un message d'erreur est affiché avec les détails.
   - Si un fichier ne contient pas de modèle déplié (pour les tables de pliage), un message d'avertissement est affiché.

5. **Nettoyage et fin de l'opération** :
   - Le fichier DFT est sauvegardé avec le nom `Dessins dft`.
   - Solid Edge reste ouvert pour permettre à l'utilisateur de visualiser ou de modifier le fichier DFT.

## Details Fonctions Excel QC

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

6. **Gestion des erreurs** :
   - En cas d'erreur pendant le traitement, un message d'erreur est affiché avec les détails.
   - Les erreurs courantes incluent des fichiers Excel mal formatés ou des fichiers DXF manquants.

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

6. **Gestion des erreurs** :
   - En cas d'erreur pendant le traitement, un message d'erreur est affiché avec les détails.
   - Les erreurs courantes incluent des fichiers Excel mal formatés ou des fichiers DXF/STEP manquants.

## Conclusion
Cette application est conçue pour simplifier et automatiser les tâches répétitives liées à Solid Edge et à la gestion des fichiers. Chaque onglet offre des fonctionnalités spécifiques pour répondre à des besoins précis, tout en garantissant une utilisation intuitive et efficace.
