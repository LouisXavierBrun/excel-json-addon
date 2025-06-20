# Excel JSON Export/Import Add-in

Module complémentaire Excel pour exporter et importer des données JSON.

## Fonctionnalités

- 📤 **Export JSON** : Exporte la première feuille du classeur au format JSON
- 📥 **Import JSON** : Importe des données JSON dans une feuille sélectionnée
- 🔄 **Gestion des révisions** : Remplace le contenu existant avec suivi des modifications
- 💾 **Téléchargement direct** : Sauvegarde les fichiers JSON exportés
- 📋 **Interface intuitive** : Interface utilisateur moderne et facile à utiliser

## Installation

### Méthode 1 : Via le manifest direct

1. **Téléchargez le fichier manifest :**
   - Clic droit sur [manifest.xml](https://VOTRE-USERNAME.github.io/excel-json-addon/manifest.xml)
   - "Enregistrer le lien sous..." → `manifest.xml`

2. **Dans Excel en ligne :**
   - Ouvrez Excel en ligne (office.com)
   - Créez ou ouvrez un classeur
   - Allez dans `Insérer` → `Applications` → `Télécharger à partir d'un fichier`
   - Sélectionnez le fichier `manifest.xml` téléchargé
   - Cliquez sur `Insérer`

### Méthode 2 : Pour Excel Desktop

1. **Ajoutez le catalogue :**
   - `Fichier` → `Options` → `Centre de gestion de la confidentialité`
   - `Paramètres du Centre de gestion de la confidentialité`
   - `Catalogues de compléments approuvés`
   - Cochez `Autoriser les catalogues de compléments sur le Web`
   - Ajoutez l'URL : `https://VOTRE-USERNAME.github.io/excel-json-addon/`

2. **Installez le complément :**
   - `Insérer` → `Mes compléments` → `Dossier partagé`
   - Sélectionnez le manifest JSON Export/Import

## Utilisation

### Export JSON
1. Ouvrez le panneau du module dans le ruban Excel
2. Assurez-vous que votre première feuille contient des données avec en-têtes
3. Cliquez sur "Exporter la première feuille"
4. Le JSON s'affiche dans la zone de texte
5. Cliquez sur "Télécharger le fichier JSON" pour sauvegarder

### Import JSON
1. Sélectionnez la feuille de destination
2. Choisissez un fichier JSON ou collez le contenu
3. Cliquez sur "Importer les données"
4. Les données sont importées avec formatage automatique

## Structure du projet

```
excel-json-addon/
├── manifest.xml          # Configuration du module
├── taskpane.html         # Interface utilisateur
├── commands.html         # Fichier de commandes
├── assets/              # Icônes du module
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   └── icon-80.png
└── README.md            # Documentation
```

## Développement

Pour modifier ce module :

1. **Clonez le repository :**
   ```bash
   git clone https://github.com/VOTRE-USERNAME/excel-json-addon.git
   ```

2. **Modifiez les fichiers selon vos besoins**

3. **Testez localement :**
   - Servez les fichiers via un serveur web
   - Utilisez le manifest local pour les tests

4. **Déployez :**
   - Poussez les modifications sur GitHub
   - GitHub Pages se met à jour automatiquement

## Support

- Compatible avec Excel en ligne et Excel Desktop
- Nécessite Excel 2016 ou supérieur
- Fonctionne avec tous les navigateurs modernes

## Licence

MIT License - Voir le fichier LICENSE pour plus de détails.
