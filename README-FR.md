# Encodeur de codes à barres UPC pour Excel

**📖 Documentation :** [EN English](README.md) | [FR Français](README-FR.md)

![Release](https://img.shields.io/github/v/release/Dolphin2ii/UPC-Barcode-Encoder-Excel-Addin)
![Downloads](https://img.shields.io/github/downloads/Dolphin2ii/UPC-Barcode-Encoder-Excel-Addin/total)
![License](https://img.shields.io/github/license/Dolphin2ii/UPC-Barcode-Encoder-Excel-Addin)

Une solution VBA Excel complète pour générer des codes à barres UPC-A et EAN-13 directement dans Excel sans dépendances externes ou logiciels payants.

## 🚀 Téléchargement rapide

📁 **[Télécharger la dernière version v1.0](https://github.com/Dolphin2ii/UPC-Barcode-Encoder-Excel-Addin/releases/latest)**

Choisissez votre version préférée :
- **UPC Excel Addin.zip** - Package complet avec toutes les langues
- **UPC Excel Addin - ENG.zip** - Version anglaise uniquement  
- **UPC Excel Addin - FR.zip** - Version française uniquement

## 🎯 Aperçu du projet

Ce projet fournit des alternatives gratuites aux logiciels de codes à barres payants (comme TEC-IT) pour générer des codes à barres UPC-A et EAN-13 dans Excel. La solution est conçue pour les environnements d'affaires nécessitant une génération de codes à barres hors ligne et sécurisée.

## ✨ Fonctionnalités

- **🔒 Sécurisé et hors ligne** : Aucune connexion externe ou transmission de données
- **📊 Formats multiples** : Supporte UPC-A (11-12 chiffres) et EAN-13 (13 chiffres)
- **🔍 Détection automatique** : Détecte automatiquement le format de code selon la longueur
- **✅ Calcul de chiffre de contrôle** : Calcule et valide automatiquement les chiffres de contrôle
- **🎨 Support de polices multiples** : Compatible avec les polices Code39 et EAN-13
- **🌍 International** : Compatible avec Excel français et autres versions localisées
- **🆓 100% gratuit** : Aucun coût de licence ou dépendance externe

## 📋 Formats de codes supportés

| Type d'entrée | Exemple | Format de sortie | Police requise |
|---------------|---------|------------------|----------------|
| UPC 11 chiffres | `82542200004` | `*825422000045*` | Code39 |
| UPC 12 chiffres | `123456789012` | `*123456789012*` | Code39 |
| EAN 13 chiffres | `6418029906397` | `*6418029906397*` | Code39 |

## 🚀 Démarrage rapide

### 1. Installer les polices

Extraire et installer les polices de codes à barres fournies :
- `Libre_Barcode_39.zip` - Police Code39 source ouverte (recommandée)

### 2. Télécharger le complément

Choisissez votre version préférée :
- 🇫🇷 **Utilisateurs français**: Télécharger `UPC Excel Addin - FR.zip`
- 🇬🇧 **Utilisateurs anglais**: Télécharger `UPC Excel Addin - ENG.zip`
- 🌍 **Bilingue**: Télécharger `UPC Excel Addin.zip` (contient les deux langues)

### 3. Installer le complément Excel

1. Ouvrir Excel
2. Aller à **Fichier** → **Options** → **Compléments**
3. Sélectionner **Compléments Excel** → **Atteindre...**
4. Cliquer **Parcourir...** et sélectionner votre fichier ZIP téléchargé (extraire d'abord)
5. Cocher la case pour activer le complément

### 4. Activer les macros (Important !)
🔒 **Information sécurité** : Ce complément contient des macros VBA pour la génération de codes à barres.

**Pourquoi les macros sont nécessaires :**
- ✅ **Traitement local uniquement** - Aucune connexion internet requise
- ✅ **Aucune transmission de données** - Tout reste sur votre ordinateur
- ✅ **Code source ouvert** - Tout le code VBA est visible et vérifiable
- ✅ **Aucune dépendance externe** - Fonctions Excel VBA pures uniquement

**Quand Excel affiche un avertissement de sécurité :**
1. Cliquer **"Activer les macros"** - C'est sécuritaire pour ce complément
2. Si bloqué : **Fichier** → **Options** → **Centre de gestion** → **Paramètres du centre de gestion**
3. Sélectionner **Paramètres des macros** → **Activer toutes les macros** (temporairement)
4. Ou ajouter l'emplacement du fichier aux **Emplacements approuvés** (recommandé pour usage permanent)

**Garantie de sécurité :** Ce complément effectue uniquement des calculs mathématiques localement. Aucun accès réseau, aucun accès fichier hors d'Excel.

### 5. Utiliser dans Excel

```excel
=EncodeUPCOrEAN(A1)    # Recommandé - Détection automatique du format
=EncodeUPCA(A1)        # UPC-A seulement (11-12 chiffres)
=EncodeEAN13(A1)       # EAN-13 seulement (13 chiffres)
```

## 📚 Référence des fonctions

### Fonctions principales

#### `EncodeUPCOrEAN(code)`

**Recommandé pour la plupart des cas d'usage**
- Détecte automatiquement UPC-A vs EAN-13 selon la longueur du code
- Retourne une chaîne formatée Code39 pour tous les types de codes
- Exemple : `=EncodeUPCOrEAN("6418029906397")` → `*6418029906397*`

#### `EncodeUPCA(code)`

**Encodage spécifique UPC-A**
- Gère uniquement les codes UPC de 11-12 chiffres
- Calcule automatiquement le chiffre de contrôle pour les codes à 11 chiffres
- Exemple : `=EncodeUPCA("82542200004")` → `*825422000045*`

#### `EncodeEAN13(code)`

**Encodage spécifique EAN-13**
- Gère les codes EAN de 12-13 chiffres
- Retourne le format brut (sans astérisques)
- Exemple : `=EncodeEAN13("6418029906397")` → `6418029906397`

#### `EncodeEAN13AsCode39(code)`

**EAN-13 avec formatage Code39**
- Formate les codes EAN-13 pour les polices Code39
- Meilleure compatibilité avec les lecteurs
- Exemple : `=EncodeEAN13AsCode39("6418029906397")` → `*6418029906397*`

### Fonctions héritées

- `EncodeUPCAAsEAN13()` - Convertit UPC-A au format EAN-13
- `EncodeUPCACompatible()` - Encodage compatible Code128

## 🛠️ Guide d'installation

### Installation de complément

1. **Télécharger** les fichiers du projet
2. **Extraire** les polices et les installer sur votre système
3. **Copier le fichier .xlam** dans le dossier de compléments par défaut d'Excel :
   - Chemin : `C:\Users\[VotreNomUtilisateur]\AppData\Roaming\Microsoft\AddIns\`
   - Ceci assure que le complément reste au bon endroit
4. **Ouvrir Excel** et aller à Fichier → Options → Compléments
5. **Parcourir** vers le fichier `.xlam` dans le dossier AddIns et l'installer
6. **Commencer à utiliser** les fonctions dans vos feuilles de calcul

### Mise à jour du complément

Si vous devez mettre à jour le complément avec de nouvelles fonctionnalités :

1. **Désactiver d'abord le complément actuel** : Fichier → Options → Compléments → Compléments Excel → Atteindre → **Décocher** votre complément UPC → OK
2. **Remplacer** l'ancien fichier `.xlam` avec la nouvelle version dans le dossier AddIns
3. **Réactiver** le complément : Fichier → Options → Compléments → Compléments Excel → Atteindre → **Cocher** votre complément UPC → OK
4. **Redémarrer Excel** pour s'assurer que les changements prennent effet

## 🔧 Dépannage

### Problèmes courants

**"Excel ne peut pas ouvrir deux classeurs portant le même nom"**
- Ceci arrive quand le complément est déjà chargé dans Excel
- Désactiver d'abord le complément : Fichier → Options → Compléments → Décocher votre complément
- Modifier le fichier `.xlam`, puis réactiver

**L'emplacement du complément est important**
- Toujours garder le fichier `.xlam` dans le dossier AddIns d'Excel
- Si vous déplacez le fichier après installation, Excel affichera des erreurs
- Utiliser le dossier : `C:\Users\[NomUtilisateur]\AppData\Roaming\Microsoft\AddIns\`

**Les codes s'affichent mais ne scannent pas**
- S'assurer d'utiliser les polices Code39 (LibreBarcode39 recommandée)
- Vérifier que la cellule contient le résultat de la formule avec astérisques : `*code*`
- Essayer la fonction `EncodeUPCOrEAN()` pour une meilleure compatibilité

**Les codes à 13 chiffres affichent une erreur**
- Utiliser `EncodeUPCOrEAN()` au lieu de `EncodeUPCA()`
- L'ancienne fonction `EncodeUPCA()` ne supporte que 11-12 chiffres

### Information sur le chiffre de contrôle

**Les codes à 11 chiffres obtiennent un chiffre supplémentaire**
- C'est le **comportement correct**
- Le chiffre supplémentaire est le chiffre de contrôle calculé
- Les codes UPC-A font toujours 12 chiffres au total
- Exemple : `82542200004` → `825422000045` (5 est le chiffre de contrôle)

## 📁 Structure du projet

```
UPC-Barcode-Encoder-Excel-Addin/
├── Excel_UPC_Barcode_Business_VBA.bas     # Code source VBA principal (pour référence)
├── Libre_Barcode_39.zip                   # Police Code39 source ouverte (recommandée)
├── UPC Excel Addin - ENG.zip              # Version anglaise pour distribution
├── UPC Excel Addin - FR.zip               # Version française pour distribution
├── UPC Excel Addin.zip                    # Complément complet empaqueté (bilingue)
├── UPC.xlam                               # Fichier complément Excel
├── LICENSE                                # Licence MIT (anglais)
├── LICENSE-FR                             # Licence MIT (français)
├── README.md                              # Documentation (anglais)
└── README-FR.md                           # Documentation (français)
```

**Pour la distribution :** Partager le fichier `UPC Excel Addin.zip` qui contient tout le nécessaire.

## 🧪 Tests

Exécuter la fonction de test intégrée pour vérifier que tout fonctionne :

1. **Ouvrir Excel** avec le complément installé
2. **Appuyer** sur `Alt + F8` pour ouvrir les macros
3. **Exécuter** `TestUPCFunctions`
4. **Vérifier** les résultats dans la Fenêtre Exécution (`Ctrl + G`)

## 🤝 Usage d'affaires

Cette solution a été développée pour remplacer les logiciels de codes à barres payants dans les environnements d'affaires :

- ✅ **Économique** : Élimine les frais de licence
- ✅ **Conforme à la sécurité** : Aucune transmission de données externes
- ✅ **Distribution facile** : Partager le fichier de complément avec les membres de l'équipe
- ✅ **Support international** : Fonctionne avec les versions localisées d'Excel

## 📈 Historique des versions

- **v2.0** (Actuelle) - Ajout du support EAN-13, détection automatique, compatibilité améliorée
- **v1.0** - Support initial UPC-A avec polices Code39

## 🛡️ Sécurité et conformité

- **Aucune connexion externe** requise
- **Aucune transmission de données** en dehors d'Excel
- **Utilise uniquement les fonctionnalités VBA Excel standard**
- **Tout le traitement se fait localement** sur l'ordinateur de l'utilisateur
- **Compatible avec les politiques de sécurité corporatives**

## 📞 Support

Pour les problèmes ou questions :

1. Vérifier la section dépannage ci-dessus
2. Exécuter la fonction de test pour vérifier l'installation
3. S'assurer que les polices sont correctement installées
4. Vérifier que vous utilisez `EncodeUPCOrEAN()` pour les types de codes mixtes

## 📄 Licence

Ce projet est sous licence MIT - voir le fichier [LICENSE-FR](LICENSE-FR) pour les détails.

Les licences des polices s'appliquent séparément :
- Polices LibreBarcode : Source ouverte (SIL Open Font License)

---

**💡 Conseil** : Pour de meilleurs résultats, utiliser la fonction `EncodeUPCOrEAN()` avec les polices LibreBarcode39. Cette combinaison fournit une compatibilité maximale avec tous les lecteurs de codes à barres.