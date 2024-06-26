# characters-list Project
Ce script Python permet de compter le nombre de lettres dans les dialogues de chaque personnage et le nombre d'occurrences des personnages dans un document Word (.docx) ou OpenDocument (.odt) ou PDF (.pdf). Il est conçu pour traiter des documents contenant des tableaux avec des colonnes spécifiques pour les personnages et les dialogues.

### Prérequis
- Python 3.12.3 ou supérieur

- Les bibliothèques Python suivantes :
  - python-docx
  - odfpy

### Installation des dépendances
Pour installer les dépendances nécessaires, exécutez la commande suivante :
```sh
pip install python-docx odfpy
```

### Utilisation

Exécutez le script à l'aide de la commande suivante dans un terminal ou une invite de commande :
```
./count_characters votre_fichier.docx
```
(ou)
```
./count_characters votre_fichier.odt
```

Au besoin tapez la commande suivante pour obtenir de l'aide :
```sh
./count_characters -h
```

### Fonctionnalités
- Compte le nombre de lettres dans les dialogues pour chaque personnage.
- Compte le nombre d'occurrences de chaque personnage.
- Exclut les personnages apparaissant moins de 5 fois, qui sont listés séparément.

#### Exemple d'affichage des résultats :
```
Nombre de lettres par personnage :
JDOE: 305 lettres
ASSI: 66 lettres

Nombre d'occurrences par personnage :
JDOE: 12 occurrences
ASSI: 7 occurrences

Total de lettres par personnage divisé par 50 :
JDOE: 6.10
ASSI: 1.32

Personnages avec moins de 5 occurrences :
DAVID: 2 occurrences
OUSSA: 1 occurrences
```

### Faire son propre exécutable
Pour créer votre propre exécutable, vous pouvez utiliser la bibliothèque PyInstaller. Pour ce faire, exécutez la commande suivante :
```sh
pyinstaller --onefile --windowed votre-script.py
```

### Compatibilité
L'executable fournit est compatible avec Windows seulement, pour MacOS il faudra exécuter le script Python. Vous pouvez l'utiliser sur des formats de fichiers '.docx' et '.odt'.

### Avertissement
Le script ne fonctionne que dans un tableau avec une colonne Timecode, une colonne Personnage et une colonne Dialogue (l'ordre est important mais pas les noms). Pensez à modifier l'en-tête du script pour qu'il corresponde à votre chemin vers Python. Cela fonctionne bien sur les fichiers '.docx' mais il se peut que cela fonctionne moins bien sur les fichiers '.odt', donc privilégiez les fichiers '.docx'.
