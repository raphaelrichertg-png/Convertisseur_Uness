# Convertisseur XLSX → UNESS (QTI 2.1)

Deux outils pour convertir une banque de QCM au format **Excel**
(template des banques Wooclap : colonnes `Type, Title, Correct, Choice, Choice, ...`)
en archive **ZIP QTI 2.1** importable dans **UNESS** (et Theia) via l'import QTI :

- **`index.html`** : page web, tout se fait dans le navigateur (aucun fichier envoyé
  sur un serveur). Si GitHub Pages est activé, elle est utilisable directement en ligne.
- **`xlsx2uness_qti.py`** : version en ligne de commande, utile en lot.

Un fichier d'exemple est fourni : **`exemple.xlsx`** (4 questions au format attendu),
à tester tel quel avec l'un des deux outils.

L'archive générée contient un fichier `.xml` par question (format `SINGLECHOICE`
à l'identique d'un export Theia confirmé fonctionnel dans UNESS) et un fichier
`pool.xml` qui centralise les références des questions.

## Version web

Ouvrez `index.html` (ou la page GitHub Pages du dépôt), glissez le fichier Excel,
vérifiez l'aperçu puis téléchargez le ZIP QTI.

## Version ligne de commande une banque de QCM au format **Excel**
(template des banques Wooclap : colonnes `Type, Title, Correct, Choice, Choice, ...`)
en archive **ZIP QTI 2.1** importable dans **UNESS** (et Theia) via l'import QTI.

L'archive générée contient un fichier `.xml` par question (format `SINGLECHOICE`
à l'identique d'un export Theia confirmé fonctionnel dans UNESS) et un fichier
`pool.xml` qui centralise les références des questions.

## Utilisation

```bash
python xlsx2uness_qti.py mesquestions.xlsx
# -> mesquestions-UNESS-QTI.zip    (à importer dans UNESS)

# options
python xlsx2uness_qti.py mesquestions.xlsx -o export_test2.zip
python xlsx2uness_qti.py mesquestions.xlsx --titre "OCE - Semestre 2 - Dossier"
python xlsx2uness_qti.py mesquestions.xlsx --feuille Questions
```

Aucune dépendance : Python 3 standard uniquement (Windows, macOS, Linux).

## Format du fichier Excel d'entrée

Feuille (idéalement nommée `Questions`), en-têtes en ligne 1 :

| Type | Title | Correct | Choice | Choice | Choice | Choice |
|---|---|---|---|---|---|---|
| Poll | Énoncé de la question | réponse attendue | proposition A | proposition B | proposition C | proposition D |

La colonne `Correct` accepte trois écrits :

1. **le texte exact de la bonne proposition** (recommandé, format banque Wooclap) ;
2. **le numéro** de la bonne proposition (`3` = 3e proposition) ;
3. **la lettre** de la bonne proposition (`B`).

Si la colonne `Correct` est vide, la 1re proposition est retenue et un
avertissement listant les questions concernées est affiché.

## Remarques sur le format de sortie (UNESS / QTI 2.1)

- chaque question est une `assessmentItem` QTI 2.1 avec
  `responseDeclaration cardinality="single"` et `correctResponse` ;
- les interactions sont de type `choiceInteraction maxChoices="1"`
  (question à réponse unique, QRU) ;
- `pool.xml` référence chaque question avec `assessmentItemRef` et un poids (`WEIGHT` = 1) ;
- l'HTML de l'énoncé et des propositions est échappé à l'identique des
  exports Theia (`<div>&lt;p&gt;…&lt;/p&gt;</div>`).

## Exemple testé

```
$ python xlsx2uness_qti.py exemple.xlsx
Feuille lue           : Questions
Questions converties   : 4
Archive QTI générée    : exemple-UNESS-QTI.zip
-> Importez ce ZIP dans UNESS via l'import QTI.
```
