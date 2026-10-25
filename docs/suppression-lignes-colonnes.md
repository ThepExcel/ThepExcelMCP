# Suppression de lignes et de colonnes dans Excel ouvert

La branche [`feat/delete-excel-rows-columns`](https://github.com/fabriceluccioniexternal/ThepExcelMCP/tree/feat/delete-excel-rows-columns) de ce fork ajoute deux actions à l'outil existant `excel_range` : `delete_rows` et `delete_columns`. Elles agissent immédiatement dans un classeur ouvert dans Excel Desktop sous Windows. Aucun paramètre de l'outil n'est ajouté.

La comparaison porte sur le `main` du projet source au commit [`bd3aac3`](https://github.com/ThepExcel/ThepExcelMCP/commit/bd3aac3e188c3f1277f8694b12b00c128db773db), vérifié le 9 octobre 2026. La proposition d'intégration est la [pull request #13](https://github.com/ThepExcel/ThepExcelMCP/pull/13). Les versions futures de `main` peuvent évoluer.

## Ce que le fork ajoute à main

| Besoin | `main` au commit de référence | Branche de fonctionnalité du fork |
|---|---|---|
| Vider des cellules | `clear` efface les valeurs et formules sélectionnées | Même comportement |
| Supprimer des lignes entières | `delete_rows` est une action inconnue | Supprime les lignes et remonte les données situées en dessous |
| Supprimer des colonnes entières | `delete_columns` est une action inconnue | Supprime les colonnes et déplace vers la gauche celles situées à droite |
| Supprimer plusieurs blocs en un appel | Pas d'action de suppression de lignes/colonnes dans `excel_range` | Accepte une sélection composée de plusieurs blocs sur une seule feuille |

Les autres outils ne sont pas remplacés : `excel_sheet(action="delete")` supprime une feuille et `excel_table(action="delete")` retire un objet Table selon ses options. Ces opérations ont une portée différente de la suppression de lignes ou de colonnes de la feuille.

## Vider des cellules ou supprimer leur ligne/colonne

```python
# Vider uniquement A2:B4 : la grille et le format restent en place.
excel_range(action="clear", range="A2:B4",
            sheet="Data", workbook="Test.xlsx")

# Supprimer les lignes 2, 3 et 4 sur TOUTE la largeur de la feuille.
excel_range(action="delete_rows", range="A2:B4",
            sheet="Data", workbook="Test.xlsx")

# Supprimer les colonnes B, C et D sur TOUTE la hauteur de la feuille.
excel_range(action="delete_columns", range="B2:D4",
            sheet="Data", workbook="Test.xlsx")
```

Ces exemples sont des alternatives, pas une séquence à exécuter sur les mêmes données. Une sélection exprimée en cellules détermine les lignes ou colonnes traversées ; elle ne limite pas la suppression aux seules cellules sélectionnées.

Excel décale les données et ajuste les références avec son comportement natif. Une référence à une cellule supprimée peut devenir `#REF!`. La suppression ne remplace pas la ligne ou la colonne par des cellules vides à sa position d'origine.

## Sélections simples, multiples et chevauchantes

| Action | `range` | Lignes/colonnes d'origine supprimées | Nombre |
|---|---|---|---|
| `delete_rows` | `A3` ou `3:3` | Ligne 3 | 1 |
| `delete_rows` | `2:4` | Lignes 2 à 4 | 3 |
| `delete_rows` | `2:4,9:9` ou `A2:B4,C9` | Lignes 2 à 4 et 9 | 4 |
| `delete_rows` | `2:4,3:6,9:9,9:9` | Lignes 2 à 6 et 9, une seule fois chacune | 6 |
| `delete_columns` | `B2` ou `B:B` | Colonne B | 1 |
| `delete_columns` | `B:D,G:G` | Colonnes B à D et G | 4 |

Les blocs se séparent par une virgule, y compris avec Excel en français. Tous les blocs sont résolus avant la première suppression. Les intervalles qui se chevauchent ou se touchent sont fusionnés. Les lignes sont supprimées du bas vers le haut et les colonnes de droite à gauche, afin d'utiliser les indices de la sélection d'origine.

Exemple : avant l'appel, `A1:A10` contient les nombres 1 à 10. Après `delete_rows` sur `2:4,9:9`, `A1:A6` contient :

| Ligne actuelle | Valeur conservée | Ligne d'origine |
|---|---|---|
| 1 | 1 | 1 |
| 2 | 5 | 5 |
| 3 | 6 | 6 |
| 4 | 7 | 7 |
| 5 | 8 | 8 |
| 6 | 10 | 10 |

La ligne 9 supprimée est donc celle de la sélection initiale, même si les autres suppressions changent ensuite les positions. De même, avec des valeurs 1 à 10 dans `A1:J1`, supprimer `B:D,G:G` laisse `[1, 5, 6, 8, 9, 10]` dans `A1:F1`.

## Choisir le classeur et la feuille

Passez `workbook` et `sheet` pour cibler explicitement le classeur ouvert et la feuille. Si ces paramètres sont omis, les règles existantes utilisent le classeur actif et sa feuille active.

Une feuille indiquée dans `range` prend priorité sur le paramètre `sheet` :

```python
excel_range(action="delete_rows", range="'Other Sheet'!A2:B4,C9",
            sheet="Data", workbook="Test.xlsx")
```

Cet appel supprime les lignes de `Other Sheet`, pas celles de `Data`. Un seul qualificatif de feuille, au début de la sélection, s'applique à tous les blocs. Une sélection comme `Data!2:4,Other!9:9` est rejetée avant toute suppression.

## Comprendre le résultat et les erreurs

Pour `delete_rows` sur `2:4,9:9`, la réponse a cette forme :

```json
{
  "deleted": {
    "workbook": "Test.xlsx",
    "sheet": "Data",
    "axis": "rows",
    "intervals": [
      {"start": 2, "end": 4, "range": "$2:$4"},
      {"start": 9, "end": 9, "range": "$9:$9"}
    ],
    "count": 4
  }
}
```

`start` et `end` sont des indices numériques à partir de 1, avant suppression. Pour les colonnes, B à D correspondent à 2 à 4 et `axis` vaut `columns`. Les intervalles retournés sont fusionnés et triés par indice croissant ; `count` compte les lignes ou colonnes distinctes supprimées, pas les cellules.

Une plage vide ou invalide empêche la suppression. Les erreurs COM, par exemple une feuille protégée, sont converties en `ToolError`. Une opération composée de plusieurs blocs n'est pas atomique : si le bloc 9 est supprimé puis que le bloc 2 à 4 échoue, le message indique les intervalles d'origine déjà supprimés et leur nombre. Aucun retour arrière automatique n'est effectué. Relisez les données avant de réessayer, car leurs positions ont changé.

Le classeur reste ouvert et **aucune sauvegarde automatique n'est effectuée**. Pour conserver les changements sur disque, utilisez ensuite explicitement `excel_workbook(action="save", workbook="Test.xlsx")`.

## Installer la bonne branche

Le `main` de ce fork est conservé au niveau du projet source. Pour obtenir les suppressions, clonez la branche de fonctionnalité :

```powershell
git clone --branch feat/delete-excel-rows-columns https://github.com/fabriceluccioniexternal/ThepExcelMCP.git
cd ThepExcelMCP
uv sync --frozen
```

Configurez votre client MCP pour lancer ce répertoire selon les [instructions du README](../README.md#install), puis redémarrez le serveur. Une installation éditable ne recharge pas un processus déjà démarré. Les liens d'installation et de publication du projet source présents dans le README désignent la version officielle ; son bundle au commit de référence n'inclut pas cette extension.

## Validation de l'implémentation

L'implémentation a été vérifiée avec 1 074 tests unitaires réussis, les contrôles Ruff `E9,F` et un test du protocole MCP. La section 29 de `tests/smoke_com.py` ajoute cinq cas Excel réels : suppressions de lignes et colonnes, plage invalide, feuille protégée et chevauchements. Les relectures vérifient les valeurs déplacées, les formules ajustées et les données conservées.

Ces cinq cas ont été exécutés dans un classeur temporaire contenant uniquement des données synthétiques, dans une instance Excel séparée. La totalité des autres sections COM n'a pas été rejouée pour cette extension.
