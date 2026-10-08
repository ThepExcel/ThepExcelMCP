<a id="francais"></a>

[Français](#francais) · [Version anglaise ci-dessous](#english-version)

<div align="center">

<img src="assets/banner-thepexcelmcp-1200x630.png" alt="ThepExcelMCP — des agents IA pilotent Excel Desktop en direct via COM" width="720">

# ThepExcelMCP

**Donnez à votre agent IA la capacité d’agir dans Excel, directement dans l’application.**

Un serveur MCP pour Windows qui pilote **Excel Desktop ouvert et en cours d’exécution** par l’automatisation COM.
Des requêtes Power Query qui s’actualisent réellement, des tableaux croisés dynamiques qui calculent,
des mesures DAX évaluées par Excel et des captures pour que l’IA puisse *voir* ce qu’elle construit.

[![CI](https://github.com/ThepExcel/ThepExcelMCP/actions/workflows/ci.yml/badge.svg)](https://github.com/ThepExcel/ThepExcelMCP/actions/workflows/ci.yml)
[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)](pyproject.toml)
[![Plateforme : Windows](https://img.shields.io/badge/platform-Windows-0078D6)](#fr-plateformes)
[![Licence : MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Tests](https://img.shields.io/badge/unit_tests-1000%2B-brightgreen)](tests/)
[![MCP](https://img.shields.io/badge/Model_Context_Protocol-server-8A2BE2)](https://modelcontextprotocol.io)

[Démarrage rapide](#fr-demarrage) ·
[Ajouts du fork](#fr-fork) ·
[Pourquoi COM en direct ?](#fr-com) ·
[Outils](#fr-outils) ·
[Installation](#fr-installation) ·
[Sécurité](#fr-securite) ·
[Dépannage](#fr-depannage)

</div>

---

<a id="fr-fork"></a>

## Extension du fork : supprimer des lignes ou des colonnes en direct

Ce fork ajoute `excel_range(action="delete_rows")` et
`excel_range(action="delete_columns")` sur la branche
[`feat/delete-excel-rows-columns`](https://github.com/fabriceluccioniexternal/ThepExcelMCP/tree/feat/delete-excel-rows-columns).
L’intégration au projet source est proposée dans la [PR #13](https://github.com/ThepExcel/ThepExcelMCP/pull/13).
La comparaison ci-dessous porte sur le `main` source au commit
[`bd3aac3`](https://github.com/ThepExcel/ThepExcelMCP/commit/bd3aac3e188c3f1277f8694b12b00c128db773db),
vérifié le 9 octobre 2026. Elle ne présume pas du contenu des futures versions du projet source.

| Opération | `main` source à `bd3aac3` | Branche de fonctionnalité de ce fork |
|---|---|---|
| `excel_range(action="clear", range="A2:B4")` | Vide les cellules sélectionnées ; conserve leur format et leur position | Même comportement |
| `excel_range(action="delete_rows", range="A2:B4")` | Action non prise en charge | Supprime les lignes 2 à 4 entières, y compris hors des colonnes A–B ; les lignes suivantes remontent |
| `excel_range(action="delete_columns", range="B2:D4")` | Action non prise en charge | Supprime les colonnes B à D entières, y compris hors des lignes 2–4 ; les colonnes à droite se déplacent vers la gauche |
| Suppression de blocs disjoints ou chevauchants | Non prise en charge | Accepte `2:4,9:9` ou `B:D,G:G` ; fusionne les intervalles qui se chevauchent ou se touchent avant suppression |

Les paramètres existants de l’outil et les règles de ciblage du classeur et de la feuille restent identiques.
Excel ajuste les références avec son comportement natif. Le classeur reste ouvert et
**n’est pas sauvegardé automatiquement**. Une opération sur plusieurs blocs peut réussir partiellement ;
les erreurs précisent les intervalles d’origine déjà supprimés.

**Guide détaillé : [Suppression de lignes et de colonnes — différence avec main](docs/suppression-lignes-colonnes.md).**
Il présente des exemples, les données avant/après, le format de réponse, les règles de ciblage et les erreurs.

Pour installer cette extension, choisissez explicitement la branche de fonctionnalité :

```powershell
git clone --branch feat/delete-excel-rows-columns https://github.com/fabriceluccioniexternal/ThepExcelMCP.git
cd ThepExcelMCP
uv sync --frozen
```

Enregistrez le serveur MCP en utilisant ce répertoire selon les instructions du client ci-dessous,
puis redémarrez-le. Les autres liens d’installation et de publication de ce README concernent le projet source ;
son bundle et le `main` inchangé de ce fork ne contiennent pas cette extension au commit de référence.

## Présentation

ThepExcelMCP est un serveur [Model Context Protocol](https://modelcontextprotocol.io) qui expose
**26 outils** couvrant l’essentiel du travail avancé dans Excel : classeurs, plages, Tables,
Power Query (création, lecture, modification et suppression du code M), tableaux croisés dynamiques,
modèle de données avec mesures DAX, graphiques, mise en forme, formats conditionnels, validation,
segments, mise en page et export PDF, protection, VBA sur activation explicite, ainsi que des copies
de sécurité non destructives et des comparaisons de plages ou de feuilles.

Il s’adresse aux personnes qui utilisent **Excel sous Windows** et souhaitent confier un véritable travail
sur leurs classeurs à un agent IA (Claude Code, Claude Desktop, Codex CLI ou tout client MCP) :
analystes automatisant les rapports mensuels, spécialistes Excel créant des traitements Power Query
par une consigne, et développeurs qui doivent construire *et vérifier* des classeurs.

Quand un agent appelle un outil, il interagit avec le processus Excel réel, **y compris le classeur
que vous avez déjà ouvert**. Il utilise l’application au lieu de modifier uniquement le fichier ZIP/XML.

<a id="fr-demarrage"></a>

## Démarrage rapide : confiez la configuration à votre agent IA

Envoyez le lien GitHub à votre agent de développement (Claude Code, Codex CLI, etc.) et demandez-lui
de configurer le serveur. Il lira ce README et pourra l’enregistrer **pour votre utilisateur**
(tous les projets) ou **pour un projet** uniquement. ❤️

```
https://github.com/ThepExcel/ThepExcelMCP
```

Pour les suppressions ajoutées dans ce fork, donnez plutôt le lien de la
[branche de fonctionnalité](https://github.com/fabriceluccioniexternal/ThepExcelMCP/tree/feat/delete-excel-rows-columns).

Une fois le serveur enregistré, vous pouvez demander à votre agent :

- *« Lis la Table de Sheet1 dans Book1 et crée un tableau croisé dynamique des ventes par Region, avec un graphique. »*
- *« Crée une requête Power Query qui charge `products.csv` et `orders.xlsx`, les joint sur ProductID et charge le résultat dans une Table. »*
- *« Ajoute la mesure DAX `Total Revenue = SUM(Orders[Amount])` au modèle de données et applique un format monétaire. »*
- *« Mets en forme l’en-tête du rapport dans Summary : gras, fond `#4472C4`, texte blanc, première ligne figée ; prends ensuite une capture pour vérifier le résultat. »*
- *« Fais d’abord une copie de sécurité du classeur, remplace ensuite toutes les occurrences de “Widget” par “Gadget” et montre-moi les différences. »*

<a id="fr-com"></a>

## Pourquoi piloter Excel ouvert plutôt qu’utiliser une bibliothèque de fichiers ?

Les bibliothèques de fichiers (`openpyxl`, outils xlsx, modification du XML) lisent et écrivent le `.xlsx`
sur disque, mais ne pilotent pas Excel. En contrôlant l’application, l’agent peut construire une solution
complète depuis un classeur vierge, avec des données en direct et les calculs réalisés par Excel :

| Fonctionnalité | Bibliothèques de fichiers (openpyxl, etc.) | ThepExcelMCP (COM en direct) |
|---|---|---|
| Lire/écrire les cellules et formules | ✅ | ✅ |
| **Évaluer les formules et tableaux dynamiques** (`XLOOKUP`, `FILTER`, plages déversées) | ❌ valeurs en cache parfois anciennes | ✅ moteur de calcul réel et relecture du déversement |
| **Power Query** : créer/modifier le code M, actualiser depuis les sources | ❌ | ✅ création/lecture/modification/suppression, actualisation et paramètres |
| **Tableaux croisés dynamiques** avec de véritables agrégations | ❌ souvent endommagés à la sauvegarde | ✅ création, gestion des champs, disposition et lecture |
| **Modèle de données / Power Pivot** : relations, mesures DAX | ❌ | ✅ et fonctions d’aide CUBEVALUE/CUBEMEMBER |
| **Graphiques**, y compris les graphiques croisés dynamiques | ⚠️ possibilités limitées et fragiles | ✅ création/configuration et export PNG |
| **Vérification visuelle** par capture du résultat | ❌ | ✅ plage / feuille / graphique → PNG |
| Travailler dans le **classeur déjà ouvert** | ❌ le fichier doit être fermé | ✅ connexion à l’instance en cours |
| Macros VBA, export PDF, segments, commentaires avec fil de discussion | ❌ | ✅ |

Chaque appel passe par un unique thread COM STA. Les règles d’Excel pour le calcul, la mise en forme
et les événements s’appliquent comme lors d’une utilisation au clavier. Grâce aux **captures**
de plages, feuilles ou graphiques, l’agent peut construire, observer et corriger son travail.

## Fonctionnement local

ThepExcelMCP est un **serveur MCP stdio exécuté comme processus sur votre ordinateur Windows**.
Il contrôle l’Excel de cette machine. Aucune inscription ni aucun compte ne sont nécessaires ;
aucune donnée du classeur n’est envoyée par ce serveur hors de votre ordinateur.

- ✅ **Claude Code**, **Claude Desktop** et **Codex CLI** s’exécutent localement et peuvent le lancer.
- ❌ Les interfaces d’agents uniquement dans le cloud, sans possibilité de lancer un processus sur votre machine, ne peuvent pas accéder à votre Excel local.

<a id="fr-outils"></a>

## Les 26 outils

Les outils sont regroupés par capacité. Chaque outil choisit son opération avec `action="..."` ;
ses docstrings précises et illustrées servent de documentation de l’API pour les modèles.

| Domaine | Outils | Ce que l’agent peut faire |
|---|---|---|
| **Classeurs, feuilles et échanges de données** | `excel_workbook` · `excel_sheet` · `excel_range` · `excel_table` | Ouvrir/créer/sauvegarder des classeurs, gérer les feuilles, lire/écrire des plages avec pagination et déversements, lecture typée ou brute plus rapide via `Value2`, tableaux dynamiques `Formula2`, `=PY()` expérimental ; gérer les Tables Excel (ListObject), ajouts en masse, tri, filtres, styles, totaux et références structurées |
| **Power Query et modèle de données** | `excel_powerquery` · `excel_datamodel` · `excel_name` | Créer/modifier/actualiser des requêtes M avec analyseur statique intégré, paramètres et chargement dans une Table ou le modèle ; Tables du modèle, relations et mesures DAX ; fonctions CUBE, plages nommées et fonctions LAMBDA |
| **Tableaux croisés dynamiques** | `excel_pivot` · `excel_slicer` | Créer depuis une plage, une Table ou le modèle, ajouter/déplacer/supprimer des champs avec agrégations réelles, gérer disposition et sous-totaux, segments et chronologies |
| **Graphiques et vérification visuelle** | `excel_chart` · `excel_screenshot` · `excel_shape` · `excel_sparkline` | Créer/configurer les graphiques, y compris croisés dynamiques, exporter en PNG ; capturer plages/feuilles/graphiques, ajouter images, zones de texte et formes automatiques, créer des graphiques sparkline dans les cellules |
| **Mise en forme, affichage et impression** | `excel_format` · `excel_conditional_format` · `excel_validation` · `excel_view` · `excel_outline` · `excel_page_setup` · `excel_comment` · `excel_hyperlink` | Polices, couleurs, bordures, formats numériques et alignement, formats conditionnels, listes de validation, volets figés, zoom, quadrillage, groupes de lignes/colonnes, mise en page et **export PDF**, notes/commentaires et liens |
| **Sécurité, audit et fonctions avancées** | `excel_snapshot` · `excel_diff` · `excel_find_replace` · `excel_protection` · `excel_vba` | Copies de sécurité non destructives (SaveCopyAs), restauration dans un autre classeur, comparaison cellule par cellule, recherche/comptage/remplacement, protection des feuilles/classeurs, gestion des modules VBA et exécution de macros avec double activation |

<details>
<summary><strong>Liste complète des actions par outil (cliquer pour développer)</strong></summary>

| Outil | Actions |
|---|---|
| `excel_workbook` | `list`, `info`, `open`, `save`, `close`, `create`, `save_as` |
| `excel_sheet` | `list`, `add`, `rename`, `delete` |
| `excel_range` | `read` (pagination, métadonnées de déversement, `value_mode=typed\|raw`), `read_spill`, `write`, `write_formula` (Formula2 / tableaux dynamiques), `write_py` (`=PY()`, expérimental), `clear`, `delete_rows`, `delete_columns` |
| `excel_table` | `list`, `create`, `read` (pagination, `value_mode=typed\|raw`), `append_rows` (redimensionnement unique avec repli par insertion), `add_column` (avec formule), `sort`, `filter`, `set_style`, `toggle_totals`, `rename`, `delete` |
| `excel_powerquery` | `list`, `get`, `create`, `update`, `delete`, `refresh`, `refresh_all`, `load_to_table`, `load_to_datamodel`, `analyze`, `analyze_raw`, `create_parameter`, `get_parameter`, `set_parameter`, `list_parameters` |
| `excel_pivot` | `list`, `create` (source : plage/Table/modèle), `add_field` (agrégation et format numérique), `remove_field`, `move_field`, `set_layout`, `refresh`, `delete`, `read` (pagination, `value_mode=typed\|raw`) |
| `excel_datamodel` | `info`, `list_tables`, `add_table`, `list_relationships`, `add_relationship`, `delete_relationship`, `list_measures`, `add_measure` (DAX), `update_measure`, `delete_measure`, `refresh`, `cube_value`, `cube_member`, `cube_formula` |
| `excel_name` | `list`, `get`, `set`, `delete` (plages nommées, constantes, LAMBDA avec indicateur `is_lambda`) |
| `excel_chart` | `list`, `create`, `configure`, `set_source`, `export_image` (PNG), `delete` |
| `excel_screenshot` | `range`, `sheet`, `chart` |
| `excel_shape` | `add_image`, `add_textbox`, `add_shape`, `list`, `move`, `delete` |
| `excel_slicer` | `add`, `add_timeline`, `list`, `delete`, `connect` |
| `excel_sparkline` | `add` (courbe/colonnes/gains-pertes), `clear`, `list` |
| `excel_format` | `font`, `fill`, `border`, `number_format`, `alignment` (dont fusion), `column_width`, `row_height`, `autofit` |
| `excel_conditional_format` | `data_bar`, `color_scale`, `icon_set`, `cell_rule`, `top_bottom`, `clear` |
| `excel_validation` | `list` (liste déroulante), `whole_number`, `decimal`, `date`, `text_length`, `custom`, `clear` |
| `excel_view` | `freeze_panes`, `unfreeze_panes`, `gridlines`, `zoom`, `headings` |
| `excel_outline` | `group_rows`, `group_columns`, `ungroup_rows`, `ungroup_columns`, `show_levels`, `clear` |
| `excel_page_setup` | `set` (orientation/papier/marges/ajustement), `print_area`, `print_titles`, `header_footer`, `export_pdf` (feuille ou classeur), `get` |
| `excel_comment` | `add`, `edit`, `reply`, `delete`, `list`, `get` (notes classiques et commentaires avec fil de discussion) |
| `excel_hyperlink` | `add` (URL/interne/e-mail/fichier), `list`, `delete` |
| `excel_protection` | `protect_sheet`, `unprotect_sheet`, `protect_workbook`, `unprotect_workbook`, `set_locked`, `status` |
| `excel_find_replace` | `find`, `count`, `replace` (plage/feuille/classeur ; casse, cellule entière) |
| `excel_diff` | `ranges`, `sheets` (valeurs, formules ou les deux ; lecture seule) |
| `excel_snapshot` | `snapshot` (SaveCopyAs), `list`, `restore` (ouvre la copie comme NOUVEAU classeur), `delete` |
| `excel_vba` | `list_modules`, `get_module`, `write_module`, `delete_module`, `run` (activation explicite, voir [Sécurité](#fr-securite)) |

Les références structurées fonctionnent directement : `excel_range(action="read", range="Orders[Amount]")`.

Supprimer des lignes ou des colonnes entières dans un classeur ouvert :

```python
excel_range(action="delete_rows", range="2:4,9:9", sheet="Data", workbook="Test.xlsx")
excel_range(action="delete_columns", range="B:D,G:G", sheet="Data", workbook="Test.xlsx")
```

Les plages de cellules sont également acceptées : supprimer les lignes pour `A2:B4,C9` retire
les lignes 2 à 4 et 9 sur toute leur largeur. Les chevauchements sont fusionnés puis les intervalles
sont supprimés dans l’ordre décroissant pour conserver les indices d’origine. Excel décale les cellules
et ajuste les références. La réponse `deleted` indique le classeur, la feuille, l’axe, les intervalles
d’origine et le nombre supprimé. Le classeur n’est pas sauvegardé. L’opération sur plusieurs blocs
n’est pas atomique : une erreur précise les intervalles déjà supprimés.

</details>

## Prérequis

- **Windows 10 / 11**
- **Microsoft 365 Excel Desktop** : démarrage automatique avec un classeur vierge si Excel n’est pas déjà ouvert (désactivation avec `THEPEXCEL_MCP_AUTOLAUNCH=0`).
- **[uv](https://docs.astral.sh/uv/)** installe et gère la version de Python requise ; vous n’avez **pas** à installer Python séparément :

  ```powershell
  powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
  ```

  Autre possibilité : `winget install --id=astral-sh.uv -e`. Fermez et rouvrez ensuite le terminal.

<a id="fr-installation"></a>

## Installation

Ces commandes installent le projet source. Pour les suppressions de ce fork, utilisez les
[commandes de sa branche de fonctionnalité](#fr-fork).

```powershell
git clone https://github.com/ThepExcel/ThepExcelMCP.git
cd ThepExcelMCP
uv sync
```

`uv sync` installe `fastmcp`, `pywin32` et `pillow` dans un environnement virtuel isolé.
Aucune configuration manuelle de pip ou de virtualenv n’est nécessaire.

### Enregistrement dans Claude Code (CLI)

```powershell
claude mcp add thepexcel-excel --scope user -- uv run --directory C:\path\to\ThepExcelMCP python -m thepexcel_mcp.server
```

Remplacez `C:\path\to\ThepExcelMCP` par le chemin de votre clone. Vérifiez avec `claude mcp list`.

Pour activer aussi VBA :

```powershell
claude mcp add thepexcel-excel --scope user `
  -e THEPEXCEL_MCP_ENABLE_VBA=1 `
  -- uv run --directory C:\path\to\ThepExcelMCP python -m thepexcel_mcp.server
```

### Enregistrement dans Claude Desktop

**Option A : bundle MCPB (un fichier à déposer par glisser-déposer). Recommandé pour Claude Desktop ;
cette installation donne aussi accès au serveur dans l’onglet Cowork.**

1. **Téléchargez `thepexcel-mcp.mcpb` depuis la [dernière publication du projet source](https://github.com/ThepExcel/ThepExcelMCP/releases/latest)** : aucun clone ni construction ne sont requis. `uv` reste nécessaire, car le bundle lance le serveur via `uv run` et ne contient pas Python. Pour construire le bundle vous-même depuis un clone :

   ```powershell
   uv run python scripts/build_mcpb.py
   ```

   Le fichier produit est `dist/thepexcel-mcp.mcpb` (environ 225 Ko : `manifest.json`, `pyproject.toml`, `uv.lock`, `src/`).

2. Dans Claude Desktop, ouvrez **Settings → Desktop app → Extensions**. Déposez `thepexcel-mcp.mcpb` dans la zone *« Drag .MCPB or .DXT files here to install »*, ou utilisez **Advanced settings → Install extension** pour choisir le fichier. Vérifiez l’aperçu du manifeste puis cliquez sur **Install**.

3. Cliquez sur **Configure**, à côté de *ThepExcel Excel MCP* dans *Installed on your computer*. Les options *Auto-launch Excel* (activée par défaut) et *Enable VBA tool* (désactivée par défaut ; nécessite l’autorisation du modèle objet VBA dans Excel) correspondent à `THEPEXCEL_MCP_AUTOLAUNCH` et `THEPEXCEL_MCP_ENABLE_VBA`.

4. Ouvrez une nouvelle conversation. Les outils `excel_*` apparaissent ; le premier appel démarre Excel sur votre machine.

Installation vérifiée le 16 août 2026 sous Windows 11 avec Claude Desktop et l’onglet Cowork :
l’extension s’installe et apparaît dans *Installed on your computer*.

> **« Failed to preview extension: Invalid manifest: server: Required; user_config: Required »** :
> votre bundle a été construit avant le 16 août 2026. L’ancien manifeste utilisait `server.type: "uv"`,
> que le schéma MCPB actuel refuse (types autorisés : `python`, `node`, `binary`, avec un `mcp_config`
> explicite). Récupérez les changements, reconstruisez avec `scripts/build_mcpb.py` et installez le
> nouveau fichier. Vérifiez le manifeste avec `npx @anthropic-ai/mcpb validate manifest.json`.

**Option B : configuration manuelle.** Modifiez `%APPDATA%\Claude\claude_desktop_config.json` :

```json
{
  "mcpServers": {
    "thepexcel-excel": {
      "command": "uv",
      "args": ["run", "--directory", "C:\\path\\to\\ThepExcelMCP", "python", "-m", "thepexcel_mcp.server"]
    }
  }
}
```

Redémarrez Claude Desktop après l’enregistrement.

<details>
<summary><strong>Enregistrement dans Codex CLI</strong></summary>

Codex utilise le même transport MCP stdio. `codex mcp add` écrit dans la configuration
**de votre utilisateur** (`~/.codex/config.toml`) :

```powershell
codex mcp add thepexcel-excel -- uv run --directory C:\path\to\ThepExcelMCP python -m thepexcel_mcp.server
```

Vérifiez avec `codex mcp list` ou `codex mcp get thepexcel-excel`. Configuration manuelle équivalente :

```toml
[mcp_servers.thepexcel-excel]
command = "uv"
args = ["run", "--directory", "C:\\path\\to\\ThepExcelMCP", "python", "-m", "thepexcel_mcp.server"]
```

**Pour un projet uniquement :** ajoutez ce bloc à `.codex/config.toml` à la racine du projet
(aucune option CLI n’existe pour cette portée). Codex doit faire confiance au projet. Les serveurs
configurés par projet sont actuellement chargés uniquement dans la CLI ; Codex Desktop lit seulement
la configuration utilisateur ([openai/codex#13025](https://github.com/openai/codex/issues/13025)).
Redémarrez la session Codex après modification : les outils MCP sont chargés au démarrage.

</details>

<details>
<summary><strong>Utilisation depuis WSL</strong></summary>

Le serveur **ne peut pas s’exécuter dans WSL** (absence de `pywin32` et de COM). Si votre client MCP
fonctionne dans WSL, enregistrez le serveur avec l’exécutable Windows `uv.exe`, pointant vers un clone
sur **Windows** :

```bash
claude mcp add thepexcel-excel --scope user -- \
  uv.exe run --directory 'C:\Tools\ThepExcelMCP' python -m thepexcel_mcp.server
```

Les flux stdio relient WSL à Windows ; le serveur et Excel s’exécutent nativement sous Windows.
Excel doit être ouvert dans votre session de bureau Windows. Cette configuration utilise le même
code Windows ; elle est proposée comme guide et n’a pas été testée séparément comme mode d’installation.

</details>

### Variables d’environnement

| Variable | Valeur par défaut | Description |
|---|---|---|
| `THEPEXCEL_MCP_AUTOLAUNCH` | `1` (activée) | Lance Excel visible avec un classeur vierge si aucune instance ne fonctionne. `0`, `false`, `no` ou `off` exigent une ouverture manuelle. |
| `THEPEXCEL_MCP_ENABLE_VBA` | Non définie (désactivée) | `1` active `excel_vba`. Désactivation par défaut pour la sécurité. |
| `THEPEXCEL_MCP_COM_TIMEOUT` | `120` | Délai maximal par appel COM, en secondes ; augmentez-le pour les actualisations lentes. |
| `THEPEXCEL_MCP_EARLYBIND` | `1` (activée pour les performances) | Liaison anticipée de l’Application COM Excel déjà connectée. Les mesures corrigées montrent un gain modeste : lectures de propriétés 9 % plus rapides lors du dernier test. `0`, `false`, `no` ou `off` imposent une liaison réellement dynamique. |
| `THEPEXCEL_MCP_TOOL_DISCOVERY` | `full` | `bm25` expose les outils d’orientation/lecture et `search_tools`/`call_tool` ; les 26 outils restent appelables et le catalogue initial est environ 94 % plus petit. |

<a id="fr-plateformes"></a>

## Plateformes prises en charge

Ce serveur pilote un processus Excel réel par COM Windows (`pywin32`) : il est
**conçu pour Windows uniquement**.

- ✅ **Windows 10 / 11** : pleinement pris en charge.
- ❌ **macOS** : Excel pour Mac n’a pas d’API COM et `pywin32` n’existe pas sur cette plateforme. Il s’agit d’une limitation de la plateforme.
- ⚠️ **WSL** : utilisable avec la configuration décrite ci-dessus ; le serveur reste un processus Windows natif.

<a id="fr-securite"></a>

## Modèle de sécurité

Le serveur intègre des protections pour encadrer le pilotage de votre Excel par un agent IA :

- **Les copies de sécurité sont non destructives.** `excel_snapshot` utilise `SaveCopyAs` pour copier le classeur sur disque sans modifier son état sauvegardé, son nom ou son chemin. `restore` ouvre la copie comme **nouveau classeur séparé**, à côté de l’original. Il ne ferme, n’écrase ni ne rétablit en place le classeur sur lequel vous travaillez. Encouragez votre agent à créer une copie avant les opérations importantes.
- **Contrôlez les changements.** `excel_diff` compare des plages ou feuilles cellule par cellule (valeurs, formules ou les deux) en lecture seule, pour vérifier l’avant/après.
- **VBA exige deux activations.** `excel_vba` nécessite `THEPEXCEL_MCP_ENABLE_VBA=1` et l’autorisation Excel *Fichier → Options → Centre de gestion de la confidentialité → Paramètres → Paramètres des macros → Accès approuvé au modèle d’objet du projet VBA*. Si l’une manque, l’erreur la précise.
- **Vérifiez les effets réels.** Les vérifications des outils modifiant le classeur portent sur l’état des cellules, formats ou fichiers ; les erreurs COM sont remontées sous forme de `ToolError` exploitable.
- **Exécution locale.** Transport stdio, votre machine, votre Excel ; aucun service réseau ni aucune télémétrie.

### Limites connues

- **Le chargement du modèle de données peut bloquer dans un contexte stdio sans interface.** `excel_datamodel(add_table)` et `excel_powerquery(load_to_datamodel)` déclenchent une actualisation Mashup nécessitant la boucle de messages de l’interface Excel. Dans une session CLI stdio, le thread COM peut rester bloqué jusqu’à l’arrêt forcé d’Excel. Cela fonctionne avec une fenêtre Excel pleinement visible, par exemple dans Claude Desktop. **Solution de repli vérifiée de bout en bout :** `load_to_table` → `excel_pivot(create, source="<table>")`. Le flux jointure Power Query entre fichiers → Table → tableau croisé → graphique croisé → segment fonctionne ainsi.
- **Les captures exigent une fenêtre Excel visible.** `excel_screenshot` et `excel_chart(export_image)` utilisent `CopyPicture`, qui nécessite l’affichage à l’écran ; une fenêtre masquée ou réduite peut produire un PNG vide.
- **Les paramètres LAMBDA** ne doivent pas ressembler à des adresses de cellules (`q1`, `x2`) : utilisez `val`, `rate`, `n`. Un ajout LAMBDA échoué peut laisser des noms cachés `_xlpm.*` bloquant les suivants ; utilisez alors un nouveau classeur.
- **Python dans Excel (`=PY()`)** s’exécute de façon asynchrone dans le cloud Microsoft. Un abonnement M365 avec Python dans Excel est requis ; l’outil n’attend pas le résultat du cloud.
- **VBA `run`** renvoie les scalaires (Long/String/Double) issus de Functions ; les Subs renvoient None. Les tableaux et objets ne sont pas pris en charge.

<a id="fr-depannage"></a>

## Dépannage

| Symptôme | Cause et solution |
|---|---|
| Le serveur ne démarre pas ; Windows bloque `thepexcel-mcp.exe` par une stratégie de contrôle des applications | Smart App Control / WDAC bloque le lanceur non signé créé par `uv` dans `.venv\Scripts`. Les commandes de ce README utilisent `python -m thepexcel_mcp.server`, qui exécute le même code sans ce lanceur. Enregistrez à nouveau le serveur si votre commande se termine encore par `thepexcel-mcp`. |
| « Excel is not running » alors que vous n’avez pas lancé Excel | Le lancement automatique est **activé par défaut**, avec une fenêtre visible et un classeur vierge. Si vous avez défini `THEPEXCEL_MCP_AUTOLAUNCH=0`, ouvrez Excel manuellement. |
| `AttributeError: ... CLSIDToClassMap` présenté comme « Excel not running » | Le cache de liaison anticipée `gen_py` de win32com est corrompu. Le serveur le répare automatiquement : il efface le cache et réessaie une fois. |
| Classeur d’une deuxième instance Excel introuvable | Le serveur recherche aussi dans la Running Object Table (ROT) de Windows si le classeur n’est pas dans la première instance. |
| Le comportement ne change pas après modification du code | Le serveur stdio garde l’ancien code en mémoire. Une installation éditable ne recharge pas le processus. Redémarrez le serveur MCP, par exemple avec une nouvelle session du client. |
| Une actualisation Power Query lente dépasse le délai | Augmentez `THEPEXCEL_MCP_COM_TIMEOUT` (secondes ; valeur par défaut 120). |
| Tous les appels se bloquent après un chargement du modèle de données | Il s’agit du blocage connu ci-dessus. Arrêtez Excel, redémarrez le serveur et utilisez `load_to_table`. |
| PNG vide | Gardez la fenêtre Excel visible et non réduite pendant `excel_screenshot` ou `export_image`. |

## Utiliser la skill fournie

Le dépôt fournit une skill Claude dans [`skills/excel-god/`](skills/excel-god/). Ce guide de stratégie
et d’orchestration aide l’agent à choisir les outils et leur ordre d’appel pour les tableaux de bord,
le nettoyage Power Query, la configuration du modèle de données et d’autres tâches courantes.

```bash
# Pour un projet
cp -r skills/excel-god /path/to/your/project/.claude/skills/

# Pour votre utilisateur, dans tous les projets
cp -r skills/excel-god ~/.claude/skills/
```

Invoquez ensuite `/excel-god` dans une session Claude Code.

## Développement

```powershell
uv sync --frozen                                              # Installer les dépendances verrouillées
uv run pytest -q                                              # Plus de 1 000 tests unitaires, COM simulé, sans Excel
uv run ruff check src scripts --select E9,F                  # Contrôles de syntaxe, imports et analyse statique
uv run python tests/smoke_com.py                              # Tests COM réels sous Windows avec Excel
uv run python tests/smoke_com.py --sections 1,2,3,4           # Sous-ensemble des sections 1 à 29
uv run --isolated --with mcp==2.0.0 python tests/protocol_smoke_v2.py  # Négociation stdio réelle
uv run python scripts/build_mcpb.py                           # Construire dist/thepexcel-mcp.mcpb
```

La suite COM réelle vérifie les effets par relecture dans Excel, avec sa propre instance.
L’exécution des 29 sections prend environ 5 à 10 minutes.

Organisation du projet :

```
src/thepexcel_mcp/
├── server.py          # FastMCP, 26 outils ; docstrings de l’API destinée aux modèles
├── session.py         # ExcelSession, thread COM STA, run_com(), recherche ROT, protections
├── domains/           # Un module par outil : classeurs, plages, Power Query, tableaux croisés...
└── analysis/          # Analyse statique du code M pour excel_powerquery
tests/                 # Tests unitaires avec COM simulé, tests du protocole et d’Excel réel
skills/excel-god/      # Skill fournie pour l’agent
scripts/build_mcpb.py  # Construction du bundle MCPB pour Claude Desktop
```

Chaque gestionnaire soumet son opération COM à un **unique thread STA** via `run_com()`.
Ce thread possède l’appartement COM, sérialise les opérations, gère les délais et supprime
les alertes Excel (`DisplayAlerts`) autour des opérations qui le nécessitent.

## Remerciements

Le projet s’appuie sur les travaux de la communauté de l’automatisation Excel et du MCP :

- **[sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel)** (MIT) : principale référence d’implémentation. Plusieurs séquences COM Excel ont été étudiées et portées de C# vers Python. Ce projet constitue aussi un serveur MCP Excel COM mature en C#/.NET.
- **[lingfan36/ai-office-mcp](https://github.com/lingfan36/ai-office-mcp)** : inspiration pour les copies de sécurité, l’annulation et la comparaison de plages.
- **[haris-musa/excel-mcp-server](https://github.com/haris-musa/excel-mcp-server)** : référence pour les conventions d’API des outils MCP Excel, avec une approche par fichiers utilisant `openpyxl`.

Les textes des licences de ces projets figurent dans [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).

## Contribuer

Travaillez sur une branche et ouvrez une pull request ; `main` est protégé. Utilisez uniquement
des **données synthétiques**, car le dépôt est public, et activez le hook de sécurité avant push.
Consultez [CONTRIBUTING.md](CONTRIBUTING.md) et le [Code de conduite](CODE_OF_CONDUCT.md).

Signalez les problèmes de sécurité de façon privée selon [SECURITY.md](SECURITY.md).

## Licence

[MIT](LICENSE). Copyright (c) 2026 ThepExcel <thepexcel@gmail.com>.

[Retour au début de la version française](#francais) · [Version anglaise ci-dessous](#english-version)

---

<a id="english-version"></a>

# English version

[Retour à la version française](#francais)

<div align="center">

<img src="assets/banner-thepexcelmcp-1200x630.png" alt="ThepExcelMCP — AI agents driving live Excel Desktop via COM" width="720">

# ThepExcelMCP

**Give your AI agent hands on the real Excel — not just the file.**

A Windows MCP server that drives a **live, running Excel Desktop** through COM automation.
Power Query that actually refreshes, PivotTables that actually pivot, DAX that actually calculates,
and screenshots so the AI can *see* what it built.

[![CI](https://github.com/ThepExcel/ThepExcelMCP/actions/workflows/ci.yml/badge.svg)](https://github.com/ThepExcel/ThepExcelMCP/actions/workflows/ci.yml)
[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue)](pyproject.toml)
[![Platform: Windows](https://img.shields.io/badge/platform-Windows-0078D6)](#platform-support)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Tests](https://img.shields.io/badge/unit_tests-1000%2B-brightgreen)](tests/)
[![MCP](https://img.shields.io/badge/Model_Context_Protocol-server-8A2BE2)](https://modelcontextprotocol.io)

[Quick start](#quick-start--hand-it-to-your-ai-agent) ·
[Fork changes](#fork-extension--live-rowcolumn-deletion) ·
[Why live COM?](#why-a-live-excel-not-a-file-library) ·
[Tools](#the-26-tools) ·
[Install](#install) ·
[Safety](#safety-model) ·
[Troubleshooting](#troubleshooting)

</div>

---

## Fork extension — live row/column deletion

This fork adds `excel_range(action="delete_rows")` and
`excel_range(action="delete_columns")` on the branch
[`feat/delete-excel-rows-columns`](https://github.com/fabriceluccioniexternal/ThepExcelMCP/tree/feat/delete-excel-rows-columns).
The change is proposed upstream in [PR #13](https://github.com/ThepExcel/ThepExcelMCP/pull/13).
The comparison below uses upstream `main` at
[`bd3aac3`](https://github.com/ThepExcel/ThepExcelMCP/commit/bd3aac3e188c3f1277f8694b12b00c128db773db),
checked on 2026-10-09; it does not assume that future upstream versions lack these actions.

| Operation | Upstream `main` at `bd3aac3` | This fork's feature branch |
|---|---|---|
| `excel_range(action="clear", range="A2:B4")` | Empties the selected cells; keeps their formatting and position | Same behavior |
| `excel_range(action="delete_rows", range="A2:B4")` | Unsupported action | Removes entire rows 2–4, including cells outside columns A–B; lower rows move up |
| `excel_range(action="delete_columns", range="B2:D4")` | Unsupported action | Removes entire columns B–D, including cells outside rows 2–4; columns to their right move left |
| Disjoint/overlapping deletion selections | Unsupported | Accepts `2:4,9:9` or `B:D,G:G`; merges overlapping/adjacent intervals before deleting |

The existing tool parameters and workbook/sheet targeting rules stay the same. Excel
adjusts references using its native deletion behavior. The workbook stays open and is
**not automatically saved**. A multi-block operation can partially succeed; errors
identify the original intervals already removed.

**Lire le guide en français : [Suppression de lignes et de colonnes — différence avec main](docs/suppression-lignes-colonnes.md).**
It includes examples, before/after data, the response format, targeting rules and failure handling.

To install this extension, use the feature branch explicitly:

```powershell
git clone --branch feat/delete-excel-rows-columns https://github.com/fabriceluccioniexternal/ThepExcelMCP.git
cd ThepExcelMCP
uv sync --frozen
```

Register the MCP server against this checkout using the client instructions below,
then restart it. The upstream installation/release links elsewhere in this README
refer to the original project; its release bundle and this fork's unchanged `main`
do not include this extension at the comparison revision.

## What is this?

ThepExcelMCP is a [Model Context Protocol](https://modelcontextprotocol.io) server that exposes
**26 tools** covering nearly everything a power user can do in Excel — workbooks, ranges, Tables,
Power Query (full M-code CRUD), PivotTables, the Data Model with DAX measures, charts, formatting,
conditional formats, validation, slicers, page setup / PDF export, protection, VBA (opt-in), and a
safety layer with non-destructive snapshots and range/sheet diffing.

It is for anyone who works in **Excel on Windows** and wants an AI agent
(Claude Code, Claude Desktop, Codex CLI, or any MCP client) to do real spreadsheet work:
analysts automating monthly reporting, Excel pros building Power Query pipelines by prompt,
and developers who need an agent to build *and verify* real workbooks.

When an agent calls a tool here, it is talking to the actual Excel process — **including the
workbook you already have open**. This is the difference between editing an XML zip file and
actually using Excel.

## Quick start — hand it to your AI agent

Send this GitHub link to your AI coding agent (Claude Code, Codex CLI, …) and tell it to take
over — it will read this README and set the MCP server up for you, as either **user scope**
(every project) or **project scope** (one project only). ❤️

> 🇹🇭 แค่ส่งลิงก์ GitHub นี้ให้ AI Agent ของคุณ แล้วบอกให้ AI จัดการต่อได้เลย จะลง MCP เป็น User Scope หรือ Project Scope ก็ได้ ❤️

```
https://github.com/ThepExcel/ThepExcelMCP
```

Once registered, just talk to your agent. Things you can say:

- *"Read the table on Sheet1 of Book1 and build a PivotTable of sales by Region, with a chart."*
- *"Write a Power Query that loads `products.csv` and `orders.xlsx`, merges them on ProductID, and loads the result into a Table."*
- *"Add a DAX measure `Total Revenue = SUM(Orders[Amount])` to the Data Model and format it as currency."*
- *"Format the report header on the Summary sheet — bold, fill `#4472C4`, white text, freeze the top row — then screenshot it so you can check it looks right."*
- *"Take a snapshot of this workbook first, then replace every occurrence of 'Widget' with 'Gadget' across the whole workbook and show me a diff."*

## Why a live Excel, not a file library?

File-based libraries (`openpyxl`, xlsx skills, raw XML editing) read and write the `.xlsx` on
disk — but they cannot *operate Excel*. Driving the real application means the agent can author
a complete solution from a blank workbook, on live data, with Excel doing the computing:

| Capability | File libraries (openpyxl etc.) | ThepExcelMCP (live COM) |
|---|---|---|
| Read/write cells & formulas | ✅ | ✅ |
| **Evaluate formulas / dynamic arrays** (`XLOOKUP`, `FILTER`, spill) | ❌ stale cached values | ✅ live calc engine, spill read-back |
| **Power Query** — create/edit M code, refresh against sources | ❌ | ✅ full CRUD + refresh + parameters |
| **PivotTables** with real aggregation | ❌ (often destroyed on save) | ✅ create, field ops, layout, read |
| **Data Model / Power Pivot** — relationships, DAX measures | ❌ | ✅ + CUBEVALUE/CUBEMEMBER helpers |
| **Charts** incl. true PivotCharts | ⚠️ limited, fragile | ✅ create/configure + PNG export |
| **Visual verification** — screenshot what was built | ❌ | ✅ range / sheet / chart → PNG |
| Work on the **workbook you already have open** | ❌ file must be closed | ✅ attaches to the running instance |
| VBA macros, PDF export, slicers, threaded comments | ❌ | ✅ |

Every call runs through a single STA COM worker thread, so Excel's own rules for calculation,
formatting, and events apply exactly as they would for a human at the keyboard. And because the
agent can **screenshot** any range, sheet, or chart, it can close the loop: build → look → fix.

## How it runs — local, not a hosted service

ThepExcelMCP is a **stdio MCP server that runs as a process on your own Windows machine** and
controls the Excel running there. Nothing to sign up for, no account, and no spreadsheet data
leaves your computer.

- ✅ **Claude Code**, **Claude Desktop**, and **Codex CLI** run locally and can launch it.
- ❌ Cloud-only agent surfaces that can't run a local process on your machine can't reach your
  local Excel, so they can't use it.

## The 26 tools

Grouped by capability — every tool is action-dispatched (`action="..."`), with precise,
example-rich docstrings that serve as the LLM-facing API.

| Area | Tools | What the agent can do |
|---|---|---|
| **Workbooks, sheets & data I/O** | `excel_workbook` · `excel_sheet` · `excel_range` · `excel_table` | Open/create/save workbooks, manage sheets, read/write ranges (paginated, spill-aware, typed or faster raw `Value2` reads, `Formula2` dynamic arrays, experimental `=PY()`), full Excel Table (ListObject) lifecycle including bulk append, sort, filter, styles, totals, and structured references |
| **Power Query & Data Model** | `excel_powerquery` · `excel_datamodel` · `excel_name` | Create/edit/refresh M queries with a built-in M static analyzer, query parameters, load to Table or Data Model; model tables, relationships, DAX measures; CUBE formula helpers; named ranges & LAMBDA functions |
| **PivotTables** | `excel_pivot` · `excel_slicer` | Create pivots from a range, Table, or the Data Model; add/move/remove fields with real aggregations; layouts & subtotals; slicers and date timelines |
| **Charts & visual verification** | `excel_chart` · `excel_screenshot` · `excel_shape` · `excel_sparkline` | Create/configure charts (incl. true PivotCharts), export chart PNGs; capture any range/sheet/chart as PNG so the agent can *see* its work; images, text boxes, AutoShapes; in-cell sparklines |
| **Formatting, layout & print** | `excel_format` · `excel_conditional_format` · `excel_validation` · `excel_view` · `excel_outline` · `excel_page_setup` · `excel_comment` · `excel_hyperlink` | Fonts/fills/borders/number formats/alignment, data bars & color scales & icon sets, dropdown validation, freeze panes/zoom/gridlines, row-column grouping, print setup + **PDF export**, notes & threaded comments, hyperlinks |
| **Safety, audit & power tools** | `excel_snapshot` · `excel_diff` · `excel_find_replace` · `excel_protection` · `excel_vba` | Non-destructive snapshots (SaveCopyAs) with safe restore, cell-by-cell diff of ranges or whole sheets, find/count/replace at range/sheet/workbook scope, sheet & workbook protection, VBA module CRUD + macro run (double opt-in) |

<details>
<summary><strong>Full action list per tool (click to expand)</strong></summary>

| Tool | Actions |
|---|---|
| `excel_workbook` | `list`, `info`, `open`, `save`, `close`, `create`, `save_as` |
| `excel_sheet` | `list`, `add`, `rename`, `delete` |
| `excel_range` | `read` (paginated, spill metadata, `value_mode=typed\|raw`), `read_spill`, `write`, `write_formula` (Formula2 / dynamic arrays), `write_py` (`=PY()`, experimental), `clear`, `delete_rows`, `delete_columns` |
| `excel_table` | `list`, `create`, `read` (paginated, `value_mode=typed\|raw`), `append_rows` (single-resize fast path with safe insertion fallback), `add_column` (with formula), `sort`, `filter`, `set_style`, `toggle_totals`, `rename`, `delete` |
| `excel_powerquery` | `list`, `get`, `create`, `update`, `delete`, `refresh`, `refresh_all`, `load_to_table`, `load_to_datamodel`, `analyze`, `analyze_raw`, `create_parameter`, `get_parameter`, `set_parameter`, `list_parameters` |
| `excel_pivot` | `list`, `create` (range/table/datamodel source), `add_field` (aggregation + number format), `remove_field`, `move_field`, `set_layout`, `refresh`, `delete`, `read` (paginated, `value_mode=typed\|raw`) |
| `excel_datamodel` | `info`, `list_tables`, `add_table`, `list_relationships`, `add_relationship`, `delete_relationship`, `list_measures`, `add_measure` (DAX), `update_measure`, `delete_measure`, `refresh`, `cube_value`, `cube_member`, `cube_formula` |
| `excel_name` | `list`, `get`, `set`, `delete` (named ranges, constants, LAMBDA — with `is_lambda` flag) |
| `excel_chart` | `list`, `create`, `configure`, `set_source`, `export_image` (PNG), `delete` |
| `excel_screenshot` | `range`, `sheet`, `chart` |
| `excel_shape` | `add_image`, `add_textbox`, `add_shape`, `list`, `move`, `delete` |
| `excel_slicer` | `add`, `add_timeline`, `list`, `delete`, `connect` |
| `excel_sparkline` | `add` (line/column/win-loss), `clear`, `list` |
| `excel_format` | `font`, `fill`, `border`, `number_format`, `alignment` (incl. merge), `column_width`, `row_height`, `autofit` |
| `excel_conditional_format` | `data_bar`, `color_scale`, `icon_set`, `cell_rule`, `top_bottom`, `clear` |
| `excel_validation` | `list` (dropdown), `whole_number`, `decimal`, `date`, `text_length`, `custom`, `clear` |
| `excel_view` | `freeze_panes`, `unfreeze_panes`, `gridlines`, `zoom`, `headings` |
| `excel_outline` | `group_rows`, `group_columns`, `ungroup_rows`, `ungroup_columns`, `show_levels`, `clear` |
| `excel_page_setup` | `set` (orientation/paper/margins/fit-to-page), `print_area`, `print_titles`, `header_footer`, `export_pdf` (sheet or workbook), `get` |
| `excel_comment` | `add`, `edit`, `reply`, `delete`, `list`, `get` (legacy notes + threaded comments) |
| `excel_hyperlink` | `add` (url/internal/email/file), `list`, `delete` |
| `excel_protection` | `protect_sheet`, `unprotect_sheet`, `protect_workbook`, `unprotect_workbook`, `set_locked`, `status` |
| `excel_find_replace` | `find`, `count`, `replace` (range/sheet/workbook scope; match-case, whole-cell) |
| `excel_diff` | `ranges`, `sheets` (values, formulas, or both — pure read) |
| `excel_snapshot` | `snapshot` (SaveCopyAs), `list`, `restore` (opens copy as a NEW workbook), `delete` |
| `excel_vba` | `list_modules`, `get_module`, `write_module`, `delete_module`, `run` (opt-in, see [Safety](#safety-model)) |

Structured references work naturally: `excel_range(action="read", range="Orders[Amount]")`.

Delete whole worksheet rows or columns in an open workbook:

```python
excel_range(action="delete_rows", range="2:4,9:9", sheet="Data", workbook="Test.xlsx")
excel_range(action="delete_columns", range="B:D,G:G", sheet="Data", workbook="Test.xlsx")
```

Cell ranges also work: deleting rows for `A2:B4,C9` removes all of rows 2-4
and 9, including cells outside the selected columns. Overlapping intervals are
merged, then removed from highest index to lowest so disjoint selections retain
their original indices. Excel shifts cells and adjusts references. The response
reports the workbook, sheet, axis, original intervals and deleted count under
`deleted`. The workbook is not saved. Multi-block deletion is not atomic: an
error reports any intervals already removed.

</details>

## Requirements

- **Windows 10 / 11**
- **Microsoft 365 Excel Desktop** — auto-launched with a blank workbook if not already running
  (opt out with `THEPEXCEL_MCP_AUTOLAUNCH=0`)
- **[uv](https://docs.astral.sh/uv/)** — installs and manages the right Python for you, so you do
  **not** need to install Python separately:

  ```powershell
  powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
  ```

  Or via winget: `winget install --id=astral-sh.uv -e`. Close and reopen your terminal afterward.

## Install

```powershell
git clone https://github.com/ThepExcel/ThepExcelMCP.git
cd ThepExcelMCP
uv sync
```

`uv sync` installs `fastmcp`, `pywin32`, and `pillow` into an isolated virtual environment.
No pip or manual virtualenv setup required.

### Register with Claude Code (CLI)

```powershell
claude mcp add thepexcel-excel --scope user -- uv run --directory C:\path\to\ThepExcelMCP python -m thepexcel_mcp.server
```

Substitute `C:\path\to\ThepExcelMCP` with your own clone path. Verify with `claude mcp list`.

To enable VBA as well:

```powershell
claude mcp add thepexcel-excel --scope user `
  -e THEPEXCEL_MCP_ENABLE_VBA=1 `
  -- uv run --directory C:\path\to\ThepExcelMCP python -m thepexcel_mcp.server
```

### Register with Claude Desktop

**Option A — MCPB bundle (one file, drag-and-drop). Recommended for Claude Desktop, and it is
the path that reaches the Cowork tab too.**

1. Get the bundle. **Download `thepexcel-mcp.mcpb` from the
   [latest release](https://github.com/ThepExcel/ThepExcelMCP/releases/latest)** — no clone, no
   build. (You still need `uv` on this machine — the bundle launches the server through `uv run`,
   it does not vendor Python.) To build it yourself from a clone instead:

   ```powershell
   uv run python scripts/build_mcpb.py
   ```

   This produces `dist/thepexcel-mcp.mcpb` (~225 KB: `manifest.json`, `pyproject.toml`,
   `uv.lock`, `src/`).

2. In Claude Desktop open **Settings → Desktop app → Extensions**, then either drag
   `thepexcel-mcp.mcpb` onto the *"Drag .MCPB or .DXT files here to install"* area or use
   **Advanced settings → Install extension** and pick the file. The preview shows the manifest;
   click **Install**.

3. **Configure** (button next to *ThepExcel Excel MCP* under *Installed on your computer*):
   *Auto-launch Excel* (default on) and *Enable VBA tool* (default off — needs the VBA project
   object model trust setting in Excel). These map to `THEPEXCEL_MCP_AUTOLAUNCH` /
   `THEPEXCEL_MCP_ENABLE_VBA`.

4. Start a new chat. The `excel_*` tools appear in the tool list; the first call launches Excel on
   your machine.

Verified 2026-08-16 on Windows 11 / Claude Desktop with the Cowork tab: install succeeds and the
extension shows under *Installed on your computer*.

> **"Failed to preview extension: Invalid manifest: server: Required; user_config: Required"** —
> you have a bundle built before 2026-08-16. Older `manifest.json` used `server.type: "uv"`, which
> the current MCPB schema does not accept (allowed: `python` / `node` / `binary` plus an explicit
> `mcp_config`). Pull, rebuild with `scripts/build_mcpb.py`, and install the new file. You can
> check any bundle's manifest with `npx @anthropic-ai/mcpb validate manifest.json`.

**Option B — manual config.** Edit `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "thepexcel-excel": {
      "command": "uv",
      "args": ["run", "--directory", "C:\\path\\to\\ThepExcelMCP", "python", "-m", "thepexcel_mcp.server"]
    }
  }
}
```

Restart Claude Desktop after saving.

<details>
<summary><strong>Register with Codex CLI</strong></summary>

Codex talks MCP over the same stdio transport. `codex mcp add` writes to your **user-scoped**
config (`~/.codex/config.toml`):

```powershell
codex mcp add thepexcel-excel -- uv run --directory C:\path\to\ThepExcelMCP python -m thepexcel_mcp.server
```

Verify: `codex mcp list` / `codex mcp get thepexcel-excel`. Equivalent manual edit:

```toml
[mcp_servers.thepexcel-excel]
command = "uv"
args = ["run", "--directory", "C:\\path\\to\\ThepExcelMCP", "python", "-m", "thepexcel_mcp.server"]
```

**Project-scoped:** add the same block to `.codex/config.toml` in the project root (there is no
CLI flag for project scope). Codex must trust the project, and project-scoped servers currently
load on the CLI only — Codex Desktop reads just the user config
([openai/codex#13025](https://github.com/openai/codex/issues/13025)). Restart the Codex session
after editing — MCP tools load at session start.

</details>

<details>
<summary><strong>Using from WSL</strong></summary>

The server itself **cannot run inside WSL** (no `pywin32`, no COM). But if your MCP client runs
inside WSL, register the server so the *command* is the Windows `uv.exe` pointing at a **Windows**
checkout:

```bash
claude mcp add thepexcel-excel --scope user -- \
  uv.exe run --directory 'C:\Tools\ThepExcelMCP' python -m thepexcel_mcp.server
```

The stdio pipes bridge the WSL→Windows boundary; the server and Excel both run natively on
Windows. Excel must be open in your Windows desktop session. This path runs the same code on
Windows under the hood — offered as guidance, not a separately tested install mode.

</details>

### Environment variables

| Variable | Default | Description |
|---|---|---|
| `THEPEXCEL_MCP_AUTOLAUNCH` | `1` (on) | Auto-launch a visible Excel (+ blank workbook) if none is running. Set `0`/`false`/`no`/`off` to require Excel be opened manually. |
| `THEPEXCEL_MCP_ENABLE_VBA` | unset (off) | Set `1` to enable the `excel_vba` tool. Off by default for security. |
| `THEPEXCEL_MCP_COM_TIMEOUT` | `120` | Per-call COM timeout in seconds. Increase for slow data refreshes. |
| `THEPEXCEL_MCP_EARLYBIND` | `1` (on) — perf default | Early-bind the already-attached Excel COM Application. Corrected wrapper-forced benchmarks showed modest gains (9% faster property reads in the latest run). Set `0`/`false`/`no`/`off` to force a truly dynamic wrapper. |
| `THEPEXCEL_MCP_TOOL_DISCOVERY` | `full` | Set `bm25` to expose only orientation/read tools plus `search_tools`/`call_tool`; the full 26-tool API remains callable and the initial catalog is about 94% smaller. |

## Platform support

This server drives a real Excel process through Windows COM (`pywin32`), so it is
**Windows-only by design**:

- ✅ **Windows 10 / 11** — fully supported.
- ❌ **macOS** — Excel for Mac has no COM automation API and `pywin32` does not exist there.
  Hard platform limitation, not a missing feature.
- ⚠️ **WSL** — supported via the cross-boundary setup above (server runs as a native Windows process).

## Safety model

Letting an AI agent drive your real Excel deserves guardrails. They are built in:

- **Snapshots are non-destructive by construction.** `excel_snapshot` uses `SaveCopyAs` — it
  streams a copy to disk **without** touching the live workbook's saved-state, name, or path.
  `restore` opens the copy as a **separate new workbook** alongside your original; it never
  closes, overwrites, or reverts the workbook you are editing. There is deliberately no
  in-place-revert code path. Encourage your agent to snapshot before risky bulk operations.
- **Audit what changed.** `excel_diff` compares two ranges or whole sheets cell-by-cell (values,
  formulas, or both) as a pure read — perfect for before/after verification.
- **VBA is double-gated.** The `excel_vba` tool requires **both** the
  `THEPEXCEL_MCP_ENABLE_VBA=1` environment variable **and** Excel's own trust setting
  (*File → Options → Trust Center → Trust Center Settings → Macro Settings → "Trust access to
  the VBA project object model"*). Without both, calls return a clear error naming the missing gate.
- **Verified effects, not just success codes.** Mutating tools read back the actual cell /
  format / file state after acting; COM errors surface as actionable `ToolError` messages.
- **Local only.** stdio transport, your machine, your Excel. No network service, no telemetry.

### Known limitations

- **Data Model load can deadlock in headless stdio contexts.** `excel_datamodel(add_table)` and
  `excel_powerquery(load_to_datamodel)` trigger a Mashup refresh that needs Excel's UI message
  pump; in a CLI stdio session this can deadlock the COM worker until Excel is force-killed
  (it works with a fully visible Excel window, e.g. under Claude Desktop). **Fallback that is
  verified end-to-end:** `load_to_table` → `excel_pivot(create, source="<table>")` — the full
  cross-file Power Query merge → Table → PivotTable → PivotChart → slicer flow works this way.
- **Screenshots need a visible Excel window.** `excel_screenshot` and `excel_chart(export_image)`
  use `CopyPicture`, which requires the window to render on screen — a minimized/hidden Excel can
  yield empty PNGs.
- **LAMBDA parameter names** must not look like cell references (`q1`, `x2`) — use `val`, `rate`,
  `n`. A failed LAMBDA add can leave hidden `_xlpm.*` names that block later adds; use a fresh
  workbook if that happens.
- **`=PY()` (Python in Excel)** is inserted but executed asynchronously by Microsoft's cloud
  service — requires an M365 subscription with Python in Excel; the tool does not wait for the
  cloud result.
- **VBA `run`** returns scalars (Long/String/Double) from Functions; Subs return None; arrays and
  objects are not supported.

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| Server won't start; Windows says an Application Control policy blocked `thepexcel-mcp.exe` | Smart App Control / WDAC blocks the unsigned launcher shim `uv` writes into `.venv\Scripts`. The commands in this README launch `python -m thepexcel_mcp.server` instead, which runs the same code without that shim — re-register with the current command if yours still ends in `thepexcel-mcp`. |
| "Excel is not running" but you didn't start it | You don't have to — auto-launch is **on by default** (visible Excel + blank workbook). If you set `THEPEXCEL_MCP_AUTOLAUNCH=0`, open Excel yourself first. |
| `AttributeError: ... CLSIDToClassMap` masquerading as "Excel not running" | Corrupt win32com `gen_py` early-binding cache. The server **self-heals** this: it clears the cache and retries once, automatically. |
| Workbook open in a *second* Excel instance not found | Handled: the server scans the Windows Running Object Table (ROT) as a fallback when a workbook isn't in the first Excel instance. |
| You edited the server code but behavior didn't change | The registered stdio server keeps old code in memory — an editable install is **not** hot reload. Restart the MCP server (e.g. start a fresh client session). |
| Slow Power Query refresh times out | Raise `THEPEXCEL_MCP_COM_TIMEOUT` (seconds; default 120). |
| All tool calls hang after a Data-Model load | The known deadlock above — force-close Excel, restart the server, and use the `load_to_table` fallback. |
| Empty screenshot PNGs | Keep the Excel window visible (not minimized) during `excel_screenshot` / `export_image`. |

## Using the bundled skill

The repo ships a Claude skill at [`skills/excel-god/`](skills/excel-god/) — a strategy and
orchestration guide that helps an agent decide which tools to call, in what order, for common
jobs (dashboards, Power Query data cleaning, Data Model setup, …).

```bash
# project-level
cp -r skills/excel-god /path/to/your/project/.claude/skills/

# user-level (available in all projects)
cp -r skills/excel-god ~/.claude/skills/
```

Then invoke it with `/excel-god` in a Claude Code session.

## Development

```powershell
uv sync --frozen                                              # install exactly the locked dependencies
uv run pytest -q                                              # 1000+ unit tests — mocked COM, no Excel needed
uv run ruff check src scripts --select E9,F                  # syntax/import/static checks
uv run python tests/smoke_com.py                              # live COM smoke suite (Windows + Excel)
uv run python tests/smoke_com.py --sections 1,2,3,4           # subset (sections 1–29)
uv run --isolated --with mcp==2.0.0 python tests/protocol_smoke_v2.py  # real stdio handshake
uv run python scripts/build_mcpb.py                           # build dist/thepexcel-mcp.mcpb
```

The live smoke suite performs real read-back verification against a running Excel (it launches
its own instance) and takes roughly 5–10 minutes for all 29 sections.

Project layout, in brief:

```
src/thepexcel_mcp/
├── server.py          # FastMCP app — 26 tool registrations; docstrings are the LLM-facing API
├── session.py         # ExcelSession — STA COM worker thread, run_com(), ROT fallback, guards
├── domains/           # one module per tool (workbook, ranges, powerquery, pivots, datamodel, …)
└── analysis/          # M-code static analyzer used by excel_powerquery
tests/                 # 1000+ unit tests (mocked COM) + protocol/live Excel smoke suites
skills/excel-god/      # bundled agent skill
scripts/build_mcpb.py  # MCPB bundle builder for Claude Desktop
```

Architecture in one line: every tool handler submits its COM callable to a single dedicated
**STA worker thread** (`run_com()`), which owns the COM apartment — serialized, timeout-guarded,
with `DisplayAlerts` suppressed around risky operations.

## Acknowledgments

This project stands on the work of others in the Excel-automation and MCP community:

- **[sbroenne/mcp-server-excel](https://github.com/sbroenne/mcp-server-excel)** (MIT) — the
  primary reference implementation; a number of Excel COM automation sequences were studied and
  ported from this C# project to Python. If you want a mature C#/.NET COM-based Excel MCP server,
  theirs is excellent.
- **[lingfan36/ai-office-mcp](https://github.com/lingfan36/ai-office-mcp)** — design inspiration
  for the snapshot/undo and range-diff tooling.
- **[haris-musa/excel-mcp-server](https://github.com/haris-musa/excel-mcp-server)** — a reference
  for MCP tool API design conventions (file-based, `openpyxl`).

Upstream license texts are reproduced in [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md).

## Contributing

Work on a branch and open a pull request — `main` is protected. Please use **synthetic data
only** (this is a public repo) and enable the pre-push safety hook. See
[CONTRIBUTING.md](CONTRIBUTING.md) and our [Code of Conduct](CODE_OF_CONDUCT.md).

Found a security issue? Please report it privately — see [SECURITY.md](SECURITY.md).

## License

[MIT](LICENSE). Copyright (c) 2026 ThepExcel <thepexcel@gmail.com>.
