# Dimensionnement hydraulique — Points hauts, intermédiaires & points bas

Application Python de dimensionnement des équipements de protection des conduites
(ventouses aux points hauts et intermédiaires, vidanges aux points bas), avec
interface graphique **CustomTkinter**, mode ligne de commande, rapports **Excel**
et **Word**, et **catalogue MCP des ventouses intégré au code**.

> Fichier principal : [`main1.py`](main1.py)

---

## Sommaire

- [Fonctionnalités](#fonctionnalités)
- [Structure du projet](#structure-du-projet)
- [Installation](#installation)
- [Utilisation](#utilisation)
- [Formats des fichiers d'entrée](#formats-des-fichiers-dentrée)
- [Paramètres de calcul](#paramètres-de-calcul)
- [Méthodologie](#méthodologie)
- [Catalogue MCP des ventouses](#catalogue-mcp-des-ventouses)
- [Sorties générées](#sorties-générées)
- [Exécutable Windows](#exécutable-windows)
- [Licence](#licence)
- [Dépannage](#dépannage)
- [Feuille de route](#feuille-de-route)
- [Auteur](#auteur)

---

## Fonctionnalités

- **Deux modes de calcul** :
  - `Profil en Long (PL)` : détection automatique des points hauts / bas sur un
    profil, ligne piézométrique critique, équipements associés.
  - `Cas Simples` : saisie directe de trois jeux de données (points hauts,
    points intermédiaires, points bas).
- **Hydraulique** : résolution itérative de **Colebrook** pour le frottement
  linéaire, pertes singulières de sortie, débits de rupture et de remplissage.
- **Dimensionnement** : choix du nombre de ventouses (`1V150`, `2V250`, …) et du
  diamètre de vidange aux points bas.
- **Interface graphique complète** (CustomTkinter) : barre de menus
  (*Fichier / Calcul / Catalogue / Aide*), barre d'outils, bouton **Importer un
  exemple**, choix interactif du dossier de sortie, ouverture automatique du
  rapport Word, onglets *Résultats*, *Graphiques* et *Catalogue MCP*, guide
  utilisateur et fenêtre À propos, barre de statut et raccourcis clavier.
- **Rapports automatiques** : classeur Excel multi-onglets et note de calcul Word
  avec tableaux et graphiques.
- **Catalogue MCP intégré** : débits, courbes, dimensions et règles d'ingénierie
  des ventouses double orifice 3 fonctions PN 10/16, exportés dans les rapports.

---

## Structure du projet

```
MK_Prandtl-Colebrook_V0/
├── main1.py                                   # Application (GUI + CLI + catalogue MCP)
├── requirements.txt                           # Dépendances Python
├── install.txt                                # Rappel d'installation / lancement
├── README.md                                  # Ce fichier
├── LICENSE                                    # Licence MIT
├── MK_Prandtl-Colebrook_V01.spec              # Recette PyInstaller de l'exe
├── assets/
│   ├── MK_icon.png                            # Icône source (1024 px)
│   └── MK_icon.ico                            # Icône Windows (7 tailles)
├── dist/
│   └── MK_Prandtl-Colebrook_V01.exe           # Exécutable Windows autonome
├── Exemple_Import_Profil_En_Long.xlsx         # Exemple d'import — mode Profil en Long
├── Exemple_Import_Cas_Simples.xlsx            # Exemple d'import — mode Cas Simples
├── Tableau_Recapitulatif_Debits_Ventouses_MCP.xlsx  # Catalogue source (3 feuilles)
├── Tableau_Recapitulatif_Debits_Ventouses_MCP.docx  # Synthèse catalogue source
├── env/                                       # Environnement virtuel fourni (Python 3.13)
└── sorties_app/                               # Rapports générés par défaut
    ├── Resultats_Hydrauliques.xlsx
    ├── Note_Calcul_Hydraulique.docx
    └── figures/
        ├── fig_pl_profil_piezo.png
        ├── fig_pl_ventouses.png
        └── fig_pl_vidanges.png
```

---

## Installation

Un environnement virtuel `env/` est déjà fourni à la racine du projet.

```powershell
# Activation de l'environnement (Windows PowerShell)
.\env\Scripts\Activate.ps1

# Installation / vérification des dépendances
python -m pip install -r requirements.txt
```

Dépendances : `customtkinter`, `pandas`, `numpy`, `matplotlib`, `openpyxl`,
`python-docx`, `Pillow` (versions minimales indiquées dans `requirements.txt`).

---

## Utilisation

### Interface graphique

```powershell
python main1.py
# équivalent : python main1.py --gui
```

1. Choisir le **mode de calcul** (Profil en Long ou Cas Simples) — aussi disponible
   dans le menu *Calcul*.
2. Charger un fichier via **« Sélectionner fichier .xlsx »**, ou cliquer sur
   **« Importer un exemple ▾ »** : le classeur d'exemple est chargé et le mode
   correspondant est présélectionné automatiquement.
3. Ajuster, si besoin, les hypothèses hydrauliques dans le panneau de gauche.
4. Indiquer le **dossier de sortie** (`sorties_app` par défaut) en le saisissant
   directement ou via **« Parcourir... »** (menu *Fichier > Choisir le dossier de sortie*).
5. Cocher ou décocher **« Ouvrir le rapport Word après génération »** et
   **« Inclure le catalogue MCP dans les rapports »** selon le besoin.
6. Cliquer sur **« Lancer le calcul & générer rapports »** (touche **F5**).
7. Consulter le journal, l'aperçu des graphiques et l'onglet **Catalogue MCP**.
   La note Word s'ouvre automatiquement si l'option est activée ; les boutons
   dédiés permettent aussi d'ouvrir le rapport et le dossier de sortie.

### Menus et raccourcis

| Menu | Actions principales |
|---|---|
| **Fichier** | Sélectionner un fichier, Importer un exemple, **Enregistrer**, **Enregistrer sous…**, Choisir le dossier de sortie, Ouvrir le dossier / le rapport, Quitter |
| **Calcul** | Lancer le calcul, changer de mode, ouvrir le rapport après génération |
| **Catalogue** | Consulter le résumé MCP, ouvrir les fichiers sources, inclure le catalogue dans les rapports |
| **Aide** | Guide utilisateur, À propos |

Raccourcis : **Ctrl+O** (fichier d'entrée), **F5** (calcul), **Ctrl+S**
(enregistrer les rapports), **Ctrl+Shift+S** (enregistrer sous un autre dossier),
**Ctrl+G** (guide).

Les commandes *Enregistrer* et *Enregistrer sous…* réexportent le dernier calcul
(Excel + Word) sans le refaire ; si aucun calcul n'a encore été lancé, *Enregistrer*
déclenche automatiquement le calcul. *Enregistrer sous…* demande un dossier, y copie
les graphiques éventuels et enregistre les rapports à cet endroit.

### Ligne de commande

```powershell
# Profil en long
python main1.py --mode pl --input-xlsx profil.xlsx --out-dir sorties_PL

# Cas simples
python main1.py --mode simple --input-xlsx cas_simples.xlsx --out-dir sorties_simples
```

| Option | Description |
|---|---|
| `--gui` | Force le lancement de l'interface graphique |
| `--mode {pl,simple}` | Mode de calcul en ligne de commande |
| `--input-xlsx` | Chemin du classeur d'entrée (obligatoire en CLI) |
| `--out-dir` | Dossier de sortie (défaut : `sorties_cli`) |

---

## Formats des fichiers d'entrée

### Mode `pl` — onglet `Profil_en_Long`

Colonnes attendues : `piquet`, `x_m`, `z_m`, `dn_mm`, `pn_bar`.

Les points hauts et points bas sont déduits automatiquement de la cote `z_m`
(par comparaison avec les piquets voisins).

### Mode `simple` — trois onglets

| Onglet | Colonnes attendues |
|---|---|
| `Points_Hauts` | `piquet_h`, `x_h`, `z_h`, `dn_mm`, `z_am`, `x_am`, `z_av`, `x_av`, `pn_bar`, (`ig1_force_pct`, `ig2_force_pct` optionnels) |
| `Points_Intermediaires` | `piquet_i`, `x_i`, `z_i`, `x_h`, `z_h`, `x_b`, `z_b`, `dn_mm` |
| `Points_Bas` | `piquet_pb`, `x_av`, `x_am`, `h_moy`, `d_equiv_mm`, `x_pb`, `z_pb`, `dn_vid_mm`, `pn_bar` |

Les colonnes `ig1_force_pct` / `ig2_force_pct` permettent de forcer un gradient
(en %) plutôt que de le déduire de la géométrie.

---

## Exemples d'import

Deux classeurs prêts à l'emploi sont fournis à la racine du projet. Ils sont
auto-documentés (feuille `Notice` ignorée par l'application) et ont été testés
avec les deux modes de calcul.

| Fichier | Mode | Contenu |
|---|---|---|
| `Exemple_Import_Profil_En_Long.xlsx` | `pl` | 17 piquets sur 4,6 km (PK 0+000 → PK 4+600), 4 points hauts, 3 points bas, 3 points intermédiaires équipés. Feuille `Profil_en_Long`. |
| `Exemple_Import_Cas_Simples.xlsx` | `simple` | 3 points hauts (dont un avec gradient forcé à 0,5 %), 3 points intermédiaires (1 cas NON, 2 cas OUI), 3 points bas. Feuilles `Points_Hauts`, `Points_Intermediaires`, `Points_Bas`. |

Utilisation immédiate en ligne de commande :

```powershell
python main1.py --mode pl --input-xlsx Exemple_Import_Profil_En_Long.xlsx --out-dir sorties_exemple_pl
python main1.py --mode simple --input-xlsx Exemple_Import_Cas_Simples.xlsx --out-dir sorties_exemple_simple
```

Dans l'interface, le bouton **« Importer un exemple ▾ »** (barre d'outils et
panneau latéral, ou menu *Fichier > Importer un exemple*) charge directement l'un
ou l'autre et active automatiquement le mode de calcul correspondant.

---

## Paramètres de calcul

| Clé | Paramètre | Valeur par défaut |
|---|---|---|
| `g` | Pesanteur (m/s²) | 9.81 |
| `nu` | Viscosité cinématique de l'eau (m²/s) | 1.1e-6 |
| `k_roughness` | Rugosité Colebrook (m) | 0.0001 (0,1 mm) |
| `dep_max` | Dépression maximale tolérée (mCE) | 3.0 |
| `v_remplissage` | Vitesse de remplissage progressive (m/s) | 0.5 |
| `mu_vidange` | Coefficient de débit de vidange | 0.7 |
| `h_seuil_interm` | Seuil de dépression déclenchant une protection intermédiaire (m) | 3.0 |
| `k_sing` | Coefficient de perte singulière de sortie | 1.0 |

---

## Méthodologie

- **Frottement** : `1/sqrt(λ) = -2·log10(k/(3,71·D) + 2,51·ν/(D·V·sqrt(λ)))`,
  résolu itérativement (`lambda_colebrook`).
- **Débit sous pente imposée** : itération couplée `V ↔ λ` (`debit_depuis_pente_colebrook`).
- **Débit avec perte singulière** : dichotomie sur
  `Δz = J_lin(V)·L + k_sing·V²/(2g)` (`debit_avec_perte_singuliere`).
- **Point haut** : `Q_dim = max(Q_rupture, Q_remplissage)` ; la ventouse est
  choisie dans `VENTOUSES_STD`, avec passage à plusieurs ventouses DN250 au-delà.
- **Point intermédiaire** : protection si la dépression `H` dépasse le seuil ;
  `Q_air = Q_total − Q_amont`.
- **Point bas** : diamètre de vidange déduit du diamètre équivalent, de la
  longueur et de la hauteur moyenne, pour des temps de vidange de 1 h, 1,5 h et 2 h.

---

## Catalogue MCP des ventouses

Le catalogue fourni à la racine (`Tableau_Recapitulatif_Debits_Ventouses_MCP.xlsx`
et `.docx`) a été **retranscrit dans `main1.py`**, section
`CATALOGUE MCP — VENTOUSES AUTOMATIQUES DOUBLE ORIFICE 3 FONCTIONS (PN 10/16)`.

### Contenu encodé

| Constante | Contenu |
|---|---|
| `CATALOGUE_MCP_SOURCE` | Métadonnées (type, PN 10/16, fonte GJS 500, pressions optimales / critiques, fichiers source) |
| `VENTOUSES_MCP` | 11 fiches DN (DN 40 → DN 300) : séries, dimensions, poids, débits évacuation/admission en Std m³/min **et** m³/h, courbe associée, notes |
| `COURBES_MCP_EVACUATION` | 9 points de pression (+0,1 → +0,9 bar) : courbes 50 / 80 / 100-150 mm, variante vanne papillon (−20 %), chandeliers |
| `COURBES_MCP_ADMISSION` | 5 points de dépression (−0,1 → −0,5 bar) : courbes 50 / 80 / 100-150 mm |
| `PURGE_PETIT_ORIFICE_MCP` | Débits de purge sous 5 → 25 bar pour 8 diamètres de tuyère (1,8 → 6,2 mm) |
| `DIMENSIONS_MCP` | 12 lignes de dimensions L, H, M, D1, D2, brides PN10-16/PN25 et poids |
| `REGLES_INGENIERIE_MCP` | Règles d'ingénierie (vanne papillon, chandeliers DN 1200/1600, équivalence corps/orifice, étanchéité 0,1 bar) |

> **Note** : ces données sont une référence catalogue. Elles n'ont **pas** modifié
> les capacités historiques de `VENTOUSES_STD` utilisées par le dimensionnement,
> afin de préserver les résultats existants. Leur utilisation pour le choix des
> ventouses pourra être activée explicitement dans une étape ultérieure.

### Fonctions d'interrogation

```python
import main1

# Fiche complète d'un DN
fiche = main1.catalogue_mcp_par_dn("DN 150")

# Capacité en Std m³/h : 'evac_opt', 'evac_crit', 'adm_opt', 'adm_crit'
q = main1.capacite_mcp_m3h("DN 150", "adm_crit")          # 5460.0

# Lecture d'une courbe (m³/min par défaut, 'm3h' possible)
q_courbe = main1.debit_mcp_courbe(0.3, "evacuation", "courbe_100_150")  # 91.0

# Tableaux pandas
catalogue = main1.catalogue_mcp_dataframe()
dimensions = main1.dimensions_mcp_dataframe()
courbes = main1.courbes_mcp_dataframe("admission")
```

### Dans l'interface

- Onglet **« Catalogue MCP »** : résumé permanent des DN, courbes, purges et règles.
- Menu **Catalogue > Consulter le catalogue MCP** : fenêtre dédiée plus large.
- Menu **Catalogue > Ouvrir le classeur / la note** : accès direct aux fichiers sources.
- Case **« Inclure le catalogue MCP dans les rapports »** : active ou non l'export
  des onglets catalogue (également dans le menu *Catalogue*).

### Présence dans les rapports

À chaque génération, le classeur Excel contient trois onglets supplémentaires :

- `Catalogue_Ventouses_MCP` (synthèse par DN),
- `Catalogue_MCP_Courbes` (évacuation + admission),
- `Catalogue_MCP_Dimensions`.

La note Word contient une section « Catalogue MCP » avec le tableau de synthèse
et les règles d'ingénierie. Ce comportement peut être désactivé avec
`exporter_excel_et_docx(..., inclure_catalogue=False)`.

---

## Sorties générées

| Fichier | Contenu |
|---|---|
| `Resultats_Hydrauliques.xlsx` | Feuilles de calcul + onglets catalogue MCP |
| `Note_Calcul_Hydraulique.docx` | Paramètres, tableaux, catalogue, graphiques |
| `figures/fig_pl_profil_piezo.png` | Profil, ligne piézométrique critique, ouvrages |
| `figures/fig_pl_ventouses.png` | Débits requis vs capacité des ventouses |
| `figures/fig_pl_vidanges.png` | Diamètres de vidange théoriques et retenus |

---

## Exécutable Windows

Un exécutable autonome est fourni dans les **GitHub Releases** du dépôt :
[`MK_Prandtl-Colebrook_V01.exe`](https://github.com/km1965/MK-Prandtl-Colebrook/releases).

Il embarque Python, toutes les dépendances, les deux classeurs d'exemple, les
fichiers sources du catalogue MCP et l'icône. Aucune installation n'est nécessaire :
téléchargez le `.exe`, double-cliquez dessus, et l'interface s'ouvre.

Les rapports sont écrits dans le dossier choisi dans l'interface (`sorties_app`
par défaut), à côté de l'exécutable. L'icône « MK » est un losange à huit côtés
avec une goutte bicolore (bleu / rouge) surmontée d'une flamme : elle rappelle
qu'une goutte d'eau peut devenir un risque pour la conduite. Le monogramme MK est
translucide, comme du verre.

Pour reconstruire l'exe depuis les sources :

```powershell
.\env\Scripts\python.exe -m pip install pyinstaller
.\env\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --windowed `
  --name MK_Prandtl-Colebrook_V01 --icon assets/MK_icon.ico `
  --add-data "Exemple_Import_Profil_En_Long.xlsx;." `
  --add-data "Exemple_Import_Cas_Simples.xlsx;." `
  --add-data "Tableau_Recapitulatif_Debits_Ventouses_MCP.xlsx;." `
  --add-data "Tableau_Recapitulatif_Debits_Ventouses_MCP.docx;." `
  --add-data "assets/MK_icon.ico;assets" main1.py
```

---

## Licence

Distribué sous licence **MIT** — voir le fichier [`LICENSE`](LICENSE).
Copyright © 2026 Mostafa Karim Soussi (km1965).

---

## Dépannage

- **`customtkinter ou Pillow n'est pas installé`** : activer `env/` puis
  `python -m pip install -r requirements.txt`.
- **`Section manquante` lors de la lecture Excel** : vérifier le nom exact des
  onglets (`Profil_en_Long` ou `Points_Hauts` / `Points_Intermediaires` / `Points_Bas`).
- **Séparateur décimal** : les champs numériques acceptent la virgule (elle est
  convertie en point automatiquement dans l'interface).
- **Aucun graphique en mode `simple`** : normal, les graphiques de synthèse sont
  générés uniquement en mode `Profil en Long`.

---

## Feuille de route

- [x] Intégration du catalogue MCP des ventouses dans le script.
- [x] Export du catalogue dans les rapports Excel et Word.
- [x] Rédaction du README.
- [x] Classeurs d'exemple d'import pour les deux modes (`pl` et `simple`).
- [x] **Barre de tâches complète dans l'interface** : menus *Fichier*, *Calcul*,
      *Catalogue*, **À propos** et **Guide utilisateur** (fenêtres modales,
      raccourcis clavier, barre de statut, accès direct au catalogue MCP et aux rapports).
- [x] **Bouton « Importer un exemple »** dans l'interface : charge directement
      `Exemple_Import_Profil_En_Long.xlsx` ou `Exemple_Import_Cas_Simples.xlsx`
      et présélectionne le mode de calcul correspondant.
- [x] **Choix interactif du dossier de sortie** (bouton *Parcourir...*).
- [x] **Ouverture automatique du rapport Word** après génération (option activable).
- [x] **Commandes Enregistrer / Enregistrer sous…** (menu *Fichier*, barre d'outils,
      Ctrl+S / Ctrl+Shift+S) pour réexporter les rapports sans recalculer.
- [x] **Icône MK** (losange 8 côtés, goutte bicolore bleu/rouge, flamme de danger,
      monogramme translucide) en PNG et ICO.
- [x] **Licence MIT** et dépôt GitHub public.
- [x] **Exécutable Windows autonome** `MK_Prandtl-Colebrook_V01.exe` (PyInstaller)
      publié dans une GitHub Release.
- [ ] Possibilité d'utiliser les capacités catalogue MCP pour le dimensionnement.

---

## Auteur

Mostafa Karim Soussi — [@km1965](https://github.com/km1965)

Dépôt public RPS 2011 : [MK-Verifications_Sismique-RPS2011](https://github.com/km1965/MK-Verifications_Sismique-RPS2011)

Identité git locale pour ce dépôt :

```powershell
git config user.name  "km1965"
git config user.email "mostafakarimsoussi1965@gmail.com"
```
