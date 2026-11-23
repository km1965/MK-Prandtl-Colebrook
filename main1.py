"""
Fichier principal : main1.py
Application complète de dimensionnement hydraulique des équipements de points hauts,
points intermédiaires (ruptures de pente H > seuil) et points bas (vidanges).

Le catalogue MCP des ventouses automatiques double orifice 3 fonctions (PN 10/16)
est encodé en données Python dans la section « CATALOGUE MCP » ci-dessous et est
exporté dans les rapports Excel/Word sous l'onglet « Catalogue_Ventouses_MCP ».

Usage :
  - Interface graphique CustomTkinter : python main1.py
  - Mode ligne de commande Profil en Long : python main1.py --mode pl --input-xlsx Exemple_Import_Profil_En_Long.xlsx --out-dir Sorties_PL
  - Mode ligne de commande Cas Simples : python main1.py --mode simple --input-xlsx Exemple_Import_Cas_Simples.xlsx --out-dir Sorties_Simples
"""

import argparse
import math
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    # Exécutable PyInstaller : BASE_DIR = dossier de l'exe (sorties),
    # RESSOURCES_DIR = dossier temporaire où sont extraites les données embarquées.
    BASE_DIR = Path(sys.executable).resolve().parent
    RESSOURCES_DIR = Path(getattr(sys, "_MEIPASS", str(BASE_DIR)))
else:
    BASE_DIR = Path(__file__).resolve().parent
    RESSOURCES_DIR = BASE_DIR

if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))


def chemin_ressource(nom: str) -> Path:
    """Résout une ressource (exemple, catalogue, icône), embarquée ou à côté de l'exe."""
    candidat = RESSOURCES_DIR / nom
    if candidat.exists():
        return candidat
    return BASE_DIR / nom

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

try:
    import customtkinter as ctk
    from tkinter import filedialog, messagebox, Menu
    from PIL import Image
    HAS_CTK = True
except ImportError:
    HAS_CTK = False


DEFAULT_PARAMS = {
    "g": 9.81,               # Pesanteur (m/s^2)
    "nu": 1.1e-6,            # Viscosité cinématique eau (m^2/s)
    "k_roughness": 0.0001,   # Rugosité Colebrook (m) -> 0.1 mm
    "dep_max": 3.0,          # Dépression maximale tolérée (mCE)
    "v_remplissage": 0.5,    # Vitesse de remplissage progressive (m/s)
    "mu_vidange": 0.7,       # Coefficient de débit de vidange
    "h_seuil_interm": 3.0,   # Seuil H (m) déclenchant la protection point intermédiaire
    "k_sing": 1.0,           # Coefficient de perte singulière de sortie dh = k * V^2 / (2g)
}

VENTOUSES_STD = [
    (50,  "DN50",  900.0),
    (80,  "DN80",  2300.0),
    (100, "DN100", 3600.0),
    (150, "DN150", 8200.0),
    (200, "DN200", 14600.0),
    (250, "DN250", 22900.0),
]


# ============================================================================
# CATALOGUE MCP — VENTOUSES AUTOMATIQUES DOUBLE ORIFICE 3 FONCTIONS (PN 10/16)
# ============================================================================
# Données retranscrites depuis les fichiers de la racine du projet :
#   - Tableau_Recapitulatif_Debits_Ventouses_MCP.xlsx (3 feuilles)
#   - Tableau_Recapitulatif_Debits_Ventouses_MCP.docx (synthèse + m³/h)
# Les débits sont exprimés en conditions « Std » (m³/min et m³/h = m³/min × 60).
# L'évacuation correspond au remplissage (surpression), l'admission à la vidange
# (dépression). Ce catalogue est une donnée de référence : il n'a pas modifié
# les hypothèses de calcul historiques de VENTOUSES_STD utilisées par le
# dimensionnement. Les fonctions ci-dessous permettent de l'interroger.

CATALOGUE_MCP_SOURCE = {
    "fabricant_type": "MCP — ventouse automatique double orifice 3 fonctions",
    "pression_nominale": "PN 10/16",
    "materiau_corps": "Fonte ductile GJS 500",
    "etancheite_min_bar": 0.1,
    "debit_optimal_bar": 0.3,
    "debit_critique_evac_bar": 0.9,
    "debit_critique_adm_bar": -0.5,
    "fichiers_source": [
        "Tableau_Recapitulatif_Debits_Ventouses_MCP.xlsx",
        "Tableau_Recapitulatif_Debits_Ventouses_MCP.docx",
    ],
}

# Synthèse par diamètre nominal (1re feuille du classeur / 1er tableau du .docx).
VENTOUSES_MCP = [
    {"dn": "DN 40", "dn_mm": 40, "serie_conduite": "60 - 250", "dimensions_lxhxm": "228 × 280 × 166",
     "poids_kg": 17.0, "evac_opt_m3min": 26.0, "evac_opt_m3h": 1560.0,
     "evac_crit_m3min": 51.0, "evac_crit_m3h": 3060.0, "adm_opt_m3min": 21.0,
     "adm_opt_m3h": 1260.0, "adm_crit_m3min": 24.5, "adm_crit_m3h": 1470.0,
     "courbe": "Courbe 50 mm", "notes": "Corps série 400 avec bride réduite DN 40/65"},
    {"dn": "DN 50", "dn_mm": 50, "serie_conduite": "60 - 250", "dimensions_lxhxm": "228 × 280 × 166",
     "poids_kg": 17.0, "evac_opt_m3min": 26.0, "evac_opt_m3h": 1560.0,
     "evac_crit_m3min": 51.0, "evac_crit_m3h": 3060.0, "adm_opt_m3min": 21.0,
     "adm_opt_m3h": 1260.0, "adm_crit_m3min": 24.5, "adm_crit_m3h": 1470.0,
     "courbe": "Courbe 50 mm", "notes": "Dimension standard de référence courbe 50 mm"},
    {"dn": "DN 60/65*", "dn_mm": 65, "serie_conduite": "60 - 250", "dimensions_lxhxm": "228 × 280 × 166",
     "poids_kg": 20.5, "evac_opt_m3min": 26.0, "evac_opt_m3h": 1560.0,
     "evac_crit_m3min": 51.0, "evac_crit_m3h": 3060.0, "adm_opt_m3min": 21.0,
     "adm_opt_m3h": 1260.0, "adm_crit_m3min": 24.5, "adm_crit_m3h": 1470.0,
     "courbe": "Courbe 50 mm", "notes": "Corps commun série 400 (H = 280 mm)"},
    {"dn": "DN 80", "dn_mm": 80, "serie_conduite": "250 - 400", "dimensions_lxhxm": "228 × 280 × 166",
     "poids_kg": 21.0, "evac_opt_m3min": 48.0, "evac_opt_m3h": 2880.0,
     "evac_crit_m3min": 96.0, "evac_crit_m3h": 5760.0, "adm_opt_m3min": 40.0,
     "adm_opt_m3h": 2400.0, "adm_crit_m3min": 49.0, "adm_crit_m3h": 2940.0,
     "courbe": "Courbe 80 mm", "notes": "Existe aussi en série 600 (H = 350 mm, 33 kg)"},
    {"dn": "DN 100", "dn_mm": 100, "serie_conduite": "250 - 400", "dimensions_lxhxm": "228 × 280 × 166",
     "poids_kg": 31.0, "evac_opt_m3min": 48.0, "evac_opt_m3h": 2880.0,
     "evac_crit_m3min": 96.0, "evac_crit_m3h": 5760.0, "adm_opt_m3min": 40.0,
     "adm_opt_m3h": 2400.0, "adm_crit_m3min": 49.0, "adm_crit_m3h": 2940.0,
     "courbe": "Courbe 80 mm (Série 400)", "notes": "Version série 400 (H = 280 mm) ; voir série 600/900"},
    {"dn": "DN 100 (Série 600/900)", "dn_mm": 100, "serie_conduite": "400 - 900",
     "dimensions_lxhxm": "253 × 420 × 265", "poids_kg": 40.0,
     "evac_opt_m3min": 91.0, "evac_opt_m3h": 5460.0, "evac_crit_m3min": 177.0,
     "evac_crit_m3h": 10620.0, "adm_opt_m3min": 74.0, "adm_opt_m3h": 4440.0,
     "adm_crit_m3min": 91.0, "adm_crit_m3h": 5460.0, "courbe": "Courbe 100/150 mm",
     "notes": "Corps grand débit (H = 420 mm) ; montage chandelier pour DN > 900"},
    {"dn": "DN 125", "dn_mm": 125, "serie_conduite": "400 - 600", "dimensions_lxhxm": "253 × 420 × 265",
     "poids_kg": 32.0, "evac_opt_m3min": 91.0, "evac_opt_m3h": 5460.0,
     "evac_crit_m3min": 177.0, "evac_crit_m3h": 10620.0, "adm_opt_m3min": 74.0,
     "adm_opt_m3h": 4440.0, "adm_crit_m3min": 91.0, "adm_crit_m3h": 5460.0,
     "courbe": "Courbe 100/150 mm", "notes": "Corps grand débit (H = 420 mm, M = 265 mm)"},
    {"dn": "DN 150", "dn_mm": 150, "serie_conduite": "400 - 600", "dimensions_lxhxm": "253 × 420 × 265",
     "poids_kg": 35.0, "evac_opt_m3min": 91.0, "evac_opt_m3h": 5460.0,
     "evac_crit_m3min": 177.0, "evac_crit_m3h": 10620.0, "adm_opt_m3min": 74.0,
     "adm_opt_m3h": 4440.0, "adm_crit_m3min": 91.0, "adm_crit_m3h": 5460.0,
     "courbe": "Courbe 100/150 mm", "notes": "Dimension standard de référence courbe 100/150 mm"},
    {"dn": "DN 200", "dn_mm": 200, "serie_conduite": "600 - 900", "dimensions_lxhxm": "253 × 420 × 265",
     "poids_kg": 50.0, "evac_opt_m3min": 91.0, "evac_opt_m3h": 5460.0,
     "evac_crit_m3min": 177.0, "evac_crit_m3h": 10620.0, "adm_opt_m3min": 74.0,
     "adm_opt_m3h": 4440.0, "adm_crit_m3min": 91.0, "adm_crit_m3h": 5460.0,
     "courbe": "Courbe 100/150 mm*", "notes": "Même corps que DN 150 (H = 420, M = 265) avec bride DN 200"},
    {"dn": "DN 250", "dn_mm": 250, "serie_conduite": "600 - 900", "dimensions_lxhxm": "253 × 420 × 265",
     "poids_kg": 60.0, "evac_opt_m3min": 91.0, "evac_opt_m3h": 5460.0,
     "evac_crit_m3min": 177.0, "evac_crit_m3h": 10620.0, "adm_opt_m3min": 74.0,
     "adm_opt_m3h": 4440.0, "adm_crit_m3min": 91.0, "adm_crit_m3h": 5460.0,
     "courbe": "Courbe 100/150 mm*", "notes": "Même corps que DN 150 (H = 420, M = 265) avec bride DN 250"},
    {"dn": "DN 300", "dn_mm": 300, "serie_conduite": "600 - 900", "dimensions_lxhxm": "253 × 420 × 265",
     "poids_kg": 75.0, "evac_opt_m3min": 91.0, "evac_opt_m3h": 5460.0,
     "evac_crit_m3min": 177.0, "evac_crit_m3h": 10620.0, "adm_opt_m3min": 74.0,
     "adm_opt_m3h": 4440.0, "adm_crit_m3min": 91.0, "adm_crit_m3h": 5460.0,
     "courbe": "Courbe 100/150 mm*", "notes": "Même corps que DN 150 (H = 420, M = 265) avec bride DN 300"},
]

# Courbes caractéristiques Grand Orifice — 2e feuille du classeur.
# Valeurs en Std m³/min ; les colonnes chandelier ne sont pas renseignées dans
# le classeur source et sont donc laissées à None ( montage décrit dans les règles).
COURBES_MCP_EVACUATION = {
    # ΔP (bar): (courbe 50 mm, courbe 80 mm, courbe 100/150 mm, 100/150 avec papillon -20 %,
    #            chandelier DN1200, chandelier DN1600)
    0.1: (14.0, 25.0, 48.0, 38.4, None, None),
    0.2: (21.0, 38.0, 71.0, 56.8, None, None),
    0.3: (26.0, 48.0, 91.0, 72.8, None, None),
    0.4: (31.0, 57.0, 108.0, 86.4, None, None),
    0.5: (35.0, 65.0, 122.0, 97.6, None, None),
    0.6: (39.0, 73.0, 136.0, 108.8, None, None),
    0.7: (43.0, 81.0, 150.0, 120.0, None, None),
    0.8: (47.0, 89.0, 164.0, 131.2, None, None),
    0.9: (51.0, 96.0, 177.0, 141.6, None, None),
}

COURBES_MCP_ADMISSION = {
    # Dépression ΔP (bar): (courbe 50 mm, courbe 80 mm, courbe 100/150 mm) en Std m³/min
    -0.1: (11.0, 20.0, 39.0),
    -0.2: (17.0, 31.0, 59.0),
    -0.3: (21.0, 40.0, 74.0),
    -0.4: (23.0, 45.0, 83.0),
    -0.5: (24.5, 49.0, 91.0),
}

COURBES_MCP_INDEX = {
    "evacuation": {"courbe_50": 0, "courbe_80": 1, "courbe_100_150": 2,
                   "courbe_100_150_papillon": 3, "chandelier_dn1200": 4,
                   "chandelier_dn1600": 5},
    "admission": {"courbe_50": 0, "courbe_80": 1, "courbe_100_150": 2},
}

# Petit orifice (purge en service / dégazage sous pression), Std m³/min — 3e feuille.
DIAMETRES_PURGE_MCP = [1.8, 2.0, 2.2, 2.5, 2.8, 3.5, 4.0, 6.2]
PURGE_PETIT_ORIFICE_MCP = {
    5.0:  [0.11, 0.15, 0.19, 0.23, 0.30, 0.46, 0.62, 1.50],
    10.0: [0.22, 0.30, 0.38, 0.46, 0.60, 0.92, 1.24, 3.00],
    15.0: [0.33, 0.45, 0.57, 0.69, 0.90, 1.38, 1.86, 4.50],
    20.0: [0.44, 0.60, 0.76, 0.92, 1.20, 1.84, 2.48, 6.00],
    25.0: [0.55, 0.75, 0.95, 1.15, 1.50, 2.30, 3.10, None],  # Ø 6.2 mm hors plage (> 20 bar)
}

# Caractéristiques dimensionnelles et pondérales — 3e feuille du classeur.
# (série, DN, L, H, M, D1, D2, D bride PN10-16, D bride PN25, poids kg)
DIMENSIONS_MCP = [
    ("400 (60-250)", "50", 228, 280, 166, 99, 125, 165, 165, 21),
    ("400 (60-250)", "65* (ou 40/65)", 228, 280, 166, 118, 145, 185, 185, 22),
    ("400 (60-250)", "80", 228, 280, 166, 132, 160, 200, 200, 25),
    ("400 (60-250)", "100", 228, 280, 166, 156, 180, 220, 235, 27),
    ("600 (250-400)", "80", 237, 350, 240, 132, 160, 200, 200, 33),
    ("600 (250-400)", "100", 237, 350, 240, 156, 180, 220, 235, 34),
    ("900 (400-900)", "100", 253, 420, 265, 156, 180, 220, 235, 40),
    ("900 (400-900)", "125", 253, 420, 265, 184, 210, 250, 270, 42),
    ("900 (400-900)", "150", 253, 420, 265, 211, 240, 285, 300, 47),
    ("900 (400-900)", "200", 253, 420, 265, 266, 295, 340, 360, 50),
    ("900 (400-900)", "250", 253, 420, 265, 319, 350, 400, 425, 60),
    ("900 (400-900)", "300", 253, 420, 265, 370, 400, 455, 485, 75),
]

REGLES_INGENIERIE_MCP = [
    "Réduction vanne papillon : l'installation d'une vanne d'isolement papillon en amont "
    "réduit le débit d'air d'environ 20 % (ex. : DN 100/150 passe de ~91 à ~73 Std m³/min à 0,3 bar).",
    "Montages en chandelier (DN conduite > 900 mm) : DN 1200 mm → 1 sortie DN 150 + 2 ventouses DN 100 "
    "(série 900) ; DN 1600 mm → 1 sortie DN 150 + 3 ventouses DN 100 (série 900).",
    "Équivalence corps/orifice : les modèles DN 200, DN 250 et DN 300 partagent la même géométrie de corps "
    "interne (L = 253, H = 420, M = 265 mm) que le DN 150, seule la bride de raccordement diffère.",
    "Pression minimale d'étanchéité du flotteur : 0,1 bar.",
]


def lambda_colebrook(k_val: float, d_val: float, v_val: float, nu_val: float) -> float:
    """Résout 1/sqrt(lambda) = -2*log10(k/(3.71*D) + 2.51*nu/(D*V*sqrt(lambda)))."""
    v_eff = max(abs(v_val), 1e-4)
    inv_sqrt_l = -2.0 * math.log10(
        k_val / (3.71 * d_val) + (2.51 * nu_val) / (d_val * v_eff * 0.02)
    )
    for _ in range(120):
        inv_sqrt_l = max(inv_sqrt_l, 1.0)
        sqrt_l = 1.0 / inv_sqrt_l
        arg = k_val / (3.71 * d_val) + (2.51 * nu_val) / (d_val * v_eff * sqrt_l)
        inv_sqrt_l = -2.0 * math.log10(max(arg, 1e-15))
    return 1.0 / (inv_sqrt_l ** 2)


def debit_depuis_pente_colebrook(j_val: float, d_m: float, params: dict):
    if j_val <= 0 or d_m <= 0:
        return 0.0, 0.0, 0.0
    v = math.sqrt(2.0 * params["g"] * d_m * j_val / 0.02)
    lam = 0.02
    for _ in range(120):
        lam = lambda_colebrook(params["k_roughness"], d_m, v, params["nu"])
        v = math.sqrt(max(2.0 * params["g"] * d_m * j_val / lam, 0.0))
    return (v * math.pi * (d_m ** 2) / 4.0) * 3600.0, v, lam


def debit_avec_perte_singuliere(dz_eff: float, longueur_m: float, d_m: float, params: dict):
    """Résout dz_eff = J_lin(V)*L + k_sing*V^2/(2g) par dichotomie."""
    if dz_eff <= 0 or longueur_m <= 0 or d_m <= 0:
        return 0.0, 0.0, 0.0
    g = params["g"]
    k_s = params.get("k_sing", 1.0)
    v_low, v_high = 0.0, math.sqrt(2.0 * g * dz_eff / max(k_s, 1e-6)) + 5.0
    for _ in range(100):
        v_mid = 0.5 * (v_low + v_high)
        lam = lambda_colebrook(params["k_roughness"], d_m, v_mid, params["nu"])
        j_lin = lam * (v_mid ** 2) / (2.0 * g * d_m)
        if j_lin * longueur_m + k_s * (v_mid ** 2) / (2.0 * g) - dz_eff < 0:
            v_low = v_mid
        else:
            v_high = v_mid
    v_sol = 0.5 * (v_low + v_high)
    lam_sol = lambda_colebrook(params["k_roughness"], d_m, v_sol, params["nu"])
    j_sol = lam_sol * (v_sol ** 2) / (2.0 * g * d_m)
    return (v_sol * math.pi * (d_m ** 2) / 4.0) * 3600.0, j_sol, v_sol


def choisir_ventouse(q_requis_m3h: float):
    for dn, _, cap in VENTOUSES_STD:
        if cap >= q_requis_m3h:
            return f"1V{dn}", cap
    top_dn, _, top_cap = VENTOUSES_STD[-1]
    nb = math.ceil(q_requis_m3h / top_cap)
    return f"{nb}V{top_dn}", nb * top_cap


# --- Fonctions d'interrogation du catalogue MCP -----------------------------

def catalogue_mcp_par_dn(dn: str):
    """Retourne la fiche catalogue MCP correspondant à un DN ('DN 150', 150, ...)."""
    dn_str = str(dn).strip().upper().replace(" ", "")
    dn_num = dn_str[2:] if dn_str.startswith("DN") else dn_str
    for fic in VENTOUSES_MCP:
        if fic["dn"].upper().replace(" ", "") == dn_str or str(fic["dn_mm"]) == dn_num:
            return dict(fic)
    return None


def capacite_mcp_m3h(dn: str, regime: str = "adm_crit"):
    """Capacité catalogue (Std m³/h) d'un DN pour un régime :
    'evac_opt', 'evac_crit', 'adm_opt' ou 'adm_crit' (défaut : admission critique)."""
    fic = catalogue_mcp_par_dn(dn)
    if fic is None:
        return None
    cle = {
        "evac_opt": "evac_opt_m3h", "evac_crit": "evac_crit_m3h",
        "adm_opt": "adm_opt_m3h", "adm_crit": "adm_crit_m3h",
    }.get(regime)
    if cle is None:
        raise ValueError(f"Régime inconnu : {regime!r}")
    return fic[cle]


def debit_mcp_courbe(delta_p_bar: float, regime: str = "evacuation",
                     courbe: str = "courbe_100_150", unite: str = "m3min"):
    """Débit d'une courbe catalogue MCP à une pression donnée.

    regime : 'evacuation' (ΔP > 0) ou 'admission' (ΔP < 0).
    courbe : clé de COURBES_MCP_INDEX[regime].
    unite  : 'm3min' (Std m³/min) ou 'm3h' (Std m³/h).
    Retourne None si le point n'est pas tabulé dans le catalogue.
    """
    tableau = COURBES_MCP_EVACUATION if regime == "evacuation" else COURBES_MCP_ADMISSION
    idx = COURBES_MCP_INDEX.get(regime, {}).get(courbe)
    if idx is None:
        raise ValueError(f"Courbe inconnue : {courbe!r} pour le régime {regime!r}")
    values = tableau.get(round(float(delta_p_bar), 1))
    if values is None or values[idx] is None:
        return None
    return values[idx] * 60.0 if unite == "m3h" else values[idx]


def catalogue_mcp_dataframe():
    """Synthèse MCP sous forme de DataFrame (une ligne par DN)."""
    return pd.DataFrame(VENTOUSES_MCP)


def dimensions_mcp_dataframe():
    """Tableau des dimensions/poids MCP sous forme de DataFrame."""
    return pd.DataFrame(DIMENSIONS_MCP, columns=[
        "serie_conduite", "dn_ventouse", "L_mm", "H_mm", "M_mm", "D1_mm",
        "D2_mm", "D_bride_PN10_16_mm", "D_bride_PN25_mm", "poids_PN10_16_kg",
    ])


def courbes_mcp_dataframe(regime: str = "evacuation"):
    """Courbes Grand Orifice MCP sous forme de DataFrame (une ligne par ΔP)."""
    if regime == "evacuation":
        tableau = COURBES_MCP_EVACUATION
        colonnes = ["delta_p_bar", "courbe_50_m3min", "courbe_80_m3min",
                    "courbe_100_150_m3min", "courbe_100_150_papillon_m3min",
                    "chandelier_dn1200_m3min", "chandelier_dn1600_m3min"]
    elif regime == "admission":
        tableau = COURBES_MCP_ADMISSION
        colonnes = ["delta_p_bar", "courbe_50_m3min", "courbe_80_m3min",
                    "courbe_100_150_m3min"]
    else:
        raise ValueError("regime doit valoir 'evacuation' ou 'admission'")
    return pd.DataFrame([[dp, *values] for dp, values in sorted(tableau.items())], columns=colonnes)


def _assembler_courbes_mcp():
    """Regroupe les courbes évacuation et admission en un seul tableau exportable."""
    evac = courbes_mcp_dataframe("evacuation")
    adm = courbes_mcp_dataframe("admission")
    evac.insert(0, "regime", "Évacuation / remplissage")
    adm.insert(0, "regime", "Admission / vidange")
    return pd.concat([evac, adm], ignore_index=True)


def calculer_mode_simple(df_ph: pd.DataFrame, df_pi: pd.DataFrame, df_pb: pd.DataFrame, params: dict):
    dep_max = params["dep_max"]
    v_rempl = params["v_remplissage"]
    h_seuil = params["h_seuil_interm"]

    res_ph = []
    for _, r in df_ph.iterrows():
        d_m = float(r["dn_mm"]) / 1000.0
        zh_prime = float(r["z_h"]) - dep_max
        xh = float(r["x_h"])
        if "ig1_force_pct" in r and pd.notna(r["ig1_force_pct"]) and float(r["ig1_force_pct"]) > 0:
            ig1 = float(r["ig1_force_pct"]) / 100.0
            q1, _, _ = debit_depuis_pente_colebrook(ig1, d_m, params)
        else:
            q1, ig1, _ = debit_avec_perte_singuliere(zh_prime - float(r["z_am"]), abs(xh - float(r["x_am"])), d_m, params)
        if "ig2_force_pct" in r and pd.notna(r["ig2_force_pct"]) and float(r["ig2_force_pct"]) > 0:
            ig2 = float(r["ig2_force_pct"]) / 100.0
            q2, _, _ = debit_depuis_pente_colebrook(ig2, d_m, params)
        else:
            q2, ig2, _ = debit_avec_perte_singuliere(zh_prime - float(r["z_av"]), abs(float(r["x_av"]) - xh), d_m, params)

        q_rupt = max(q1, q2)
        q_rempl = v_rempl * math.pi * (d_m ** 2) / 4.0 * 3600.0
        q_dim = max(q_rupt, q_rempl)
        equip, cap = choisir_ventouse(q_dim)
        res_ph.append({
            "piquet_h": r["piquet_h"], "x_h_m": round(xh, 2), "z_h_m": round(float(r["z_h"]), 3),
            "dn_mm": int(r["dn_mm"]), "pn_bar": int(r.get("pn_bar", 16)),
            "ig1_pct": round(ig1 * 100, 3), "ig2_pct": round(ig2 * 100, 3),
            "q_rupture_m3h": round(q_rupt, 2), "q_rempl_m3h": round(q_rempl, 2),
            "q_dim_m3h": round(q_dim, 2), "equipement": equip, "capacite_m3h": cap,
        })

    res_pi = []
    for _, r in df_pi.iterrows():
        xi, zi = float(r["x_i"]), float(r["z_i"])
        xh, zh = float(r["x_h"]), float(r["z_h"])
        xb, zb = float(r["x_b"]), float(r["z_b"])
        d_m = float(r["dn_mm"]) / 1000.0
        zh_prime = zh - dep_max
        z_piezo = zh_prime + (zb - zh_prime) * ((xi - xh) / (xb - xh)) if abs(xb - xh) > 1e-6 else zh_prime
        h_dep = zi - z_piezo
        if h_dep > h_seuil:
            zi_prime = zi - dep_max
            ig1_i = max((zh_prime - zi_prime) / abs(xi - xh), 0.0) if abs(xi - xh) > 0 else 0.0
            q_a, _, _ = debit_depuis_pente_colebrook(ig1_i, d_m, params)
            q_tot, ig2_i, _ = debit_avec_perte_singuliere(zi_prime - zb, abs(xb - xi), d_m, params)
            q_i = max(q_tot - q_a, 0.0)
            equip_i, cap_i = choisir_ventouse(q_i) if q_i > 0 else ("Aucune", 0.0)
            besoin = "OUI (H > seuil)"
        else:
            ig1_i = ig2_i = q_a = q_tot = q_i = cap_i = 0.0
            equip_i, besoin = "Aucune", "NON (H <= seuil)"
        res_pi.append({
            "piquet_i": r["piquet_i"], "x_i_m": round(xi, 2), "z_i_m": round(zi, 3),
            "depression_h_m": round(h_dep, 2), "protection_requise": besoin,
            "qa_m3h": round(q_a, 2), "q_total_m3h": round(q_tot, 2),
            "qi_air_m3h": round(q_i, 2), "equipement": equip_i, "capacite_m3h": cap_i,
        })

    res_pb = []
    for idx, r in df_pb.iterrows():
        long_m = abs(float(r["x_av"]) - float(r["x_am"]))
        h_moy = max(float(r["h_moy"]), 1e-3)
        d_eq = float(r["d_equiv_mm"])
        base = (d_eq / 1000.0) * math.sqrt(long_m / math.sqrt(h_moy))
        dn_vid = int(r.get("dn_vid_mm", 100))
        res_pb.append({
            "n": idx + 1, "piquet_pb": r["piquet_pb"],
            "x_pb_m": round(float(r["x_pb"]), 2), "z_pb_m": round(float(r["z_pb"]), 3),
            "longueur_m": round(long_m, 2), "h_moy_m": round(h_moy, 2), "d_equiv_mm": round(d_eq, 1),
            "d_1h_mm": round(9.5 * base, 1), "d_1_5h_mm": round(7.7 * base, 1), "d_2h_mm": round(6.7 * base, 1),
            "dn_vidange_mm": dn_vid, "pn_bar": int(r.get("pn_bar", 16)),
        })

    return pd.DataFrame(res_ph), pd.DataFrame(res_pi), pd.DataFrame(res_pb)


def calculer_profil_en_long(df_pl: pd.DataFrame, params: dict):
    df = df_pl.sort_values("x_m").reset_index(drop=True).copy()
    n = len(df)
    dep_max, v_rempl, h_seuil = params["dep_max"], params["v_remplissage"], params["h_seuil_interm"]

    types = []
    for i in range(n):
        zc = df.loc[i, "z_m"]
        if i == 0:
            types.append("POINT_HAUT" if zc > df.loc[i + 1, "z_m"] else "POINT_BAS")
        elif i == n - 1:
            types.append("POINT_HAUT" if zc > df.loc[i - 1, "z_m"] else "POINT_BAS")
        else:
            zp, zn = df.loc[i - 1, "z_m"], df.loc[i + 1, "z_m"]
            if zc >= zp and zc > zn:
                types.append("POINT_HAUT")
            elif zc <= zp and zc < zn:
                types.append("POINT_BAS")
            else:
                types.append("SOMMET_COURANT")
    df["type_sommet"] = types

    idx_ph = df.index[df["type_sommet"] == "POINT_HAUT"].tolist()
    idx_pb = df.index[df["type_sommet"] == "POINT_BAS"].tolist()

    res_ph, ph_infos = [], []
    for ih in idx_ph:
        xh, zh = float(df.loc[ih, "x_m"]), float(df.loc[ih, "z_m"])
        d_m = float(df.loc[ih, "dn_mm"]) / 1000.0
        zh_prime = zh - dep_max
        ib_am = [ib for ib in idx_pb if ib < ih][-1] if [ib for ib in idx_pb if ib < ih] else None
        ib_av = [ib for ib in idx_pb if ib > ih][0] if [ib for ib in idx_pb if ib > ih] else None
        q1, ig1, _ = debit_avec_perte_singuliere(zh_prime - float(df.loc[ib_am, "z_m"]), xh - float(df.loc[ib_am, "x_m"]), d_m, params) if ib_am is not None else (0.0, 0.0, 0.0)
        q2, ig2, _ = debit_avec_perte_singuliere(zh_prime - float(df.loc[ib_av, "z_m"]), float(df.loc[ib_av, "x_m"]) - xh, d_m, params) if ib_av is not None else (0.0, 0.0, 0.0)
        q_rupt = max(q1, q2)
        q_rempl = v_rempl * math.pi * (d_m ** 2) / 4.0 * 3600.0
        q_dim = max(q_rupt, q_rempl)
        equip, cap = choisir_ventouse(q_dim)
        res_ph.append({
            "piquet_h": df.loc[ih, "piquet"], "x_h_m": round(xh, 2), "z_h_m": round(zh, 3),
            "dn_mm": int(df.loc[ih, "dn_mm"]), "pn_bar": int(df.loc[ih, "pn_bar"]),
            "ig1_pct": round(ig1 * 100, 3), "ig2_pct": round(ig2 * 100, 3),
            "q_rupture_m3h": round(q_rupt, 2), "q_rempl_m3h": round(q_rempl, 2),
            "q_dim_m3h": round(q_dim, 2), "equipement": equip, "capacite_m3h": cap,
        })
        ph_infos.append({"ih": ih, "ib_am": ib_am, "ib_av": ib_av, "xh": xh, "zh_prime": zh_prime})

    res_pi = []
    piezo_min = df["z_m"] - dep_max
    for info in ph_infos:
        ih, xh, zh_prime = info["ih"], info["xh"], info["zh_prime"]
        for side, ib in [("AVAL", info["ib_av"]), ("AMONT", info["ib_am"])]:
            if ib is None:
                continue
            rng = range(ih + 1, ib) if side == "AVAL" else range(ib + 1, ih)
            xb, zb = float(df.loc[ib, "x_m"]), float(df.loc[ib, "z_m"])
            for k in rng:
                xk, zk = float(df.loc[k, "x_m"]), float(df.loc[k, "z_m"])
                d_m = float(df.loc[k, "dn_mm"]) / 1000.0
                z_line = zh_prime + (zb - zh_prime) * (abs(xk - xh) / abs(xb - xh))
                piezo_min.loc[k] = min(piezo_min.loc[k], z_line)
                h_dep = zk - z_line
                if h_dep > h_seuil:
                    zk_prime = zk - dep_max
                    ig1_i = max((zh_prime - zk_prime) / abs(xk - xh), 0.0)
                    q_a, _, _ = debit_depuis_pente_colebrook(ig1_i, d_m, params)
                    q_tot, ig2_i, _ = debit_avec_perte_singuliere(zk_prime - zb, abs(xb - xk), d_m, params)
                    q_i = max(q_tot - q_a, 0.0)
                    equip_i, cap_i = choisir_ventouse(q_i) if q_i > 0 else ("1V100", 3600.0)
                    res_pi.append({
                        "piquet_i": df.loc[k, "piquet"], "versant": f"{df.loc[ih, 'piquet']} -> {df.loc[ib, 'piquet']}",
                        "x_i_m": round(xk, 2), "z_i_m": round(zk, 3), "depression_h_m": round(h_dep, 2),
                        "qa_m3h": round(q_a, 2), "q_total_m3h": round(q_tot, 2),
                        "qi_air_m3h": round(q_i, 2), "equipement": equip_i, "capacite_m3h": cap_i,
                    })
    df["z_piezo_min_m"] = piezo_min

    res_pb = []
    for cnt, ib in enumerate(idx_pb, start=1):
        zb = float(df.loc[ib, "z_m"])
        i_start = [ih for ih in idx_ph if ih <= ib][-1] if [ih for ih in idx_ph if ih <= ib] else 0
        i_end = [ih for ih in idx_ph if ih >= ib][0] if [ih for ih in idx_ph if ih >= ib] else (n - 1)
        if i_end <= i_start:
            continue
        sub = df.loc[i_start:i_end]
        xv, zv, dv = sub["x_m"].to_numpy(float), sub["z_m"].to_numpy(float), sub["dn_mm"].to_numpy(float)
        long_m = float(xv[-1] - xv[0])
        dx = np.diff(xv)
        h_moy = max(float(np.sum(np.maximum(0.5 * (zv[:-1] + zv[1:]) - zb, 0.0) * dx) / long_m), 0.1)
        d_eq = float(np.sum(0.5 * (dv[:-1] + dv[1:]) * dx) / long_m)
        base = (d_eq / 1000.0) * math.sqrt(long_m / math.sqrt(h_moy))
        d_1_5h = 7.7 * base
        dn_ret = next((dn for dn in [100, 150, 200, 250, 300] if dn >= d_1_5h), 300)
        res_pb.append({
            "n": cnt, "piquet_pb": df.loc[ib, "piquet"], "x_pb_m": round(float(df.loc[ib, "x_m"]), 2),
            "z_pb_m": round(zb, 3), "longueur_m": round(long_m, 2), "h_moy_m": round(h_moy, 2),
            "d_equiv_mm": round(d_eq, 1), "d_1h_mm": round(9.5 * base, 1),
            "d_1_5h_mm": round(d_1_5h, 1), "d_2h_mm": round(6.7 * base, 1),
            "dn_vidange_retenu_mm": dn_ret, "pn_bar": int(df.loc[ib, "pn_bar"]),
        })
    return df, pd.DataFrame(res_ph), pd.DataFrame(res_pi), pd.DataFrame(res_pb)


def generer_figures_pl(df_out: pd.DataFrame, df_ph: pd.DataFrame, df_pi: pd.DataFrame, df_pb: pd.DataFrame, fig_dir: Path):
    fig_dir.mkdir(parents=True, exist_ok=True)
    p1 = fig_dir / "fig_pl_profil_piezo.png"
    fig, ax = plt.subplots(figsize=(10.5, 4.5), dpi=180)
    ax.plot(df_out["x_m"], df_out["z_m"], color="#0F172A", linewidth=1.8, label="Profil en long Z (m)")
    ax.plot(df_out["x_m"], df_out["z_piezo_min_m"], color="#DC2626", linestyle="--", linewidth=1.3, label="Ligne piézométrique critique")
    ph_pts = df_out[df_out["type_sommet"] == "POINT_HAUT"]
    pb_pts = df_out[df_out["type_sommet"] == "POINT_BAS"]
    ax.scatter(ph_pts["x_m"], ph_pts["z_m"], marker="^", s=60, color="#1D4ED8", label="Points hauts")
    ax.scatter(pb_pts["x_m"], pb_pts["z_m"], marker="v", s=60, color="#D97706", label="Points bas")
    if not df_pi.empty:
        ax.scatter(df_pi["x_i_m"], df_pi["z_i_m"], marker="D", s=55, color="#0D9488", label="Points intermédiaires équipés")
    ax.set_xlabel("Abscisse X (m)")
    ax.set_ylabel("Cote Z (m)")
    ax.set_title("Profil en long, ligne piézométrique et ouvrages")
    ax.grid(True, linestyle=":", alpha=0.5)
    ax.legend(fontsize=8)
    plt.tight_layout()
    plt.savefig(p1)
    plt.close(fig)

    p2 = fig_dir / "fig_pl_ventouses.png"
    fig, ax = plt.subplots(figsize=(10.5, 4.0), dpi=180)
    lbls = list(df_ph["piquet_h"]) + ([f"{x} (PI)" for x in df_pi["piquet_i"]] if not df_pi.empty else [])
    q_r = list(df_ph["q_dim_m3h"]) + (list(df_pi["qi_air_m3h"]) if not df_pi.empty else [])
    c_r = list(df_ph["capacite_m3h"]) + (list(df_pi["capacite_m3h"]) if not df_pi.empty else [])
    xi = np.arange(len(lbls))
    ax.bar(xi - 0.18, q_r, 0.36, label="Débit air requis (m³/h)", color="#1E40AF")
    ax.bar(xi + 0.18, c_r, 0.36, label="Capacité ventouse (m³/h)", color="#0D9488")
    ax.set_xticks(xi)
    ax.set_xticklabels(lbls, rotation=45, ha="right", fontsize=8)
    ax.set_ylabel("Débit (m³/h)")
    ax.set_title("Débits d'air requis vs capacité ventouse")
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    ax.legend(fontsize=8)
    plt.tight_layout()
    plt.savefig(p2)
    plt.close(fig)

    p3 = fig_dir / "fig_pl_vidanges.png"
    fig, ax = plt.subplots(figsize=(10.5, 4.0), dpi=180)
    xb = np.arange(len(df_pb))
    ax.plot(xb, df_pb["d_1h_mm"], marker="o", color="#DC2626", label="d théorique T = 1,0 h (mm)")
    ax.plot(xb, df_pb["d_1_5h_mm"], marker="s", color="#2563EB", label="d théorique T = 1,5 h (mm)")
    ax.plot(xb, df_pb["d_2h_mm"], marker="^", color="#059669", label="d théorique T = 2,0 h (mm)")
    ax.bar(xb, df_pb["dn_vidange_retenu_mm"], width=0.35, alpha=0.25, color="#0F172A", label="DN vidange retenu (mm)")
    ax.set_xticks(xb)
    ax.set_xticklabels(df_pb["piquet_pb"], rotation=45, ha="right", fontsize=8)
    ax.set_ylabel("Diamètre (mm)")
    ax.set_title("Diamètres de vidange aux points bas")
    ax.grid(axis="y", linestyle=":", alpha=0.5)
    ax.legend(fontsize=8)
    plt.tight_layout()
    plt.savefig(p3)
    plt.close(fig)
    return [p1, p2, p3]


def exporter_excel_et_docx(sheets_dict: dict, figs: list, out_dir: Path, titre: str,
                           params: dict, inclure_catalogue: bool = True):
    out_dir.mkdir(parents=True, exist_ok=True)
    xlsx_path = out_dir / "Resultats_Hydrauliques.xlsx"
    docx_path = out_dir / "Note_Calcul_Hydraulique.docx"

    wb = openpyxl.Workbook()
    ws0 = wb.active
    first = True
    navy = PatternFill("solid", fgColor="1E293B")
    zebra = PatternFill("solid", fgColor="F8FAFC")
    hfont = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    bfont = Font(name="Calibri", size=10, bold=False, color="0F172A")
    border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    def _ecrire_feuille(ws, df):
        if df is None or df.empty:
            ws["A1"] = "Aucune donnée dans cette feuille."
            return
        for c_i, col in enumerate(df.columns, start=1):
            c = ws.cell(1, c_i, col)
            c.fill = navy
            c.font = hfont
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = border
        for r_i, (_, row) in enumerate(df.iterrows(), start=2):
            for c_i, val in enumerate(row.values, start=1):
                c = ws.cell(r_i, c_i, None if pd.isna(val) else val)
                c.font = bfont
                c.border = border
                if r_i % 2 == 1:
                    c.fill = zebra
        ws.freeze_panes = "A2"
        for col in ws.columns:
            max_w = max((len(str(c.value or "")) for c in col), default=12)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_w + 3, 13), 30)

    for name, df in sheets_dict.items():
        ws = ws0 if first else wb.create_sheet()
        first = False
        ws.title = name[:31]
        _ecrire_feuille(ws, df)

    # Catalogue MCP : ajout systématique en fin de classeur (données de référence).
    if inclure_catalogue:
        for name, df in [
            ("Catalogue_Ventouses_MCP", catalogue_mcp_dataframe()),
            ("Catalogue_MCP_Dimensions", dimensions_mcp_dataframe()),
            ("Catalogue_MCP_Courbes", _assembler_courbes_mcp()),
        ]:
            ws = ws0 if first else wb.create_sheet()
            first = False
            ws.title = name[:31]
            _ecrire_feuille(ws, df)

    wb.save(xlsx_path)

    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10.5)
    doc.styles["Normal"].font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)

    h1 = doc.add_heading(titre, level=1)
    for r in h1.runs:
        r.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)

    doc.add_heading("Paramètres de calcul retenus", level=2)
    doc.add_paragraph(
        f"Dépression max : {params['dep_max']} mCE | g = {params['g']} m/s² | "
        f"nu = {params['nu']} m²/s | Rugosité k = {params['k_roughness']} m | "
        f"Vitesse remplissage = {params['v_remplissage']} m/s | mu vidange = {params['mu_vidange']} | "
        f"Seuil point intermédiaire H = {params['h_seuil_interm']} m."
    )

    for name, df in sheets_dict.items():
        if df is None or df.empty or name == "Profil_Detail":
            continue
        doc.add_heading(f"Tableau — {name}", level=2)
        cols_show = list(df.columns[:8])
        tbl = doc.add_table(rows=1, cols=len(cols_show))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, col_n in enumerate(cols_show):
            tbl.rows[0].cells[i].text = str(col_n)
        for _, r in df.head(25).iterrows():
            rc = tbl.add_row().cells
            for i, col_n in enumerate(cols_show):
                val = r[col_n]
                rc[i].text = "" if pd.isna(val) else str(val)

    if inclure_catalogue:
        doc.add_heading("Catalogue MCP — Ventouses automatiques double orifice (PN 10/16)", level=2)
        doc.add_paragraph(
            f"Type : {CATALOGUE_MCP_SOURCE['fabricant_type']} | "
            f"Corps : {CATALOGUE_MCP_SOURCE['materiau_corps']} | "
            f"Étanchéité min. : {CATALOGUE_MCP_SOURCE['etancheite_min_bar']} bar | "
            f"Débit optimal : ΔP = {CATALOGUE_MCP_SOURCE['debit_optimal_bar']} bar | "
            f"Débit critique : +{CATALOGUE_MCP_SOURCE['debit_critique_evac_bar']} bar (évac.) / "
            f"{CATALOGUE_MCP_SOURCE['debit_critique_adm_bar']} bar (adm.)."
        )
        cols_catalogue = ["dn", "serie_conduite", "evac_opt_m3h", "evac_crit_m3h",
                          "adm_opt_m3h", "adm_crit_m3h", "poids_kg"]
        tbl_cat = doc.add_table(rows=1, cols=len(cols_catalogue))
        tbl_cat.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, col_n in enumerate(cols_catalogue):
            tbl_cat.rows[0].cells[i].text = str(col_n)
        for fic in VENTOUSES_MCP:
            rc = tbl_cat.add_row().cells
            for i, col_n in enumerate(cols_catalogue):
                rc[i].text = str(fic[col_n])
        doc.add_paragraph(
            "Les courbes détaillées (évacuation, admission, vanne papillon, chandeliers), "
            "les débits du petit orifice de purge et les dimensions sont fournis dans le classeur "
            "Excel, onglets « Catalogue_MCP_Courbes » et « Catalogue_MCP_Dimensions »."
        )
        for regle in REGLES_INGENIERIE_MCP:
            doc.add_paragraph(regle, style="List Bullet")

    doc.add_heading("Graphiques de synthèse", level=2)
    for fp in figs:
        if Path(fp).exists():
            doc.add_picture(str(fp), width=Inches(6.2))
            cp = doc.add_paragraph(Path(fp).stem)
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cp.runs[0].italic = True
            cp.runs[0].font.size = Pt(9)

    doc.save(docx_path)
    return xlsx_path, docx_path


def _ouvrir_chemin_systeme(target: Path):
    path_str = str(target.resolve())
    sys_name = platform.system()
    if sys_name == "Windows":
        os.startfile(path_str)  # type: ignore[attr-defined]
    elif sys_name == "Darwin":
        subprocess.run(["open", path_str], check=False)
    else:
        subprocess.run(["xdg-open", path_str], check=False)


VERSION_APP = "1.0.0"

GUIDE_UTILISATEUR = """GUIDE UTILISATEUR — Dimensionnement hydraulique des points hauts, intermédiaires et bas

1. DÉMARRAGE
   Lancez l'application avec « python main1.py » (environnement virtuel env/ activé).
   La barre de menus en haut donne accès à Fichier, Calcul, Catalogue et Aide.
   La barre d'outils ci-dessous reprend les actions les plus fréquentes.

2. CHOISIR LE MODE DE CALCUL
   • Profil en Long (PL) : à partir d'une liste de piquets (x_m, z_m), l'application
     détecte automatiquement les points hauts, les points bas et les points
     intermédiaires à protéger. Feuille Excel attendue : « Profil_en_Long ».
   • Cas Simples : vous décrivez directement 3 jeux de données dans 3 feuilles :
     « Points_Hauts », « Points_Intermediaires » et « Points_Bas ».
   Le mode peut aussi être changé depuis le menu Calcul.

3. IMPORTER UN FICHIER
   • Bouton « Sélectionner fichier .xlsx » ou menu Fichier > Sélectionner un fichier.
   • Bouton « Importer un exemple » : choisissez « Profil en Long » ou « Cas Simples ».
     Le classeur d'exemple est chargé ET le mode de calcul correspondant est
     présélectionné automatiquement. Les exemples sont livrés à la racine du projet.
   • Raccourci : Ctrl+O.

4. DOSSIER DE SORTIE
   Le champ « Dossier de sortie » contient « sorties_app » par défaut (dossier relatif
   créé à côté de main1.py). Cliquez sur « Parcourir... » ou utilisez le menu
   Fichier > Choisir le dossier de sortie pour sélectionner n'importe quel dossier.
   Vous pouvez aussi saisir directement un chemin absolu dans le champ.

5. HYPOTHÈSES HYDRAULIQUES
   Ajustez si besoin : dépression max, rugosité k, viscosité nu, pesanteur g,
   vitesse de remplissage, coefficient de vidange, seuil des points intermédiaires
   et coefficient de perte singulière. Les valeurs par défaut sont documentées
   dans le README. Le séparateur décimal « , » est accepté.

6. LANCER LE CALCUL
   Cliquez sur « Lancer le calcul & générer rapports » (ou F5, ou menu Calcul).
   Le journal « Résultats & Journal » affiche les tableaux détaillés et les chemins
   des fichiers générés. Si l'option est activée, la note Word s'ouvre
   automatiquement à la fin du calcul.

7. RAPPORTS ET GRAPHIQUES
   • « Enregistrer » (Ctrl+S) : réécrit le classeur Excel et la note Word dans le
     dossier de sortie courant, sans relancer le calcul. Si aucun calcul n'a encore
     été fait, le calcul est lancé automatiquement.
   • « Enregistrer sous... » (Ctrl+Shift+S) : choisit un autre dossier, y copie les
     graphiques si besoin, puis enregistre les deux rapports. Le nouveau dossier
     devient le dossier de sortie courant.
   • « Ouvrir la note Word (.docx) » : note de calcul avec tableaux, catalogue MCP
     et graphiques.
   • « Ouvrir le dossier de sortie » : accès au classeur Excel et aux figures.
   • Onglet « Aperçu Graphique » : profil piézométrique, débits d'air et vidanges
     (uniquement en mode Profil en Long).

8. CATALOGUE MCP
   Le menu Catalogue permet de consulter un résumé des ventouses DN 40 à DN 300,
   d'ouvrir les fichiers sources (.xlsx / .docx) et d'activer ou désactiver
   l'ajout du catalogue dans les rapports. Les données sont aussi accessibles
   depuis Python via les fonctions catalogue_mcp_par_dn(), capacite_mcp_m3h(), etc.

9. RACCOURCIS
   Ctrl+O        : sélectionner un fichier d'entrée
   F5            : lancer le calcul
   Ctrl+S        : enregistrer les rapports dans le dossier courant
   Ctrl+Shift+S  : enregistrer les rapports sous un autre dossier
   Ctrl+G        : afficher ce guide
   Échap         : fermer la fenêtre active

10. EN CAS DE PROBLÈME
   Vérifiez le nom exact des feuilles Excel, la présence des colonnes attendues
   (voir la feuille « Notice » des exemples) et l'activation de l'environnement
   virtuel. Les messages d'erreur s'affichent dans une boîte de dialogue et le
   détail est journalisé dans la barre de statut.
"""

TEXTE_A_PROPOS = """À PROPOS — Dimensionnement hydraulique des équipements de points hauts

Version : {version}

Application de dimensionnement des ventouses aux points hauts et intermédiaires,
et des vidanges aux points bas, basée sur la résolution itérative de l'équation
de Colebrook et sur les règles de protection des conduites.

Fonctionnalités principales :
  • Modes « Profil en Long » et « Cas Simples »
  • Interface CustomTkinter avec barre de menus et barre d'outils
  • Exemples d'import intégrés (Fichier > Importer un exemple)
  • Choix du dossier de sortie et ouverture automatique du rapport Word
  • Catalogue MCP des ventouses double orifice 3 fonctions (PN 10/16) intégré
  • Export Excel multi-onglets et note de calcul Word

Auteur : Mostafa Karim Soussi (@km1965)
GitHub : https://github.com/km1965
Dépôt RPS 2011 : https://github.com/km1965/MK-Verifications_Sismique-RPS2011

Bibliothèques : customtkinter, pandas, numpy, matplotlib, openpyxl, python-docx, Pillow.

Sources du catalogue : Tableau_Recapitulatif_Debits_Ventouses_MCP.xlsx et .docx
fournis à la racine du projet.
""".format(version=VERSION_APP)


if HAS_CTK:
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme("blue")

    class HydrauliqueApp(ctk.CTk):
        def __init__(self):
            super().__init__()
            self.title(f"Dimensionnement Points Hauts, Intermédiaires & Points Bas — v{VERSION_APP}")
            self.geometry("1180x790")
            self.minsize(1000, 700)
            icone = chemin_ressource("assets/MK_icon.ico")
            if icone.exists():
                try:
                    self.iconbitmap(str(icone))
                except Exception:
                    pass  # icône facultative selon la plateforme

            self.input_xlsx_path = ""
            self.param_entries = {}
            self.last_figs = []
            self.last_out_dir = None
            self.last_xlsx_path = None
            self.last_docx_path = None
            self.last_resultats = None   # cache des DataFrames pour « Enregistrer »
            self.last_params = None
            self.last_titre = ""
            self.mode_var = ctk.StringVar(value="pl")
            self.auto_open_var = ctk.BooleanVar(value=True)
            self.inclure_catalogue_var = ctk.BooleanVar(value=True)
            self.status_var = ctk.StringVar(value="Prêt")

            self._build_menu()
            self._build_layout()
            self._bind_raccourcis()

        def _bind_raccourcis(self):
            self.bind("<Control-o>", lambda _e: self._choisir_fichier())
            self.bind("<Control-O>", lambda _e: self._choisir_fichier())
            self.bind("<F5>", lambda _e: self._lancer_calcul())
            self.bind("<Control-s>", lambda _e: self._enregistrer())
            self.bind("<Control-S>", lambda _e: self._enregistrer())
            self.bind("<Control-Shift-S>", lambda _e: self._enregistrer_sous())
            self.bind("<Control-g>", lambda _e: self._afficher_guide())
            self.bind("<Control-G>", lambda _e: self._afficher_guide())

        def _set_status(self, message: str):
            self.status_var.set(message)
            self.update_idletasks()

        def _journal(self, message: str):
            self.txt_resultats.insert("end", message)
            self.txt_resultats.see("end")

        def _build_menu(self):
            """Barre de menus native : Fichier / Calcul / Catalogue / Aide."""
            menubar = Menu(self, tearoff=0)

            menu_fichier = Menu(menubar, tearoff=0)
            menu_fichier.add_command(label="Sélectionner un fichier .xlsx...", accelerator="Ctrl+O",
                                     command=self._choisir_fichier)
            self.menu_exemples = Menu(menu_fichier, tearoff=0)
            self.menu_exemples.add_command(label="Profil en Long (17 piquets)",
                                           command=lambda: self._importer_exemple("pl"))
            self.menu_exemples.add_command(label="Cas Simples (3 jeux de données)",
                                           command=lambda: self._importer_exemple("simple"))
            menu_fichier.add_cascade(label="Importer un exemple", menu=self.menu_exemples)
            menu_fichier.add_separator()
            menu_fichier.add_command(label="Enregistrer les rapports", accelerator="Ctrl+S",
                                     command=self._enregistrer)
            menu_fichier.add_command(label="Enregistrer sous...", accelerator="Ctrl+Shift+S",
                                     command=self._enregistrer_sous)
            self.menu_enregistrer = (menu_fichier, 3, 4)  # indices « Enregistrer » / « Enregistrer sous »
            menu_fichier.add_separator()
            menu_fichier.add_command(label="Choisir le dossier de sortie...",
                                     command=self._choisir_dossier_sortie)
            menu_fichier.add_command(label="Ouvrir le dossier de sortie",
                                     command=self._ouvrir_dossier_sortie)
            menu_fichier.add_command(label="Ouvrir la note Word (.docx)",
                                     command=self._ouvrir_rapport_docx)
            menu_fichier.add_separator()
            menu_fichier.add_command(label="Quitter", command=self.destroy)
            menubar.add_cascade(label="Fichier", menu=menu_fichier)

            menu_calcul = Menu(menubar, tearoff=0)
            menu_calcul.add_command(label="Lancer le calcul & générer rapports", accelerator="F5",
                                    command=self._lancer_calcul)
            menu_calcul.add_separator()
            menu_calcul.add_radiobutton(label="Mode Profil en Long", variable=self.mode_var,
                                        value="pl", command=self._mode_change)
            menu_calcul.add_radiobutton(label="Mode Cas Simples", variable=self.mode_var,
                                        value="simple", command=self._mode_change)
            menu_calcul.add_separator()
            menu_calcul.add_checkbutton(label="Ouvrir le rapport Word après génération",
                                        variable=self.auto_open_var)
            menubar.add_cascade(label="Calcul", menu=menu_calcul)

            menu_catalogue = Menu(menubar, tearoff=0)
            menu_catalogue.add_command(label="Consulter le catalogue MCP (résumé)",
                                       command=self._afficher_catalogue)
            menu_catalogue.add_separator()
            menu_catalogue.add_command(label="Ouvrir le classeur catalogue (.xlsx)",
                                       command=lambda: self._ouvrir_source_catalogue(".xlsx"))
            menu_catalogue.add_command(label="Ouvrir la note catalogue (.docx)",
                                       command=lambda: self._ouvrir_source_catalogue(".docx"))
            menu_catalogue.add_separator()
            menu_catalogue.add_checkbutton(label="Inclure le catalogue dans les rapports",
                                           variable=self.inclure_catalogue_var)
            menubar.add_cascade(label="Catalogue", menu=menu_catalogue)

            menu_aide = Menu(menubar, tearoff=0)
            menu_aide.add_command(label="Guide utilisateur", accelerator="Ctrl+G",
                                  command=self._afficher_guide)
            menu_aide.add_command(label="À propos", command=self._afficher_a_propos)
            menubar.add_cascade(label="Aide", menu=menu_aide)

            self.config(menu=menubar)

        def _invalider_cache_resultats(self):
            """Le calcul en mémoire ne correspond plus au contexte : forcer un nouveau calcul."""
            self.last_resultats = None
            self.last_params = None
            self.last_titre = ""

        def _mode_change(self):
            mode = "Profil en Long" if self.mode_var.get() == "pl" else "Cas Simples"
            self._invalider_cache_resultats()
            self._set_status(f"Mode de calcul : {mode}")

        def _build_layout(self):
            self.grid_columnconfigure(0, weight=0, minsize=340)
            self.grid_columnconfigure(1, weight=1)
            self.grid_rowconfigure(0, weight=0)  # barre d'outils
            self.grid_rowconfigure(1, weight=1)  # contenu
            self.grid_rowconfigure(2, weight=0)  # barre de statut

            # --- Barre d'outils (raccourcis des menus) ---
            toolbar = ctk.CTkFrame(self, height=44, corner_radius=8)
            toolbar.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=(8, 4))
            for c in range(10):
                toolbar.grid_columnconfigure(c, weight=0)
            toolbar.grid_columnconfigure(9, weight=1)

            self.btn_exemple = ctk.CTkButton(
                toolbar, text="Importer un exemple ▾", width=175,
                fg_color="#7C3AED", hover_color="#6D28D9", command=self._afficher_menu_exemples
            )
            self.btn_exemple.grid(row=0, column=0, padx=(6, 3), pady=6)
            ctk.CTkButton(toolbar, text="Calculer (F5)", width=125, command=self._lancer_calcul).grid(
                row=0, column=1, padx=3, pady=6
            )
            ctk.CTkButton(toolbar, text="Dossier sortie", width=125, fg_color="#334155",
                          hover_color="#1E293B", command=self._choisir_dossier_sortie).grid(
                row=0, column=2, padx=3, pady=6
            )
            ctk.CTkButton(toolbar, text="Rapport Word", width=125, fg_color="#0F766E",
                          hover_color="#115E59", command=self._ouvrir_rapport_docx).grid(
                row=0, column=3, padx=3, pady=6
            )
            self.btn_save = ctk.CTkButton(toolbar, text="Enregistrer", width=115,
                                          fg_color="#2563EB", hover_color="#1D4ED8",
                                          command=self._enregistrer)
            self.btn_save.grid(row=0, column=4, padx=3, pady=6)
            self.btn_save_as = ctk.CTkButton(toolbar, text="Enregistrer sous…", width=135,
                                             fg_color="#2563EB", hover_color="#1D4ED8",
                                             command=self._enregistrer_sous)
            self.btn_save_as.grid(row=0, column=5, padx=3, pady=6)
            ctk.CTkFrame(toolbar, width=2, fg_color="#475569").grid(row=0, column=6, padx=8, pady=8)
            ctk.CTkButton(toolbar, text="Guide", width=80, command=self._afficher_guide).grid(
                row=0, column=7, padx=3, pady=6
            )
            ctk.CTkButton(toolbar, text="À propos", width=90, command=self._afficher_a_propos).grid(
                row=0, column=8, padx=3, pady=6
            )

            # --- Panneau latéral : paramètres et import ---
            sidebar = ctk.CTkScrollableFrame(self, label_text="Paramètres & Import (.xlsx)")
            sidebar.grid(row=1, column=0, sticky="nsew", padx=(12, 6), pady=(6, 6))
            sidebar.grid_columnconfigure(1, weight=1)

            ctk.CTkLabel(sidebar, text="Mode de calcul :", font=ctk.CTkFont(weight="bold")).grid(
                row=0, column=0, columnspan=2, sticky="w", padx=6, pady=(6, 2)
            )
            ctk.CTkRadioButton(sidebar, text="Profil en Long (PL)", variable=self.mode_var, value="pl").grid(
                row=1, column=0, sticky="w", padx=6, pady=2
            )
            ctk.CTkRadioButton(sidebar, text="Cas Simples (3 onglets)", variable=self.mode_var, value="simple").grid(
                row=1, column=1, sticky="w", padx=6, pady=2
            )

            ctk.CTkButton(sidebar, text="Sélectionner fichier .xlsx", command=self._choisir_fichier).grid(
                row=2, column=0, columnspan=2, sticky="ew", padx=6, pady=(10, 3)
            )
            ctk.CTkButton(sidebar, text="Importer un exemple ▾", fg_color="#7C3AED",
                          hover_color="#6D28D9", command=self._afficher_menu_exemples).grid(
                row=3, column=0, columnspan=2, sticky="ew", padx=6, pady=3
            )
            self.lbl_fichier = ctk.CTkLabel(sidebar, text="Aucun fichier sélectionné", text_color="gray", wraplength=290)
            self.lbl_fichier.grid(row=4, column=0, columnspan=2, sticky="w", padx=6, pady=(4, 8))

            ctk.CTkLabel(sidebar, text="Hypothèses hydrauliques :", font=ctk.CTkFont(weight="bold")).grid(
                row=5, column=0, columnspan=2, sticky="w", padx=6, pady=(6, 4)
            )

            champs = [
                ("dep_max", "Dépression max (mCE)", str(DEFAULT_PARAMS["dep_max"])),
                ("k_roughness", "Rugosité k (m)", str(DEFAULT_PARAMS["k_roughness"])),
                ("nu", "Viscosité nu (m²/s)", str(DEFAULT_PARAMS["nu"])),
                ("g", "Pesanteur g (m/s²)", str(DEFAULT_PARAMS["g"])),
                ("v_remplissage", "Vitesse rempl. (m/s)", str(DEFAULT_PARAMS["v_remplissage"])),
                ("mu_vidange", "Coeff. vidange mu", str(DEFAULT_PARAMS["mu_vidange"])),
                ("h_seuil_interm", "Seuil point interm. (m)", str(DEFAULT_PARAMS["h_seuil_interm"])),
                ("k_sing", "Coeff. singulier k", str(DEFAULT_PARAMS["k_sing"])),
            ]
            for idx, (key, label_txt, default_val) in enumerate(champs, start=6):
                ctk.CTkLabel(sidebar, text=label_txt).grid(row=idx, column=0, sticky="w", padx=6, pady=3)
                ent = ctk.CTkEntry(sidebar, width=110)
                ent.insert(0, default_val)
                ent.grid(row=idx, column=1, sticky="e", padx=6, pady=3)
                self.param_entries[key] = ent

            row = 14
            ctk.CTkLabel(sidebar, text="Dossier de sortie :", font=ctk.CTkFont(weight="bold")).grid(
                row=row, column=0, columnspan=2, sticky="w", padx=6, pady=(10, 2)
            )
            row += 1
            frame_sortie = ctk.CTkFrame(sidebar, fg_color="transparent")
            frame_sortie.grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=(0, 6))
            frame_sortie.grid_columnconfigure(0, weight=1)
            self.ent_outdir = ctk.CTkEntry(frame_sortie)
            self.ent_outdir.insert(0, "sorties_app")
            self.ent_outdir.grid(row=0, column=0, sticky="ew")
            ctk.CTkButton(frame_sortie, text="Parcourir...", width=95,
                          command=self._choisir_dossier_sortie).grid(row=0, column=1, padx=(6, 0))

            row += 1
            ctk.CTkCheckBox(sidebar, text="Ouvrir le rapport Word après génération",
                            variable=self.auto_open_var).grid(
                row=row, column=0, columnspan=2, sticky="w", padx=6, pady=2
            )
            row += 1
            ctk.CTkCheckBox(sidebar, text="Inclure le catalogue MCP dans les rapports",
                            variable=self.inclure_catalogue_var).grid(
                row=row, column=0, columnspan=2, sticky="w", padx=6, pady=2
            )

            row += 1
            self.btn_run = ctk.CTkButton(
                sidebar, text="Lancer le calcul & générer rapports", height=38,
                font=ctk.CTkFont(weight="bold"), command=self._lancer_calcul
            )
            self.btn_run.grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=(8, 6))

            row += 1
            self.btn_open_dir = ctk.CTkButton(
                sidebar, text="Ouvrir le dossier de sortie", fg_color="#334155", hover_color="#1E293B",
                state="disabled", command=self._ouvrir_dossier_sortie
            )
            self.btn_open_dir.grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=3)

            row += 1
            self.btn_open_docx = ctk.CTkButton(
                sidebar, text="Ouvrir la note Word (.docx)", fg_color="#0F766E", hover_color="#115E59",
                state="disabled", command=self._ouvrir_rapport_docx
            )
            self.btn_open_docx.grid(row=row, column=0, columnspan=2, sticky="ew", padx=6, pady=(3, 10))

            # --- Onglets principaux ---
            self.tabs = ctk.CTkTabview(self)
            self.tabs.grid(row=1, column=1, sticky="nsew", padx=(6, 12), pady=(6, 6))
            tab_res = self.tabs.add("Résultats & Journal")
            tab_fig = self.tabs.add("Aperçu Graphique")
            tab_cat = self.tabs.add("Catalogue MCP")

            tab_res.grid_columnconfigure(0, weight=1)
            tab_res.grid_rowconfigure(0, weight=1)
            self.txt_resultats = ctk.CTkTextbox(tab_res, font=ctk.CTkFont(family="Consolas", size=11))
            self.txt_resultats.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
            self.txt_resultats.insert(
                "1.0",
                "Prêt. Sélectionnez un fichier .xlsx ou importez un exemple, puis lancez le calcul (F5).\n"
                "Astuce : le menu Aide > Guide utilisateur décrit toutes les fonctionnalités.\n"
            )

            tab_fig.grid_columnconfigure(0, weight=1)
            tab_fig.grid_rowconfigure(1, weight=1)
            self.fig_selector = ctk.CTkSegmentedButton(tab_fig, values=[], command=self._afficher_figure)
            self.fig_selector.grid(row=0, column=0, padx=6, pady=6)
            self.lbl_image = ctk.CTkLabel(tab_fig, text="Aucun graphique généré pour le moment.")
            self.lbl_image.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)

            tab_cat.grid_columnconfigure(0, weight=1)
            tab_cat.grid_rowconfigure(0, weight=1)
            self.txt_catalogue = ctk.CTkTextbox(tab_cat, font=ctk.CTkFont(family="Consolas", size=11))
            self.txt_catalogue.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
            self.txt_catalogue.insert("1.0", self._texte_catalogue())
            self.txt_catalogue.configure(state="disabled")

            # --- Barre de statut ---
            status = ctk.CTkFrame(self, height=26, corner_radius=0)
            status.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 6))
            status.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(status, textvariable=self.status_var, anchor="w",
                         font=ctk.CTkFont(size=11)).grid(row=0, column=0, sticky="w", padx=10, pady=3)

        def _afficher_menu_exemples(self):
            """Ouvre le menu des exemples sous le bouton d'import."""
            try:
                x = self.btn_exemple.winfo_rootx()
                y = self.btn_exemple.winfo_rooty() + self.btn_exemple.winfo_height() + 2
                self.menu_exemples.tk_popup(x, y)
            finally:
                self.menu_exemples.grab_release()

        def _importer_exemple(self, mode: str):
            """Charge un classeur d'exemple et présélectionne le mode correspondant."""
            nom = ("Exemple_Import_Profil_En_Long.xlsx" if mode == "pl"
                   else "Exemple_Import_Cas_Simples.xlsx")
            chemin = chemin_ressource(nom)
            if not chemin.exists():
                messagebox.showerror(
                    "Exemple introuvable",
                    f"Le fichier {nom} est introuvable à la racine du projet :\n{BASE_DIR}"
                )
                return
            self.input_xlsx_path = str(chemin)
            self.lbl_fichier.configure(text=f"{nom} (exemple)", text_color=("black", "white"))
            self.mode_var.set(mode)
            self._mode_change()
            mode_txt = "Profil en Long" if mode == "pl" else "Cas Simples"
            self._journal(f"Exemple importé : {nom}\nMode activé : {mode_txt}\n")
            self.tabs.set("Résultats & Journal")
            self._set_status(f"Exemple chargé : {nom} — mode {mode_txt}")

        def _choisir_dossier_sortie(self):
            """Sélection graphique du dossier où écrire les rapports."""
            kwargs = {"title": "Choisir le dossier d'enregistrement des sorties"}
            initial = self.ent_outdir.get().strip()
            if initial and Path(initial).is_dir():
                kwargs["initialdir"] = initial
            dossier = filedialog.askdirectory(**kwargs)
            if dossier:
                self.ent_outdir.delete(0, "end")
                self.ent_outdir.insert(0, dossier)
                self._set_status(f"Dossier de sortie : {dossier}")

        def _texte_catalogue(self) -> str:
            """Résumé textuel du catalogue MCP (onglet et fenêtre dédiée)."""
            lignes = [
                "CATALOGUE MCP — VENTOUSES AUTOMATIQUES DOUBLE ORIFICE 3 FONCTIONS (PN 10/16)",
                "=" * 92,
                "",
                f"Corps : {CATALOGUE_MCP_SOURCE['materiau_corps']} | "
                f"Étanchéité min. : {CATALOGUE_MCP_SOURCE['etancheite_min_bar']} bar",
                f"Débit optimal : ΔP = {CATALOGUE_MCP_SOURCE['debit_optimal_bar']} bar | "
                f"Critique : +{CATALOGUE_MCP_SOURCE['debit_critique_evac_bar']} bar (évac.) / "
                f"{CATALOGUE_MCP_SOURCE['debit_critique_adm_bar']} bar (adm.)",
                "",
                "DÉBITS PAR DN (Std m³/h)",
                "-" * 92,
                f"{'DN':<26}{'Évac. opt.':>12}{'Évac. crit.':>13}{'Adm. opt.':>12}{'Adm. crit.':>13}{'Poids kg':>10}",
            ]
            for fic in VENTOUSES_MCP:
                lignes.append(
                    f"{fic['dn']:<26}{fic['evac_opt_m3h']:>12.0f}{fic['evac_crit_m3h']:>13.0f}"
                    f"{fic['adm_opt_m3h']:>12.0f}{fic['adm_crit_m3h']:>13.0f}{fic['poids_kg']:>10.1f}"
                )
            lignes += ["", "COURBES GRAND ORIFICE À 0,3 bar (Std m³/min)", "-" * 92]
            for dp in (0.3, 0.9):
                v = COURBES_MCP_EVACUATION[dp]
                lignes.append(
                    f"ΔP = +{dp:.1f} bar : courbe 50 mm = {v[0]:.0f} | courbe 80 mm = {v[1]:.0f} | "
                    f"courbe 100/150 = {v[2]:.0f} | avec papillon = {v[3]:.1f}"
                )
            v = COURBES_MCP_ADMISSION[-0.5]
            lignes.append(
                f"ΔP = -0.5 bar : courbe 50 mm = {v[0]:.1f} | courbe 80 mm = {v[1]:.1f} | "
                f"courbe 100/150 = {v[2]:.1f}"
            )
            lignes += ["", "PURGE PETIT ORIFICE (Std m³/min)", "-" * 92]
            tete = "Pression (bar) | " + " | ".join(f"Ø{d} mm" for d in DIAMETRES_PURGE_MCP)
            lignes.append(tete)
            for pression, debits in PURGE_PETIT_ORIFICE_MCP.items():
                vals = ["hors plage" if v is None else f"{v:.2f}" for v in debits]
                lignes.append(f"{pression:>13.0f} | " + " | ".join(f"{v:>7}" for v in vals))
            lignes += ["", "RÈGLES D'INGÉNIERIE", "-" * 92]
            for i, regle in enumerate(REGLES_INGENIERIE_MCP, 1):
                lignes.append(f"{i}. {regle}")
            lignes += [
                "",
                "Fichiers sources : " + " , ".join(CATALOGUE_MCP_SOURCE["fichiers_source"]),
                "Les données complètes sont exportées dans les rapports Excel/Word.",
            ]
            return "\n".join(lignes)

        def _afficher_catalogue(self):
            self._afficher_fenetre_texte(
                "Catalogue MCP — Ventouses double orifice 3 fonctions",
                self._texte_catalogue(), largeur=1040, hauteur=720
            )

        def _ouvrir_source_catalogue(self, extension: str):
            fichier = next(
                (f for f in CATALOGUE_MCP_SOURCE["fichiers_source"]
                 if f.lower().endswith(extension.lower())), None
            )
            chemin = chemin_ressource(fichier) if fichier else None
            if chemin and chemin.exists():
                _ouvrir_chemin_systeme(chemin)
                self._set_status(f"Ouverture de {fichier}")
            else:
                messagebox.showwarning(
                    "Catalogue MCP",
                    f"Fichier source {extension} introuvable à la racine du projet."
                )

        def _afficher_fenetre_texte(self, titre: str, contenu: str,
                                    largeur: int = 900, hauteur: int = 640):
            """Fenêtre modale générique pour le guide, À propos et le catalogue."""
            win = ctk.CTkToplevel(self)
            win.title(titre)
            win.geometry(f"{largeur}x{hauteur}")
            win.transient(self)
            win.grid_columnconfigure(0, weight=1)
            win.grid_rowconfigure(0, weight=1)

            txt = ctk.CTkTextbox(win, font=ctk.CTkFont(family="Consolas", size=11), wrap="word")
            txt.grid(row=0, column=0, sticky="nsew", padx=12, pady=(12, 6))
            txt.insert("1.0", contenu)
            txt.configure(state="disabled")

            ctk.CTkButton(win, text="Fermer", width=110, command=win.destroy).grid(
                row=1, column=0, pady=(0, 12)
            )
            win.grab_set()
            win.after(80, win.lift)

        def _afficher_guide(self):
            self._afficher_fenetre_texte("Guide utilisateur", GUIDE_UTILISATEUR, 950, 720)

        def _afficher_a_propos(self):
            self._afficher_fenetre_texte("À propos", TEXTE_A_PROPOS, 820, 520)

        def _enregistrer(self, dossier=None):
            """Réexporte le dernier calcul (Excel + Word) sans le refaire."""
            if self.last_resultats is None:
                if dossier:
                    self.ent_outdir.delete(0, "end")
                    self.ent_outdir.insert(0, str(dossier))
                self._journal("Aucun résultat en mémoire : lancement du calcul...\n")
                self._lancer_calcul()
                return

            cible = Path(dossier) if dossier else Path(self.ent_outdir.get().strip() or "sorties_app")
            self._set_status(f"Enregistrement dans {cible}...")
            try:
                figs = list(self.last_figs)
                if figs:
                    src_figs = (self.last_out_dir / "figures").resolve() if self.last_out_dir else None
                    dst_figs = (cible / "figures").resolve()
                    if src_figs != dst_figs:
                        dst_figs.mkdir(parents=True, exist_ok=True)
                        copies = []
                        for fig in figs:
                            src, dst = Path(fig), dst_figs / Path(fig).name
                            if src.resolve() != dst.resolve():
                                shutil.copy2(src, dst)
                            copies.append(dst)
                        figs = copies

                xlsx_p, docx_p = exporter_excel_et_docx(
                    self.last_resultats, figs, cible, self.last_titre, self.last_params,
                    inclure_catalogue=self.inclure_catalogue_var.get()
                )
                self.last_out_dir = cible
                self.last_xlsx_path = Path(xlsx_p)
                self.last_docx_path = Path(docx_p)
                self.last_figs = figs
                self.btn_open_dir.configure(state="normal")
                self.btn_open_docx.configure(state="normal")
                self._journal(f"Enregistré dans {cible} :\n- Excel : {xlsx_p}\n- Word  : {docx_p}\n")
                self._set_status(f"Enregistré dans {cible}")
                messagebox.showinfo("Enregistrer", f"Rapports enregistrés dans :\n{cible}")
            except Exception as exc:
                self._set_status(f"Erreur enregistrement : {exc}")
                self._journal(f"ERREUR enregistrement : {exc}\n")
                messagebox.showerror("Enregistrer", str(exc))

        def _enregistrer_sous(self):
            """Demande un dossier puis y enregistre les rapports (Ctrl+Shift+S)."""
            initial = self.ent_outdir.get().strip() or "sorties_app"
            kwargs = {"title": "Enregistrer les rapports dans un dossier"}
            if Path(initial).is_dir():
                kwargs["initialdir"] = initial
            dossier = filedialog.askdirectory(**kwargs)
            if dossier:
                self.ent_outdir.delete(0, "end")
                self.ent_outdir.insert(0, dossier)
                self._enregistrer(dossier)

        def _choisir_fichier(self):
            fpath = filedialog.askopenfilename(
                title="Choisir un fichier Excel d'entrée",
                filetypes=[("Fichiers Excel", "*.xlsx *.xls")]
            )
            if fpath:
                self.input_xlsx_path = fpath
                self._invalider_cache_resultats()
                self.lbl_fichier.configure(text=Path(fpath).name, text_color=("black", "white"))
                self._journal(f"Fichier d'entrée : {fpath}\n")
                self._set_status(f"Fichier chargé : {Path(fpath).name}")

        def _lire_parametres(self) -> dict:
            p = {}
            for k, ent in self.param_entries.items():
                p[k] = float(ent.get().strip().replace(",", "."))
            return p

        def _afficher_figure(self, fig_name: str):
            for fp in self.last_figs:
                if Path(fp).stem == fig_name and Path(fp).exists():
                    img = Image.open(fp)
                    ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=(760, 350))
                    self.lbl_image.configure(image=ctk_img, text="")
                    self.lbl_image.image = ctk_img
                    return

        def _ouvrir_dossier_sortie(self):
            if self.last_out_dir and self.last_out_dir.exists():
                _ouvrir_chemin_systeme(self.last_out_dir)

        def _ouvrir_rapport_docx(self):
            if self.last_docx_path and self.last_docx_path.exists():
                _ouvrir_chemin_systeme(self.last_docx_path)

        def _lancer_calcul(self):
            if not self.input_xlsx_path or not Path(self.input_xlsx_path).exists():
                messagebox.showwarning("Fichier manquant", "Veuillez d'abord sélectionner un fichier .xlsx valide.")
                return
            self._set_status("Calcul en cours...")
            try:
                params = self._lire_parametres()
                out_dir = Path(self.ent_outdir.get().strip() or "sorties_app")
                mode = self.mode_var.get()

                self.txt_resultats.delete("1.0", "end")
                if mode == "pl":
                    df_pl = pd.read_excel(self.input_xlsx_path, sheet_name="Profil_en_Long")
                    df_out, res_ph, res_pi, res_pb = calculer_profil_en_long(df_pl, params)
                    figs = generer_figures_pl(df_out, res_ph, res_pi, res_pb, out_dir / "figures")
                    xlsx_p, docx_p = exporter_excel_et_docx(
                        {
                            "Profil_Detail": df_out,
                            "Points_Hauts": res_ph,
                            "Points_Intermediaires": res_pi,
                            "Points_Bas": res_pb,
                        },
                        figs, out_dir, "Note de calcul — Profil en Long Complet", params,
                        inclure_catalogue=self.inclure_catalogue_var.get()
                    )
                    self.txt_resultats.insert("end", "=== MODE PROFIL EN LONG ===\n")
                    self.txt_resultats.insert("end", f"Points hauts détectés : {len(res_ph)}\n")
                    self.txt_resultats.insert("end", f"Points intermédiaires équipés : {len(res_pi)}\n")
                    self.txt_resultats.insert("end", f"Points bas (vidanges) : {len(res_pb)}\n\n")
                    self.txt_resultats.insert("end", "--- Points Hauts ---\n" + res_ph.to_string(index=False) + "\n\n")
                    if not res_pi.empty:
                        self.txt_resultats.insert("end", "--- Points Intermédiaires ---\n" + res_pi.to_string(index=False) + "\n\n")
                    self.txt_resultats.insert("end", "--- Points Bas ---\n" + res_pb.to_string(index=False) + "\n\n")
                else:
                    df_ph = pd.read_excel(self.input_xlsx_path, sheet_name="Points_Hauts")
                    df_pi = pd.read_excel(self.input_xlsx_path, sheet_name="Points_Intermediaires")
                    df_pb = pd.read_excel(self.input_xlsx_path, sheet_name="Points_Bas")
                    res_ph, res_pi, res_pb = calculer_mode_simple(df_ph, df_pi, df_pb, params)
                    figs = []
                    xlsx_p, docx_p = exporter_excel_et_docx(
                        {"Points_Hauts": res_ph, "Points_Intermediaires": res_pi, "Points_Bas": res_pb},
                        figs, out_dir, "Note de calcul — Cas Simples", params,
                        inclure_catalogue=self.inclure_catalogue_var.get()
                    )
                    self.txt_resultats.insert("end", "=== MODE CAS SIMPLES ===\n")
                    self.txt_resultats.insert("end", "--- Points Hauts ---\n" + res_ph.to_string(index=False) + "\n\n")
                    self.txt_resultats.insert("end", "--- Points Intermédiaires ---\n" + res_pi.to_string(index=False) + "\n\n")
                    self.txt_resultats.insert("end", "--- Points Bas ---\n" + res_pb.to_string(index=False) + "\n\n")

                # Mise en cache pour les actions « Enregistrer » / « Enregistrer sous ».
                self.last_params = params
                if mode == "pl":
                    self.last_resultats = {
                        "Profil_Detail": df_out,
                        "Points_Hauts": res_ph,
                        "Points_Intermediaires": res_pi,
                        "Points_Bas": res_pb,
                    }
                    self.last_titre = "Note de calcul — Profil en Long Complet"
                else:
                    self.last_resultats = {
                        "Points_Hauts": res_ph,
                        "Points_Intermediaires": res_pi,
                        "Points_Bas": res_pb,
                    }
                    self.last_titre = "Note de calcul — Cas Simples"

                self.last_out_dir = out_dir
                self.last_xlsx_path = Path(xlsx_p)
                self.last_docx_path = Path(docx_p)
                self.btn_open_dir.configure(state="normal")
                self.btn_open_docx.configure(state="normal")

                self.txt_resultats.insert("end", f"Fichiers générés :\n- Excel : {xlsx_p}\n- Word  : {docx_p}\n")
                self.last_figs = figs
                if figs:
                    names = [Path(f).stem for f in figs]
                    self.fig_selector.configure(values=names)
                    self.fig_selector.set(names[0])
                    self._afficher_figure(names[0])

                self._set_status(f"Terminé — rapports dans {out_dir}")
                messagebox.showinfo("Succès", f"Calcul terminé.\nRapports enregistrés dans : {out_dir}")
                if self.auto_open_var.get() and self.last_docx_path and self.last_docx_path.exists():
                    self._journal("Ouverture automatique de la note Word...\n")
                    self._ouvrir_rapport_docx()
            except Exception as exc:
                self._set_status(f"Erreur : {exc}")
                self._journal(f"ERREUR : {exc}\n")
                messagebox.showerror("Erreur de calcul", str(exc))


def main():
    parser = argparse.ArgumentParser(description="main1.py — Calcul équipements Points Hauts, Intermédiaires & Bas")
    parser.add_argument("--gui", action="store_true", help="Lancer l'interface graphique CustomTkinter")
    parser.add_argument("--mode", choices=["simple", "pl"], help="Mode ligne de commande : 'simple' ou 'pl'")
    parser.add_argument("--input-xlsx", type=str, help="Chemin du fichier Excel (.xlsx) d'entrée")
    parser.add_argument("--out-dir", type=str, default="sorties_cli", help="Dossier de sortie")
    args = parser.parse_args()

    if args.mode is None or args.gui:
        if not HAS_CTK:
            raise RuntimeError("customtkinter ou Pillow n'est pas installé. Lancez : pip install -r requirements.txt")
        app = HydrauliqueApp()
        app.mainloop()
        return

    if not args.input_xlsx:
        raise ValueError("Veuillez préciser --input-xlsx en mode ligne de commande (--mode).")

    out_dir = Path(args.out_dir)
    if args.mode == "simple":
        df_ph = pd.read_excel(args.input_xlsx, sheet_name="Points_Hauts")
        df_pi = pd.read_excel(args.input_xlsx, sheet_name="Points_Intermediaires")
        df_pb = pd.read_excel(args.input_xlsx, sheet_name="Points_Bas")
        res_ph, res_pi, res_pb = calculer_mode_simple(df_ph, df_pi, df_pb, DEFAULT_PARAMS)
        xlsx_p, docx_p = exporter_excel_et_docx(
            {"Points_Hauts": res_ph, "Points_Intermediaires": res_pi, "Points_Bas": res_pb},
            [], out_dir, "Note de calcul — Cas Simples", DEFAULT_PARAMS
        )
    else:
        df_pl = pd.read_excel(args.input_xlsx, sheet_name="Profil_en_Long")
        df_out, res_ph, res_pi, res_pb = calculer_profil_en_long(df_pl, DEFAULT_PARAMS)
        figs = generer_figures_pl(df_out, res_ph, res_pi, res_pb, out_dir / "figures")
        xlsx_p, docx_p = exporter_excel_et_docx(
            {
                "Profil_Detail": df_out,
                "Points_Hauts": res_ph,
                "Points_Intermediaires": res_pi,
                "Points_Bas": res_pb,
            },
            figs, out_dir, "Note de calcul — Profil en Long Complet", DEFAULT_PARAMS
        )

    message = f"Calcul terminé avec succès -> Excel : {xlsx_p} | Word : {docx_p}"
    # En mode --windowed (exe), sys.stdout est absent : ne pas faire échouer le calcul.
    try:
        if sys.stdout is not None:
            print(message)
    except Exception:
        pass
    try:
        (out_dir / "calcul_termine.txt").write_text(message + "\n", encoding="utf-8")
    except Exception:
        pass


if __name__ == "__main__":
    try:
        main()
    except Exception:
        import traceback
        import tempfile

        log = Path(tempfile.gettempdir()) / "MK_Prandtl-Colebrook_erreur.log"
        try:
            log.write_text(traceback.format_exc(), encoding="utf-8")
        except Exception:
            pass
        if HAS_CTK:
            try:
                from tkinter import messagebox as _mb
                _mb.showerror("Erreur", f"Une erreur est survenue.\nDétail : {log}")
            except Exception:
                pass
        raise