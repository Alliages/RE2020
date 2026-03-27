"""
Calculateur de seuils RE2020 — version Streamlit
Basé sur le fichier Excel de Guillaume Meunier (v2.1, mars 2026)
Document non officiel — GNU GPL v3
"""

import streamlit as st

st.set_page_config(
    page_title="Calculateur Seuils RE2020",
    page_icon="🏗️",
    layout="wide",
)

# ─── Données de référence ──────────────────────────────────────────────────────

ZONES = {
    "Ain": ("H1c", 1), "Aisne": ("H1a", 2), "Allier": ("H1c", 3),
    "Alpes-de-Haute-Provence": ("H2d", 4), "Alpes-Maritimes": ("H3", 6),
    "Ardèche": ("H2d", 7), "Ardennes": ("H1b", 8), "Ariège": ("H2c", 9),
    "Aube": ("H1b", 10), "Aude": ("H3", 11), "Aveyron": ("H2c", 12),
    "Bas-Rhin": ("H1b", 67), "Bouches-du-Rhône": ("H3", 13),
    "Calvados": ("H1a", 14), "Cantal": ("H1c", 15), "Charente": ("H2b", 16),
    "Charente-Maritime": ("H2b", 17), "Cher": ("H2b", 18),
    "Corrèze": ("H1c", 19), "Corse Haute-Corse": ("H3", "2B"),
    "Corse-du-Sud": ("H3", "2A"), "Côte-d'Or": ("H1c", 21),
    "Côtes-d'Armor": ("H2a", 22), "Creuse": ("H1c", 23),
    "Deux-Sèvres": ("H2b", 79), "Dordogne": ("H2c", 24),
    "Doubs": ("H1c", 25), "Drôme": ("H2d", 26), "Eure": ("H1a", 27),
    "Eure-et-Loir": ("H1a", 28), "Finistère": ("H2a", 29),
    "Gard": ("H3", 30), "Haute-Garonne": ("H2c", 31), "Gers": ("H2c", 32),
    "Gironde": ("H2b", 33), "Haut-Rhin": ("H1b", 68),
    "Haute-Loire": ("H1c", 43), "Haute-Marne": ("H1b", 52),
    "Haute-Saône": ("H1c", 70), "Haute-Savoie": ("H1c", 74),
    "Haute-Vienne": ("H1c", 87), "Hautes-Alpes": ("H2d", 5),
    "Hautes-Pyrénées": ("H2c", 65), "Hérault": ("H3", 34),
    "Ille-et-Vilaine": ("H2a", 35), "Indre": ("H2b", 36),
    "Indre-et-Loire": ("H2b", 37), "Isère": ("H1c", 38), "Jura": ("H1c", 39),
    "Landes": ("H2b", 40), "Loir-et-Cher": ("H2b", 41), "Loire": ("H1c", 42),
    "Loire-Atlantique": ("H2a", 44), "Loiret": ("H1b", 45),
    "Lot": ("H2c", 46), "Lot-et-Garonne": ("H2c", 47), "Lozère": ("H1c", 48),
    "Maine-et-Loire": ("H2a", 49), "Manche": ("H1a", 50), "Marne": ("H1b", 51),
    "Mayenne": ("H2a", 53), "Meurthe-et-Moselle": ("H1b", 54),
    "Meuse": ("H1b", 55), "Morbihan": ("H2a", 56), "Moselle": ("H1b", 57),
    "Nièvre": ("H1b", 58), "Nord": ("H1a", 59), "Oise": ("H1a", 60),
    "Orne": ("H1a", 61), "Paris": ("H1a", 75), "Pas-de-Calais": ("H1a", 62),
    "Puy-de-Dôme": ("H1c", 63), "Pyrénées-Atlantiques": ("H2c", 64),
    "Pyrénées-Orientales": ("H3", 66), "Saône-et-Loire": ("H1c", 71),
    "Sarthe": ("H2b", 72), "Savoie": ("H1c", 73),
    "Seine-et-Marne": ("H1a", 77), "Seine-Maritime": ("H1a", 76),
    "Seine-Saint-Denis": ("H1a", 93), "Somme": ("H1a", 80),
    "Tarn": ("H2c", 81), "Tarn-et-Garonne": ("H2c", 82),
    "Territoire de Belfort": ("H1b", 90), "Val-de-Marne": ("H1a", 94),
    "Val-d'Oise": ("H1a", 95), "Var": ("H3", 83), "Vaucluse": ("H2d", 84),
    "Vendée": ("H2b", 85), "Vienne": ("H2b", 86), "Vosges": ("H1b", 88),
    "Yonne": ("H1b", 89), "Yvelines": ("H1a", 78),
}

USAGE_LABELS = {
    "Maisons_individuelles_ou_accolées": "Maisons individuelles ou accolées",
    "Logement_collectif": "Logement collectif",
    "Bureaux": "Bureaux",
    "Enseignement_primaire_ou_secondaire": "Enseignement primaire ou secondaire",
    "Médiathèques_et_bibliothèques": "Médiathèques et bibliothèques",
    "Bâtiments_universitaires": "Bâtiments universitaires et atypiques",
    "Hôtels": "Hôtels",
    "Etablissements_petite_enfance": "Établissements accueil petite enfance",
    "Restaurants": "Restaurants",
    "Commerces": "Commerces",
    "Vestiaires": "Vestiaires seuls",
    "Etablissements_sanitaires": "Établissements sanitaires avec hébergement",
    "Etablissements_sante": "Établissements de santé",
    "Aérogares": "Aérogares",
    "Industries": "Industries ou artisanats",
    "Sport_municipal": "Établissements sportifs municipaux/scolaires",
    "Restaurants_scolaires": "Restaurants scolaires",
    "Sport_prive": "Établissements sportifs privés",
}

BBIO_MOYENS = {
    "Logement_collectif": 65,
    "Maisons_individuelles_ou_accolées": 63,
    "Bureaux": 95,
    "Enseignement_primaire_ou_secondaire - Enseignement primaire": 68,
    "Enseignement_primaire_ou_secondaire - Enseignement secondaire": 68,
    "Médiathèques_et_bibliothèques": 117,
    "Bâtiments_universitaires": 122,
    "Hôtels - 0,1,2 - nuit": 76,
    "Hôtels - 3,4,5 - nuit": 76,
    "Hôtels - 0,1,2 - jour": 134,
    "Hôtels - 3,4,5 - jour": 163,
    "Etablissements_petite_enfance": 139,
    "Restaurants - continu": 245,
    "Restaurants - 1r5j": 100,
    "Restaurants - 2r7j": 206,
    "Restaurants - 2r6j": 177,
    "Commerces": 170,
    "Vestiaires": 225,
    "Etablissements_sanitaires": 174,
    "Etablissements_sante - nuit": 164,
    "Etablissements_sante - jour": 133,
    "Aérogares": 248,
    "Industries - 3x8": 257,
    "Industries - 8h18": 69,
    "Sport_municipal": 94,
    "Restaurants_scolaires - 1r5j": 76,
    "Restaurants_scolaires - 3r5j": 76,
    "Sport_prive": 94,
}

CEP_NR_MOYENS = {
    "Maisons_individuelles_ou_accolées": 55,
    "Logement_collectif": 70,
    "Bureaux": 75,
    "Enseignement_primaire_ou_secondaire - Enseignement primaire": 65,
    "Enseignement_primaire_ou_secondaire - Enseignement secondaire": 63,
    "Médiathèques_et_bibliothèques": 93,
    "Bâtiments_universitaires": 102,
    "Hôtels - 0,1,2 - nuit": 121,
    "Hôtels - 3,4,5 - nuit": 118,
    "Hôtels - 0,1,2 - jour": 235,
    "Hôtels - 3,4,5 - jour": 234,
    "Etablissements_petite_enfance": 150,
    "Restaurants - continu": 282,
    "Restaurants - 1r5j": 132,
    "Restaurants - 2r7j": 219,
    "Restaurants - 2r6j": 214,
    "Commerces": 163,
    "Vestiaires": 242,
    "Etablissements_sanitaires": 190,
    "Etablissements_sante - nuit": 274,
    "Etablissements_sante - jour": 165,
    "Aérogares": 191,
    "Industries - 3x8": 290,
    "Industries - 8h18": 94,
    "Sport_municipal": 94,
    "Restaurants_scolaires - 1r5j": 94,
    "Restaurants_scolaires - 3r5j": 94,
    "Sport_prive": 94,
}

CEP_MOYENS = {
    "Maisons_individuelles_ou_accolées": 75,
    "Logement_collectif": 85,
    "Bureaux": 85,
    "Enseignement_primaire_ou_secondaire - Enseignement primaire": 72,
    "Enseignement_primaire_ou_secondaire - Enseignement secondaire": 72,
    "Médiathèques_et_bibliothèques": 105,
    "Bâtiments_universitaires": 112,
    "Hôtels - 0,1,2 - nuit": 144,
    "Hôtels - 3,4,5 - nuit": 138,
    "Hôtels - 0,1,2 - jour": 252,
    "Hôtels - 3,4,5 - jour": 281,
    "Etablissements_petite_enfance": 182,
    "Restaurants - continu": 578,
    "Restaurants - 1r5j": 275,
    "Restaurants - 2r7j": 446,
    "Restaurants - 2r6j": 412,
    "Commerces": 182,
    "Vestiaires": 306,
    "Etablissements_sanitaires": 252,
    "Etablissements_sante - nuit": 302,
    "Etablissements_sante - jour": 180,
    "Aérogares": 253,
    "Industries - 3x8": 365,
    "Industries - 8h18": 116,
    "Sport_municipal": 116,
    "Restaurants_scolaires - 1r5j": 116,
    "Restaurants_scolaires - 3r5j": 116,
    "Sport_prive": 116,
}

IC_ENERGIE_MOYENS = {
    "Logement_collectif": {"rcu_classe_2022_2024": 560, "rcu_classe_2025+": 320, "rcu_autre": 260, "no_rcu": 260},
    "Maisons_individuelles_ou_accolées": {"rcu_classe_2022_2024": 200, "rcu_classe_2025+": 200, "rcu_autre": 160, "no_rcu": 160},
    "Bureaux": {"rcu_classe_2022_2024": 280, "rcu_classe_2025+": 200, "rcu_autre": 200, "no_rcu": 200},
    "Enseignement_primaire_ou_secondaire - Enseignement primaire": {"rcu_classe_2022_2024": 240, "rcu_classe_2025+": 200, "rcu_autre": 140, "no_rcu": 140},
    "Enseignement_primaire_ou_secondaire - Enseignement secondaire": {"rcu_classe_2022_2024": 240, "rcu_classe_2025+": 200, "rcu_autre": 140, "no_rcu": 140},
    "Médiathèques_et_bibliothèques": {"rcu_classe_2025+": 285, "rcu_autre": 285, "no_rcu": 285},
    "Bâtiments_universitaires": {"rcu_classe_2025+": 190, "rcu_autre": 190, "no_rcu": 190},
    "Hôtels - 0,1,2 - nuit": {"rcu_classe_2022_2024": 490, "rcu_autre": 390, "no_rcu": 390},
    "Hôtels - 3,4,5 - nuit": {"rcu_classe_2022_2024": 485, "rcu_autre": 350, "no_rcu": 350},
    "Hôtels - 0,1,2 - jour": {"rcu_classe_2022_2024": 595, "rcu_autre": 495, "no_rcu": 495},
    "Hôtels - 3,4,5 - jour": {"rcu_classe_2022_2024": 630, "rcu_autre": 520, "no_rcu": 520},
    "Etablissements_petite_enfance": {"rcu_classe_2025+": 680, "rcu_autre": 680, "no_rcu": 680},
    "Restaurants - continu": {"rcu_classe_2025+": 570, "rcu_autre": 570, "no_rcu": 570},
    "Restaurants - 1r5j": {"rcu_autre": 570, "no_rcu": 570},
    "Restaurants - 2r7j": {"rcu_autre": 570, "no_rcu": 570},
    "Restaurants - 2r6j": {"rcu_autre": 570, "no_rcu": 570},
    "Commerces": {"rcu_autre": 400, "no_rcu": 400},
    "Vestiaires": {"rcu_autre": 600, "no_rcu": 600},
    "Etablissements_sanitaires": {"rcu_autre": 400, "no_rcu": 400},
    "Etablissements_sante - nuit": {"rcu_autre": 400, "no_rcu": 400},
    "Etablissements_sante - jour": {"rcu_autre": 400, "no_rcu": 400},
    "Aérogares": {"rcu_autre": 400, "no_rcu": 400},
    "Industries - 3x8": {"rcu_autre": 400, "no_rcu": 400},
    "Industries - 8h18": {"rcu_autre": 400, "no_rcu": 400},
    "Sport_municipal": {"rcu_autre": 400, "no_rcu": 400},
    "Restaurants_scolaires - 1r5j": {"rcu_autre": 400, "no_rcu": 400},
    "Restaurants_scolaires - 3r5j": {"rcu_autre": 400, "no_rcu": 400},
    "Sport_prive": {"rcu_autre": 400, "no_rcu": 400},
}

# Mccat (en unité entière, à diviser par 1000 pour obtenir le ratio)
MCCAT_ENERGIE = {
    "Logement_collectif": {1: 3, 2: 2, 3: None},
    "Maisons_individuelles_ou_accolées": {1: 0, 2: 0, 3: None},
    "Bureaux": {1: 6, 2: 4, 3: None},
    "Enseignement_primaire_ou_secondaire - Enseignement primaire": {1: 9, 2: 6, 3: 7},
    "Enseignement_primaire_ou_secondaire - Enseignement secondaire": {1: 9, 2: 8, 3: 8},
    "Médiathèques_et_bibliothèques": {1: 12, 2: 10, 3: 9},
    "Bâtiments_universitaires": {1: 15, 2: 12, 3: None},
    "Hôtels - 0,1,2 - nuit": {1: 18, 2: 14, 3: 15},
    "Hôtels - 3,4,5 - nuit": {1: 21, 2: 14, 3: 15},
    "Hôtels - 0,1,2 - jour": {1: 24, 2: 14, 3: 15},
    "Hôtels - 3,4,5 - jour": {1: 27, 2: 14, 3: 15},
    "Etablissements_petite_enfance": {1: 30, 2: 16, 3: 18},
    "Restaurants - continu": {1: 33, 2: 18, 3: 21},
    "Restaurants - 1r5j": {1: 36, 2: 18, 3: 21},
    "Restaurants - 2r7j": {1: 39, 2: 18, 3: 21},
    "Restaurants - 2r6j": {1: 42, 2: 18, 3: 21},
    "Commerces": {1: 45, 2: 20, 3: 24},
    "Vestiaires": {1: 48, 2: 22, 3: 27},
    "Etablissements_sanitaires": {1: 51, 2: 24, 3: 30},
    "Etablissements_sante - nuit": {1: 54, 2: 26, 3: 33},
    "Etablissements_sante - jour": {1: 57, 2: 26, 3: 33},
    "Aérogares": {1: 60, 2: 28, 3: 36},
    "Industries - 3x8": {1: 63, 2: 30, 3: 39},
    "Industries - 8h18": {1: 66, 2: 30, 3: 39},
    "Sport_municipal": {1: 69, 2: 32, 3: 42},
    "Restaurants_scolaires - 1r5j": {1: 72, 2: None, 3: None},
    "Restaurants_scolaires - 3r5j": {1: 72, 2: None, 3: None},
    "Sport_prive": {1: 75, 2: None, 3: None},
}

ICC_MOYENS = {
    "Maisons_individuelles_ou_accolées": {"2022_2024": 640, "2025_2027": 530, "2028_2030": 451.25, "2031+": 428.69},
    "Logement_collectif": {"2022_2024": 740, "2025_2027": 650, "2028_2030": 551, "2031+": 523.45},
    "Bureaux": {"2022_2024": 980, "2025_2027": 810, "2028_2030": 674.5, "2031+": 640.78},
    "Enseignement_primaire_ou_secondaire": {"2022_2024": 900, "2025_2027": 770, "2028_2030": 646, "2031+": 613.7},
    "Médiathèques_et_bibliothèques": {"2025_2027": 940, "2028_2030": 745.75, "2031+": 708.46},
    "Bâtiments_universitaires": {"2025_2027": 940, "2028_2030": 750.5, "2031+": 712.98},
    "Hôtels": {"2025_2027": 820, "2028_2030": 646, "2031+": 613.7},
    "Etablissements_petite_enfance": {"2025_2027": 950, "2028_2030": 741, "2031+": 703.95},
    "Restaurants": {"2025_2027": 800, "2028_2030": 636.5, "2031+": 604.68},
    "Commerces": {"2025_2027": 800, "2028_2030": 636.5, "2031+": 604.68},
    "Vestiaires": {"2025_2027": 1050, "2028_2030": 855, "2031+": 812.25},
    "Etablissements_sanitaires": {"2025_2027": 880, "2028_2030": 722, "2031+": 685.9},
    "Etablissements_sante": {"2025_2027": 880, "2028_2030": 722, "2031+": 685.9},
    "Aérogares": {"2025_2027": 1120, "2028_2030": 902.5, "2031+": 857.38},
    "Industries": {"2025_2027": 840, "2028_2030": 660.25, "2031+": 627.24},
    "Sport_municipal": {"2025_2027": 900, "2028_2030": 722, "2031+": 685.9},
    "Restaurants_scolaires": {"2025_2027": 800, "2028_2030": 636.5, "2031+": 604.68},
    "Sport_prive": {"2025_2027": 900, "2028_2030": 722, "2031+": 685.9},
}

DH_MAX = {
    "Maisons_individuelles_ou_accolées": 1250,
    "Logement_collectif": 1250,
    "Bureaux": 1150,
    "Enseignement_primaire_ou_secondaire - Enseignement primaire": 900,
    "Enseignement_primaire_ou_secondaire - Enseignement secondaire": 900,
    "Médiathèques_et_bibliothèques": 900,
    "Bâtiments_universitaires": 900,
    "Hôtels - 0,1,2 - nuit": 300,
    "Hôtels - 3,4,5 - nuit": 300,
    "Hôtels - 0,1,2 - jour": 1300,
    "Hôtels - 3,4,5 - jour": 1300,
    "Etablissements_petite_enfance": 550,
    "Restaurants - continu": 2500,
    "Restaurants - 1r5j": 250,
    "Restaurants - 2r7j": 1600,
    "Restaurants - 2r6j": 1250,
    "Commerces": 3300,
    "Vestiaires": 1000,
    "Etablissements_sanitaires": 900,
    "Etablissements_sante - nuit": 900,
    "Etablissements_sante - jour": 1250,
    "Aérogares": 12100,
    "Industries - 3x8": 3200,
    "Industries - 8h18": 900,
    "Sport_municipal": 2000,
    "Restaurants_scolaires - 1r5j": 40,
    "Restaurants_scolaires - 3r5j": 260,
    "Sport_prive": 2000,
}

DH_EFFINERGIE = {
    "Maisons_individuelles_ou_accolées": 600,
    "Logement_collectif": 600,
    "Bureaux": 600,
    "Enseignement_primaire_ou_secondaire - Enseignement primaire": 600,
    "Enseignement_primaire_ou_secondaire - Enseignement secondaire": 600,
}


# ─── Fonctions de calcul ───────────────────────────────────────────────────────

def get_detailed_usage(usage, hotel_detail, resto_detail, resto_sco_detail,
                       industrie_detail, enseignement_detail, sante_detail):
    if usage == "Hôtels":
        return f"Hôtels - {hotel_detail}"
    if usage == "Restaurants":
        return f"Restaurants - {resto_detail}"
    if usage == "Restaurants_scolaires":
        return f"Restaurants_scolaires - {resto_sco_detail}"
    if usage == "Industries":
        return f"Industries - {industrie_detail}"
    if usage == "Enseignement_primaire_ou_secondaire":
        return f"Enseignement_primaire_ou_secondaire - Enseignement {enseignement_detail}"
    if usage == "Etablissements_sante":
        return f"Etablissements_sante - {sante_detail}"
    return usage


def get_icc_key(usage):
    mapping = {
        "Enseignement_primaire_ou_secondaire": "Enseignement_primaire_ou_secondaire",
        "Hôtels": "Hôtels",
        "Industries": "Industries",
        "Restaurants": "Restaurants",
        "Restaurants_scolaires": "Restaurants_scolaires",
        "Etablissements_sante": "Etablissements_sante",
    }
    return mapping.get(usage, usage)


def calculate(
    usage, du, zone, sref, nl, hsp, sagr, bruit, categorie, inertie,
    annee_pc, igh, rcu, rcu_classe, clim, iclot1, iclot2, iclot13, icded,
):
    is_logement = usage in ("Maisons_individuelles_ou_accolées", "Logement_collectif")
    is_ml = usage == "Maisons_individuelles_ou_accolées"
    is_lc = usage == "Logement_collectif"
    smoy = sref / nl if is_logement and nl > 0 else None

    # ── BBIO ─────────────────────────────────────────────────────────────────
    bbio_moy = BBIO_MOYENS.get(du) or BBIO_MOYENS.get(usage)

    mbgeo = mbcombles = mbsurf_moy = mbsurf_tot = mbbruit = mbhsp = 0.0

    if bbio_moy is not None:
        if is_lc and zone in ("H2d", "H3"):
            mbgeo = 0.10
        if is_lc and smoy is not None:
            mbsurf_moy = max(-0.15, min(0.15, (65 - smoy) / 65 * 0.15))
        if is_lc:
            mbsurf_tot = max(-0.05, min(0.05, (1500 - sref) / 10000))
        bruit_coeff = {"BR1": 0, "BR2": 0.04, "BR3": 0.08}
        mbbruit = bruit_coeff.get(bruit, 0) if is_logement else 0
        if is_lc and hsp > 2.5:
            mbhsp = min(0.10, (hsp - 2.5) / 2.5 * 0.04)
        elif is_ml and hsp > 2.5:
            mbhsp = min(0.10, (hsp - 2.5) / 2.5 * 0.08)

    sum_mb = mbgeo + mbcombles + mbsurf_moy + mbsurf_tot + mbbruit + mbhsp
    bbio_max = bbio_moy * (1 + sum_mb) if bbio_moy is not None else None

    # ── ÉNERGIE ──────────────────────────────────────────────────────────────
    cepnr_moy = CEP_NR_MOYENS.get(du) or CEP_NR_MOYENS.get(usage)
    cep_moy = CEP_MOYENS.get(du) or CEP_MOYENS.get(usage)

    mcgeo = mccombles = mcsurf_moy = mcsurf_tot = mccat = mchsp = 0.0

    if cepnr_moy is not None:
        if zone in ("H2d", "H3") and is_lc:
            mcgeo = 0.10
        if is_lc and smoy is not None:
            mcsurf_moy = max(-0.10, min(0.10, (70 - smoy) / 70 * 0.10))
        if is_lc:
            mcsurf_tot = max(-0.10, min(0.0, (1500 - sref) / 15000))
        cat_map = MCCAT_ENERGIE.get(du) or MCCAT_ENERGIE.get(usage)
        if cat_map:
            raw = cat_map.get(categorie)
            mccat = (raw / 1000) if raw is not None else 0.0
        if is_lc and hsp > 2.5:
            mchsp = 0.02
        elif is_ml and hsp > 2.5:
            mchsp = 0.04

    sum_mc = mcgeo + mccombles + mcsurf_moy + mcsurf_tot + mccat + mchsp
    cepnr_max = cepnr_moy * (1 + sum_mc) if cepnr_moy is not None else None
    cep_max = cep_moy * (1 + sum_mc) if cep_moy is not None else None

    # ── ICénergie ────────────────────────────────────────────────────────────
    icmap = IC_ENERGIE_MOYENS.get(du) or IC_ENERGIE_MOYENS.get(usage)
    ic_moy = None
    if icmap:
        if not rcu:
            ic_moy = icmap.get("no_rcu", icmap.get("rcu_autre"))
        elif not rcu_classe:
            ic_moy = icmap.get("rcu_autre", icmap.get("no_rcu"))
        else:
            if annee_pc == "2022_2024":
                ic_moy = icmap.get("rcu_classe_2022_2024", icmap.get("rcu_autre"))
            else:
                ic_moy = icmap.get("rcu_classe_2025+", icmap.get("rcu_autre"))
    icenergie_max = ic_moy * (1 + sum_mc) if ic_moy is not None else None

    # ── ICconstruction ────────────────────────────────────────────────────────
    icc_key = get_icc_key(usage)
    icc_table = ICC_MOYENS.get(icc_key)
    icc_moy = icc_table.get(annee_pc) if icc_table else None
    if icc_moy is None and icc_table:
        icc_moy = icc_table.get("2028_2030")
    if icc_moy and igh:
        icc_moy *= 0.95

    misurf_moy = misurf_tot = 0.0
    if is_lc and smoy is not None:
        misurf_moy = max(-0.10, min(0.10, (67 - smoy) / 67 * 0.10))
    if is_lc:
        misurf_tot = max(-0.10, min(0.0, (1500 - sref) / 15000))

    miinfra = iclot2
    mivrd = iclot1
    mipv = -iclot13
    mided = 0.0
    if is_lc and annee_pc == "2025_2027":
        mided = -15.0
    miagrement = 0.0
    if is_lc and sref > 0 and (sagr / sref) > 0.10:
        miagrement = 2.75
    mi_hsp = 7.6 if (is_lc and hsp > 2.5) else 0.0
    miclim_rcu = 25.0 if (rcu and rcu_classe and clim and is_lc) else 0.0

    icc_max = None
    if icc_moy is not None:
        base_part = icc_moy * (1 + misurf_moy + misurf_tot)
        icc_max = base_part + miinfra + mivrd + mipv + mided + miagrement + mi_hsp + miclim_rcu

    # ── DH ───────────────────────────────────────────────────────────────────
    dh_cat = DH_MAX.get(du) or DH_MAX.get(usage)
    dh_max = dh_cat
    dh_effinergie = DH_EFFINERGIE.get(du) or DH_EFFINERGIE.get(usage)

    coeffs = {
        "bbio_moy": bbio_moy, "mbgeo": mbgeo, "mbcombles": mbcombles,
        "mbsurf_moy": mbsurf_moy, "mbsurf_tot": mbsurf_tot,
        "mbbruit": mbbruit, "mbhsp": mbhsp, "sum_mb": sum_mb,
        "cepnr_moy": cepnr_moy, "cep_moy": cep_moy,
        "mcgeo": mcgeo, "mcsurf_moy": mcsurf_moy, "mcsurf_tot": mcsurf_tot,
        "mccat": mccat, "mchsp": mchsp, "sum_mc": sum_mc,
        "ic_moy": ic_moy, "icc_moy": icc_moy,
        "misurf_moy": misurf_moy, "misurf_tot": misurf_tot,
        "miinfra": miinfra, "mivrd": mivrd, "mipv": mipv,
        "mided": mided, "miagrement": miagrement, "mi_hsp": mi_hsp,
        "miclim_rcu": miclim_rcu, "smoy": smoy,
    }

    return {
        "bbio_max": bbio_max, "cepnr_max": cepnr_max, "cep_max": cep_max,
        "icenergie_max": icenergie_max, "icc_max": icc_max,
        "dh_max": dh_max, "dh_effinergie": dh_effinergie,
        "coeffs": coeffs,
    }


# ─── Interface Streamlit ───────────────────────────────────────────────────────

st.title("🏗️ Calculateur Seuils RE2020")
st.caption("Version 2.1 — mars 2026 · Document non officiel basé sur les textes réglementaires RE2020 · GNU GPL v3")

st.warning(
    "⚠️ Outil non officiel réalisé d'après le fichier de Guillaume Meunier. "
    "À utiliser avec toutes les vérifications nécessaires sur legifrance.gouv.fr",
    icon="⚠️",
)

# ── Colonnes principales ──────────────────────────────────────────────────────
col_left, col_right = st.columns(2, gap="large")

with col_left:
    st.subheader("🏢 Bâtiment")

    usage = st.selectbox(
        "Typologie d'usage",
        options=list(USAGE_LABELS.keys()),
        format_func=lambda k: USAGE_LABELS[k],
        index=1,  # Logement collectif
    )

    hotel_detail = "0,1,2 - nuit"
    resto_detail = "continu"
    resto_sco_detail = "1r5j"
    industrie_detail = "3x8"
    enseignement_detail = "primaire"
    sante_detail = "nuit"

    if usage == "Hôtels":
        hotel_detail = st.selectbox("Détail hôtel", [
            "0,1,2 - nuit", "3,4,5 - nuit", "0,1,2 - jour", "3,4,5 - jour"
        ])
    elif usage == "Restaurants":
        resto_map = {
            "En continu, 18h/j, 7j/7": "continu",
            "1 repas/j, 5j/7": "1r5j",
            "2 repas/j, 7j/7": "2r7j",
            "2 repas/j, 6j/7": "2r6j",
        }
        resto_detail = resto_map[st.selectbox("Type de restaurant", list(resto_map.keys()))]
    elif usage == "Restaurants_scolaires":
        rsto_sco_map = {"1 repas/j, 5j/7": "1r5j", "3 repas/j, 5j/7": "3r5j"}
        resto_sco_detail = rsto_sco_map[st.selectbox("Type restaurant scolaire", list(rsto_sco_map.keys()))]
    elif usage == "Industries":
        industrie_detail = st.selectbox("Horaires", {"3×8h": "3x8", "8h à 18h": "8h18"})
        industrie_map = {"3×8h": "3x8", "8h à 18h": "8h18"}
        industrie_detail = industrie_map[st.selectbox("Horaires", list(industrie_map.keys()))]
    elif usage == "Enseignement_primaire_ou_secondaire":
        ens_map = {"Enseignement primaire": "primaire", "Enseignement secondaire": "secondaire"}
        enseignement_detail = ens_map[st.selectbox("Niveau d'enseignement", list(ens_map.keys()))]
    elif usage == "Etablissements_sante":
        sante_map = {"Partie nuit": "nuit", "Partie jour": "jour"}
        sante_detail = sante_map[st.selectbox("Partie de bâtiment", list(sante_map.keys()))]

    dept_list = sorted(ZONES.keys())
    default_dept = dept_list.index("Val-d'Oise") if "Val-d'Oise" in dept_list else 0
    departement = st.selectbox(
        "Département",
        options=dept_list,
        index=default_dept,
        format_func=lambda d: f"{d} ({ZONES[d][1]})",
    )
    zone = ZONES[departement][0]
    st.info(f"🌍 Zone climatique : **{zone}**")

    elev = st.radio("Élévation du site", ["< 400 m", "≥ 400 m"], horizontal=True)
    annee_map = {
        "2022 à 2024": "2022_2024",
        "2025 à 2027": "2025_2027",
        "2028 à 2030": "2028_2030",
        "À partir de 2031": "2031+",
    }
    annee_label = st.selectbox("Année de dépôt du PC", list(annee_map.keys()), index=2)
    annee_pc = annee_map[annee_label]
    igh = st.checkbox("IGH (bâtiment de grande hauteur) ?", value=False)

with col_right:
    st.subheader("📐 Surfaces & Géométrie")

    sref = st.number_input("Surface de référence Sref (m² SHAB)", min_value=1.0, value=4000.0, step=50.0)

    is_logement = usage in ("Maisons_individuelles_ou_accolées", "Logement_collectif")
    is_lc = usage == "Logement_collectif"

    nl, hsp, sagr, bruit, categorie_int, inertie = 1, 2.5, 0.0, "BR1", 1, "lourde"

    if is_logement:
        nl = st.number_input("Nombre de logements (NL)", min_value=1, value=60)
        smoy_display = sref / nl if nl > 0 else 0
        st.caption(f"Surface moyenne des logements : {smoy_display:.1f} m²")

        hsp = st.number_input("Hauteur sous plafond moyenne HSP (m)", min_value=2.0, value=2.7, step=0.1)
        sagr = st.number_input("Surface d'agrément extérieur Sagr (m²)", min_value=0.0, value=0.0, step=10.0)
        if sref > 0:
            st.caption(f"Taux Sagr/Sref = {sagr/sref*100:.1f} %")

        bruit = st.selectbox("Exposition au bruit", ["BR1 — Non exposé", "BR2 — Exposé voie ferrée", "BR3 — Exposé route/aéroport"])
        bruit = bruit.split(" ")[0]

        cat_label = st.selectbox("Catégorie de contraintes extérieures", ["Catégorie 1", "Catégorie 2", "Catégorie 3"])
        categorie_int = int(cat_label.split()[-1])

        inertie = st.radio("Classe d'inertie (DH Effinergie)", ["Lourde ou très lourde", "Très légère, légère ou moyenne"])
        inertie = "lourde" if "Lourde" in inertie else "legere"
    else:
        cat_label = st.selectbox("Catégorie de contraintes extérieures", ["Catégorie 1", "Catégorie 2", "Catégorie 3"])
        categorie_int = int(cat_label.split()[-1])

    st.subheader("⚡ Énergie")
    rcu = st.checkbox("Raccordé à un réseau de chaleur urbain (RCU) ?", value=False)
    rcu_classe = False
    clim = False
    if rcu:
        rcu_classe = st.checkbox("Ce réseau est-il classé ?", value=True)
        if rcu_classe:
            clim = st.checkbox("Climatisé (avec RCU classé) ?", value=True)

    st.subheader("🧱 Matière / Construction")
    c1, c2 = st.columns(2)
    with c1:
        iclot1 = st.number_input("Iclot VRD (kgCO₂e/m²)", value=50.0, step=5.0)
        iclot13 = st.number_input("Iclot photovoltaïque (kgCO₂e/m²)", value=0.0, step=5.0)
    with c2:
        iclot2 = st.number_input("Iclot fondation/infra (kgCO₂e/m²)", value=70.0, step=5.0)
        icded = st.number_input("Impact DED et forfaits (kgCO₂e/m²)", value=300.0, step=10.0)

# ─── Calculs ─────────────────────────────────────────────────────────────────

du = get_detailed_usage(
    usage, hotel_detail, resto_detail, resto_sco_detail,
    industrie_detail, enseignement_detail, sante_detail,
)

results = calculate(
    usage=usage, du=du, zone=zone, sref=sref, nl=nl, hsp=hsp, sagr=sagr,
    bruit=bruit, categorie=categorie_int, inertie=inertie,
    annee_pc=annee_pc, igh=igh, rcu=rcu, rcu_classe=rcu_classe, clim=clim,
    iclot1=iclot1, iclot2=iclot2, iclot13=iclot13, icded=icded,
)

# ─── Affichage des résultats ──────────────────────────────────────────────────

st.divider()
st.subheader("📊 Résultats — Seuils obligatoires RE2020")
st.caption(f"Usage : **{USAGE_LABELS.get(usage, usage)}** · Détail : `{du}` · Zone : **{zone}** · PC déposé : **{annee_label}**")


def fmt(v, decimals=1):
    if v is None:
        return "—"
    return f"{v:.{decimals}f}"


mc1, mc2, mc3, mc4, mc5, mc6 = st.columns(6)

with mc1:
    st.metric("Bbio max", fmt(results["bbio_max"]) + " pts")
with mc2:
    st.metric("Cep,nr max", fmt(results["cepnr_max"]) + " kWh EP/(m²·an)")
with mc3:
    st.metric("Cep max", fmt(results["cep_max"]) + " kWh EP/(m²·an)")
with mc4:
    st.metric("Icénergie max", fmt(results["icenergie_max"]) + " kgCO₂e/(m²·50ans)")
with mc5:
    st.metric("Icconstruction max", fmt(results["icc_max"]) + " kgCO₂e/m²", help="Sur 50 ans")
with mc6:
    dh_str = fmt(results["dh_max"], 0) + " °C·h"
    if results["dh_effinergie"]:
        dh_str += f"\n(Effinergie : {results['dh_effinergie']} °C·h)"
    st.metric("DH max", fmt(results["dh_max"], 0) + " °C·h",
              delta=f"Effinergie : {results['dh_effinergie']} °C·h" if results["dh_effinergie"] else None,
              delta_color="off")

# ─── Détail des coefficients ──────────────────────────────────────────────────

c = results["coeffs"]

with st.expander("🔍 Détail des coefficients de modulation"):
    d1, d2, d3 = st.columns(3)

    with d1:
        st.markdown("**Bioclimatisme (Bbio)**")
        st.table({
            "Coefficient": ["Bbio_maxmoyen", "Mbgéo", "Mbsurf_moy", "Mbsurf_tot", "Mbbruit", "MbHSP", "**Σ Mb**"],
            "Valeur": [
                f"{c['bbio_moy']}",
                f"{c['mbgeo']*100:.2f} %",
                f"{c['mbsurf_moy']*100:.2f} %",
                f"{c['mbsurf_tot']*100:.2f} %",
                f"{c['mbbruit']*100:.2f} %",
                f"{c['mbhsp']*100:.2f} %",
                f"**{c['sum_mb']*100:.2f} %**",
            ],
        })
        if results["bbio_max"]:
            st.markdown(f"→ **Bbio_max = {c['bbio_moy']} × (1 + {c['sum_mb']:.4f}) = {results['bbio_max']:.2f}**")

    with d2:
        st.markdown("**Énergie (Cep, Icénergie)**")
        st.table({
            "Coefficient": ["Cep,nr_maxmoyen", "Cep_maxmoyen", "Icénergie_moy", "Mcgéo", "Mcsurf_moy", "Mcsurf_tot", "Mccat", "McHSP", "**Σ Mc**"],
            "Valeur": [
                f"{c['cepnr_moy']}",
                f"{c['cep_moy']}",
                f"{c['ic_moy']}",
                f"{c['mcgeo']*100:.2f} %",
                f"{c['mcsurf_moy']*100:.2f} %",
                f"{c['mcsurf_tot']*100:.2f} %",
                f"{c['mccat']*100:.2f} %",
                f"{c['mchsp']*100:.2f} %",
                f"**{c['sum_mc']*100:.2f} %**",
            ],
        })

    with d3:
        st.markdown("**Matière (Icconstruction)**")
        st.table({
            "Coefficient": ["Icc_maxmoyen", "Misurf_moy", "Misurf_tot", "Miinfra (Iclot2)", "Mivrd (Iclot1)", "Mipv", "Mided", "Miagrément", "Mi_HSP", "Miclim_RCU"],
            "Valeur": [
                f"{c['icc_moy']:.2f}" if c['icc_moy'] else "—",
                f"{c['misurf_moy']*100:.2f} %",
                f"{c['misurf_tot']*100:.2f} %",
                f"+{c['miinfra']:.0f}",
                f"+{c['mivrd']:.0f}",
                f"{c['mipv']:.0f}",
                f"{c['mided']:.0f}",
                f"+{c['miagrement']:.2f}",
                f"+{c['mi_hsp']:.1f}",
                f"+{c['miclim_rcu']:.0f}",
            ],
        })
        if results["icc_max"]:
            st.markdown(
                f"→ **Icc_max = {c['icc_moy']:.2f} × (1 + {c['misurf_moy']:.4f} + {c['misurf_tot']:.4f})"
                f" + {c['miinfra']:.0f} + {c['mivrd']:.0f} {c['mipv']:.0f}"
                f" + {c['mided']:.0f} + {c['miagrement']:.2f} + {c['mi_hsp']:.1f} + {c['miclim_rcu']:.0f}"
                f" = **{results['icc_max']:.2f}**"
            )

st.divider()
st.caption(
    "Réalisé d'après le fichier Excel de Guillaume Meunier (v2.1). "
    "Document non officiel — partagé sous licence GNU GPL v3. "
    "Vérifiez toujours les résultats avec les textes officiels sur legifrance.gouv.fr"
)
