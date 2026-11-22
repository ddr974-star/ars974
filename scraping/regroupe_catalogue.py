#!/usr/bin/env python3
"""Regroupe les déclinaisons d'un même article (tailles, angles, débits…) en une seule fiche à choix.

Remplit les colonnes « Famille » et « Choix 1 » de catalogue/catalogue.xlsx pour les lignes
qui ne sont pas déjà dans une famille. Deux lignes sont regroupées quand leur nom ne diffère
que par des mesures (16 / 20, 90° / 180°, ½'' / ¾''…) et qu'elles sont dans la même catégorie.

Usage :
    python3 regroupe_catalogue.py            # aperçu, ne modifie rien
    python3 regroupe_catalogue.py --ecrire   # écrit dans catalogue.xlsx
Ensuite : cd ../catalogue && npm run export   (ou double-clic sur exporter.bat)
"""
import re
import sys
from pathlib import Path

import openpyxl

CATALOGUE = Path(__file__).resolve().parent.parent / "catalogue" / "catalogue.xlsx"

# Modèles différents qui se ressemblent : on ne les regroupe pas (mots du nom, en minuscules)
NE_PAS_REGROUPER = ["turbine 3504", "turbine 8005", "eagle 900", "eagle 950", "g880", "g885",
                    "turbine t5", "turbine t7", "série 552", "série 752", "i 25 ultra", "i 40 ultra", "0 à 60 l/h"]

# Nom de famille imposé quand le nom automatique serait maladroit (début du nom -> nom de famille)
NOMS = {
    "collier de prise en charge": "Collier de prise en charge série pro",
    "coude articulé acme": "Coude articulé ACME × BSP (30 cm)",
    "goutteur 0,33 en ligne dripline ø16": "Goutteur 0,33 L/h en ligne Dripline Ø16 noir",
    "goutteur 0.33 en ligne dripline ø 16": "Goutteur 0,33 L/h en ligne Dripline Ø16 marron",
    "régulateur de pression pré-taré": "Régulateur de pression pré-taré ¾'' Rain Bird",
    "régulateur de pression pn15": "Régulateur de pression PN15 / PN25",
    "micro asperseur rotor spray": "Micro asperseur Rotor spray 360°",
    "nez de robinet": "Nez de robinet femelle Claber",
    "régulateur de pression pn25": "Régulateur de pression PN25 avec union",
    "tuyau goutteurs intégrés hydropc": "Tuyau goutteurs intégrés HydroPC 2,2 L/h Rivulis",
    "adaptateur acme": "Adaptateur ACME × BSP Rain Bird",
    "câble multiconducteur 0,8mm x": "Câble multiconducteur 0,8 mm² – 75 m",
    "câble multiconducteur 0,8mm²": "Câble multiconducteur 0,8 mm²",
    "richdel": "Electrovanne Richdel Irritrol",
    "pga": "Electrovanne PGA Rain Bird",
    "pgv-1": "Electrovanne PGV Hunter",
    "clarinette dura": "Clarinette DURA 1''",
}

MESURE = r"\d+(?:[.,]\d+)?(?:\s*(?:mm|m3/h|l/h|bars?|cm|m)(?![a-zà-ü]))?|[½¼¾]|''|\""
UNITES = {"mm", "l/h", "m", "bar", "bars", "cm", "sorties", "sortie", "m3/h", "conducteurs", "mètres",
          "vannes", "stations", "voies"}
mesure = lambda w: bool(re.search(r"\d|[½¼¾]", w))


def cle(nom):
    n = nom.lower()
    n = re.sub(MESURE, "#", n)
    n = re.sub(r"(#\s*([x×/-]|de|à)?\s*)+", "# ", n)
    return re.sub(r"\s+", " ", n).strip()


def mots(t):
    return t.split(" ")


def decoupe(noms):
    """Retire le début et la fin communs (mot par mot) : reste la partie qui change."""
    m = [mots(n) for n in noms]
    debut = 0
    while all(len(x) > debut for x in m) and len({x[debut].lower() for x in m}) == 1:
        debut += 1
    fraction = lambda w: bool(re.match(r"^(\d+/\d+|[½¼¾])", w))
    # le début commun ne coupe pas une mesure en deux (1" | 1/2…)
    while debut and mesure(m[0][debut - 1]) and any(len(x) > debut and fraction(x[debut]) for x in m):
        debut -= 1
    # un nom entièrement compris dans le début commun (« Mamelon 1" » / « Mamelon 1" x 3/4" ») :
    # la dernière mesure redevient un choix
    while debut and mesure(m[0][debut - 1]) and any(len(x) == debut for x in m):
        debut -= 1
    fin = 0
    while all(len(x) > debut + fin for x in m) and len({x[-1 - fin].lower() for x in m}) == 1:
        fin += 1
    while fin and m[0][len(m[0]) - fin].lower() in UNITES:
        fin -= 1  # l'unité reste dans le choix : « 40 l/h », « 2 sorties »
    choix = [" ".join(x[debut:len(x) - fin]).strip(" -–,") or n for x, n in zip(m, noms)]
    nom = " ".join(m[0][:debut] + m[0][len(m[0]) - fin:]).strip(" -–,Ø")
    nom = nom[:1].upper() + nom[1:]
    return nom, choix


def ordre(choix):
    """Tri naturel : 16 avant 20, 90° avant 180°, ½'' avant ¾'' avant 1'' avant 1''1/4."""
    t = choix.replace("½", " 1/2").replace("¼", " 1/4").replace("¾", " 3/4").replace(",", ".")
    t = t.replace("''", '"')
    vals = []
    for m in re.finditer(r'(\d+)\s*"\s*(\d+)/(\d+)|(\d+)/(\d+)|(\d+(?:\.\d+)?)', t):
        if m.group(1):
            vals.append(int(m.group(1)) + int(m.group(2)) / int(m.group(3)))
        elif m.group(4):
            vals.append(int(m.group(4)) / int(m.group(5)))
        else:
            vals.append(float(m.group(6)))
    return vals or [0]


def description(desc, choix_tous):
    """Description commune : on retire les phrases propres à une seule déclinaison."""
    if not desc:
        return desc
    phrases = re.split(r"(?<=[.!?])\s+", desc)
    gardees = [p for p in phrases if not any(c and c in p for c in choix_tous)]
    debut = "Disponible en " + ", ".join(choix_tous[:-1]) + " et " + choix_tous[-1] + "."
    return (debut + " " + " ".join(gardees)).strip()


def techniques(spec, choix_tous):
    if not spec:
        return spec
    return "\n".join(l for l in spec.split("\n") if not any(c and c in l for c in choix_tous))


def main(ecrire):
    wb = openpyxl.load_workbook(CATALOGUE)
    ws = wb["Produits"]
    t = {str(c.value or "").split(" (")[0].strip().lower(): c.column for c in ws[1]}
    col = lambda debut: next(v for k, v in t.items() if k.startswith(debut))
    C_CAT, C_NOM, C_FAM, C_CH1 = col("catégorie"), col("nom du produit"), col("famille"), col("choix 1")
    C_DESC, C_TECH, C_VIS = col("description"), col("caractéristiques techniques"), col("visible")
    C_PRIX_TTC, C_PRIX_HT = col("prix ttc"), col("prix ht")
    ncol = ws.max_column

    groupes = {}
    for r in range(2, ws.max_row + 1):
        nom = ws.cell(r, C_NOM).value
        if not nom or ws.cell(r, C_FAM).value or str(ws.cell(r, C_VIS).value).lower() == "non":
            continue
        if any(x in nom.lower() for x in NE_PAS_REGROUPER):
            continue
        groupes.setdefault((ws.cell(r, C_CAT).value, cle(nom)), []).append(r)
    groupes = [v for v in groupes.values() if len(v) > 1]

    # Familles portant le même nom imposé : fusionnées (ex. les deux séries de colliers)
    fusion = {}
    for lignes in groupes:
        noms = [ws.cell(r, C_NOM).value for r in lignes]
        auto, _ = decoupe(noms)
        nom_fam = next((v for k, v in NOMS.items() if k in noms[0].lower()), auto)
        fusion.setdefault((ws.cell(lignes[0], C_CAT).value, nom_fam), []).extend(lignes)

    total = 0
    for (cat, nom_fam), lignes in sorted(fusion.items()):
        lignes = sorted(lignes)
        noms = [ws.cell(r, C_NOM).value for r in lignes]
        _, choix = decoupe(noms)
        if len(set(choix)) < len(choix):  # choix en double : on garde le nom complet
            choix = noms
        # On range les lignes dans l'ordre naturel des choix, dans les mêmes emplacements
        valeurs = [[ws.cell(r, c).value for c in range(1, ncol + 1)] for r in lignes]
        tri = sorted(range(len(lignes)), key=lambda i: ordre(choix[i]))
        valeurs = [valeurs[i] for i in tri]
        choix = [choix[i] for i in tri]
        desc = next((v[C_DESC - 1] for v in valeurs if v[C_DESC - 1]), None)
        tech = next((v[C_TECH - 1] for v in valeurs if v[C_TECH - 1]), None)
        print(f"{cat:16} {nom_fam}  →  {' · '.join(choix)}")
        total += len(lignes)
        if not ecrire:
            continue
        for j, (r, v) in enumerate(zip(lignes, valeurs)):
            for c in range(1, ncol + 1):
                ws.cell(r, c).value = v[c - 1]
            lettre = openpyxl.utils.get_column_letter(C_PRIX_HT)
            ws.cell(r, C_PRIX_TTC).value = f'=IF({lettre}{r}="","",ROUND({lettre}{r}*(1+Réglages!$B$2),2))'
            ws.cell(r, C_FAM).value = nom_fam
            ws.cell(r, C_CH1).value = choix[j]
            ws.cell(r, C_DESC).value = description(desc, choix) if j == 0 else None
            ws.cell(r, C_TECH).value = techniques(tech, choix) if j == 0 else None
    print(f"\n{len(fusion)} familles, {total} lignes regroupées.")
    if ecrire:
        wb.save(CATALOGUE)
        print("catalogue.xlsx enregistré.")
    else:
        print("Aperçu seulement : relancez avec --ecrire pour enregistrer.")


if __name__ == "__main__":
    main("--ecrire" in sys.argv)
