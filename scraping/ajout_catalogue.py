#!/usr/bin/env python3
"""Ajoute les fiches récupérées sur Arrosage Distribution dans catalogue/catalogue.xlsx.

- Prix HT = prix TTC du site / 1,20 (TVA métropole) ; la boutique ajoute ensuite sa TVA.
- Les déclinaisons d'un même produit (même gamme) forment une « Famille » à choix.
- Photos : fichiers <référence>_1.jpg déjà copiés dans boutique/images.
- Les références déjà présentes dans le catalogue sont ignorées (relance sans doublon).

Usage : python3 ajout_catalogue.py [--filtre TEXTE] [--categorie "Électrovannes"] sortie/fichier.jsonl [...]
  --filtre     ne garde que les fiches dont la rubrique du site contient TEXTE (ex. lectrovanne)
  --categorie  range toutes ces fiches dans cette catégorie de la boutique
Ensuite : cd ../catalogue && npm run export   (ou double-clic sur exporter.bat)
"""
import json
import re
import sys
from copy import copy
from pathlib import Path

import openpyxl

ICI = Path(__file__).resolve().parent
CATALOGUE = ICI.parent / "catalogue" / "catalogue.xlsx"
IMAGES = ICI.parent / "boutique" / "images"
TVA_SITE = 0.20

# Sous-catégorie du site -> catégorie de la boutique
CATEGORIES = {
    "Raccords": "Raccords",
    "Régulateur de pression": "Accessoires",
    "Goutteurs": "Goutte-à-goutte",
    "Micro-asperseurs": "Goutte-à-goutte",
    "Micro-tubes et piques de maintien": "Goutte-à-goutte",
    "Outils de pose": "Goutte-à-goutte",
    "Turbines": "Arroseurs",
    "Tuyères": "Arroseurs",
    "Accessoires": "Accessoires",
}


def categorie(d):
    derniere = d["categorie"].split(" > ")[-1]
    return CATEGORIES.get(derniere, "Goutte-à-goutte" if "goutte" in d["url"] else "Arroseurs")


def description(d, maxi=700):
    t = d["description"] or d["description_courte"]
    t = re.sub(r"\n+", " ", t).strip()
    if len(t) > maxi:
        coupe = t[:maxi]
        t = coupe[: coupe.rfind(". ") + 1] or coupe.rsplit(" ", 1)[0] + "…"
    return t


def techniques(d):
    lignes = []
    for k, v in d["caracteristiques"].items():
        v = v.replace("\n", ", ")
        if len(v) > 150:
            v = v[:147].rsplit(" ", 1)[0] + "…"
        lignes.append(f"{k} : {v}")
    return "\n".join(lignes)


def choix(noms):
    """Partie distinctive de chaque nom (le début commun est retiré)."""
    if len(noms) < 2:
        return noms
    pref = noms[0]
    for n in noms[1:]:
        while not n.startswith(pref):
            pref = pref[:-1]
    pref = pref[: pref.rfind(" ") + 1] if " " in pref else ""
    res = [n[len(pref):].strip(" -–,") or n for n in noms]
    return res if len(set(res)) == len(res) else noms


def main(fichiers, filtre=None, forcee=None):
    fiches = {}
    for f in fichiers:
        for ligne in open(f, encoding="utf-8"):
            d = json.loads(ligne)
            if filtre and filtre.lower() not in d["categorie"].lower():
                continue
            fiches[d["reference"]] = d
    global categorie
    if forcee:
        categorie = lambda d: forcee

    wb = openpyxl.load_workbook(CATALOGUE)
    ws = wb["Produits"]
    titres = {str(c.value or "").split(" (")[0].strip().lower(): c.column for c in ws[1]}
    col = lambda debut: next(v for k, v in titres.items() if k.startswith(debut))
    C = {k: col(k) for k in ["visible", "catégorie", "nom du produit", "référence", "prix ht",
                             "prix ttc", "stock", "photo", "description", "caractéristiques techniques",
                             "famille", "choix 1"]}

    existantes = {str(ws.cell(r, C["référence"]).value or "").lower() for r in range(2, ws.max_row + 1)}
    derniere = max(r for r in range(1, ws.max_row + 1)
                   if any(ws.cell(r, c).value not in (None, "") for c in (C["nom du produit"], C["référence"], C["famille"])))
    modele = derniere

    # Regroupement par gamme
    groupes = {}
    for d in sorted(fiches.values(), key=lambda d: (categorie(d), d["produit"], d["nom"])):
        if d["reference"].lower() in existantes:
            continue
        groupes.setdefault((categorie(d), d["produit"] or d["nom"]), []).append(d)

    r = derniere
    ajoutes = familles = sans_photo = 0
    for (cat, gamme), lot in groupes.items():
        c1 = choix([d["nom"] for d in lot])
        for i, d in enumerate(lot):
            r += 1
            for c in range(1, ws.max_column + 1):  # même mise en forme que la ligne modèle
                src, dst = ws.cell(modele, c), ws.cell(r, c)
                if src.has_style:
                    dst._style = copy(src._style)
            try:
                ht = round(float(d["prix_ttc"]) / (1 + TVA_SITE), 2)
            except (TypeError, ValueError):
                ht = None
            photo = f"{re.sub(r'[^A-Za-z0-9.-]+', '_', d['reference'])}_1.jpg"
            if not (IMAGES / photo).exists():
                photo, sans_photo = None, sans_photo + 1
            premiere = i == 0
            ws.cell(r, C["visible"], "oui")
            ws.cell(r, C["catégorie"], cat)
            ws.cell(r, C["nom du produit"], d["nom"])
            ws.cell(r, C["référence"], d["reference"])
            ws.cell(r, C["prix ht"], ht)
            ws.cell(r, C["prix ttc"], f'=IF(E{r}="","",ROUND(E{r}*(1+Réglages!$B$2),2))')
            ws.cell(r, C["stock"], "En stock" if d["disponibilite"] == "InStock" else "Sous 5 jours")
            ws.cell(r, C["photo"], photo)
            if premiere:
                ws.cell(r, C["description"], description(d))
                ws.cell(r, C["caractéristiques techniques"], techniques(d))
            if len(lot) > 1:
                ws.cell(r, C["famille"], gamme)
                ws.cell(r, C["choix 1"], c1[i])
            ajoutes += 1
        familles += len(lot) > 1

    # Listes déroulantes et contrôles : étendus aux nouvelles lignes
    for dv in ws.data_validations.dataValidation:
        for rg in list(dv.sqref.ranges):
            if rg.max_row < r:
                dv.sqref.remove(rg)
                dv.sqref.add(f"{rg.coord.split(':')[0]}:{rg.coord.split(':')[-1].rstrip('0123456789')}{r}")
    wb.save(CATALOGUE)
    print(f"{ajoutes} produits ajoutés (lignes {derniere + 1} à {r}), dont {familles} fiches à choix ; "
          f"{sans_photo} sans photo.")


if __name__ == "__main__":
    args, opts = [], {}
    it = iter(sys.argv[1:])
    for a in it:
        if a in ("--filtre", "--categorie"):
            opts[a[2:]] = next(it)
        else:
            args.append(a)
    if not args:
        sys.exit(__doc__)
    main(args, opts.get("filtre"), opts.get("categorie"))
