#!/usr/bin/env python3
"""Récupère les fiches produits de www.arrosage-distribution.fr.

Étape 1 (collecte) : lit le plan du site (sitemap) puis chaque fiche produit,
et enregistre les données dans un fichier JSON Lines (une ligne par fiche).
La collecte reprend là où elle s'est arrêtée si on la relance.

Étape 2 (images) : télécharge les photos, redimensionnées à 1200 px max.
Étape 3 (export) : construit un classeur Excel avec miniatures intégrées.

Usage :
    python3 scraper_arrosage_distribution.py collecte [--categorie goutte-a-goutte] [--limite N] [--delai 1.0]
    python3 scraper_arrosage_distribution.py images
    python3 scraper_arrosage_distribution.py export
"""
import argparse
import html
import io
import json
import re
import sys
import time
from pathlib import Path

import requests
from bs4 import BeautifulSoup

BASE = "https://www.arrosage-distribution.fr"
SITEMAP = BASE + "/sitemap/products.xml"
ICI = Path(__file__).resolve().parent
SORTIE = ICI / "sortie"
DONNEES = SORTIE / "arrosage_distribution.jsonl"  # redéfini selon la rubrique
MINIATURES = SORTIE / "miniatures"
IMAGES = SORTIE / "images"
TAILLE_MAX = 1200  # px, côté le plus long des images enregistrées
EXCEL = SORTIE / "arrosage_distribution.xlsx"  # redéfini selon la rubrique


def choisit_rubrique(categorie):
    """Un fichier de données et un classeur par rubrique."""
    global DONNEES, EXCEL
    if "," in categorie:
        categorie = "autres_familles"
    if categorie and categorie != "goutte-a-goutte":
        DONNEES = SORTIE / f"arrosage_distribution_{categorie}.jsonl"
        EXCEL = SORTIE / f"arrosage_distribution_{categorie}.xlsx"

session = requests.Session()
session.headers["User-Agent"] = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120 Safari/537.36"
)


def get(url, essais=4):
    for i in range(essais):
        try:
            r = session.get(url, timeout=30)
            if r.status_code == 200:
                return r
            if r.status_code == 404:
                return None
        except requests.RequestException:
            pass
        time.sleep(2 * (i + 1))
    return None


def nettoie(txt):
    if txt is None:
        return ""
    txt = html.unescape(html.unescape(str(txt)))
    txt = re.sub(r"<br\s*/?>|</p>|</li>|</h\d>", "\n", txt)
    txt = re.sub(r"<[^>]+>", " ", txt)
    txt = re.sub(r"[ \t\xa0]+", " ", txt)
    txt = re.sub(r"\s*\n\s*", "\n", txt)
    return txt.strip()


def urls_sitemap():
    r = get(SITEMAP)
    if r is None:
        sys.exit("Impossible de lire le sitemap")
    return re.findall(r"<loc>([^<]+)</loc>", r.text)


def urls_categorie(categorie, delai):
    """Liste les fiches d'une rubrique en parcourant ses pages (?page=2, 3…)."""
    urls, page = [], 1
    while True:
        r = get(f"{BASE}/category/{categorie}?page={page}")
        if r is None:
            break
        nouvelles = [BASE + h if h.startswith("/") else h
                     for h in re.findall(r'href="((?:https://www\.arrosage-distribution\.fr)?/products/[^"?#]+)"', r.text)]
        nouvelles = [u for u in dict.fromkeys(nouvelles) if u not in urls]
        if not nouvelles:
            break
        urls += nouvelles
        print(f"  page {page} : {len(nouvelles)} fiches", flush=True)
        page += 1
        time.sleep(delai)
    return urls


def lit_jsonld(soup):
    blocs = []
    for s in soup.find_all("script", type="application/ld+json"):
        try:
            blocs.append(json.loads(s.string or "", strict=False))
        except (json.JSONDecodeError, TypeError):
            continue
    return blocs


def analyse(url, page):
    soup = BeautifulSoup(page, "html.parser")
    ld = lit_jsonld(soup)
    produit = next((b for b in ld if b.get("@type") in ("Product", "ProductGroup")), {})
    fil = next((b for b in ld if b.get("@type") == "BreadcrumbList"), {})
    chemin = [e.get("name", "") for e in fil.get("itemListElement", [])][1:-1]

    # Variante correspondant à cette URL (les fiches à déclinaisons
    # partagent la même page groupe).
    variante = soup.find(attrs={"data-link": url})
    sku = variante.get("data-code") if variante else produit.get("sku", "")
    nom = variante.get("data-name") if variante else produit.get("name", "")
    court = variante.get("data-short-description") if variante else ""
    if not court:
        bloc = soup.find(id="sylius-product-short-description")
        court = bloc.decode_contents() if bloc else ""

    prix, dispo = "", ""
    offres = [produit] + produit.get("hasVariant", [])
    for o in offres:
        if o.get("sku") == sku or o is produit:
            off = o.get("offers", {})
            if off.get("price"):
                prix = off.get("price")
                dispo = off.get("availability", "").rsplit("/", 1)[-1]
    if not prix:
        off = produit.get("offers", {})
        prix = off.get("price") or off.get("lowPrice", "")

    desc = soup.select_one("details#description")
    if desc and desc.summary:
        desc.summary.decompose()
    caract = {}
    for tr in soup.select("#sylius-product-attributes tr"):
        tds = tr.find_all("td")
        if len(tds) == 2:
            caract[nettoie(tds[0].get_text())] = nettoie(tds[1].get_text())

    images = []
    for a in soup.select('a[href*="/media/cache/sylius_shop_product_original/"]'):
        if a["href"] not in images:
            images.append(a["href"])
    if not images and produit.get("image"):
        images.append(produit["image"])

    pdfs = sorted({a["href"] for a in soup.select('a[href$=".pdf"]')})

    return {
        "url": url,
        "reference": sku,
        "nom": nettoie(nom),
        "produit": nettoie(produit.get("name", "")),
        "marque": (produit.get("brand") or {}).get("name", "") or caract.get("Marque", ""),
        "categorie": " > ".join(nettoie(c) for c in chemin),
        "prix_ttc": prix,
        "disponibilite": dispo,
        "description_courte": nettoie(court),
        "description": nettoie(desc.decode_contents()) if desc else nettoie(produit.get("description")),
        "caracteristiques": caract,
        "images": images,
        "documents_pdf": pdfs,
    }


def collecte(limite, delai, categorie):
    """categorie : une rubrique, plusieurs séparées par des virgules, ou "" (tout le site).

    Anti-doublon : toute fiche déjà présente dans un fichier .jsonl de sortie/
    (même adresse ou même référence) est ignorée.
    """
    SORTIE.mkdir(exist_ok=True)
    faites, refs = set(), set()
    for fichier in SORTIE.glob("*.jsonl"):
        for ligne in fichier.open(encoding="utf-8"):
            d = json.loads(ligne)
            faites.add(d["url"])
            refs.add(d["reference"])
    liste = []  # (url, famille)
    for cat in [c for c in categorie.split(",") if c] or [""]:
        if cat:
            print(f"Rubrique {cat} :", flush=True)
            liste += [(u, cat) for u in urls_categorie(cat, delai)]
        else:
            liste += [(u, "") for u in urls_sitemap()]
    vues, urls = set(), []
    for u, fam in liste:  # un produit présent dans deux rubriques n'est pris qu'une fois
        if u not in faites and u not in vues:
            vues.add(u)
            urls.append((u, fam))
    if limite:
        urls = urls[:limite]
    print(f"{len(faites)} fiches déjà récupérées, {len(urls)} à faire", flush=True)
    doublons = 0
    with DONNEES.open("a", encoding="utf-8") as f:
        for i, (url, fam) in enumerate(urls, 1):
            r = get(url)
            if r is not None:
                try:
                    d = analyse(url, r.text)
                    if d["reference"] and d["reference"] in refs:
                        doublons += 1
                    else:
                        refs.add(d["reference"])
                        d["famille"] = fam
                        f.write(json.dumps(d, ensure_ascii=False) + "\n")
                        f.flush()
                except Exception as e:  # une fiche mal formée ne bloque pas le reste
                    print(f"  erreur sur {url} : {e}", flush=True)
            else:
                print(f"  échec : {url}", flush=True)
            if i % 50 == 0:
                print(f"{i}/{len(urls)}", flush=True)
            time.sleep(delai)
    print(f"Terminé. {doublons} doublons de référence ignorés.", flush=True)


def miniature(url_image, delai):
    """Télécharge la petite vignette du site (pour l'affichage dans Excel)."""
    MINIATURES.mkdir(parents=True, exist_ok=True)
    petit = url_image.replace("sylius_shop_product_original", "sylius_shop_product_small_thumbnail")
    petit = petit.replace("sylius_shop_product_large_thumbnail", "sylius_shop_product_small_thumbnail")
    nom = re.sub(r"[^A-Za-z0-9]+", "_", petit.split("/media/cache/")[-1])
    chemin = MINIATURES / nom
    if not chemin.exists():
        r = get(petit)
        time.sleep(delai)
        if r is None:
            return None
        chemin.write_bytes(r.content)
    return chemin


def nom_image(d, i):
    return f"{re.sub(r'[^A-Za-z0-9.-]+', '_', d['reference'])}_{i}.jpg"


def images(delai):
    """Télécharge les images d'origine, redimensionnées (JPEG fond blanc)."""
    from PIL import Image as PILImage
    IMAGES.mkdir(parents=True, exist_ok=True)
    n = 0
    for ligne in DONNEES.open(encoding="utf-8"):
        d = json.loads(ligne)
        for i, url in enumerate(d["images"], 1):
            cible = IMAGES / nom_image(d, i)
            if cible.exists():
                continue
            r = get(url)
            time.sleep(delai)
            if r is None:
                print(f"  échec image : {url}", flush=True)
                continue
            im = PILImage.open(io.BytesIO(r.content))
            if im.mode in ("RGBA", "LA", "P"):
                im = im.convert("RGBA")
                fond = PILImage.new("RGB", im.size, "white")
                fond.paste(im, mask=im.split()[-1])
                im = fond
            im = im.convert("RGB")
            im.thumbnail((TAILLE_MAX, TAILLE_MAX))
            im.save(cible, "JPEG", quality=88, optimize=True)
            n += 1
    print(f"{n} images enregistrées dans {IMAGES}")


def export(delai):
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Alignment, Font, PatternFill
    from PIL import Image as PILImage

    fiches = {}
    for ligne in DONNEES.open(encoding="utf-8"):
        d = json.loads(ligne)
        fiches[d["url"]] = d
    fiches = sorted(fiches.values(), key=lambda d: (d["categorie"], d["produit"], d["nom"]))

    wb = Workbook()
    ws = wb.active
    ws.title = "Produits"
    entetes = ["Photo", "Référence", "Nom", "Produit (gamme)", "Marque", "Catégorie",
               "Prix TTC (€)", "Disponibilité", "Description courte", "Description",
               "Caractéristiques", "Fichiers images (dossier images)", "URL images",
               "Documents PDF", "URL fiche"]
    largeurs = [14, 14, 40, 32, 14, 40, 11, 12, 40, 60, 50, 28, 50, 40, 50]
    ws.append(entetes)
    for i, l in enumerate(largeurs):
        ws.column_dimensions[chr(65 + i)].width = l
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2E7D32")
        c.alignment = Alignment(vertical="center", wrap_text=True)
    ws.freeze_panes = "C2"

    wc = wb.create_sheet("Caractéristiques")
    wc.append(["Référence", "Nom", "Caractéristique", "Valeur"])
    for c in wc[1]:
        c.font = Font(bold=True)
    for col, l in zip("ABCD", [14, 40, 30, 50]):
        wc.column_dimensions[col].width = l

    for n, d in enumerate(fiches, start=2):
        try:
            prix = float(d["prix_ttc"])
        except (TypeError, ValueError):
            prix = d["prix_ttc"]
        ws.append([
            "", d["reference"], d["nom"], d["produit"], d["marque"], d["categorie"], prix,
            d["disponibilite"], d["description_courte"], d["description"][:32000],
            "\n".join(f"{k} : {v}" for k, v in d["caracteristiques"].items()),
            "\n".join(nom_image(d, i) for i in range(1, len(d["images"]) + 1)),
            "\n".join(d["images"]),
            "\n".join(d["documents_pdf"]), d["url"],
        ])
        ws.row_dimensions[n].height = 78
        for c in ws[n]:
            c.alignment = Alignment(vertical="top", wrap_text=True)
        ws.cell(n, 15).hyperlink = d["url"]
        for k, v in d["caracteristiques"].items():
            wc.append([d["reference"], d["nom"], k, v])
        if d["images"]:
            chemin = IMAGES / nom_image(d, 1)
            if not chemin.exists():
                chemin = miniature(d["images"][0], delai)
            if chemin:
                try:
                    im = PILImage.open(chemin).convert("RGB")
                    im.thumbnail((96, 96))
                    tampon = io.BytesIO()
                    im.save(tampon, "PNG")
                    img = XLImage(tampon)
                    ws.add_image(img, f"A{n}")
                except Exception:
                    pass
        if n % 200 == 0:
            print(f"export {n - 1}/{len(fiches)}", flush=True)

    ws.auto_filter.ref = f"A1:O{len(fiches) + 1}"
    wb.save(EXCEL)
    print(f"Classeur écrit : {EXCEL} ({len(fiches)} fiches)")


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("etape", choices=["collecte", "images", "export"])
    p.add_argument("--limite", type=int, default=0, help="nombre max de fiches (test)")
    p.add_argument("--delai", type=float, default=1.0, help="pause entre requêtes (s)")
    p.add_argument("--categorie", default="goutte-a-goutte",
                   help="rubrique du site (ex. goutte-a-goutte) ; 'tout' = catalogue complet")
    a = p.parse_args()
    cat = "" if a.categorie == "tout" else a.categorie
    choisit_rubrique(cat or "tout")
    if a.etape == "collecte":
        collecte(a.limite, a.delai, cat)
    elif a.etape == "images":
        images(a.delai)
    else:
        export(a.delai)
