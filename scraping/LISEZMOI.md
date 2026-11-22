# Scraping Arrosage Distribution – rubrique goutte-à-goutte

Script : `scraper_arrosage_distribution.py` (Python 3, `requests`, `beautifulsoup4`, `openpyxl`, `Pillow`).

```
python3 scraper_arrosage_distribution.py collecte   # fiches → sortie/arrosage_distribution.jsonl
python3 scraper_arrosage_distribution.py images     # photos → sortie/images/ (JPEG, 1200 px max)
python3 scraper_arrosage_distribution.py export     # tableau → sortie/arrosage_distribution.xlsx
```

Options : `--categorie goutte-a-goutte` (par défaut ; `tout` = catalogue complet),
`--limite N` (test), `--delai 1.0` (pause entre requêtes, en secondes).
La collecte reprend là où elle s'est arrêtée ; supprimer `sortie/` pour repartir de zéro.

Classeur Excel :
- onglet **Produits** : photo, référence, nom, gamme, marque, catégorie, prix TTC,
  disponibilité, descriptions, caractéristiques, noms des fichiers images, URL ;
- onglet **Caractéristiques** : une ligne par caractéristique (pratique pour filtrer).

**Droits** : les mentions légales du site interdisent toute reproduction sans
autorisation préalable, et les photos proviennent des fabricants (Rain Bird,
Palaplast, Rivulis, Netafim…). Ces données servent à la consultation interne ;
pour publier les photos sur la boutique, obtenir l'accord d'Arrosage Distribution
ou récupérer les visuels auprès des fabricants / du grossiste.

## Rubriques déjà récupérées

| Rubrique | Option `--categorie` | Fiches | Classeur |
|---|---|---|---|
| Goutte-à-goutte | `goutte-a-goutte` | 103 | `sortie/arrosage_distribution.xlsx` |
| Turbines et tuyères | `arrroseurs-turbines-tuyeres` | 122 | `sortie/arrosage_distribution_arrroseurs-turbines-tuyeres.xlsx` |
| Toutes les autres familles (166 sous-rubriques) | liste séparée par des virgules | 3 034 | `sortie/arrosage_distribution_autres_familles.xlsx` |

(L'option reprend l'orthographe de l'adresse du site, « arrroseurs » avec 3 r.)
Toutes les photos sont dans `sortie/images/` (1200 px pour les deux premières rubriques, 800 px pour les autres).
Seules celles du goutte-à-goutte et des turbines sont copiées dans `boutique/images/`.

Anti-doublon : la collecte ignore toute fiche déjà présente dans un fichier `.jsonl` de `sortie/`
(même adresse ou même référence), et un produit listé dans plusieurs rubriques n'est pris qu'une fois.
