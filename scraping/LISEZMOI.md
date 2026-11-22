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
