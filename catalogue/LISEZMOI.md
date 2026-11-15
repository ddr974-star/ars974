# Catalogue de la boutique

Les produits, les prix et le stock se saisissent dans un tableur : `catalogue.xlsx`.

## Mettre à jour la boutique

1. Ouvrez `catalogue.xlsx` et modifiez l'onglet **Produits** (voir l'onglet « Mode d'emploi »).
2. Enregistrez (Ctrl+S) et **fermez** Excel.
3. Double-cliquez sur `exporter.bat`. Le fichier `boutique/produits.js` est recréé.
4. Rechargez la boutique. Si une erreur s'affiche, elle indique la ligne à corriger : rien n'est publié tant qu'elle n'est pas résolue.

Tableur : Excel, LibreOffice Calc (gratuit) ou tout logiciel qui enregistre en .xlsx. Gardez le format .xlsx à l'enregistrement.

Prérequis : Node.js installé sur l'ordinateur. Au premier lancement, `exporter.bat` installe seul ce qui manque.

## Ce que contient le classeur

| Onglet | Rôle |
|---|---|
| Produits | Une ligne par produit : catégorie, nom, référence, prix HT, stock, photo, description, caractéristiques. |
| Réglages | TVA, seuil de livraison offerte, frais de livraison. |
| Mode d'emploi | Rappel des étapes. |

## Photos

Trois façons d'ajouter une photo, au choix pour chaque produit :

1. **La coller dans le tableur**, sur la ligne du produit, dans la colonne Photo (LibreOffice : Insertion > Image, puis clic droit > Ancrer > À la cellule). L'export l'enregistre dans `boutique/images` sous le nom de la référence (par exemple `50pcja2s00c.jpg`).
2. **Écrire le nom d'un fichier** déjà présent dans `boutique/images` dans la colonne Photo.

3. **Déposer le fichier dans `boutique/images` en le nommant avec la référence** du produit (par exemple `50pcja2s00c.jpg`, en minuscules). Rien à saisir dans le tableur : l'export le rattache tout seul. C'est la méthode la plus simple et la plus fiable.

Priorité : photo collée, puis nom écrit dans la colonne Photo, puis fichier nommé d'après la référence. Insérez les photos après avoir fini de trier le tableau : elles ne suivent pas toujours leur ligne si on trie ensuite. Réduisez-les à environ 800 pixels de côté. Les images « dans la cellule » d'Excel 365 ne sont pas lues : utilisez une image ordinaire posée sur la cellule.

## Articles à variantes (plusieurs débits ou tailles)

Donnez le même nom dans la colonne **Famille** à toutes les lignes d'un même article : elles forment une seule fiche avec des boutons de choix (goutteurs NESTOS, CETA, XB-PC…).

- **Choix 1** : le choix proposé (par exemple `2 L/h`). Obligatoire.
- **Couleur du choix 1** : pastille de couleur (noir, vert, rouge, bleu, orange, marron, violet, blanc, beige, gris, jaune, rose, ou un code `#2E9E4F`).
- **Pièces par lot** : affiche le prix à l'unité (« Soit 0,3132 € l'unité »).
- **Points forts** : une ligne par point, sous la forme `Titre : texte`.
- La description, les caractéristiques techniques et les points forts se saisissent sur la **première ligne** de la famille.
- Une ligne sans Famille reste un article simple.

## Mise en ligne (NAS Synology)

Copiez le dossier `boutique` (avec `produits.js` et `images`) dans le dossier de Web Station.
Le dossier `catalogue` reste sur votre ordinateur : il n'a pas à être publié.

## Limites actuelles

- Les catégories sont fixes. Pour en ajouter une, il faut modifier `index.html`.
