---
name: presentation
description: Améliorer la présentation visuelle des pages HTML de la boutique et des fiches produit (typographie, espacements, hiérarchie, couleurs, responsive mobile/iPhone). À utiliser dès qu'on parle de design, mise en page, look ou rendu.
---

# Améliorer la présentation

Fichiers concernés : `boutique/index.html`, `fiches/` et `fiches/ecommerce/`.

## Contraintes du projet
- Garder le fond clair, même en mode sombre.
- Les photos ne doivent jamais déborder sur le texte (vérifier Safari et Chrome iOS).
- Le bandeau défilant avec le téléphone reste au-dessus de l'en-tête.
- Ne pas changer les textes, prix ni références produit sans demande.

## Méthode
1. Lire le fichier et repérer les problèmes : hiérarchie floue, espacements irréguliers, textes trop petits, contrastes faibles.
2. Proposer des changements précis avant de tout réécrire. Rester cohérent avec la charte existante.
3. Appliquer des modifications ciblées dans le CSS plutôt qu'une refonte du HTML.

## Règles de style
- Typographie : 2 polices maximum, titres bien distincts du corps, corps à 16 px minimum.
- Espacements : une échelle unique (8, 16, 24, 32, 48 px), plus d'air entre les sections.
- Couleurs : une couleur d'accent, des contrastes lisibles (texte foncé sur fond clair).
- Cartes produit : photos de même ratio, prix en évidence, bouton d'action clair.
- Mobile d'abord : tester à 375 px de large, cibles tactiles de 44 px minimum, aucun défilement horizontal.

## Vérification
Après chaque changement, ouvrir la page et contrôler le rendu à 375 px et à 1280 px.
