// Lit catalogue.xlsx et crée ../boutique/produits.js (lu par la boutique).
// Utilisation : double-clic sur exporter.bat, ou "npm run export" dans ce dossier.
const fs = require('fs');
const crypto = require('crypto');
const path = require('path');
const ExcelJS = require('exceljs');

// Les variables CATALOGUE_* ne servent qu'aux essais ; en temps normal on garde les chemins par défaut.
const XLSX = process.env.CATALOGUE_XLSX || path.join(__dirname, 'catalogue.xlsx');
const SORTIE = process.env.CATALOGUE_SORTIE || path.join(__dirname, '..', 'boutique', 'produits.js');
const IMAGES = process.env.CATALOGUE_IMAGES || path.join(__dirname, '..', 'boutique', 'images');

// Libellé affiché dans le tableur -> identifiant utilisé par la boutique
const CATEGORIES = {
  'Tuyaux': 'tuyaux', 'Raccords': 'raccords', 'Goutte-à-goutte': 'goutte', 'Arroseurs': 'arroseurs',
  'Programmateurs': 'programmateurs', 'Électrovannes': 'vannes', 'Filtration': 'filtration',
  'Pompes': 'pompes', 'Accessoires': 'accessoires'
};
const STOCKS = { 'en stock': 'in', 'dernières pièces': 'low', 'sous 5 jours': 'order' };

const erreurs = [];
const avertissements = [];

function texte(cell) {
  const v = cell.value;
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && v.result !== undefined) return String(v.result).trim();
  return String(cell.text !== undefined ? cell.text : v).trim();
}
function nombre(cell) {
  let v = cell.value;
  if (v && typeof v === 'object' && v.result !== undefined) v = v.result;
  if (typeof v === 'number') return v;
  const t = String(v === null || v === undefined ? '' : v).replace(/\s/g, '').replace('€', '').replace(',', '.');
  if (t === '') return null;
  const n = Number(t);
  return Number.isFinite(n) ? n : NaN;
}
const arrondi = (x) => Math.round(x * 100) / 100;

async function main() {
  if (!fs.existsSync(XLSX)) { console.error('Fichier introuvable : ' + XLSX); process.exit(1); }
  const wb = new ExcelJS.Workbook();
  try { await wb.xlsx.readFile(XLSX); }
  catch (e) { console.error('Impossible de lire catalogue.xlsx. Est-il encore ouvert dans Excel ? (' + e.message + ')'); process.exit(1); }

  // Réglages
  const rg = wb.getWorksheet('Réglages');
  if (!rg) { console.error('Onglet « Réglages » introuvable.'); process.exit(1); }
  const tva = nombre(rg.getCell('B2'));
  const franco = nombre(rg.getCell('B3'));
  const frais = nombre(rg.getCell('B4'));
  if (!(tva >= 0 && tva < 1)) erreurs.push('Réglages : la TVA doit être un pourcentage (par exemple 8,5 %).');
  if (!(franco >= 0)) erreurs.push('Réglages : « Livraison offerte à partir de » doit être un nombre.');
  if (!(frais >= 0)) erreurs.push('Réglages : « Frais de livraison » doit être un nombre.');

  // Produits
  const ws = wb.getWorksheet('Produits');
  if (!ws) { console.error('Onglet « Produits » introuvable.'); process.exit(1); }
  // Images collées dans le tableur : rattachées à la ligne où elles sont posées
  const imagesParLigne = {};
  for (const im of ws.getImages()) {
    const media = wb.getImage(im.imageId);
    if (!media || !media.buffer) continue;
    const tl = im.range && im.range.tl ? im.range.tl : {};
    const l0 = tl.nativeRow !== undefined ? tl.nativeRow : tl.row;
    if (l0 === undefined) continue;
    const ext = String(media.extension || 'png').toLowerCase().replace('jpeg', 'jpg');
    (imagesParLigne[Math.floor(l0) + 1] = imagesParLigne[Math.floor(l0) + 1] || []).push({ buffer: Buffer.from(media.buffer), ext: ext });
  }
  // Photos déposées dans boutique/images avec la référence comme nom (par exemple 81rac01000c.jpg) : rattachées automatiquement
  const parReference = {};
  if (fs.existsSync(IMAGES)) for (const f of fs.readdirSync(IMAGES)) {
    const m = /^(.+).(jpe?g|png|webp)$/i.exec(f);
    if (m) parReference[m[1].toLowerCase().replace(/[^a-z0-9._-]+/g, '-')] = f;
  }
  let imagesEcrites = 0;
  const empreinte = (b) => crypto.createHash('md5').update(b).digest('hex');
  function ecrireImage(fichier, buffer) {
    const chemin = path.join(IMAGES, fichier);
    if (fs.existsSync(chemin) && empreinte(fs.readFileSync(chemin)) === empreinte(buffer)) return;
    fs.mkdirSync(IMAGES, { recursive: true });
    fs.writeFileSync(chemin, buffer);
    imagesEcrites++;
  }
  // Les colonnes sont repérées par leur titre : on peut en ajouter, retirer ou déplacer sans casser l'export
  const norm = (t) => String(t === null || t === undefined ? '' : t).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();
  const TITRES = {
    visible: 'visible', categorie: 'categorie', nom: 'nom du produit', spec: 'caracteristiques (sous le nom)',
    ref: 'reference', ht: 'prix ht', stock: 'stock', photo: 'photo', desc: 'description', techniques: 'caracteristiques techniques',
    famille: 'famille', choix1: 'choix 1', couleur: 'couleur', choix2: 'choix 2', pieces: 'pieces', points: 'points forts'
  };
  // Couleurs reconnues dans la colonne « Couleur du choix 1 » (ou un code comme #2E9E4F)
  const COULEURS = {
    noir: '#1B1B1B', vert: '#2E9E4F', marron: '#7A4B2A', violet: '#9B7FD9', rouge: '#D64541', bleu: '#2F7FD6',
    orange: '#E8892B', blanc: '#F2F2F2', gris: '#8A8F8C', jaune: '#E6C229', beige: '#D9C8A3', rose: '#E58FB0'
  };
  const COL = {};
  ws.getRow(1).eachCell((cell, n) => {
    const t = norm(texte(cell));
    for (const k of Object.keys(TITRES)) if (COL[k] === undefined && t.startsWith(TITRES[k])) COL[k] = n;
  });
  for (const k of ['categorie', 'nom', 'ht']) {
    if (!COL[k]) { console.error('Colonne « ' + TITRES[k] + ' » introuvable dans l\'onglet Produits (ligne 1).'); process.exit(1); }
  }
  const cel = (row, k) => COL[k] ? row.getCell(COL[k]) : { value: null, text: '' };
  const produits = [];
  const familles = {};
  const refsVues = {};
  const lignesDeTexte = (t, sep) => t.split(/\r?\n/).map(l => l.trim()).filter(Boolean).map(l => {
    const i = l.indexOf(sep);
    return i > 0 ? [l.slice(0, i).trim(), l.slice(i + 1).trim()] : [l, ''];
  });
  ws.eachRow((row, n) => {
    if (n === 1) return;
    const nom = texte(cel(row, 'nom'));
    const ref = texte(cel(row, 'ref'));
    const nomFamille = texte(cel(row, 'famille'));
    const vide = !nom && !ref && !nomFamille && !texte(cel(row, 'categorie')) && nombre(cel(row, 'ht')) === null;
    if (vide) return;
    const ou = 'Ligne ' + n + (nom || nomFamille ? ' (' + (nom || nomFamille) + ')' : '');

    if (texte(cel(row, 'visible')).toLowerCase() === 'non') return; // produit masqué

    const catLib = texte(cel(row, 'categorie'));
    const cat = CATEGORIES[catLib];
    if (!cat) { erreurs.push(ou + ' : catégorie « ' + catLib + ' » inconnue. Choisissez-la dans la liste déroulante.'); return; }
    if (!nom && !nomFamille) { erreurs.push(ou + ' : le nom est vide.'); return; }
    const ht = nombre(cel(row, 'ht'));
    if (ht === null || Number.isNaN(ht) || ht < 0) { erreurs.push(ou + ' : le prix HT est vide ou n\'est pas un nombre.'); return; }
    const stockLib = texte(cel(row, 'stock')) || 'En stock';
    const stock = STOCKS[stockLib.toLowerCase()];
    if (!stock) { erreurs.push(ou + ' : stock « ' + stockLib + ' » inconnu. Choisissez-le dans la liste déroulante.'); return; }

    // Photo : une image collée dans la ligne est prioritaire, sinon on lit le nom de fichier de la colonne Photo
    // Une virgule à la place du point avant l'extension (« photo,jpg ») est corrigée
    const photo = texte(cel(row, 'photo')).replace(/,(jpe?g|png|webp)$/i, '.$1');
    const collees = imagesParLigne[n] || [];
    let imgFichier = '';
    if (collees.length) {
      if (collees.length > 1) avertissements.push(ou + ' : ' + collees.length + ' images sur cette ligne, seule la première est utilisée.');
      const base = (ref || nom || nomFamille || 'ligne-' + n).toLowerCase().replace(/[^a-z0-9._-]+/g, '-').replace(/^-+|-+$/g, '');
      const fichier = base + '.' + collees[0].ext;
      ecrireImage(fichier, collees[0].buffer);
      imgFichier = 'images/' + fichier;
      if (collees[0].buffer.length > 500 * 1024) avertissements.push(ou + ' : la photo pèse ' + Math.round(collees[0].buffer.length / 1024) + ' Ko. Réduisez-la (environ 800 pixels de côté) pour une boutique plus rapide.');
    } else if (!photo && ref && parReference[ref.toLowerCase().replace(/[^a-z0-9._-]+/g, '-')]) {
      imgFichier = 'images/' + parReference[ref.toLowerCase().replace(/[^a-z0-9._-]+/g, '-')];
    } else if (photo) {
      const parRef = ref && parReference[ref.toLowerCase().replace(/[^a-z0-9._-]+/g, '-')];
      if (fs.existsSync(path.join(IMAGES, photo))) imgFichier = 'images/' + photo;
      else if (parRef) { imgFichier = 'images/' + parRef; avertissements.push(ou + " : la photo « " + photo + " » n'existe pas ; « " + parRef + " » (nommée d'après la référence) est utilisée à la place. Corrigez le nom dans la colonne Photo."); }
      else avertissements.push(ou + ' : la photo « ' + photo + ' » n\'existe pas dans boutique/images (produit exporté sans photo).');
    }
    if (ref) {
      if (refsVues[ref]) avertissements.push(ou + ' : la référence « ' + ref + ' » est déjà utilisée à la ligne ' + refsVues[ref] + '.');
      refsVues[ref] = n;
    } else {
      avertissements.push(ou + ' : pas de référence.');
    }

    const specs = lignesDeTexte(texte(cel(row, 'techniques')), ':');

    // Ligne d'une famille : regroupée avec les autres lignes de même famille (une seule fiche à variantes)
    if (nomFamille) {
      const c1 = texte(cel(row, 'choix1'));
      const c2 = texte(cel(row, 'choix2'));
      if (!c1) { erreurs.push(ou + ' : « Choix 1 » est vide (par exemple « 2 L/h »). Il est obligatoire pour une famille.'); return; }
      const couleur = texte(cel(row, 'couleur'));
      let hex = '';
      if (couleur) {
        hex = /^#[0-9a-f]{3,8}$/i.test(couleur) ? couleur : COULEURS[norm(couleur)];
        if (!hex) { avertissements.push(ou + ' : couleur « ' + couleur + ' » inconnue (couleurs : ' + Object.keys(COULEURS).join(', ') + ').'); hex = ''; }
      }
      let pieces = nombre(cel(row, 'pieces'));
      if (pieces !== null && (Number.isNaN(pieces) || pieces <= 0)) { avertissements.push(ou + ' : « Pièces par lot » n\'est pas un nombre.'); pieces = null; }
      let f = familles[nomFamille];
      if (!f) {
        f = familles[nomFamille] = { famille: true, cat: cat, title: nomFamille, desc: '', specs: [], points: [], variantes: [] };
        produits.push(f);
      }
      if (f.cat !== cat) avertissements.push(ou + ' : la catégorie diffère de la première ligne de la famille « ' + nomFamille + ' » (celle-ci est conservée).');
      if (!f.desc) f.desc = texte(cel(row, 'desc'));
      if (!f.specs.length) f.specs = specs;
      if (!f.points.length) f.points = lignesDeTexte(texte(cel(row, 'points')), ':');
      if (f.variantes.some(v => v.c1 === c1 && v.c2 === c2)) avertissements.push(ou + ' : la combinaison « ' + c1 + ' » + « ' + c2 + ' » existe déjà dans la famille « ' + nomFamille + ' ».');
      f.variantes.push({ c1: c1, hex: hex, c2: c2, ref: ref, ht: ht, price: arrondi(ht * (1 + tva)), stock: stock, img: imgFichier, pieces: pieces || 0 });
      return;
    }

    produits.push({
      cat: cat,
      name: nom,
      spec: texte(cel(row, 'spec')),
      ref: ref,
      ht: ht,
      price: arrondi(ht * (1 + tva)),
      stock: stock,
      img: imgFichier,
      desc: texte(cel(row, 'desc')),
      specs: specs
    });
  });

  if (erreurs.length) {
    console.error('\nL\'export est annulé : corrigez ces erreurs dans catalogue.xlsx puis recommencez.\n');
    erreurs.forEach(e => console.error('  ✗ ' + e));
    process.exit(1);
  }

  const data = { reglages: { tva: tva, livraisonOfferte: franco, fraisLivraison: frais }, produits: produits };
  const contenu = '/* Fichier généré par catalogue/exporter.js à partir de catalogue.xlsx. Ne pas modifier à la main. */\n' +
    'window.CATALOGUE = ' + JSON.stringify(data, null, 1) + ';\n';
  fs.writeFileSync(SORTIE, contenu, 'utf8');

  // Nouveau numéro de version dans la page : les navigateurs rechargent les prix à jour
  const PAGE = process.env.CATALOGUE_PAGE || path.join(__dirname, '..', 'boutique', 'index.html');
  let page = fs.readFileSync(PAGE, 'utf8');
  const versionnee = page.replace(/produits\.js\?v=[^"']*/, 'produits.js?v=' + Date.now());
  if (versionnee !== page) fs.writeFileSync(PAGE, versionnee, 'utf8');
  else avertissements.push('index.html ne contient pas la ligne <script src="produits.js?v=...">.');

  if (imagesEcrites) console.log('\n' + imagesEcrites + ' photo(s) enregistrée(s) depuis le tableur dans boutique/images.');
  const nbFam = produits.filter(p => p.famille).length;
  console.log('\n✓ ' + produits.length + ' articles exportés vers boutique/produits.js' +
    (nbFam ? ' (dont ' + nbFam + ' fiches à variantes, ' + produits.filter(p => p.famille).reduce((s, f) => s + f.variantes.length, 0) + ' variantes)' : ''));
  if (avertissements.length) {
    console.log('\nÀ vérifier :');
    avertissements.forEach(a => console.log('  ! ' + a));
  }
}

main().catch(e => { console.error('Erreur : ' + e.message); process.exit(1); });
