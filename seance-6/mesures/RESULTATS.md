# Séance 6 — résultats des mesures du 12 octobre 2026

Mesures faites avec `outils/seance6_mesures.js`, dans Google Chrome, sur un profil vierge (aucun historique, aucun compte connecté). Elles sont là comme exemple : **tes résultats seront différents**, c'est normal — refais-les.

## Question 2 — Cookies (lemonde.fr)

Fichier : `cookies.json`

| Moment | Cookies | dont du site | dont de tiers | Domaines tiers |
|---|---|---|---|---|
| À l'arrivée, avant tout choix sur le bandeau | 10 | 10 | 0 | 0 |
| Après « Accepter et continuer », page défilée | 215 | 26 | 189 | 70 |

Dans le même temps, la page est passée de 5 à 119 sites tiers contactés. Parmi les domaines qui déposent des cookies : régies publicitaires et mesure d'audience (Criteo, Amazon, Adform, The Trade Desk, Bing…).

## Question 4 — Navigation au clavier (parisnanterre.fr)

Fichiers : `clavier_essai3.json`, `clavier_essai4.json` (chaque arrêt de la touche Tab y est décrit)

- La page d'accueil compte 140 liens ; **aucun ne contient le mot « Contact »**. Le plus proche, « Annuaire », est dans le pied de page.
- « Annuaire » est atteint au **80ᵉ appui** sur Tab (résultat identique sur trois essais).
- Premier arrêt : le bouton de langue. Il n'y a pas de lien d'évitement du type « Aller au contenu ».
- Le bandeau des cookies n'est atteint qu'au 10ᵉ appui.
- 16 à 18 arrêts tombent sur des éléments invisibles à l'écran (diapositives masquées des carrousels).
- Le repère de focus est le contour par défaut du navigateur, d'un pixel.

## Question 6 — Empreinte carbone (websitecarbon.com)

Page testée : parisnanterre.fr — résultat affiché le 12 octobre 2026, dernier test du site le 7 octobre 2026.

| Indicateur | Valeur |
|---|---|
| Note | F (échelle de A+ à F) |
| Position | plus polluant que 86 % des pages testées |
| CO2 par visite | 0,81 g |
| Sur un an, avec 10 000 visites par mois | 97,09 kg de CO2, soit ce que 5 arbres absorbent en un an |
| Hébergement | pas d'énergie verte détectée |

## Question 1 — Do Not Track (épisode 1)

Le site interroge son propre service de profilage dès l'ouverture de l'épisode. Il en reçoit :

- la ville, le pays et les coordonnées approximatives, déduits de l'adresse IP ;
- le fournisseur d'accès à Internet ;
- le navigateur, sa version et le système d'exploitation ;
- la météo du lieu, pour personnaliser l'épisode.

Chaque page envoie aussi la langue du navigateur, la taille de l'écran, l'heure et un identifiant de visiteur à trois outils de suivi : AT Internet (Xiti), Google Analytics et DoubleClick.

Les valeurs relevées ne sont pas publiées ici : ce sont celles de la personne qui a fait le test. Relève les tiennes.

## Question 5 — Simulateur Atalan

Simulateurs proposés : daltonisme, malvoyance, cécité, surdité, handicap moteur. Testé : malvoyance, filtre DMLA — une large tache grise masque le centre de l'écran, seul le pourtour reste lisible.
