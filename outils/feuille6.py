# -*- coding: utf-8 -*-
# Seance 6 : feuille de reponses. Les reponses de l'activite 6 reprennent des
# mesures reelles faites le 12/10/2026 avec seance6_mesures.js (resultats
# bruts conserves dans seance-6/mesures/).
# Usage : python feuille6.py <enonce.docx> <sortie.docx> <identite.json>
import json
import re
import sys

from feuille_commun import Feuille

SRC, OUT = sys.argv[1], sys.argv[2]
ID = json.load(open(sys.argv[3], encoding="utf-8"))
NB = " "


def typo(t):
    t = re.sub(r" ([?:;!»%])", NB + r"\1", t)
    t = t.replace("« ", "«" + NB)
    t = re.sub(r"(\d) (\d{3})", r"\1" + NB + r"\2", t)
    return t


f = Feuille(SRC)


def rep(fragment, *lignes, occurrence=1):
    f.repondre(fragment, [typo(l) for l in lignes], occurrence)


f.apres(f.premier(), f.identite("%s — %s" % (ID["nom_prenom"], ID["numero"])))

# ---------------------------------------------------------------- ACTIVITE 1
rep("Découvrez les personnes qui ont les mêmes goûts",
    "Repères de paramétrage (coursenligne est une plate-forme Moodle) :",
    "• chemin : menu utilisateur (mon nom, en haut à droite) > Profil > Modifier le profil ;",
    "• rubrique « Généraux » : le réglage « Affichage de l'adresse de courriel » permet de la masquer ou de ne la montrer qu'aux participants de ses cours ; la ville et la description sont facultatives, on ne les remplit que si l'on accepte qu'elles soient vues ;",
    "• rubrique « Avatar utilisateur » : l'image déposée est visible de tous les utilisateurs, d'où le choix d'une photo ou d'un avatar neutre et correct ;",
    "• rubrique « Centres d'intérêt » : chaque tag saisi devient, une fois le profil enregistré, un lien cliquable qui affiche les utilisateurs ayant choisi le même tag.",
    "Principe retenu : ne rendre visible que le strict nécessaire, comme le recommande le cours pour toute inscription à un service en ligne.")

# ---------------------------------------------------------------- ACTIVITE 2
rep("Quelles informations l'internaute doit-il saisir",
    "Le formulaire d'inscription (facebook.com, consulté le 12/10/2026) demande : le prénom, le nom, la date de naissance (jour, mois, année), le genre, un numéro de mobile ou une adresse e-mail, et un mot de passe.",
    "Avant même d'utiliser le service, on livre donc son identité civile, son âge et un moyen de contact.")

rep("badge de vérification bleu",
    "Le badge bleu signifie que Facebook (Meta) a vérifié l'authenticité de la Page : il s'agit bien de la Page officielle de la CNIL, et non d'une imitation ou d'une page de fans.",
    "C'est un repère de confiance sur l'identité de l'émetteur, pas un jugement sur le contenu publié.")

rep("quelles sont les conséquences respectives",
    "« J'aime » : j'affiche publiquement mon soutien. La Page figure parmi mes mentions J'aime, je suis compté parmi ses fans, mes amis peuvent le voir, et je suis automatiquement abonné à ses publications.",
    "« S'abonner » : je reçois les publications de la Page dans mon fil d'actualité, sans déclarer que je l'aime.")

rep("Quel type d'exploitation et d'utilisation des données",
    "Facebook fournit aux administrateurs des Pages des statistiques sur leur audience : nombre de personnes touchées, réactions, et profil des fans sous forme agrégée (âge, genre, ville, pays, langue, heures de connexion).",
    "La Page peut s'en servir pour cibler ses publications et ses publicités vers ses fans, leurs amis et des personnes au profil comparable. Ma mention J'aime peut aussi être montrée à mes amis à côté d'une publicité de la Page.")

rep("Peut-on suivre l'actualité d'une page",
    "En partie seulement. S'abonner sans aimer évite d'afficher une mention J'aime sur mon profil et de servir de recommandation auprès de mes amis.",
    "Mais Facebook enregistre quand même mon abonnement et toutes mes interactions (publications vues, clics, commentaires) : elles alimentent les statistiques de la Page et le ciblage publicitaire. On ne suit donc jamais une Page sans laisser de traces.")

rep("Diriez-vous que les services en ligne",
    "Oui. Un service gratuit comme Facebook se finance par la publicité : il vend aux annonceurs des espaces publicitaires et un ciblage alimenté par les données de ses utilisateurs (profil, centres d'intérêt, navigation).",
    "Les données sont la matière première ; l'audience et le ciblage sont le produit vendu. D'où la formule : « si c'est gratuit, c'est vous le produit ».")

rep("Que permet Google Analytics",
    "Google Analytics est un outil de mesure d'audience. Il indique à une entreprise combien de visiteurs viennent sur son site, d'où ils viennent (moteur de recherche, réseau social, publicité), avec quel appareil, quelles pages ils consultent, combien de temps ils restent et s'ils vont jusqu'à l'achat.",
    "Elle peut ainsi améliorer son site et mesurer l'efficacité de ses campagnes. En contrepartie, la navigation des visiteurs est suivie au moyen de cookies.")

# ---------------------------------------------------------------- ACTIVITE 3
rep("Explorez et notez les différents services",
    "Sur l'ENT (ent.parisnanterre.fr) : la messagerie universitaire ; Mon emploi du temps (horaires et salles, y compris pour les examens) ; Mon dossier Étudiant (notes et résultats) ; la Scolarité (certificat de scolarité) ; l'accès aux cours en ligne (coursenligne) et aux ressources des bibliothèques. Le bouton Personnaliser permet de réorganiser l'interface.",
    "Sur le site de la DRI (dri.parisnanterre.fr, consulté le 12/10/2026) : l'accès au système d'information (activation et gestion du compte, wifi eduroam, eduspot et UPNguest, connexion aux postes de travail, authentification multifacteur, bonnes pratiques de sécurité) ; les applications (communication et organisation, travail collaboratif avec le stockage Microsoft 365 / OneDrive, visioconférence) ; l'équipement (postes de travail, logiciels) ; la demande d'assistance.")

# ---------------------------------------------------------------- ACTIVITE 4
rep("d'après les 2 options proposées",
    "L'annuaire santé permet de rechercher un professionnel de santé. Le formulaire propose deux critères, dont au moins un doit être rempli : le nom d'un professionnel de santé ou d'un centre de santé, ou bien une profession (médecin généraliste, chirurgien-dentiste, infirmier, masseur-kinésithérapeute…), à compléter par un lieu : rue, ville ou code postal.",
    "(Site consulté le 12/10/2026 : il a été refondu, les deux « options » de l'énoncé sont devenues ces deux critères de recherche.)")

rep("Dispositions générales en bas de page",
    "Cette rubrique s'appelle aujourd'hui « CGU ». Le service, gratuit et facultatif, permet de trouver les coordonnées des professionnels de santé exerçant à titre libéral et des centres de santé, de connaître leur secteur conventionnel et de savoir s'ils acceptent la carte Vitale.",
    "Il répond à la mission d'information des assurés confiée à l'Assurance Maladie pour faciliter l'accès aux soins.")

rep("Mentions légales en bas de page",
    "La CNIL (Commission nationale de l'informatique et des libertés), autorité indépendante auprès de laquelle toute personne peut déposer une réclamation.",
    "Le responsable du traitement est la Cnam, qui dispose d'un délégué à la protection des données (DPO).")

rep("Cliquez sur Accéder à mon compte",
    "Le protocole HTTPS, c'est-à-dire HTTP sécurisé par le chiffrement TLS (anciennement SSL) : l'adresse de la page de connexion commence par https:// et le navigateur affiche un cadenas.",
    "Les identifiants et les données de santé circulent chiffrés entre le navigateur et le serveur, et le certificat garantit que l'on se trouve bien sur le site de l'Assurance Maladie.")

rep("En quoi le site Internet ameli.fr facilite",
    "Le compte ameli dématérialise les démarches : suivi des remboursements, téléchargement d'attestations (droits, indemnités journalières), commande d'une carte Vitale ou d'une carte européenne d'assurance maladie, déclaration d'un changement de situation, messagerie avec sa caisse.",
    "L'assuré agit 24 h/24 sans se déplacer ni envoyer de courrier ; les professionnels de santé télétransmettent les feuilles de soins, ce qui raccourcit les délais de remboursement et réduit le papier.")

# ---------------------------------------------------------------- ACTIVITE 5
rep("Quelle est la différence entre une donnée personnelle",
    "Une donnée personnelle est toute information se rapportant à une personne physique identifiée ou identifiable (nom, numéro de téléphone, adresse e-mail…).",
    "Une donnée sensible est une catégorie particulière de donnée personnelle, que le RGPD interdit en principe de recueillir ou d'utiliser, sauf exceptions, parce que sa divulgation expose à des discriminations.",
    "Deux exemples : les données de santé (un diagnostic, un arrêt maladie) et les convictions religieuses. On peut citer aussi les opinions politiques, l'origine ethnique, l'orientation sexuelle ou les données biométriques.")

rep("Vous prenez une photo avec votre smartphone",
    "La date et l'heure de la prise de vue, la position GPS, le modèle du téléphone et les réglages de l'appareil, les dimensions et le poids du fichier.",
    "S'y ajoutent, à l'envoi, l'expéditeur, le destinataire, la date et l'heure du message.")

rep("Le Big Data se définit par la règle",
    "Volume : la masse gigantesque de données produites et stockées.",
    "Vélocité : la vitesse à laquelle elles sont créées, collectées et analysées, souvent en temps réel.",
    "Variété : la diversité de leurs formes et de leurs sources (textes, photos, vidéos, données de capteurs, traces de navigation).")

rep("Vous découvrez que votre nom apparaît dans les résultats",
    "Le droit au déréférencement : je demande au moteur de recherche de ne plus associer mon nom et mon prénom à ce lien dans ses résultats.",
    "Le blog, lui, reste en ligne : pour le faire disparaître, il faut exercer en plus le droit à l'effacement auprès de son éditeur ou de son hébergeur.")

rep("Expliquez avec vos propres mots la différence fondamentale",
    "L'authentification vérifie qu'une personne est bien celle qu'elle prétend être : mon visage est comparé à un seul modèle, le mien (déverrouiller son téléphone, passer un contrôle d'accès).",
    "L'identification cherche à retrouver qui est une personne parmi d'autres : un visage est comparé à tous les modèles d'une base de données (reconnaître quelqu'un dans une foule).",
    "La première répond à la question « est-ce bien vous ? », la seconde à « qui êtes-vous ? ».")

rep("Vous remarquez que YouTube ou Netflix",
    "Un algorithme de recommandation, de type auto-apprenant : son comportement évolue en fonction des données qu'on lui fournit (mes vidéos vues, mes recherches, mes « j'aime »).",
    "Il relève de l'intelligence artificielle, et plus précisément de l'apprentissage automatique (machine learning).")

rep("Citez trois types d'historiques",
    "L'historique du navigateur (pages visitées, fichiers téléchargés), l'historique du moteur de recherche (requêtes saisies) et l'historique de géolocalisation du smartphone. On peut ajouter l'historique des vidéos regardées ou des publications aimées.",
    "Un cookie est une petite suite d'informations, identifiée par un nom, qu'un site dépose dans mon navigateur. Celui-ci la conserve, puis la renvoie au site à chaque visite : cela permet de me reconnaître (panier, session), mais aussi de tracer ma navigation à des fins statistiques ou publicitaires.")

rep("Face à l'hyper-connectivité",
    "Le droit à la déconnexion (loi Travail de 2016, en vigueur depuis le 1er janvier 2017) : le salarié n'est pas tenu de répondre aux courriels, appels et messages professionnels en dehors de son temps de travail, afin de protéger ses temps de repos et sa vie personnelle.")

rep("Qu'est-ce que l'accessibilité numérique",
    "C'est le fait de concevoir les sites, les applications et les documents pour que tout le monde puisse les utiliser, y compris les personnes en situation de handicap.",
    "Une personne aveugle navigue avec un lecteur d'écran qui lit la page à voix haute : il faut donc des textes alternatifs sur les images, des titres structurés et une navigation possible au clavier. Dans un PowerPoint : contrastes suffisants, polices lisibles, ordre de lecture logique, vidéos sous-titrées.")

rep("où sont stockées physiquement les ressources",
    "Dans des centres de données (data centers) : de vastes bâtiments remplis de serveurs, reliés au réseau par des câbles et répartis dans le monde entier.",
    "Ces serveurs fonctionnent 24 h/24, consomment beaucoup d'électricité et doivent être refroidis en permanence ; les données y sont en outre dupliquées par sécurité. Comme on ne voit ni les machines ni l'énergie dépensée, cette pollution reste invisible pour l'utilisateur.")

# ---------------------------------------------------------------- ACTIVITE 6
rep("Quelles sont les données que ce site a réussi à capter",
    "En ouvrant l'épisode 1, « Routines matinales », le 12/10/2026, le site connaissait déjà, sans que j'aie rien saisi : ma ville et mon pays, déduits de mon adresse IP (%s) ; mon fournisseur d'accès à Internet (%s) ; mon navigateur et son numéro de version (%s) ; mon système d'exploitation (%s) ; la langue de mon navigateur (français) ; la taille de mon écran et l'heure de ma visite."
    % (ID["dnt_ville"], ID["dnt_fai"], ID["dnt_navigateur"], ID["dnt_systeme"]),
    "Il a aussi récupéré la météo de ma ville pour personnaliser l'épisode, m'a attribué un identifiant de visiteur par cookie et a signalé ma visite à trois outils de suivi : AT Internet (Xiti), Google Analytics et DoubleClick.")

rep("Combien de cookies sont déposés sur votre navigateur",
    "Site testé : lemonde.fr, le 12/10/2026, dans un navigateur sans historique.",
    "À l'arrivée, avant tout choix sur le bandeau : 10 cookies, tous déposés par le site lui-même, aucun par un tiers.",
    "Après avoir cliqué sur « Accepter et continuer » et fait défiler la page : 215 cookies, dont 26 du site et 189 déposés par 70 domaines tiers (régies publicitaires et mesure d'audience : Criteo, Amazon, Adform, The Trade Desk, Bing…).",
    "Près de neuf cookies sur dix viennent donc de sociétés dont je n'ai jamais visité le site.")

rep("Quelles sont les 3 grandes catégories",
    "L'historique de navigation ; les cookies et autres données des sites ; les images et fichiers en cache.",
    "(Les paramètres avancés ajoutent l'historique des téléchargements, les mots de passe enregistrés et les données de saisie automatique.)")

rep("Avez-vous réussi ?",
    "Pas complètement. Sur www.parisnanterre.fr (testé le 12/10/2026), la page d'accueil ne contient aucun lien « Contact » : le plus proche, « Annuaire », se trouve dans le pied de page et je ne l'ai atteint qu'au 80e appui sur la touche Tab.",
    "Difficultés rencontrées :",
    "• aucun lien d'évitement du type « Aller au contenu » : le premier Tab tombe sur le bouton de langue, puis il faut traverser tout le menu ;",
    "• la fenêtre d'information et le bandeau cookies recouvrent la page, mais le clavier ne les atteint qu'au 10e appui ;",
    "• une vingtaine d'arrêts (16 à 18 selon les essais) se font sur des éléments invisibles à l'écran, les diapositives masquées des carrousels : on ne sait plus où l'on est ;",
    "• le repère de focus n'est qu'un fin contour d'un pixel, difficile à suivre des yeux.")

rep("Quel simulateur avez-vous testé",
    "J'ai testé le simulateur « Malvoyance » avec le filtre DMLA (dégénérescence maculaire liée à l'âge) : une large tache masque le centre de l'écran, là où se pose le regard, et seul le pourtour reste visible.",
    "Un Captcha qui demande de sélectionner des images devient alors impossible, puisque ce que l'on fixe est précisément ce qui disparaît ; sans alternative (Captcha audio, question en texte), l'utilisateur est bloqué et ne peut ni s'inscrire ni se connecter. De même, un site aux textes peu contrastés, ou qui se déforment quand on les agrandit, lui devient illisible.")

rep("Quel est le score écologique du site",
    "Site testé : www.parisnanterre.fr, le portail de mon université.",
    "Note écologique : F, la plus mauvaise de l'échelle (de A+ à F) : la page est plus polluante que 86 % des pages web testées. Chaque visite produit 0,81 g de CO2.",
    "Avec 10 000 visites par mois, cela représente environ 97 kg de CO2 par an, soit ce que 5 arbres absorbent en une année. L'outil signale aussi que l'hébergement n'utilise pas d'énergie verte.",
    "(Résultat affiché le 12/10/2026 ; dernier test du site le 7 octobre 2026.)")

rep("citez deux \"écogestes\" immédiats",
    "1. Faire le tri : supprimer les mails inutiles, en commençant par ceux qui contiennent de grosses pièces jointes, puis vider les dossiers Spam et Corbeille.",
    "2. Me désinscrire des newsletters que je ne lis jamais, pour ne plus stocker de messages inutiles.",
    "Et à l'avenir : envoyer un lien de partage plutôt qu'une pièce jointe.")

f.enregistrer(OUT)
