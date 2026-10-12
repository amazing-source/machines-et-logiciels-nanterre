# -*- coding: utf-8 -*-
# Met le kit a ton nom, sans ouvrir Word, Excel ni PowerPoint.
#
#   python outils/remplir.py              copie les modeles dans a-deposer/, avec
#                                         ton identite et les bons noms de fichiers
#   python outils/remplir.py --appliquer  reporte les valeurs ecrites dans
#                                         a-deposer/cases.json
#   python outils/remplir.py --verifier   liste ce qui reste a completer
#
# L'identite est lue dans outils/identite.json (copie de identite.exemple.json).
# Les modeles des dossiers seance-N ne sont jamais modifies. Un fichier deja
# present dans a-deposer/ n'est jamais ecrase (ta narration y est peut-etre
# enregistree), sauf avec --ecraser.
# Python 3 seul : aucune bibliotheque a installer.
import html
import io
import json
import re
import sys
import zipfile
from pathlib import Path

RACINE = Path(__file__).resolve().parent.parent
SORTIE = RACINE / "a-deposer"
IDENTITE = RACINE / "outils" / "identite.json"
CASES = SORTIE / "cases.json"
OFFICE = (".docx", ".xlsx", ".pptx")

# cases d'identite : remplies d'office a partir de identite.json
FIXES = (("[NOM Prénom]", "nom_prenom"), ("[Prénom Nom]", "prenom_nom"), ("[P. Nom]", "signature"),
         ("[N° étudiant]", "numero"), ("[Filière]", "filiere"))
# cases remplies seulement si identite.json les renseigne (seance 6, question 1)
OPTION = (("[ta ville, ton pays]", "dnt_ville"), ("[ton fournisseur d'accès]", "dnt_fai"),
          ("[ton navigateur]", "dnt_navigateur"), ("[ton système d'exploitation]", "dnt_systeme"))
# toute case encore vide : identite, ou valeur a relever sur sa propre machine
MOTIF = re.compile(r"\[(?:NOM Prénom|Prénom Nom|P\. Nom|N° étudiant|Filière)\]"
                   r"|\[(?:À COMPLÉTER|taille|largeur|hauteur|nom de ton|ta |ton |tonnom|Vérifie |Cette activité )[^\[\]<>]{0,700}\]")
# remarques du modele : pas de valeur a donner, elles se suppriment
NOTE = ("[Vérifie ", "[Cette activité ")
MODE_D_EMPLOI = ("Pour chaque case, ecris la valeur dans \"valeur\" (du texte, entre guillemets). "
                 "null = pas encore remplie ; \"\" = la case est supprimee (c'est le reglage des remarques du modele). "
                 "\"contexte\" montre la phrase, la case concernee entre >>> et <<<. "
                 "Ne change rien d'autre, puis lance : python outils/remplir.py --appliquer")


def vide(valeur):
    return valeur is None or "[" in str(valeur) or not str(valeur).strip()


def lire_identite():
    if not IDENTITE.exists():
        sys.exit("Il manque outils/identite.json : copie outils/identite.exemple.json sous ce nom et remplis-le.")
    idt = json.loads(IDENTITE.read_text(encoding="utf-8-sig"))
    manque = [c for c in ("nom_prenom", "numero", "filiere") if vide(idt.get(c))]
    if manque:
        sys.exit("outils/identite.json : remplis d'abord " + ", ".join(manque) + ".")
    # le nom s'ecrit en capitales (DUPONT Marie) : c'est ce qui le separe du prenom
    mots = idt["nom_prenom"].split()
    nom = [m for m in mots if m == m.upper() and any(c.isalpha() for c in m)]
    if not nom or len(nom) == len(mots):
        nom = mots[:1]
    prenom = " ".join(m for m in mots if m not in nom)
    nom = " ".join(nom).title()
    deduit = {"prenom_nom": (prenom + " " + nom).strip(), "signature": (prenom[:1] + ". " if prenom else "") + nom}
    for cle, valeur in deduit.items():
        if vide(idt.get(cle)):
            idt[cle] = valeur
            print("  %s non renseigne, deduit : %s" % (cle, valeur))
    if vide(idt.get("fichier_numero")) or idt["fichier_numero"] == "NUMERO":
        idt["fichier_numero"] = idt["numero"]
    return idt


def table(idt):
    t = {case: idt[cle] for case, cle in FIXES}
    t.update({case: idt[cle] for case, cle in OPTION if not vide(idt.get(cle))})
    return {html.escape(c, quote=False): html.escape(str(v), quote=False) for c, v in t.items()}


def reecrire(data, fonction, prefixe=""):
    """Recopie une archive Office en passant chaque partie XML par fonction(nom, texte)."""
    src = zipfile.ZipFile(io.BytesIO(data))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in src.infolist():
            d = src.read(item.filename)
            bas = item.filename.lower()
            if bas.endswith((".xml", ".rels")):
                d = fonction(prefixe + item.filename, d.decode("utf-8")).encode("utf-8")
            elif bas.endswith(OFFICE):  # classeur des donnees d'un graphique
                d = reecrire(d, fonction, prefixe + item.filename + "!")
            z.writestr(item, d)
    return out.getvalue()


def parties(data, prefixe=""):
    z = zipfile.ZipFile(io.BytesIO(data))
    for nom in z.namelist():
        if nom.lower().endswith(OFFICE):
            yield from parties(z.read(nom), prefixe + nom + "!")
        elif nom.lower().endswith(".xml"):
            yield prefixe + nom, z.read(nom).decode("utf-8")


def sans_balises(xml):
    return re.sub(r"<[^>]+>", "", xml)


def clair(xml):
    return html.unescape(sans_balises(xml))


def fichiers():
    return sorted(p for p in SORTIE.rglob("*") if p.suffix.lower() in OFFICE and not p.name.startswith("~$"))


def releve(chemin):
    """Cases encore vides d'un fichier : (partie, rang dans la partie, case, contexte)."""
    for partie, xml in parties(chemin.read_bytes()):
        for rang, m in enumerate(MOTIF.finditer(xml), 1):
            debut = max(xml.rfind(b, 0, m.start()) for b in ("<w:p>", "<w:p ", "<a:p>", "<si>"))
            fins = [i for i in (xml.find(b, m.end()) for b in ("</w:p>", "</a:p>", "</si>")) if i >= 0]
            avant, apres = clair(xml[max(debut, 0):m.start()]), clair(xml[m.end():min(fins or [len(xml)])])
            yield partie, rang, clair(m.group()), avant[-250:] + ">>>" + clair(m.group()) + "<<<" + apres[:150]


def inventaire(garder):
    """Ecrit a-deposer/cases.json. garder : conserve les valeurs deja ecrites
    dans l'ancien fichier, tant que les cases du document n'ont pas bouge."""
    ancien = {}
    if garder and CASES.exists():
        for c in json.loads(CASES.read_text(encoding="utf-8-sig")).get("cases", []):
            ancien.setdefault(c["fichier"], []).append(c)
    cases = []
    for f in fichiers():
        rel = f.relative_to(SORTIE).as_posix()
        trouve = [{"fichier": rel, "partie": p, "rang": r, "case": c, "contexte": ctx,
                   "valeur": "" if c.startswith(NOTE) else None} for p, r, c, ctx in releve(f)]
        avant = ancien.get(rel, [])
        if [(c["partie"], c["rang"], c["case"]) for c in avant] == [(c["partie"], c["rang"], c["case"]) for c in trouve]:
            for c, a in zip(trouve, avant):
                c["valeur"] = a.get("valeur")
        cases += trouve
    SORTIE.mkdir(exist_ok=True)
    CASES.write_text(json.dumps({"_mode_d_emploi": MODE_D_EMPLOI, "cases": cases}, ensure_ascii=False, indent=2), encoding="utf-8")
    return cases


def copier(ecraser):
    idt = lire_identite()
    t = table(idt)
    nom_fichier = re.sub(r'[\\/:*?"<>|]', "", idt["nom_prenom"]).strip().replace(" ", "_")

    def remplacer(partie, xml):
        for case, valeur in t.items():
            xml = xml.replace(case, valeur)
        return xml

    for src in sorted(RACINE.glob("seance-*/*")):
        if src.suffix.lower() not in OFFICE or src.name.startswith("~$"):
            continue
        dst = SORTIE / src.parent.name / src.name.replace("NOM_Prenom", nom_fichier).replace("NUMERO", str(idt["fichier_numero"]))
        if dst.exists() and not ecraser:
            print("  conserve : %s (deja present)" % dst.relative_to(RACINE).as_posix())
            continue
        dst.parent.mkdir(parents=True, exist_ok=True)
        dst.write_bytes(reecrire(src.read_bytes(), remplacer))
        print("  ecrit    : %s" % dst.relative_to(RACINE).as_posix())
    bilan(inventaire(garder=True))


def appliquer():
    if not CASES.exists():
        sys.exit("Pas de a-deposer/cases.json : lance d'abord  python outils/remplir.py")
    voulu = {}
    for c in json.loads(CASES.read_text(encoding="utf-8-sig"))["cases"]:
        if c.get("valeur") is not None:
            voulu.setdefault(c["fichier"], {}).setdefault(c["partie"], {})[c["rang"]] = c
    # tout est controle avant la premiere ecriture : un fichier n'est jamais
    # rempli a moitie
    pret, n = {}, 0
    for rel, par_partie in voulu.items():
        def poser(partie, xml, par_partie=par_partie, rel=rel):
            trouves = list(MOTIF.finditer(xml))
            for rang in sorted(par_partie.get(partie, {}), reverse=True):
                c = par_partie[partie][rang]
                if rang > len(trouves) or clair(trouves[rang - 1].group()) != c["case"]:
                    sys.exit("%s a change depuis le releve des cases : relance  python outils/remplir.py  puis remplis de nouveau cases.json." % rel)
                m = trouves[rang - 1]
                xml = xml[:m.start()] + html.escape(str(c["valeur"]), quote=False) + xml[m.end():]
            return xml
        pret[rel] = reecrire((SORTIE / rel).read_bytes(), poser)
        n += sum(len(v) for v in par_partie.values())
    for rel, data in pret.items():
        (SORTIE / rel).write_bytes(data)
    print("  %d case(s) remplie(s) dans %d fichier(s)" % (n, len(pret)))
    bilan(inventaire(garder=False))


def narration(data):
    """Diapositives qui portent un enregistrement audio, sur le total."""
    z = zipfile.ZipFile(io.BytesIO(data))
    noms = set(z.namelist())
    diapos = [n for n in noms if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)]
    avec = 0
    for d in diapos:
        rels = d.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
        if rels in noms and re.search(r'Target="[^"]+\.(?:m4a|wav|mp3|wma|aac)"', z.read(rels).decode("utf-8"), re.I):
            avec += 1
    return avec, len(diapos)


def bilan(cases=None):
    """Etat de a-deposer/ : cases restantes par fichier, narration du projet."""
    reste = 0
    print()
    for f in fichiers():
        data = f.read_bytes()
        n = sum(len(MOTIF.findall(sans_balises(xml))) for _, xml in parties(data))
        reste += n
        print("  %-13s %s" % ("complet" if not n else "%d case(s)" % n, f.relative_to(SORTIE).as_posix()))
        if f.name.startswith("PROJET_MULTIMEDIA") and f.suffix.lower() == ".pptx":
            avec, total = narration(data)
            if avec < total:
                print("  %-13s narration : %d diapositive(s) sur %d -> A ENREGISTRER AVEC TA VOIX" % ("", avec, total))
    print()
    if reste:
        print("Il reste %d case(s) : ce sont des valeurs a relever chez toi." % reste)
        if cases is not None:
            print("Elles sont listees dans a-deposer/cases.json : ecris les valeurs, puis  python outils/remplir.py --appliquer")
            if reste > len(cases):
                print("(%d d'entre elles sont a remplir a la main dans le fichier : Ctrl + H, cherche [ )" % (reste - len(cases)))
    else:
        print("Plus aucune case a remplir.")
    print("Reste a faire toi-meme : enregistrer ta narration (projet de la seance 5), deposer sur coursenligne, finir PIX.")
    return 1 if reste else 0


if __name__ == "__main__":
    sys.stdout.reconfigure(errors="replace")
    if "--appliquer" in sys.argv:
        appliquer()
    elif "--verifier" in sys.argv:
        sys.exit(bilan())
    else:
        copier("--ecraser" in sys.argv)
