# -*- coding: utf-8 -*-
# Seance 5, projet : telechargement du jeu de donnees ouvert Parcoursup 2025
# (portail data.enseignementsup-recherche.gouv.fr, Licence Ouverte v2.0),
# restreint aux formations d'Ile-de-France.
# Usage : python projet5_telecharge.py <dossier_de_sortie>
import csv
import io
import os
import sys

import requests

OUT = sys.argv[1]
os.makedirs(OUT, exist_ok=True)

BASE = ("https://data.enseignementsup-recherche.gouv.fr/api/explore/v2.1/"
        "catalog/datasets/fr-esr-parcoursup")

meta = requests.get(BASE, timeout=60).json()
print("Jeu      :", meta["metas"]["default"]["title"])
print("Modifie  :", meta["metas"]["default"]["modified"])
print("Licence  :", meta["metas"]["default"]["license"])
print("Lignes   :", meta["metas"]["default"]["records_count"])

r = requests.get(BASE + "/exports/csv", timeout=180, params={
    "where": 'region_etab_aff="Ile-de-France"',
    "delimiter": ";",
    "use_labels": "false",
    "limit": -1,
})
r.raise_for_status()
raw = os.path.join(OUT, "fr-esr-parcoursup_2025_ile-de-france_brut.csv")
with open(raw, "wb") as f:
    f.write(r.content)

rows = list(csv.DictReader(io.StringIO(r.content.decode("utf-8-sig")), delimiter=";"))
print("Extrait  :", len(rows), "formations,", len(rows[0]), "colonnes ->", raw)

# dictionnaire des champs (nom technique -> libelle officiel)
with open(os.path.join(OUT, "dictionnaire_des_champs.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.writer(f, delimiter=";")
    w.writerow(["champ", "type", "libelle"])
    for fld in meta["fields"]:
        w.writerow([fld["name"], fld["type"], fld["label"]])
