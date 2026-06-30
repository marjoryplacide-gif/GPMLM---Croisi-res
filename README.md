# Planning des escales paquebots — GPMLM

Application qui génère automatiquement le fichier Excel "planning des
escales paquebots" destiné aux agents, à partir d'une extraction VIGIEsip
(.csv).

Remplace le travail manuel décrit dans le mode opératoire d'Yvon
("Mise en page fichier croisières Vigiesip.docx").

## Utilisation (une fois l'appli en ligne)

1. Ouvrir le lien de l'application dans un navigateur
2. Déposer le fichier .csv exporté depuis VIGIEsip
3. Vérifier/ajuster les dates de début et fin de saison
4. Cliquer sur "Générer le planning"
5. Cliquer sur "Télécharger le fichier Excel"

Aucune installation, aucune commande à taper.

## Tester en local (pour Marjo, avant déploiement)

```bash
pip install -r requirements.txt
streamlit run app.py
```

Un onglet de navigateur s'ouvre automatiquement sur l'application.

## Déployer en ligne (Streamlit Community Cloud)

Même procédure que pour le projet "Droits de port" :

1. Créer un dépôt GitHub contenant ces fichiers (`app.py`, `traitement.py`,
   `requirements.txt`)
2. Sur https://share.streamlit.io, connecter ce dépôt
3. Indiquer `app.py` comme fichier principal
4. L'application est accessible via un lien à partager à n'importe quel
   collègue — aucune installation nécessaire de leur côté

## Fichiers du projet

- `app.py` — l'interface (dépôt de fichier, options, téléchargement)
- `traitement.py` — la logique de transformation (lecture du CSV,
  reconstruction des escales, filtrage, mise en forme Excel)
- `requirements.txt` — librairies Python nécessaires

## Point en suspens

Le fichier final d'Yvon mentionne des postes (GQ, SP, BEACH...) absents de
l'extraction VIGIEsip (qui n'en contient que 4 : Pointe Simon Est/Ouest,
Tourelles, Quai des Annexes). Origine à confirmer avec Yvon/Lauriane.
