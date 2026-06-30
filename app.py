"""
Application Streamlit : génère le planning des escales paquebots (Excel)
à partir d'une extraction VIGIEsip (.csv).

Lancer en local :
    streamlit run app.py
"""

import datetime as dt

import streamlit as st

from traitement import FichierVigieInvalide, transformer

st.set_page_config(
    page_title="Planning escales paquebots - GPMLM",
    page_icon="🚢",
    layout="centered",
)

st.title("🚢 Planning des escales paquebots")
st.caption(
    "Génère automatiquement le fichier Excel destiné aux agents, "
    "à partir d'une extraction VIGIEsip (Prévisions Paquebots)."
)

st.divider()

# ---------------------------------------------------------------------------
# Étape 1 : dépôt du fichier
# ---------------------------------------------------------------------------
st.subheader("1. Déposez votre extraction VIGIEsip")
fichier = st.file_uploader(
    "Fichier .csv exporté depuis VIGIEsip (Agenda > Prévisions Paquebots > Édition CSV)",
    type=["csv"],
)

# ---------------------------------------------------------------------------
# Étape 2 : période de la saison
# ---------------------------------------------------------------------------
st.subheader("2. Période de la saison")

annee_actuelle = dt.date.today().year
mois_actuel = dt.date.today().month
annee_saison_defaut = annee_actuelle if mois_actuel >= 10 else annee_actuelle - 1

col1, col2 = st.columns(2)
with col1:
    date_debut = st.date_input(
        "Début de saison",
        value=dt.date(annee_saison_defaut, 10, 1),
        format="DD/MM/YYYY",
    )
with col2:
    date_fin = st.date_input(
        "Fin de saison",
        value=dt.date(annee_saison_defaut + 1, 9, 30),
        format="DD/MM/YYYY",
    )

titre_par_defaut = f"SAISON CROISIERES {date_debut.year}-{date_fin.year} - ESCALES GPMLM"
titre = st.text_input("Titre du document", value=titre_par_defaut)

st.divider()

# ---------------------------------------------------------------------------
# Étape 3 : génération et téléchargement
# ---------------------------------------------------------------------------
st.subheader("3. Générer le fichier")

if fichier is None:
    st.info("Déposez un fichier CSV ci-dessus pour activer la génération.")
else:
    if date_fin <= date_debut:
        st.error("La date de fin doit être après la date de début.")
    else:
        if st.button("Générer le planning", type="primary", use_container_width=True):
            try:
                with st.spinner("Génération du fichier en cours..."):
                    contenu_xlsx, nb_escales, postes_inconnus = transformer(
                        fichier, date_debut, date_fin, titre
                    )

                st.success(f"Fichier généré avec succès — {nb_escales} escales trouvées.")

                if postes_inconnus:
                    st.warning(
                        "Poste(s) non reconnu(s) dans l'extraction, conservé(s) "
                        f"tel(s) quel(s) dans le fichier final : {', '.join(postes_inconnus)}"
                    )

                nom_fichier = f"Planning_escales_{date_debut.year}-{date_fin.year}.xlsx"
                st.download_button(
                    label="📥 Télécharger le fichier Excel",
                    data=contenu_xlsx,
                    file_name=nom_fichier,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            except FichierVigieInvalide as e:
                st.error(f"Le fichier déposé n'a pas pu être traité : {e}")
            except Exception as e:
                st.error(
                    "Une erreur inattendue s'est produite pendant la génération. "
                    f"Détail technique : {e}"
                )

st.divider()
with st.expander("ℹ️ Comment obtenir le fichier CSV depuis VIGIEsip ?"):
    st.markdown(
        """
        1. Sur VIGIEsip, cliquer sur **Agenda** puis **Prévisions Paquebots**
        2. Définir les dates de l'extraction (une saison va du 1er octobre au 30 septembre)
        3. Cliquer sur **Impression**, puis **Édition CSV**
        4. Déposer ce fichier ci-dessus
        """
    )
