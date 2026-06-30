"""
Application Streamlit : génère le planning des escales paquebots (Excel)
à partir d'une extraction VIGIEsip (.csv).

Lancer en local :
    streamlit run app.py
"""

import streamlit as st

from traitement import FichierVigieInvalide, transformer

st.set_page_config(
    page_title="Planning escales paquebots - GPMLM",
    layout="centered",
)

st.title("Planning des escales paquebots")
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

titre_personnalise = st.text_input(
    "Titre du document (laisser vide pour un titre automatique)",
    value="",
)

st.divider()

# ---------------------------------------------------------------------------
# Étape 2 : génération et téléchargement
# ---------------------------------------------------------------------------
st.subheader("2. Générer le fichier")

if fichier is None:
    st.info("Déposez un fichier CSV ci-dessus pour activer la génération.")
else:
    if st.button("Générer le planning", type="primary", use_container_width=True):
        try:
            with st.spinner("Génération du fichier en cours..."):
                titre = titre_personnalise.strip() or None
                contenu_xlsx, nb_escales, postes_inconnus, date_debut, date_fin = transformer(
                    fichier, titre
                )

            st.success(
                f"Fichier généré avec succès — {nb_escales} escales trouvées, "
                f"saison du {date_debut.strftime('%d/%m/%Y')} au {date_fin.strftime('%d/%m/%Y')}."
            )

            if postes_inconnus:
                st.warning(
                    "Poste(s) non reconnu(s) dans l'extraction, conservé(s) "
                    f"tel(s) quel(s) dans le fichier final : {', '.join(postes_inconnus)}"
                )

            nom_fichier = f"Planning_escales_{date_debut.year}-{date_fin.year}.xlsx"
            st.download_button(
                label="Télécharger le fichier Excel",
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
with st.expander("Comment obtenir le fichier CSV depuis VIGIEsip ?"):
    st.markdown(
        """
        1. Sur VIGIEsip, cliquer sur **Agenda** puis **Prévisions Paquebots**
        2. Définir les dates de l'extraction (une saison va du 1er octobre au 30 septembre)
        3. Cliquer sur **Impression**, puis **Édition CSV**
        4. Déposer ce fichier ci-dessus
        """
    )
