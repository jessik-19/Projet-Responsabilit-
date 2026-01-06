# app_responsabilites.py
import streamlit as st
import pandas as pd
from io import BytesIO

# ======================================================
# APPARENCE
# ======================================================
def load_custom_css():
    st.markdown("""
    <style>
        .main {background-color: #f9f9f9;}
        h1, h2, h3 {font-family: 'Segoe UI', sans-serif;}
        .stButton>button {
            background-color: #2d86ff;
            color: white;
            border-radius: 6px;
            padding: 8px 16px;
            border: none;
        }
        .stDownloadButton>button {
            background-color: #34c759;
            color: white;
            border-radius: 6px;
        }
    </style>
    """, unsafe_allow_html=True)

# ======================================================
# TRAITEMENT METIER (INCHANGÉ)
# ======================================================
def traitement_responsabilites(df_extraction, df_profils, df_users):

    df_users = df_users.copy()
    df_users.columns = df_users.columns.astype(str).str.strip()

    if "Flag" in df_users.columns:
        df_users = df_users[df_users["Flag"] == 1]

    df_users["Nom utilisateur"] = df_users["Nom utilisateur"].astype(str).str.strip().str.upper()
    df_users["Profil"] = df_users["Profil"].astype(str).str.strip()

    df_profils = df_profils.copy()
    df_profils.columns = df_profils.columns.astype(str).str.strip()
    df_profils["Profil"] = df_profils["Profil"].astype(str).str.strip()
    df_profils["Responsabilite"] = df_profils["Responsabilite"].astype(str).str.strip().str.upper()

    dict_attendues = (
        df_profils
        .groupby("Profil")["Responsabilite"]
        .apply(set)
        .to_dict()
    )

    df_extraction["Nom utilisateur"] = df_extraction["Nom utilisateur"].astype(str).str.strip().str.upper()
    df_extraction["Responsabilite"] = df_extraction["Responsabilite"].astype(str).str.strip().str.upper()

    resumes, manquants, en_trop = [], [], []

    for _, row in df_users.iterrows():
        user = row["Nom utilisateur"]
        profil = row["Profil"]

        attendues = dict_attendues.get(profil, set())
        reelles = set(df_extraction[df_extraction["Nom utilisateur"] == user]["Responsabilite"])
        nb_dans_grand_back = len(df_extraction[df_extraction["Nom utilisateur"] == user])

        manq = attendues - reelles
        trop = reelles - attendues

        resumes.append({
            "Utilisateur": user,
            "Profil": profil,
            "Resp. dans Grand Back": nb_dans_grand_back,
            "Resp. attendues (profil type)": len(attendues),
            "Resp. OK": len(attendues & reelles),
            "Resp. manquantes": len(manq),
            "Resp. en trop": len(trop)
        })

        for r in sorted(manq):
            manquants.append({"Utilisateur": user, "Responsabilite manquante": r})

        for r in sorted(trop):
            en_trop.append({"Utilisateur": user, "Responsabilite non attendue": r})

    return (
        pd.DataFrame(resumes).sort_values("Utilisateur"),
        pd.DataFrame(manquants),
        pd.DataFrame(en_trop)
    )

# ======================================================
# APPLICATION STREAMLIT
# ======================================================
def main():
    st.set_page_config(page_title="Contrôle Grand Back", layout="wide")
    load_custom_css()

    col1, col2 = st.columns([1, 4])
    with col1:
        st.image("logo_accor.png", width=120)
    with col2:
        st.markdown("<h1 style='text-align:center;'>Contrôle des responsabilités Grand Back</h1>", unsafe_allow_html=True)

    st.subheader("1. Import des fichiers")

    c1, c2, c3 = st.columns(3)

    with c1:
        fichier_extraction = st.file_uploader("Extraction Grand Back FR", type="xlsx")
        st.caption("Fichier Grand Back brut (1 colonne avec des ;)")
        st.image("images/exemple_extraction.png", use_container_width=True)

    with c2:
        fichier_profils = st.file_uploader("Profils types – Responsabilités Grand Back", type="xlsx")
        st.caption("Feuille attendue : Responsabilités Grand Back")
        st.image("images/exemple_acces.png", use_container_width=True)

    with c3:
        fichier_users = st.file_uploader("Liste utilisateurs Direction", type="xlsx")
        st.caption("Colonnes attendues : Nom utilisateur / Profil (+ Flag optionnel)")
        st.image("images/exemple_equipe.png", use_container_width=True)

    if fichier_extraction and fichier_profils and fichier_users:

        # ==================================================
        # EXTRACTION GRAND BACK
        # ==================================================
        df_raw = pd.read_excel(fichier_extraction)

        if df_raw.shape[1] != 1:
            st.error("❌ Le fichier Extraction Grand Back doit contenir une seule colonne.")
            st.info(f"Colonnes trouvées : {list(df_raw.columns)}")
            st.stop()

        df_split = df_raw.iloc[:, 0].astype(str).str.split(";", expand=True)

        if df_split.shape[1] < 5:
            st.error("❌ Format invalide : séparation par ';' incorrecte.")
            st.stop()

        df_extraction = pd.DataFrame({
            "Nom utilisateur": df_split[0],
            "Responsabilite": df_split[4]
        })

        df_extraction = df_extraction[
            (df_extraction["Nom utilisateur"] != "") &
            (df_extraction["Responsabilite"] != "")
        ]

        st.success(" Extraction Grand Back chargée correctement")

  
        # ==================================================
        # PROFILS — UX METIER (FEUILLES + COLONNES)
        # ==================================================
        xls = pd.ExcelFile(fichier_profils)
        feuilles_trouvees = xls.sheet_names

        st.info("📄 Feuille détectée dans le fichier Profils")

        feuille_attendue = "Responsabilités Grand Back"

        # --- Contrôle de la feuille ---
        if feuille_attendue not in feuilles_trouvees:
            st.error("❌ Mauvais fichier Profils déposé")

            st.markdown(
                f"""
                **Feuille attendue :**
                - {feuille_attendue}

                **Feuille détectée :**
                - {feuilles_trouvees[0]}

                👉 Merci de déposer le fichier **Accès aux outils – Profils types et outils**.
                """
            )
            st.stop()

        st.success(" Feuille correcte détectée")

        # --- Lecture de la feuille ---
        df_profils = pd.read_excel(
            fichier_profils,
            sheet_name="Responsabilités Grand Back"
        )

        # Nettoyage des noms
        df_profils.columns = (
            df_profils.columns.astype(str)
            .str.strip()
            .str.upper()
        )

        # 🔁 MAPPING DES NOMS METIER → NOMS INTERNES
        mapping_colonnes_profils = {
            "PROFIL TYPE": "Profil",
            "PROFIL": "Profil",
            "RESPONSABILITÉ GRAND BACK": "Responsabilite",
            "RESPONSABILITE GRAND BACK": "Responsabilite",
            "RESPONSABILITÉ": "Responsabilite",
            "RESPONSABILITE": "Responsabilite"
        }

        df_profils = df_profils.rename(columns=mapping_colonnes_profils)

        #  Vérification métier
        colonnes_attendues = {"Profil", "Responsabilite"}
        colonnes_trouvees = set(df_profils.columns)

        if not colonnes_attendues.issubset(colonnes_trouvees):
            st.error("❌ Colonnes incorrectes dans le fichier Profils")
            st.info(f"Colonnes attendues : {', '.join(colonnes_attendues)}")
            st.info(f"Colonnes trouvées : {', '.join(colonnes_trouvees)}")
            st.stop()

        st.success(" Colonnes Profils reconnues et normalisées")


        colonnes_manquantes = colonnes_attendues - colonnes_trouvees

        if colonnes_manquantes:
            st.error("❌ Colonnes manquantes dans le fichier Profils")
            st.markdown(
                f"""
                **Colonnes manquantes :**
                - {', '.join(colonnes_manquantes)}

                👉 Merci de vérifier le fichier Profils.
                """
            )
            st.stop()

        st.success(" Colonnes du fichier Profils conformes")


        # ==================================================
        # UTILISATEURS — CONTRÔLE COLONNES
        # ==================================================
        df_users = pd.read_excel(fichier_users)
        df_users.columns = df_users.columns.astype(str).str.strip().str.upper()

        df_users = df_users.rename(columns={
            "LISTE DES UTILISATEURS": "Nom utilisateur",
            "NOM UTILISATEUR": "Nom utilisateur",
            "UTILISATEUR": "Nom utilisateur",
            "NOM": "Nom utilisateur",
            "PROFIL TYPE": "Profil",
            "PROFIL": "Profil"
        })

        colonnes_attendues_users = {"Nom utilisateur", "Profil"}
        colonnes_trouvees_users = set(df_users.columns)

        if not colonnes_attendues_users.issubset(colonnes_trouvees_users):
            st.error("❌ Colonnes manquantes dans le fichier Utilisateurs")
            st.info(f"Colonnes attendues : {colonnes_attendues_users}")
            st.info(f"Colonnes trouvées : {colonnes_trouvees_users}")
            st.stop()

        st.success(" Fichier Utilisateurs conforme")

        # ==================================================
        # TRAITEMENT
        # ==================================================
        df_resume, df_manq, df_trop = traitement_responsabilites(
            df_extraction, df_profils, df_users
        )
        st.write("")
        st.write("")
        st.subheader("2. Résultats")
        st.dataframe(df_resume, use_container_width=True)
        #st.dataframe(df_manq, use_container_width=True)
        #st.dataframe(df_trop, use_container_width=True)

        # ==================================================
        # EXPORT
        # ==================================================
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_resume.to_excel(writer, sheet_name="Resume", index=False)
            df_manq.to_excel(writer, sheet_name="Manquantes", index=False)
            df_trop.to_excel(writer, sheet_name="En_trop", index=False)

        output.seek(0)
        st.write("")
        st.write("")
        st.download_button(
            "Télécharger le rapport Excel",
            data=output,
            file_name="rapport_controle_grand_back.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

if __name__ == "__main__":
    main()
