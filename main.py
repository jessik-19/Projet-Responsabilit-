# app_responsabilites.py
import streamlit as st
import pandas as pd
from io import BytesIO
# --- Apparence ---
def load_custom_css():
   st.markdown("""
<style>
       .main {background-color: #f9f9f9;}
       h1, h2, h3 {
           color: #1a1a1a;
           font-family: 'Segoe UI', sans-serif;
       }
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

# CSS personnalisé
def load_css():
   with open("style.css") as f:
       st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# --- Traitement principal ---
def traitement_responsabilites(df_extraction, df_profils, df_direction):
    
    # Préparer les utilisateurs à analyser (flag == 1)
   df_utilisateurs = df_direction[df_direction.iloc[:, 2] == 1].copy()
   df_utilisateurs["Nom utilisateur"] = df_utilisateurs.iloc[:, 0].astype(str).str.strip().str.upper()
   df_utilisateurs["Profil"] = df_utilisateurs.iloc[:, 1].astype(str).str.strip()
   
   # Nettoyer les responsabilités attendues par profil
   df_profils_clean = df_profils.copy()
   df_profils_clean.columns = df_profils_clean.iloc[0]
   df_profils_clean = df_profils_clean[1:]
   df_profils_clean = df_profils_clean.rename(columns={df_profils_clean.columns[0]: "Profil"})
   df_profils_clean = df_profils_clean.set_index("Profil")
   dict_responsabilites_attendues = {}
   for profil in df_profils_clean.index.unique():
       valeurs = df_profils_clean.loc[profil].values.flatten().tolist()
       valeurs_propres = [str(v).strip().upper() for v in valeurs if pd.notna(v)]
       dict_responsabilites_attendues[profil] = set(valeurs_propres)
   
   # Nettoyer les responsabilités réelles
   df_extraction["Nom utilisateur"] = df_extraction["Nom utilisateur"].astype(str).str.strip().str.upper()
   df_extraction["Responsabilité"] = df_extraction["Responsabilité"].astype(str).str.strip().str.upper()
   
   # Comparaison
   lignes_resumes, lignes_manquants, lignes_en_trop = [], [], []
   for _, row in df_utilisateurs.iterrows():
       user = row["Nom utilisateur"]
       profil = row["Profil"]
       attendues = dict_responsabilites_attendues.get(profil, set())
       reelles = set(df_extraction[df_extraction["Nom utilisateur"] == user]["Responsabilité"])
       manquantes = attendues - reelles
       en_trop = reelles - attendues
       lignes_resumes.append({
           "Utilisateur": user,
           "Profil": profil,
           "Nb responsabilités attendues": len(attendues),
           "Nb présentes": len(reelles & attendues),
           "Nb manquantes": len(manquantes),
           "Nb en trop": len(en_trop)
       })
       for resp in sorted(manquantes):
           lignes_manquants.append({"Utilisateur": user, "Responsabilité manquante": resp})
       for resp in sorted(en_trop):
           lignes_en_trop.append({"Utilisateur": user, "Responsabilité non attendue": resp})
   df_resume = pd.DataFrame(lignes_resumes).sort_values("Utilisateur")
   df_manquants = pd.DataFrame(lignes_manquants).sort_values(["Utilisateur", "Responsabilité manquante"])
   df_en_trop = pd.DataFrame(lignes_en_trop).sort_values(["Utilisateur", "Responsabilité non attendue"])
   return df_resume, df_manquants, df_en_trop


# --- App principale ---
def main():
    st.set_page_config(page_title=" Contrôle Grand Back", layout="wide")
    load_custom_css()
   
   
    col1, col2 = st.columns([1, 4])  # Logo à gauche, texte à droite
    with col1:
        st.image("logo_accor.png", width=120)
    with col2:
        st.markdown("""
    <div style='padding-top: 15px;'>
    <h1 style='text-align: center;margin-bottom: 0;'>Application interne - Contrôle des responsabilités Grand Back</h1>
    </div>
    """, unsafe_allow_html=True)
    # Texte "Bienvenue" centré
    st.markdown("""
    <div style='text-align: center; margin-top: -10px; font-size: 18px;'>
        Vérifie les responsabilités attribuées aux utilisateurs selon leur</strong> profil type.
    </div>
    """, unsafe_allow_html=True)

    st.markdown("<br><br>", unsafe_allow_html=True)#ajoutez de l'espace 

    st.subheader("1. Importer vos fichiers")

    st.markdown("<br>", unsafe_allow_html=True)#ajoutez de l'espace 
   
    col1, col2, col3 = st.columns(3)
    with col1:
        fichier_extraction = st.file_uploader(" Extraction Grand Back FR", type="xlsx")
        st.caption("Fichiers Grand Back contenant tous les Responsabilitén par Utilisateur.")
        st.image("images/exemple_extraction.png", use_container_width=True)
    with col2:
        fichier_profils = st.file_uploader(" Accès aux outils - Profils types et outils", type="xlsx")
        st.caption("Fichiers Accès aux outils contenant tous les Profils et Responsabilitén.")
        st.image("images/exemple_acces.png", use_container_width=True)
    with col3:
        fichier_utilisateurs = st.file_uploader(" Liste utilisateurs Direction", type="xlsx")
        st.caption("Fichiers avec la liste des utilisateurss avec une colonne de flag.")
        st.image("images/exemple_equipe.png", use_container_width=True)
    if fichier_extraction and fichier_profils and fichier_utilisateurs:

        try:

            df_extraction = pd.read_excel(fichier_extraction)

            #  Nettoyage complet

            df_extraction.columns = (

                df_extraction.columns

                .astype(str)

                .str.strip()

                .str.replace("\xa0", " ", regex=True)

                .str.replace("Unnamed: ", "", regex=True)

                .str.replace(r"\d+", "", regex=True)

                .str.normalize('NFKD')

                .str.encode('ascii', errors='ignore')

                .str.decode('utf-8')

            )

            colonnes_trouvees = [c for c in df_extraction.columns if c.strip() != ""]

            #  Seulement les colonnes vraiment utilisées dans ton traitement

            colonnes_utiles = {"Nom utilisateur", "Responsabilite"}  

            colonnes_manquantes = colonnes_utiles - set(colonnes_trouvees)

            if colonnes_manquantes:

                st.warning(f"Colonnes manquantes dans le fichier Extraction Grand Back: {', '.join(colonnes_manquantes)}")
                st.info(f"Colonnes détectées : {', '.join(colonnes_trouvees)}")

            else:

                st.success("Fichier Référence correctement chargé !")

                st.info(f"Colonnes détectées : {', '.join(colonnes_trouvees)}")

        except Exception as e:

            st.error(f"Erreur dans le fichier Référence : {e}")
            


        # === 2️Fichier Back ===

        try:

            xls = pd.ExcelFile(fichier_profils)
            if "Responsabilités Grand Back" in xls.sheet_names:
                df_profils_clean = pd.read_excel(xls, sheet_name="Responsabilités Grand Back")
                st.info("Feuille chargée : Responsabilités Grand Back")
            else:
                st.error("La feuille 'Responsabilités Grand Back' est introuvable dans le fichier Accès aux outils.")
                #st.stop()

            df_profils_clean.columns = (

                df_profils_clean.columns

                .astype(str)

               .str.strip()

               .str.replace("\xa0", " ", regex=True)

               .str.replace("Unnamed: ", "", regex=True)

               .str.replace(r"\d+", "", regex=True)

               .str.normalize('NFKD')

               .str.encode('ascii', errors='ignore')

               .str.decode('utf-8')

            )

            colonnes_trouvees = [c for c in df_profils_clean.columns if c.strip() != ""]

            colonnes_utiles = {"Profil type", "Responsabilite Grand Back"}  # celles que ton traitement utilise vraiment

            colonnes_manquantes = colonnes_utiles - set(colonnes_trouvees)

            if colonnes_manquantes:

                st.warning(f"Colonnes manquantes dans le fichier Accès aux outils: {', '.join(colonnes_manquantes)}")
                st.info(f"Colonnes détectées : {', '.join(colonnes_trouvees)}")

            else:

                st.success("Fichier Back correctement chargé !")

                st.info(f"Colonnes détectées : {', '.join(colonnes_trouvees)}")

        except Exception as e:

           st.error(f"Erreur dans le fichier Back : {e}")


        # === 3️Fichier Équipe ===

        try:

            df_utilisateurs = pd.read_excel(fichier_utilisateurs)

            df_utilisateurs.columns = (

                df_utilisateurs.columns

                .astype(str)

                .str.strip()

                .str.replace("\xa0", " ", regex=True)

                .str.replace("Unnamed: ", "", regex=True)

                .str.replace(r"\d+", "", regex=True)

                .str.normalize('NFKD')

                .str.encode('ascii', errors='ignore')

                .str.decode('utf-8')

            )

            colonnes_trouvees = [c for c in df_utilisateurs.columns if c.strip() != ""]

            colonnes_utiles = {"LISTE DES UTILISATEURS", "PROFIL TYPE", "FLAG DIRECTION COMPTABLE CORPORATE"}

            colonnes_manquantes = colonnes_utiles - set(colonnes_trouvees)

            if colonnes_manquantes:

                st.warning(f"Colonnes manquantes dans le fichier liste utilisateurs Direction : {', '.join(colonnes_manquantes)}")
                st.info(f"Colonnes détectées : {', '.join(colonnes_trouvees)}")

            else:

                st.success("Fichier Équipe correctement chargé !")

                st.info(f"Colonnes détectées : {', '.join(colonnes_trouvees)}")

        except Exception as e:

            st.error(f"Erreur dans le fichier Équipe : {e}")


        try:
            df_extraction = pd.read_excel(fichier_extraction, usecols=[0, 4])
            df_extraction.columns = ["Nom utilisateur", "Responsabilité"]
            df_profils = pd.read_excel(fichier_profils, sheet_name="Responsabilités Grand Back")
            df_utilisateurs = pd.read_excel(fichier_utilisateurs)
            st.success(" Fichiers chargés avec succès")
            df_resume, df_manquants, df_en_trop = traitement_responsabilites(
                df_extraction, df_profils, df_utilisateurs
            )
            st.header("2️. Résultats de la comparaison")
            st.subheader(" a. Vue d'ensemble")
            st.dataframe(df_resume, use_container_width=True)
            st.subheader(" b. Responsabilités manquantes")
            st.dataframe(df_manquants, use_container_width=True)
            st.subheader(" c. Responsabilités non attendues")
            st.dataframe(df_en_trop, use_container_width=True)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
               df_resume.to_excel(writer, sheet_name="Résumé", index=False)
               df_manquants.to_excel(writer, sheet_name="Manquantes", index=False)
               df_en_trop.to_excel(writer, sheet_name="En trop", index=False)
            output.seek(0)
            st.download_button(
                label=" Télécharger le rapport Excel",
                data=output,
                file_name="rapport_responsabilites.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
           )
        except Exception as e:
            st.error(f"Une erreur est survenue : {e}")
if __name__ == "__main__":
   main()