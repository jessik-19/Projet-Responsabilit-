# app_responsabilites.py
import streamlit as st
import pandas as pd
from io import BytesIO

# ======================================================
# APPARENCE
# ======================================================
def load_custom_css():
    st.markdown(
        """
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
        """,
        unsafe_allow_html=True,
    )

# ======================================================
# LECTURE FICHIER (EXCEL OU HTML DEGUISE)
# ======================================================
def _looks_like_html(data: bytes) -> bool:
    head = data[:2000].lstrip().lower()
    return head.startswith(b"<html") or b"<table" in head or b"<!doctype html" in head

def read_table_auto(uploaded_file, sheet_name=None):
    """
    Lit un fichier uploadé :
    - vrai Excel .xls -> xlrd
    - vrai Excel .xlsx -> openpyxl
    - faux .xls (HTML SharePoint) -> read_html (bs4/html5lib)
    """
    if uploaded_file is None:
        raise ValueError("Aucun fichier reçu.")

    data = uploaded_file.getvalue()
    if not data:
        raise ValueError("Fichier vide (0 octet).")

    name = (uploaded_file.name or "").lower()

    # --- Cas HTML déguisé ---
    if _looks_like_html(data):
        # pandas utilisera beautifulsoup4/html5lib si installés
        tables = pd.read_html(BytesIO(data))
        # prendre la 1ère table non vide
        for t in tables:
            if isinstance(t, pd.DataFrame) and not t.empty:
                return t
        raise ValueError("Fichier HTML détecté mais aucune table exploitable trouvée.")

    # --- Cas Excel ---
    if name.endswith(".xls"):
        return pd.read_excel(BytesIO(data), sheet_name=sheet_name, engine="xlrd")
    else:
        return pd.read_excel(BytesIO(data), sheet_name=sheet_name, engine="openpyxl")

# ======================================================
# NORMALISATION HEADER (quand colonnes = 0,1,2,...)
# ======================================================
def normalize_header_if_needed(df: pd.DataFrame) -> pd.DataFrame:
    """
    Si le fichier vient de read_html, souvent les colonnes deviennent 0..n.
    On essaie de retrouver la ligne qui contient "Nom utilisateur" et la mettre en header.
    """
    if df is None or df.empty:
        return df

    # si colonnes déjà textuelles (pas seulement 0..n), on laisse
    if all(isinstance(c, str) for c in df.columns) and any("nom" in str(c).lower() for c in df.columns):
        return df

    # Chercher une ligne qui ressemble à un header
    target_keywords = ["nom utilisateur", "responsabil", "application"]
    best_idx = None

    for i in range(min(len(df), 30)):  # on scanne max 30 premières lignes
        row_txt = " | ".join([str(x).strip().lower() for x in df.iloc[i].tolist()])
        if any(k in row_txt for k in target_keywords):
            best_idx = i
            break

    if best_idx is not None:
        new_cols = [str(x).strip() for x in df.iloc[best_idx].tolist()]
        df2 = df.iloc[best_idx + 1 :].copy()
        df2.columns = new_cols
        df2 = df2.reset_index(drop=True)
        return df2

    return df

# ======================================================
# EXTRACTION GRAND BACK (ancien ; ou nouveau multi-colonnes)
# ======================================================
def build_extraction_df(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Retourne un DF avec 2 colonnes:
    - Nom utilisateur
    - Responsabilite
    Supporte:
    - Ancien format: 1 colonne avec des ';'
    - Nouveau format: colonnes déjà séparées (ou header à reconstruire)
    """
    if df_raw is None or df_raw.empty:
        raise ValueError("Extraction vide ou illisible.")

    df_raw = normalize_header_if_needed(df_raw)

    # CAS 1: une seule colonne => split par ;
    if df_raw.shape[1] == 1:
        s = df_raw.iloc[:, 0].astype(str)
        df_split = s.str.split(";", expand=True)

        if df_split.shape[1] < 5:
            raise ValueError("Format 1-colonne détecté mais séparation ';' invalide (< 5 champs).")

        df_extraction = pd.DataFrame({
            "Nom utilisateur": df_split.iloc[:, 0],
            "Responsabilite": df_split.iloc[:, 4],
        })
        return df_extraction

    # CAS 2: multi-colonnes
    # Normaliser noms de colonnes
    cols = [str(c).strip() for c in df_raw.columns]
    df = df_raw.copy()
    df.columns = cols

    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.replace("\u00a0", " ", regex=False)  # nbsp
        .str.upper()
    )

    mapping = {
        "NOM UTILISATEUR": "Nom utilisateur",
        "UTILISATEUR": "Nom utilisateur",
        "USER": "Nom utilisateur",
        "USERNAME": "Nom utilisateur",

        "RESPONSABILITE": "Responsabilite",
        "RESPONSABILITÉ": "Responsabilite",
        "RESPONSABILITE GRAND BACK": "Responsabilite",
        "RESPONSABILITÉ GRAND BACK": "Responsabilite",
        "RESPONSABILITY": "Responsabilite",
        "RESPONSIBILITY": "Responsabilite",
    }

    df = df.rename(columns=mapping)

    # Si on n’a pas trouvé par noms, on fallback par POSITION (comme ton fichier)
    # D’après ton aperçu : A=Nom utilisateur, E=Responsabilité => index 0 et 4
    if ("Nom utilisateur" not in df.columns) or ("Responsabilite" not in df.columns):
        if df.shape[1] >= 5:
            df_extraction = pd.DataFrame({
                "Nom utilisateur": df.iloc[:, 0],
                "Responsabilite": df.iloc[:, 4],
            })
            return df_extraction
        else:
            raise ValueError(f"Nouveau format détecté mais colonnes insuffisantes: {list(df.columns)}")

    df_extraction = df[["Nom utilisateur", "Responsabilite"]].copy()
    return df_extraction

# ======================================================
# TRAITEMENT METIER
# ======================================================
def traitement_responsabilites(df_extraction, df_profils, df_users):
    # USERS
    df_users = df_users.copy()
    df_users.columns = df_users.columns.astype(str).str.strip()

    # flag (robuste)
    if any(c.upper() == "FLAG" for c in df_users.columns):
        col_flag = [c for c in df_users.columns if c.upper() == "FLAG"][0]
        df_users = df_users[df_users[col_flag] == 1]

    # Normaliser colonnes
    if "Nom utilisateur" not in df_users.columns and "NOM UTILISATEUR" in df_users.columns:
        df_users = df_users.rename(columns={"NOM UTILISATEUR": "Nom utilisateur"})
    if "Profil" not in df_users.columns and "PROFIL" in df_users.columns:
        df_users = df_users.rename(columns={"PROFIL": "Profil"})

    df_users["Nom utilisateur"] = df_users["Nom utilisateur"].astype(str).str.strip().str.upper()
    df_users["Profil"] = df_users["Profil"].astype(str).str.strip()

    # PROFILS
    df_profils = df_profils.copy()
    df_profils.columns = df_profils.columns.astype(str).str.strip()

    df_profils["Profil"] = df_profils["Profil"].astype(str).str.strip()
    df_profils["Responsabilite"] = df_profils["Responsabilite"].astype(str).str.strip().str.upper()

    dict_attendues = (
        df_profils.groupby("Profil")["Responsabilite"]
        .apply(set)
        .to_dict()
    )

    # EXTRACTION
    df_extraction = df_extraction.copy()
    df_extraction["Nom utilisateur"] = df_extraction["Nom utilisateur"].astype(str).str.strip().str.upper()
    df_extraction["Responsabilite"] = df_extraction["Responsabilite"].astype(str).str.strip().str.upper()

    resumes, manquants, en_trop = [], [], []

    for _, row in df_users.iterrows():
        user = row["Nom utilisateur"]
        profil = row["Profil"]

        attendues = dict_attendues.get(profil, set())
        reelles = set(df_extraction[df_extraction["Nom utilisateur"] == user]["Responsabilite"])
        nb_dans_grand_back = (df_extraction["Nom utilisateur"] == user).sum()

        manq = attendues - reelles
        trop = reelles - attendues

        resumes.append({
            "Utilisateur": user,
            "Profil": profil,
            "Resp. dans Grand Back": int(nb_dans_grand_back),
            "Resp. attendues (profil type)": len(attendues),
            "Resp. OK": len(attendues & reelles),
            "Resp. manquantes": len(manq),
            "Resp. en trop": len(trop),
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
        fichier_extraction = st.file_uploader("Extraction Grand Back FR", type=["xls", "xlsx", "html", "htm"])
        st.caption("Ancien format (1 colonne avec ;) OU nouveau format (table SharePoint/Excel Online)")
        st.image("images/exemple_extraction.png", use_container_width=True)

    with c2:
        fichier_profils = st.file_uploader("Profils types – Responsabilités Grand Back", type=["xlsx"])
        st.caption("Feuille attendue : Responsabilités Grand Back")
        st.image("images/exemple_acces.png", use_container_width=True)

    with c3:
        fichier_users = st.file_uploader("Liste utilisateurs Direction", type=["xlsx"])
        st.caption("Colonnes attendues : Nom utilisateur / Profil (+ Flag optionnel)")
        st.image("images/exemple_equipe.png", use_container_width=True)

    if not (fichier_extraction and fichier_profils and fichier_users):
        return

    # 1) EXTRACTION
    try:
        df_raw = read_table_auto(fichier_extraction, sheet_name=0)
        df_extraction = build_extraction_df(df_raw)

        # nettoyage
        df_extraction["Nom utilisateur"] = df_extraction["Nom utilisateur"].astype(str).str.strip()
        df_extraction["Responsabilite"] = df_extraction["Responsabilite"].astype(str).str.strip()
        df_extraction = df_extraction[(df_extraction["Nom utilisateur"] != "") & (df_extraction["Responsabilite"] != "")]

        st.success("✅ Extraction Grand Back chargée correctement")
        st.caption(f"Lignes extraction : {len(df_extraction)}")

    except Exception as e:
        st.error(f"❌ Impossible de lire le fichier Extraction Grand Back : {e}")
        st.stop()

    # 2) PROFILS (sheet spécifique)
    try:
        xls = pd.ExcelFile(BytesIO(fichier_profils.getvalue()), engine="openpyxl")
        feuilles = xls.sheet_names
        st.info(f"📄 Feuilles détectées dans le fichier Profils : {', '.join(feuilles)}")

        feuille_attendue = "Responsabilités Grand Back"
        if feuille_attendue not in feuilles:
            st.error("❌ Mauvais fichier Profils déposé")
            st.stop()

        df_profils = pd.read_excel(BytesIO(fichier_profils.getvalue()), sheet_name=feuille_attendue, engine="openpyxl")

    except Exception as e:
        st.error(f"❌ Impossible de lire le fichier Profils : {e}")
        st.stop()

    # normalisation profils
    df_profils.columns = df_profils.columns.astype(str).str.strip().str.upper()
    mapping_profils = {
        "PROFIL TYPE": "Profil",
        "PROFIL": "Profil",
        "RESPONSABILITÉ GRAND BACK": "Responsabilite",
        "RESPONSABILITE GRAND BACK": "Responsabilite",
        "RESPONSABILITÉ": "Responsabilite",
        "RESPONSABILITE": "Responsabilite",
    }
    df_profils = df_profils.rename(columns=mapping_profils)
    if not {"Profil", "Responsabilite"}.issubset(df_profils.columns):
        st.error(f"❌ Colonnes incorrectes dans Profils. Trouvées: {list(df_profils.columns)}")
        st.stop()
    st.success("✅ Colonnes Profils reconnues et normalisées")

    # 3) USERS
    try:
        df_users = pd.read_excel(BytesIO(fichier_users.getvalue()), sheet_name=0, engine="openpyxl")
    except Exception as e:
        st.error(f"❌ Impossible de lire le fichier Utilisateurs : {e}")
        st.stop()

    df_users.columns = df_users.columns.astype(str).str.strip().str.upper()
    df_users = df_users.rename(columns={
        "LISTE DES UTILISATEURS": "Nom utilisateur",
        "NOM UTILISATEUR": "Nom utilisateur",
        "UTILISATEUR": "Nom utilisateur",
        "NOM": "Nom utilisateur",
        "PROFIL TYPE": "Profil",
        "PROFIL": "Profil",
        "FLAG": "Flag",
    })

    if not {"Nom utilisateur", "Profil"}.issubset(df_users.columns):
        st.error(f"❌ Colonnes manquantes dans Utilisateurs. Trouvées: {list(df_users.columns)}")
        st.stop()
    st.success("✅ Fichier Utilisateurs conforme")

    # 4) TRAITEMENT
    df_resume, df_manq, df_trop = traitement_responsabilites(df_extraction, df_profils, df_users)

    st.subheader("2. Résultats")
    st.dataframe(df_resume, use_container_width=True)

    # 5) EXPORT
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_resume.to_excel(writer, sheet_name="Resume", index=False)
        df_manq.to_excel(writer, sheet_name="Manquantes", index=False)
        df_trop.to_excel(writer, sheet_name="En_trop", index=False)
    output.seek(0)

    st.download_button(
        "📥 Télécharger le rapport Excel",
        data=output,
        file_name="rapport_controle_grand_back.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

if __name__ == "__main__":
    main()