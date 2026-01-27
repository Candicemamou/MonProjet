
import re
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import plotly.graph_objects as go
from pandas import DataFrame as _DataFrame
import os
import io
import zipfile
import unicodedata


st.set_page_config(page_title='Analyse Fiches IA4Doc', layout='wide')


#Nettoie les cellules des fiches de test
def _normalize_text(s: object) -> str:
    """Normalise un texte (minuscule, sans accents, espaces simplifiés)."""
    if s is None:
        return ""
    s= str(s).strip().lower()
    # Retire les accents
    s=unicodedata.normalize("NFKD", s)
    s= "".join(ch for ch in s if not unicodedata.combining(ch))
    s= re.sub(r"\s+", " ", s)
    return s

def verdict_to_score(verdict: object) -> float:
    """Mappe strictement le verdict vers un score : bon=1, partiel=0.5, mauvais=0."""
    if not isinstance(verdict, str):
        return np.nan
    
    v = verdict.strip() # On garde juste le retrait des espaces inutiles
    
    if v == "Bon":
        return 1.0
    if v == "Partiellement bon":
        return 0.5
    if v == "Mauvais":
        return 0.0
        
    return np.nan # Si c'est vide ou écrit autrement

# ---------------------------------------------------------------------
def infer_fo_cols(df: pd.DataFrame) -> list[str]:
    """Détecte automatiquement les colonnes fonctions outils (fo1, fo2, ...), triées numériquement"""
    cols = []
    for c in df.columns:
        cs = str(c).strip()
        if re.match(r"^fo\d+$", cs):
            cols.append(cs)
    def _k(x: str) -> int:
        m = re.search(r"(\d+)$", x)
        return int(m.group(1)) if m else 0
    return sorted(cols, key=_k)


def excel_sheet_name(title: str) -> str:
    """Nom des tableaux-> feuilles pour export excel (même noms que ceux des tableaux affichés sur le streamlit) : 
    - max 31 caractères
    - suppression des caractères interdits : : \ / ? * [ ]
    """
    if not isinstance(title, str):
        title = str(title)
    # Caractères interdits Excel
    title = re.sub(r"[:\\/\?\*\[\]]", "", title)
    title = title.strip().strip("'")
    if not title:
        title = "Feuil1"
    return title[:31]


# =====================================================================
# STYLES / COULEURS
# =====================================================================

#Couleurs pour les verdicts
def color_verdict(val):
    if not isinstance(val, str):
        return ""
    v = val.strip().lower()
    if v == "bon":
        return "background-color: #c6efce; color: #006100; font-weight: bold;"
    if "partiellement" in v:
        return "background-color: #ffe699; color: #7f6000; font-weight: bold;"
    if "mauvais" in v:
        return "background-color: #fd4a4a; color: #9c0006; font-weight: bold;"
    return ""

# Regarde dans le tableau avec le nb de tests et performance à atteindre
def extract_pct(cell):
    if not isinstance(cell, str):
        return None
    cell = cell.strip()
    if not cell or cell in ("tbd", "NA"):
        return None
    if "(" not in cell or "%" not in cell:
        return None
    try:
        inside = cell.split("(")[1].split(")")[0]   # ex: '0.7%' ou '25.0%'
        inside = inside.replace("%", "").replace(",", ".")
        return float(inside)
    except Exception:
        return None


def color_tab3(val):
    pct = extract_pct(val)
    if pct is None:
        return ""
    # pct est en pourcentage (0–100)
    if pct < 25:
        return "background-color: #fd4a4a;"               # rouge
    elif pct < 50:
        return "background-color: #f4b183;"               # orange
    elif pct < 75:
        return "background-color: #fff2cc;"               # jaune
    elif pct < 90:
        return "background-color: #c6efce;"               # vert clair
    else:
        return "background-color: #00b050; color: white;" # vert foncé


#tableau justesse
def build_tab2bis(data, ref_base):
    #seulement les fs réellement testés
    fs_testes = data["fs_id"].astype(str).str.strip().unique()
    tab2bis = ref_base[ref_base["N° fs"].isin(fs_testes)].copy()

    #Calcul de la moyenne des scores (0 à 100) avec bon=100%, partiel=50%, mauvais=0%
    tmp = data.copy()
    tmp["__score__"] = tmp["verdict_doc"].apply(verdict_to_score) * 100.0
    taux_fs = (
        tmp.groupby("fs_id")["__score__"]
        .mean()
        .rename(lambda x: str(x).strip())
        .to_dict()
    )

    # Remplit les cases avec le % calculé
    fo_cols = infer_fo_cols(ref_base)
    def fill_cell(row, fo):
        val = row[fo]
        if val != "x":
            return ""
        fs = row["N° fs"]
        actual = taux_fs.get(fs)
        if actual is None or pd.isna(actual):
            return ""
        return f"{actual:.1f}%"

    for fo in fo_cols:
        tab2bis[fo] = tab2bis.apply(lambda r: fill_cell(r, fo), axis=1)

    return tab2bis, fo_cols


def style_tab2bis(df, fo_cols):
    """Applique les couleurs sur le Tableau Justesse."""
    styles = _DataFrame("", index=df.index, columns=df.columns)

    for i in df.index:
        for fo in fo_cols:
            val = df.at[i, fo]
            if not isinstance(val, str) or not val:
                continue
            if val in ("tbd", "NA"):
                continue
            try:
                pct = float(val.replace("%", "").replace(",", "."))
            except Exception:
                continue
            #en fonction du pourcentage    
            if pct < 25:
                styles.at[i, fo] = "background-color: #fd4a4a;"
            elif pct < 50:
                styles.at[i, fo] = "background-color: #f4b183;"
            elif pct < 75:
                styles.at[i, fo] = "background-color: #fff2cc;"
            elif pct < 90:
                styles.at[i, fo] = "background-color: #c6efce;"
            else:
                styles.at[i, fo] = "background-color: #00b050; color: white;"

    return styles


def norm_verdict(x: str) -> str:
    """Simplifie le verdict pour les statistiques (compter combien de 'bon', 'mauvais'...)"""
    if not isinstance(x, str):
        return "none"
    x = x.strip().lower()
    if x in ("", "nan", "none"):
        return "none"
    if x == "bon":
        return "bon"
    if "partiel" in x:
        return "partiel"
    if "mauvais" in x:
        return "mauvais"
    return "none"



def extract_modification(ref):
    """Extrait le code modification (ex: m01) depuis la référence."""
    if not isinstance(ref, str):
        return ""
    ref = ref.strip()
    match = re.search(r"(m\d{2})$", ref, re.IGNORECASE)
    return match.group(1) if match else ""


def extract_doc_id(ref: str) -> str:
    """Extrait un identifiant de document depuis la référence (nom de fichier).
    Heuristique : cherche d'abord un motif explicite (doc/document/coedm/ref + chiffres),
    sinon prend la première séquence de >=4 chiffres, sinon retourne la référence entière.
    """
    if not isinstance(ref, str):
        return ""
    s = ref.strip()
    m = re.search(r"(?:doc|document|coedm|ref)[-_ ]*(\d+)", s, flags=re.IGNORECASE)
    if m:
        return m.group(1)
    m = re.search(r"(\d{4,})", s)
    if m:
        return m.group(1)
    return s



# =====================================================================
# PARSE FICHE
# =====================================================================

#permet de lire les cases des fiches de test
def parse_fiche(file):
    filename = file.name

    if filename.startswith("~$"):
        return pd.DataFrame()  # fichier temporaire Excel => on ignore (j'ai déja rencontré le cas, d'où cette sécurité)

    # Exemples de noms :
    # - fs60-IA-v00-TUF-FFP_PV_...-Fiche-v01-CM.xlsx
    # - fs1-IA-v00-NR-...-Fiche-v01-CM.xlsx (fs 1 à 3 chiffres)
    filename_clean = str(filename).strip()

    # --- 1. Analyse du NOM du fichier ---
    # Vérifie si le nom respecte le format strict fsXX-IA-vXX..
    pat = r"^(fs\d{1,3})-IA-v\d+-(TUF|NR)-(.+)-Fiche-v\d+-[A-Za-z]+\.xlsx$" # si le nom des fiches change, c'est ici qu'il faut faire les modifs !
    m = re.match(pat, filename_clean, flags=re.IGNORECASE)
    if not m:
        # Si le nom est mauvais, le script s'arrête ici pour ce fichier
        raise ValueError(
            "Nom de fichier invalide. Format attendu : "
            "fs<1-3 chiffres>-IA-vXX-(TUF|NR)-<ref>-Fiche-vXX-<initiales>.xlsx"
        )

    fs_id = m.group(1).lower()
    fiche_type = m.group(2).upper()
    ref_from_filename = m.group(3)
    doc_id = extract_doc_id(ref_from_filename)

    # après upload des fiches de test, l'outil va lire la feuille "Template Fiche de Test" <- si le nom de la feuille excel change c'est ici qu'il faut modifier 
    df = pd.read_excel(file, sheet_name="Template Fiche de Test", header=None)

    # -----------------------------------------------------------------
    # Détection du décalage : ligne 18 contient-elle "*Searchable:" ? (après la nouvelle version du template des fiches de test)
    # -----------------------------------------------------------------
    has_searchable_row = False
    try:
        row18 = df.iloc[17]  # ligne 18 (index 17)
        for val in row18:
            if isinstance(val, str) and "Searchable" in val:
                has_searchable_row = True
                break
    except Exception:
        pass

    offset = 0 if has_searchable_row else -1

    def r(base_row):
        # applique le décalage seulement après la ligne 18
        if base_row > 17:
            return base_row + offset
        return base_row

    # ===========================
    # Champs "avant 18" Récupération des infos générales (En-tête de fiche)
    # ===========================
    type_test = df.iloc[4, 1] if pd.notna(df.iloc[4, 1]) else None
    type_doc = df.iloc[14, 1] if pd.notna(df.iloc[14, 1]) else None
    ref_coedm = ref_from_filename

    # Label fonctionnalité (ligne 24 → index 23, après 18 → offset)
    label_fonctionnalite = None
    try:
        val_lbl = df.iloc[r(23), 1]
        if pd.notna(val_lbl):
            label_fonctionnalite = val_lbl
    except Exception:
        pass

    if label_fonctionnalite is None or (
        isinstance(label_fonctionnalite, float) and pd.isna(label_fonctionnalite)
    ):
        label_fonctionnalite = fs_id

    # Date (B10 → ligne 10, index 9) = avant 18
    date_test = None
    try:
        val_date = df.iloc[9, 1]
        if not pd.isna(val_date):
            date_test = pd.to_datetime(val_date, errors="coerce")
    except Exception:
        pass

    # Site (B9 → ligne 9, index 8, colonne 1)
    site = None
    try:
        val_site = df.iloc[8, 1]  # B9
        if not pd.isna(val_site):
            site = str(val_site).strip()
            if site == "LTA":
                site = "STMA"
    except Exception:
        pass


    # Classe documentaire (D15 → ligne 15, index 14) = avant 18
    classe_documentaire = None
    try:
        val_cls = df.iloc[14, 3]
        if pd.isna(val_cls):
            classe_documentaire = None
        elif isinstance(val_cls, str):
            v = val_cls.strip()
            if not v:
                classe_documentaire = None
            elif v.lower() == "non":
                classe_documentaire = "Non Searchable"
            elif v.lower() == "oui":
                classe_documentaire = "Searchable"
            else:
                classe_documentaire = v
        else:
            classe_documentaire = val_cls
    except Exception:
        pass

    # Nombre de pages total (ligne 17 → index 16)
    try:
        nb_pages_total = int(df.iloc[16, 1])
    except Exception:
        nb_pages_total = None

    # Commentaire additionnel (ligne 46 → index 45, après 18 → offset)
    commentaire_add = None
    try:
        comment_row = df.iloc[r(45)]
        for val in comment_row:
            if isinstance(val, str) and val.strip():
                commentaire_add = val.strip()
                break
    except Exception:
        pass

    # Verdict doc : lignes 6 à 12 (index 5..11)
    verdict_doc = None
    try:
        for rr in range(5, 12):
            val = df.iloc[rr, 2]
            if isinstance(val, str) and val.strip():
                verdict_doc = val.strip()
                break
    except Exception:
        pass

    verdict_score = verdict_to_score(verdict_doc)

    # Nom testeur : lignes 6 et 7 (index 5 & 6)
    nom_testeur = None
    try:
        val7 = df.iloc[6, 1]
        val6 = df.iloc[5, 1]
        out = []
        if isinstance(val7, str) and val7.strip():
            out.append(val7.strip())
        if isinstance(val6, str) and val6.strip():
            out.append(val6.strip())
        if out:
            nom_testeur = " ".join(out)
    except Exception:
        pass

    # Fonctionnalité : 25A / 26A avec offset
    fonctionnalite = None
    try:
        cell_26A = df.iloc[r(25), 0]  # ligne 26 (index 25)
        cell_25A = df.iloc[r(24), 0]  # ligne 25 (index 24)
        sentinel = "[NR] Nombre de test de répétabilité requis"

        if isinstance(cell_26A, str) and cell_26A.strip() and cell_26A.strip() != sentinel:
            fonctionnalite = cell_26A.strip()
        elif isinstance(cell_25A, str) and cell_25A.strip():
            fonctionnalite = cell_25A.strip()
    except Exception:
        pass

    # Noms des tests (ligne 31 → index 30, après 18 → offset)
    try:
        tests = list(df.iloc[r(30), 2:6].dropna())
    except Exception:
        tests = []

    records = []
    # Pour chaque test trouvé...
    for k in range(len(tests)):
        col = 2 + k # Décalage de colonnes (C=2, D=3...)

        # lignes 32..37 → index 31..36 → après 18 → offset
        # On récupère les lignes de résultats (Temps, FN, FP...)
        row_h = r(31)
        row_m = r(32)
        row_j = r(33)
        row_fn = r(34)
        row_fp = r(35)
        row_inc = r(36)

        raw_vals = [
            df.iloc[row_h, col],   # humain
            df.iloc[row_m, col],   # machine
            df.iloc[row_j, col],   # justes
            df.iloc[row_fn, col],  # fn
            df.iloc[row_fp, col],  # fp
            df.iloc[row_inc, col], # incertaines
        ]

        # Test totalement vide → ignoré
        if all(pd.isna(v) for v in raw_vals):
            continue

        def to_float_or_none(v):
            try:
                return float(v) if not pd.isna(v) else None
            except Exception:
                return None

        humain = to_float_or_none(raw_vals[0])
        machine = to_float_or_none(raw_vals[1])
        justes = to_float_or_none(raw_vals[2])
        fn = to_float_or_none(raw_vals[3])
        fp = to_float_or_none(raw_vals[4])
        inc = to_float_or_none(raw_vals[5])

        # Création de l'objet résultat
        rec = {
            "fs_id": fs_id,
            "ref_coedm": ref_coedm,
            "doc_id": doc_id,
            "fiche_type": fiche_type,
            "fiche_name": filename,
            "date_test": date_test,
            "type_test": type_test,
            "type_document": type_doc,
            "classe_documentaire": classe_documentaire,
            "site": site,
            "label_fonctionnalite": label_fonctionnalite,
            "nb_pages_total": nb_pages_total,
            "commentaire_additionnel": commentaire_add,
            "verdict_doc": verdict_doc,
            "verdict_score": verdict_score,
            "nom_testeur": nom_testeur,
            "fonctionnalite": fonctionnalite,
            "test_label": f"Test {k+1}",
            "temps_humain_s": humain,
            "temps_machine_s": machine,
            "nb_pages_justes": justes,
            "fn": fn,
            "fp": fp,
            "incertaines": inc,
        }
        # Calcul des pourcentages
        if nb_pages_total and nb_pages_total != 0:
            rec["taux_justes"] = (justes / nb_pages_total * 100) if justes is not None else None
            rec["taux_fn"] = (fn / nb_pages_total * 100) if fn is not None else None
            rec["taux_fp"] = (fp / nb_pages_total * 100) if fp is not None else None
            rec["taux_incertaines"] = (inc / nb_pages_total * 100) if inc is not None else None
        else:
            rec["taux_justes"] = None
            rec["taux_fn"] = None
            rec["taux_fp"] = None
            rec["taux_incertaines"] = None


        if humain is not None and machine is not None:
            rec["gain_temps_s"] = humain - machine
        else:
            rec["gain_temps_s"] = None

        records.append(rec)

    return pd.DataFrame(records)


# =====================================================================
# INTERFACE STREAMLIT
# =====================================================================

st.title("Analyse automatique des fiches IA4Doc")

# Zone de glisser-déposer des fichiers
uploaded_files = st.file_uploader(
    "Fichiers Excel ou ZIP",
    type=["xlsx", "zip"],
    accept_multiple_files=True
)

def iter_excel_files(uploaded_files):
    """
    Génère des objets fichier Excel (avec un attribut .name)
    à partir de ce que l'utilisateur a uploadé :
    - .xlsx directs
    - .zip contenant des .xlsx
    """
    for up in uploaded_files:
        fname = up.name.lower()

        # 1) Cas ZIP
        if fname.endswith(".zip"):
            try:
                with zipfile.ZipFile(up) as zf:
                    for member in zf.namelist():
                        if not member.lower().endswith(".xlsx"):
                            continue
                        data = zf.read(member)
                        bio = io.BytesIO(data)
                        # On donne un "name" pour que parse_fiche puisse l'utiliser
                        bio.name = os.path.basename(member)
                        yield bio
            except Exception as e:
                st.error(f"Erreur en lisant le zip {up.name} : {e}")

        # 2) Cas Excel direct
        elif fname.endswith(".xlsx"):
            yield up



if uploaded_files:

    # --- Chargement et Analyse des données ---
    # On stocke les erreurs dans la session pour pouvoir les supprimer (bouton ❌)
    if "file_errors" not in st.session_state:
        st.session_state["file_errors"] = {}

    def add_file_error(fname: str, msg: str):
        # Garde la dernière erreur pour ce fichier (simple)
        st.session_state["file_errors"][fname] = msg

    # ------------------------------------------------------
    # PARSE DES FICHES + anti-doublon
    # ------------------------------------------------------
    dfs = []
    seen_files = set()

    for file in iter_excel_files(uploaded_files):
        # Anti-doublon de fichier
        if file.name in seen_files:
            add_file_error(f"Fiche déjà chargée : {file.name} (ignorée)")
            continue
        seen_files.add(file.name)

        # Le fichier référentiel (pourScript-tableauxJeremie.xlsx) n'est pas une fiche de test
        if file.name.lower().startswith("pourscript-tableauxjeremie"):
            continue

        try:
            # APPEL AU LECTEUR (parse_fiche)
            df = parse_fiche(file)
            if not df.empty:
                dfs.append(df)
        except Exception as e:
            add_file_error(file.name, str(e))

    # Affiche les erreurs et permet de les enlever (juste l'affichage, pas besoin de re-uploader)
    if st.session_state["file_errors"]:
        st.subheader("Erreurs détectées")
        for fname, msg in list(st.session_state["file_errors"].items()):
            c1, c2 = st.columns([12, 1])
            c1.error(f"{fname} : {msg}")
            if c2.button("❌", key=f"rm_err_{fname}"):
                del st.session_state["file_errors"][fname]

    if not dfs:
        st.warning("Aucune fiche valide.")
        st.stop()

    data = pd.concat(dfs, ignore_index=True)

    # ID simple pour pouvoir exclure des lignes plus tard (1 ligne = 1 test)
    data = data.reset_index(drop=True)
    data["test_uid"] = data.index.astype(int)

    # ------------------------------------------------------
    # NR / TUF : règle métier
    # ------------------------------------------------------
    # - Pour la PERF : si NR existe -> NR gagne ; si plusieurs NR -> on garde la plus récente
    # - Si pas de NR : on garde la plus récente des TUF
    # - IMPORTANT : on ne jette rien côté volume. Tous les tests restent dans "data".
    #
    # Donc ici on fabrique juste un sous-jeu "data_calc" utilisé pour les moyennes de score.
    overrides_report = pd.DataFrame()
    data["date_test"] = pd.to_datetime(data.get("date_test"), errors="coerce")

    def _pick_latest_fiche(g: pd.DataFrame, fiche_type: str):
        # Renvoie le nom de fiche la plus récente pour un type donné (NR ou TUF)
        sub = g[g["fiche_type"].astype(str).str.upper() == fiche_type].copy()
        if sub.empty:
            return None
        # Une fiche = plusieurs lignes (Test 1, Test 2...), donc on déduplique au niveau fichier
        sub = sub.sort_values("date_test", ascending=False)
        return sub["fiche_name"].iloc[0]

    def _pick_best_fiche(g: pd.DataFrame) -> str:
        # D'abord NR (le plus récent), sinon TUF (le plus récent), sinon le plus récent tout court
        fnr = _pick_latest_fiche(g, "NR")
        if fnr:
            return fnr
        ftuf = _pick_latest_fiche(g, "TUF")
        if ftuf:
            return ftuf
        g = g.sort_values("date_test", ascending=False)
        return g["fiche_name"].iloc[0]

    rows = []
    best_fiches = []

    if all(c in data.columns for c in ["doc_id", "fs_id", "fiche_type", "fiche_name"]):
        for (doc, fs), g in data.groupby(["doc_id", "fs_id"], dropna=False):
            best = _pick_best_fiche(g)
            best_fiches.append(best)

            # Petit rapport NR vs TUF (utile en audit)
            fnr = _pick_latest_fiche(g, "NR")
            ftuf = _pick_latest_fiche(g, "TUF")
            if fnr and ftuf:
                # score = verdict_score (déjà entre 0 et 1)
                nr_score = pd.to_numeric(g[g["fiche_name"] == fnr]["verdict_score"], errors="coerce").mean()
                tuf_score = pd.to_numeric(g[g["fiche_name"] == ftuf]["verdict_score"], errors="coerce").mean()

                nr_verdict = g[g["fiche_name"] == fnr]["verdict_doc"].dropna().astype(str)
                tuf_verdict = g[g["fiche_name"] == ftuf]["verdict_doc"].dropna().astype(str)

                rows.append({
                    "doc_id": doc,
                    "fs_id": fs,
                    "fiche_NR": fnr,
                    "fiche_TUF": ftuf,
                    "verdict_NR": nr_verdict.mode().iloc[0] if len(nr_verdict) else "",
                    "verdict_TUF": tuf_verdict.mode().iloc[0] if len(tuf_verdict) else "",
                    "perf_NR": (nr_score * 100) if pd.notna(nr_score) else np.nan,
                    "perf_TUF": (tuf_score * 100) if pd.notna(tuf_score) else np.nan,
                    "delta_NR_minus_TUF": ((nr_score - tuf_score) * 100) if pd.notna(nr_score) and pd.notna(tuf_score) else np.nan,
                })

        if rows:
            overrides_report = pd.DataFrame(rows)

        # data_calc = uniquement les fiches "retenues" pour les calculs de perf
        data_calc = data[data["fiche_name"].isin(best_fiches)].copy()
    else:
        data_calc = data.copy()

    data["fs_id"] = data["fs_id"].astype(str).str.strip()


    # ------------------------------------------------------
    # Gestion des exclusions (tests à ignorer)
    # ------------------------------------------------------

    # --- Gestion des exclusions manuelles (Cocher pour ignorer) ---
    if "excluded_test_uids" not in st.session_state:
        st.session_state["excluded_test_uids"] = set()

    excluded = st.session_state["excluded_test_uids"]
    data_f = data[~data["test_uid"].isin(excluded)].copy()

    # Même logique de filtre mais pour les calculs de perf (NR plus récent, etc.)
    # On applique aussi les exclusions manuelles.
    data_calc_f = data_calc[~data_calc["test_uid"].isin(excluded)].copy()

    # Colonnes "données extraites" (utilisées aussi pour les exports)
    clean_cols_internal = [
        "test_uid",
        "fs_id",
        "doc_id",
        "fiche_type",
        "fiche_name",
        "ref_coedm",
        "date_test",
        "nom_testeur",
        "type_test",
        "type_document",
        "classe_documentaire",
        "site",
        "fonctionnalite",
        "nb_pages_total",
        "test_label",
        "temps_humain_s",
        "temps_machine_s",
        "nb_pages_justes",
        "fn",
        "fp",
        "incertaines",
        "taux_justes",
        "taux_fn",
        "taux_fp",
        "taux_incertaines",
        "gain_temps_s",
        "verdict_doc",
    ]

    clean_cols = [c for c in clean_cols_internal if c not in ("test_uid", "doc_id")]

    # ==================================================================
    # CHARGEMENT RÉFÉRENTIEL (Feuil1 / Feuil2)
    # ==================================================================

    # C'est là qu'on utilise "infer_fo_cols" pour trouver les colonnes fo1...
    ref_xls = pd.ExcelFile("pourScript-tableauxJeremie.xlsx") # tableau excel à fournir pour le référentiel des fonctionnalités outil
    ref1 = ref_xls.parse("Feuil1")
    ref2 = ref_xls.parse("Feuil2")

    # Nettoyage colonnes vides
    ref1 = ref1.loc[:, ~ref1.columns.str.contains("Unnamed")]
    ref2 = ref2.loc[:, ~ref2.columns.str.contains("Unnamed")]

    # Unifier la colonne complexité (compatibilité anciennes versions)
    comp1 = ref2["complexité.1"] if "complexité.1" in ref2.columns else None
    comp0 = ref2["complexité"] if "complexité" in ref2.columns else None

    if comp1 is not None and comp0 is not None:
        ref2["complexité_unifiee"] = comp1.where(comp1.notna(), comp0)
    elif comp1 is not None:
        ref2["complexité_unifiee"] = comp1
    else:
        ref2["complexité_unifiee"] = comp0

    ref2["complexité_unifiee"] = ref2["complexité_unifiee"].fillna("tbd")
    ref2 = ref2.drop(columns=[c for c in ["complexité", "complexité.1"] if c in ref2.columns])
    ref2 = ref2.rename(columns={"complexité_unifiee": "complexité"})

    # Merge référentiel
    ref_full = ref1.merge(ref2, on="N° fs", how="left")
    ref_full["N° fs"] = ref_full["N° fs"].astype(str).str.strip()

    # Colonnes FO détectées automatiquement (fo1..foN)
    fo_cols = infer_fo_cols(ref_full)

    base_cols = ["N° fs"] + fo_cols + ["complexité"]
    ref_base = ref_full[base_cols].copy()

    # Assure qu'il y a au moins une colonne cochée 'x' 
    def ensure_at_least_one_fo(row: pd.Series) -> pd.Series:
        if not fo_cols:
            return row
        if not any(str(row.get(fo, "")).strip().lower() == "x" for fo in fo_cols):
            row[fo_cols[0]] = "x" #<- sinon on ajoute par défaut une croix sur la fo1
        return row

    ref_base = ref_base.apply(ensure_at_least_one_fo, axis=1)

    # ==================================================================
    # Critères / nb tests : depuis Excel si disponible, sinon fallback
    # ==================================================================


    #A revoir
    # --- Récupération des Critères de Réussite (99% etc) ---
    # Fonctions helpers pour lire les %, int ou 'tbd'
    def _to_percent_str(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "tbd"
        if isinstance(v, str):
            s = v.strip()
            if not s:
                return "tbd"
            # déjà "99%" ?
            if "%" in s:
                return s
            # "0.99" ?
            try:
                f = float(s.replace(",", "."))
                if 0 <= f <= 1.2:
                    return f"{f*100:.0f}%"
                if 1.2 < f <= 100:
                    return f"{f:.0f}%"
            except Exception:
                pass
            return s
        if isinstance(v, (int, float)):
            if 0 <= v <= 1.2:
                return f"{v*100:.0f}%"
            if 1.2 < v <= 100:
                return f"{v:.0f}%"
        return "tbd"

    def _to_int_or_tbd(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "tbd"
        if isinstance(v, str):
            s = v.strip()
            if not s or s.lower() == "tbd":
                return "tbd"
            if s.upper() == "NA":
                return "NA"
            try:
                return int(float(s.replace(",", ".")))
            except Exception:
                return s
        if isinstance(v, (int, float)):
            try:
                return int(v)
            except Exception:
                return "tbd"
        return "tbd"
    
    # Dictionnaire par défaut
    crit = {
        "pc":  {"nb": 764, "critere": "99%", "echelle": "moins de 1% de FN soit 99% de TN"},
        "lc":  {"nb": 764, "critere": "99%", "echelle": "moins de 1% de FN soit 99% de TN"},
        "c":   {"nb": 252, "critere": "97%", "echelle": "moins de 3% de FN soit 97% de TN"},
        "tbd": {"nb": "tbd", "critere": "tbd", "echelle": ""},
        "NA":  {"nb": "NA", "critere": "NA", "echelle": ""},
    }

    # Si l'Excel contient les colonnes, on enrichit/écrase le mapping
    nb_col = None
    for c in ["nb test par fonction outil", "nb_tests", "nb test", "nb"]:
        if c in ref2.columns:
            nb_col = c
            break

    crit_col = None
    for c in ["critère 1", "critere 1", "critère", "critere", "critere %", "critère %"]:
        if c in ref2.columns:
            crit_col = c
            break

    echelle_col = None
    for c in ["échelle", "echelle"]:
        if c in ref2.columns:
            echelle_col = c
            break

    if nb_col or crit_col or echelle_col:
        tmp_crit = ref2[["complexité"] + [c for c in [nb_col, crit_col, echelle_col] if c]].copy()
        tmp_crit["complexité"] = tmp_crit["complexité"].astype(str).str.strip()
        tmp_crit = tmp_crit.dropna(subset=["complexité"])
        tmp_crit = tmp_crit[tmp_crit["complexité"].str.lower().ne("nan")]

        for cx, grp in tmp_crit.groupby("complexité"):
            cx_norm = str(cx).strip()
            info = crit.get(cx_norm, {"nb": "tbd", "critere": "tbd", "echelle": ""}).copy()

            if nb_col and nb_col in grp.columns:
                val_nb = grp[nb_col].dropna().iloc[0] if grp[nb_col].dropna().shape[0] else None
                info["nb"] = _to_int_or_tbd(val_nb)

            if crit_col and crit_col in grp.columns:
                val_cr = grp[crit_col].dropna().iloc[0] if grp[crit_col].dropna().shape[0] else None
                info["critere"] = _to_percent_str(val_cr)

            if echelle_col and echelle_col in grp.columns:
                val_ec = grp[echelle_col].dropna().iloc[0] if grp[echelle_col].dropna().shape[0] else None
                info["echelle"] = "" if val_ec is None or (isinstance(val_ec, float) and pd.isna(val_ec)) else str(val_ec)

            crit[cx_norm] = info

    # =====================================================================
    # 1) Tableau — Performance & quantité de tests par FS
    # =====================================================================
    st.subheader("1 — Perf & nb tests par FS")
    show_new = st.toggle("Afficher / masquer (1)", value=True, key="show_section_1")
    if show_new:
        # Quantité de tests : au niveau "test" (une ligne = un test)
        qte_tab = (
            data_f.groupby("fs_id")
            .agg(quantite_tests=("test_label", "count"))
            .reset_index()
        )

        # Performance : au niveau "fiche" (un fichier = un verdict)
        perf_base = data_calc_f.drop_duplicates(subset=["fs_id", "fiche_name"]).copy()
        perf_tab = (
            perf_base.groupby("fs_id")
            .agg(performance_score=("verdict_score", "mean"))
            .reset_index()
        )
        perf_tab["performance"] = perf_tab["performance_score"] * 100

        new_tab = (
            qte_tab.merge(perf_tab[["fs_id", "performance"]], on="fs_id", how="left")
            .rename(columns={"fs_id": "fsXX"})
        )

        # Tri naturel fs1, fs2, fs10...
        def _fs_key(s):
            m = re.search(r"(\d+)$", str(s))
            return int(m.group(1)) if m else 10**9

        new_tab = new_tab.sort_values("fsXX", key=lambda col: col.map(_fs_key))

        # Affichage avec % (mais on garde aussi une version numérique pour le graphique)
        new_tab["performance_num"] = pd.to_numeric(new_tab["performance"], errors="coerce")
        new_tab_display = new_tab[["fsXX", "performance_num", "quantite_tests"]].copy()
        new_tab_display["performance"] = new_tab_display["performance_num"].map(
            lambda x: "" if pd.isna(x) else f"{x:.1f}%"
        )
        new_tab_display = new_tab_display.rename(columns={"quantite_tests": "quantité de tests"})
        st.dataframe(new_tab_display[["fsXX", "performance", "quantité de tests"]], use_container_width=True)

        # Graphique double axe (comme la capture) : quantité (gauche) + performance % (droite)

        # Graphique double axe (comme la capture) : 2 barres par FS
        # - Quantité (bleu) sur l'axe gauche
        # - Performance (%) (rouge) sur l'axe droit
        # Graphique interactif (comme la capture) : 2 barres par FS
        # - Quantité (bleu clair) sur l'axe gauche
        # - Performance (%) (bleu foncé) sur l'axe droit
        # + Survol : détail par classe documentaire
        fs_order = new_tab["fsXX"].astype(str).tolist()
        qte = pd.to_numeric(new_tab["quantite_tests"], errors="coerce").fillna(0).astype(int).tolist()
        perf = pd.to_numeric(new_tab["performance_num"], errors="coerce")
        perf_pct = perf.fillna(0).tolist()


        # Détails par classe documentaire
        tmp_detail = data_f.copy()
        tmp_detail["fs_id"] = tmp_detail["fs_id"].astype(str).str.strip()
        tmp_detail["classe_documentaire"] = tmp_detail.get("classe_documentaire", "").fillna("Inconnu").astype(str)

        tmp_perf_detail = data_calc_f.drop_duplicates(subset=["fs_id", "fiche_name"]).copy()
        tmp_perf_detail["classe_documentaire"] = tmp_perf_detail.get("classe_documentaire", "").fillna("Inconnu").astype(str)
        perf_by_class = (
            tmp_perf_detail.groupby(["fs_id", "classe_documentaire"])["verdict_score"]
            .mean()
            .reset_index()
        )
        perf_by_class["perf_pct"] = perf_by_class["verdict_score"] * 100
        qte_by_class = (
            tmp_detail.groupby(["fs_id", "classe_documentaire"])            .size()            .reset_index(name="nb")
        )

        def _hover_perf(fs: str, overall: float) -> str:
            sub = perf_by_class[perf_by_class["fs_id"] == fs].copy()
            sub = sub.sort_values("classe_documentaire")
            lines = [f"<b>{overall:.1f}%</b> (performance totale)"]
            for _, r in sub.iterrows():
                val = r["perf_pct"]
                if pd.notna(val):
                    lines.append(f"Classe doc {r['classe_documentaire']} : {val:.1f}%")
            return "<br>".join(lines)

        def _hover_qte(fs: str, total: int) -> str:
            sub = qte_by_class[qte_by_class["fs_id"] == fs].copy()
            sub = sub.sort_values("classe_documentaire")
            lines = [f"<b>{total} tests</b> (total)"]
            for _, r in sub.iterrows():
                lines.append(f"Classe doc {r['classe_documentaire']} : {int(r['nb'])}")
            return "<br>".join(lines)

        hover_perf = [_hover_perf(fs, float(p)) for fs, p in zip(fs_order, perf_pct)]
        hover_qte = [_hover_qte(fs, int(t)) for fs, t in zip(fs_order, qte)]

        fig_bar = go.Figure()
        fig_bar.add_bar(
            x=fs_order,
            y=qte,
            name="Quantité de tests",
            marker_color="#a7d2ff",
            hovertext=hover_qte,
            hoverinfo="text",
            offsetgroup="qte",
            alignmentgroup="fs",
        )
        fig_bar.add_bar(
            x=fs_order,
            y=perf_pct,
            name="Performance (%)",
            marker_color="#063970",
            yaxis="y2",
            hovertext=hover_perf,
            hoverinfo="text",
            offsetgroup="perf",
            alignmentgroup="fs",
        )

        fig_bar.update_layout(
            barmode="group",
            xaxis=dict(title="FS"),
            yaxis=dict(title="Quantité de tests", rangemode="tozero"),
            yaxis2=dict(title="Performance (%)", overlaying="y", side="right", range=[0, 100]),
            legend=dict(orientation="h", yanchor="bottom", y=1.15, xanchor="center", x=0.5),
            margin=dict(t=80, b=40, l=40, r=40),
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    # =====================================================================
    # 2) KPI globaux
    # =====================================================================
    # =====================================================================
    # 2) KPI globaux
    # =====================================================================
    st.subheader("2 — KPI globaux")

    # KPIs (sur les données filtrées)
    nb_tests_total = int(len(data_f))
    nb_docs = int(data_f["ref_coedm"].nunique()) if "ref_coedm" in data_f.columns else 0
    nb_fs = int(data_f["fs_id"].nunique()) if "fs_id" in data_f.columns else 0
    nb_testeurs = int(data_f["nom_testeur"].nunique()) if "nom_testeur" in data_f.columns else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Tests", nb_tests_total)
    c2.metric("Documents testés", nb_docs)
    c3.metric("FS testées", nb_fs)
    c4.metric("Testeurs", nb_testeurs)

    # -----------------------------------------------------------------
    # Remplacements NR vs TUF (NR écrase TUF) + warning si NR < TUF
    # -----------------------------------------------------------------
    if isinstance(overrides_report, pd.DataFrame) and len(overrides_report):
        st.subheader("Remplacements NR vs TUF")
        st.info(
            f"{int(len(overrides_report))} cas : une fiche NR a remplacé une fiche TUF (même doc_id + même FS)."
        )

        disp_all = overrides_report.copy()

        # Numériser pour détecter NR < TUF
        for c in ["perf_TUF", "perf_NR", "delta_NR_minus_TUF"]:
            if c in disp_all.columns:
                disp_all[c] = pd.to_numeric(disp_all[c], errors="coerce")

        worse = disp_all.copy()
        if "delta_NR_minus_TUF" in worse.columns:
            worse = worse[worse["delta_NR_minus_TUF"] < 0]

        if len(worse):
            st.warning("⚠️ Attention : NR < TUF sur la performance moyenne pour certains remplacements.")

        disp_show = disp_all.copy()
        for c in ["perf_TUF", "perf_NR", "delta_NR_minus_TUF"]:
            if c in disp_show.columns:
                disp_show[c] = disp_show[c].map(lambda x: "" if pd.isna(x) else f"{x:.1f}%")

        st.dataframe(
            disp_show[
                ["doc_id", "fs_id", "fiche_TUF", "fiche_NR", "verdict_TUF", "verdict_NR", "perf_TUF", "perf_NR", "delta_NR_minus_TUF"]
            ],
            use_container_width=True,
        )

# ---------------------------------------------------------------------
    # Graphique unique : volume de tests + performance dans le temps
    # (axe X datetime => dates espacées proportionnellement)
    # ---------------------------------------------------------------------
    # ---------------------------------------------------------------------
    # Graphique : évolution CUMULÉE du nombre total de tests (courbe qui grimpe)
    # (axe X datetime => dates espacées proportionnellement)
    # ---------------------------------------------------------------------
    st.subheader("4 — Évolution cumulée tests")

    if "date_test" in data_f.columns:
        import matplotlib.dates as mdates

        tmp = data_f.copy()
        tmp["date_test"] = pd.to_datetime(tmp["date_test"], errors="coerce")
        tmp = tmp.dropna(subset=["date_test"])

        if len(tmp):
            # Agrégation par jour
            daily = (
                tmp.groupby(tmp["date_test"].dt.floor("D"))
                .agg(nb_tests=("date_test", "size"))
                .reset_index()
                .rename(columns={"date_test": "date"})
                .sort_values("date")
            )

            daily_cum = daily[["date", "nb_tests"]].copy()
            daily_cum["nb_tests_cum"] = daily_cum["nb_tests"].cumsum()

            fig2, ax2 = plt.subplots(figsize=(11, 4))
            ax2.plot(
                daily_cum["date"],
                daily_cum["nb_tests_cum"],
                marker="o",
                color="tab:blue",
                label="Total cumulé de tests",
            )
            ax2.set_ylabel("Nombre total de tests")
            ax2.grid(True, axis="y", alpha=0.3)

            locator2 = mdates.AutoDateLocator()
            ax2.xaxis.set_major_locator(locator2)
            ax2.xaxis.set_major_formatter(mdates.ConciseDateFormatter(locator2))
            fig2.autofmt_xdate()

            ax2.legend(loc="upper center", bbox_to_anchor=(0.5, 1.18), ncol=1, frameon=False)
            st.pyplot(fig2, use_container_width=True)
        else:
            st.info("Aucune date exploitable dans la colonne 'date_test'.")
    else:
        st.info("Aucune colonne 'date_test' détectée : impossible d'afficher l'évolution des tests dans le temps.")

    st.subheader("5 — Justesse moyenne FS")
    show_tab2bis = st.toggle("Afficher / masquer", value=True, key="show_2bis")
    if show_tab2bis:
        tab2bis, fo_cols_2bis = build_tab2bis(data_calc_f, ref_base)
        tab2bis_display = tab2bis[["N° fs"] + fo_cols_2bis]
        styles_2bis = style_tab2bis(tab2bis_display, fo_cols_2bis)
        styler_2bis = tab2bis_display.style.apply(lambda _: styles_2bis, axis=None)
        st.dataframe(styler_2bis, use_container_width=True)

    # =====================================================================
    # 4) Tableau 3 (progression)
    # =====================================================================
    st.subheader("6 — Progression tests FS")
    show_tab3 = st.toggle("Afficher / masquer", value=True, key="show_3")
    if show_tab3:
        fs_testes = data_f["fs_id"].unique()
        tab3 = ref_base[ref_base["N° fs"].isin(fs_testes)].copy()

        tests_counts = data_f.groupby("fs_id").size().to_dict()

        tmp = data_f.copy()
        tmp["vcat"] = tmp["verdict_doc"].apply(norm_verdict)

        tver = tmp.groupby(["fs_id", "vcat"]).size().unstack(fill_value=0)
        tests_vdict = {fs: row.to_dict() for fs, row in tver.iterrows()}

        def convert_cell_progress(row, fo):
            val = row[fo]
            cx = row["complexité"]
            fs = row["N° fs"]
            if val != "x":
                return ""
            done = tests_counts.get(fs, 0)
            info = crit.get(cx, {})
            total = info.get("nb", "tbd")

            if isinstance(total, str) and total in ("tbd", "NA"):
                return f"{done}/{total}"

            if isinstance(total, (int, float)) and total > 0:
                pct = done / total
                return f"{done}/{int(total)} ({pct:.1%})"

            return str(done)

        for fo in fo_cols:
            tab3[fo] = tab3.apply(lambda row: convert_cell_progress(row, fo), axis=1)

        def ratio_cat(fs, cat):
            fs = str(fs)
            total = tests_counts.get(fs, 0)
            if total == 0:
                return ""
            n = tests_vdict.get(fs, {}).get(cat, 0)
            return f"{int(n)}/{int(total)}"

        tab3["Bon"] = tab3["N° fs"].map(lambda fs: ratio_cat(fs, "bon"))
        tab3["Partiellement bon"] = tab3["N° fs"].map(lambda fs: ratio_cat(fs, "partiel"))
        tab3["Mauvais"] = tab3["N° fs"].map(lambda fs: ratio_cat(fs, "mauvais"))
        tab3["None"] = tab3["N° fs"].map(lambda fs: ratio_cat(fs, "none"))

        tab3_display = tab3.drop(columns=["complexité"])
        st.dataframe(tab3_display, use_container_width=True)

    # =====================================================================
    # 5) Tableau 1 — Référentiel brut
    # =====================================================================
    st.subheader("7 — Référentiel brut")
    show_tab1 = st.toggle("Afficher / masquer", value=False, key="show_1")
    if show_tab1:
        st.dataframe(ref_base)

    # =====================================================================
    # 6) Tableau 2 — Critères par fonctionnalité (SANS code couleur)
    # =====================================================================
    st.subheader("8 — Critères de réussite par fonctionnalité")
    show_tab2 = st.toggle("Afficher / masquer", value=False, key="show_2")
    if show_tab2:
        tab2 = ref_base.copy()

        def convert_cell_percent(row, fo):
            val = row[fo]
            cx = row["complexité"]
            if val != "x":
                return ""
            if pd.isna(cx) or cx in ("tbd", "NA"):
                return "tbd"
            info = crit.get(cx)
            if not info:
                return "tbd"
            return info.get("critere", "tbd")

        for fo in fo_cols:
            tab2[fo] = tab2.apply(lambda row: convert_cell_percent(row, fo), axis=1)

        tab2 = tab2.drop(columns=["complexité"])
        st.dataframe(tab2)

    # (Section KPI globaux déplacée en 2)



    # =====================================================================
    # 8) Réussite par classe doc et fonctionnalités
    # =====================================================================
    st.subheader("9 — Réussite par classe documentaire et fonctionnalité")
    show_cd = st.toggle("Afficher / masquer réussite par classe doc", value=False)
    if show_cd:
        if "classe_documentaire" in data_f.columns:
            tmp_cd = data_f.copy()
            tmp_cd["vcat"] = tmp_cd["verdict_doc"].apply(norm_verdict)

            tot_cd = (
                tmp_cd.groupby(["fs_id", "classe_documentaire"])
                .size()
                .rename("total_tests")
            )

            bon_cd = (
                tmp_cd[tmp_cd["vcat"] == "bon"]
                .groupby(["fs_id", "classe_documentaire"])
                .size()
                .rename("bon_tests")
            )

            cd_df = pd.concat([tot_cd, bon_cd], axis=1).fillna(0)
            cd_df["bon_tests"] = cd_df["bon_tests"].astype(int)
            cd_df["total_tests"] = cd_df["total_tests"].astype(int)

            def fmt_ratio(row):
                if row["total_tests"] == 0:
                    return ""
                return f"{row['bon_tests']}/{row['total_tests']}"

            cd_df["ratio"] = cd_df.apply(fmt_ratio, axis=1)

            pivot_cd = cd_df.reset_index().pivot(
                index="fs_id",
                columns="classe_documentaire",
                values="ratio"
            )

            st.dataframe(pivot_cd)
        else:
            st.info("Aucune classe documentaire trouvée dans les fiches.")

    # =====================================================================
    # 9) Taux de justesse par classe doc
    # =====================================================================
    st.subheader("10 — Justesse par classe doc")
    show_taux_cd = st.toggle("Afficher / masquer taux par classe doc", value=False)
    if show_taux_cd:
        if "classe_documentaire" in data_f.columns:
            classe_summary = (
                data_f.groupby("classe_documentaire")
                .agg(
                    nb_tests=("test_label", "count"),
                    taux_justes_moy=("taux_justes", "mean"),
                )
                .reset_index()
            )
            classe_summary["taux_justes_moy"] = classe_summary["taux_justes_moy"].round(2)
            st.dataframe(classe_summary)
        else:
            st.info("Aucune classe documentaire trouvée dans les fiches.")
            classe_summary = pd.DataFrame()

    # =====================================================================
    # 10) KPI par fonctionnalité
    # =====================================================================
    st.subheader("11 — KPI par fonctionnalité")
    show_kpi_fct = st.toggle("Afficher / masquer KPI par fonctionnalité", value=False)
    if show_kpi_fct:
        kpi_fct = (
            data_f.groupby("fs_id")
            .agg(
                nb_tests=("test_label", "count"),
                temps_humain_moy=("temps_humain_s", "mean"),
                temps_machine_moy=("temps_machine_s", "mean"),
                taux_justes_moy=("taux_justes", "mean"),
                nb_docs=("ref_coedm", lambda s: s.nunique()),
            )
            .reset_index()
        )

        tmp_kpi = data_f.copy()
        tmp_kpi["vcat"] = tmp_kpi["verdict_doc"].apply(norm_verdict)

        counts = (
            tmp_kpi.groupby(["fs_id", "vcat"])
            .size()
            .unstack(fill_value=0)
            .reset_index()
        )

        for c in ["bon", "partiel", "mauvais", "none"]:
            if c not in counts.columns:
                counts[c] = 0

        counts["total_tests"] = counts[["bon", "partiel", "mauvais", "none"]].sum(axis=1)

        def ratio_str(n, total):
            if total == 0:
                return ""
            return f"{int(n)}/{int(total)}"

        counts["bon_ratio"] = counts.apply(lambda r: ratio_str(r.get("bon", 0), r["total_tests"]), axis=1)
        counts["partiel_ratio"] = counts.apply(lambda r: ratio_str(r.get("partiel", 0), r["total_tests"]), axis=1)
        counts["mauvais_ratio"] = counts.apply(lambda r: ratio_str(r.get("mauvais", 0), r["total_tests"]), axis=1)

        kpi_fct = kpi_fct.merge(
            counts[["fs_id", "bon_ratio", "partiel_ratio", "mauvais_ratio"]],
            on="fs_id",
            how="left",
        )

        kpi_fct = kpi_fct.rename(columns={
            "bon_ratio": "bon",
            "partiel_ratio": "partiellement bon",
            "mauvais_ratio": "mauvais",
        })

        st.dataframe(kpi_fct)

    # =====================================================================
    # 11) Commentaires additionnels
    # =====================================================================
    st.subheader("12 — Commentaires additionnels détectés")
    show_comments = st.toggle("Afficher / masquer commentaires", value=False)
    if show_comments:
        comment_rows = (
            data_f[["ref_coedm", "verdict_doc", "commentaire_additionnel"]]
            .dropna(subset=["commentaire_additionnel"])
            .drop_duplicates()
        )

        if comment_rows.empty:
            st.info("Aucun commentaire dans les fiches.")
        else:
            options = []
            for idx, row in comment_rows.iterrows():
                ref = row["ref_coedm"]
                verdict = row["verdict_doc"] or "Non renseigné"
                label = f"{ref} — Verdict : {verdict}"
                options.append((label, idx))

            labels = [o[0] for o in options]
            choice_label = st.selectbox(
                "Sélectionner un document pour voir le commentaire",
                labels,
            )

            chosen_idx = dict(options)[choice_label]
            chosen_row = comment_rows.loc[chosen_idx]

            st.info(
                f"📄 {chosen_row['ref_coedm']} — Verdict : {chosen_row['verdict_doc'] or 'Non renseigné'}\n\n"
                f"📝 Commentaire : {chosen_row['commentaire_additionnel']}"
            )

    # =====================================================================
    # 12) Données extraites + exclusions (❌)
    # =====================================================================
    st.subheader("13 — Données extraites")
    show_data = st.toggle("Afficher / masquer données extraites", value=False)
    if show_data:

        # Affichage complet (non filtré) — SANS doc_id / test_uid
        df_clean_all = data[clean_cols].copy()
        if "date_test" in df_clean_all.columns:
            df_clean_all["date_test"] = pd.to_datetime(df_clean_all["date_test"], errors="coerce").dt.date

        num_cols = df_clean_all.select_dtypes(include=["float", "int"]).columns
        df_clean_all[num_cols] = df_clean_all[num_cols].round(2)

        # Editor pour exclure des tests (ID interne via l'index, non affiché)
        df_editor = data[["test_uid", "fs_id", "ref_coedm", "test_label", "verdict_doc"]].copy()
        df_editor = df_editor.set_index("test_uid")
        df_editor["❌ Exclure"] = df_editor.index.isin(st.session_state["excluded_test_uids"])

        st.caption("Cochez ❌ Exclure pour ignorer des lignes (cela mettra à jour tous les tableaux).")
        edited = st.data_editor(
            df_editor[["fs_id", "ref_coedm", "test_label", "verdict_doc", "❌ Exclure"]],
            hide_index=True,
            column_config={
                "fs_id": st.column_config.TextColumn("FS", disabled=True),
                "ref_coedm": st.column_config.TextColumn("Document", disabled=True),
                "test_label": st.column_config.TextColumn("Test", disabled=True),
                "verdict_doc": st.column_config.TextColumn("Verdict", disabled=True),
                "❌ Exclure": st.column_config.CheckboxColumn("❌ Exclure"),
            },
            disabled=["fs_id", "ref_coedm", "test_label", "verdict_doc"],
            use_container_width=True,
        )

        # Mettre à jour la session_state
        new_excluded = set(edited.index[edited["❌ Exclure"]].astype(int).tolist())
        st.session_state["excluded_test_uids"] = new_excluded

        # Affichage data filtrée, avec couleur verdict (SANS doc_id / test_uid)
        df_clean_f = data_f[clean_cols].copy()
        if "date_test" in df_clean_f.columns:
            df_clean_f["date_test"] = pd.to_datetime(df_clean_f["date_test"], errors="coerce").dt.date
        num_cols_f = df_clean_f.select_dtypes(include=["float", "int"]).columns
        df_clean_f[num_cols_f] = df_clean_f[num_cols_f].round(2)

        st.dataframe(df_clean_f.style.applymap(color_verdict, subset=["verdict_doc"]))
    st.subheader("14 — Résultats par type doc")
    show_type = st.toggle("Afficher / masquer résultats par type de doc", value=False)
    if show_type:
        summary = (
            data_f.groupby("type_document")
            .agg(
                nb_tests=("test_label", "count"),
                temps_humain_moy=("temps_humain_s", "mean"),
                temps_machine_moy=("temps_machine_s", "mean"),
                taux_justes_moy=("taux_justes", "mean"),
                taux_fn_moy=("taux_fn", "mean"),
                taux_fp_moy=("taux_fp", "mean"),
            )
            .reset_index()
        )
        st.dataframe(summary)

    # =====================================================================
    # 14) Exports
    # =====================================================================
    st.subheader("15 — Export")
    # 2) Export suivi recettage (Excel) — filtré
    recettage_df = data_f[
        ["fs_id", "verdict_doc", "classe_documentaire", "site", "ref_coedm"]
    ].copy()

    recettage_df["score"] = recettage_df["verdict_doc"].apply(verdict_to_score)
    recettage_df["Modification"] = recettage_df["ref_coedm"].apply(extract_modification)

    recettage_df = recettage_df[
        ["fs_id", "verdict_doc", "score", "classe_documentaire", "site", "Modification"]
    ]

    from io import BytesIO
    output_recettage = BytesIO()
    with pd.ExcelWriter(output_recettage, engine="xlsxwriter") as writer:
        recettage_df.to_excel(writer, sheet_name="Suivi_recettage", index=False)

    st.download_button(
        "Export suivi recettage (Excel)",
        data=output_recettage.getvalue(),
        file_name="suivi_recettage.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # 3) Excel multi-feuilles avec un peu de mise en forme
    # --- Reconstruire Tableau 2bis / Tableau 3 (filtré) ---
    tab2bis_raw, fo_cols_2bis = build_tab2bis(data_calc_f, ref_base)
    tab2bis_xlsx = tab2bis_raw[["N° fs"] + fo_cols_2bis]

    fs_testes = data_f["fs_id"].unique()
    tab3_x = ref_base[ref_base["N° fs"].isin(fs_testes)].copy()
    tests_counts = data_f.groupby("fs_id").size().to_dict()

    def convert_cell_progress_export(row, fo):
        val = row[fo]
        cx = row["complexité"]
        fs = row["N° fs"]
        if val != "x":
            return ""
        done = tests_counts.get(fs, 0)
        info = crit.get(cx, {})
        total = info.get("nb", "tbd")

        if isinstance(total, str) and total in ("tbd", "NA"):
            return f"{done}/{total}"

        if isinstance(total, (int, float)) and total > 0:
            pct = done / total
            return f"{done}/{int(total)} ({pct:.1%})"

        return str(done)

    for fo in fo_cols:
        tab3_x[fo] = tab3_x.apply(lambda row: convert_cell_progress_export(row, fo), axis=1)

    tmp_export = data_f.copy()
    tmp_export["vcat"] = tmp_export["verdict_doc"].apply(norm_verdict)
    tver = tmp_export.groupby(["fs_id", "vcat"]).size().unstack(fill_value=0)
    tests_vdict = {fs: row.to_dict() for fs, row in tver.iterrows()}

    def ratio_cat(fs, cat):
        fs = str(fs)
        total = tests_counts.get(fs, 0)
        if total == 0:
            return ""
        n = tests_vdict.get(fs, {}).get(cat, 0)
        return f"{int(n)}/{int(total)}"

    tab3_x["Bon"] = tab3_x["N° fs"].map(lambda fs: ratio_cat(fs, "bon"))
    tab3_x["Partiellement bon"] = tab3_x["N° fs"].map(lambda fs: ratio_cat(fs, "partiel"))
    tab3_x["Mauvais"] = tab3_x["N° fs"].map(lambda fs: ratio_cat(fs, "mauvais"))
    tab3_x["None"] = tab3_x["N° fs"].map(lambda fs: ratio_cat(fs, "none"))

    tab3_display = tab3_x.drop(columns=["complexité"])

    # Données extraites filtrées
    clean_cols_export = clean_cols
    df_clean_export = data_f[clean_cols_export].copy()
    if "date_test" in df_clean_export.columns:
        df_clean_export["date_test"] = pd.to_datetime(df_clean_export["date_test"], errors="coerce").dt.date
    num_cols = df_clean_export.select_dtypes(include=["float", "int"]).columns
    df_clean_export[num_cols] = df_clean_export[num_cols].round(2)

    # KPI par fonctionnalité (filtré)
    kpi_fct_export = (
        data_f.groupby("fs_id")
        .agg(
            nb_tests=("test_label", "count"),
            temps_humain_moy=("temps_humain_s", "mean"),
            temps_machine_moy=("temps_machine_s", "mean"),
            taux_justes_moy=("taux_justes", "mean"),
            nb_docs=("ref_coedm", lambda s: s.nunique()),
        )
        .reset_index()
    )

    # Résultat par type doc (filtré)
    summary_export = (
        data_f.groupby("type_document")
        .agg(
            nb_tests=("test_label", "count"),
            temps_humain_moy=("temps_humain_s", "mean"),
            temps_machine_moy=("temps_machine_s", "mean"),
            taux_justes_moy=("taux_justes", "mean"),
            taux_fn_moy=("taux_fn", "mean"),
            taux_fp_moy=("taux_fp", "mean"),
        )
        .reset_index()
    )

    # Tableau perf/quantité par FS (filtré) — pour l'export
    perfqte_export = (
        data_f.groupby('fs_id')
        .agg(performance=('taux_justes','mean'), quantite_tests=('test_label','count'))
        .reset_index()
        .rename(columns={'fs_id':'fs'})
    )
    perfqte_export['performance'] = pd.to_numeric(perfqte_export['performance'], errors='coerce')
    perfqte_export['performance_%'] = perfqte_export['performance'].map(lambda x: '' if pd.isna(x) else f"{x:.1f}%")

    output = BytesIO()

    # Titres (doivent être identiques aux sous-titres Streamlit et compatibles Excel)
    TITLE_PERF_FS = "1 — Perf & nb tests par FS"
    TITLE_JUSTESSE_FS = "5 — Justesse moyenne FS"
    TITLE_PROG_FS = "6 — Progression tests FS"
    TITLE_REPL = "Remplacements NR vs TUF"
    TITLE_JUSTESSE_CD = "10 — Justesse par classe doc"
    TITLE_KPI_FCT = "11 — KPI par fonctionnalité"
    TITLE_DATA = "13 — Données extraites"
    TITLE_TYPE_DOC = "14 — Résultats par type doc"

    # 1) Perf / quantité par FS (filtré)
    perf_fs_export = (
        perfqte_export[["fs", "performance_%", "quantite_tests"]]
        .rename(columns={"fs": "fsXX", "performance_%": "performance", "quantite_tests": "quantité de tests"})
        .copy()
    )

    # 10) Justesse par classe doc (filtré)
    justesse_cd_export = pd.DataFrame()
    if "classe_documentaire" in data_f.columns:
        justesse_cd_export = (
            data_f.groupby("classe_documentaire")
            .agg(
                nb_tests=("test_label", "count"),
                taux_justes_moy=("taux_justes", "mean"),
            )
            .reset_index()
        )
        justesse_cd_export["taux_justes_moy"] = justesse_cd_export["taux_justes_moy"].round(2)

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        perf_fs_export.to_excel(writer, sheet_name=excel_sheet_name(TITLE_PERF_FS), index=False)
        tab2bis_xlsx.to_excel(writer, sheet_name=excel_sheet_name(TITLE_JUSTESSE_FS), index=False)
        tab3_display.to_excel(writer, sheet_name=excel_sheet_name(TITLE_PROG_FS), index=False)

        if isinstance(overrides_report, pd.DataFrame) and len(overrides_report):
            overrides_report.to_excel(writer, sheet_name=excel_sheet_name(TITLE_REPL), index=False)

        if isinstance(justesse_cd_export, pd.DataFrame) and len(justesse_cd_export):
            justesse_cd_export.to_excel(writer, sheet_name=excel_sheet_name(TITLE_JUSTESSE_CD), index=False)

        kpi_fct_export.to_excel(writer, sheet_name=excel_sheet_name(TITLE_KPI_FCT), index=False)
        summary_export.to_excel(writer, sheet_name=excel_sheet_name(TITLE_TYPE_DOC), index=False)
        df_clean_export.to_excel(writer, sheet_name=excel_sheet_name(TITLE_DATA), index=False)

        # Formats Excel (couleurs)
        workbook = writer.book

        fmt_red = workbook.add_format({"bg_color": "#fb3f35"})
        fmt_orange = workbook.add_format({"bg_color": "#F4B183"})
        fmt_yellow = workbook.add_format({"bg_color": "#FFF2CC"})
        fmt_light_green = workbook.add_format({"bg_color": "#C6EFCE"})
        fmt_dark_green = workbook.add_format({"bg_color": "#00B050", "font_color": "#FFFFFF"})

        fmt_verdict_bon = workbook.add_format({"bg_color": "#C6EFCE", "font_color": "#006100", "bold": True})
        fmt_verdict_partiel = workbook.add_format({"bg_color": "#FFE699", "font_color": "#7F6000", "bold": True})
        fmt_verdict_mauvais = workbook.add_format({"bg_color": "#fb3f35", "font_color": "#9C0006", "bold": True})

        # Couleurs % dans "Justesse moyenne FS"
        ws2 = writer.sheets[excel_sheet_name(TITLE_JUSTESSE_FS)]
        n_rows_2 = len(tab2bis_xlsx)

        def _simple_pct(val):
            if not isinstance(val, str):
                return None
            s = val.strip().replace("%", "").replace(",", ".")
            if not s:
                return None
            try:
                return float(s)
            except Exception:
                return None

        for r in range(n_rows_2):
            for fo in fo_cols_2bis:
                c_idx = tab2bis_xlsx.columns.get_loc(fo)
                val = tab2bis_xlsx.iloc[r, c_idx]
                pct = _simple_pct(val)
                if pct is None:
                    continue
                if pct < 25:
                    fmt = fmt_red
                elif pct < 50:
                    fmt = fmt_orange
                elif pct < 75:
                    fmt = fmt_yellow
                elif pct < 90:
                    fmt = fmt_light_green
                else:
                    fmt = fmt_dark_green
                ws2.write(r + 1, c_idx, val, fmt)

        # Couleurs verdict dans "Données extraites"
        wsD = writer.sheets[excel_sheet_name(TITLE_DATA)]
        n_rows_D = len(df_clean_export)
        verdict_col_idx = df_clean_export.columns.get_loc("verdict_doc")

        for r in range(n_rows_D):
            val = df_clean_export.iloc[r, verdict_col_idx]
            if not isinstance(val, str):
                continue
            v = val.strip().lower()
            if v == "bon":
                fmt = fmt_verdict_bon
            elif "partiellement" in v:
                fmt = fmt_verdict_partiel
            elif "mauvais" in v:
                fmt = fmt_verdict_mauvais
            else:
                continue
            wsD.write(r + 1, verdict_col_idx, val, fmt)


    st.download_button(
        "Exporter les tableaux (Excel)",
        data=output.getvalue(),
        file_name="IA4Doc_tableaux_complets_filtre.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
