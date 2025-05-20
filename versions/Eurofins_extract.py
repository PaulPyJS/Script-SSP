import pandas as pd
import unicodedata
import re
from collections import defaultdict

import json
import sys
import os


# === Script : EXTRACT VALUE WITH KEYWORD IN AN EXCEL FORMAT - TABLEURS EUROFINS ===
# = v1 : Test import from Excel raw DF-Excel and keyword-based extract
# = v1.5 : Multiple keywords test
# = v2 : Normalizing text and splitting data into keywords based on value or sum
# = v3 : Extraction validated for general keyword with random pick
# = v4 : Adapting the script to allow HAP to be separated from Naphtalene/HAP
# = v5 : DEBUG
# = v6 : Adding SUM calculation based on JSON file to allow local memory
# = v7 : Link to UI and using json for keywords
#
excel_path = sys.argv[1]
keywords_file = sys.argv[2]
sheet_name = sys.argv[3] if len(sys.argv) > 3 else None


export_mode = len(sys.argv) > 3 and sys.argv[3] == "--export"

with open(keywords_file, "r", encoding="utf-8") as f:
    config_json = json.load(f)

    if isinstance(config_json, dict):
        general_keywords = config_json.get("keywords_valides", [])
        groupes_personnalises = config_json.get("groupes_personnalises", {})
    elif isinstance(config_json, list):
        general_keywords = config_json
        groupes_personnalises = {}
    else:
        raise ValueError("Format de configuration JSON non reconnu.")

if sheet_name:
    df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
else:
    xls = pd.ExcelFile(excel_path)
    df_raw = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None)
print("✅ Fichier chargé.")

# Based on Excel format from Eurofins, without blank parts
# Starting with CodeEuro/CodeArt/Date + Data from L5-C4 index 4,3
headers = ['Code Eurofins', 'Code Artelia', 'Date prélèvement'] + df_raw.iloc[4, 3:].tolist()

header_data_only = [str(col).strip() for col in df_raw.iloc[4, 3:].tolist()]
header_index_map = {i: header_data_only[i] for i in range(len(header_data_only))}

# Data values from line 6, index 5
df = df_raw.iloc[5:].copy()
df.columns = headers
df = df.reset_index(drop=True)

df = df.dropna(axis=1, how='all')
df.columns = pd.Index([str(c).strip() for c in df.columns])

header_index_to_df_col_index = {}

for i, col in enumerate(header_data_only):
    if i + 3 >= len(df.columns):
        continue

    df_col_name = df.columns[i + 3]
    for i in range(len(header_data_only)):
        if i + 3 < len(df.columns):
            header_index_to_df_col_index[i] = i + 3






# ==== CLEANING FUNCTIONS ====
#
# Normalizing text in the headers but will also be applied to text user input
def normalize(text):
    if pd.isna(text):
        return ""
    text = str(text).lower()
    text = unicodedata.normalize('NFD', text).encode('ascii', 'ignore').decode('utf-8')
    # Adding spaces before and after to be able to use it to separate word like ethylbenzene
    return " " + text.strip() + " "

def clean_tokens(text):
    if pd.isna(text):
        return []
    text = str(text).lower()
    text = unicodedata.normalize('NFD', text).encode('ascii', 'ignore').decode('utf-8')  # remove accents
    return re.findall(r'[a-z0-9]+', text)  # keep only alphanum words

# ==================== DEBUG FUNCTION =============================
# Used to convert index to Excel columns
def col_idx_to_excel_letter(idx):
    letter = ''
    while idx >= 0:
        letter = chr(idx % 26 + 65) + letter
        idx = idx // 26 - 1
    return letter
# ==================== DEBUG FUNCTION =============================


def charger_groupes_parametres(FICHIER_GROUPES_PARAMETRES = "sum.json"):
    if not os.path.exists(FICHIER_GROUPES_PARAMETRES):
        print(f"⚠️ Fichier '{FICHIER_GROUPES_PARAMETRES}' non trouvé. Création d'un fichier vide.")
        with open(FICHIER_GROUPES_PARAMETRES, "w") as f:
            json.dump({}, f)
        return {}
    with open(FICHIER_GROUPES_PARAMETRES, "r", encoding="utf-8") as f:
        return json.load(f)

def sauvegarder_groupes_parametres(groupes, FICHIER_GROUPES_PARAMETRES = "sum.json"):
    with open(FICHIER_GROUPES_PARAMETRES, "w", encoding="utf-8") as f:
        json.dump(groupes, f, indent=2, ensure_ascii=False)





# ==== PROCESSING FUNCTIONS ====
#
# Looking for the general keywords to match headers with it - returning index of columns
def get_matching_columns(headers, keywords):
    matched = {kw: [] for kw in keywords}
    summed = {kw: [] for kw in keywords}

    # Adding keywords for global data
    sum_keywords = {"somme", "total", "addition", "synthese", "synthèse"}

    for i, col in enumerate(headers):
        # No % data
        if "%" in str(col):
            continue

        tokens_col = clean_tokens(col)

        for kw in keywords:
            tokens_kw = clean_tokens(kw)
            # print(f"🔍 TEST kw='{kw}' | col='{col}' | tokens_kw={tokens_kw} | tokens_col={tokens_col}")
            if all(tok in tokens_col for tok in tokens_kw):
                # print(f"✅ VRAI MATCH : kw='{kw}' match avec col='{col}'")
                matched[kw].append(i)
                col_norm = col.lower().replace(" ", "")
                if any(word in tokens_col for word in sum_keywords) or "c5-c10" in col_norm or "c10-c40" in col_norm:
                    summed[kw].append(i)

    return matched, summed

def afficher_colonnes_detectees(columns_dict, df, titre="Colonnes détectées"):
    print(f"\n📊 {titre} :")
    for kw, index_list in columns_dict.items():
        if index_list:
            display = []
            for idx in index_list:
                try:
                    df_col_index = idx + 3  # Décalage logique vers df.columns
                    if df_col_index >= len(df.columns):
                        continue
                    col_name = df.columns[df_col_index]
                    col_letter = col_idx_to_excel_letter(df_col_index)
                    display.append(f"{col_letter} ({col_name})")
                except Exception as e:
                    print(f"⚠️ Erreur pour index '{idx}': {e}")
                    continue
            print(f"✅ {kw} → {display}")


def extraire_valeurs_generales(df, matched_columns, general_keywords):
    resultats = {}
    for kw in general_keywords:
        if kw.lower() == "hap":
            continue
        if matched_columns.get(kw):
            for idx in matched_columns[kw]:
                df_col_index = idx + 3
                if df_col_index >= len(df.columns):
                    continue
                col_name = df.columns[df_col_index]
                for i in range(len(df)):
                    try:
                        val = df.at[i, col_name]
                        if isinstance(val, pd.Series):
                            val = val[val.notna()]
                            val = val[val.astype(str).str.strip() != ""]
                            if val.empty:
                                continue
                            val = val.iloc[0]
                        elif pd.isna(val) or str(val).strip() == "":
                            continue
                    except Exception:
                        continue

                    val_str = str(val).strip()
                    if val_str.startswith("<"):
                        val_str = f"<LQ ({val_str})"

                    artelia = df.at[i, 'Code Artelia']
                    if pd.isna(artelia):
                        continue

                    if artelia not in resultats:
                        resultats[artelia] = {}

                    resultats[artelia][kw] = val_str
    return resultats

def extraire_valeurs_hap(df, matched_columns):
    resultats = {}
    def normalize(text):
        text = str(text).lower()
        return unicodedata.normalize('NFD', text).encode('ascii', 'ignore').decode('utf-8')

    for idx in matched_columns.get("hap", []):
        df_col_index = idx + 3
        if df_col_index >= len(df.columns):
            continue
        col_name = df.columns[df_col_index]
        col_name_norm = normalize(col_name)
        true_kw = "hap + naphtalène" if "naphtalene" in col_name_norm else "hap"
        # print(f"\n🔍 Colonne détectée : '{col_name}' → Classée comme '{true_kw}'")
        for i in range(len(df)):
            try:
                val = df.at[i, col_name]
                if isinstance(val, pd.Series):
                    val = val[val.notna()]
                    val = val[val.astype(str).str.strip() != ""]
                    if val.empty:
                        continue
                    val = val.iloc[0]
                elif pd.isna(val) or str(val).strip() == "":
                    continue
            except Exception:
                continue

            val_str = str(val).strip()
            if val_str.startswith("<"):
                val_str = f"<LQ ({val_str})"

            artelia = df.at[i, 'Code Artelia']
            if pd.isna(artelia):
                continue

            if artelia not in resultats:
                resultats[artelia] = {}

            resultats[artelia][true_kw] = val_str
            # print(f"✅ Ajout : Artelia = {artelia} | {true_kw} = {val_str}")
    return resultats


def afficher_groupement_par_artelia(df, resultats_artelia):
    print("\n📥 Résumé regroupé par code Artelia :")
    tous_codes_artelia = df['Code Artelia'].dropna().unique()

    for artelia in tous_codes_artelia:
        mesures = resultats_artelia.get(artelia, {})
        if not mesures:
            print(f"Artelia: {artelia} | Pas d'analyse détectée")
        else:
            ligne = f"Artelia: {artelia}"
            for comp, val in mesures.items():
                ligne += f" | {comp} = {val}"
            print(ligne)

    print("\n🔎 DEBUG — Valeurs HAP par échantillon :")
    for artelia in tous_codes_artelia:
        hap_val = resultats_artelia.get(artelia, {}).get("hap", "—")
        hap_naphtalene_val = resultats_artelia.get(artelia, {}).get("hap + naphtalène", "—")
        print(f"[HAP] - Artelia: {artelia} | HAP = {hap_val} | HAP + naphtalène = {hap_naphtalene_val}")


def additionner_parametres(df, resultats_artelia, liste_parametres, nom_somme="somme personnalisée"):
    for artelia in df['Code Artelia'].dropna().unique():
        total = 0.0
        valeurs_utilisees = 0

        for param in liste_parametres:
            valeur = resultats_artelia.get(artelia, {}).get(param)
            if valeur and not str(valeur).strip().startswith("<"):
                try:
                    total += float(str(valeur).replace(",", "."))
                    valeurs_utilisees += 1
                except ValueError:
                    print(f"⚠️ Valeur non convertible ignorée : '{valeur}' pour '{param}' (échantillon {artelia})")
                    continue

        if valeurs_utilisees > 0:
            resultats_artelia[artelia][nom_somme] = round(total, 3)



# = PROCESSING PART = #
#
sum = charger_groupes_parametres()

matched_columns, summed_columns = get_matching_columns(df.columns[3:], general_keywords)

afficher_colonnes_detectees(matched_columns, df, titre="Colonnes détectées par mot-clé")
afficher_colonnes_detectees(summed_columns, df, titre="Colonnes qui semblent déjà contenir une somme")

resultats_generaux = extraire_valeurs_generales(df, matched_columns, general_keywords)
resultats_hap = extraire_valeurs_hap(df, matched_columns)

resultats_artelia = defaultdict(dict)
for dico in [resultats_generaux, resultats_hap]:
    for k, v in dico.items():
        resultats_artelia[k].update(v)


for nom_somme, liste_parametres in groupes_personnalises.items():
    if nom_somme in general_keywords:  # c’est-à-dire dans la zone de droite
        additionner_parametres(df, resultats_artelia, liste_parametres, nom_somme)

with open("résumé_extraction.json", "w", encoding="utf-8") as f:
    json.dump(resultats_artelia, f, indent=2, ensure_ascii=False)

afficher_groupement_par_artelia(df, resultats_artelia)

# Based on flag : if --export is existent then export mode is activated
if export_mode:
    df_export = pd.DataFrame.from_dict(resultats_artelia, orient='index')
    df_export.reset_index(inplace=True)
    df_export.rename(columns={"index": "Code Artelia"}, inplace=True)

    colonnes_finales = []

    for col in general_keywords:
        if col == "Code Artelia":
            continue
        if col in df_export.columns:
            colonnes_finales.append(col)

    if "Code Artelia" not in colonnes_finales:
        colonnes_finales = ["Code Artelia"] + colonnes_finales
    else:
        colonnes_finales.insert(0, "Code Artelia")

    if "Code Eurofins" in general_keywords:
        df_export = df_export.merge(df[["Code Artelia", "Code Eurofins"]], on="Code Artelia", how="left")
        colonnes_finales.insert(1, "Code Eurofins")

    if "Date prélèvement" in general_keywords:
        df_export = df_export.merge(df[["Code Artelia", "Date prélèvement"]], on="Code Artelia", how="left")
        if "Code Eurofins" in general_keywords:
            colonnes_finales.insert(2, "Date prélèvement")
        else:
            colonnes_finales.insert(1, "Date prélèvement")


    # EXPORT

    df_export = df_export.loc[:, list(dict.fromkeys(colonnes_finales))]

    dossier_source = os.path.dirname(excel_path)
    nom_base = os.path.splitext(os.path.basename(excel_path))[0]
    horodatage = pd.Timestamp.today().strftime('%Y%m%d_%H%M')
    nom_fichier = f"{nom_base}_{horodatage}.xlsx"
    chemin_complet = os.path.join(dossier_source, nom_fichier)

    df_export.to_excel(chemin_complet, index=False)
    print(f"\n✅ Export terminé : {chemin_complet}")


with open("matched_columns.json", "w", encoding="utf-8") as f:
    json.dump(matched_columns, f, indent=2, ensure_ascii=False)

