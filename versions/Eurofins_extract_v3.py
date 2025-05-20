import pandas as pd
import unicodedata
import re
import numpy as np


# === Script : EXTRACT VALUE WITH KEYWORD IN AN EXCEL FORMAT - TABLEURS EUROFINS ===
# = v1 : Test import from Excel raw DF-Excel and keyword-based extract
# = v1.5 : Multiple keywords test
# = v2 : Normalizing text and splitting data into keywords based on value or sum
# = v3 : Extraction validated for general keyword with random pick
#

file_path = "00_INPUT/Résultats Eurofins 06052025.xlsm"
sheet_name = "Comparer les échantillons"

df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
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






# = FUNCTIONS PART
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
            #print(f"🔍 TEST kw='{kw}' | col='{col}' | tokens_kw={tokens_kw} | tokens_col={tokens_col}")
            if all(tok in tokens_col for tok in tokens_kw):
                #print(f"✅ VRAI MATCH : kw='{kw}' match avec col='{col}'")
                matched[kw].append(i)
                col_norm = col.lower().replace(" ", "")
                if any(word in tokens_col for word in sum_keywords) or "c5-c10" in col_norm or "c10-c40" in col_norm:
                    summed[kw].append(i)

    return matched, summed

def afficher_colonnes_detectees(columns_dict, header_index_map, df, titre="Colonnes détectées"):
    print(f"\n📊 {titre} :")
    for kw, index_list in columns_dict.items():
        if index_list:
            display = []
            for i in index_list:
                col_name = header_index_map.get(i)

                if not col_name or col_name not in df.columns:
                    continue

                try:
                    # ✅ Affichage uniquement : on revient à l’index Excel réel
                    col_letter = col_idx_to_excel_letter(i + 3)
                    display.append(f"{col_letter} ({col_name})")

                except Exception as e:
                    print(f"⚠️ Erreur pour colonne '{col_name}': {e}")
                    continue

            print(f"✅ {kw} → {display}")


# ==================== DEBUG FUNCTION =============================
# Used to convert index to Excel columns
def col_idx_to_excel_letter(idx):
    letter = ''
    while idx >= 0:
        letter = chr(idx % 26 + 65) + letter
        idx = idx // 26 - 1
    return letter
# ==================== DEBUG FUNCTION =============================




# = PROCESSING PART
#
general_keywords = [
    "benzene",
    "hap",
    "somme des hap",
    "somme 15 hap",
    "c5 - c10",
    "c10 - c16",
    "c10 - c40",
    "c10 - c12",
    "c12 - c16",
    "c16 - c20",
    "c20 - c24",
    "c24 - c28",
    "c28 - c32",
    "c32 - c36",
    "c36 - c40",
    "nc10 - nc16",
    "nc16 - nc22",
    "nc22 - nc30",
    "nc30 - nc40",
    "hct nc10 - nc16",
    "hct >nc16 - nc22",
    "indice hydrocarbures (c10 - c40)"
]

matched_columns, summed_columns = get_matching_columns(header_data_only, general_keywords)
# 🔎 On filtre les index invalides
max_index = len(df.columns) - 1
for kw in general_keywords:
    matched_columns[kw] = [i for i in matched_columns[kw] if i + 3 <= max_index]
    summed_columns[kw] = [i for i in summed_columns[kw] if i + 3 <= max_index]
print("\n📊 Vérification finale :")
for kw in general_keywords:
    print(f"🔹 {kw} → {len(matched_columns[kw])} colonnes valides conservées")
print("\n🔎 Vérification des index détectés...")
for kw in general_keywords:
    for idx_col in matched_columns[kw]:
        df_col_index = idx_col + 3  # Décalage logique
        if df_col_index >= len(df.columns):
            print(f"❌ Problème : kw '{kw}' → idx brut {idx_col} (+3 → {df_col_index}) dépasse df.columns (len={len(df.columns)})")

print("\n=== DEBUG HEADERS NORMALISÉS ===")
for i, col in enumerate(df_raw.iloc[4]):
    print(i, "→", normalize(col).lower())

afficher_colonnes_detectees(matched_columns, header_index_map, df, titre="Colonnes détectées par mot-clé")
afficher_colonnes_detectees(summed_columns, header_index_map, df, titre="Colonnes qui semblent déjà contenir une somme")

print("\n📥 Extraction des valeurs par mot-clé avec contexte (Eurofins / Artelia / Date) :")
for kw in general_keywords:
    if matched_columns[kw]:
        print(f"\n🔍 Résultats pour {kw} :")
        for idx_col in matched_columns[kw]:
            if idx_col not in header_index_map:
                print(f"⚠️ Index {idx_col} introuvable dans header_index_map")
                continue

            col_name = header_index_map[idx_col]
            if col_name not in df.columns:
                print(f"⚠️ col_name '{col_name}' pas trouvé dans df.columns")
                continue

            print(f"🧪 DEBUG EXTRACTION → kw={kw} | idx_col={idx_col} | col_name={col_name}")

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
                except Exception as e:
                    print(f"⚠️ Erreur accès cellule ({i}, {col_name}) : {e}")
                    continue

                val_str = str(val).strip()

                if val_str.startswith("<"):
                    val_str = f"<LQ ({val_str})"

                eurofins = df.at[i, 'Code Eurofins']
                artelia = df.at[i, 'Code Artelia']
                date = df.at[i, 'Date prélèvement']

                print(f"  - Eurofins: {eurofins} | Artelia: {artelia} | Date: {date} | {col_name} = {val_str}")



# EXPORTING PART
# if df_results.empty:
#     print("⚠ Aucun résultat trouvé.")
# else:
#     print("\n✅ Aperçu des résultats :")
#     print(df_results.head())
#
#     columns_to_export = ["Code Artelia"] + [kw for kw in keywords]
#     df_results_export = df_results[columns_to_export]
#
#     df_results_export.to_excel("résultats_triés.xlsx", index=False)
#     print("✅ Résultats enregistrés dans 'résultats_triés.xlsx' (résumé complet logique VBA)")