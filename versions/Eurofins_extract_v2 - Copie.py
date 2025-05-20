import pandas as pd
import unicodedata
import re
import string


# === Script : EXTRACT VALUE WITH KEYWORD IN AN EXCEL FORMAT - TABLEURS EUROFINS ===
# = v1 : Test import from Excel raw DF-Excel and keyword-based extract
# = v1.5 : Multiple keywords test
# = v2 : Normalizing text and splitting data into keywords based on value or sum
#

file_path = "00_INPUT/Résultats Eurofins 06052025.xlsm"
sheet_name = "Comparer les échantillons"

df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
print("✅ Fichier chargé.")

# Based on Excel format from Eurofins, without blank parts
# Starting with CodeEuro/CodeArt/Date + Data from L5-C4 index 4,3
headers = ['Code Eurofins', 'Code Artelia', 'Date prélèvement'] + df_raw.iloc[4, 3:].tolist()

# Data values from line 6, index 5
df = df_raw.iloc[5:].copy()
df.columns = headers
df = df.reset_index(drop=True)

df = df.dropna(axis=1, how='all')
df.columns = [str(c).strip() for c in df.columns]

header_row = df_raw.iloc[4].tolist()





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
            if all(tok in tokens_col for tok in tokens_kw):  # tous les mots du mot-clé sont dans les tokens
                matched[kw].append(i)
                col_norm = col.lower().replace(" ", "")
                if any(word in tokens_col for word in sum_keywords) or "c5-c10" in col_norm or "c10-c40" in col_norm:
                    summed[kw].append(i)

    return matched, summed




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

matched_columns, summed_columns = get_matching_columns(header_row, general_keywords)

print("\n=== DEBUG HEADERS NORMALISÉS ===")
for i, col in enumerate(df_raw.iloc[4]):
    print(i, "→", normalize(col).lower())

print("\n📌 Colonnes détectées par mot-clé :")
for kw in general_keywords:
    if matched_columns[kw]:
        cols = matched_columns[kw]
        display = [f"{col_idx_to_excel_letter(i)} ({header_row[i]})" for i in cols]
        print(f"🔹 {kw} → {display}")

print("\n📊 Colonnes qui semblent déjà contenir une somme :")
for kw in general_keywords:
    if summed_columns[kw]:
        cols = summed_columns[kw]
        display = [f"{col_idx_to_excel_letter(i)} ({header_row[i]})" for i in cols]
        print(f"✅ {kw} → {display}")







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