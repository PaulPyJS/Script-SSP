import pandas as pd
import unicodedata
import re


# === Script : EXTRACT VALUE WITH KEYWORD IN AN EXCEL FORMAT - TABLEURS EUROFINS ===
# = v1 : Test import from Excel raw DF-Excel and keyword-based extract
# = v1.5 : Multiple keywords test
# = v2 : Normalizing text and splitting data into keywords based on value or sum
# = v3 : Extraction validated for general keyword with random pick
# = v4 : Adapting the script to allow HAP to be separated from Naphtalene/HAP
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
            #print(f"🔍 TEST kw='{kw}' | col='{col}' | tokens_kw={tokens_kw} | tokens_col={tokens_col}")
            if all(tok in tokens_col for tok in tokens_kw):
                #print(f"✅ VRAI MATCH : kw='{kw}' match avec col='{col}'")
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







# = PROCESSING PART = #
#
general_keywords = [
    "benzene",
    "hap", # Special treatment for logical issue regarding naphta

    # SHORT CARBON
    "c5 - c10", # = TOTAL !
    "c6 - c8", # ALI
    "c8 - c10", # ALI ou TOTAL
    "c5 - c6", # ALI
    "c9 - c10", # ARO
    "c6 - c9", # ARO
    "c5 - c8", # = TOTAL !

    # LONG CARBON
    "c10 - c40", # = TOTAL !

    # 8 PARTS
    "c10 - c12",
    "c12 - c16",
    "c16 - c20",
    "c20 - c24",
    "c24 - c28",
    "c28 - c32",
    "c32 - c36",
    "c36 - c40",
    # 4 PARTS
    "nc10 - nc16",
    "nc16 - nc22",
    "nc22 - nc30",
    "nc30 - nc40",
]

matched_columns, summed_columns = get_matching_columns(df.columns[3:], general_keywords)

for kw in general_keywords:
    print(f"🔹 {kw} → {len(matched_columns[kw])} colonnes valides conservées")

afficher_colonnes_detectees(matched_columns, df, titre="Colonnes détectées par mot-clé")
afficher_colonnes_detectees(summed_columns, df, titre="Colonnes qui semblent déjà contenir une somme")


print("\n📥 Résumé regroupé par code Artelia :")
resultats_artelia = {}

# Extraction for all general component except HAP
for kw in general_keywords:
    if kw.lower() == "hap":
        continue  # on le gère à part après

    if matched_columns.get(kw):
        for idx in matched_columns[kw]:
            df_col_index = idx + 3  # car matched_columns est basé sur df.columns[3:]
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

                if artelia not in resultats_artelia:
                    resultats_artelia[artelia] = {}

                resultats_artelia[artelia][kw] = val_str

# Extraction for HAP and Naphtalene
if matched_columns.get("hap"):
    for idx in matched_columns["hap"]:
        df_col_index = idx + 3  # passage de l’index relatif au df.columns
        if df_col_index >= len(df.columns):
            continue
        col_name = df.columns[df_col_index]

        col_name_norm = normalize(col_name)
        is_naphtalene = "naphtalene" in col_name_norm
        true_kw = "hap + naphtalène" if is_naphtalene else "hap"

        print(f"\n🔍 Colonne détectée : '{col_name}'")
        print(f"   → Classée comme '{true_kw}'")

        for i in range(len(df)):
            try:
                val = df.at[i, col_name]
                # même logique pour filtrer les valeurs valides
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

            if artelia not in resultats_artelia:
                resultats_artelia[artelia] = {}

            resultats_artelia[artelia][true_kw] = val_str
            print(f"✅ Ajout : Artelia = {artelia} | {true_kw} = {val_str}")


# Grouping data
tous_codes_artelia = df['Code Artelia'].dropna().unique()

for artelia in tous_codes_artelia:
    mesures = resultats_artelia.get(artelia, {})
    if not mesures:
        print(f"🎯 Artelia: {artelia} | Pas d'analyse détectée")
    else:
        ligne = f"🎯 Artelia: {artelia}"
        for comp, val in mesures.items():
            ligne += f" | {comp} = {val}"
        print(ligne)

print("\n🔎 DEBUG — Valeurs HAP par échantillon :")
for artelia in tous_codes_artelia:
    hap_val = resultats_artelia.get(artelia, {}).get("hap", "—")
    hap_naphtalene_val = resultats_artelia.get(artelia, {}).get("hap + naphtalène", "—")

    print(f"🧪 Artelia: {artelia} | HAP = {hap_val} | HAP + naphtalène = {hap_naphtalene_val}")



# = EXPORTING PART
#
# df_export = pd.DataFrame.from_dict(resultats_artelia, orient='index')
#
# # Réinitialisation de l’index (Code Artelia devient une colonne)
# df_export.reset_index(inplace=True)
# df_export.rename(columns={"index": "Code Artelia"}, inplace=True)
#
# # On s’assure que tous les codes Artelia sont inclus, même sans valeurs
# df_export_complet = pd.DataFrame({'Code Artelia': tous_codes_artelia})
# df_export = pd.merge(df_export_complet, df_export, on='Code Artelia', how='left')
#
# # ✅ Export dans un fichier Excel
# output_file = "résumé_extraction_artelia.xlsx"
# df_export.to_excel(output_file, index=False)
# print(f"\n✅ Résumé exporté dans le fichier : {output_file}")