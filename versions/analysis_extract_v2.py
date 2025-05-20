import pandas as pd
import json
import os
import unicodedata
import re
from collections import defaultdict


# === Script : EXTRACT VALUE WITH KEYWORD IN AN EXCEL FORMAT - TABLEURS MULTIPLE PAR CLASSES ===
# = v1.0 : Test import from Excel raw DF-Excel and keyword-based extract
        # = v1.05 : Multiple keywords test
    # = v1.2 : Normalizing text and splitting data into keywords based on value or sum
    # = v1.3 : Extraction validated for general keyword with random pick
    # = v1.4 : Adapting the script to allow HAP to be separated from Naphtalene/HAP
    # = v1.5 : DEBUG
    # = v1.6 : Adding SUM calculation based on JSON file to allow local memory
    # = v1.7 : Link to UI and using json for keywords
# = v2.0 : PASSAGE FORMAT CLASSES DEPUIS EUROFINS_EXTRACT.PY
#
class BaseExtract:
    def __init__(self, excel_path, config_path, sheet_name=None):
        self.excel_path = excel_path
        self.config_path = config_path
        self.sheet_name = sheet_name
        self.df = None
        self.keywords_valides = []
        self.groupes_personnalises = {}
        self.resultats_artelia = defaultdict(dict)
        self.matched_columns = {}

    def load_keywords(self):
        with open(self.config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            if isinstance(data, dict):
                self.keywords_valides = data.get("keywords_valides", [])
                self.groupes_personnalises = data.get("groupes_personnalises", {})
            elif isinstance(data, list):
                self.keywords_valides = data
                self.groupes_personnalises = {}
            else:
                raise ValueError("Format de configuration JSON non reconnu.")

    def normalize(self, text):
        if pd.isna(text):
            return ""
        try:
            text = str(text)
        except Exception:
            return ""
        text = text.lower()
        return unicodedata.normalize('NFD', text).encode('ascii', 'ignore').decode('utf-8')
    def clean_tokens(self, text):
        return re.findall(r'[a-z0-9]+', self.normalize(text))

    def get_matching_columns(self, columns):
        matched = {kw: [] for kw in self.keywords_valides}
        for i, col in enumerate(columns):
            tokens_col = self.clean_tokens(col)
            for kw in self.keywords_valides:
                tokens_kw = self.clean_tokens(kw)
                if all(tok in tokens_col for tok in tokens_kw):
                    matched[kw].append(i)
        return matched

    def additionner_parametres(self, liste_parametres, nom_somme="somme personnalisée"):
        for artelia in self.resultats_artelia:
            total = 0.0
            valeurs_utilisees = 0
            for param in liste_parametres:
                valeur = self.resultats_artelia[artelia].get(param)
                if valeur and not str(valeur).strip().startswith("<"):
                    try:
                        total += float(str(valeur).replace(",", "."))
                        valeurs_utilisees += 1
                    except ValueError:
                        continue
            if valeurs_utilisees > 0:
                self.resultats_artelia[artelia][nom_somme] = round(total, 3)

    def export(self):
        dossier = os.path.dirname(self.excel_path)
        nom_base = os.path.splitext(os.path.basename(self.excel_path))[0]
        horodatage = pd.Timestamp.today().strftime('%Y%m%d_%H%M')
        nom_fichier = os.path.join(dossier, f"{nom_base}_résumé_extraction_{horodatage}.xlsx")

        df_export = pd.DataFrame.from_dict(self.resultats_artelia, orient='index')
        df_export.reset_index(inplace=True)
        df_export.rename(columns={"index": "Code Artelia"}, inplace=True)

        colonnes_finales = []

        for col in self.keywords_valides:
            if col == "Code Artelia":
                continue
            if col in df_export.columns:
                colonnes_finales.append(col)

        # Colonnes fixes à insérer
        if "Code Artelia" not in colonnes_finales:
            colonnes_finales.insert(0, "Code Artelia")

        if "Code Eurofins" in self.keywords_valides:
            df_export = df_export.merge(self.df[["Code Artelia", "Code Eurofins"]], on="Code Artelia", how="left")
            colonnes_finales.insert(1, "Code Eurofins")

        if "Date prélèvement" in self.keywords_valides:
            df_export = df_export.merge(self.df[["Code Artelia", "Date prélèvement"]], on="Code Artelia", how="left")
            idx = 2 if "Code Eurofins" in self.keywords_valides else 1
            colonnes_finales.insert(idx, "Date prélèvement")

        colonnes_finales = list(dict.fromkeys(colonnes_finales))
        df_export = df_export.loc[:, [col for col in colonnes_finales if col in df_export.columns]]
        df_export.to_excel(nom_fichier, index=False)
        print(f"✅ Export terminé : {nom_fichier}")
        with open("résumé_extraction.json", "w", encoding="utf-8") as f:
            json.dump(self.resultats_artelia, f, indent=2, ensure_ascii=False)

    def load_data(self):
        raise NotImplementedError

    def extract(self):
        raise NotImplementedError

    def get_matched_columns(self):
        return self.matched_columns

# ================================================== #
# ==================== EUROFINS ==================== #
# ================================================== #

class EurofinsExtract(BaseExtract):
    def load_data(self):
        df_raw = pd.read_excel(self.excel_path, sheet_name=self.sheet_name, header=None)
        headers = ['Code Eurofins', 'Code Artelia', 'Date prélèvement'] + df_raw.iloc[4, 3:].tolist()
        df = df_raw.iloc[5:].copy()
        df.columns = headers
        df = df.dropna(axis=1, how='all')
        df.columns = pd.Index([str(c).strip() for c in df.columns])
        self.df = df.reset_index(drop=True)
        if 'Code Artelia' not in df.columns:
            raise ValueError("❌ La colonne 'Code Artelia' est absente du fichier. Vérifiez le format.")

    def extract(self):
        # Matching des colonnes (on obtient des indices de colonnes relatives à self.df.columns[3:])
        all_matched = self.get_matching_columns(self.df.columns[3:])
        self.matched_columns = all_matched

        # === Extraction des valeurs générales (hors HAP) ===
        for i in range(len(self.df)):
            artelia = self.df.at[i, 'Code Artelia']
            if pd.isna(artelia):
                continue

            for kw, index_list in all_matched.items():
                if kw.lower() == "hap":
                    continue
                for idx in index_list:
                    df_col_index = idx + 3  # alignement avec df.columns (comme dans le script original)
                    if df_col_index >= len(self.df.columns):
                        continue
                    col_name = self.df.columns[df_col_index]
                    try:
                        val = self.df.at[i, col_name]
                        if isinstance(val, pd.Series):
                            val = val.dropna().astype(str).str.strip()
                            if val.empty:
                                continue
                            val = val.iloc[0]
                        elif pd.isna(val) or str(val).strip() == "":
                            continue

                        val_str = str(val).strip()
                        if val_str.startswith("<"):
                            val_str = f"<LQ ({val_str})"

                        self.resultats_artelia[artelia][kw] = val_str
                        break  # une seule valeur par mot-clé
                    except Exception:
                        continue

        # === Extraction spécifique HAP ===
        hap_matched = all_matched.get("hap", [])
        for idx in hap_matched:
            df_col_index = idx + 3
            if df_col_index >= len(self.df.columns):
                continue
            col_name = self.df.columns[df_col_index]
            col_name_norm = self.normalize(col_name)
            true_kw = "hap + naphtalène" if "naphtalene" in col_name_norm else "hap"
            if true_kw not in self.keywords_valides:
                continue

            for i in range(len(self.df)):
                try:
                    val = self.df.at[i, col_name]
                    if isinstance(val, pd.Series):
                        val = val.dropna().astype(str).str.strip()
                        if val.empty:
                            continue
                        val = val.iloc[0]
                    elif pd.isna(val) or str(val).strip() == "":
                        continue

                    val_str = str(val).strip()
                    if val_str.startswith("<"):
                        val_str = f"<LQ ({val_str})"

                    artelia = self.df.at[i, 'Code Artelia']
                    if pd.isna(artelia):
                        continue

                    self.resultats_artelia[artelia][true_kw] = val_str
                except Exception:
                    continue

        # === Sommes personnalisées ===
        for nom, liste in self.groupes_personnalises.items():
            if nom in self.keywords_valides:
                self.additionner_parametres(liste, nom)


# ================================================== #
# ==================== AGROLAB ===================== #
# ================================================== #

class AgrolabExtract(BaseExtract):
    def load_data(self):
        df_raw = pd.read_excel(self.excel_path, sheet_name=self.sheet_name, header=None)
        df_raw = df_raw.dropna(how='all')
        df_raw = df_raw.dropna(axis=1, how='all')
        df_raw.columns = df_raw.iloc[:, 0]
        df = df_raw.iloc[:, 1:].T
        df.columns = [str(c).strip() for c in df.columns]
        self.df = df.reset_index(drop=True)

    def extract(self):
        matched_columns = self.get_matching_columns(self.df.columns)
        self.matched_columns = self.get_matching_columns(self.df.columns)
        for i, row in self.df.iterrows():
            artelia = f"ECHANTILLON_{i+1}"
            for kw, cols in matched_columns.items():
                for col in cols:
                    try:
                        val = self.df.at[i, col]
                        if isinstance(val, pd.Series):
                            val = val.dropna().astype(str).str.strip()
                            if val.empty:
                                continue
                            val = val.iloc[0]
                        elif pd.isna(val) or str(val).strip() == "":
                            continue

                        val_str = str(val).strip()
                        if val_str.startswith("<"):
                            val_str = f"<LQ ({val_str})"
                        self.resultats_artelia[artelia][kw] = val_str
                        break
                    except Exception:
                        continue
        for nom, params in self.groupes_personnalises.items():
            self.additionner_parametres(params, nom)