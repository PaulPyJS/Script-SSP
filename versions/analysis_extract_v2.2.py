import pandas as pd
import json
import os
from collections import defaultdict

from extract_utils import normalize, clean_tokens, extraire_valeur, formater_valeur


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
    # = v2.1 : Ajout classe type AGROLAB
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
                self.colonnes_finales_export = data.get("keywords_valides", [])
                self.groupes_personnalises = data.get("groupes_personnalises", {})

                self.keywords_valides = list({kw.split("→")[0].strip() for kw in self.colonnes_finales_export})
            elif isinstance(data, list):
                self.colonnes_finales_export = data
                self.keywords_valides = data
                self.groupes_personnalises = {}
            else:
                raise ValueError("Format de configuration JSON non reconnu.")

        for group in self.groupes_personnalises.values():
            for param in group:
                if param not in self.keywords_valides:
                    self.keywords_valides.append(param)

        self.mapping_all = {}
        for kw in self.keywords_valides:
            if "→ all" in kw:
                original_kw = kw.split("→")[0].strip()
                self.mapping_all[original_kw] = kw

        colonnes_fixes = ["Code Artelia", "Code Eurofins", "Date prélèvement"]
        for col in colonnes_fixes:
            if col not in self.keywords_valides:
                self.keywords_valides.insert(0, col)

    def get_matching_columns(self, columns):
        matched = {kw: [] for kw in self.keywords_valides}
        for i, col in enumerate(columns):
            if "%" in str(col):
                continue
            tokens_col = clean_tokens(col)
            for kw in self.keywords_valides:
                tokens_kw = clean_tokens(kw)
                if all(tok in tokens_col for tok in tokens_kw):
                    matched[kw].append((i, col))
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

        for col in self.colonnes_finales_export:
            if col in df_export.columns and col not in colonnes_finales:
                colonnes_finales.append(col)
            elif "→" in col:
                # Exemple : "benzene → all" → tester "benzene"
                kw_simple = col.split("→")[0].strip()
                if kw_simple in df_export.columns and kw_simple not in colonnes_finales:
                    colonnes_finales.append(kw_simple)

            # Ajout des groupes calculés qui seraient présents dans les colonnes
        for nom_groupe in self.groupes_personnalises:
            if nom_groupe in df_export.columns and nom_groupe not in colonnes_finales:
                colonnes_finales.append(nom_groupe)

        if "Code Artelia" not in colonnes_finales:
            colonnes_finales.insert(0, "Code Artelia")

        if any("Code Eurofins" in col for col in self.colonnes_finales_export):
            if "Code Eurofins" not in self.df.columns:
                print("⚠️ 'Code Eurofins' absent du DataFrame source")
            else:
                df_export = df_export.merge(
                    self.df[["Code Artelia", "Code Eurofins"]],
                    on="Code Artelia", how="left"
                )
                if "Code Eurofins" not in colonnes_finales:
                    colonnes_finales.insert(1, "Code Eurofins")

        if any("Date prélèvement" in col for col in self.colonnes_finales_export):
            if "Date prélèvement" not in self.df.columns:
                print("⚠️ 'Date prélèvement' absent du DataFrame source")
            else:
                df_export = df_export.merge(
                    self.df[["Code Artelia", "Date prélèvement"]],
                    on="Code Artelia", how="left"
                )
                idx = colonnes_finales.index("Code Eurofins") + 1 if "Code Eurofins" in colonnes_finales else 1
                if "Date prélèvement" not in colonnes_finales:
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
        all_matched = self.get_matching_columns(self.df.columns[3:])
        self.matched_columns = all_matched

        for i in range(len(self.df)):
            artelia = self.df.at[i, 'Code Artelia']
            if pd.isna(artelia):
                continue

            for kw, col_infos in all_matched.items():
                if kw.lower() == "hap":
                    continue

                if f"{kw} → all" in self.keywords_valides:
                    # Cas "→ all" : on prend la première valeur trouvée
                    for idx, nom_col in col_infos:
                        df_col_index = idx + 3
                        if df_col_index >= len(self.df.columns):
                            continue
                        col_name = self.df.columns[df_col_index]
                        val = extraire_valeur(self.df.at[i, col_name])
                        if val is not None:
                            self.resultats_artelia[artelia][f"{kw} → all"] = formater_valeur(val)
                            break
                else:
                    # Cas "→ nom exact"
                    if kw not in self.keywords_valides:
                        continue
                    for idx, nom_col in col_infos:
                        df_col_index = idx + 3
                        if df_col_index >= len(self.df.columns):
                            continue
                        col_name = self.df.columns[df_col_index]
                        val = extraire_valeur(self.df.at[i, col_name])
                        if val is not None:
                            self.resultats_artelia[artelia][kw] = formater_valeur(val)
                            break

        # === HAP ===
        hap_matched = all_matched.get("hap", [])
        for idx, nom_col in hap_matched:
            df_col_index = idx + 3
            if df_col_index >= len(self.df.columns):
                continue
            col_name = self.df.columns[df_col_index]
            col_name_norm = normalize(col_name)
            true_kw = "hap + naphtalène" if "naphtalene" in col_name_norm else "hap"
            true_kw = true_kw.lower()
            if true_kw not in self.keywords_valides:
                continue

            if f"{true_kw} → all" in self.keywords_valides:
                nom_final = f"{true_kw} → all"
            elif true_kw in self.keywords_valides:
                nom_final = true_kw
            else:
                continue

            for i in range(len(self.df)):
                try:
                    val = extraire_valeur(self.df.at[i, col_name])
                    if val is None:
                        continue
                    val_str = formater_valeur(val)

                    artelia = self.df.at[i, 'Code Artelia']
                    if pd.isna(artelia):
                        continue

                    self.resultats_artelia[artelia][nom_final] = val_str
                except Exception:
                    continue

        # === Sommes personnalisées ===
        for nom, liste in self.groupes_personnalises.items():
            if nom in self.keywords_valides:
                self.additionner_parametres(liste, nom)


# ================================================== #
# ==================== AGROLAB ===================== #
# ================================================== #

class RowExtract(BaseExtract):
    def load_data(self):
        df_raw = pd.read_excel(self.excel_path, sheet_name=self.sheet_name, header=None)

        noms_echantillons = df_raw.iloc[7, 4:].tolist()
        numeros = df_raw.iloc[6, 4:].tolist()
        dates = df_raw.iloc[8, 4:].tolist()
        ligne10 = df_raw.iloc[9, 4:].tolist()

        # If Limite exist in row 10 then suppressing data
        colonnes_a_exclure = []
        for i, val in enumerate(ligne10):
                if isinstance(val, str) and "limite" in val.lower():
                    colonnes_a_exclure.append(i)

        for i in sorted(colonnes_a_exclure, reverse=True):
            del noms_echantillons[i]
            del numeros[i]
            del dates[i]

        # Values data from E11
        valeurs = df_raw.iloc[10:, 4:].copy()
        if colonnes_a_exclure:
            valeurs.drop(valeurs.columns[colonnes_a_exclure], axis=1, inplace=True)

        valeurs.columns = noms_echantillons
        valeurs.index = df_raw.iloc[10:, 0].tolist()  # colonne A → paramètres

        self.df = valeurs.T.reset_index(drop=True)
        self.df.columns.name = None
        self.df["Nom échantillon"] = noms_echantillons
        self.df["Numéro échantillon"] = numeros
        self.df["Date"] = dates
        self.df["Code Artelia"] = [str(n).strip() for n in noms_echantillons]

    def extract(self):
        matched_columns = self.get_matching_columns(self.df.columns)
        self.matched_columns = matched_columns

        # =========== DEBUG ==================
        # afficher_colonnes_detectees(matched_columns, titre="🔍 Colonnes correspondant aux mots-clés")
        # ====================================

        for i, row in self.df.iterrows():
            artelia = str(row["Nom échantillon"]).strip()

            for kw, col_infos in matched_columns.items():
                if f"{kw} → all" in self.keywords_valides:
                    for col_idx, nom_col in col_infos:
                        val = extraire_valeur(self.df.iloc[i, col_idx])
                        if val is not None:
                            self.resultats_artelia[artelia][f"{kw} → all"] = formater_valeur(val)
                            break
                else:
                    if kw not in self.keywords_valides:
                        continue
                    for col_idx, nom_col in col_infos:
                        val = extraire_valeur(self.df.iloc[i, col_idx])
                        if val is not None:
                            self.resultats_artelia[artelia][kw] = formater_valeur(val)
                            break

        for nom, params in self.groupes_personnalises.items():
            self.additionner_parametres(params, nom)





# =========================== DEBUG ====================================
def afficher_colonnes_detectees(columns_dict, titre="Colonnes détectées"):
    print(f"\n📊 {titre} :")
    for kw, cols in columns_dict.items():
        if cols:
            print(f"✅ {kw} → colonnes index : {cols}")
        else:
            print(f"⚠️ {kw} → Aucune colonne détectée")