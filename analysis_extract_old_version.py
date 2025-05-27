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
    # = v2.1 : Adding class type AGROLAB
    # = v2.2 : Using Rows and Columns from user to configure type of table - suppressing Agrolab/Eurofins type
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
        self.mapping_all = {}

    # FROM: config JSON file
    # TO:   setup: keywords, groups, mappings, fixed columns
    def load_keywords(self):
        with open(self.config_path, "r", encoding="utf-8") as f:
            data = json.load(f)

            if isinstance(data, dict):
                self.colonnes_finales_export = data.get("keywords_valides", [])
                self.keywords_raw = self.colonnes_finales_export
                self.groupes_personnalises = data.get("groupes_personnalises", {})
                self.keywords_raw = self.colonnes_finales_export

            # OLD FORMAT == KEEP FOR TESTS
            elif isinstance(data, list):
                self.colonnes_finales_export = data
                self.groupes_personnalises = {}

            else:
                raise ValueError("Format de configuration JSON non reconnu.")

            # =====================================================================

            # "Keywords_valides" creation
            if all(isinstance(kw, str) for kw in self.colonnes_finales_export):
                self.keywords_valides = [
                    kw.split("→")[0].strip() if "→" in kw else kw
                    for kw in self.colonnes_finales_export
                    if not kw.startswith("Code Artelia →")
                ]
            else:
                self.keywords_valides = list({
                    kw.split("→")[0].strip()
                    for kw in self.colonnes_finales_export
                    if isinstance(kw, str) and "→" in kw
                })
            # ========================== DEBUG ===============================
            # print("\n🧪 DEBUG - keywords_valides issus du JSON :")
            # for kw in self.keywords_valides:
            #     print(f"   🔹 {kw}")
            # ========================== DEBUG ===============================

        # Special mapping for "→ all" in kw
        for kw in self.keywords_valides:
            if "→ all" in kw:
                original_kw = kw.split("→")[0].strip()
                self.mapping_all[original_kw] = kw
        # Special mapping for fixed Columns =================================== From v0 to check
        colonnes_fixes = ["Code Artelia", "Code Eurofins", "Date prélèvement"]
        for col in colonnes_fixes:
            if col not in self.keywords_valides:
                self.keywords_valides.insert(0, col)

        # =====================================================================

        # Multiple match with kw in cible
        self.correspondances_explicites = {}  # type: dict[str, list[str]]
        for kw_raw in self.colonnes_finales_export:
            if isinstance(kw_raw, str) and "→" in kw_raw:
                kw, cible = [x.strip() for x in kw_raw.split("→", 1)]
                self.correspondances_explicites.setdefault(kw, []).append(cible)

    # FROM: row Excel columns name list
    # TO:   dict {keyword: [index, column_name]}
    def get_matching_columns(self, columns):
        # Create dict with keyword is a key associated with empty list = blocking multiple match
        matched = {kw: [] for kw in self.keywords_valides}

        # On all col from Excel file
        for i, col in enumerate(columns):
            if "%" in str(col):
                continue

            # Normalized token logic for cleaning
            tokens_col = clean_tokens(col)
            for kw in self.keywords_valides:
                tokens_kw = clean_tokens(kw)
                if all(tok in tokens_col for tok in tokens_kw):
                    # Adding index/column match for each without any multiple match
                    if (i, col) not in matched[kw]:
                        matched[kw].append((i, col))
        return matched

    # FROM: config JSON file extract excel
    # TO:   dict {keyword: [index, column_name]} for existing columns only
    def get_selected_from_config(self, columns, correspondances_explicites):
        matched = {}
        for kw, cibles in correspondances_explicites.items():
            matched[kw] = []
            for cible in cibles:
                if cible.strip() == "→ all":
                    continue
                for i, col in enumerate(columns):
                    if pd.isna(col):
                        continue
                    if str(col).strip() == str(cible).strip():
                        matched[kw].append((i, col))
        return matched

    def additionner_parametres(self, liste_parametres, nom_somme="somme personnalisée"):
        for artelia in self.resultats_artelia:
            total = 0.0
            valeurs_utilisees = 0

            for param in liste_parametres:
                # Nettoyer et normaliser les noms
                if "→" in param:
                    base = param.split("→")[0].strip()
                    version_all = f"{base} → all"
                else:
                    base = param.strip()
                    version_all = None

                valeur = None

                if version_all and version_all in self.resultats_artelia[artelia]:
                    valeur = self.resultats_artelia[artelia][version_all]
                elif base in self.resultats_artelia[artelia]:
                    valeur = self.resultats_artelia[artelia][base]


                if valeur and not str(valeur).strip().startswith("<"):
                    try:
                        total += float(str(valeur).replace(",", "."))
                        valeurs_utilisees += 1
                    except ValueError:
                        continue

            if valeurs_utilisees > 0:
                total_round = round(total, 3)
                self.resultats_artelia[artelia][nom_somme] = total_round
                print(f"DEBUG {artelia} – {nom_somme} : {total_round}")

    def _merge_column(self, df_export, col_name, colonnes_finales):
        if col_name not in self.df.columns:
            print(f"⚠️ '{col_name}' absent du DataFrame source")
            return df_export, colonnes_finales

        df_source = self.df[["Code Artelia", col_name]].drop_duplicates(subset=["Code Artelia"]).copy()
        df_export["_CodeArtelia_upper"] = df_export["Code Artelia"].str.upper()
        df_source["_CodeArtelia_upper"] = df_source["Code Artelia"].str.upper()

        df_export = df_export.merge(
            df_source,
            left_on="Code Artelia",
            right_index=True,
            how="left",
            suffixes=("", "_drop")
        )

        df_export.drop([c for c in df_export.columns if c.endswith("_drop")], axis=1, inplace=True)
        df_export.drop(columns=["_CodeArtelia_upper"], inplace=True)
        df_source.drop(columns=["_CodeArtelia_upper"], inplace=True)

        if col_name not in colonnes_finales:
            idx = 1
            if "Code Eurofins" in colonnes_finales and col_name != "Code Eurofins":
                idx = colonnes_finales.index("Code Eurofins") + 1
            colonnes_finales.insert(idx, col_name)
        return df_export, colonnes_finales


    def export(self):
        dossier = os.path.dirname(self.excel_path)
        nom_base = os.path.splitext(os.path.basename(self.excel_path))[0]
        horodatage = pd.Timestamp.today().strftime('%Y%m%d_%H%M')
        nom_fichier = os.path.join(dossier, f"{nom_base}_résumé_extraction_{horodatage}.xlsx")

        colonnes_finales = []
        df_export = pd.DataFrame.from_dict(self.resultats_artelia, orient='index')
        print("\n📦 CONTENU DE resultats_artelia:")
        import pprint
        pprint.pprint(self.resultats_artelia)

        print("\n🧪 DEBUG : Colonnes initiales du df_export")
        print(df_export.columns.tolist())

        if "Code Artelia" in df_export.columns:
            df_export.rename(columns={"Code Artelia": "Code Artelia_orig"}, inplace=True)

        df_export.reset_index(inplace=True)
        df_export.rename(columns={"index": "Code Artelia"}, inplace=True)

        colonnes_finales = [col for col in self.colonnes_finales_export if col in df_export.columns]

        print("\n🧪 DEBUG : colonnes_finales après filtrage par colonnes_finales_export")
        print(colonnes_finales)

        colonnes_finales = list(dict.fromkeys(colonnes_finales))

        if "Code Artelia" not in colonnes_finales:
            colonnes_finales.insert(0, "Code Artelia")

        if any("Code Eurofins" in col for col in self.colonnes_finales_export):
            df_export, colonnes_finales = self._merge_column(df_export, "Code Eurofins", colonnes_finales)

        if any("Date prélèvement" in col for col in self.colonnes_finales_export):
            df_export, colonnes_finales = self._merge_column(df_export, "Date prélèvement", colonnes_finales)

        # Excluding double values
        colonnes_finales = list(dict.fromkeys(colonnes_finales))

        print("\n🧪 DEBUG : colonnes_finales après insertions et nettoyage doublons")
        print(colonnes_finales)

        # Filter out groups elements if not required in the valid_keywords
        colonnes_a_exclure = set()

        print("\n🧪 DEBUG : keywords_valides")
        print(self.keywords_valides)

        for nom_groupe, sous_elements in self.groupes_personnalises.items():
            if nom_groupe not in self.keywords_valides:
                print(f"   ❌ Le groupe '{nom_groupe}' n'est pas sélectionné, on ne supprime rien")
                continue

            for sous in sous_elements:
                base_sous = sous.split("→")[0].strip()
                if base_sous not in self.keywords_valides:
                    print(f"   🚫 Suppression des colonnes liées à : {base_sous}")
                    colonnes_a_exclure.add(base_sous)

        print("\n🧪 DEBUG : Colonnes à exclure détectées")
        print(colonnes_a_exclure)

        colonnes_finales = [col for col in colonnes_finales if col not in colonnes_a_exclure]

        colonnes_autorisees = set(self.colonnes_finales_export + ["Code Artelia", "Code Eurofins", "Date prélèvement"])
        df_export = df_export.loc[:, [col for col in df_export.columns if col in colonnes_autorisees]]

        print("\n🧪 DEBUG : colonnes dans df_export juste avant Excel")
        print(df_export.columns.tolist())

        df_export.to_excel(nom_fichier, index=False)

        print(f"✅ Export terminé : {nom_fichier}")

        # =========================== DEBUG ======================================
        # with open("resume_extraction.json", "w", encoding="utf-8") as f:
        #     json.dump(self.resultats_artelia, f, indent=2, ensure_ascii=False)

    def _extract_hap(self):
        hap_matched = self.matched_columns.get("hap", [])
        for idx, nom_col in hap_matched:
            if idx >= len(self.df.columns):
                continue
            col_name = self.df.columns[idx]
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

    def load_data(self):
        raise NotImplementedError

    def extract(self):
        raise NotImplementedError

    def get_matched_columns(self):
        return self.matched_columns




# ================================================== #
# ==================== COLUMNS ===================== #
# ================================================== #

class ColumnsExtract(BaseExtract):
    def __init__(self, excel_path, config_path, sheet_name=None, col_config=None):
        super().__init__(excel_path, config_path, sheet_name)
        self.row_config = col_config or {}

    def load_data(self):
        df_raw = pd.read_excel(self.excel_path, sheet_name=self.sheet_name, header=None)
        r = self.row_config

        # Paramètres : ligne contenant les noms des paramètres
        noms_parametres = df_raw.iloc[r["param_row"], r["data_start_col"]:].tolist()

        # Ligne limite (si elle existe)
        ligne_limites = df_raw.iloc[r["row_limites"], r["data_start_col"]:].tolist() if r.get("row_limites") is not None else []

        # Suppression des colonnes marquées "limite"
        colonnes_a_exclure = []
        for i, val in enumerate(ligne_limites):
            if isinstance(val, str) and "limite" in val.lower():
                colonnes_a_exclure.append(i)

        for i in sorted(colonnes_a_exclure, reverse=True):
            del noms_parametres[i]

        valeurs = df_raw.iloc[r["data_start_row"]:, r["data_start_col"]:].copy()
        if colonnes_a_exclure:
            colonnes_a_supprimer = [valeurs.columns[i] for i in colonnes_a_exclure]
            valeurs.drop(columns=colonnes_a_supprimer, inplace=True)
            noms_parametres = [n for i, n in enumerate(noms_parametres) if i not in colonnes_a_exclure]

        # Appliquer les noms de colonnes
        valeurs.columns = noms_parametres

        # Ajouter infos échantillons
        noms_echantillons = df_raw.iloc[r["data_start_row"]:, r["nom_col"]].tolist()
        valeurs["Nom échantillon"] = noms_echantillons
        valeurs["Code Artelia"] = [str(n).strip() for n in noms_echantillons]

        # Champs optionnels (ex : date, code labo…)
        for nom_champ, (row, col) in r.get("optionnels", {}).items():
            try:
                valeurs[nom_champ] = df_raw.iloc[row:, col].reset_index(drop=True)
            except Exception as e:
                print(f"⚠️ Erreur chargement champ optionnel '{nom_champ}': {e}")

        self.df = valeurs.reset_index(drop=True)
        print("\n🔍 Vérification des Code Artelia lus dans load_data():")
        print(self.df["Code Artelia"])

    def extract(self):
        self.matched_columns = self.get_matching_columns(self.df.columns)
        matched_columns_all = self.get_matching_columns(self.df.columns)

        for i, row in self.df.iterrows():
            artelia = str(row["Nom échantillon"]).strip()

            for kw, cibles in self.correspondances_explicites.items():
                if "→ all" in cibles:
                    colonnes_possibles = matched_columns_all.get(kw, [])
                    for col_idx, nom_col in colonnes_possibles:
                        val = extraire_valeur(self.df.iloc[i, col_idx])
                        if val is not None and not str(val).strip().startswith("<"):
                            val_format = formater_valeur(val)
                            if kw not in self.resultats_artelia[artelia]:
                                self.resultats_artelia[artelia][kw] = val_format
                                print(f"{artelia} / {kw} → {self.resultats_artelia[artelia].get(kw)}")
                            break

                    if kw not in self.colonnes_finales_export:
                        self.colonnes_finales_export.append(kw)

                    continue

                suffixe = 1
                utilisees = set()

                for cible in cibles:
                    trouve = False
                    for col_idx, nom_col in matched_columns_all.get(kw, []):
                        if nom_col.strip() == cible.strip() and col_idx not in utilisees:
                            val = extraire_valeur(self.df.iloc[i, col_idx])
                            if val is not None:
                                val_format = formater_valeur(val)

                                nom_final = f"{kw}_{suffixe}" if len(cibles) > 1 else kw
                                self.resultats_artelia[artelia][nom_final] = val_format

                                if nom_final not in self.colonnes_finales_export:
                                    self.colonnes_finales_export.append(nom_final)

                                utilisees.add(col_idx)
                                trouve = True
                                break  # on passe à la cible suivante
                    suffixe += 1

        self.colonnes_finales_export = list(dict.fromkeys(self.colonnes_finales_export))

        self._extract_hap()

        for nom, liste in self.groupes_personnalises.items():
            if nom in self.colonnes_finales_export:
                self.additionner_parametres(liste, nom)




# ================================================== #
# ===================== ROWS ======================= #
# ================================================== #

class RowsExtract(BaseExtract):
    def __init__(self, excel_path, config_path, sheet_name=None, row_config=None):
        super().__init__(excel_path, config_path, sheet_name)
        self.row_config = row_config or {}

    def load_data(self):
        df_raw = pd.read_excel(self.excel_path, sheet_name=self.sheet_name, header=None)
        r = self.row_config

        noms_echantillons = df_raw.iloc[r["nom_row"], r["data_start_col"]:].tolist()

        ligne_limites = df_raw.iloc[r["row_limites"], r["data_start_col"]:].tolist() if r.get(
            "row_limites") is not None else []

        # If Limite exist in row 10 then suppressing data
        colonnes_a_exclure = []
        for i, val in enumerate(ligne_limites):
                if isinstance(val, str) and "limite" in val.lower():
                    colonnes_a_exclure.append(i)

        for i in sorted(colonnes_a_exclure, reverse=True):
            del noms_echantillons[i]

        valeurs = df_raw.iloc[r["data_start_row"]:, r["data_start_col"]:].copy()
        if colonnes_a_exclure:
            colonnes_a_supprimer = [valeurs.columns[i] for i in colonnes_a_exclure]
            valeurs.drop(columns=colonnes_a_supprimer, inplace=True)
            noms_echantillons = [n for i, n in enumerate(noms_echantillons) if i not in colonnes_a_exclure]

        row_param = r["param_row"]
        col_param = r["param_col"]

        valeurs.columns = noms_echantillons
        index_raw = df_raw.iloc[row_param:, col_param].tolist()

        if len(index_raw) > valeurs.shape[0]:
            print("📏 [DEBUG] Troncature des noms de paramètres")
            index_raw = index_raw[:valeurs.shape[0]]
        elif len(index_raw) < valeurs.shape[0]:
            print("📏 [DEBUG] Complétion des noms de paramètres")
            index_raw += [""] * (valeurs.shape[0] - len(index_raw))

        valeurs = valeurs.iloc[:len(index_raw)].copy()
        valeurs.index = index_raw

        self.df = valeurs.T.reset_index(drop=True)
        self.df.columns.name = None
        self.df["Nom échantillon"] = noms_echantillons
        self.df["Code Artelia"] = [str(n).strip() for n in noms_echantillons]

        # Optional part from UI 1 type
        for nom_champ, (row, col) in r.get("optionnels", {}).items():
            try:
                self.df[nom_champ] = df_raw.iloc[row, col:]
                if colonnes_a_exclure:
                    self.df[nom_champ] = self.df[nom_champ].drop(index=colonnes_a_exclure).reset_index(drop=True)
                else:
                    self.df[nom_champ] = self.df[nom_champ].reset_index(drop=True)
            except Exception as e:
                print(f"⚠️ Erreur chargement champ optionnel '{nom_champ}': {e}")

    def extract(self):
        matched_columns = self.get_matching_columns(self.df.columns)
        self.matched_columns = matched_columns

        print("\n📊 [DEBUG] Colonnes associées aux mots-clés (matched_columns) :")
        for kw, lst in matched_columns.items():
            print(f"  🔍 {kw} → {lst}")

        for i, row in self.df.iterrows():
            artelia = str(row["Nom échantillon"]).strip()

            for kw, col_infos in matched_columns.items():
                selections = [nom_col for _, nom_col in col_infos]
                # ========================= DEBUG =======================================
                # print(f"\n🔎 [DEBUG] Mot-clé : {kw}")
                # print(f"     ↪ col_infos : {col_infos}")
                # print(f"     ↪ selections (noms colonnes associés) : {selections}")
                # ========================= DEBUG =======================================

                if len(selections) == 1 and selections[0].strip().endswith("→ all"):
                    # 🔹 CAS SPÉCIAL : sélection unique "kw → all"
                    cible = selections[0].split("→", 1)[1].strip()
                    for col_idx, nom_col in col_infos:
                        if nom_col.strip() == cible:
                            val = extraire_valeur(self.df.iloc[i, col_idx])
                            if val is not None and not str(val).strip().startswith("<"):
                                val_format = formater_valeur(val)
                                print(f"     ✅ Valeur all trouvée : {val_format} (col {col_idx})")
                                self.resultats_artelia[artelia][f"{kw} → all"] = val_format
                                break

                elif len(selections) >= 1:
                    # 🔁 CAS GÉNÉRAL : sélection multiple (y compris all)
                    suffixe = 1
                    for sel in selections:
                        if "→" in sel:
                            cible = sel.split("→", 1)[1].strip()
                        else:
                            cible = sel.strip()

                        if sel.strip().endswith("→ all") and len(selections) == 1:
                            nom_final = kw
                        elif sel.strip().endswith("→ all"):
                            nom_final = f"{kw}_{suffixe}"
                        elif len(selections) == 1:
                            nom_final = kw
                        else:
                            nom_final = f"{kw}_{suffixe}"

                        # Creating a list of every kw with all its matching
                        for col_idx, nom_col in col_infos:
                            if nom_col.strip() == cible:
                                val = extraire_valeur(self.df.iloc[i, col_idx])
                                if val is not None:
                                    val_format = formater_valeur(val)
                                    self.resultats_artelia[artelia][nom_final] = val_format

                                    # Insert based on the user's UI2 order
                                    if nom_final not in self.colonnes_finales_export:
                                        if kw in self.colonnes_finales_export:
                                            idx = self.colonnes_finales_export.index(kw)
                                            self.colonnes_finales_export.insert(idx + suffixe, nom_final)
                                        else:
                                            self.colonnes_finales_export.append(nom_final)
                        suffixe += 1

        self._extract_hap()

        for nom, liste in self.groupes_personnalises.items():
            if nom in self.colonnes_finales_export:
                self.additionner_parametres(liste, nom)




# =========================== DEBUG ====================================
def afficher_colonnes_detectees(columns_dict, titre="Colonnes détectées"):
    print(f"\n📊 {titre} :")
    for kw, cols in columns_dict.items():
        if cols:
            print(f"✅ {kw} → colonnes index : {cols}")
        else:
            print(f"⚠️ {kw} → Aucune colonne détectée")