import pandas as pd
import os
import json
from extract_utils import clean_tokens, cell_to_index, values_lq_or_none


class BaseExtract:
    def __init__(self, excel_path, json_config_path, sheet_name, config, input_zone_gauche = None):
        self.excel_path = excel_path
        self.json_config_path = json_config_path
        self.sheet_name = sheet_name
        self.config = config
        self.df = None
        self.resultats = {}
        self.keywords_valides = []
        self.groupes_personnalises = {}
        self.input_zone_gauche = input_zone_gauche or []


    # INPUT:
    #   path_config (str): Path to JSON file containing a simple list of keywords.
    # OUTPUT:
    #   keywords (list[str]): List of base keywords, unchanged (not normalized here).
    @staticmethod
    def load_keywords_ui1(path_config: str) -> list[str]:
        with open(path_config, "r", encoding="utf-8") as f:
            data = json.load(f)

        if not isinstance(data, list):
            raise ValueError("Expected a JSON list of keywords as input.")

        return [kw.strip() for kw in data if isinstance(kw, str)]


    # INPUT:
    #   columns (list[str]): List of column names from a row in Excel.
    #   keywords (list[str]): List of base keywords to match against columns.
    # OUTPUT:
    #   matched (dict[str, list[tuple[int, str]]]):
    #       Dict of matches: {keyword → list of (index, column_name) found}.
    #   multiple_matches: list[str] → keywords with multiple column matches
    @staticmethod
    def get_matching_columns(columns: list[str], keywords: list[str]) -> tuple[dict[str, list[tuple[int, str]]], list[str]]:
        matched = {kw: [] for kw in keywords}
        multiple_matches = []

        for i, col in enumerate(columns):
            if "%" in str(col):
                continue
            tokens_col = clean_tokens(str(col))
            for kw in keywords:
                tokens_kw = clean_tokens(kw)
                if all(tok in tokens_col for tok in tokens_kw):
                    if (i, col) not in matched[kw]:
                        matched[kw].append((i, col))

        # NEW LIST FOR :  → all if multiple match for 1 kw
        for kw, matches in matched.items():
            if len(matches) > 1:
                multiple_matches.append(kw)

        return matched, multiple_matches

    # INPUT:
    #   item (str): The keyword to extract (can be "→ all", "→ (index, name)", or "→ column name").
    #   df (pd.DataFrame): The full Excel DataFrame.
    #   idx_row (int): The index of the current row (sample) to extract from.
    #   noms_colonnes (list): List of column headers at the param_row.
    #   correspondances_input (dict): Dict containing keyword → list of (col_idx, col_name) mappings.
    # OUTPUT:
    #   Value (str, float, or "") from the row
    def extract_values(self, item, df, noms_reference, correspondances_input, axis, idx=None):
        # STEP 1 try : → all
        # On item : item is membre for groupes_personnalises or kw for keyword_valides
        idx_col, idx_ligne = (idx, None) if axis == "rows" else (None, idx)

        if "→ all" in item:
            # Reminder : correspondance_input = {KW → all : [(idx, nom),(idx, nom),()]}
            #            item is membre for groups & kw for base  = example : "toluene → (68, Toluène)"
            #                                                              or "toluene → all"
            #
            match_possibles = correspondances_input.get(item.strip(), [])
            print(f"\n🔄 Traitement '{item}' avec colonnes possibles :", match_possibles)

            valeurs_possibles = []
            # Using valeur_possible to take multiples ones but using the first [0]
            for idx_possibles, nom in match_possibles:
                try:
                    if axis == "rows":
                        val = df.iat[idx_possibles, idx_col]
                        if pd.notna(val) and str(val).strip() != "":
                            valeurs_possibles.append(val)
                    elif axis == "columns":
                        val = df.iat[idx_ligne, idx_possibles]
                        if pd.notna(val) and str(val).strip() != "":
                            valeurs_possibles.append(val)
                except Exception as e:
                    print(f"❌ Erreur accès {axis} '{nom}' : {e}")
                    continue
            # Using first value found
            return values_lq_or_none(valeurs_possibles[0]) if valeurs_possibles else ""

        # STEP 2 : → + real column name
        elif "→" in item:
            try:
                _, cible = map(str.strip, item.split("→", 1))
                if cible.startswith("(") and "," in cible:
                    idx_str, _ = cible.strip("()").split(",", 1)
                    idx_possible = int(idx_str.strip())
                    val = df.iat[idx_ligne, idx_possible] if axis == "columns" else df.iat[idx_possible, idx_col]
                    return values_lq_or_none(val)
                else:
                    idx_possible = noms_reference.index(cible)
                    val = df.iat[idx_ligne, idx_possible] if axis == "columns" else df.iat[idx_possible, idx_col]
                    return values_lq_or_none(val)
            except Exception as e:
                print(f"❌ Erreur sur item '{item}' : {e}")
                return ""

        # STEP 3 : No → in data = " " securize the ransomize
        elif "(" not in item:
                return ""

        # STEP 4 : No → in data but ( ok = Part from the randomiz
        else:
            correspondances = correspondances_input.get(item, [])
            if len(correspondances) == 1:
                idx_ref, nom = correspondances[0]
                try:
                    val = df.iat[idx_ligne, idx_ref] if axis == "columns" else df.iat[idx_ref, idx_col]
                    return values_lq_or_none(val)
                except Exception as e:
                    print(f"❌ Erreur fallback simple sur '{item}' : {e}")
                    return ""
            else:
                print(f"⚠️ '{item}' sans → ignoré car correspondances multiples ou absentes : {correspondances}")
                return ""


    def load_data(self):
        self.df = pd.read_excel(self.excel_path, sheet_name=self.sheet_name, header=None)

    def load_keywords_ui2(self):
        with open(self.json_config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        print("✔️ LOAD_KEYWORD_UI2 : Contenu JSON chargé :", data)

        self.keywords_valides = data.get("keywords_valides", [])
        self.groupes_personnalises = data.get("groupes_personnalises", {})
        self.ordre_colonnes = data.get(
            "ordre_selection",
            self.keywords_valides + list(self.groupes_personnalises.keys())
        )
        print("✅ LOAD_KEYWORD_UI2 : Groupes chargés :", self.groupes_personnalises)
        print("✅ LOAD_KEYWORD_UI2 : Ordre colonnes :", self.ordre_colonnes)

    def export(self, output_path="export_resultats.xlsx"):
        if not self.resultats:
            print("⚠️ Aucun résultat à exporter.")
            return

        df_export = pd.DataFrame.from_dict(self.resultats, orient="index")

        if "Code Artelia" in df_export.columns:
            df_export = df_export.drop(columns=["Code Artelia"])

        # Tri explicite selon l'ordre souhaité (zone droite)
        colonnes_finales = [col for col in self.ordre_colonnes if col in df_export.columns]
        df_export = df_export[colonnes_finales]

        df_export.index.name = "Code Artelia"
        df_export.to_excel(output_path)
        print(f"✅ Résultats exportés dans {output_path}")



# ======================================================================================= #
# ====================================== COLUMNS ======================================== #
# ======================================================================================= #

class ColumnsExtract(BaseExtract):
    def __init__(self, excel_path, json_config_path, sheet_name, col_config):
        super().__init__(excel_path, json_config_path, sheet_name, col_config)
        self.col_config = col_config

    # INPUT:
    #   self.df: Pandas DataFrame loaded from the Excel sheet
    #   self.keywords_valides: List of selected keywords to extract
    #   self.groupes_personnalises: Dict of custom groups (sums of keywords)
    #   self.col_config: Dict specifying key row and column indices (e.g. nom_row, nom_col, param_row)
    #   self.json_config_path: Path to JSON file containing input and output mapping zones
    # OUTPUT:
    #   self.resultats : values per sample, including group sums.
    def extract(self):
        self.resultats = {}
        df = self.df
        cfg = self.col_config

        nom_row = cfg["nom_row"]
        nom_col = cfg["nom_col"]
        param_row = cfg["param_row"]

        noms_colonnes = list(df.iloc[param_row])
        print("EXTRACT : Groupes chargés dans l'extract:", self.groupes_personnalises)

        # STEP 0 : Using [output_zone_droite] to recalculate based on [input_zone_gauche]
        #           (just (matched) to avoid all)
        all_correspondances = {}

        base_keywords = [
            kw.split("→")[0].strip()
            for kw in self.keywords_valides
            if "→ all" in kw
        ]

        for membres in self.groupes_personnalises.values():
            base_keywords.extend([
                m.split("→")[0].strip()
                for m in membres
                if "→ all" in m
            ])
        base_keywords = list(set(base_keywords))# Security

        # New detection
        noms_colonnes = list(df.iloc[self.col_config["param_row"]])
        matched, _ = self.get_matching_columns(noms_colonnes, base_keywords)
        print("\n🔍 MATCHED COLUMNS POUR '→ all' :")
        for kw, correspondances in matched.items():
            print(f"  {kw} → {[nom for _, nom in correspondances]}")

        # From matched
        for kw, correspondances in matched.items():
            all_correspondances[f"{kw} → all"] = [(col_idx, col) for col_idx, col in correspondances]

        # To code_artelia
        for idx_ligne in range(nom_row, len(df)):
            code_artelia = df.iat[idx_ligne, nom_col]
            if not isinstance(code_artelia, str) or not code_artelia.strip():
                continue

            self.resultats[code_artelia] = {}

            # STEP 1 : Groups to be processed independantly
            for nom_groupe, membres in self.groupes_personnalises.items():
                total = 0
                count = 0
                for membre in membres:
                    val = self.extract_values(
                        item=membre,
                        df=df,
                        idx=idx_ligne,
                        noms_reference=noms_colonnes,
                        correspondances_input=all_correspondances,
                        axis="columns"
                    )
                    self.resultats[code_artelia][membre] = val
                    try:
                        total += float(str(val).replace(",", "."))
                        count += 1
                    except:
                        continue
                self.resultats[code_artelia][nom_groupe] = (
                    values_lq_or_none(total) if count > 0 else ""
                )

            # STEP 2 : Simple match : keyword_valides part
            for kw in self.keywords_valides:
                if kw in self.groupes_personnalises:
                    continue
                val = self.extract_values(
                    item=kw,
                    df=df,
                    idx=idx_ligne,
                    noms_reference=noms_colonnes,
                    correspondances_input=all_correspondances,
                    axis="columns"
                )
                self.resultats[code_artelia][kw] = val



# ======================================================================================= #
# ======================================== ROWS ========================================= #
# ======================================================================================= #

class RowsExtract(BaseExtract):
    def __init__(self, excel_path, json_config_path, sheet_name, row_config):
        super().__init__(excel_path, json_config_path, sheet_name, row_config)
        self.row_config = row_config  # Exemple: {"col_nom_param": 1, "col_valeur": 2, "start_row": 8}

    # INPUT:
    #   self.df: Excel sheet loaded as a DataFrame.
    #   self.keywords_valides: List of parameters to extract.
    #   self.groupes_personnalises: Custom parameter groups (sum of several parameters).
    #   self.row_config: Dict defining key columns (e.g. col_nom_param, col_valeur).
    #   self.json_config_path: Path to the JSON config used for input/output mappings.
    #
    # OUTPUT:
    #   self.resultats: Dict mapping each sample code to its extracted parameters (and group values).
    def extract(self):
        self.resultats = {}
        df = self.df
        cfg = self.row_config

        nom_row = cfg["nom_row"]
        nom_col = cfg["nom_col"]
        param_col = cfg["param_col"]
        param_row = cfg["param_row"]
        data_start_col = cfg["data_start_col"]

        # STEP 1 : Extracting the "→ all" needed
        #
        all_correspondances = {}
        base_keywords = [
            kw.split("→")[0].strip()
            for kw in self.keywords_valides
            if "→ all" in kw
        ]

        for membres in self.groupes_personnalises.values():
            base_keywords.extend([
                m.split("→")[0].strip()
                for m in membres
                if "→ all" in m
            ])

        base_keywords = list(set(base_keywords))
        noms_parametres = df.iloc[param_row:, param_col].tolist() #all cells from param_col from nom_row
        matched, _ = self.get_matching_columns(noms_parametres, base_keywords)

        # all_correspondances output - {KW → all : [(idx, nom),(idx, nom),()]}
        for kw, correspondances in matched.items():
            all_correspondances[f"{kw} → all"] = [(idx, nom) for idx, nom in correspondances]


        # STEP 2 : Looking for code and values & creating list of results
        #
        for idx_col in range(data_start_col, df.shape[1]):
            code_artelia = df.iloc[nom_row, idx_col]
            if not isinstance(code_artelia, str) or not code_artelia.strip():
                continue
            self.resultats[code_artelia] = {}

            # STEP 2.1 : Looking for values for each kind in keywords_finals
            # GROUPS
            for nom_groupe, membres in self.groupes_personnalises.items():
                total = 0
                count = 0
                for membre in membres:
                    val_membre = self.extract_values(
                        item=membre,
                        df=df,
                        idx=idx_col,
                        noms_reference=noms_parametres,
                        correspondances_input=all_correspondances,
                        axis="rows",
                    )
                    self.resultats[code_artelia][membre] = val_membre
                    try:
                        total += float(str(val_membre).replace(",", "."))
                        count += 1
                    except:
                        continue
                self.resultats[code_artelia][nom_groupe] = (
                    values_lq_or_none(total) if count > 0 else ""
                )

            # SIMPLE
            for kw in self.keywords_valides:
                if kw in self.groupes_personnalises:
                    continue
                val_kw = self.extract_values(
                    item=kw,
                    df=df,
                    idx=idx_col,
                    noms_reference=noms_parametres,
                    correspondances_input=all_correspondances,
                    axis="rows",
                )
                self.resultats[code_artelia][kw] = val_kw












# ========================================== DEBUGGING TEST =========================================================
# ========================================== DEBUGGING TEST =========================================================
# ========================================== DEBUGGING TEST =========================================================
# ========================================== DEBUGGING TEST =========================================================
# ========================================== DEBUGGING TEST =========================================================
# ========================================== DEBUGGING TEST =========================================================
# ========================================== DEBUGGING TEST =========================================================
# ========================================== DEBUGGING TEST =========================================================
# ========================================== DEBUGGING TEST =========================================================
# ========================================== DEBUGGING TEST =========================================================
if __name__ == "__main__":
    file_path = os.path.join(os.path.dirname(__file__), "Résultats Eurofins.xlsm")
    if not os.path.exists(file_path):
        print("❌ Fichier 'Résultats Eurofins.xlsm' introuvable à côté du script.")
        exit()

    config = {
        "cell_nom_echantillon": "B6",
        "cell_parametres": "D5",
        "cell_data_start": "D6"
    }

    keywords_path = os.path.join(os.path.dirname(__file__), "keywords.json")
    keywords_valides = BaseExtract.load_keywords_ui1(keywords_path)

    df = pd.read_excel(file_path, sheet_name=0, header=None)
    row_start, col_start = cell_to_index(config["cell_parametres"])

    row = df.iloc[row_start]
    matched, multi = BaseExtract.get_matching_columns(row, keywords_valides)

    print("\n--- Résultat final ---")
    for kw, matches in matched.items():
        for col_idx, nom_col in matches:
            print(f"{kw}  →     {col_idx}     :     '{nom_col}'")

    if multi:
        print("\n🔁 Keywords with multiple matches (→ all):")
        for kw in multi:
            print(f"{kw} → all  →     '', 'all'")