import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import sys
import os
import warnings
import pandas as pd
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

from analysis_extract import ColumnsExtract, RowsExtract
from ui_post_extract import ouvrir_ui_post_extract
from extract_utils import cell_to_index, normalize

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DOSSIER_DATA = os.path.join(BASE_DIR, "00_Cache")
os.makedirs(DOSSIER_DATA, exist_ok=True)

# UI 1
FICHIER_SESSION = os.path.join(DOSSIER_DATA, "last_session.json")
FICHIER_TEMP_KEYWORDS = os.path.join(DOSSIER_DATA, "temp_keywords.json")
TYPES_EXTRACTION = ["Colonnes", "Lignes"]

# UI 2
FICHIER_LAST_CONFIG = os.path.join(DOSSIER_DATA, "last_config_extract.json")
FICHIER_LAST_TYPE_CONFIG = os.path.join(DOSSIER_DATA, "last_type_config.json")


# === Script : MAIN UI - TABLEURS MULTIPLES ===
# = v1 : Import Excel and JSON for keywords + able to modify keywords
    # = v1.05 : Modify type of extraction possible for future
    # = v1.2 : Adding second UI to verify keywords extraction based on Excel columns letter and adding it to extract
    # = v1.25 : Adding JSON config_extract file - memory for extraction will
    # = v1.3 : Adding grouping component to export SUM
    # = v1.4 : Adding randomizing area to manually verify data before Excel extraction
    # = v1.45 : Debug and modifying output location
# = v2 : PASSAGE FORMAT CLASSES VIA ANALYSIS_EXTRACT.py
    # = v2.1 : Clean function with extract_utils.py and ui_post_extract.py
    # = v2.2 : Letting user choose between Columns or Rows data, adding UI for parameters
    # = v2.3 : Debug on Excel exporting only 1 column per keyword even if multiple were selected
#

class ExtractApp:
    def __init__(self, master):
        self.master = master
        master.title("Analyses géochimiques")
        master.geometry("300x450")
        master.resizable(False, False)

        self.keyword_file = ""
        self.excel_file = ""
        self.keywords = []
        self.row_config = None
        self.col_config = None

        self.setup_ui()
        self.charger_derniere_session()

    def setup_ui(self):
        main_frame = tk.Frame(self.master)
        main_frame.pack(expand=True)

        # KEYWORDS
        tk.Label(main_frame, text="MOTS-CLÉS", font=("Segoe UI", 10, "bold")).pack()
        separator = tk.Frame(main_frame, height=0.5, bd=0, relief='sunken', bg='gray')
        separator.pack(fill="x", padx=20, pady=5)
        self.label_keywords = tk.Label(main_frame, text="Aucun fichier chargé", fg="red")
        self.label_keywords.pack()
        frame_kw = tk.Frame(main_frame)
        frame_kw.pack(pady=5)
        tk.Button(frame_kw, text="Sélectionner", command=self.charger_keywords).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_kw, text="Modifier", command=self.ouvrir_editeur_keywords).pack(side=tk.LEFT, padx=5)
        separator = tk.Frame(main_frame, height=0.5, bd=0, relief='sunken', bg='gray')
        separator.pack(fill="x", padx=20, pady=5)

        # EXCEL
        tk.Label(main_frame, text="FICHIER EXCEL", font=("Segoe UI", 10, "bold")).pack(pady=(15, 0))
        separator = tk.Frame(main_frame, height=0.5, bd=0, relief='sunken', bg='gray')
        separator.pack(fill="x", padx=20, pady=5)
        self.label_excel = tk.Label(main_frame, text="Aucun fichier chargé", fg="red")
        self.label_excel.pack()
        tk.Label(main_frame, text="Feuille Excel", font=("Segoe UI", 9, "bold")).pack()
        self.sheet_var = tk.StringVar()
        self.menu_sheets = ttk.Combobox(main_frame, textvariable=self.sheet_var, state="disabled", width=30)
        self.menu_sheets.pack(pady=(0, 10))
        tk.Button(main_frame, text="Sélectionner", command=self.choisir_fichier_excel).pack(pady=5)

        # TYPE D'EXTRACTION
        separator = tk.Frame(main_frame, height=0.5, bd=0, relief='sunken', bg='gray')
        separator.pack(fill="x", padx=20, pady=10)
        tk.Label(main_frame, text="TYPE D'EXTRACTION", font=("Segoe UI", 9, "bold")).pack(pady=(15, 5))
        frame_type = tk.Frame(main_frame)
        frame_type.pack(pady=5)

        self.type_var = tk.StringVar(value="Colonnes")
        self.menu_type = ttk.Combobox(frame_type, textvariable=self.type_var, values=TYPES_EXTRACTION, state="readonly",
                                      width=20)
        self.menu_type.pack(side=tk.LEFT, padx=(0, 5))

        btn_configurer_type = tk.Button(frame_type, text="Configurer", command=self.ouvrir_popup_type)
        btn_configurer_type.pack(side=tk.LEFT)

        tk.Button(frame_type, text="💾", command=self.sauver_config_type).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_type, text="📂", command=self.charger_config_type).pack(side=tk.LEFT, padx=5)

        # BOUTON EXTRACT
        tk.Button(self.master, text="EXTRACT", bg="green", fg="white", font=("Segoe UI", 10, "bold"), command=self.lancer_extraction).pack(pady=10)

        tk.Label(self.master, text="© Paul Ancian – 2025", font=("Segoe UI", 7), fg="gray") \
            .pack(side="bottom", pady=(5, 10))

    def charger_keywords(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if not file_path:
            return
        self.keyword_file = file_path
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)

                if isinstance(data, list):
                    self.keywords = data
                elif isinstance(data, dict) and "keywords_valides" in data:
                    self.keywords = data["keywords_valides"]
                else:
                    raise ValueError("Format de fichier JSON non reconnu.")

                self.label_keywords.config(text=os.path.basename(file_path), fg="black")
                self.sauvegarder_session()
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur chargement JSON :\n{e}")

    def ouvrir_editeur_keywords(self):
        if not self.keyword_file or not os.path.exists(self.keyword_file):
            messagebox.showwarning("Erreur", "Aucun fichier JSON chargé.")
            return

        editeur = tk.Toplevel(self.master)
        editeur.title("Modifier les mots-clés")
        editeur.geometry("400x300")

        listbox = tk.Listbox(editeur)
        listbox.pack(expand=True, fill=tk.BOTH, padx=10, pady=10)

        for kw in self.keywords:
            listbox.insert(tk.END, kw)

        def ajouter():
            mot = simple_input("Ajouter un mot-clé")
            if mot:
                listbox.insert(tk.END, mot)

        def supprimer():
            selection = listbox.curselection()
            if selection:
                listbox.delete(selection[0])

        def valider():
            self.keywords = list(listbox.get(0, tk.END))
            try:
                with open(self.keyword_file, "w", encoding="utf-8") as f:
                    json.dump(self.keywords, f, indent=2, ensure_ascii=False)
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur sauvegarde : {e}")
            editeur.destroy()

        frame_btn = tk.Frame(editeur)
        frame_btn.pack(pady=5)
        tk.Button(frame_btn, text="+", command=ajouter).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_btn, text="-", command=supprimer).pack(side=tk.LEFT, padx=5)
        tk.Button(frame_btn, text="Valider", command=valider).pack(side=tk.LEFT, padx=5)

    def choisir_fichier_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xls *.xlsx *.xlsm")])
        if file_path:
            self.excel_file = file_path
            self.label_excel.config(text=os.path.basename(file_path), fg="black")

            try:
                import pandas as pd
                sheets = pd.ExcelFile(file_path).sheet_names
                self.menu_sheets["values"] = sheets
                self.sheet_var.set(sheets[0])  # valeur par défaut = première feuille
                self.menu_sheets.config(state="readonly")
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de lire les feuilles :\n{e}")



    # NEW v2.2 = STEP ASKING USER FOR CELLS DATA - First cell with value, not headers
    #
    def ouvrir_popup_type(self):
        extraction_type = self.type_var.get()
        if extraction_type == "Lignes":
            def continuer_extraction(row_config):
                self.row_config = row_config
                with open(FICHIER_LAST_TYPE_CONFIG, "w", encoding="utf-8") as f:
                    json.dump(self.row_config, f, indent=2, ensure_ascii=False)

            self.create_type_extract_popup(self.master, continuer_extraction, config_init=self.row_config)

        elif extraction_type == "Colonnes":
            def continuer_extraction(col_config):
                self.col_config = col_config
                with open(FICHIER_LAST_TYPE_CONFIG, "w", encoding="utf-8") as f:
                    json.dump(self.col_config, f, indent=2, ensure_ascii=False)

            self.create_type_extract_popup(self.master, continuer_extraction, config_init=self.col_config)

        else:
            messagebox.showinfo("Info", f"Aucune configuration requise pour le type : {extraction_type}")



    def create_type_extract_popup(self, parent, callback, config_init=None):
        popup = tk.Toplevel(parent)
        popup.title("Configuration : Type")
        popup.geometry("240x300")
        popup.grab_set()

        entries = {}
        valeurs_par_defaut = config_init or {}

        def create_labeled_entry(frame, label_text, key, row):
            var = tk.StringVar()
            var.set(valeurs_par_defaut.get(key, ""))
            tk.Label(frame, text=label_text, anchor="w").grid(row=row, column=0, sticky="w", padx=(10, 5), pady=5)
            entry = tk.Entry(frame, textvariable=var, width=8)
            entry.grid(row=row, column=1, sticky="w", pady=5)
            entries[key] = var

        frame_main = tk.Frame(popup)
        frame_main.columnconfigure(1, weight=1)
        frame_main.pack(fill="both", expand=True, padx=10, pady=10)

        # Principals
        create_labeled_entry(frame_main, "ID Artelia :", "cell_nom_echantillon", row=0)
        create_labeled_entry(frame_main, "Noms Paramètres :", "cell_parametres", row=1)
        create_labeled_entry(frame_main, "Valeurs :", "cell_data_start", row=2)


        tk.Label(frame_main, text="Exemple : G8", fg="gray").grid(
            row=3, column=0, columnspan=2, padx=10, pady=(0, 5)
        )

        # Optionals, user choose name and cells
        def add_optional_field():
            popup_opt = tk.Toplevel(popup)
            popup_opt.title("Ajouter un champ optionnel")
            popup_opt.geometry("200x170")
            popup_opt.grab_set()

            var_name = tk.StringVar()
            var_cell = tk.StringVar()

            tk.Label(popup_opt, text="Nom du champ (ex: code_labo)").pack(pady=5)
            tk.Entry(popup_opt, textvariable=var_name, width=15).pack(pady=5)

            tk.Label(popup_opt, text="Cellule (ex: F9)").pack(pady=5)
            tk.Entry(popup_opt, textvariable=var_cell, width=8).pack(pady=5)


            def ajouter():
                name = var_name.get().strip()
                cell = var_cell.get().strip()
                if name and cell:
                    optional_fields[name] = cell
                    tk.Label(frame_optional, text=f"{name} : {cell}").pack(anchor="w")
                    popup_opt.destroy()
                else:
                    messagebox.showerror("Erreur", "Veuillez remplir les deux champs.")

            tk.Button(popup_opt, text="Ajouter", command=ajouter).pack(pady=10)

        tk.Button(frame_main, text="Ajouter champ optionnel", command=add_optional_field) \
            .grid(row=4, column=0, columnspan=2, pady=5)


        optional_fields = {}
        frame_optional = tk.LabelFrame(frame_main, text="Champs optionnels", padx=10, pady=10)
        frame_optional.grid(row=5, column=0, columnspan=2, padx=10, pady=5, sticky="w")

        for nom, cell in valeurs_par_defaut.get("optionnels", {}).items():
            optional_fields[nom] = cell
            tk.Label(frame_optional, text=f"{nom} : {cell}").pack(anchor="w")

        def valider():
            config = {k: v.get().strip() for k, v in entries.items()}
            config["optionnels"] = optional_fields
            if config.get("cell_limite", "") == "":
                config["cell_limite"] = None
            popup.destroy()
            callback(config)

        tk.Button(frame_main, text="Valider", command=valider, bg="green", fg="white") \
            .grid(row=8, column=0, columnspan=2, pady=10)


    def get_current_config(self):
        extraction_type = self.type_var.get().lower()
        if extraction_type == "colonnes":
            return self.col_config
        elif extraction_type == "lignes":
            return self.row_config
        return None

    def sauver_config_type(self):
        config = self.get_current_config()
        if not config:
            messagebox.showwarning("Aucune config", "Aucune configuration à sauvegarder.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".json",
                                            initialfile=f"config_type_{self.type_var.get().lower()}.json")
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            messagebox.showinfo("Succès", f"Configuration sauvegardée dans :\n{path}")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de sauvegarder :\n{e}")

    def charger_config_type(self):
        path = filedialog.askopenfilename(filetypes=[("Fichier JSON", "*.json")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                config = json.load(f)
            if self.type_var.get().lower() == "lignes":
                if not isinstance(config, dict):
                    messagebox.showerror("Erreur", "La configuration pour 'Lignes' doit être un dictionnaire JSON.")
                    return
                self.row_config = config
            else:
                if not isinstance(config, dict):
                    messagebox.showerror("Erreur", "La configuration pour 'Colonnes' doit être un dictionnaire JSON.")
                    return
                self.col_config = config
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de charger la configuration :\n{e}")



    def lancer_extraction(self):
        def detecter_correspondances(df_columns, keywords):
            correspondances = {}
            for kw in keywords:
                correspondances[kw] = []
                for col in df_columns:
                    col_norm = normalize(col)
                    kw_norm = normalize(kw)
                    if kw_norm in col_norm:
                        nom_affichage = f"{kw} → {col.strip()}"
                        correspondances[kw].append(nom_affichage)
            return correspondances

        if not self.excel_file:
            messagebox.showwarning("Fichier Excel manquant", "Veuillez sélectionner un fichier Excel.")
            return
        if not self.keywords:
            messagebox.showwarning("Mots-clés manquants", "Aucun mot-clé chargé.")
            return

        extraction_type = self.type_var.get()
        sheet_name = self.sheet_var.get()
        if not self.sheet_var.get() or self.sheet_var.get() not in self.menu_sheets["values"]:
            messagebox.showerror("Erreur", "Veuillez sélectionner une feuille Excel valide.")
            return

        config_extraction = self.get_current_config()
        if not isinstance(config_extraction, dict):
            messagebox.showerror("Erreur", "La configuration d'extraction est invalide (non dict).")
            return

        if not config_extraction:
            messagebox.showwarning("Configuration manquante",
                                   "Veuillez d'abord configurer l'extraction via le bouton 'Configurer'.")
            return

        df_temp = pd.read_excel(self.excel_file, sheet_name=sheet_name, nrows=10, header=None)

        if extraction_type.lower() == "colonnes":
            ligne_param = config_extraction.get("param_row", 0)
            colonne_param = config_extraction.get("param_col", 0)
            colonnes_detectees = df_temp.iloc[ligne_param, colonne_param:].astype(str).tolist()
        else:
            ligne_param = config_extraction.get("param_row", 0)
            colonne_param = config_extraction.get("param_col", 0)
            colonnes_detectees = df_temp.iloc[ligne_param:, colonne_param].astype(str).tolist()

        correspondances = detecter_correspondances(colonnes_detectees, self.keywords)
        groupes = {}

        with open(FICHIER_TEMP_KEYWORDS, "w", encoding="utf-8") as f:
            json.dump({
                "keywords_valides": correspondances,
                "groupes_personnalises": groupes
            }, f, indent=2, ensure_ascii=False)

        try:
            cell_param = config_extraction.get("cell_parametres", "A1")
            cell_nom = config_extraction.get("cell_nom_echantillon", "A1")
            cell_data_start = config_extraction.get("cell_data_start", "A1")

            r_param, c_param = cell_to_index(cell_param)
            r_nom, c_nom = cell_to_index(cell_nom)
            r_data, c_data = cell_to_index(cell_data_start)

            optionnels_brut = config_extraction.get("optionnels", {})
            optionnels = {}
            for k, v in optionnels_brut.items():
                if isinstance(v, str) and v.strip() and v.strip().lower() != "none":
                    try:
                        optionnels[k] = cell_to_index(v)
                    except Exception as e:
                        print(f"⚠️ Optionnel ignoré ({k}): valeur invalide '{v}' ({e})")

            config = {
                "param_row": r_param,
                "param_col": c_param,
                "nom_row": r_nom,
                "nom_col": c_nom,
                "data_start_row": r_data,
                "data_start_col": c_data,
                "optionnels": optionnels
            }

            if extraction_type.lower() == "colonnes":
                extractor = ColumnsExtract(self.excel_file, FICHIER_TEMP_KEYWORDS, sheet_name, col_config=config)
            elif extraction_type.lower() == "lignes":
                extractor = RowsExtract(self.excel_file, FICHIER_TEMP_KEYWORDS, sheet_name, row_config=config)
            else:
                messagebox.showerror("Erreur", f"Type d'extraction '{extraction_type}' non supporté.")
                return

            extractor.load_keywords()
            extractor.load_data()
            extractor.extract()

            print("Résultat extrait:", extractor.resultats_artelia)

            if extractor.df is None or extractor.df.empty:
                messagebox.showwarning("Extraction vide", "Aucune donnée extraite depuis le fichier.")
                return

            ouvrir_ui_post_extract(
                extractor.get_matched_columns(),
                extraction_type,
                self.excel_file,
                extractor.resultats_artelia,
                sheet_name,
                extractor.df,
                mapping_all=extractor.mapping_all,
                config_extraction=config_extraction
            )

        except Exception as e:
            messagebox.showerror("Erreur durant l'extraction", str(e))



    def sauvegarder_session(self):
        try:
            with open(FICHIER_SESSION, "w", encoding="utf-8") as f:
                json.dump({"keyword_file": self.keyword_file}, f)
        except Exception as e:
            print(f"⚠️ Erreur sauvegarde session : {e}")

    def charger_derniere_session(self):
        if not os.path.exists(FICHIER_SESSION):
            return
        try:
            with open(FICHIER_SESSION, "r", encoding="utf-8") as f:
                data = json.load(f)
                if "keyword_file" in data and os.path.exists(data["keyword_file"]):
                    self.keyword_file = data["keyword_file"]
                    with open(self.keyword_file, "r", encoding="utf-8") as kf:
                        self.keywords = json.load(kf)
                    self.label_keywords.config(text=os.path.basename(self.keyword_file), fg="black")
        except Exception as e:
            print(f"⚠️ Erreur chargement session : {e}")


        if os.path.exists(FICHIER_LAST_TYPE_CONFIG):
            try:
                with open(FICHIER_LAST_TYPE_CONFIG, "r", encoding="utf-8") as f:
                    config = json.load(f)

                if self.type_var.get().lower() == "lignes":
                    self.row_config = config
                    print("🔁 Configuration lignes rechargée :", self.row_config)
                else:
                    self.col_config = config
                    print("🔁 Configuration colonnes rechargée :", self.col_config)
            except Exception as e:
                print(f"⚠️ Erreur chargement dernière config extraction : {e}")




def simple_input(title):
    popup = tk.Toplevel()
    popup.title(title)
    popup.geometry("300x100")
    var = tk.StringVar()
    tk.Entry(popup, textvariable=var).pack(pady=10)
    tk.Button(popup, text="OK", command=popup.destroy).pack()
    popup.grab_set()
    popup.wait_window()
    return var.get().strip()


if __name__ == "__main__":
    root = tk.Tk()
    app = ExtractApp(root)
    root.mainloop()

