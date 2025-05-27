import tkinter as tk
from tkinter import filedialog, messagebox
import json
import os
import sys
from collections import defaultdict
import random

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

DOSSIER_DATA = os.path.join(BASE_DIR, "00_Cache")
os.makedirs(DOSSIER_DATA, exist_ok=True)

FICHIER_LAST_CONFIG = os.path.join(DOSSIER_DATA, "last_config_extract.json")
temp_json = os.path.join(DOSSIER_DATA, "final_keywords.json")


def save_last_config(path):
    with open(FICHIER_LAST_CONFIG, "w") as f:
        json.dump({"last_config": path}, f)


def load_last_config():
    if os.path.exists(FICHIER_LAST_CONFIG):
        with open(FICHIER_LAST_CONFIG, "r") as f:
            return json.load(f).get("last_config", "")
    return ""


def ouvrir_ui_post_extract(matched_columns, extraction_type, excel_file, resultats_artelia, sheet_name, df, mapping_all, config_extraction=None):
    affichage_mapping = {}
    mapping_all = getattr(resultats_artelia, 'mapping_all', {})

    for kw, colonnes in matched_columns.items():
        for col in colonnes:
            label_affiche = f"{kw} → {col}"
            affichage_mapping[label_affiche] = kw

    def charger_config():
        path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                # NEW VERSION JSON : Release v1.1
                config_kw_dict = data.get("keywords_valides", {})
                mots = []
                for kw, correspondances in config_kw_dict.items():
                    for corr in correspondances:
                        mots.append(f"{kw} → {corr}")

                groupes.clear()
                groupes.update(data.get("groupes_personnalises", {}))

                config_path.set(path)
                save_last_config(path)

                zone_droite.delete(0, tk.END)
                zone_gauche.delete(0, tk.END)

                for label in libelles_formates:
                    ref = affichage_mapping.get(label, label)
                    if isinstance(ref, tuple):
                        ref_str = f"{ref[0]} → {ref[1]}"
                    else:
                        ref_str = ref
                    if ref_str in mots:
                        zone_droite.insert(tk.END, label)
                    else:
                        zone_gauche.insert(tk.END, label)

                for nom_groupe in groupes:
                    if nom_groupe not in zone_droite.get(0, tk.END):
                        zone_droite.insert(tk.END, nom_groupe)

                afficher_groupes()

        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de charger le fichier :\n{e}")

    def ajouter_mots():
        for i in zone_gauche.curselection()[::-1]:
            kw = zone_gauche.get(i)
            zone_droite.insert(tk.END, kw)
            zone_gauche.delete(i)

    def retirer_mots():
        for i in zone_droite.curselection()[::-1]:
            kw = zone_droite.get(i)
            zone_gauche.insert(tk.END, kw)
            zone_droite.delete(i)

    def generer_config():
        labels_selectionnes = list(zone_droite.get(0, tk.END))
        if not labels_selectionnes:
            messagebox.showwarning("Aucun mot-clé", "Veuillez sélectionner au moins un mot-clé.")
            return

        mots_dict = defaultdict(list)

        # SIMPLES VALUES
        for label in labels_selectionnes:
            if label in groupes:
                continue

            kw = affichage_mapping.get(label, label)

            if isinstance(kw, tuple):
                base, cible = kw
                mots_dict[base].append(cible)
            elif isinstance(kw, str):
                if "→" in kw:
                    base, cible = [x.strip() for x in kw.split("→", 1)]
                    mots_dict[base].append(cible)
                elif "→ all" in kw:
                    base = kw.split("→")[0].strip()
                    # NEW VERSION : Release v1.1 : Adding all matching param instead of just "all"
                    if base in matched_columns:
                        for _, vrai_nom in matched_columns[base]:
                            if vrai_nom not in mots_dict[base]:
                                mots_dict[base].append(vrai_nom)
                else:
                    mots_dict[kw].append(kw)

        # GROUPS VALUES
        for group_name, sous_params in groupes.items():
            for param in sous_params:
                if "→" in param:
                    base, suffix = [x.strip() for x in param.split("→", 1)]
                    if suffix == "all" and base in matched_columns:
                        if base not in mots_dict:
                            mots_dict[base] = ["all"]
                        elif "all" not in mots_dict[base]:
                            mots_dict[base].append("all")

        groupes_corriges = {}
        for nom_groupe, sous_params in groupes.items():
            nouveaux_sous_params = []
            for sp in sous_params:
                if "→" in sp:
                    base, cible = [x.strip() for x in sp.split("→", 1)]
                    if cible == "all":
                        # Remplacer par tous les noms réels
                        if base in matched_columns:
                            for _, vrai_nom in matched_columns[base]:
                                nouveaux_sous_params.append(f"{base} → {vrai_nom}")
                    else:
                        nouveaux_sous_params.append(sp)
                else:
                    nouveaux_sous_params.append(sp)
            groupes_corriges[nom_groupe] = nouveaux_sous_params


        # DEBUG : Afficher dans la console le contenu final de mots_dict
        print("=== DEBUG: mots_dict final ===")
        for k, v in mots_dict.items():
            print(f"{k}: {v}")
        print("=== DEBUG: matched_columns ===")
        for k, v in matched_columns.items():
            print(f"{k}: {v}")
        print("=== FIN DEBUG ===")

        # JSON Generation
        path = filedialog.asksaveasfilename(defaultextension=".json", initialfile="config_extract.json")
        if not path:
            return

        with open(path, "w", encoding="utf-8") as f:
            json.dump({
                "keywords_valides": dict(mots_dict),
                "groupes_personnalises": groupes_corriges
            }, f, indent=2, ensure_ascii=False)

        config_path.set(path)
        save_last_config(path)
        messagebox.showinfo("Succès", f"Configuration sauvegardée dans :\n{path}")


    def extraire_en_excel():
        labels_selectionnes = list(zone_droite.get(0, tk.END))
        mots_dict = {}

        for label in labels_selectionnes:
            kw = affichage_mapping.get(label, label)

            if isinstance(kw, tuple):
                base, cible = kw
            elif isinstance(kw, str) and "→" in kw:
                base, cible = [x.strip() for x in kw.split("→", 1)]
            else:
                base, cible = kw, kw

            if base not in mots_dict:
                mots_dict[base] = []
            if cible not in mots_dict[base]:
                mots_dict[base].append(cible)

        if not mots_dict:
            messagebox.showwarning("Aucun mot-clé", "Veuillez sélectionner au moins un mot-clé.")
            return

        config = None
        if config_extraction:
            try:
                from extract_utils import cell_to_index

                if extraction_type.lower() == "lignes":
                    r_nom, c_nom = cell_to_index(config_extraction["cell_nom_echantillon"])
                    r_data, c_data = cell_to_index(config_extraction["cell_data_start"])
                    r_param, c_param = cell_to_index(config_extraction["cell_parametres"])
                    r_limite = None
                    cell_limite = config_extraction.get("cell_limite")
                    if cell_limite and str(cell_limite).strip().lower() != "none":
                        r_limite, _ = cell_to_index(cell_limite)

                    optionnels_brut = config_extraction.get("optionnels", {})
                    optionnels = {
                        k: cell_to_index(v) for k, v in optionnels_brut.items()
                        if isinstance(v, str) and v.strip().lower() != "none"
                    }

                    config = {
                        "nom_row": r_nom,
                        "nom_col": c_nom,
                        "param_row": r_param,
                        "param_col": c_param,
                        "row_limites": r_limite,
                        "data_start_row": r_data,
                        "data_start_col": c_data,
                        "optionnels": optionnels,
                    }

                elif extraction_type.lower() == "colonnes":
                    r_param, c_param = cell_to_index(config_extraction["cell_parametres"])
                    r_nom, c_nom = cell_to_index(config_extraction["cell_nom_echantillon"])
                    r_data, c_data = cell_to_index(config_extraction["cell_data_start"])
                    r_limite, c_limite = (None, None)
                    cell_limite = config_extraction.get("cell_limite")
                    if cell_limite and str(cell_limite).strip().lower() != "none":
                        r_limite, c_limite = cell_to_index(cell_limite)

                    optionnels_brut = config_extraction.get("optionnels", {})
                    optionnels = {
                        k: cell_to_index(v) for k, v in optionnels_brut.items()
                        if isinstance(v, str) and v.strip().lower() != "none"
                    }

                    config = {
                        "param_row": r_param,
                        "param_col": c_param,
                        "nom_row": r_nom,
                        "nom_col": c_nom,
                        "data_start_row": r_data,
                        "data_start_col": c_data,
                        "limite_row": r_limite,
                        "limite_col": c_limite,
                        "optionnels": optionnels
                    }

            except Exception as e:
                messagebox.showerror("Erreur config_extraction", f"Erreur lors de la conversion de la config :\n{e}")
                return

        with open(temp_json, "w", encoding="utf-8") as f:
            json.dump({
                "keywords_valides": mots_dict,
                "groupes_personnalises": groupes
            }, f, indent=2, ensure_ascii=False)

        if extraction_type.lower() == "colonnes":
            from analysis_extract import ColumnsExtract
            extractor = ColumnsExtract(excel_file, temp_json, sheet_name, col_config=config)
        elif extraction_type.lower() == "lignes":
            from analysis_extract import RowsExtract
            extractor = RowsExtract(excel_file, temp_json, sheet_name, row_config=config)
        else:
            messagebox.showinfo("Non pris en charge", f"Type '{extraction_type}' non encore implémenté.")
            return

        extractor.load_keywords()
        extractor.load_data()
        extractor.extract()
        extractor.export()

    # GROUPING DATA SUM FUNCTION
    #
    def editer_groupe(nom=None):
        def valider():
            nom_groupe = entry_nom.get().strip()
            if not nom_groupe:
                messagebox.showwarning("Nom manquant", "Veuillez entrer un nom de groupe.")
                return

            selection = listbox.curselection()
            mots_selectionnes = []
            for i in selection:
                item = listbox.get(i)
                val = reverse_mapping.get(item)
                if isinstance(val, tuple):
                    mots_selectionnes.append(f"{val[0]} → {val[1]}")
                else:
                    mots_selectionnes.append(val)

            if not mots_selectionnes:
                messagebox.showwarning("Aucun mot-clé", "Sélectionnez au moins un mot-clé.")
                return

            groupes[nom_groupe] = mots_selectionnes
            fenetre.destroy()
            afficher_groupes()

        fenetre = tk.Toplevel()
        fenetre.title("Créer / Modifier un groupe")
        tk.Label(fenetre, text="Nom du groupe :").pack(pady=5)
        entry_nom = tk.Entry(fenetre)
        entry_nom.pack(pady=5)
        if nom:
            entry_nom.insert(0, nom)

        tk.Label(fenetre, text="Mots-clés disponibles :").pack()
        listbox = tk.Listbox(fenetre, selectmode=tk.MULTIPLE, width=40, height=10)
        listbox.pack(padx=10, pady=5)

        reverse_mapping = {}
        libelles_groupables = []

        # Basic copy from keyword matching generation
        for kw, correspondances in matched_columns.items():
            if not correspondances:
                label = kw
                libelles_groupables.append(label)
                reverse_mapping[label] = kw
            else:
                label_all = f"{kw} → all"
                libelles_groupables.append(label_all)
                reverse_mapping[label_all] = (kw, "all")
                for _, vrai_nom in correspondances:
                    label = f"{kw} → {vrai_nom}"
                    libelles_groupables.append(label)
                    reverse_mapping[label] = (kw, vrai_nom)

        # Listbox from copy
        for label in libelles_groupables:
            listbox.insert(tk.END, label)
        # Memory for reopening
        if nom and nom in groupes:
            mots_du_groupe = groupes[nom]
            for i, label in enumerate(listbox.get(0, tk.END)):
                val = reverse_mapping.get(label)
                val_str = f"{val[0]} → {val[1]}" if isinstance(val, tuple) else str(val)
                if val_str in mots_du_groupe:
                    listbox.selection_set(i)

        tk.Button(fenetre, text="Valider", command=valider).pack(pady=10)


    def supprimer_groupe(nom):
        if messagebox.askyesno("Confirmer suppression", f"Supprimer le groupe « {nom} » ?"):
            groupes.pop(nom, None)
            afficher_groupes()

    def ajouter_groupe_a_selection(nom_groupe):
        if nom_groupe in groupes:
            if nom_groupe not in zone_droite.get(0, tk.END):
                zone_droite.insert(tk.END, nom_groupe)
            if nom_groupe in zone_gauche.get(0, tk.END):
                idx = zone_gauche.get(0, tk.END).index(nom_groupe)
                zone_gauche.delete(idx)

    def afficher_groupes():
        # Efface tout dans la vraie zone liste
        for widget in frame_groupes_liste.winfo_children():
            widget.destroy()

        if not groupes:
            lbl = tk.Label(frame_groupes_liste, text="Aucun groupe défini", fg="gray")
            lbl.pack()
            return

        for nom in groupes:
            row = tk.Frame(frame_groupes_liste)
            row.pack(fill="x", pady=1)
            tk.Label(row, text=nom, width=25, anchor="w").pack(side=tk.LEFT)
            tk.Button(row, text="✏️", command=lambda n=nom: editer_groupe(n)).pack(side=tk.LEFT, padx=2)
            tk.Button(row, text="❌", command=lambda n=nom: supprimer_groupe(n)).pack(side=tk.LEFT, padx=2)
            tk.Button(row, text="➕", command=lambda n=nom: ajouter_groupe_a_selection(n)).pack(side=tk.LEFT, padx=10)


    # = MAIN WINDOW
    #
    fenetre = tk.Toplevel()

    fenetre.title("Sélection des paramètres à extraire")
    fenetre.geometry("550x750")

    config_path = tk.StringVar(value="config_extract.json")
    groupes = {}

    tk.Label(fenetre, text="CONFIGURATION :", font=("Segoe UI", 10, "bold")).pack(pady=2)
    frame_conf = tk.Frame(fenetre)
    frame_conf.pack()
    tk.Label(frame_conf, textvariable=config_path, fg="blue").pack(side=tk.LEFT)
    tk.Button(frame_conf, text="SÉLECTIONNER", command=charger_config).pack(side=tk.LEFT, padx=5)

    # PARAMETERS AREA TO SELECT FROM LISTS
    #
    frame_zones = tk.Frame(fenetre)
    frame_zones.pack(pady=15, fill="both", expand=True)

    tk.Label(frame_zones, text="PARAMÈTRES DISPONIBLES", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, padx=10,
                                                                                            pady=(0, 5))
    tk.Label(frame_zones, text="PARAMÈTRES À EXPORTER", font=("Segoe UI", 9, "bold")).grid(row=0, column=2, padx=10,
                                                                                           pady=(0, 5))
    zone_gauche = tk.Listbox(frame_zones, selectmode=tk.EXTENDED, width=40, height=15)
    zone_droite = tk.Listbox(frame_zones, selectmode=tk.EXTENDED, width=40, height=15)
    zone_gauche.grid(row=1, column=0, padx=10)
    zone_droite.grid(row=1, column=2, padx=10)

    # DRAG AND DROP ON RIGHT AREA TO BE ABLE TO MODIFY THE JSON DIRECTLY INTO THE UI
    #
    def on_start_drag(event):
        widget = event.widget
        widget.drag_start_index = widget.nearest(event.y)

    def on_drag_motion(event):
        widget = event.widget
        current_index = widget.nearest(event.y)
        start_index = getattr(widget, "drag_start_index", None)

        if start_index is not None and current_index != start_index:
            items = list(widget.get(0, tk.END))
            item = items.pop(start_index)
            items.insert(current_index, item)

            widget.delete(0, tk.END)
            for i in items:
                widget.insert(tk.END, i)

            widget.drag_start_index = current_index

    # Bind drag & drop à la zone de droite
    zone_droite.bind("<Button-1>", on_start_drag)
    zone_droite.bind("<B1-Motion>", on_drag_motion)

    frame_btn = tk.Frame(frame_zones)
    frame_btn.grid(row=1, column=1)
    tk.Button(frame_btn, text="→", command=ajouter_mots).pack(pady=10)
    tk.Button(frame_btn, text="←", command=retirer_mots).pack(pady=10)

    # RANDOM VERIFICATION AREA
    #
    frame_verification = tk.LabelFrame(fenetre, text="🔎 Vérification d'une ligne extraite",
                                       font=("Segoe UI", 9, "bold"))
    frame_verification.pack(pady=10, fill="x", padx=15)

    text_resultat = tk.Text(frame_verification, height=2, width=80, font=("Courier", 9), wrap="word", bg="#f0f0f0",
                            state="disabled")
    text_resultat.pack(pady=5, padx=10)


    def afficher_ligne_random():
        try:
            if not resultats_artelia:
                raise ValueError("Aucune donnée disponible.")

            code = random.choice(list(resultats_artelia.keys()))
            mesures = resultats_artelia[code]

            texte = f"Code Artelia: {code}"
            if not mesures:
                texte += " | Pas d'analyse détectée"
            else:
                comp, val = random.choice(list(mesures.items()))
                texte += f" | {comp} = {val}"

            text_resultat.configure(state="normal")
            text_resultat.delete("1.0", tk.END)
            text_resultat.insert(tk.END, texte)
            text_resultat.configure(state="disabled")

        except Exception as e:
            text_resultat.configure(state="normal")
            text_resultat.delete("1.0", tk.END)
            text_resultat.insert(tk.END, f"⚠️ Erreur : {e}")
            text_resultat.configure(state="disabled")

    tk.Button(frame_verification, text="Randomize", command=afficher_ligne_random).pack(pady=5)


    # GROUP AREA TO GATHER PARAMETERS TO PREPARE SUM
    #
    frame_groupes = tk.LabelFrame(fenetre, text="🔄 GROUPES DE SOMME PERSONNALISÉS", font=("Segoe UI", 9, "bold"))
    frame_groupes.pack(pady=10, fill="x", padx=15)

    frame_groupes_liste = tk.Frame(frame_groupes)
    frame_groupes_liste.pack(anchor="w", padx=10, pady=5)

    tk.Button(frame_groupes, text="+ Créer un groupe", command=lambda: editer_groupe()).pack(pady=5)

    frame_bottom = tk.Frame(fenetre)
    frame_bottom.pack(pady=15)
    tk.Button(frame_bottom, text="💾 GÉNÉRER JSON", width=20, command=generer_config).pack(side=tk.LEFT, padx=20)
    tk.Button(frame_bottom, text="📤 EXTRAIRE EN EXCEL", width=20, command=extraire_en_excel).pack(side=tk.LEFT, padx=20)

    tk.Label(fenetre, text="© Paul Ancian – 2025", font=("Segoe UI", 7), fg="gray") \
        .pack(side="bottom", pady=(5, 10))


    # == INIT ==
    affichage_mapping = {}
    libelles_formates = []

    for kw, correspondances in matched_columns.items():
        if not correspondances:
            label = kw
            libelles_formates.append(label)
            affichage_mapping[label] = kw
        else:
            label_all = f"{kw} → all"
            libelles_formates.append(label_all)
            affichage_mapping[label_all] = (kw, "all")

            for idx, vrai_nom in correspondances:
                label = f"{kw} → {vrai_nom}"
                libelles_formates.append(label)
                affichage_mapping[label] = (kw, vrai_nom)

    # JSON CONFIG : From old session or from importing JSON
    config_kw = []
    last_used = load_last_config()
    if last_used and os.path.exists(last_used):
        try:
            with open(last_used, "r", encoding="utf-8") as f:
                data = json.load(f)
                config_path.set(last_used)
                config_kw_dict = data.get("keywords_valides", {})
                config_kw = []

                for kw, correspondances in config_kw_dict.items():
                    if isinstance(correspondances, list):
                        # Keep multiple instance of the same cells for cases like multiple same analysis
                        for corr in correspondances:
                            config_kw.append(f"{kw} → {corr}")

                groupes.update(data.get("groupes_personnalises", {}))
        except:
            config_kw = []

    zone_droite.delete(0, tk.END)
    zone_gauche.delete(0, tk.END)

    for label in libelles_formates:
        ref = affichage_mapping.get(label, label)
        if isinstance(ref, tuple):
            ref = f"{ref[0]} → {ref[1]}"
        if ref in config_kw:
            zone_droite.insert(tk.END, label)
        else:
            zone_gauche.insert(tk.END, label)

    for nom_groupe in groupes:
        if nom_groupe not in zone_droite.get(0, tk.END):
            zone_droite.insert(tk.END, nom_groupe)

    afficher_groupes()