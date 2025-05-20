import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import os
import subprocess

# UI 1
FICHIER_SESSION = "last_session.json"
FICHIER_TEMP_KEYWORDS = "temp_keywords.json"
TYPES_EXTRACTION = ["Eurofins", "TEST"]

# UI 2
FICHIER_LAST_CONFIG = "last_config_extract.json"


# === Script : MAIN UI - TABLEURS EUROFINS ===
# = v1 : Import Excel and JSON for keywords + able to modify keywords
# = v1.5 : Modify type of extraction possible for future
# = v2 : Adding second UI to verify keywords extraction based on Excel columns letter and adding it to extract
# = v2.5 : Adding JSON config_extract file - memory for extraction will
# = v3 : Adding grouping component to export SUM
# = v4 : Adding randomizing area to manually verify data before Excel extraction
# = v4.5 : Debug and modifying output location
#
class ExtractApp:
    def __init__(self, master):
        self.master = master
        master.title("Analyses géochimiques")
        master.geometry("300x400")
        master.resizable(False, False)

        self.keyword_file = ""
        self.excel_file = ""
        self.keywords = []

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
        tk.Button(main_frame, text="Sélectionner", command=self.choisir_fichier_excel).pack(pady=5)

        # TYPE D'EXTRACTION
        separator = tk.Frame(main_frame, height=0.5, bd=0, relief='sunken', bg='gray')
        separator.pack(fill="x", padx=20, pady=10)
        tk.Label(main_frame, text="TYPE D'EXTRACTION", font=("Segoe UI", 9, "bold")).pack(pady=(15, 5))
        frame_type = tk.Frame(main_frame)
        frame_type.pack()
        self.type_var = tk.StringVar(value="Eurofins")
        self.menu_type = ttk.Combobox(frame_type, textvariable=self.type_var, values=TYPES_EXTRACTION, state="readonly", width=20)
        self.menu_type.pack()

        # BOUTON EXTRACT
        tk.Button(self.master, text="EXTRACT", bg="green", fg="white", font=("Segoe UI", 10, "bold"), command=self.lancer_extraction).pack(pady=10)

    def charger_keywords(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if not file_path:
            return
        self.keyword_file = file_path
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                self.keywords = json.load(f)
                if not isinstance(self.keywords, list):
                    raise ValueError("Le fichier JSON doit contenir une liste.")
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

    def lancer_extraction(self):
        if not self.excel_file:
            messagebox.showwarning("Fichier Excel manquant", "Veuillez sélectionner un fichier Excel.")
            return

        if not self.keywords:
            messagebox.showwarning("Mots-clés manquants", "Aucun mot-clé chargé.")
            return

        with open(FICHIER_TEMP_KEYWORDS, "w", encoding="utf-8") as f:
            json.dump(self.keywords, f, indent=2, ensure_ascii=False)

        extraction_type = self.type_var.get()
        if extraction_type == "Eurofins":
            subprocess.run(["python", "Eurofins_extract.py", self.excel_file, FICHIER_TEMP_KEYWORDS])

            # Charger les colonnes détectées depuis un fichier temporaire produit par Eurofins_extract.py
            if os.path.exists("matched_columns.json"):
                with open("matched_columns.json", "r", encoding="utf-8") as f:
                    matched_columns = json.load(f)
                ouvrir_ui_post_extract(matched_columns, extraction_type, self.excel_file)
            else:
                messagebox.showwarning("Colonnes introuvables", "Le fichier matched_columns.json n'a pas été trouvé.")
        else:
            messagebox.showinfo("Non implémenté", f"Le type '{extraction_type}' n'est pas encore pris en charge.")

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





# ======================================= UI 2 ==========================================
#
def ouvrir_ui_post_extract(matched_columns: dict, extraction_type: str, excel_file: str):
    def save_last_config(path):
        with open(FICHIER_LAST_CONFIG, "w") as f:
            json.dump({"last_config": path}, f)

    def load_last_config():
        if os.path.exists(FICHIER_LAST_CONFIG):
            with open(FICHIER_LAST_CONFIG, "r") as f:
                return json.load(f).get("last_config", "")
        return ""

    def charger_config():
        path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                mots = data.get("keywords_valides", [])
                groupes.clear()
                groupes.update(data.get("groupes_personnalises", {}))

                zone_droite.delete(0, tk.END)
                zone_gauche.delete(0, tk.END)
                for kw in mots:
                    zone_droite.insert(tk.END, kw)
                for kw in tous_keywords:
                    if kw not in mots:
                        zone_gauche.insert(tk.END, kw)

                config_path.set(path)
                save_last_config(path)
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
        mots = list(zone_droite.get(0, tk.END))
        if not mots:
            messagebox.showwarning("Aucun mot-clé", "Veuillez sélectionner au moins un mot-clé.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".json", initialfile="config_extract.json")
        if not path:
            return
        with open(path, "w", encoding="utf-8") as f:
            json.dump({
                "keywords_valides": mots,
                "groupes_personnalises": groupes
            }, f, indent=2, ensure_ascii=False)
        config_path.set(path)
        save_last_config(path)
        messagebox.showinfo("Succès", f"Configuration sauvegardée dans :\n{path}")

    def extraire_en_excel():
        mots = list(zone_droite.get(0, tk.END))
        if not mots:
            messagebox.showwarning("Aucun mot-clé", "Veuillez sélectionner au moins un mot-clé.")
            return
        temp_json = "final_keywords.json"
        with open(temp_json, "w", encoding="utf-8") as f:
            json.dump({
                "keywords_valides": mots,
                "groupes_personnalises": groupes
            }, f, indent=2, ensure_ascii=False)
        if extraction_type.lower() == "eurofins":
            subprocess.run(["python", "Eurofins_extract.py", excel_file, temp_json, "--export"])
        else:
            messagebox.showinfo("Non pris en charge", f"Type '{extraction_type}' non encore implémenté.")



    # GROUPING DATA SUM FUNCTION
    #
    def editer_groupe(nom=None):
        def valider():
            nom_groupe = entry_nom.get().strip()
            if not nom_groupe:
                messagebox.showwarning("Nom manquant", "Veuillez entrer un nom de groupe.")
                return
            selection = listbox.curselection()
            mots = [zone_gauche.get(i) for i in selection]
            if not mots:
                messagebox.showwarning("Aucun mot-clé", "Sélectionnez au moins un mot-clé.")
                return
            groupes[nom_groupe] = mots
            fenetre.destroy()
            afficher_groupes()

        fenetre = tk.Toplevel()
        fenetre.title("Créer / Modifier un groupe")
        tk.Label(fenetre, text="Nom du groupe :").pack(pady=5)
        entry_nom = tk.Entry(fenetre)
        entry_nom.pack(pady=5)
        if nom:
            entry_nom.insert(0, nom)

        tk.Label(fenetre, text="Mots-clés à inclure (depuis la zone de gauche) :").pack()
        listbox = tk.Listbox(fenetre, selectmode=tk.MULTIPLE, width=40, height=10)
        listbox.pack(padx=10, pady=5)

        # Using UI1 list to gather data
        keywords_source = list(zone_gauche.get(0, tk.END)) + list(zone_droite.get(0, tk.END))
        unique_keywords = []
        seen = set()

        for kw in keywords_source:
            if kw not in seen:
                seen.add(kw)
                unique_keywords.append(kw)
                listbox.insert(tk.END, kw)

        # Sélection uniquement sur ces valeurs
        if nom and nom in groupes:
            mots_du_groupe = groupes[nom]
            for i, kw in enumerate(unique_keywords):
                if kw in mots_du_groupe:
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
    fenetre.geometry("550x700")

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





    # ========================================================================================================
    # RESULTS DIRECT FROM EUROFINS_EXTRACT ================================ NEED ADJUSTMENT FOR MULTIPLE CASES
    # ========================================================================================================

    resultats_artelia = {}

    try:
        with open("résumé_extraction.json", "r", encoding="utf-8") as f:
            resultats_artelia = json.load(f)
    except Exception as e:
        print(f"⚠️ Impossible de charger les résultats : {e}")

    def afficher_ligne_random():
        try:
            import random

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




    # == INIT ==
    tous_keywords = list(matched_columns.keys())

    colonnes_fixes = ["Code Eurofins", "Code Artelia", "Date prélèvement"]

    for kw in tous_keywords:
        if kw not in colonnes_fixes and kw not in zone_droite.get(0, tk.END):
            zone_gauche.insert(tk.END, kw)

    last_used = load_last_config()
    if last_used and os.path.exists(last_used):
        config_path.set(last_used)
        with open(last_used, "r", encoding="utf-8") as f:
            data = json.load(f)
            mots = data.get("keywords_valides", [])
            groupes.update(data.get("groupes_personnalises", {}))
            for kw in mots:
                zone_droite.insert(tk.END, kw)
            for kw in tous_keywords:
                if kw not in mots:
                    zone_gauche.insert(tk.END, kw)
    else:
        default_keys = ["Code Eurofins", "Code Artelia", "Date prélèvement"]
        for kw in default_keys:
            zone_droite.insert(tk.END, kw)
        for kw in tous_keywords:
            if kw not in default_keys:
                zone_gauche.insert(tk.END, kw)

    afficher_groupes()




if __name__ == "__main__":
    root = tk.Tk()
    app = ExtractApp(root)
    root.mainloop()

