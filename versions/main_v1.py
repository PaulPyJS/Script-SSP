import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import os
import subprocess

FICHIER_SESSION = "last_session.json"
FICHIER_TEMP_KEYWORDS = "temp_keywords.json"
TYPES_EXTRACTION = ["Eurofins", "TEST"]


# === Script : MAIN UI - TABLEURS EUROFINS ===
# = v1 : Import Excel and JSON for keywords + able to modify keywords
# = v1.5 : Modify type of extraction possible for future
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

if __name__ == "__main__":
    root = tk.Tk()
    app = ExtractApp(root)
    root.mainloop()

