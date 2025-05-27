import unicodedata
import re
import json
import os
import pandas as pd


def extraire_valeur(cellule):
    if isinstance(cellule, pd.Series):
        cellule = cellule.dropna().astype(str).str.strip()
        return cellule.iloc[0] if not cellule.empty else None
    if pd.isna(cellule) or str(cellule).strip() == "":
        return None
    return str(cellule).strip()

def normalize(text):
    if text is None or not isinstance(text, str):
        text = str(text) if text is not None else ""
    text = text.lower()
    return unicodedata.normalize('NFD', text).encode('ascii', 'ignore').decode('utf-8')

def clean_tokens(text):
    return re.findall(r'[a-z0-9]+', normalize(text))


def col_idx_to_excel_letter(idx):
    letter = ''
    while idx >= 0:
        letter = chr(idx % 26 + 65) + letter
        idx = idx // 26 - 1
    return letter

def cell_to_index(cell: str) -> tuple[int, int]:
    letters = ''.join([c for c in cell if c.isalpha()])
    digits = ''.join([c for c in cell if c.isdigit()])

    col = 0
    for i, char in enumerate(reversed(letters.upper())):
        col += (ord(char) - ord('A') + 1) * (26 ** i)
    col -= 1  # 0-indexed

    row = int(digits) - 1

    return row, col

def sauvegarder_json(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

def charger_json(path: str) -> dict:
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def charger_groupes_parametres(fichier="sum.json"):
    if not os.path.exists(fichier):
        sauvegarder_json(fichier, {})
    return charger_json(fichier)

def sauvegarder_groupes_parametres(groupes, fichier="sum.json"):
    sauvegarder_json(fichier, groupes)


def formater_valeur(val):
    if pd.isna(val):
        return "<LQ"

    val_str = str(val).strip().lower()

    if val_str.startswith("<"):
        return f"<LQ ({val_str})"

    if val_str in {"n.d.", "n.d", "nd", "-", "n.d,", "n.d..", ""}:
        return "<LQ"

    return val_str

def is_label_all(label_info):
    return isinstance(label_info, tuple) and label_info[1] == "all"