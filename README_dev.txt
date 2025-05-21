# GeoChemical Analysis Extractor – Developer Notes

This project is a modular Python application used to extract, filter, and export structured geochemical data from Excel files (Eurofins, Agrolab, etc.) using a GUI (Tkinter).

---

## 🧱 Project Structure

| File/Folder               | Purpose                                                       |
|---------------------------|---------------------------------------------------------------|
| `main.py`                 | Main launcher and UI 1 (keywords, Excel, config type)         |
| `ui_post_extract.py`      | UI 2: parameter selection, grouping, and Excel export         |
| `analysis_extract.py`     | Core logic for data extraction (RowsExtract, ColumnsExtract)  |
| `extract_utils.py`        | Utility functions (cell parsing, keyword normalization)       |
| `config_extract.json`     | Default keyword/group config (used in UI 2)                   |
| `config_type_*.json`      | Cell coordinate presets for "lignes" and "colonnes" modes     |
| `temp_keywords.json`      | Temporary file storing active keywords before extraction      |
| `résumé_extraction.json`  | Raw results dump for debug and validation                     |
| `sum.json`                | (Optional) predefined groups of parameters to sum             |

---

## ⚙️ Runtime Logic

1. **UI 1** (`main.py`)
   - Load keyword JSON
   - Load Excel file and sheet
   - Select extraction type (columns or rows)
   - Configure cell positions (param, nom, data)
   - Launch extractor

2. **Extraction**
   - Depending on type, either `ColumnsExtract` or `RowsExtract` is used
   - Keywords are matched against headers
   - Results stored in `self.resultats_artelia`

3. **UI 2** (`ui_post_extract.py`)
   - Lets user refine the list of parameters and groups
   - Generates new config file (`config_extract.json`) or temporary export JSON
   - Final Excel export is triggered and written via `export()`

---

## 🧪 Testing Tips

- Use Excel test files with well-known layouts
- Add `print(self.df.head())` inside `load_data()` for preview
- Check `résumé_extraction.json` to inspect raw results
- Use drag-and-drop reordering in UI 2 for debugging export order

---

## 📦 Setup (Development)

You’ll need Python 3.10+ installed.

Install dependencies:

```bash
pip install -r requirements.txt