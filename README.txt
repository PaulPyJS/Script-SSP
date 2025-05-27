# GeoChemical Analysis Extractor

This desktop application allows you to extract, select, and export specific geochemical analysis parameters from heterogeneous Excel files (initially based on multiple Eurofins and Agrolab tabs)

### Features

- Load your Excel analysis file 
- Select or edit keyword filters
- Input your own extract type based on rows or columns type + cells where the data start /!\ NOT HEADERS /!\
	- Save your extract_type in .json to reuse it 
- Choose what to export via interface and create groups to add values
	- Save your export_type in .json to reuse it 
- Export results as a clean `.xlsx` file
	- If values nd, -, <LQ, then result will be <LQ. if no data at all then empty cell


# ======================================================================================================================================================


## How to Use

1. Launch the application by double-clicking on the `.exe` (no installation required).
2. In the first window:
   - Select a **keyword file** (`.json`) or create one using `Modify`
   - Load your **Excel file** 
   - Select the **extraction type** in Rows if your parameters are displayed row by row (or in Columns if param displayed col by col)
   - Adapt the type depending on your exact Excel format by input Cells of the first data for each mandatory parameters
   - Add optional parameters if required (still in dev v2.2)
   - Click **EXTRACT**
3. In the second window:
   - See all detected keywords and matching parameters
   - Move the parameters you want to export to the right side
   - Create your own personal groups to be able to add parameters and create sum
   - Click **GENERATE JSON** to save your selection and groups and being able to reuse it 
   - Click **EXTRACT TO EXCEL** to export your results


# ======================================================================================================================================================


## Files and Folders
  
| File                         | Description                                             |
|------------------------------|---------------------------------------------------------|
| `main.exe`                   | The application launcher                                |
| `config_extract_*.json`      | Excel extraction file for UI 2                          |
| `config_type_*.json`         | Type extraction file for UI 1                           |
| `temp_keywords.json`         | Temporary file used during extraction                   |
| `final_keywords.json`        | Temporary file used for Excel extraction                |
| `last_*`                     | Multiple files used to keep memory in between session   |


# ======================================================================================================================================================


## 📦 Requirements (if running from source)

- Python 3.10+
- `pandas >= 1.3`
- `openpyxl >= 3.0`

(Install with `pip install -r requirements.txt`)

---

## ❓ Need help?

Contact: **Paul Ancian**  
📧 paul.a88@hotmail.fr