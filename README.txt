# GeoChemical Analysis Extractor

This desktop application allows you to extract, select, and export specific geochemical analysis parameters from Excel files provided by Eurofins (or other labs in future versions).

### Features

- Load your Excel analysis file
- Select or edit keyword filters
- Automatically detect matching columns
- Choose what to export via interface
- Export results as a clean `.xlsx` file


# ======================================================================================================================================================


## How to Use

1. Launch the application by double-clicking on the `.exe` (no installation required).
2. In the first window:
   - Select a **keyword file** (`.json`) or create one using `Modify`
   - Load your **Excel file** 
   - Select the **extraction type** (currently only available: `Eurofins`)
   - Click **EXTRACT**
3. In the second window:
   - See all detected keywords and matching parameters
   - Move the parameters you want to export to the right side
   - Click **GENERATE JSON** to save your selection and being able to reuse it (optional)
   - Click **EXTRACT TO EXCEL** to export your results


# ======================================================================================================================================================


## Files and Folders

| File                        | Description                              |
|-----------------------------|------------------------------------------|
| `main.exe`                  | Launchable application                   |
| `config_extract.json`       | Last used keyword list (auto loaded)     |
| `sum.json` (optional)       | For custom grouped sums                  |
| `export_analyses_*.xlsx`    | Your final Excel output                  |


# ======================================================================================================================================================


## Need help?

For assistance or feedback, please contact Paul Ancian at @paul.a88@hotmail.fr