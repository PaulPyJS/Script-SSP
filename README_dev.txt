# GeoChemical Analysis Extractor – Developer Notes

This is a modular Python application for extracting and exporting parameters from structured Excel reports (Eurofins, etc.), with a Tkinter-based UI.

---

## Structure

| File                     | Purpose                                    |
|--------------------------|--------------------------------------------|
| `main.py`                | Main UI (Tkinter) with config + launch     |
| `Eurofins_extract.py`    | Extraction + optional export logic         |
| `sum.json`               | Groups of parameters to be summed          |
| `matched_columns.json`   | Intermediate file with matched keywords    |
| `final_keywords.json`    | List of keywords selected for export       |

---

## Installation (for development)

Make sure you have Python 3.8+ installed.  
Install the required packages:

```bash
pip install -r requirements.txt