from analysis_extract import RowsExtract

extractor = RowsExtract(
    excel_path="test_hap.xlsx",
    config_path="test_config.json",
    sheet_name=0,  # ✅ Ceci résout ton problème
    row_config={
        "nom_row": 0,            # ligne des noms d’échantillons (SP01, SP02)
        "nom_col": 1,            # colonne où ça commence (SP01 est en colonne 1)
        "param_row": 2,          # première ligne contenant un paramètre
        "param_col": 0,          # les noms des paramètres sont en colonne A
        "row_limites": None,
        "data_start_row": 2,     # les valeurs commencent ligne 2 (après les noms)
        "data_start_col": 1,     # les valeurs commencent colonne B
        "optionnels": {}
    }
)

extractor.load_keywords()
extractor.load_data()
extractor.extract()

from pprint import pprint
print("\n=== Résultat Artelia ===")
pprint(extractor.resultats_artelia)