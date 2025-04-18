import os
import sys
from datetime import datetime

import descargar_reporte
from crear_principal import crear_tabla_principal
from distribucion_trastornos import reporte_distribucion_trastornos
from edad_trastorno import reporte_distribucion_edades_trastornos
from estado_origen_trastorno import reporte_distribucion_estado_origen_trastorno
from estado_residencia_trastorno import reporte_distribucion_estado_residencia_trastorno
from factores_en_trastornos import reporte_caracteristicas_asociadas
from firebase_utils import (
    delete_all_data,
    get_all_data,
    initialize_db,
    save_data_as_csv,
    save_data_as_json,
)
from sexo_trastorno import reporte_distribucion_sexo_trastornos


def resource_path(relative_path):
    """Get absolute path to resource (for PyInstaller)"""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


if __name__ == "__main__":
    path = resource_path("firebase-admin-sdk.json")
    db = initialize_db(path)
    data = get_all_data(db)
    json_file = save_data_as_json(data, directory=".", prefix="data")
    csv_file = save_data_as_csv(data, directory=".", prefix="data")
    report_file = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    crear_tabla_principal(data, report_file)
    reporte_distribucion_edades_trastornos(report_file)
    reporte_distribucion_sexo_trastornos(report_file)
    reporte_distribucion_estado_origen_trastorno(report_file)
    reporte_distribucion_estado_residencia_trastorno(report_file)
    reporte_distribucion_trastornos(report_file)
    reporte_caracteristicas_asociadas(report_file)
    confirm = input("Delete all data from Firestore? (yes/no): ").strip().lower()
    if confirm == "yes":
        delete_all_data(db)
    descargar_reporte.auto_download_file(report_file)
