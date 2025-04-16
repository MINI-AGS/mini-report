import os
import sys

import descargar_reporte
from crear_principal import crear_tabla_principal
from distribucion_trastornos import reporte_distribucion_trastornos
from edad_trastorno import reporte_distribucion_edades_trastornos
from estado_origen_trastorno import reporte_distribucion_estado_origen_trastorno
from estado_residencia_trastorno import reporte_distribucion_estado_residencia_trastorno
from factores_en_trastornos import reporte_caracteristicas_asociadas
from get_data import get_firebase_data
from sexo_trastorno import reporte_distribucion_sexo_trastornos


def resource_path(relative_path):
    """Get absolute path to resource (for PyInstaller)"""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


if __name__ == "__main__":
    path = resource_path("firebase-admin-sdk.json")
    data = get_firebase_data(path)
    crear_tabla_principal(data)
    reporte_distribucion_edades_trastornos()
    reporte_distribucion_sexo_trastornos()
    reporte_distribucion_estado_origen_trastorno()
    reporte_distribucion_estado_residencia_trastorno()
    reporte_distribucion_trastornos()
    reporte_caracteristicas_asociadas()
    descargar_reporte.auto_download_file()
