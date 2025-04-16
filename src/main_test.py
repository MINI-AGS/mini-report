import os, sys
import descargar_reporte
from crear_principal import crear_tabla_principal
from distribucion_trastornos import reporte_distribucion_trastornos
from edad_trastorno import reporte_distribucion_edades_trastornos
from estado_origen_trastorno import reporte_distribucion_estado_origen_trastorno
from estado_residencia_trastorno import reporte_distribucion_estado_residencia_trastorno
from factores_en_trastornos import reporte_caracteristicas_asociadas
from pref_sexual_trastorno import reporte_distribucion_trastornos_pref_sexual
from sexo_trastorno import reporte_distribucion_sexo_trastornos


def obtener_ruta_archivo(nombre_archivo):
    if getattr(sys, 'frozen', False):
        # Ejecutable (.exe)
        base_path = os.path.dirname(sys.executable)
    else:
        # Modo script
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, nombre_archivo)


# Usamos la función para obtener la ruta del JSON
path = obtener_ruta_archivo("imaginari_data.json")

# Ejecutar todos los archivos
if __name__ == "__main__":
    print("Usando ruta:", path)
    crear_tabla_principal(path)
    reporte_distribucion_edades_trastornos()
    reporte_distribucion_sexo_trastornos()
    reporte_distribucion_estado_origen_trastorno()
    reporte_distribucion_estado_residencia_trastorno()
    reporte_distribucion_trastornos_pref_sexual()
    reporte_distribucion_trastornos()
    reporte_caracteristicas_asociadas()
    descargar_reporte.auto_download_file()
    