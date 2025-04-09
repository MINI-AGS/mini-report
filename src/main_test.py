import descargar_reporte
from crear_principal import crear_tabla_principal
from distribucion_trastornos import reporte_distribucion_trastornos
from edad_trastorno import reporte_distribucion_edades_trastornos
from estado_origen_trastorno import reporte_distribucion_estado_origen_trastorno
from estado_residencia_trastorno import reporte_distribucion_estado_residencia_trastorno
from factores_en_trastornos import reporte_caracteristicas_asociadas
from get_test_data import get_test_data
from sexo_trastorno import reporte_distribucion_sexo_trastornos

if __name__ == "__main__":
    data = get_test_data("false_data/jtest_data.json")
    crear_tabla_principal(data)
    reporte_distribucion_edades_trastornos()
    reporte_distribucion_sexo_trastornos()
    reporte_distribucion_estado_origen_trastorno()
    reporte_distribucion_estado_residencia_trastorno()
    reporte_distribucion_trastornos()
    reporte_caracteristicas_asociadas()
    descargar_reporte.auto_download_file()
