from crear_principal import crear_tabla_principal
from distribucion_trastornos import reporte_distribucion_trastornos
from edad_trastorno import reporte_distribucion_edades_trastornos
from sexo_trastorno import reporte_distribucion_sexo_trastornos
from estado_origen_trastorno import reporte_distribucion_estado_origen_trastorno
from estado_residencia_trastorno import reporte_distribucion_estado_residencia_trastorno
from factores_en_trastornos import reporte_caracteristicas_asociadas
import descargar_reporte
import threading

path = "false_data/jtest_data.json"
# Ejecutar todos los archivos
# Excepto "descargar reporte.py", en lo que encuentro una mejor manera de hacer la descarga.
if __name__ == "__main__":
    crear_tabla_principal(path)
    reporte_distribucion_edades_trastornos()
    reporte_distribucion_sexo_trastornos()
    reporte_distribucion_estado_origen_trastorno()
    reporte_distribucion_estado_residencia_trastorno()
    reporte_distribucion_trastornos()
    reporte_caracteristicas_asociadas()
    # Iniciar servidor en un hilo para no bloquear la ejecución
    descargar_reporte.iniciar_servidor() # Espera a que el hilo del servidor termine

