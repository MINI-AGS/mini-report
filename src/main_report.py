from crear_principal import crear_tabla_principal
from crear_distribucion_trastornos import reporte_distribucion_trastornos

# Ejecutar todos los archivos
# Excepto "descargar reporte.py", en lo que encuentro una mejor manera de hacer la descarga.
if __name__ == "__main__":
    crear_tabla_principal()
    reporte_distribucion_trastornos()