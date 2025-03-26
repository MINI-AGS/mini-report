import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
import os

def crear_principal():
    # La siguiente parte se sustituira por una conexion y comprobaciones a la base de datos
    ##############################################################################################
    path = "False_data/imaginari_data.txt"

    # Verificacion del archivo
    if not os.path.exists(path):
        print("Error al encontrar el archivo")
        return
    
    # Leer el archivo
    try:
        with open(path, encoding="utf-8", errors="ignore") as archivo:
            lineas = [linea.strip().split("|") for linea in archivo]
    
    except Exception as e:
        print(f"Error al intentar leer el archivo: {e}")
        return
    ##############################################################################################
    
    # Crear un linbro de trabajo
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reporte de las encuestas"

    # Encabezados de la tabla (sujetos a cambios)
    encabezados = ["Nombre", "Edad", "Sexo", "Preferencia sexual", "Estado de origen", 
                   "Estado de residencia", "Respuesta del cuestionario", "Diagnóstico"]
    
    # Estilos para los encabezados
    header_font = Font(bold=True, color="FFFFFF")