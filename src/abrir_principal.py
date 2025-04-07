import openpyxl

# Función para cargar el archivo principal generado de "crear_principal.py"
def cargar_principal():
    try:
        wb = openpyxl.load_workbook("reporte_principal.xlsx")
    except FileNotFoundError:
        print("El archivo no se encuentra en la carpeta.")
        return  None
    
    ws = wb.active
    ws.title = "Reporte de Encuesta"

    # Se retorna el libro y hoja principal donde se encuentra la información.
    return wb, ws