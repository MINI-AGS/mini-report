import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
from form_tabla import crear_tabla  # Importa la función desde el archivo externo

# Crear la tabla principal, donde se muestran los resultados de la encuesta y sus datos relevantes.
# No se incluyeron las respuestas de cada pregunta.
# La tabla generado aqui se utilizara para generar el resto de reportes, para no tener que estar 
# llamando a la base de datos en cada reporte.
def crear_tabla_principal():
    # La siguiente seccion se remplazara por el llamado y lectura de la base de datos.
    ############################################################################################
    path = "False_data/imaginari_data.txt"

    # Verificar que el archivo existe
    if not os.path.exists(path):
        print("Error: El archivo no existe.")
        return

    # Leer el archivo
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as archivo:
            lineas = [linea.strip().split("|") for linea in archivo]
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
        return
    ############################################################################################

    # Crear un nuevo libro de trabajo
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reporte de Encuesta"

    # Definir encabezados de la tabla
    encabezados = ["Nombre", "Edad", "Sexo", "Preferencia sexual", "Estado de origen", 
                   "Estado de residencia"]
    
    trastornos = ["Episodio depresivo mayor", "Trastorno distímico", 
                   "Riesgo de suicidio", "Episodio (hipo)maníaco", "Trastorno de angustia", 
                   "Agorafobia", "Fobia social", "Trastorno obsesivo-compulsivo", 
                   "Estado por estrés postraumático", "Abuso y dependencia de alcohol", 
                   "Trastornos asociados al uso de sustancias psicoactivas no alcohólicas", 
                   "Trastornos psicóticos", "Anorexia nerviosa", "Bulimia nerviosa", 
                   "Trastorno de ansiedad generalizada", "Trastorno antisocial de la personalidad"]
    
    encabezados += trastornos

    # Índices de los trastornos 
    indice_trastornos = list(range(7, len(encabezados) + 1))

    # Estilos para encabezados
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(left=Side(style="medium"), right=Side(style="medium"),
                         top=Side(style="medium"), bottom=Side(style="medium"))

    # Agregar encabezados
    for col_num, encabezado in enumerate(encabezados, 1):
        celda = ws.cell(row=1, column=col_num, value=encabezado)
        celda.font = header_font
        celda.fill = header_fill
        celda.alignment = header_alignment
        celda.border = thin_border

    # Agregar los datos a la tabla
    for row_num, fila in enumerate(lineas, 2):
        if len(fila) != len(encabezados):
            fila += [""] * (len(encabezados) - len(fila))  # En caso de no haber dato se rellena con un vacio
        
        # Contar cantidad de trastornos del entrevistado actual
        cantidad_si = sum(1 for i in indice_trastornos if i - 1 < len(fila) and fila[i - 1].strip() == "Si")
        
        # Definir color en función para darle al usuario en funcion de la cantidad de trastornos que padece
        if cantidad_si > 0:
            rojo = 255  
            naranja = max(0, 255 - (cantidad_si * 25))  # Disminuye progresivamente de 255 a 0

            color_hex = f"{rojo:02X}{naranja:02X}00"  
            fill_color = PatternFill(start_color=color_hex, fill_type="solid")
        else:
            # No se pone color a usuarios que no sufren ningun trastorno
            fill_color = None

        for col_num, dato in enumerate(fila, 1):
            celda = ws.cell(row=row_num, column=col_num, value=dato)
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            celda.border = thin_border
            # Aplicar color de acuerdo a la cantidad de trastornos al nombre del entrevistado
            if fill_color and col_num == 1:  
                celda.fill = fill_color
            # Se pintan de amarillos los "Si", para identificar más facilmente los trastornos padecidos
            if dato.strip() == "Si":
                celda.fill = PatternFill(start_color="FFD700", fill_type="solid") 

    # Ajustar el ancho de las columnas
    for col in range(1, len(encabezados) + 1):
        max_length = max(len(str(ws.cell(row=row, column=col).value or "")) for row in range(1, len(lineas) + 2))
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = min(max_length + 2, 30)

    # Ajustar altura de filas
    for row in range(1, len(lineas) + 2):
        ws.row_dimensions[row].height = 40 if row == 1 else 20

    # Aplicar formato de tabla 
    crear_tabla(ws, type_table="TableStyleLight11", table_name="Tabla_general")

    # Guardar el archivo Excel 
    output_path = "reporte_principal.xlsx"
    try:
        wb.save(output_path)
        wb.close()
        print(f"Página principal guardada en: '{output_path}'")
    except Exception as e:
        print(f"Error al guardar el archivo: {e}")

if __name__ == "__main__":
    crear_tabla_principal()
