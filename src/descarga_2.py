import os
import subprocess
from flask import Flask, send_file

app = Flask(__name__)

# Ruta del archivo generado
output_file = "reporte_principal.xlsx"

@app.route('/descargar_reporte')
def descargar_reporte():
    # Ejecutar script que genera el reporte
    try:
        subprocess.run(["python", "tu_script_principal.py"], check=True)
    except subprocess.CalledProcessError as e:
        return f"Error ejecutando el script: {e}"

    # Verificar si se generó el archivo
    if not os.path.exists(output_file):
        return "El archivo no fue generado."

    # Descargar el archivo
    return send_file(output_file, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True, port=5000)
