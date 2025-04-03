import http.server
import socketserver
import os

DIRECTORY = "./"
FILENAME = "reporte_principal.xlsx"
PORT = 8000

class CustomHandler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path == "/download":
            file_path = os.path.join(DIRECTORY, FILENAME)
            if os.path.exists(file_path):
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition", f"attachment; filename={FILENAME}")
                self.end_headers()
                
                # Enviar archivo
                with open(file_path, "rb") as file:
                    self.wfile.write(file.read())

                # Eliminar archivo después de enviarlo
                os.remove(file_path)
                print(f"Archivo {FILENAME} eliminado después de la descarga.")
            else:
                self.send_response(404)
                self.end_headers()
                self.wfile.write(b"Error: Archivo no encontrado.")
        else:
            self.send_response(404)
            self.end_headers()
            self.wfile.write(b"Ruta no encontrada.")

def iniciar_servidor():
    with socketserver.TCPServer(("", PORT), CustomHandler) as httpd:
        print(f"Servidor iniciado en http://localhost:{PORT}/download")
        httpd.serve_forever()
