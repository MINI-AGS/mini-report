import os
import platform
import shutil
import subprocess

FILENAME = "reporte_principal.xlsx"
DEST_DIR = os.path.expanduser("~/Downloads")  # or any path you want
SOURCE_PATH = os.path.join(".", FILENAME)
DEST_PATH = os.path.join(DEST_DIR, FILENAME)


def auto_download_file():
    if os.path.exists(SOURCE_PATH):
        shutil.copy(SOURCE_PATH, DEST_PATH)
        print(f"Reporte guardado en {DEST_PATH}")
    else:
        print("Error: Ocurrio un error al guardar el reporte.")


if __name__ == "__main__":
    auto_download_file()
