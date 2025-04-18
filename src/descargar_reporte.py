import os
import shutil


def auto_download_file(filename):
    DEST_DIR = os.path.expanduser("~/Downloads")  # or any path you want
    SOURCE_PATH = os.path.join(".", filename)
    DEST_PATH = os.path.join(DEST_DIR, filename)

    if os.path.exists(SOURCE_PATH):
        shutil.copy(SOURCE_PATH, DEST_PATH)
        print(f"Reporte guardado en {DEST_PATH}")
    else:
        print("Error: Ocurrio un error al guardar el reporte.")


if __name__ == "__main__":
    auto_download_file()
