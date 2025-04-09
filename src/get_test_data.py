import json
import os


def get_test_data(path):

    if not os.path.exists(path):
        print("Error: El archivo no existe.")
        return

    # Leer el archivo JSON
    try:
        with open(path, "r", encoding="utf-8") as archivo:
            data = json.load(archivo)
    except Exception as e:
        print(f"Error al leer el archivo: {e}")
        return

    return data


if __name__ == "__main__":
    get_test_data()
