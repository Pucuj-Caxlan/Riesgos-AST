from flask import Flask, request, jsonify
from openpyxl import load_workbook
import os

app = Flask(__name__)

@app.route("/llenar_riesgo", methods=["POST"])
def llenar_riesgo():
    try:
        datos = request.get_json()

        # Ruta al archivo Excel dentro de la carpeta tmp
        path_excel = "tmp/AST_WM.xlsx"

        # Abrir el archivo
        wb = load_workbook(path_excel)
        ws = wb.active

        # Buscar la siguiente fila vacía (columna A)
        fila = 5
        while ws.cell(row=fila, column=1).value:
            fila += 1

        # Lista de campos en orden de columnas
        campos = [
            "actividad", "riesgos_detectados", "frecuencia",
            "severidad", "impacto", "medidas_control"
        ]

        # Insertar los datos en la fila encontrada
        for i, campo in enumerate(campos, start=1):
            ws.cell(row=fila, column=i).value = datos[campo]

        wb.save(path_excel)
        return jsonify({"mensaje": "Registro exitoso", "fila_insertada": fila}), 200

    except Exception as e:
        return jsonify({"mensaje": f"Error: {str(e)}"}), 500

# Corre el servidor Flask
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
