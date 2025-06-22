from flask import Flask, request, jsonify
from openpyxl import load_workbook
import os
import shutil

app = Flask(__name__)

# Copia inicial del archivo al entorno de escritura
ORIG_PATH = "AST_WM.xlsx"
TMP_PATH = "/tmp/AST_WM.xlsx"

if not os.path.exists(TMP_PATH):
    shutil.copyfile(ORIG_PATH, TMP_PATH)

@app.route("/llenar_riesgo", methods=["POST"])
def llenar_riesgo():
    try:
        datos = request.get_json()
        wb = load_workbook(TMP_PATH)
        ws = wb.active

        # Eliminar fila 5
        ws.delete_rows(5)
        # Insertar nueva fila vacía
        ws.insert_rows(5)

        campos = [
            "actividad", "riesgos_detectados", "frecuencia",
            "severidad", "impacto", "medidas_control"
        ]

        for i, campo in enumerate(campos, start=1):
            ws.cell(row=5, column=i).value = datos[campo]

        wb.save(TMP_PATH)
        return jsonify({"mensaje": "Registro exitoso", "fila_insertada": 5}), 200

    except Exception as e:
        return jsonify({"mensaje": f"Error: {str(e)}"}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
