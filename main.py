from flask import Flask, request, jsonify, send_file
from openpyxl import load_workbook
import os

app = Flask(__name__)

# Ruta segura para escribir dentro del entorno Render
RUTA_EXCEL = "tmp/AST_WM.xlsx"

@app.route("/llenar_riesgo", methods=["POST"])
def llenar_riesgo():
    try:
        datos = request.get_json()
        wb = load_workbook(RUTA_EXCEL)
        ws = wb.active

        # Eliminar fila 5 en caso de que ya exista y tenga datos
        ws.delete_rows(5)
        ws.insert_rows(5)

        # Lista de campos esperados
        campos = [
            "actividad", "riesgos_detectados", "frecuencia",
            "severidad", "impacto", "medidas_control"
        ]

        for i, campo in enumerate(campos, start=1):
            ws.cell(row=5, column=i).value = datos[campo]

        wb.save(RUTA_EXCEL)
        return jsonify({"mensaje": "Registro exitoso", "fila_insertada": 5}), 200

    except Exception as e:
        return jsonify({"mensaje": f"Error: {str(e)}"}), 500

@app.route("/descargar_excel", methods=["GET"])
def descargar_excel():
    try:
        return send_file(RUTA_EXCEL, as_attachment=True)
    except Exception as e:
        return jsonify({"mensaje": f"Error al descargar: {str(e)}"}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
