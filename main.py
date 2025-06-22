from flask import Flask, request, jsonify, send_file
from openpyxl import load_workbook
import shutil
import os

app = Flask(__name__)

@app.route("/llenar_riesgo", methods=["POST"])
def llenar_riesgo():
    try:
        # Copiar archivo a /tmp para poder modificarlo
        ORIG_PATH = "AST_WM.xlsx"
        TMP_PATH = "/tmp/AST_WM.xlsx"
        shutil.copyfile(ORIG_PATH, TMP_PATH)

        # Cargar y modificar el archivo desde /tmp
        wb = load_workbook(TMP_PATH)
        ws = wb.active

        # Eliminar y reiniciar fila 5
        ws.delete_rows(5)
        ws.insert_rows(5)

        datos = request.get_json()
        campos = [
            "actividad", "riesgos_detectados", "frecuencia",
            "severidad", "impacto", "medidas_control"
        ]

        for i, campo in enumerate(campos, start=1):
            ws.cell(row=5, column=i).value = datos[campo]

        wb.save(TMP_PATH)

        return jsonify({
            "mensaje": "Registro exitoso en AST_WM.xlsx",
            "archivo_modificado": "/tmp/AST_WM.xlsx"
        }), 200

    except Exception as e:
        return jsonify({"mensaje": f"Error: {str(e)}"}), 500


@app.route("/descargar_excel", methods=["GET"])
def descargar_excel():
    TMP_PATH = "/tmp/AST_WM.xlsx"
    if os.path.exists(TMP_PATH):
        return send_file(TMP_PATH, as_attachment=True)
    else:
        return jsonify({"mensaje": "Archivo no encontrado"}), 404


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
