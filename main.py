from flask import Flask, request, jsonify, send_from_directory
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import os

app = Flask(__name__)

ARCHIVO = "tmp/AST_WM.xlsx"
FILA_INSERTAR = 5  # Fila donde se agregan las actividades

@app.route("/llenar_riesgo", methods=["POST"])
def llenar_riesgo():
    try:
        datos = request.get_json()

        wb = load_workbook(ARCHIVO)
        ws = wb.active

        # Insertar nueva fila en la posición especificada
        ws.insert_rows(FILA_INSERTAR)

        # Campos en el orden correcto de columnas (A a L = 12 columnas)
        columnas = [
            "actividad",  # A
            "condiciones",  # B
            "cond_seguridad",  # C
            "instrucciones",  # D
            "tipo_factor",  # E
            "causas",  # F
            "analisis",  # G
            "frecuencia",  # H
            "severidad",  # I
            "impacto",  # J
            "medidas",  # K
            "observaciones"  # L (opcional)
        ]

        # Llenar y aplicar formato (alineación justificada o centrado según longitud)
        for i, campo in enumerate(columnas, start=1):
            valor = datos.get(campo, "")
            celda = ws.cell(row=FILA_INSERTAR, column=i, value=valor)

            if len(str(valor)) > 20:
                celda.alignment = Alignment(horizontal="justify", vertical="top", wrap_text=True)
            else:
                celda.alignment = Alignment(horizontal="center", vertical="center")

        wb.save(ARCHIVO)
        return jsonify({"mensaje": "Riesgo registrado correctamente", "fila_insertada": FILA_INSERTAR}), 200

    except Exception as e:
        return jsonify({"mensaje": f"Error: {str(e)}"}), 500

@app.route("/static/AST_WM.xlsx")
def descargar_archivo():
    return send_from_directory("tmp", "AST_WM.xlsx", as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
