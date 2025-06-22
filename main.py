from flask import Flask, request, jsonify
import pandas as pd
from openpyxl import load_workbook
import os

app = Flask(__name__)

EXCEL_PATH = "AST_WM.xlsx"

@app.route("/")
def index():
    return "Servidor de Riesgos y AST activo."

@app.route("/llenar_riesgo", methods=["POST"])
def llenar_riesgo():
    data = request.json
    wb = load_workbook(EXCEL_PATH)
    ws = wb["Actual"]

    fila = 5
    while ws[f"A{fila}"].value is not None:
        fila += 1

    columnas = {
        "actividad": "A",
        "condiciones_instalaciones": "B",
        "condiciones_seguridad": "C",
        "procedimientos_existentes": "D",
        "tipo_factor_riesgo": "E",
        "causas_posibles": "F",
        "riesgos_detectados": "G",
        "frecuencia": "H",
        "severidad": "I",
        "impacto": "J",
        "medidas_control": "K"
    }

    for campo, letra in columnas.items():
        if campo in data:
            ws[f"{letra}{fila}"] = data[campo]

    wb.save(EXCEL_PATH)
    return jsonify({"mensaje": "Riesgo registrado correctamente", "fila_insertada": fila})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)