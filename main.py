from flask import Flask, request, jsonify
from flask_cors import CORS
import pandas as pd
import os

app = Flask(__name__)
CORS(app)

EXCEL_PATH = "AST_WM.xlsx"

@app.route("/llenar_riesgo", methods=["POST"])
def llenar_riesgo():
    try:
        # Datos JSON recibidos desde el GPT
        datos = request.get_json()

        # Si no existe el archivo, crear uno con encabezados
        if not os.path.exists(EXCEL_PATH):
            columnas = [
                "Actividad", "Riesgos detectados", "Frecuencia",
                "Severidad", "Impacto", "Medidas de control"
            ]
            df = pd.DataFrame(columns=columnas)
            df.to_excel(EXCEL_PATH, index=False)

        # Cargar archivo existente
        df = pd.read_excel(EXCEL_PATH)

        # Si la fila 5 ya contiene datos, borrarla
        if len(df) >= 5:
            df.drop(index=4, inplace=True)  # Fila 5 es índice 4
            df.reset_index(drop=True, inplace=True)

        # Agregar la nueva fila
        nueva_fila = {
            "Actividad": datos["actividad"],
            "Riesgos detectados": datos["riesgos_detectados"],
            "Frecuencia": datos["frecuencia"],
            "Severidad": datos["severidad"],
            "Impacto": datos["impacto"],
            "Medidas de control": datos["medidas_control"]
        }
        df.loc[len(df)] = nueva_fila

        # Guardar cambios
        df.to_excel(EXCEL_PATH, index=False)

        return jsonify({
            "mensaje": "Registro exitoso",
            "fila_insertada": len(df)
        }), 200

    except Exception as e:
        return jsonify({
            "mensaje": f"Error al registrar el riesgo: {str(e)}"
        }), 500

if __name__ == "__main__":
    app.run(debug=False, port=10000, host="0.0.0.0")
