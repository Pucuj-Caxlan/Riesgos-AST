from flask import Flask, request, jsonify
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import os

app = Flask(__name__)

@app.route("/llenar_riesgo", methods=["POST"])
def llenar_riesgo():
    try:
        datos = request.get_json()

        # Ruta relativa al archivo dentro de /tmp
        archivo_excel = "tmp/AST_WM.xlsx"
        wb = load_workbook(archivo_excel)
        ws = wb.active

        # Eliminar fila 5 y dejarla lista
        ws.delete_rows(5)
        ws.insert_rows(5)

        # Campos en orden, incluyendo los que ahora deseas llenar
        campos = [
            "actividad",                    # A
            "riesgos_detectados",          # B
            "condiciones_seguridad",       # C (nuevo)
            "frecuencia",                  # D
            "severidad",                   # E
            "impacto",                     # F
            "medidas_control"             # G
        ]

        for i, campo in enumerate(campos, start=1):
            valor = datos.get(campo, "")
            valor = valor.replace("*", "").strip()  # Limpia asteriscos
            celda = ws.cell(row=5, column=i)
            celda.value = valor

            # Aplica formato: justificar si largo, centrar si corto
            if len(valor) > 30:
                celda.alignment = Alignment(horizontal='justify', vertical='top', wrap_text=True)
            else:
                celda.alignment = Alignment(horizontal='center', vertical='center')

        wb.save(archivo_excel)
        return jsonify({"mensaje": "Registro exitoso", "fila_insertada": 5}), 200

    except Exception as e:
        return jsonify({"mensaje": f"Error: {str(e)}"}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
