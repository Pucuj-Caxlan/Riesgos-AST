from flask import Flask, request, jsonify
from openpyxl import load_workbook

app = Flask(__name__)

@app.route("/llenar_riesgo", methods=["POST"])
def llenar_riesgo():
    try:
        datos = request.get_json()

        # Lista de campos requeridos y su orden en columnas
        campos = [
            "actividad", "riesgos_detectados", "frecuencia",
            "severidad", "impacto", "medidas_control"
        ]

        # Validar que todos los campos estén presentes y no vacíos
        for campo in campos:
            if campo not in datos or not str(datos[campo]).strip():
                return jsonify({
                    "mensaje": f"Falta o está vacío el campo requerido: '{campo}'"
                }), 400

        wb = load_workbook("AST_WM.xlsx")
        ws = wb.active

        # Eliminar fila 5 (que contiene celdas combinadas)
        ws.delete_rows(5)

        # Insertar nueva fila vacía en la posición 5
        ws.insert_rows(5)

        # Insertar datos en la nueva fila
        for i, campo in enumerate(campos, start=1):
            ws.cell(row=5, column=i).value = datos[campo]

        wb.save("AST_WM.xlsx")
        return jsonify({"mensaje": "Registro exitoso", "fila_insertada": 5}), 200

    except Exception as e:
        return jsonify({"mensaje": f"Error: {str(e)}"}), 500

# Punto final bien indentado
if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
