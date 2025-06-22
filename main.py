from flask import Flask, request, jsonify
from openpyxl import load_workbook

app = Flask(__name__)  # <--- ESTA LÍNEA ES LA QUE FALTA

@app.route("/llenar_riesgo", methods=["POST"])
def llenar_riesgo():
    try:
        datos = request.get_json()
        wb = load_workbook("AST_WM.xlsx")
        ws = wb.active

        # Eliminar fila 5
        ws.delete_rows(5)

        # Insertar nueva fila vacía en la posición 5
        ws.insert_rows(5)

        # Lista de campos en orden de columnas
        campos = [
            "actividad", "riesgos_detectados", "frecuencia",
            "severidad", "impacto", "medidas_control"
        ]

        for i, campo in enumerate(campos, start=1):
            ws.cell(row=5, column=i).value = datos[campo]

        wb.save("AST_WM.xlsx")
        return jsonify({"mensaje": "Registro exitoso", "fila_insertada": 5}), 200

    except Exception as e:
        return jsonify({"mensaje": f"Error: {str(e)}"}), 500
        if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))  # Render asigna el puerto por variable de entorno
    app.run(host="0.0.0.0", port=port)

