@app.route("/llenar_riesgo", methods=["POST"])
def llenar_riesgo():
    try:
        datos = request.get_json()
        wb = load_workbook("AST_WM.xlsx")
        ws = wb.active

        # Eliminar fila 5 (que contiene celdas combinadas)
        ws.delete_rows(5)

        # Insertar nueva fila vacía en la posición 5
        ws.insert_rows(5)

        # Lista de campos en orden de columnas
        campos = [
            "actividad", "riesgos_detectados", "frecuencia",
            "severidad", "impacto", "medidas_control"
        ]

        # Insertar datos en la nueva fila
        for i, campo in enumerate(campos, start=1):
            ws.cell(row=5, column=i).value = datos[campo]

        wb.save("AST_WM.xlsx")
        return jsonify({"mensaje": "Registro exitoso", "fila_insertada": 5}), 200

    except Exception as e:
        return jsonify({"mensaje": f"Error: {str(e)}"}), 500
