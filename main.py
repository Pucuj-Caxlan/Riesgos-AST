from flask import Flask, request, jsonify, send_from_directory
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import os

app = Flask(__name__)

@app.route("/llenar_riesgo", methods=["POST"])
def llenar_riesgo():
    try:
        datos = request.get_json()

        # Ruta del archivo en carpeta tmp (Render permite escritura ahí)
        excel_path = "/tmp/AST_WM.xlsx"
        template_path = "tmp/AST_WM.xlsx"

        # Si el archivo no existe en tmp, cópialo desde el repo
        if not os.path.exists(excel_path):
            from shutil import copyfile
            copyfile(template_path, excel_path)

        wb = load_workbook(excel_path)
        ws = wb.active

        # Elimina la fila 5 si ya tiene contenido (usualmente contiene celdas combinadas)
        ws.delete_rows(5)
        ws.insert_rows(5)

        # Mapeo en orden de columnas A–L (1–12)
        columnas = {
            1: datos["actividad"],
            2: datos["riesgos_detectados"],
            3: "B",     # Condiciones de herramientas y equipo (según STPS)
            4: "SI",    # Instrucciones de seguridad
            5: "MEDIO", # Tipo de factor de riesgo
            6: "Lesiones por impacto, cortes, proyecciones",
            7: datos["frecuencia"],
            8: datos["severidad"],
            9: datos["impacto"],
            10: datos["medidas_control"],
        }

        # Insertar y formatear celdas
        for col, valor in columnas.items():
            celda = ws.cell(row=5, column=col)
            celda.value = valor
            if len(str(valor)) <= 5:
                celda.alignment = Alignment(horizontal="center", vertical="center")
            else:
                celda.alignment = Alignment(horizontal="justify", vertical="top", wrap_text=True)

        wb.save(excel_path)
        return jsonify({"mensaje": "Registro exitoso", "fila_insertada": 5}), 200

    except Exception as e:
        return jsonify({"mensaje": f"Error: {str(e)}"}), 500

@app.route("/static/AST_WM.xlsx", methods=["GET"])
def descargar_archivo():
    return send_from_directory("/tmp", "AST_WM.xlsx", as_attachment=True)

# Configuración para despliegue en Render
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
