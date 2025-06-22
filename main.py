from flask import Flask, request, jsonify
import openpyxl
import os

app = Flask(__name__)

@app.route('/llenar_riesgo', methods=['POST'])
def llenar_riesgo():
    data = request.json
    archivo_excel = 'AST_WM.xlsx'

    # Cargar archivo y hoja activa
    wb = openpyxl.load_workbook(archivo_excel)
    ws = wb.active

    # Borrar cualquier contenido desde la fila 5 hacia abajo (si existe)
    max_row = ws.max_row
    if max_row >= 5:
        for row in ws.iter_rows(min_row=5, max_row=max_row):
            for cell in row:
                cell.value = None

    # Insertar nueva fila desde la fila 5
    nueva_fila = [
        "Actividad registrada por IA",
        data.get("riesgos_detectados", ""),
        data.get("frecuencia", ""),
        data.get("severidad", ""),
        data.get("impacto", ""),
        data.get("medidas_control", "")
    ]
    ws.append(nueva_fila)

    wb.save(archivo_excel)
    return jsonify({
        "mensaje": "Registro exitoso",
        "fila_insertada": ws.max_row
    })

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
