from flask import Flask, request, jsonify
import openpyxl
import os

app = Flask(__name__)

@app.route('/llenar_riesgo', methods=['POST'])
def llenar_riesgo():
    try:
        data = request.json
        archivo_excel = 'AST_WM.xlsx'
        wb = openpyxl.load_workbook(archivo_excel)
        ws = wb.active

        # Limpiar desde fila 5 hacia abajo
        for row in ws.iter_rows(min_row=5, max_row=ws.max_row):
            for cell in row:
                cell.value = None

        # Insertar nueva fila manualmente en fila 5
        ws.cell(row=5, column=1, value="Actividad registrada por IA")
        ws.cell(row=5, column=2, value=data.get("riesgos_detectados", ""))
        ws.cell(row=5, column=3, value=data.get("frecuencia", ""))
        ws.cell(row=5, column=4, value=data.get("severidad", ""))
        ws.cell(row=5, column=5, value=data.get("impacto", ""))
        ws.cell(row=5, column=6, value=data.get("medidas_control", ""))

        wb.save(archivo_excel)

        return jsonify({
            "mensaje": "Registro exitoso",
            "fila_insertada": 5
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
