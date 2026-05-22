from flask import Flask, request, jsonify
from flask_cors import CORS
import openpyxl, io, base64
from openpyxl.styles import Alignment, Font
from openpyxl.cell.rich_text import CellRichText

app = Flask(__name__)
CORS(app)

TEMPLATE_B64 = open('template.b64').read().strip()
CAP_MESES = {12:1, 4:3, 2:6}

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

@app.route('/generar-excel', methods=['POST'])
def generar_excel():
    try:
        d = request.json
        orig = base64.b64decode(TEMPLATE_B64)
        wb = openpyxl.load_workbook(io.BytesIO(orig))
        ws1 = wb['Datos']
        ws2 = wb['Matriz de Sensibilidad']

        cm = CAP_MESES.get(int(d.get('cap', 12)), 1)

        # Borrar comentarios
        ws1.comments = {}
        ws2.comments = {}

        # Celdas amarillas
        ws1['C7'] = float(d['tasa']) / 100
        ws1['E7'] = 12
        ws1['E8'] = cm
        ws1['E12'] = cm
        ws1['E13'] = 1
        ws1['H7'] = float(d['cok']) / 100
        ws1['E18'] = float(d['precio'])
        ws1['E19'] = float(d['pctIni']) / 100
        ws1['E22'] = int(d['plazo'])
        ws1['E28'] = float(d['gastos'])
        ws1['E34'] = float(d['alq'])
        ws1['E38'] = float(d['mnt'])
        ws1['E39'] = float(d['imp'])
        ws1['E40'] = float(d['mng'])
        ws1['E46'] = float(d['h1'])
        ws1['E48'] = float(d['h2'])
        ws1['E50'] = float(d['h3'])

        # Matriz apreciaciones
        ws2['D3'] = float(d['a1']) / 100
        ws2['E3'] = float(d['a2']) / 100
        ws2['F3'] = float(d['a3']) / 100
        ws2['G3'] = float(d['a4']) / 100

        # Wrap text columna K
        wrap = Alignment(wrap_text=True, vertical='top')
        for row in ws1.iter_rows():
            for cell in row:
                if cell.column == 11 and cell.value:
                    cell.alignment = wrap

        # Fuente F3 G3 Matriz
        font_base = Font(name='Aptos Narrow', size=11)
        ws2['F3'].font = font_base
        ws2['G3'].font = font_base

        # Anchos columnas Datos
        for col, w in [('A',10.664),('E',13.109),('F',12.219),('J',10.664),
                       ('K',10.887),('L',10.664),('N',18.777),('O',10.664)]:
            ws1.column_dimensions[col].width = w

        # Anchos columnas Matriz
        for col, w in [('A',10.664),('B',16.664),('C',2.887),('D',12.219),('H',10.664)]:
            ws2.column_dimensions[col].width = w

        # Alturas filas Datos
        for row, h in [(24,48.0),(44,30.0),(69,45.0),(71,50.25)]:
            ws1.row_dimensions[row].height = h

        # Alturas filas Matriz
        ws2.row_dimensions[6].height = 22.5

        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        result_b64 = base64.b64encode(buf.read()).decode('utf-8')

        return jsonify({'status': 'ok', 'file': result_b64})

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
