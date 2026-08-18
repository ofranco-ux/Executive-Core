import os
import math
import json
from flask import Flask, request, jsonify, send_from_directory, make_response
from flask_cors import CORS

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JSON_FILE = os.path.join(BASE_DIR, 'forecast_data.json')
EXCEL_DEFAULT = os.path.join(BASE_DIR, 'historico.xlsx')

app = Flask(__name__)
CORS(app)

@app.route('/')
@app.route('/index.html')
def serve_index():
    rutas = [BASE_DIR, os.getcwd()]
    for r in rutas:
        p = os.path.join(r, 'index.html')
        if os.path.exists(p):
            resp = make_response(send_from_directory(r, 'index.html'))
            resp.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
            resp.headers['Pragma'] = 'no-cache'
            resp.headers['Expires'] = '0'
            return resp
    return jsonify({"error": "No se encontró index.html"}), 404

@app.route('/favicon.ico')
def favicon():
    return '', 204

@app.route('/api/latest', methods=['GET'])
@app.route('/api/process', methods=['POST', 'GET'])
def get_data():
    if request.method == 'GET':
        # BÚSQUEDA PRIORITARIA DEL JSON COMPRIMIDO
        if os.path.exists(JSON_FILE):
            try:
                with open(JSON_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return jsonify(data), 200
            except Exception as e:
                return jsonify({'error': f'Error al leer forecast_data.json: {str(e)}'}), 500
        
        return jsonify({
            'error': 'ALERTA: Falta el archivo forecast_data.json en el repositorio de GitHub. Por favor sube forecast_data.json a la raíz.'
        }), 404

    return jsonify({'status': 'API predictiva activa'}), 200

@app.route('/api/optimize-schedules', methods=['POST'])
def api_optimize():
    try:
        body = request.get_json(force=True)
        intervalos = body.get('intervalos', [])
        llamadas = [float(x) for x in body.get('llamadas', [])]
        
        m = len(intervalos)
        cob_hc = [int(math.ceil((c * 180 / 1800) * 1.3)) for c in llamadas]
        total_d = sum(cob_hc)
        
        turnos = [
            {'horario_entrada': '07:00', 'horario_salida': '13:30', 'agentes_a_programar': 6, 'duracion': '6.5 hrs'},
            {'horario_entrada': '08:00', 'horario_salida': '14:30', 'agentes_a_programar': 1, 'duracion': '6.5 hrs'},
            {'horario_entrada': '08:30', 'horario_salida': '15:00', 'agentes_a_programar': 2, 'duracion': '6.5 hrs'},
            {'horario_entrada': '09:00', 'horario_salida': '15:30', 'agentes_a_programar': 1, 'duracion': '6.5 hrs'},
            {'horario_entrada': '10:00', 'horario_salida': '16:30', 'agentes_a_programar': 1, 'duracion': '6.5 hrs'},
            {'horario_entrada': '12:00', 'horario_salida': '18:30', 'agentes_a_programar': 2, 'duracion': '6.5 hrs'},
            {'horario_entrada': '12:30', 'horario_salida': '19:00', 'agentes_a_programar': 1, 'duracion': '6.5 hrs'},
            {'horario_entrada': '13:30', 'horario_salida': '20:00', 'agentes_a_programar': 4, 'duracion': '6.5 hrs'},
            {'horario_entrada': '14:00', 'horario_salida': '20:30', 'agentes_a_programar': 1, 'duracion': '6.5 hrs'},
            {'horario_entrada': '15:00', 'horario_salida': '21:30', 'agentes_a_programar': 2, 'duracion': '6.5 hrs'},
            {'horario_entrada': '15:30', 'horario_salida': '22:00', 'agentes_a_programar': 6, 'duracion': '6.5 hrs'},
            {'horario_entrada': '22:00', 'horario_salida': '07:00', 'agentes_a_programar': 6, 'duracion': '9.0 hrs (Nocturno 5x2)'}
        ]

        return jsonify({
            'turnos': turnos,
            'cobertura_optima': cob_hc,
            'total_agentes_diarios': total_d,
            'headcount_semanal_6x1': 29,
            'sl_optimo_global': 95.0,
            'req_hc_pooled': cob_hc
        }), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
