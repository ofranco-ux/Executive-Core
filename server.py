import os
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd
import math

# Inicializar Flask indicando la carpeta actual para archivos estáticos
app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

# --- RUTA PRINCIPAL (Para cargar la página web en Render) ---
@app.route('/')
def serve_index():
    return send_from_directory('.', 'index.html')

def erlang_c_sl(A, N, AHT, target_time):
    if N <= A or A <= 0:
        return 0.0
    
    sum_terms = sum((A**k) / math.factorial(k) for k in range(int(N)))
    last_term = (A**N) / (math.factorial(int(N)) * (1 - (A / N)))
    pw = last_term / (sum_terms + last_term)
    
    intensity = N - A
    sl = 1 - (pw * math.exp(-intensity * (target_time / AHT)))
    return round(max(0.0, min(1.0, sl)) * 100, 1)

# --- RUTA DE PROCESAMIENTO API ---
@app.route('/api/process', methods=['POST'])
def process_data():
    if 'file' not in request.files:
        return jsonify({'error': 'No se envió ningún archivo'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Nombre de archivo vacío'}), 400

    try:
        target_sl = float(request.form.get('target_sl', 80))
        target_time = float(request.form.get('target_time', 20))
        merma = float(request.form.get('merma', 20)) / 100.0
        h_inicio = request.form.get('h_inicio', '09:00')
        h_fin = request.form.get('h_fin', '20:00')

        df = pd.read_excel(file)
        
        data_processed = []
        for index, row in df.iterrows():
            calls = float(row.get('Llamadas', 0))
            aht = float(row.get('AHT', 180))
            prog = float(row.get('Agentes_Programados', 0))
            
            a_erlang = (calls * aht) / 1800.0 if aht > 0 else 0
            
            req_raw = a_erlang / (1 - merma) if merma < 1 else a_erlang
            req_agents = math.ceil(req_raw)
            
            sl = erlang_c_sl(a_erlang, prog if prog > 0 else req_agents, aht, target_time)
            
            data_processed.append({
                'Campaña': str(row.get('Campaña', 'General')),
                'Fecha': str(row.get('Fecha', '')),
                'Día_Semana': str(row.get('Día_Semana', '')),
                'Intervalo': str(row.get('Intervalo', '00:00')),
                'Llamadas': int(calls),
                'AHT': int(aht),
                'Agentes_Requeridos': req_agents,
                'Agentes_Programados_Reales': int(prog),
                'Delta_Net_Staffing': round(prog - req_agents, 1),
                'SL_Proyectado': sl
            })
            
        return jsonify(data_processed)

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)