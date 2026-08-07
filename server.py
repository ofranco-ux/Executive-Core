import os
import math
import gc
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

@app.route('/')
def serve_index():
    return send_from_directory('.', 'index.html')

def erlang_c_sl_optimizado(A, N, AHT, target_time):
    if N <= A or A <= 0 or N <= 0:
        return 0.0
    
    try:
        sum_terms = 1.0
        current_term = 1.0
        int_N = int(N)
        
        # Limite preventivo para evitar loops largos
        if int_N > 1000:
            int_N = 1000

        for k in range(1, int_N):
            current_term *= (A / k)
            sum_terms += current_term
            
        last_term = current_term * (A / N) / (1.0 - (A / N))
        pw = last_term / (sum_terms + last_term)
        
        intensity = N - A
        sl = 1.0 - (pw * math.exp(-intensity * (target_time / AHT)))
        return round(max(0.0, min(100.0, sl * 100.0)), 1)
    except (OverflowError, ZeroDivisionError):
        return 0.0

@app.route('/api/process', methods=['POST'])
@app.route('/api/process/', methods=['POST'])
def process_data():
    if 'file' not in request.files:
        return jsonify({'error': 'No se recibió ningún archivo.'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Nombre de archivo vacío.'}), 400

    try:
        target_sl = float(request.form.get('target_sl', 80))
        target_time = float(request.form.get('target_time', 20))
        merma = float(request.form.get('merma', 20)) / 100.0

        # Intentar leer Excel de la forma más liviana posible
        try:
            df = pd.read_excel(file, engine='calamine')
        except Exception:
            file.seek(0)
            df = pd.read_excel(file, engine='openpyxl')

        df.columns = [str(col).strip() for col in df.columns]

        data_processed = []
        
        # Procesamiento por dict para minimizar overhead de pandas en memoria
        records = df.to_dict(orient='records')
        del df  # Liberar dataframe de la memoria RAM
        gc.collect()

        for row in records:
            calls = float(row.get('Llamadas', 0) or 0)
            aht = float(row.get('AHT', 180) or 180)
            prog = float(row.get('Agentes_Programados', 0) or 0)
            
            a_erlang = (calls * aht) / 1800.0 if aht > 0 else 0.0
            req_raw = a_erlang / (1.0 - merma) if merma < 1.0 else a_erlang
            req_agents = math.ceil(req_raw)
            
            eval_agents = prog if prog > 0 else req_agents
            sl = erlang_c_sl_optimizado(a_erlang, eval_agents, aht, target_time)
            
            data_processed.append({
                'Campaña': str(row.get('Campaña', row.get('Campana', 'General'))),
                'Fecha': str(row.get('Fecha', '')),
                'Día_Semana': str(row.get('Día_Semana', row.get('Dia_Semana', ''))),
                'Intervalo': str(row.get('Intervalo', '00:00')),
                'Llamadas': int(calls),
                'AHT': int(aht),
                'Agentes_Requeridos': req_agents,
                'Agentes_Programados_Reales': int(prog),
                'Delta_Net_Staffing': round(prog - req_agents, 1),
                'SL_Proyectado': sl
            })
            
        del records
        gc.collect()  # Forzar liberación de RAM final

        return jsonify(data_processed)

    except Exception as e:
        gc.collect()
        return jsonify({'error': f"Error al procesar el archivo: {str(e)}"}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
