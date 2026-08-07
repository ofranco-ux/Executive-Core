import os
import math
import gc
import re
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

# --- RUTAS FRONTEND ---
@app.route('/')
@app.route('/index.html')
def serve_index():
    return send_from_directory('.', 'index.html')

def clean_num(val, default=0.0):
    if pd.isna(val) or val is None:
        return default
    try:
        val_str = str(val).strip().replace(',', '.')
        val_str = re.sub(r'[^0-9.]', '', val_str)
        return float(val_str) if val_str else default
    except Exception:
        return default

def erlang_c_sl_optimizado(A, N, AHT, target_time):
    if N <= A or A <= 0 or N <= 0:
        return 0.0
    try:
        sum_terms = 1.0
        current_term = 1.0
        int_N = min(int(N), 1000)

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

def encontrar_columna(df, posibles_nombres):
    columnas_df = {str(c).strip().lower(): c for c in df.columns}
    for pos in posibles_nombres:
        pos_clean = pos.strip().lower()
        if pos_clean in columnas_df:
            return columnas_df[pos_clean]
    return None

# --- RUTAS BACKEND API (Registradas para evitar el error 404) ---
@app.route('/api/process', methods=['POST', 'GET'])
@app.route('/api/process/', methods=['POST', 'GET'])
def process_data():
    if request.method == 'GET':
        return jsonify({'status': 'API operativa en Render'}), 200

    if 'file' not in request.files:
        return jsonify({'error': 'No se recibió ningún archivo.'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'Nombre de archivo vacío.'}), 400

    try:
        target_sl = clean_num(request.form.get('target_sl'), 80.0)
        target_time = clean_num(request.form.get('target_time'), 20.0)
        merma = clean_num(request.form.get('merma'), 20.0) / 100.0

        try:
            df = pd.read_excel(file, engine='calamine')
        except Exception:
            file.seek(0)
            df = pd.read_excel(file, engine='openpyxl')

        col_calls = encontrar_columna(df, ['Llamadas', 'Calls', 'Volumen', 'Ofrecidas', 'Llamadas_Ofrecidas'])
        col_aht = encontrar_columna(df, ['AHT', 'TMO', 'Handle_Time', 'Tiempo_Manejo', 'AHT_Segs'])
        col_prog = encontrar_columna(df, ['Agentes_Programados', 'Programados', 'Agentes', 'Roster', 'FTEs', 'Agentes_Programados_Reales'])
        col_camp = encontrar_columna(df, ['Campaña', 'Campana', 'Skill', 'Servicio', 'Project'])
        col_fecha = encontrar_columna(df, ['Fecha', 'Date', 'Day'])
        col_inter = encontrar_columna(df, ['Intervalo', 'Hora', 'Interval', 'Half_Hour'])
        col_dia = encontrar_columna(df, ['Día_Semana', 'Dia_Semana', 'Dia', 'Day_Of_Week'])

        if col_fecha:
            df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce').dt.strftime('%Y-%m-%d')

        records = df.to_dict(orient='records')
        del df
        gc.collect()

        data_processed = []

        for row in records:
            calls = clean_num(row.get(col_calls) if col_calls else 0, 0.0)
            aht = clean_num(row.get(col_aht) if col_aht else 180, 180.0)
            prog = clean_num(row.get(col_prog) if col_prog else 0, 0.0)
            
            campana_val = str(row.get(col_camp, 'General')) if col_camp and not pd.isna(row.get(col_camp)) else 'General'
            fecha_val = str(row.get(col_fecha, '')).split(' ')[0] if col_fecha and not pd.isna(row.get(col_fecha)) else ''
            dia_val = str(row.get(col_dia, '')) if col_dia and not pd.isna(row.get(col_dia)) else ''
            intervalo_val = str(row.get(col_inter, '00:00')) if col_inter and not pd.isna(row.get(col_inter)) else '00:00'

            a_erlang = (calls * aht) / 1800.0 if aht > 0 else 0.0
            req_raw = a_erlang / (1.0 - merma) if merma < 1.0 else a_erlang
            req_agents = math.ceil(req_raw)
            
            eval_agents = prog if prog > 0 else req_agents
            sl = erlang_c_sl_optimizado(a_erlang, eval_agents, aht, target_time)
            
            data_processed.append({
                'Campaña': campana_val,
                'Fecha': fecha_val,
                'Día_Semana': dia_val,
                'Intervalo': intervalo_val,
                'Llamadas': int(calls),
                'AHT': int(aht),
                'Agentes_Requeridos': req_agents,
                'Agentes_Programados_Reales': int(prog),
                'Delta_Net_Staffing': round(prog - req_agents, 1),
                'SL_Proyectado': sl
            })
            
        del records
        gc.collect()

        return jsonify(data_processed)

    except Exception as e:
        gc.collect()
        return jsonify({'error': f"Error al procesar el archivo: {str(e)}"}), 500

# Ruta comodín para evitar errores 404 en sub-rutas
@app.errorhandler(404)
def not_found(e):
    return jsonify({'error': 'La ruta solicitada no existe en el servidor Python'}), 404

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
