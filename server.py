import os
import math
import gc
import re
from datetime import datetime, time
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd

app = Flask(__name__, static_folder='.', static_url_path='')
CORS(app)

@app.route('/')
@app.route('/index.html')
def serve_index():
    return send_from_directory('.', 'index.html')

@app.route('/favicon.ico')
def favicon():
    return '', 204

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

def parse_time_str(t_str):
    """ Convierte HH:MM a minutos desde medianoche """
    try:
        parts = str(t_str).strip().split(':')
        return int(parts[0]) * 60 + int(parts[1])
    except Exception:
        return None

def construir_matriz_plantilla(xls_file):
    """
    Lee la hoja Platilla/Plantilla y construye un mapa de capacidad:
    (Campaña, Día_Semana_LOWER, Intervalo_HH:MM) -> Agentes_Presentes
    """
    try:
        sheet_names = xls_file.sheet_names
        sheet_plantilla = None
        for s in sheet_names:
            if 'plat' in s.lower() or 'plan' in s.lower():
                sheet_plantilla = s
                break

        if not sheet_plantilla:
            return {}

        df_p = pd.read_excel(xls_file, sheet_name=sheet_plantilla)
        dias_cols = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
        
        malla = {}

        for _, row in df_p.iterrows():
            camp = str(row.get('Campaña', 'General')).strip()
            for dia in dias_cols:
                if dia not in df_p.columns:
                    continue
                horario = str(row.get(dia, '')).strip()
                if not horario or 'descanso' in horario.lower() or '-' not in horario:
                    continue
                
                try:
                    h_in, h_out = horario.split('-')
                    m_in = parse_time_str(h_in)
                    m_out = parse_time_str(h_out)

                    if m_in is not None and m_out is not None:
                        # Iterar intervalos de 30 minutos
                        cur = m_in
                        while cur < m_out:
                            hh = cur // 60
                            mm = cur % 60
                            inter_str = f"{hh:02d}:{mm:02d}"
                            
                            key = (camp.lower(), dia.lower(), inter_str)
                            malla[key] = malla.get(key, 0) + 1
                            cur += 30
                except Exception:
                    continue

        return malla
    except Exception as e:
        print("Error parseando plantilla:", e)
        return {}

def encontrar_columna(df, posibles_nombres):
    columnas_df = {str(c).strip().lower(): c for c in df.columns}
    for pos in posibles_nombres:
        pos_clean = pos.strip().lower()
        if pos_clean in columnas_df:
            return columnas_df[pos_clean]
    return None

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

        xls_file = pd.ExcelFile(file)
        
        # 1. Mapear capacidad real desde la hoja Platilla
        matriz_roster = construir_matriz_plantilla(xls_file)

        # 2. Leer hoja de llamadas
        sheet_calls = xls_file.sheet_names[0]
        for s in xls_file.sheet_names:
            if 'llam' in s.lower() or 'hist' in s.lower():
                sheet_calls = s
                break

        df = pd.read_excel(xls_file, sheet_name=sheet_calls)

        col_calls = encontrar_columna(df, ['Recibidas', 'Llamadas', 'Calls', 'Volumen', 'Ofrecidas'])
        col_aht = encontrar_columna(df, ['AHT', 'TMO', 'Handle_Time'])
        col_prog = encontrar_columna(df, ['Agentes_Programados', 'Programados', 'Agentes', 'Roster'])
        col_camp = encontrar_columna(df, ['Campaña', 'Campana', 'Ring Group', 'Skill'])
        col_fecha = encontrar_columna(df, ['Fecha', 'Date'])
        col_inter = encontrar_columna(df, ['Intervalo', 'Hora'])
        col_dia = encontrar_columna(df, ['Día', 'Dia', 'Día_Semana', 'Dia_Semana'])

        if col_fecha:
            df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce').dt.strftime('%Y-%m-%d')

        records = df.to_dict(orient='records')
        del df
        gc.collect()

        data_processed = []

        for row in records:
            calls = clean_num(row.get(col_calls) if col_calls else 0, 0.0)
            aht = clean_num(row.get(col_aht) if col_aht else 180, 180.0)
            
            campana_val = str(row.get(col_camp, 'General')) if col_camp and not pd.isna(row.get(col_camp)) else 'General'
            fecha_val = str(row.get(col_fecha, '')).split(' ')[0] if col_fecha and not pd.isna(row.get(col_fecha)) else ''
            dia_val = str(row.get(col_dia, '')).strip() if col_dia and not pd.isna(row.get(col_dia)) else ''
            
            inter_raw = str(row.get(col_inter, '00:00')) if col_inter and not pd.isna(row.get(col_inter)) else '00:00'
            intervalo_val = str(inter_raw).strip()
            if len(intervalo_val.split(':')) == 3:
                intervalo_val = ':'.join(intervalo_val.split(':')[:2])

            a_erlang = (calls * aht) / 1800.0 if aht > 0 else 0.0
            req_raw = a_erlang / (1.0 - merma) if merma < 1.0 else a_erlang
            req_agents = math.ceil(req_raw)

            # 3. Determinar Agentes Programados desde la hoja Platilla o la columna
            key_roster = (campana_val.lower(), dia_val.lower(), intervalo_val)
            if key_roster in matriz_roster:
                prog = matriz_roster[key_roster]
            else:
                prog_val = row.get(col_prog) if col_prog else None
                if prog_val is not None and not pd.isna(prog_val):
                    prog = int(clean_num(prog_val, req_agents))
                else:
                    prog = req_agents

            sl = erlang_c_sl_optimizado(a_erlang, prog, aht, target_time)
            
            data_processed.append({
                'Campaña': campana_val,
                'Fecha': fecha_val,
                'Día_Semana': dia_val,
                'Intervalo': intervalo_val,
                'Llamadas': int(calls),
                'AHT': int(aht),
                'Agentes_Requeridos': req_agents,
                'Agentes_Programados_Reales': prog,
                'Delta_Net_Staffing': round(prog - req_agents, 1),
                'SL_Proyectado': sl
            })
            
        del records
        gc.collect()

        return jsonify(data_processed)

    except Exception as e:
        gc.collect()
        return jsonify({'error': f"Error al procesar el archivo: {str(e)}"}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
