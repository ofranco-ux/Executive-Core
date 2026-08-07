import os
import math
import gc
import re
from datetime import datetime, timedelta
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd

# Definir la ruta absoluta del directorio base para evitar pantallas en blanco / 404
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__, static_folder=BASE_DIR, static_url_path='')
CORS(app)

# --- RUTAS FRONTEND & ARCHIVOS ESTÁTICOS ---
@app.route('/')
@app.route('/index.html')
def serve_index():
    return send_from_directory(BASE_DIR, 'index.html')

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
    """
    Cálculo iterativo eficiente de Erlang C para evitar desbordamiento 
    de memoria (Memory Limit 512MB de Render).
    """
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
    """ Convierte formato HH:MM a minutos desde medianoche """
    try:
        parts = str(t_str).strip().split(':')
        return int(parts[0]) * 60 + int(parts[1])
    except Exception:
        return None

def construir_matriz_plantilla(xls_file):
    """
    Parsea la hoja Platilla / Plantilla y construye un mapa de capacidad:
    (Campaña, Día_Semana, Intervalo) -> Agentes_Presentes
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
        print("Error procesando hoja plantilla:", e)
        return {}

def encontrar_columna(df, posibles_nombres):
    columnas_df = {str(c).strip().lower(): c for c in df.columns}
    for pos in posibles_nombres:
        pos_clean = pos.strip().lower()
        if pos_clean in columnas_df:
            return columnas_df[pos_clean]
    return None

# --- RUTAS BACKEND API ---
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
        dias_futuros = int(clean_num(request.form.get('dias'), 30))

        xls_file = pd.ExcelFile(file)
        
        # 1. Mapear capacidad real desde la hoja Platilla
        matriz_roster = construir_matriz_plantilla(xls_file)

        # 2. Cargar datos históricos de tráfico
        sheet_calls = xls_file.sheet_names[0]
        for s in xls_file.sheet_names:
            if 'llam' in s.lower() or 'hist' in s.lower():
                sheet_calls = s
                break

        df = pd.read_excel(xls_file, sheet_name=sheet_calls)

        col_calls = encontrar_columna(df, ['Recibidas', 'Llamadas', 'Calls', 'Volumen', 'Ofrecidas'])
        col_aht = encontrar_columna(df, ['AHT', 'TMO', 'Handle_Time'])
        col_camp = encontrar_columna(df, ['Campaña', 'Campana', 'Ring Group', 'Skill'])
        col_inter = encontrar_columna(df, ['Intervalo', 'Hora'])
        col_dia = encontrar_columna(df, ['Día', 'Dia', 'Día_Semana', 'Dia_Semana'])

        records = df.to_dict(orient='records')
        del df
        gc.collect()

        # 3. Construir perfiles promedios por (Campaña, Día_Semana, Intervalo)
        perfiles_historicos = {}
        for row in records:
            calls = clean_num(row.get(col_calls) if col_calls else 0, 0.0)
            aht = clean_num(row.get(col_aht) if col_aht else 180, 180.0)
            campana_val = str(row.get(col_camp, 'General')).strip() if col_camp and not pd.isna(row.get(col_camp)) else 'General'
            dia_val = str(row.get(col_dia, '')).strip().lower() if col_dia and not pd.isna(row.get(col_dia)) else ''
            
            inter_raw = str(row.get(col_inter, '00:00')).strip() if col_inter and not pd.isna(row.get(col_inter)) else '00:00'
            intervalo_val = str(inter_raw).strip()
            if len(intervalo_val.split(':')) == 3:
                intervalo_val = ':'.join(intervalo_val.split(':')[:2])

            key = (campana_val, dia_val, intervalo_val)
            if key not in perfiles_historicos:
                perfiles_historicos[key] = {'calls': [], 'aht': []}
            perfiles_historicos[key]['calls'].append(calls)
            perfiles_historicos[key]['aht'].append(aht)

        mapa_promedios = {}
        for k, v in perfiles_historicos.items():
            mapa_promedios[k] = {
                'calls': sum(v['calls']) / len(v['calls']) if v['calls'] else 0,
                'aht': sum(v['aht']) / len(v['aht']) if v['aht'] else 180
            }

        del records
        del perfiles_historicos
        gc.collect()

        dias_espanol = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo']

        # 4. Generar Proyección Futura desde HOY hacia N Días
        data_processed = []
        fecha_inicio = datetime.now()

        campanas_unicas = sorted(list(set(k[0] for k in mapa_promedios.keys())))
        intervalos_unicos = sorted(list(set(k[2] for k in mapa_promedios.keys())))

        for d in range(dias_futuros):
            fecha_actual = fecha_inicio + timedelta(days=d)
            str_fecha = fecha_actual.strftime('%Y-%m-%d')
            nombre_dia = dias_espanol[fecha_actual.weekday()]

            for camp in campanas_unicas:
                for inter in intervalos_unicos:
                    key_hist = (camp, nombre_dia, inter)
                    
                    if key_hist in mapa_promedios:
                        calls = mapa_promedios[key_hist]['calls']
                        aht = mapa_promedios[key_hist]['aht']
                    else:
                        calls = 0.0
                        aht = 180.0

                    # Cálculo Erlang C y requerimiento
                    a_erlang = (calls * aht) / 1800.0 if aht > 0 else 0.0
                    req_raw = a_erlang / (1.0 - merma) if merma < 1.0 else a_erlang
                    req_agents = math.ceil(req_raw)

                    # Obtener programados desde la matriz Platilla
                    key_roster = (camp.lower(), nombre_dia.lower(), inter)
                    prog = matriz_roster.get(key_roster, req_agents)

                    sl = erlang_c_sl_optimizado(a_erlang, prog, aht, target_time)

                    data_processed.append({
                        'Campaña': camp,
                        'Fecha': str_fecha,
                        'Día_Semana': nombre_dia.capitalize(),
                        'Intervalo': inter,
                        'Llamadas': int(round(calls)),
                        'AHT': int(round(aht)),
                        'Agentes_Requeridos': req_agents,
                        'Agentes_Programados_Reales': prog,
                        'Delta_Net_Staffing': round(prog - req_agents, 1),
                        'SL_Proyectado': sl
                    })

        gc.collect()
        return jsonify(data_processed)

    except Exception as e:
        gc.collect()
        return jsonify({'error': f"Error al procesar el archivo: {str(e)}"}), 500

@app.errorhandler(404)
def not_found(e):
    return jsonify({'error': 'La ruta solicitada no existe en el servidor Python'}), 404

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
