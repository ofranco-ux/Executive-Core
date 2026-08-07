import os
import math
import gc
import re
from datetime import datetime, timedelta
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd

# Definir la ruta absoluta del directorio base
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
    try:
        parts = str(t_str).strip().split(':')
        return int(parts[0]) * 60 + int(parts[1])
    except Exception:
        return None

def construir_matriz_plantilla(xls_file):
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

def calcular_tendencia_lineal(valores):
    """ Regresión lineal simple sin depender de librerías externas """
    n = len(valores)
    if n < 2:
        return 0.0, valores[0] if n == 1 else 0.0
    x = list(range(n))
    sum_x = sum(x)
    sum_y = sum(valores)
    sum_xy = sum(i * j for i, j in zip(x, valores))
    sum_x2 = sum(i ** 2 for i in x)
    
    denom = (n * sum_x2 - sum_x ** 2)
    if denom == 0:
        return 0.0, sum_y / n
    slope = (n * sum_xy - sum_x * sum_y) / denom
    intercept = (sum_y - slope * sum_x) / n
    return slope, intercept

# --- RUTAS BACKEND API ---
@app.route('/api/process', methods=['POST', 'GET'])
@app.route('/api/process/', methods=['POST', 'GET'])
def process_data():
    if request.method == 'GET':
        return jsonify({'status': 'API predictiva activa en Render'}), 200

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
        matriz_roster = construir_matriz_plantilla(xls_file)

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
        col_fecha = encontrar_columna(df, ['Fecha', 'Date'])

        # Normalizar fechas
        df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce')
        df = df.dropna(subset=[col_fecha])

        fecha_maxima = df[col_fecha].max()
        fecha_inicio_forecast = fecha_maxima + timedelta(days=1)

        # -------------------------------------------------------------
        # MODELO PREDICTIVO SERIE TEMPORAL POR CAMPAÑA
        # -------------------------------------------------------------
        df_diario = df.groupby([col_fecha, col_camp])[col_calls].sum().reset_index()
        campanas_unicas = df[col_camp].unique()

        modelos_pronostico = {}
        for camp in campanas_unicas:
            sub = df_diario[df_diario[col_camp] == camp].sort_values(col_fecha)
            volumenes = sub[col_calls].tolist()
            if len(volumenes) > 1:
                slope, intercept = calcular_tendencia_lineal(volumenes)
                promedio_base = sum(volumenes) / len(volumenes)
            else:
                slope, intercept = 0.0, volumenes[0] if volumenes else 100.0
                promedio_base = intercept
            
            modelos_pronostico[camp] = {
                'slope': slope,
                'intercept': intercept,
                'len_hist': len(volumenes),
                'promedio_base': promedio_base
            }

        # -------------------------------------------------------------
        # PROFILE CURVE INTRADÍA (FACTOR DÍA_SEMANA E INTERVALO)
        # -------------------------------------------------------------
        df['Dia_Semana_Clean'] = df[col_dia].astype(str).str.strip().str.lower() if col_dia else df[col_fecha].dt.day_name().str.lower()
        
        df['Inter_Clean'] = df[col_inter].astype(str).str.strip()
        df['Inter_Clean'] = df['Inter_Clean'].apply(lambda x: ':'.join(x.split(':')[:2]) if len(x.split(':')) == 3 else x)

        perfil_intradia = df.groupby([col_camp, 'Dia_Semana_Clean', 'Inter_Clean']).agg(
            avg_calls=(col_calls, 'mean'),
            avg_aht=(col_aht, 'mean')
        ).reset_index()

        totales_dia = perfil_intradia.groupby([col_camp, 'Dia_Semana_Clean'])['avg_calls'].transform('sum')
        perfil_intradia['weight'] = [
            (c / t) if t > 0 else 0 
            for c, t in zip(perfil_intradia['avg_calls'], totales_dia)
        ]

        mapa_perfil = {}
        for _, r in perfil_intradia.iterrows():
            key = (r[col_camp], r['Dia_Semana_Clean'], r['Inter_Clean'])
            mapa_perfil[key] = {'weight': r['weight'], 'aht': r['avg_aht']}

        intervalos_unicos = sorted(df['Inter_Clean'].unique())

        del df, df_diario
        gc.collect()

        # -------------------------------------------------------------
        # GENERACIÓN DEL FORECAST PROYECTADO A FUTURO
        # -------------------------------------------------------------
        dias_espanol = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo']
        data_processed = []

        for d in range(dias_futuros):
            fecha_actual = fecha_inicio_forecast + timedelta(days=d)
            str_fecha = fecha_actual.strftime('%Y-%m-%d')
            nombre_dia = dias_espanol[fecha_actual.weekday()]

            for camp in campanas_unicas:
                mod = modelos_pronostico.get(camp, {'slope': 0.0, 'intercept': 100.0, 'len_hist': 10, 'promedio_base': 100.0})
                idx_futuro = mod['len_hist'] + d
                
                # Predicción con tendencia
                volumen_predicho_diario = mod['intercept'] + mod['slope'] * idx_futuro
                if volumen_predicho_diario < (mod['promedio_base'] * 0.2):
                    volumen_predicho_diario = mod['promedio_base']

                for inter in intervalos_unicos:
                    key_p = (camp, nombre_dia, inter)
                    info_p = mapa_perfil.get(key_p, {'weight': 1.0 / max(len(intervalos_unicos), 1), 'aht': 180.0})

                    calls = volumen_predicho_diario * info_p['weight']
                    aht = info_p['aht'] if info_p['aht'] > 0 else 180.0

                    a_erlang = (calls * aht) / 1800.0 if aht > 0 else 0.0
                    req_raw = a_erlang / (1.0 - merma) if merma < 1.0 else a_erlang
                    req_agents = math.ceil(req_raw)

                    key_roster = (str(camp).lower(), nombre_dia.lower(), inter)
                    prog = matriz_roster.get(key_roster, req_agents)

                    sl = erlang_c_sl_optimizado(a_erlang, prog, aht, target_time)

                    data_processed.append({
                        'Campaña': str(camp),
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
