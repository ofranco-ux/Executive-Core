import os
import math
import gc
import re
import json
from datetime import datetime, timedelta
from flask import Flask, request, jsonify, send_from_directory, make_response
from flask_cors import CORS
import pandas as pd
import numpy as np

# --- LIBRERÍAS DE MACHINE LEARNING ---
import holidays
from sklearn.ensemble import RandomForestRegressor

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_FILE_IN = os.path.join(BASE_DIR, 'forecast_cache_in.json')
CACHE_FILE_OUT = os.path.join(BASE_DIR, 'forecast_cache_out.json')
CACHE_FILE_CHAT = os.path.join(BASE_DIR, 'forecast_cache_chat.json')
CONFIG_FILE = os.path.join(BASE_DIR, 'wfm_config.json') 
EXCEL_DEFAULT = os.path.join(BASE_DIR, 'historico.xlsx')

app = Flask(__name__)
CORS(app)

VENTANAS_SERVICIO = {
    'ambulancia servicios': {'inicio': 0 * 60, 'fin': 24 * 60},
    'asignación hogar': {'inicio': 0 * 60, 'fin': 24 * 60},
    'asignacion hogar': {'inicio': 0 * 60, 'fin': 24 * 60},
    'asignación vial': {'inicio': 0 * 60, 'fin': 24 * 60},
    'asignacion vial': {'inicio': 0 * 60, 'fin': 24 * 60},
    'coppel servicios': {'inicio': 0 * 60, 'fin': 24 * 60},
    'liverpool servicios': {'inicio': 0 * 60, 'fin': 24 * 60},
    'multicampañas': {'inicio': 0 * 60, 'fin': 24 * 60},
    'multicampanas': {'inicio': 0 * 60, 'fin': 24 * 60},
    'seguimiento hogar': {'inicio': 0 * 60, 'fin': 24 * 60},
    'seguimiento vial': {'inicio': 0 * 60, 'fin': 24 * 60},
    'suburbia servicios': {'inicio': 0 * 60, 'fin': 24 * 60},
    'experiencias liverpool': {'inicio': 9 * 60, 'fin': 21 * 60},
    'experiencias suburbia': {'inicio': 9 * 60, 'fin': 21 * 60},
    'retenciones suburbia': {'inicio': 9 * 60, 'fin': 20 * 60},
    'retenciones liverpool': {'inicio': 9 * 60, 'fin': 20 * 60}
}

# =====================================================================
# 🧠 MOTOR DE MACHINE LEARNING CON ATRIBUTOS TEMPORALES Y EXÓGENOS
# =====================================================================
def pronosticar_con_machine_learning(df_diario_campana, dias_futuros, fecha_inicio_forecast, col_fecha, col_calls):
    """
    Entrena un modelo autoregresivo Random Forest para predecir el volumen
    de llamadas de forma diaria incorporando lags, promedios móviles y festivos de México.
    """
    df_ml = df_diario_campana.sort_values(col_fecha).copy()
    
    anos_presentes = list(df_ml[col_fecha].dt.year.unique())
    anos_presentes.append(fecha_inicio_forecast.year)
    anos_presentes.append((fecha_inicio_forecast + timedelta(days=dias_futuros)).year)
    anos_unicos = list(set(anos_presentes))
    
    festivos_pais = holidays.CountryHoliday('MX', years=anos_unicos)
    
    df_ml['lag_1'] = df_ml[col_calls].shift(1)
    df_ml['lag_2'] = df_ml[col_calls].shift(2)
    df_ml['lag_7'] = df_ml[col_calls].shift(7)
    df_ml['lag_14'] = df_ml[col_calls].shift(14)
    
    df_ml['rolling_mean_7'] = df_ml[col_calls].shift(1).rolling(window=7).mean()
    df_ml['rolling_mean_30'] = df_ml[col_calls].shift(1).rolling(window=30).mean()
    
    df_ml['dia_semana'] = df_ml[col_fecha].dt.weekday
    df_ml['dia_mes'] = df_ml[col_fecha].dt.day
    df_ml['es_fin_de_mes'] = df_ml[col_fecha].dt.is_month_end.astype(int)
    df_ml['es_festivo'] = df_ml[col_fecha].apply(lambda x: 1 if x in festivos_pais else 0)
    
    df_train = df_ml.dropna().copy()
    
    if len(df_train) < 14:
        promedio_seguro = df_diario_campana[col_calls].mean()
        return [max(0.0, promedio_seguro)] * dias_futuros

    features = ['lag_1', 'lag_2', 'lag_7', 'lag_14', 'rolling_mean_7', 'rolling_mean_30', 
                'dia_semana', 'dia_mes', 'es_fin_de_mes', 'es_festivo']
    
    X_train = df_train[features]
    y_train = df_train[col_calls]
    
    modelo = RandomForestRegressor(n_estimators=100, random_state=42, max_depth=10)
    modelo.fit(X_train, y_train)
    
    historial_simulado = df_ml.to_dict('records')
    preds_finales = []
    fecha_actual = fecha_inicio_forecast
    
    for d in range(dias_futuros):
        vols_recientes = [r[col_calls] for r in historial_simulado]
        rm_7 = np.mean(vols_recientes[-7:]) if len(vols_recientes) >= 7 else np.mean(vols_recientes)
        rm_30 = np.mean(vols_recientes[-30:]) if len(vols_recientes) >= 30 else np.mean(vols_recientes)
        
        X_pred = pd.DataFrame([{
            'lag_1': historial_simulado[-1][col_calls],
            'lag_2': historial_simulado[-2][col_calls] if len(historial_simulado) > 1 else historial_simulado[-1][col_calls],
            'lag_7': historial_simulado[-7][col_calls] if len(historial_simulado) > 6 else historial_simulado[-1][col_calls],
            'lag_14': historial_simulado[-14][col_calls] if len(historial_simulado) > 13 else historial_simulado[-1][col_calls],
            'rolling_mean_7': rm_7,
            'rolling_mean_30': rm_30,
            'dia_semana': fecha_actual.weekday(),
            'dia_mes': fecha_actual.day,
            'es_fin_de_mes': 1 if (fecha_actual + timedelta(days=1)).day == 1 else 0,
            'es_festivo': 1 if fecha_actual in festivos_pais else 0
        }])
        
        pred_vol = float(modelo.predict(X_pred[features]))
        pred_vol = max(0.0, pred_vol)
        preds_finales.append(pred_vol)
        
        historial_simulado.append({col_fecha: fecha_actual, col_calls: pred_vol})
        fecha_actual += timedelta(days=1)
        
    return preds_finales

# =====================================================================
# ⚙️ MÓDULOS DE PROCESAMIENTO AUXILIARES Y ERLANG C
# =====================================================================
def buscar_archivo_excel():
    if os.path.exists(EXCEL_DEFAULT): return EXCEL_DEFAULT
    try:
        archivos = [f for f in os.listdir(BASE_DIR) if f.lower().endswith('.xlsx') and not f.startswith('~')]
        if not archivos: return None
        for f in archivos:
            if 'historico' in f.lower(): return os.path.join(BASE_DIR, f)
        return os.path.join(BASE_DIR, archivos[0])
    except: return None

@app.route('/')
@app.route('/index.html')
def serve_index():
    rutas_a_buscar = [BASE_DIR, os.getcwd(), os.path.dirname(BASE_DIR)]
    for ruta in rutas_a_buscar:
        target_path = os.path.join(ruta, 'index.html')
        if os.path.exists(target_path):
            response = make_response(send_from_directory(ruta, 'index.html'))
            response.headers['Cache-Control'] = 'no-cache, no-store, must-revalidate'
            response.headers['Pragma'] = 'no-cache'
            response.headers['Expires'] = '0'
            return response
    return jsonify({"error": "ALERTA CRÍTICA: No se encontró el archivo index.html."}), 404

@app.route('/favicon.ico')
def favicon(): return '', 204

@app.route('/api/config', methods=['GET', 'POST'])
def manage_config():
    if request.method == 'POST':
        try:
            new_config = request.get_json(force=True)
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f: json.dump(new_config, f)
            return jsonify({'status': 'Guardado'}), 200
        except Exception as e: return jsonify({'error': str(e)}), 500
    else:
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f: return jsonify(json.load(f)), 200
            except: pass
        return jsonify({'targetSl': 80, 'targetTime': 20, 'merma': 30}), 200

def clean_num(val, default=0.0):
    if pd.isna(val) or val is None: return default
    try:
        val_str = str(val).strip().replace(',', '.')
        val_str = re.sub(r'[^0-9.]', '', val_str)
        return float(val_str) if val_str else default
    except: return default

def parse_aht_to_seconds(val):
    if pd.isna(val) or val is None: return 180.0
    if isinstance(val, (int, float)): return float(val) if float(val) > 15 else float(val) * 60.0
    val_str = str(val).strip()
    if ':' in val_str:
        p = val_str.split(':')
        try:
            if len(p) == 3: return int(p[0]) * 3600 + int(p[1]) * 60 + float(p[2])
            elif len(p) == 2: return int(p[0]) * 60 + float(p[1])
        except: pass
    return clean_num(val_str, 180.0)

def format_aht_str(seconds):
    if pd.isna(seconds) or seconds is None or seconds <= 0: return "00:00:00"
    secs = int(round(seconds))
    return f"{secs // 3600:02d}:{(secs % 3600) // 60:02d}:{secs % 60:02d}"

def clean_interval_str(val):
    try:
        if pd.isna(val): return "00:00"
        val_str = str(val).strip()
        if hasattr(val, 'hour') and hasattr(val, 'minute'): hh, mm = val.hour, val.minute
        else:
            m = re.search(r'(\d{1,2}):(\d{2})', val_str)
            if m: hh, mm = int(m.group(1)), int(m.group(2))
            else: return "00:00"
        if mm < 15: mm_round = 0
        elif mm < 45: mm_round = 30
        else:
            mm_round = 0
            hh = (hh + 1) % 24
        return f"{hh:02d}:{mm_round:02d}"
    except: return "00:00"

ERLANG_CACHE = {}
def erlang_c_sl_optimizado(A, N, AHT, target_time):
    if N <= A or A <= 0 or N <= 0: return 0.0
    key = (round(A, 2), N, round(AHT, 1), target_time)
    if key in ERLANG_CACHE: return ERLANG_CACHE[key]
    try:
        sum_terms, current_term = 1.0, 1.0
        int_N = min(int(N), 1000)
        for k in range(1, int_N):
            current_term *= (A / k)
            sum_terms += current_term
        last_term = current_term * (A / N) / (1.0 - (A / N))
        pw = last_term / (sum_terms + last_term)
        sl = 1.0 - (pw * math.exp(-(N - A) * (target_time / AHT)))
        resultado = round(max(0.0, min(100.0, sl * 100.0)), 1)
        ERLANG_CACHE[key] = resultado
        return resultado
    except: return 0.0

def calcular_agentes_requeridos_erlang_c(A, aht, target_time, target_sl):
    if A <= 0 or aht <= 0: return 0
    n = max(1, int(math.floor(A)) + 1)
    if A > 50: n = max(n, int(math.floor(A + math.sqrt(A))))
    while n < 3000:
        if erlang_c_sl_optimizado(A, n, aht, target_time) >= target_sl: 
            return n
