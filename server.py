import os
import math
import gc
import re
from datetime import datetime, timedelta
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd
import numpy as np

# Ruta absoluta del directorio base
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
    """ Cálculo iterativo eficiente de Erlang C """
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

        df_p = pd.read_excel(xls_file, sheet_name=sheet_plantilla, engine='openpyxl')
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

# ---------------------------------------------------------------------
# MOTOR ML 1: REGRESIÓN RIDGE AUTORREGRESIVA L2
# ---------------------------------------------------------------------
def entrenar_ridge_ml(X, y, l2_reg=10.0):
    X_b = np.c_[np.ones((X.shape[0], 1)), X]
    mean = np.mean(X_b[:, 1:], axis=0)
    std = np.std(X_b[:, 1:], axis=0) + 1e-8
    
    X_norm = X_b.copy()
    X_norm[:, 1:] = (X_b[:, 1:] - mean) / std
    
    I = np.eye(X_norm.shape[1])
    I[0, 0] = 0.0
    
    try:
        weights = np.linalg.inv(X_norm.T @ X_norm + l2_reg * I) @ X_norm.T @ y
    except np.linalg.LinAlgError:
        weights = np.linalg.pinv(X_norm.T @ X_norm + l2_reg * I) @ X_norm.T @ y
        
    return weights, mean, std

def predecir_ridge_ml(weights, mean, std, X_new):
    X_b = np.c_[np.ones((X_new.shape[0], 1)), X_new]
    X_norm = X_b.copy()
    X_norm[:, 1:] = (X_b[:, 1:] - mean) / std
    pred = X_norm @ weights
    return float(pred[0])

def extraer_features_fecha(fecha, volumenes_hist, trend_idx):
    day_of_week = fecha.weekday()
    day_of_month = fecha.day
    is_weekend = 1.0 if day_of_week >= 5 else 0.0
    is_quincena = 1.0 if day_of_month in [1, 15, 16, 30, 31] else 0.0
    
    lag_1 = volumenes_hist[-1] if len(volumenes_hist) >= 1 else 100.0
    lag_7 = volumenes_hist[-7] if len(volumenes_hist) >= 7 else lag_1
    lag_14 = volumenes_hist[-14] if len(volumenes_hist) >= 14 else lag_7

    dow_encoded = [1.0 if day_of_week == i else 0.0 for i in range(7)]
    return [lag_1, lag_7, lag_14, float(day_of_month), is_weekend, is_quincena, float(trend_idx)] + dow_encoded

# ---------------------------------------------------------------------
# MOTOR ML 2: HOLT-WINTERS TRIPLE SMOOTHING CON GRID SEARCH Y WMAPE
# ---------------------------------------------------------------------
def holt_winters_fit_predict(series, season_len=7, alpha=0.2, beta=0.1, gamma=0.3, n_preds=30):
    n = len(series)
    if n < season_len * 2:
        return [np.mean(series) if len(series) > 0 else 100.0] * n_preds
    
    level = np.mean(series[:season_len])
    trend = (np.mean(series[season_len:2*season_len]) - np.mean(series[:season_len])) / season_len
    seasonals = [series[i] - level for i in range(season_len)]
    
    for i in range(n):
        val = series[i]
        last_level, last_trend = level, trend
        st_prev = seasonals[i % season_len]
        
        level = alpha * (val - st_prev) + (1 - alpha) * (last_level + last_trend)
        trend = beta * (level - last_level) + (1 - beta) * last_trend
        seasonals[i % season_len] = gamma * (val - level) + (1 - gamma) * st_prev
        
    preds = []
    for m in range(1, n_preds + 1):
        p = level + m * trend + seasonals[(n + m - 1) % season_len]
        preds.append(max(0.0, float(p)))
    return preds

def grid_search_auto_hw(series, n_preds=30):
    """ Encuentra los mejores parámetros alpha, beta, gamma que minimizan el WMAPE """
    if len(series) < 21:
        return holt_winters_fit_predict(series, n_preds=n_preds)
    
    train = np.array(series[:-14])
    val_true = np.array(series[-14:])
    
    best_wmape = float('inf')
    best_params = (0.2, 0.05, 0.2)
    
    alphas = [0.1, 0.2, 0.3]
    betas = [0.01, 0.05, 0.1]
    gammas = [0.1, 0.2, 0.3, 0.5]
    
    sum_true = np.sum(val_true) if np.sum(val_true) > 0 else 1.0

    for a in alphas:
        for b in betas:
            for g in gammas:
                p_val = np.array(holt_winters_fit_predict(train, season_len=7, alpha=a, beta=b, gamma=g, n_preds=14))
                wmape = (np.sum(np.abs(val_true - p_val)) / sum_true) * 100
                if wmape < best_wmape:
                    best_wmape = wmape
                    best_params = (a, b, g)
                    
    a_opt, b_opt, g_opt = best_params
    return holt_winters_fit_predict(series, season_len=7, alpha=a_opt, beta=b_opt, gamma=g_opt, n_preds=n_preds)

def limpiar_outliers_iqr(series_list):
    """ Limpia valores atípicos mediante el rango intercuartílico (IQR) """
    if len(series_list) < 14:
        return list(series_list)
    arr = np.array(series_list)
    q25, q75 = np.percentile(arr, 25), np.percentile(arr, 75)
    iqr = q75 - q25
    lower, upper = q25 - 1.5 * iqr, q75 + 1.5 * iqr
    return np.clip(arr, lower, upper).tolist()

# --- RUTAS BACKEND API ---
@app.route('/api/process', methods=['POST', 'GET'])
@app.route('/api/process/', methods=['POST', 'GET'])
def process_data():
    if request.method == 'GET':
        return jsonify({'status': 'API predictiva avanzada activa en Render'}), 200

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

        xls_file = pd.ExcelFile(file, engine='openpyxl')
        matriz_roster = construir_matriz_plantilla(xls_file)

        sheet_calls = xls_file.sheet_names[0]
        for s in xls_file.sheet_names:
            if 'llam' in s.lower() or 'hist' in s.lower():
                sheet_calls = s
                break

        df = pd.read_excel(xls_file, sheet_name=sheet_calls, engine='openpyxl')

        col_calls = encontrar_columna(df, ['Recibidas', 'Llamadas', 'Calls', 'Volumen', 'Ofrecidas'])
        col_aht = encontrar_columna(df, ['AHT', 'TMO', 'Handle_Time'])
        col_camp = encontrar_columna(df, ['Campaña', 'Campana', 'Ring Group', 'Skill'])
        col_inter = encontrar_columna(df, ['Intervalo', 'Hora'])
        col_dia = encontrar_columna(df, ['Día', 'Dia', 'Día_Semana', 'Dia_Semana'])
        col_fecha = encontrar_columna(df, ['Fecha', 'Date'])

        df[col_fecha] = pd.to_datetime(df[col_fecha], errors='coerce')
        df = df.dropna(subset=[col_fecha])

        fecha_maxima = df[col_fecha].max()
        fecha_inicio_forecast = fecha_maxima + timedelta(days=1)

        aht_global_campana = df.groupby(col_camp)[col_aht].apply(lambda x: x[x > 0].mean() if len(x[x > 0]) > 0 else 180.0).to_dict()

        # -------------------------------------------------------------
        # 1. ENTRENAMIENTO ENSAMBLE OPTIMIZADO (GRID SEARCH HW + RIDGE)
        # -------------------------------------------------------------
        df_diario = df.groupby([col_fecha, col_camp])[col_calls].sum().reset_index()
        campanas_unicas = df[col_camp].unique()

        modelos_ml = {}
        historial_volumenes = {}
        hw_forecasts = {}

        for camp in campanas_unicas:
            sub = df_diario[df_diario[col_camp] == camp].sort_values(col_fecha).reset_index(drop=True)
            fechas_list = sub[col_fecha].tolist()
            raw_volumenes = sub[col_calls].tolist()
            
            volumenes_list = limpiar_outliers_iqr(raw_volumenes)
            historial_volumenes[camp] = list(volumenes_list)

            hw_forecasts[camp] = grid_search_auto_hw(volumenes_list, n_preds=dias_futuros)

            X_data, y_data = [], []
            for i in range(14, len(sub)):
                f = fechas_list[i]
                v_hist = volumenes_list[:i]
                feat = extraer_features_fecha(f, v_hist, trend_idx=i)
                X_data.append(feat)
                y_data.append(volumenes_list[i])

            if len(X_data) > 10:
                X_arr, y_arr = np.array(X_data), np.array(y_data)
                weights, mean, std = entrenar_ridge_ml(X_arr, y_arr, l2_reg=10.0)
                modelos_ml[camp] = {
                    'weights': weights,
                    'mean': mean,
                    'std': std,
                    'promedio_base': np.mean(y_arr)
                }
            else:
                modelos_ml[camp] = None

        # -------------------------------------------------------------
        # 2. PROFILE CURVE INTRADÍA
        # -------------------------------------------------------------
        df['Dia_Semana_Clean'] = df[col_dia].astype(str).str.strip().str.lower() if col_dia else df[col_fecha].dt.day_name().str.lower()
        
        df['Inter_Clean'] = df[col_inter].astype(str).str.strip()
        df['Inter_Clean'] = df['Inter_Clean'].apply(lambda x: ':'.join(x.split(':')[:2]) if len(x.split(':')) == 3 else x)

        max_date_hist = df[col_fecha].max()
        df_reciente = df[df[col_fecha] >= (max_date_hist - timedelta(days=60))]

        perfil_intradia = df_reciente.groupby([col_camp, 'Dia_Semana_Clean', 'Inter_Clean']).agg(
            avg_calls=(col_calls, 'mean'),
            avg_aht=(col_aht, lambda x: x[x > 0].mean() if len(x[x > 0]) > 0 else 0)
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

        del df, df_diario, df_reciente
        gc.collect()

        # -------------------------------------------------------------
        # 3. GENERACIÓN DEL FORECAST ENSAMBLADO FINAL
        # -------------------------------------------------------------
        dias_espanol = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo']
        data_processed = []

        for d in range(dias_futuros):
            fecha_actual = fecha_inicio_forecast + timedelta(days=d)
            str_fecha = fecha_actual.strftime('%Y-%m-%d')
            nombre_dia = dias_espanol[fecha_actual.weekday()]

            for camp in campanas_unicas:
                hist_vol = historial_volumenes[camp]
                trend_idx = len(hist_vol)
                feat_futuras = np.array([extraer_features_fecha(fecha_actual, hist_vol, trend_idx)])

                model_info = modelos_ml.get(camp)
                if model_info:
                    vol_ridge = predecir_ridge_ml(
                        model_info['weights'], model_info['mean'], model_info['std'], feat_futuras
                    )
                    vol_ridge = max(vol_ridge, model_info['promedio_base'] * 0.15)
                else:
                    vol_ridge = np.mean(hist_vol[-7:]) if hist_vol else 100.0

                vol_hw = hw_forecasts[camp][d] if d < len(hw_forecasts[camp]) else vol_ridge

                volumen_predicho_diario = (0.65 * vol_hw) + (0.35 * vol_ridge)

                historial_volumenes[camp].append(volumen_predicho_diario)

                for inter in intervalos_unicos:
                    key_p = (camp, nombre_dia, inter)
                    info_p = mapa_perfil.get(key_p, {'weight': 1.0 / max(len(intervalos_unicos), 1), 'aht': 0})

                    calls = volumen_predicho_diario * info_p['weight']
                    
                    aht = info_p['aht']
                    if aht <= 0 or pd.isna(aht):
                        aht = aht_global_campana.get(camp, 180.0)

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
