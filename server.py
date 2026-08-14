import os
import math
import gc
import re
import json
from datetime import datetime, timedelta
from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import pandas as pd
import numpy as np

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CACHE_FILE = os.path.join(BASE_DIR, 'forecast_cache.json')
EXCEL_DEFAULT = os.path.join(BASE_DIR, 'historico.xlsx')

app = Flask(__name__, static_folder=BASE_DIR, static_url_path='')
CORS(app)

VENTANAS_SERVICIO = {
    'experiencias liverpool': {'inicio': 9 * 60, 'fin': 21 * 60},
    'experiencias suburbia':  {'inicio': 9 * 60, 'fin': 21 * 60},
    'retenciones liverpool':   {'inicio': 9 * 60, 'fin': 20 * 60},
    'retenciones suburbia':    {'inicio': 9 * 60, 'fin': 20 * 60},
    'Coppel Servicios':        {'inicio': 0 * 60, 'fin': 24 * 60}
}

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

def format_aht_str(seconds):
    if pd.isna(seconds) or seconds is None or seconds <= 0:
        return "00:00"
    secs = int(round(seconds))
    hrs = secs // 3600
    mins = (secs % 3600) // 60
    return f"{hrs:02d}:{mins:02d}"

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

def calcular_agentes_requeridos_erlang_c(A, aht, target_time, target_sl):
    if A <= 0 or aht <= 0:
        return 0
    n = max(1, int(math.floor(A)) + 1)
    while n < 1000:
        sl = erlang_c_sl_optimizado(A, n, aht, target_time)
        if sl >= target_sl:
            return n
        n += 1
    return n

def parse_time_str(t_str):
    try:
        parts = str(t_str).strip().split(':')
        return int(parts[0]) * 60 + int(parts[1])
    except Exception:
        return None

def esta_en_ventana_servicio(campana, intervalo_str):
    camp_key = str(campana).strip().lower()
    minutos_inter = parse_time_str(intervalo_str)
    if minutos_inter is None:
        return True
    ventana = VENTANAS_SERVICIO.get(camp_key)
    if not ventana:
        return True
    return ventana['inicio'] <= minutos_inter < ventana['fin']

def obtener_ventana_global(campanas_lista):
    inicios, fines = [], []
    for c in campanas_lista:
        c_key = str(c).strip().lower()
        if c_key in VENTANAS_SERVICIO:
            inicios.append(VENTANAS_SERVICIO[c_key]['inicio'])
            fines.append(VENTANAS_SERVICIO[c_key]['fin'])
    inicio_global = min(inicios) if inicios else 9 * 60
    fin_global = max(fines) if fines else 21 * 60
    return inicio_global, fin_global

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
    n_rows = X_new.shape[0] if hasattr(X_new, 'shape') else len(X_new)
    X_b = np.c_[np.ones((n_rows, 1)), X_new]
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
    if len(series) < 21:
        return holt_winters_fit_predict(series, n_preds=n_preds)
    train = np.array(series[:-14])
    val_true = np.array(series[-14:])
    best_wmape = float('inf')
    best_params = (0.2, 0.05, 0.2)
    sum_true = np.sum(val_true) if np.sum(val_true) > 0 else 1.0

    for a in [0.1, 0.2, 0.3]:
        for b in [0.01, 0.05, 0.1]:
            for g in [0.1, 0.2, 0.3, 0.5]:
                p_val = np.array(holt_winters_fit_predict(train, season_len=7, alpha=a, beta=b, gamma=g, n_preds=14))
                wmape = (np.sum(np.abs(val_true - p_val)) / sum_true) * 100
                if wmape < best_wmape:
                    best_wmape = wmape
                    best_params = (a, b, g)
                    
    a_opt, b_opt, g_opt = best_params
    return holt_winters_fit_predict(series, season_len=7, alpha=a_opt, beta=b_opt, gamma=g_opt, n_preds=n_preds)

def limpiar_outliers_iqr(series_list):
    if len(series_list) < 14:
        return list(series_list)
    arr = np.array(series_list)
    q25, q75 = np.percentile(arr, 25), np.percentile(arr, 75)
    iqr = q75 - q25
    lower, upper = q25 - 1.5 * iqr, q75 + 1.5 * iqr
    return np.clip(arr, lower, upper).tolist()

def procesar_archivo_excel(file_source, target_sl=80.0, target_time=20.0, merma=0.20, dias_futuros=30):
    xls_file = pd.ExcelFile(file_source, engine='openpyxl')
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

    df_diario = df.groupby([col_fecha, col_camp])[col_calls].sum().reset_index()
    campanas_unicas = df[col_camp].unique()

    modelos_ml, historial_volumenes, hw_forecasts = {}, {}, {}

    for camp in campanas_unicas:
        sub = df_diario[df_diario[col_camp] == camp].sort_values(col_fecha).reset_index(drop=True)
        fechas_list = sub[col_fecha].tolist()
        volumenes_list = limpiar_outliers_iqr(sub[col_calls].tolist())
        historial_volumenes[camp] = list(volumenes_list)
        hw_forecasts[camp] = grid_search_auto_hw(volumenes_list, n_preds=dias_futuros)

        X_data, y_data = [], []
        for i in range(14, len(sub)):
            f = fechas_list[i]
            feat = extraer_features_fecha(f, volumenes_list[:i], trend_idx=i)
            X_data.append(feat)
            y_data.append(volumenes_list[i])

        if len(X_data) > 10:
            X_arr, y_arr = np.array(X_data), np.array(y_data)
            weights, mean, std = entrenar_ridge_ml(X_arr, y_arr, l2_reg=10.0)
            modelos_ml[camp] = {'weights': weights, 'mean': mean, 'std': std, 'promedio_base': np.mean(y_arr)}
        else:
            modelos_ml[camp] = None

    df['Dia_Semana_Clean'] = df[col_dia].astype(str).str.strip().str.lower() if col_dia else df[col_fecha].dt.day_name().str.lower()
    df['Inter_Clean'] = df[col_inter].astype(str).str.strip().apply(lambda x: ':'.join(x.split(':')[:2]) if len(x.split(':')) == 3 else x)

    df['En_Ventana'] = df.apply(lambda r: esta_en_ventana_servicio(r[col_camp], r['Inter_Clean']), axis=1)
    df_filtrado = df[df['En_Ventana']].copy()

    max_date_hist = df_filtrado[col_fecha].max()
    df_reciente = df_filtrado[df_filtrado[col_fecha] >= (max_date_hist - timedelta(days=60))]

    perfil_intradia = df_reciente.groupby([col_camp, 'Dia_Semana_Clean', 'Inter_Clean']).agg(
        avg_calls=(col_calls, 'mean'),
        avg_aht=(col_aht, lambda x: x[x > 0].mean() if len(x[x > 0]) > 0 else 0)
    ).reset_index()

    totales_dia = perfil_intradia.groupby([col_camp, 'Dia_Semana_Clean'])['avg_calls'].transform('sum')
    perfil_intradia['weight'] = [(c / t) if t > 0 else 0 for c, t in zip(perfil_intradia['avg_calls'], totales_dia)]

    mapa_perfil = {}
    for _, r in perfil_intradia.iterrows():
        key = (r[col_camp], r['Dia_Semana_Clean'], r['Inter_Clean'])
        mapa_perfil[key] = {'weight': r['weight'], 'aht': r['avg_aht']}

    intervalos_operativos_por_camp = {}
    for camp in campanas_unicas:
        inters_camp = df_filtrado[df_filtrado[col_camp] == camp]['Inter_Clean'].unique().tolist()
        intervalos_operativos_por_camp[camp] = sorted([i for i in inters_camp if esta_en_ventana_servicio(camp, i)])

    del df, df_diario, df_filtrado, df_reciente
    gc.collect()

    factor_asistencia = max(0.01, 1.0 - merma)
    dias_espanol = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo']
    data_processed = []

    for d in range(dias_futuros):
        fecha_actual = fecha_inicio_forecast + timedelta(days=d)
        str_fecha = fecha_actual.strftime('%Y-%m-%d')
        nombre_dia = dias_espanol[fecha_actual.weekday()]

        for camp in campanas_unicas:
            hist_vol = historial_volumenes[camp]
            feat_futuras = np.array([extraer_features_fecha(fecha_actual, hist_vol, len(hist_vol))])

            model_info = modelos_ml.get(camp)
            if model_info:
                vol_ridge = predecir_ridge_ml(model_info['weights'], model_info['mean'], model_info['std'], feat_futuras)
                vol_ridge = max(vol_ridge, model_info['promedio_base'] * 0.15)
            else:
                vol_ridge = np.mean(hist_vol[-7:]) if hist_vol else 100.0

            vol_hw = hw_forecasts[camp][d] if d < len(hw_forecasts[camp]) else vol_ridge
            volumen_predicho_diario = (0.65 * vol_hw) + (0.35 * vol_ridge)
            historial_volumenes[camp].append(volumen_predicho_diario)

            intervalos_validos = intervalos_operativos_por_camp.get(camp, [])

            for inter in intervalos_validos:
                key_p = (camp, nombre_dia, inter)
                info_p = mapa_perfil.get(key_p, {'weight': 0.0, 'aht': 0.0})
                calls = volumen_predicho_diario * info_p['weight']
                aht = info_p['aht'] if (info_p['aht'] > 0 and not pd.isna(info_p['aht'])) else aht_global_campana.get(camp, 180.0)

                a_erlang = (calls * aht) / 1800.0 if (aht > 0 and calls > 0) else 0.0
                
                # REQUERIDO EN HEADCOUNT (HC) APEGADO STRICTAMENTE AL TARGET SL DINÁMICO
                req_ftes = calcular_agentes_requeridos_erlang_c(a_erlang, aht, target_time, target_sl) if calls > 0 else 0
                req_hc = math.ceil(req_ftes / factor_asistencia) if req_ftes > 0 else 0

                key_roster = (str(camp).lower(), nombre_dia.lower(), inter)
                prog_nominal_hc = matriz_roster.get(key_roster, req_hc)
                
                prog_efectivo_raw = prog_nominal_hc * factor_asistencia if calls > 0 else 0.0
                prog_efectivo_int = int(round(prog_nominal_hc))

                sl = erlang_c_sl_optimizado(a_erlang, prog_efectivo_raw, aht, target_time) if calls > 0 else 100.0
                delta_net_hc = int(prog_efectivo_int - req_hc) if calls > 0 else 0

                data_processed.append({
                    'Campaña': str(camp),
                    'Fecha': str_fecha,
                    'Día_Semana': nombre_dia.capitalize(),
                    'Intervalo': inter,
                    'Llamadas': int(round(calls)),
                    'AHT': format_aht_str(aht),
                    'AHT_Segundos': int(round(aht)),
                    'Agentes_Requeridos': req_hc,
                    'Agentes_Programados_Reales': prog_efectivo_int,
                    'Delta_Net_Staffing': delta_net_hc,
                    'SL_Proyectado': sl
                })

    try:
        with open(CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(data_processed, f)
    except Exception as err:
        print("Error guardando cache:", err)

    return data_processed

# --- MOTOR DE OPTIMIZACIÓN LÍMITES STRICTOS (EVITA EL SOBRE-DIMENSIONAMIENTO EN TARDE) ---
def resolver_turnos_optimos(intervalos, req_vector, campanas_activas, llamadas_vec=None, aht_vec=None, 
                            target_sl=80.0, target_time=20.0, merma=0.20, duracion_jornada=8.0, 
                            es_nocturno=False):
    m = len(intervalos)
    if m == 0:
        return [], [0]*m, 0, 0, 100.0, [100.0]*m, 100.0, 100.0

    llamadas_arr = np.array(llamadas_vec, dtype=float) if llamadas_vec is not None else np.zeros(m)
    aht_arr = np.array(aht_vec, dtype=float) if aht_vec is not None else np.full(m, 180.0)
    tot_llamadas = np.sum(llamadas_arr)
    factor_asistencia = max(0.01, 1.0 - merma)
    req_hc_arr = np.array(req_vector, dtype=float)

    cob_hc = np.zeros(m, dtype=float)
    x_turnos_dict = {}

    agentes_nocturnos_totales_hc = 0
    agentes_diurnos_totales_hc = 0

    target_sl_dinamico = float(target_sl)

    # --- PASO 1: TURNO NOCTURNO FIJO (22:00 A 07:00 / 5x2) ---
    if es_nocturno:
        label_jornada_noc = "9.0 hrs (Nocturno 5x2)"
        indices_nocturnos = []

        for j in range(m):
            min_in = parse_time_str(intervalos[j])
            if min_in is not None:
                if min_in >= (22 * 60) or min_in < (7 * 60):
                    indices_nocturnos.append(j)

        if len(indices_nocturnos) > 0:
            agentes_noc_hc = 1
            while agentes_noc_hc <= 200:
                cob_temp_ftes = agentes_noc_hc * factor_asistencia
                sl_acum, llamadas_noc = 0.0, 0.0
                for idx in indices_nocturnos:
                    c = llamadas_arr[idx]
                    aht_s = aht_arr[idx]
                    a_erl = (c * aht_s) / 1800.0 if (c > 0 and aht_s > 0) else 0.0
                    sl_v = erlang_c_sl_optimizado(a_erl, cob_temp_ftes, aht_s, target_time) if c > 0 else 100.0
                    sl_acum += (c * sl_v)
                    llamadas_noc += c

                sl_prom_noc = (sl_acum / llamadas_noc) if llamadas_noc > 0 else 100.0
                if sl_prom_noc >= target_sl_dinamico:
                    break
                agentes_noc_hc += 1

            key_turno_noc = ("22:00", "07:00", label_jornada_noc)
            x_turnos_dict[key_turno_noc] = agentes_noc_hc
            agentes_nocturnos_totales_hc = agentes_noc_hc

            for idx in indices_nocturnos:
                cob_hc[idx] = agentes_noc_hc

    # --- PASO 2: TURNOS DIURNOS (PROGRAMACIÓN AJUSTADA SIN CASCADA) ---
    duracion_jornada = float(duracion_jornada)
    SHIFT_BLOCKS = int(round(duracion_jornada * 2))
    duracion_minutos = int(round(duracion_jornada * 60))
    label_jornada_diurna = f"{duracion_jornada:.1f} hrs".replace('.0', '')

    min_diurno_inicio = 7 * 60    # 07:00
    min_diurno_limite = 22 * 60   # 22:00
    min_entrada_maxima = min_diurno_limite - duracion_minutos

    for j in range(m):
        min_in = parse_time_str(intervalos[j])
        if min_in is None:
            continue

        if min_in >= min_diurno_inicio and min_in < min_diurno_limite:
            c = llamadas_arr[j]
            aht_s = aht_arr[j]
            a_erl = (c * aht_s) / 1800.0 if (c > 0 and aht_s > 0) else 0.0
            
            # Requerimiento exacto en HC para este intervalo
            req_ftes_j = calcular_agentes_requeridos_erlang_c(a_erl, aht_s, target_time, target_sl_dinamico) if c > 0 else 0
            req_hc_j = math.ceil(req_ftes_j / factor_asistencia) if req_ftes_j > 0 else 0

            # Evaluar cuál es el déficit real considerando lo que ya se acumuló previamente
            deficit_hc = req_hc_j - cob_hc[j]

            # SOLO programar si existe un déficit real y si agregar ese turno no crea un sobre-exceso masivo en el bloque
            if deficit_hc > 0:
                agentes_hc_a_programar = math.ceil(deficit_hc)

                # Verificar que en los bloques futuros del turno no sobrepasemos por más de 2 agentes la demanda requerida de esos bloques
                idx_inicio_real = j
                for search_idx in range(m):
                    if parse_time_str(intervalos[search_idx]) == min_in:
                        idx_inicio_real = search_idx
                        break

                max_exceso_permitido = 2
                se_puede_agregar = True
                
                # Control de sobre-dimensionamiento
                for t in range(idx_inicio_real, min(idx_inicio_real + SHIFT_BLOCKS, m)):
                    min_t = parse_time_str(intervalos[t])
                    if min_t is not None and min_t >= min_diurno_limite:
                        break
                    
                    cobertura_futura = cob_hc[t] + agentes_hc_a_programar
                    req_futuro = req_hc_arr[t]
                    
                    if req_futuro > 0 and (cobertura_futura - req_futuro) > max_exceso_permitido:
                        # Si meter este turno completo causa un exceso masivo en la tarde, limitamos los agentes
                        agentes_hc_a_programar = max(1, int(req_futuro + max_exceso_permitido - cob_hc[t]))
                        if agentes_hc_a_programar <= 0:
                            se_puede_agregar = False
                        break

                if se_puede_agregar and agentes_hc_a_programar > 0:
                    min_entrada_efectiva = min_in
                    if min_entrada_efectiva > min_entrada_maxima:
                        min_entrada_efectiva = max(min_diurno_inicio, min_entrada_maxima)

                    min_out = min_entrada_efectiva + duracion_minutos
                    h_in_str = f"{(int(min_entrada_efectiva // 60)):02d}:{(int(min_entrada_efectiva % 60)):02d}"
                    h_out_str = f"{(int(min_out // 60)):02d}:{(int(min_out % 60)):02d}"

                    key_turno = (h_in_str, h_out_str, label_jornada_diurna)
                    x_turnos_dict[key_turno] = x_turnos_dict.get(key_turno, 0) + agentes_hc_a_programar

                    for t in range(idx_inicio_real, min(idx_inicio_real + SHIFT_BLOCKS, m)):
                        min_t = parse_time_str(intervalos[t])
                        if min_t is not None and min_t >= min_diurno_limite:
                            break
                        cob_hc[t] += agentes_hc_a_programar

    # --- PASO 3: METRICAS Y PROYECCIÓN FINAL DE SL ---
    sl_optimo_vector = []
    for i in range(m):
        c = llamadas_arr[i]
        aht_s = aht_arr[i]
        n_opt_ftes = cob_hc[i] * factor_asistencia
        a_erl = (c * aht_s) / 1800.0 if (c > 0 and aht_s > 0) else 0.0
        sl_val = erlang_c_sl_optimizado(a_erl, n_opt_ftes, aht_s, target_time) if c > 0 else 100.0
        sl_optimo_vector.append(sl_val)

    sl_arr = np.array(sl_optimo_vector)
    sl_optimo_global = float(np.sum(llamadas_arr * sl_arr) / tot_llamadas) if tot_llamadas > 0 else 100.0
    sl_optimo_global = round(sl_optimo_global, 1)

    cobertura_hc_entera = np.round(cob_hc).astype(int).tolist()
    turnos_sugeridos = []
    total_agentes_diarios_hc = 0

    for (h_in, h_out, label_dur), qty in x_turnos_dict.items():
        if qty > 0:
            turnos_sugeridos.append({
                'horario_entrada': h_in,
                'horario_salida': h_out,
                'agentes_a_programar': qty,
                'duracion': label_dur
            })
            total_agentes_diarios_hc += qty
            if "Nocturno" not in label_dur:
                agentes_diurnos_totales_hc += qty

    hc_nocturno = math.ceil(agentes_nocturnos_totales_hc * (7.0 / 5.0))
    hc_diurno = math.ceil(agentes_diurnos_totales_hc * (7.0 / 6.0))
    headcount_semanal_requerido = hc_nocturno + hc_diurno

    total_req_hc = np.sum(req_hc_arr)
    total_prog_hc = np.sum(cob_hc)
    staffing_level_optimo = round(float((total_prog_hc / total_req_hc * 100.0)), 1) if total_req_hc > 0 else 100.0
    eficiencia = round(min(100.0, (total_req_hc / total_prog_hc * 100.0)), 1) if total_prog_hc > 0 else 100.0

    return turnos_sugeridos, cobertura_hc_entera, total_agentes_diarios_hc, headcount_semanal_requerido, eficiencia, sl_optimo_vector, sl_optimo_global, staffing_level_optimo

# --- ENDPOINTS ---

@app.route('/api/optimize-schedules', methods=['POST'])
def api_optimize_schedules():
    try:
        body = request.get_json(force=True)
        intervalos = body.get('intervalos', [])
        requeridos = body.get('requeridos', [])
        campanas = body.get('campanas', [])
        llamadas = body.get('llamadas', [])
        ahts = body.get('ahts', [])
        target_sl = float(body.get('target_sl', 80.0))
        target_time = float(body.get('target_time', 20.0))
        merma = float(body.get('merma', 30.0)) / 100.0
        duracion_jornada = float(body.get('duracion_jornada', 8.0))
        es_nocturno = bool(body.get('es_nocturno', False))

        if not intervalos or not requeridos or len(intervalos) != len(requeridos):
            return jsonify({'error': 'Datos incompletos'}), 400

        turnos, cob_optima, total_diario, total_hc, eficiencia, sl_vec, sl_global, staff_level = resolver_turnos_optimos(
            intervalos, requeridos, campanas, llamadas_vec=llamadas, aht_vec=ahts, 
            target_sl=target_sl, target_time=target_time, merma=merma, 
            duracion_jornada=duracion_jornada, es_nocturno=es_nocturno
        )

        return jsonify({
            'turnos': turnos,
            'cobertura_optima': cob_optima,
            'total_agentes_diarios': total_diario,
            'headcount_semanal_6x1': total_hc,
            'eficiencia_cobertura': eficiencia,
            'sl_optimo_vector': sl_vec,
            'sl_optimo_global': sl_global,
            'staffing_level_optimo': staff_level
        }), 200

    except Exception as e:
        return jsonify({'error': f'Error optimizando turnos: {str(e)}'}), 500

@app.route('/api/optimize-schedules-weekly', methods=['POST'])
def api_optimize_schedules_weekly():
    try:
        body = request.get_json(force=True)
        semana_data = body.get('semana_data', [])
        campanas = body.get('campanas', [])
        target_sl = float(body.get('target_sl', 80.0))
        target_time = float(body.get('target_time', 20.0))
        merma = float(body.get('merma', 30.0)) / 100.0
        duracion_jornada = float(body.get('duracion_jornada', 8.0))
        es_nocturno = bool(body.get('es_nocturno', False))

        malla_semanal = {}
        headcount_max_dia = 0

        for dia_info in semana_data:
            nombre_dia = dia_info.get('dia_semana')
            intervalos = dia_info.get('intervalos', [])
            requeridos = dia_info.get('requeridos', [])
            llamadas = dia_info.get('llamadas', [])
            ahts = dia_info.get('ahts', [])

            turnos, cob_opt, total_d, hc_req, efic, sl_vec, sl_g, staff_lvl = resolver_turnos_optimos(
                intervalos, requeridos, campanas, llamadas_vec=llamadas, aht_vec=ahts,
                target_sl=target_sl, target_time=target_time, merma=merma, 
                duracion_jornada=duracion_jornada, es_nocturno=es_nocturno
            )

            malla_semanal[nombre_dia] = {
                'turnos': turnos,
                'cobertura': cob_opt,
                'agentes_dia': total_d,
                'sl_global_dia': sl_g,
                'staffing_level_dia': staff_lvl,
                'intervalos': intervalos,
                'requeridos': requeridos,
                'sl_vector': sl_vec
            }
            headcount_max_dia = max(headcount_max_dia, total_d)

        factor_hc = (7.0 / 5.0) if es_nocturno else (7.0 / 6.0)
        hc_semanal_total = math.ceil(headcount_max_dia * factor_hc)

        return jsonify({
            'malla_semanal': malla_semanal,
            'headcount_semanal_requerido': hc_semanal_total,
            'agentes_promedio_dia': math.ceil(sum(m['agentes_dia'] for m in malla_semanal.values()) / len(semana_data)) if semana_data else 0
        }), 200

    except Exception as e:
        return jsonify({'error': f'Error en proyección semanal: {str(e)}'}), 500

@app.route('/api/latest', methods=['GET'])
def get_latest_forecast():
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            if data and len(data) > 0:
                return jsonify(data), 200
        except Exception:
            pass

    if os.path.exists(EXCEL_DEFAULT):
        try:
            data = procesar_archivo_excel(EXCEL_DEFAULT)
            return jsonify(data), 200
        except Exception as e:
            return jsonify({'error': f'Error procesando historico.xlsx automático: {str(e)}'}), 500

    return jsonify({'error': 'No se encontró historico.xlsx en GitHub.'}), 404

@app.route('/api/process', methods=['POST', 'GET'])
@app.route('/api/process/', methods=['POST', 'GET'])
def process_data():
    if request.method == 'GET':
        return jsonify({'status': 'API predictiva activa'}), 200

    target_sl = clean_num(request.form.get('target_sl'), 80.0)
    target_time = clean_num(request.form.get('target_time'), 20.0)
    merma = clean_num(request.form.get('merma'), 20.0) / 100.0
    dias_futuros = int(clean_num(request.form.get('dias'), 30))

    if 'file' in request.files and request.files['file'].filename != '':
        file_source = request.files['file']
    elif os.path.exists(EXCEL_DEFAULT):
        file_source = EXCEL_DEFAULT
    else:
        return jsonify({'error': 'No se recibió archivo ni existe historico.xlsx.'}), 400

    try:
        data_processed = procesar_archivo_excel(file_source, target_sl, target_time, merma, dias_futuros)
        gc.collect()
        return jsonify(data_processed)
    except Exception as e:
        gc.collect()
        return jsonify({'error': f"Error: {str(e)}"}), 500

@app.errorhandler(404)
def not_found(e):
    return jsonify({'error': 'La ruta solicitada no existe'}), 404

if os.path.exists(EXCEL_DEFAULT) and not os.path.exists(CACHE_FILE):
    try:
        print("Procesando historico.xlsx inicial...")
        procesar_archivo_excel(EXCEL_DEFAULT)
    except Exception as e:
        print("Error inicial:", e)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
